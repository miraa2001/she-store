begin;

alter table public.purchases add column if not exists received_amount numeric(14,2) not null default 0,
  add column if not exists remitted_amount numeric(14,2) not null default 0,
  add column if not exists payment_model_ready boolean not null default false;
alter table public.instant_pickups add column if not exists received_amount numeric(14,2) not null default 0,
  add column if not exists remitted_amount numeric(14,2) not null default 0,
  add column if not exists payment_model_ready boolean not null default false;

create table if not exists public.purchase_payments (
  id uuid primary key default gen_random_uuid(),
  purchase_id uuid references public.purchases(id) on delete restrict,
  instant_pickup_id uuid references public.instant_pickups(id) on delete restrict,
  amount numeric(14,2) not null check (amount >= 0 and amount::text not in ('NaN','Infinity','-Infinity')),
  pickup_point text not null,
  legacy boolean not null default false,
  received_on_pickup boolean not null default false,
  created_at timestamptz not null default clock_timestamp(),
  created_by uuid references auth.users(id),
  remitted_at timestamptz,
  remitted_by uuid references auth.users(id),
  cash_recognized_amount numeric(14,2) not null default 0,
  cash_transaction_id uuid references public.home_cash_transactions(id),
  cancelled_at timestamptz,
  cancelled_by uuid references auth.users(id),
  cancellation_reason text,
  cancellation_cash_id uuid references public.home_cash_transactions(id),
  check ((purchase_id is null) <> (instant_pickup_id is null)),
  check (legacy or amount > 0),
  check (cash_recognized_amount between 0 and amount)
);
create index if not exists purchase_payments_purchase_idx on public.purchase_payments(purchase_id,created_at,id);
create index if not exists purchase_payments_instant_idx on public.purchase_payments(instant_pickup_id,created_at,id);
create unique index if not exists purchase_payments_legacy_purchase on public.purchase_payments(purchase_id) where legacy;
create unique index if not exists purchase_payments_legacy_instant on public.purchase_payments(instant_pickup_id) where legacy;

-- Freeze the pre-installation baseline; never replay it into Home Cash.
lock table public.purchases, public.instant_pickups in share row exclusive mode;
insert into public.purchase_payments(purchase_id,amount,pickup_point,legacy,remitted_at,cash_recognized_amount)
select p.id,round(coalesce(p.paid_price,p.price,0),2),coalesce(p.pickup_point,'من البيت'),true,
  case when p.collected or public.home_cash_receipt_source(p.pickup_point)='home' then clock_timestamp() end,
  case when coalesce(r.received and not r.excluded_before_start,false) then r.recognized_amount else 0 end
from public.purchases p left join public.home_cash_receipts r on r.purchase_id=p.id
where not p.payment_model_ready and (p.picked_up or p.collected) on conflict do nothing;
insert into public.purchase_payments(instant_pickup_id,amount,pickup_point,legacy,remitted_at,cash_recognized_amount)
select p.id,p.price,p.pickup_point,true,
  case when p.collected or public.home_cash_receipt_source(p.pickup_point)='home' then clock_timestamp() end,
  case when coalesce(r.received and not r.excluded_before_start,false) then r.recognized_amount else 0 end
from public.instant_pickups p left join public.instant_pickup_cash_receipts r on r.instant_pickup_id=p.id
where not p.payment_model_ready and (p.picked_up or p.collected) on conflict do nothing;

-- Cash is now driven by payment/collection events, not by the parcel checkbox.
drop trigger if exists purchases_home_cash_receipt on public.purchases;
drop trigger if exists instant_pickups_home_cash on public.instant_pickups;
update public.purchases p set received_amount=coalesce((select sum(amount) from public.purchase_payments where purchase_id=p.id),0),
  remitted_amount=coalesce((select sum(amount) from public.purchase_payments where purchase_id=p.id and remitted_at is not null),0),
  payment_model_ready=true where not payment_model_ready;
update public.instant_pickups p set received_amount=coalesce((select sum(amount) from public.purchase_payments where instant_pickup_id=p.id),0),
  remitted_amount=coalesce((select sum(amount) from public.purchase_payments where instant_pickup_id=p.id and remitted_at is not null),0),
  payment_model_ready=true where not payment_model_ready;
alter table public.purchases alter column payment_model_ready set default true;
alter table public.instant_pickups alter column payment_model_ready set default true;

create or replace function public.can_read_purchase_payments(p_point text)
returns boolean language sql stable security definer set search_path='' as $$
  select coalesce((select case
    when r.role='admin' then true
    when r.role='viewer' and lower(u.email)='reem@she-store.com' then public.home_cash_receipt_source(p_point) in ('home','nablus','delivery')
    when r.role='viewer' and lower(u.email)='rawand@she-store.com' then public.home_cash_receipt_source(p_point) in ('home','delivery')
    when r.role='pickup' and lower(u.email)='maryamti@she-store.com' then public.home_cash_receipt_source(p_point)='maryamti'
    when r.role='pickup' and lower(u.email)='nablus@she-store.com' then public.home_cash_receipt_source(p_point)='nablus'
    else false end from public.user_roles r join auth.users u on u.id=r.user_id where r.user_id=auth.uid()),false);
$$;
create or replace function public.can_record_purchase_payments(p_point text)
returns boolean language sql stable security definer set search_path='' as $$
  select coalesce((select r.role='admin' or (r.role='viewer' and lower(u.email)='reem@she-store.com'
    and public.home_cash_receipt_source(p_point) in ('home','nablus','delivery'))
    from public.user_roles r join auth.users u on u.id=r.user_id where r.user_id=auth.uid()),false);
$$;
alter table public.purchase_payments enable row level security;
revoke all on public.purchase_payments from anon,authenticated;
grant select on public.purchase_payments to authenticated;
drop policy if exists payment_location_read on public.purchase_payments;
create policy payment_location_read on public.purchase_payments for select to authenticated
  using(public.can_read_purchase_payments(pickup_point));

create or replace function public.guard_purchase_payments()
returns trigger language plpgsql set search_path='' as $$
declare instant boolean := tg_table_name='instant_pickups'; agreed numeric; has_history boolean;
begin
  if tg_op='DELETE' then
    if exists(select 1 from public.purchase_payments where purchase_id=old.id or instant_pickup_id=old.id) then
      raise exception 'لا يمكن حذف مشترى له سجل دفعات';
    end if;
    return old;
  end if;
  if current_user in ('authenticated','anon') then
    if tg_op='INSERT' then
      if new.received_amount<>0 or new.remitted_amount<>0 or new.picked_up or new.collected then
        raise exception 'يجب تسجيل الاستلام من نافذة الدفعات'; end if;
    elsif new.received_amount is distinct from old.received_amount or new.remitted_amount is distinct from old.remitted_amount
      or new.picked_up is distinct from old.picked_up or new.picked_up_at is distinct from old.picked_up_at
      or new.payment_model_ready is distinct from old.payment_model_ready then
      raise exception 'يجب تسجيل الاستلام من نافذة الدفعات';
    end if;
  end if;
  if tg_op='UPDATE' then
    if new.picked_up is distinct from old.picked_up and new.received_amount=old.received_amount and new.received_amount>0 then
      raise exception 'للتراجع عن الاستلام يجب إلغاء الدفعات أولاً'; end if;
    if not instant and coalesce((to_jsonb(old)->>'ready_for_pickup')::boolean,false)
      and not coalesce((to_jsonb(new)->>'ready_for_pickup')::boolean,false)
      and exists(select 1 from public.purchase_payments where purchase_id=old.id) then
      raise exception 'لا يمكن إلغاء الجاهزية لمشترى له سجل دفعات'; end if;
    if new.pickup_point is distinct from old.pickup_point then
      select exists(select 1 from public.purchase_payments where purchase_id=old.id or instant_pickup_id=old.id) into has_history;
      if has_history then raise exception 'لا يمكن نقل مشترى بعد تسجيل دفعة'; end if;
    end if;
  end if;
  agreed := case when instant then (to_jsonb(new)->>'price')::numeric
    else coalesce((to_jsonb(new)->>'paid_price')::numeric,(to_jsonb(new)->>'price')::numeric,0) end;
  if new.received_amount<0 or new.remitted_amount<0 or new.remitted_amount>new.received_amount or new.received_amount>round(agreed,2) then
    raise exception 'السعر لا يمكن أن يكون أقل من الدفعات المسجلة'; end if;
  if new.picked_up and new.received_amount<=0 and agreed>0 then raise exception 'يجب تسجيل دفعة قبل الاستلام'; end if;
  new.payment_model_ready := true;
  new.collected := new.picked_up and new.received_amount>=round(agreed,2) and new.remitted_amount>=new.received_amount;
  new.collected_at := case when new.collected then coalesce(new.collected_at,clock_timestamp()) else null end;
  return new;
end;
$$;
drop trigger if exists purchases_payment_guard on public.purchases;
create trigger purchases_payment_guard before insert or update or delete on public.purchases
  for each row execute function public.guard_purchase_payments();
drop trigger if exists instant_payment_guard on public.instant_pickups;
create trigger instant_payment_guard before insert or update or delete on public.instant_pickups
  for each row execute function public.guard_purchase_payments();

-- Internal helpers use a fixed allowlist of table names, never user-provided SQL.
create or replace function public.payment_target(p_id uuid,p_instant boolean)
returns jsonb language plpgsql security definer set search_path='' as $$
declare target jsonb;
begin
  execute format('select to_jsonb(p) from public.%I p where id=$1 for update',
    case when p_instant then 'instant_pickups' else 'purchases' end) into target using p_id;
  if target is null then raise exception 'المشترى غير موجود'; end if;
  return target;
end;
$$;
create or replace function public.refresh_purchase_payment_totals(p_id uuid,p_instant boolean,p_receive boolean default false)
returns jsonb language plpgsql security definer set search_path='' as $$
declare paid numeric; remitted numeric; result jsonb;
begin
  select coalesce(sum(amount),0),coalesce(sum(amount) filter(where remitted_at is not null),0) into paid,remitted
    from public.purchase_payments where cancelled_at is null
    and (case when p_instant then instant_pickup_id=p_id else purchase_id=p_id end);
  execute format('update public.%I p set received_amount=$2,remitted_amount=$3,
    picked_up=case when $4 then true when $2=0 then false else picked_up end,
    picked_up_at=case when $4 then coalesce(picked_up_at,clock_timestamp()) when $2=0 then null else picked_up_at end
    where id=$1 returning to_jsonb(p)',case when p_instant then 'instant_pickups' else 'purchases' end)
    into result using p_id,paid,remitted,p_receive;
  return result;
end;
$$;
create or replace function public.payment_cash_entry(p_target jsonb,p_instant boolean,p_amount numeric,p_note text)
returns uuid language plpgsql security definer set search_path='' as $$
declare entry uuid; actor text; order_title text;
begin
  select email into actor from auth.users where id=auth.uid();
  if not p_instant then select order_name into order_title from public.orders where id=(p_target->>'order_id')::uuid; end if;
  insert into public.home_cash_transactions(kind,amount,source,purchase_id,instant_pickup_id,order_name,customer_name,note,created_by,actor_email)
    values(case when p_amount>0 then 'receipt' else 'receipt_adjustment' end,p_amount,
      public.home_cash_receipt_source(p_target->>'pickup_point'),
      case when not p_instant then (p_target->>'id')::uuid end,
      case when p_instant then (p_target->>'id')::uuid end,order_title,p_target->>'customer_name',p_note,auth.uid(),actor)
    returning id into entry;
  return entry;
end;
$$;

create or replace function public.get_purchase_payment_state(p_id uuid,p_instant boolean default false)
returns jsonb language plpgsql security definer set search_path='' as $$
declare target jsonb; entries jsonb;
begin
  -- No row locks for this read-only request.
  if p_instant is null then raise exception 'نوع المشترى غير صحيح'; end if;
  execute format('select to_jsonb(p) from public.%I p where id=$1',case when p_instant then 'instant_pickups' else 'purchases' end) into target using p_id;
  if target is null or not public.can_read_purchase_payments(target->>'pickup_point') then raise exception 'المشترى غير متاح'; end if;
  select coalesce(jsonb_agg(to_jsonb(p) order by created_at desc,id desc),'[]'::jsonb) into entries
    from public.purchase_payments p where case when p_instant then instant_pickup_id=p_id else purchase_id=p_id end;
  return jsonb_build_object('purchase',target,'payments',entries);
end;
$$;
create or replace function public.record_purchase_payment(p_request_id uuid,p_id uuid,p_amount numeric,p_receive boolean,p_instant boolean default false)
returns jsonb language plpgsql security definer set search_path='' as $$
declare target jsonb; prior public.purchase_payments%rowtype; agreed numeric; source text; cash_id uuid; changed boolean;
begin
  perform 1 from public.home_cash_account where id=1 for update;
  target := public.payment_target(p_id,p_instant);
  if not public.can_record_purchase_payments(target->>'pickup_point') then raise exception 'تسجيل الدفعات متاح لرهف وريم فقط ضمن أماكن الاستلام المسموحة'; end if;
  if p_request_id is null or p_instant is null or p_amount is null or p_amount<=0 or p_amount<>round(p_amount,2)
    or p_amount::text in ('NaN','Infinity','-Infinity') or p_receive is null then raise exception 'أدخلي مبلغاً صحيحاً أكبر من صفر بمنزلتين عشريتين'; end if;
  select * into prior from public.purchase_payments where id=p_request_id;
  if found then
    if prior.amount<>p_amount or prior.received_on_pickup<>p_receive or prior.created_by is distinct from auth.uid()
      or (case when p_instant then prior.instant_pickup_id is distinct from p_id else prior.purchase_id is distinct from p_id end)
      then raise exception 'معرّف الدفعة مستخدم بقيم مختلفة'; end if;
    return jsonb_build_object('purchase',target,'changed_receipt',false);
  end if;
  if (target->>'collected')::boolean then raise exception 'المشترى مكتمل الدفع والتحصيل'; end if;
  if not p_instant and not coalesce((target->>'ready_for_pickup')::boolean,false) then raise exception 'المشترى غير جاهز للاستلام'; end if;
  if not (target->>'picked_up')::boolean and not p_receive then raise exception 'يجب تسجيل أول دفعة مع الاستلام'; end if;
  agreed := case when p_instant then (target->>'price')::numeric else coalesce((target->>'paid_price')::numeric,(target->>'price')::numeric,0) end;
  if p_amount>round(agreed,2)-(target->>'received_amount')::numeric then raise exception 'الدفعة أكبر من المبلغ المتبقي'; end if;
  source := public.home_cash_receipt_source(target->>'pickup_point');
  if source is null then raise exception 'مكان الاستلام غير صحيح'; end if;
  if source='home' then cash_id := public.payment_cash_entry(target,p_instant,p_amount,'دفعة من الزبون - البيت'); end if;
  insert into public.purchase_payments(id,purchase_id,instant_pickup_id,amount,pickup_point,received_on_pickup,created_by,
    remitted_at,remitted_by,cash_recognized_amount,cash_transaction_id)
    values(p_request_id,case when not p_instant then p_id end,case when p_instant then p_id end,p_amount,target->>'pickup_point',p_receive,auth.uid(),
      case when source='home' then clock_timestamp() end,case when source='home' then auth.uid() end,
      case when source='home' then p_amount else 0 end,cash_id);
  changed := p_receive and not (target->>'picked_up')::boolean;
  target := public.refresh_purchase_payment_totals(p_id,p_instant,p_receive);
  return jsonb_build_object('purchase',target,'changed_receipt',changed);
end;
$$;
create or replace function public.collect_purchase_payments(p_ids uuid[],p_instant boolean default false)
returns numeric language plpgsql security definer set search_path='' as $$
declare item uuid; target jsonb; entry public.purchase_payments%rowtype; cash_id uuid; total numeric:=0;
begin
  if not coalesce(public.is_admin(),false) then raise exception 'التحصيل متاح لرهف فقط'; end if;
  if p_instant is null then raise exception 'نوع المشترى غير صحيح'; end if;
  if p_ids is null or cardinality(p_ids)=0 or array_position(p_ids,null) is not null then raise exception 'اختاري مشتريات للتحصيل'; end if;
  perform 1 from public.home_cash_account where id=1 for update;
  for item in select distinct x from unnest(p_ids) x order by x loop
    target := public.payment_target(item,p_instant);
    for entry in select * from public.purchase_payments where cancelled_at is null and remitted_at is null
      and (case when p_instant then instant_pickup_id=item else purchase_id=item end) order by id for update loop
      cash_id := public.payment_cash_entry(target,p_instant,entry.amount,'تحصيل دفعة من نقطة الاستلام أو التوصيل');
      update public.purchase_payments set remitted_at=clock_timestamp(),remitted_by=auth.uid(),
        cash_recognized_amount=amount,cash_transaction_id=cash_id where id=entry.id;
      total := total+entry.amount;
    end loop;
    perform public.refresh_purchase_payment_totals(item,p_instant);
  end loop;
  return total;
end;
$$;
create or replace function public.cancel_purchase_payment(p_payment_id uuid,p_reason text)
returns jsonb language plpgsql security definer set search_path='' as $$
declare entry public.purchase_payments%rowtype; target jsonb; instant boolean; target_id uuid; cash_id uuid;
begin
  if not coalesce(public.is_admin(),false) then raise exception 'إلغاء الدفعات متاح لرهف فقط'; end if;
  if p_reason is null or length(btrim(p_reason)) not between 1 and 500 then raise exception 'أدخلي سبب الإلغاء'; end if;
  perform 1 from public.home_cash_account where id=1 for update;
  select * into entry from public.purchase_payments where id=p_payment_id for update;
  if not found then raise exception 'الدفعة غير موجودة'; end if;
  instant := entry.instant_pickup_id is not null; target_id := coalesce(entry.instant_pickup_id,entry.purchase_id);
  target := public.payment_target(target_id,instant);
  if entry.cancelled_at is not null then return target; end if;
  if entry.cash_recognized_amount>0 then
    cash_id := public.payment_cash_entry(target,instant,-entry.cash_recognized_amount,'إلغاء دفعة: '||btrim(p_reason));
  end if;
  update public.purchase_payments set cancelled_at=clock_timestamp(),cancelled_by=auth.uid(),
    cancellation_reason=btrim(p_reason),cancellation_cash_id=cash_id where id=entry.id;
  return public.refresh_purchase_payment_totals(target_id,instant);
end;
$$;

revoke all on function public.can_read_purchase_payments(text),public.can_record_purchase_payments(text),
  public.payment_target(uuid,boolean),public.refresh_purchase_payment_totals(uuid,boolean,boolean),
  public.payment_cash_entry(jsonb,boolean,numeric,text),public.guard_purchase_payments(),
  public.get_purchase_payment_state(uuid,boolean),public.record_purchase_payment(uuid,uuid,numeric,boolean,boolean),
  public.collect_purchase_payments(uuid[],boolean),public.cancel_purchase_payment(uuid,text) from public,anon,authenticated;
grant execute on function public.can_read_purchase_payments(text),public.get_purchase_payment_state(uuid,boolean),
  public.record_purchase_payment(uuid,uuid,numeric,boolean,boolean),public.collect_purchase_payments(uuid[],boolean),
  public.cancel_purchase_payment(uuid,text) to authenticated;

-- Recompute completion without posting money; price and profit columns stay unchanged.
update public.purchases set payment_model_ready=true;
update public.instant_pickups set payment_model_ready=true;
notify pgrst,'reload schema';
commit;
