begin;

-- Home Cash must be installed first; order-backed purchases remain unchanged.
do $$ begin
  if to_regclass('public.home_cash_transactions') is null then
    raise exception 'Run the Home Cash migration before this migration.';
  end if;
end $$;

create table if not exists public.instant_pickups (
  id uuid primary key default gen_random_uuid(),
  customer_name text not null check (length(btrim(customer_name)) between 1 and 200),
  price numeric(14, 2) not null check (price >= 0 and price::text not in ('NaN', 'Infinity', '-Infinity')),
  pickup_point text not null check (pickup_point in ('من البيت', 'من نقطة الاستلام', 'نقطة استلام نابلس', 'توصيل')),
  picked_up boolean not null default false,
  picked_up_at timestamptz,
  collected boolean not null default false,
  collected_at timestamptz,
  created_at timestamptz not null default clock_timestamp(),
  created_by uuid references auth.users(id) on delete set null,
  check (not collected or picked_up)
);
create index if not exists instant_pickups_pending_idx
  on public.instant_pickups (pickup_point, created_at desc, id) where not collected;

alter table public.instant_pickups enable row level security;
revoke all on public.instant_pickups from anon, authenticated;
grant select on public.instant_pickups to authenticated;

create or replace function public.can_access_instant_pickup(p_pickup_point text)
returns boolean language sql stable security definer set search_path = '' as $$
  select coalesce((select case
    when r.role = 'admin' then true
    when r.role = 'viewer' then p_pickup_point in ('من البيت', 'توصيل')
      or (lower(u.email) not like '%rawand%' and p_pickup_point = 'نقطة استلام نابلس')
    when r.role = 'pickup' and lower(u.email) = 'nablus@she-store.com' then p_pickup_point = 'نقطة استلام نابلس'
    when r.role = 'pickup' and lower(u.email) = 'maryamti@she-store.com' then p_pickup_point = 'من نقطة الاستلام'
    else false end
  from public.user_roles r join auth.users u on u.id = r.user_id
  where r.user_id = auth.uid()), false);
$$;
revoke all on function public.can_access_instant_pickup(text) from public, anon, authenticated;
grant execute on function public.can_access_instant_pickup(text) to authenticated;
drop policy if exists instant_pickups_location_read on public.instant_pickups;
create policy instant_pickups_location_read on public.instant_pickups
  for select to authenticated using (public.can_access_instant_pickup(pickup_point));

alter table public.home_cash_transactions add column if not exists instant_pickup_id uuid
  references public.instant_pickups(id) on delete set null;

create table if not exists public.instant_pickup_cash_receipts (
  instant_pickup_id uuid primary key,
  received boolean not null default false,
  recognized_amount numeric(14, 2) not null default 0,
  source text,
  excluded_before_start boolean not null default false
);
alter table public.instant_pickup_cash_receipts enable row level security;
revoke all on public.instant_pickup_cash_receipts from anon, authenticated;

lock table public.instant_pickups in share row exclusive mode;
insert into public.instant_pickup_cash_receipts
  (instant_pickup_id, received, recognized_amount, source, excluded_before_start)
select p.id, eligible.received, case when eligible.received then p.price else 0 end,
  case when eligible.received then location.source else null end, eligible.received
from public.instant_pickups p
cross join lateral (select public.home_cash_receipt_source(p.pickup_point) as source) location
cross join lateral (select case when location.source = 'home' then p.picked_up or p.collected
  else p.collected end as received) eligible
on conflict (instant_pickup_id) do nothing;

create or replace function public.sync_instant_pickup_home_cash()
returns trigger language plpgsql security definer set search_path = '' as $$
declare
  previous public.instant_pickup_cash_receipts%rowtype;
  next_source text;
  next_received boolean;
  next_amount numeric(14, 2);
  change_amount numeric(14, 2);
  event_note text;
  cash_actor_email text;
begin
  perform 1 from public.home_cash_account where id = 1 for update;
  if not found then raise exception 'تتبع مصاري البيت غير متاح'; end if;
  insert into public.instant_pickup_cash_receipts (instant_pickup_id) values (new.id) on conflict do nothing;
  select * into previous from public.instant_pickup_cash_receipts where instant_pickup_id = new.id;
  -- A parcel transfer does not relocate cash already received.
  next_source := case when previous.received then previous.source else public.home_cash_receipt_source(new.pickup_point) end;
  next_received := case when next_source = 'home' then new.picked_up or new.collected else new.collected end;
  next_amount := case when next_received then new.price else 0 end;
  change_amount := next_amount - previous.recognized_amount;
  update public.instant_pickup_cash_receipts set received = next_received, recognized_amount = next_amount,
    source = case when next_received then next_source else null end,
    excluded_before_start = previous.excluded_before_start and next_received
  where instant_pickup_id = new.id;
  if previous.excluded_before_start then return new; end if;
  if change_amount = 0 and previous.received = next_received then
    if tg_op <> 'UPDATE' then return new; end if;
    if not (previous.received and previous.source = 'home' and new.collected is distinct from old.collected) then return new; end if;
  end if;
  event_note := case when not previous.received and next_received then 'استلام أو تحصيل مستلم فوري'
    when previous.received and not next_received then 'تراجع عن استلام أو تحصيل مستلم فوري'
    when change_amount <> 0 then 'تعديل مبلغ مستلم فوري'
    else 'تأكيد تحصيل مستلم فوري من البيت' end;
  select email into cash_actor_email from auth.users where id = auth.uid();
  insert into public.home_cash_transactions
    (kind, amount, source, instant_pickup_id, order_name, customer_name, note, created_by, actor_email)
  values (case when not previous.received and next_received then 'receipt' else 'receipt_adjustment' end,
    change_amount, coalesce(next_source, previous.source), new.id, 'استلام فوري', new.customer_name,
    event_note, auth.uid(), cash_actor_email);
  return new;
end;
$$;
revoke all on function public.sync_instant_pickup_home_cash() from public, anon, authenticated;
drop trigger if exists instant_pickups_home_cash on public.instant_pickups;
create trigger instant_pickups_home_cash after insert or update of picked_up, collected, price
  on public.instant_pickups for each row execute function public.sync_instant_pickup_home_cash();

create or replace function public.create_instant_pickup(
  p_request_id uuid, p_customer_name text, p_price numeric, p_pickup_point text
)
returns public.instant_pickups language plpgsql security definer set search_path = '' as $$
declare
  saved public.instant_pickups%rowtype;
begin
  if not coalesce(public.is_admin(), false) then raise exception 'لا توجد صلاحية' using errcode = '42501'; end if;
  if p_request_id is null or p_customer_name is null or length(btrim(p_customer_name)) not between 1 and 200
    or p_price is null or p_price < 0 or p_price <> round(p_price, 2)
    or p_price::text in ('NaN', 'Infinity', '-Infinity')
    or p_pickup_point is null or p_pickup_point not in ('من البيت', 'من نقطة الاستلام', 'نقطة استلام نابلس', 'توصيل')
  then raise exception 'بيانات المستلم غير صحيحة'; end if;
  insert into public.instant_pickups (id, customer_name, price, pickup_point, created_by)
  values (p_request_id, btrim(p_customer_name), p_price, p_pickup_point, auth.uid())
  on conflict (id) do nothing returning * into saved;
  if not found then
    select * into saved from public.instant_pickups where id = p_request_id;
    if saved.customer_name is distinct from btrim(p_customer_name) or saved.price is distinct from p_price
      or saved.pickup_point is distinct from p_pickup_point then raise exception 'هذا المستلم محفوظ مسبقاً بقيم مختلفة'; end if;
  end if;
  return saved;
end;
$$;

create or replace function public.update_instant_pickup(p_id uuid, p_action text, p_pickup_point text default null)
returns public.instant_pickups language plpgsql security definer set search_path = '' as $$
declare
  saved public.instant_pickups%rowtype;
  admin_user boolean := coalesce(public.is_admin(), false);
  pickup_user boolean;
begin
  select * into saved from public.instant_pickups where id = p_id for update;
  if not found or not public.can_access_instant_pickup(saved.pickup_point) then
    raise exception 'المستلم غير متاح' using errcode = '42501';
  end if;
  select exists (select 1 from public.user_roles where user_id = auth.uid() and role = 'pickup') into pickup_user;
  if p_action in ('receive', 'undo_receive') then
    if not (admin_user or pickup_user) then raise exception 'لا توجد صلاحية' using errcode = '42501'; end if;
    if saved.collected then raise exception 'لا يمكن تعديل مستلم تم تحصيله'; end if;
    update public.instant_pickups set picked_up = (p_action = 'receive'),
      picked_up_at = case when p_action = 'receive' then coalesce(picked_up_at, clock_timestamp()) else null end
    where id = p_id returning * into saved;
  elsif p_action = 'collect' then
    if not admin_user then raise exception 'لا توجد صلاحية' using errcode = '42501'; end if;
    if not saved.picked_up then raise exception 'يجب استلام المشترى قبل التحصيل'; end if;
    if saved.collected then return saved; end if;
    update public.instant_pickups set collected = true, collected_at = clock_timestamp()
    where id = p_id returning * into saved;
  elsif p_action = 'transfer' then
    if not admin_user then raise exception 'لا توجد صلاحية' using errcode = '42501'; end if;
    if saved.collected then raise exception 'لا يمكن نقل مستلم تم تحصيله'; end if;
    if p_pickup_point is null or p_pickup_point not in ('من البيت', 'من نقطة الاستلام', 'نقطة استلام نابلس', 'توصيل') then
      raise exception 'نقطة الاستلام غير صحيحة';
    end if;
    update public.instant_pickups set pickup_point = p_pickup_point where id = p_id returning * into saved;
  else raise exception 'الإجراء غير صحيح'; end if;
  return saved;
end;
$$;

create or replace function public.collect_instant_pickups(p_ids uuid[])
returns integer language plpgsql security definer set search_path = '' as $$
declare
  requested integer;
  eligible integer;
  changed integer;
begin
  if not coalesce(public.is_admin(), false) then raise exception 'لا توجد صلاحية' using errcode = '42501'; end if;
  select count(distinct item) into requested from unnest(p_ids) item;
  if requested = 0 or array_position(p_ids, null) is not null then raise exception 'لا توجد مشتريات للتحصيل'; end if;
  perform 1 from public.instant_pickups where id = any(p_ids) order by id for update;
  select count(*) into eligible from public.instant_pickups where id = any(p_ids) and picked_up;
  if eligible <> requested then raise exception 'يجب استلام المشتريات قبل التحصيل'; end if;
  update public.instant_pickups set collected = true, collected_at = clock_timestamp()
  where id = any(p_ids) and not collected;
  get diagnostics changed = row_count;
  return changed;
end;
$$;

revoke all on function public.create_instant_pickup(uuid, text, numeric, text),
  public.update_instant_pickup(uuid, text, text), public.collect_instant_pickups(uuid[])
  from public, anon, authenticated;
grant execute on function public.create_instant_pickup(uuid, text, numeric, text),
  public.update_instant_pickup(uuid, text, text), public.collect_instant_pickups(uuid[])
  to authenticated;

notify pgrst, 'reload schema';
commit;
