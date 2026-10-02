begin;

create table if not exists public.home_cash_account (
  id smallint primary key check (id = 1),
  started_at timestamptz not null default clock_timestamp()
);

create table if not exists public.home_cash_receipts (
  purchase_id uuid primary key,
  received boolean not null default false,
  recognized_amount numeric(14, 2) not null default 0,
  source text,
  excluded_before_start boolean not null default false
);

create table if not exists public.home_cash_transactions (
  id uuid primary key default gen_random_uuid(),
  kind text not null check (kind in ('tracking_started', 'receipt', 'receipt_adjustment', 'expense', 'expense_reversal')),
  amount numeric(14, 2) not null,
  category text check (category in ('bags', 'pins', 'name_stickers', 'postal_delivery', 'atm_deposit')),
  quantity integer check (quantity > 0),
  source text,
  purchase_id uuid references public.purchases(id) on delete set null,
  order_name text,
  customer_name text,
  note text,
  reverses_id uuid unique references public.home_cash_transactions(id),
  created_by uuid references auth.users(id) on delete set null,
  actor_email text,
  created_at timestamptz not null default clock_timestamp(),
  check (kind <> 'expense' or (amount < 0 and category is not null and quantity is not null)),
  check (kind <> 'expense_reversal' or (amount > 0 and reverses_id is not null))
);

create index if not exists home_cash_transactions_created_idx
  on public.home_cash_transactions (created_at desc, id desc);
create unique index if not exists home_cash_single_start_idx
  on public.home_cash_transactions (kind) where kind = 'tracking_started';

alter table public.home_cash_account enable row level security;
alter table public.home_cash_receipts enable row level security;
alter table public.home_cash_transactions enable row level security;

drop policy if exists home_cash_account_admin_read on public.home_cash_account;
create policy home_cash_account_admin_read on public.home_cash_account
  for select to authenticated using (public.is_admin());
drop policy if exists home_cash_transactions_admin_read on public.home_cash_transactions;
create policy home_cash_transactions_admin_read on public.home_cash_transactions
  for select to authenticated using (public.is_admin());

revoke all on public.home_cash_account, public.home_cash_receipts, public.home_cash_transactions from anon, authenticated;
grant select on public.home_cash_account, public.home_cash_transactions to authenticated;

create or replace function public.home_cash_receipt_source(p_pickup text)
returns text language sql immutable set search_path = '' as $$
  select case btrim(p_pickup)
    when 'من البيت' then 'home'
    when 'توصيل' then 'delivery'
    when 'نقطة استلام نابلس' then 'nablus'
    when 'من نقطة الاستلام' then 'maryamti'
    when 'من نقطة الاستلام - مريمتي' then 'maryamti'
    when 'من نقطة الاستلام - La Aura' then 'maryamti'
    else null
  end;
$$;
revoke all on function public.home_cash_receipt_source(text) from public, anon, authenticated;

insert into public.home_cash_account (id) values (1) on conflict (id) do nothing;
insert into public.home_cash_transactions (kind, amount, note)
values ('tracking_started', 0, 'بدء تتبع مصاري البيت من صفر')
on conflict do nothing;

-- Existing receipts form a baseline, without adding historical money to the new balance.
lock table public.purchases in share row exclusive mode;
insert into public.home_cash_receipts (purchase_id, received, recognized_amount, source, excluded_before_start)
select p.id, eligible.received,
  case when eligible.received then round(coalesce(p.paid_price, p.price, 0), 2) else 0 end,
  case when eligible.received then location.source else null end,
  eligible.received
from public.purchases p
cross join lateral (select public.home_cash_receipt_source(p.pickup_point) as source) location
cross join lateral (select coalesce(case when location.source = 'home'
  then p.picked_up or p.collected else location.source is not null and p.collected end, false) as received) eligible
on conflict (purchase_id) do nothing;

create or replace function public.sync_home_cash_receipt()
returns trigger language plpgsql security definer set search_path = '' as $$
declare
  previous public.home_cash_receipts%rowtype;
  next_source text;
  next_received boolean;
  next_amount numeric(14, 2);
  change_amount numeric(14, 2);
  event_note text;
  cash_order_name text;
  cash_actor_email text;
begin
  perform 1 from public.home_cash_account where id = 1 for update;
  if not found then return new; end if;

  insert into public.home_cash_receipts (purchase_id) values (new.id) on conflict do nothing;
  select * into previous from public.home_cash_receipts where purchase_id = new.id;

  -- Transfers change the parcel location, not cash already received at home.
  next_source := case when previous.received then previous.source
    else public.home_cash_receipt_source(new.pickup_point) end;
  next_received := coalesce(case when next_source = 'home'
    then new.picked_up or new.collected else next_source is not null and new.collected end, false);
  next_amount := case when next_received then round(coalesce(new.paid_price, new.price, 0), 2) else 0 end;
  change_amount := next_amount - previous.recognized_amount;

  update public.home_cash_receipts set
    received = next_received,
    recognized_amount = next_amount,
    source = case when next_received then next_source else null end,
    excluded_before_start = previous.excluded_before_start and next_received
  where purchase_id = new.id;

  if previous.excluded_before_start then return new; end if;
  if change_amount = 0 and previous.received = next_received then
    if tg_op <> 'UPDATE' then return new; end if;
    if not (previous.received and previous.source = 'home' and new.collected is distinct from old.collected) then
      return new;
    end if;
  end if;

  event_note := case
    when not previous.received and next_received and next_source = 'home' then 'استلام مشترى من البيت'
    when not previous.received and next_received then 'تحصيل من نقطة الاستلام أو التوصيل'
    when previous.received and not next_received then 'تراجع عن استلام أو تحصيل'
    when change_amount <> 0 then 'تعديل المبلغ المدفوع'
    else 'تأكيد أو تراجع عن تحصيل مشترى مستلم من البيت'
  end;
  select order_name into cash_order_name from public.orders where id = new.order_id;
  select email into cash_actor_email from auth.users where id = auth.uid();
  insert into public.home_cash_transactions
    (kind, amount, source, purchase_id, order_name, customer_name, note, created_by, actor_email)
  values (case when not previous.received and next_received then 'receipt' else 'receipt_adjustment' end,
    change_amount, coalesce(next_source, previous.source), new.id, cash_order_name, new.customer_name,
    event_note, auth.uid(), cash_actor_email);
  return new;
end;
$$;
revoke all on function public.sync_home_cash_receipt() from public, anon, authenticated;
drop trigger if exists purchases_home_cash_receipt on public.purchases;
create trigger purchases_home_cash_receipt after insert or update of picked_up, collected, paid_price, price
  on public.purchases for each row execute function public.sync_home_cash_receipt();

create or replace function public.get_home_cash_summary()
returns jsonb language sql stable security invoker set search_path = '' as $$
  select jsonb_build_object(
    'started_at', a.started_at,
    'balance', coalesce(sum(t.amount), 0),
    'income', coalesce(sum(t.amount) filter (where t.kind in ('receipt', 'receipt_adjustment')), 0),
    'expenses', -coalesce(sum(t.amount) filter (where t.kind in ('expense', 'expense_reversal')), 0)
  ) from public.home_cash_account a left join public.home_cash_transactions t on true
  where a.id = 1 group by a.started_at;
$$;

create or replace function public.get_home_cash_activity(p_offset integer default 0, p_limit integer default 26)
returns table (id uuid, kind text, amount numeric, category text, quantity integer, source text,
  order_name text, customer_name text, note text, actor_email text, created_at timestamptz,
  balance_after numeric, is_voided boolean)
language sql stable security invoker set search_path = '' as $$
  select t.id, t.kind, t.amount, t.category, t.quantity, t.source, t.order_name, t.customer_name,
    t.note, t.actor_email, t.created_at,
    sum(t.amount) over (order by t.created_at, t.id rows unbounded preceding),
    exists (select 1 from public.home_cash_transactions r where r.reverses_id = t.id)
  from public.home_cash_transactions t
  order by t.created_at desc, t.id desc
  limit least(greatest(p_limit, 1), 100) offset greatest(p_offset, 0);
$$;

create or replace function public.record_home_cash_expense(
  p_request_id uuid, p_category text, p_amount numeric, p_quantity integer, p_note text default null
)
returns public.home_cash_transactions language plpgsql security definer set search_path = '' as $$
declare
  existing public.home_cash_transactions%rowtype;
  saved public.home_cash_transactions%rowtype;
  available numeric;
  cash_actor_email text;
begin
  if not coalesce(public.is_admin(), false) then raise exception 'لا توجد صلاحية' using errcode = '42501'; end if;
  if p_request_id is null or p_amount is null or p_amount <= 0 or p_amount <> round(p_amount, 2)
    or p_amount::text in ('NaN', 'Infinity', '-Infinity') or p_quantity is null or p_quantity <= 0
    or p_category is null or p_category not in ('bags', 'pins', 'name_stickers', 'postal_delivery', 'atm_deposit')
    or length(coalesce(p_note, '')) > 500 then raise exception 'بيانات المصروف غير صحيحة'; end if;

  perform 1 from public.home_cash_account where id = 1 for update;
  if not found then raise exception 'تتبع مصاري البيت غير متاح'; end if;
  select * into existing from public.home_cash_transactions where id = p_request_id;
  if found then
    if existing.kind = 'expense' and existing.amount = -p_amount and existing.category = p_category
      and existing.quantity = p_quantity and coalesce(existing.note, '') = coalesce(btrim(p_note), '') then return existing; end if;
    raise exception 'هذا المصروف محفوظ مسبقاً بقيم مختلفة';
  end if;
  select coalesce(sum(amount), 0) into available from public.home_cash_transactions;
  if available < p_amount then raise exception 'المبلغ أكبر من الرصيد المتاح في البيت'; end if;
  select email into cash_actor_email from auth.users where id = auth.uid();
  insert into public.home_cash_transactions (id, kind, amount, category, quantity, note, created_by, actor_email)
  values (p_request_id, 'expense', -p_amount, p_category, p_quantity, nullif(btrim(p_note), ''), auth.uid(), cash_actor_email)
  returning * into saved;
  return saved;
end;
$$;

create or replace function public.cancel_home_cash_expense(p_expense_id uuid)
returns public.home_cash_transactions language plpgsql security definer set search_path = '' as $$
declare
  original public.home_cash_transactions%rowtype;
  saved public.home_cash_transactions%rowtype;
  cash_actor_email text;
begin
  if not coalesce(public.is_admin(), false) then raise exception 'لا توجد صلاحية' using errcode = '42501'; end if;
  perform 1 from public.home_cash_account where id = 1 for update;
  select * into original from public.home_cash_transactions where id = p_expense_id and kind = 'expense';
  if not found then raise exception 'المصروف غير موجود'; end if;
  select * into saved from public.home_cash_transactions where reverses_id = p_expense_id;
  if found then return saved; end if;
  select email into cash_actor_email from auth.users where id = auth.uid();
  insert into public.home_cash_transactions (kind, amount, category, quantity, note, reverses_id, created_by, actor_email)
  values ('expense_reversal', -original.amount, original.category, original.quantity, 'إلغاء مصروف أو إيداع',
    original.id, auth.uid(), cash_actor_email) returning * into saved;
  return saved;
end;
$$;

revoke all on function public.get_home_cash_summary(), public.get_home_cash_activity(integer, integer),
  public.record_home_cash_expense(uuid, text, numeric, integer, text), public.cancel_home_cash_expense(uuid)
  from public, anon, authenticated;
grant execute on function public.get_home_cash_summary(), public.get_home_cash_activity(integer, integer),
  public.record_home_cash_expense(uuid, text, numeric, integer, text), public.cancel_home_cash_expense(uuid)
  to authenticated;

notify pgrst, 'reload schema';
commit;
