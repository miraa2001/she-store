begin;

alter table public.orders
  add column if not exists marketing_fee numeric(14,2) not null default 0,
  add column if not exists home_profit_percent numeric(5,2),
  add column if not exists rahaf_profit_percent numeric(5,2),
  add column if not exists mira_profit_percent numeric(5,2);
alter table public.orders drop constraint if exists orders_marketing_fee_check;
alter table public.orders add constraint orders_marketing_fee_check
  check (marketing_fee >= 0 and marketing_fee::text not in ('NaN','Infinity','-Infinity'));
alter table public.orders drop constraint if exists orders_profit_percent_check;
alter table public.orders add constraint orders_profit_percent_check check (
  (home_profit_percent is null and rahaf_profit_percent is null and mira_profit_percent is null)
  or (home_profit_percent is not null and rahaf_profit_percent is not null and mira_profit_percent is not null
    and home_profit_percent between 0 and 100 and rahaf_profit_percent between 0 and 100
    and mira_profit_percent between 0 and 100
    and home_profit_percent + rahaf_profit_percent + mira_profit_percent = 100)
);

create table if not exists public.profit_tracking_state (
  id smallint primary key check (id = 1), started_at timestamptz not null default clock_timestamp()
);
-- Only the first installation clears legacy manual profits, never a migration rerun.
do $$ begin
  insert into public.profit_tracking_state (id) values (1) on conflict do nothing;
  if found then
    update public.orders set total_profit = 0, mira_profit = 0, rahaf_profit = 0;
  end if;
end $$;

-- No order FK: deleting an order must not erase earnings that were already paid.
create table if not exists public.order_profit_accruals (
  order_id uuid primary key, order_name text not null,
  configured boolean not null, purchase_value numeric(14,2) not null,
  spent numeric(14,2) not null, postal_fee numeric(14,2) not null,
  marketing_fee numeric(14,2) not null, distributable_profit numeric(14,2) not null,
  home_percent numeric(5,2), rahaf_percent numeric(5,2), mira_percent numeric(5,2),
  home_profit numeric(14,2) not null, rahaf_profit numeric(14,2) not null, mira_profit numeric(14,2) not null,
  updated_at timestamptz not null default clock_timestamp()
);
create table if not exists public.profit_payouts (
  id uuid primary key, party text not null check (party in ('home','rahaf','mira')),
  amount numeric(14,2) not null check (amount > 0 and amount::text not in ('NaN','Infinity','-Infinity')),
  deduct_home_cash boolean not null,
  cash_transaction_id uuid unique references public.home_cash_transactions(id),
  created_by uuid references auth.users(id) on delete set null,
  actor_email text, created_at timestamptz not null default clock_timestamp(),
  check (deduct_home_cash = (cash_transaction_id is not null))
);
create index if not exists profit_payouts_created_idx on public.profit_payouts(created_at desc,id desc);

alter table public.profit_tracking_state enable row level security;
alter table public.order_profit_accruals enable row level security;
alter table public.profit_payouts enable row level security;
revoke all on public.profit_tracking_state, public.order_profit_accruals, public.profit_payouts from anon, authenticated;
grant select on public.order_profit_accruals, public.profit_payouts to authenticated;
drop policy if exists profit_accruals_admin_read on public.order_profit_accruals;
create policy profit_accruals_admin_read on public.order_profit_accruals for select to authenticated using (public.is_admin());
drop policy if exists profit_payouts_admin_read on public.profit_payouts;
create policy profit_payouts_admin_read on public.profit_payouts for select to authenticated using (public.is_admin());

alter table public.home_cash_transactions drop constraint if exists home_cash_transactions_kind_check;
alter table public.home_cash_transactions add constraint home_cash_transactions_kind_check
  check (kind in ('tracking_started','receipt','receipt_adjustment','expense','expense_reversal','profit_payout'));

create or replace function public.refresh_order_profit(p_order_id uuid)
returns void language plpgsql security definer set search_path = '' as $$
declare
  o public.orders%rowtype;
  purchase_total numeric;
  net_profit numeric;
  rahaf_amount numeric := 0;
  mira_amount numeric := 0;
  home_amount numeric := 0;
begin
  -- The same lock as the cash ledger serializes payouts, expenses, and earnings changes.
  perform 1 from public.home_cash_account where id = 1 for update;
  select * into o from public.orders where id = p_order_id;
  if not found then return; end if;
  select round(coalesce(sum(coalesce(paid_price,price,0)),0),2) into purchase_total
    from public.purchases where order_id = p_order_id;
  net_profit := greatest(0, purchase_total + round(coalesce(o.postal_fee,0),2) - round(coalesce(o.spent_amount,0),2) - o.marketing_fee);
  if o.home_profit_percent is not null then
    -- Largest-remainder allocation keeps the three rounded shares equal to net profit.
    with shares(party, exact_cents, priority) as (
      values ('home',net_profit*o.home_profit_percent,1),
        ('rahaf',net_profit*o.rahaf_profit_percent,2), ('mira',net_profit*o.mira_profit_percent,3)
    ), ranked as (
      select *, row_number() over(order by exact_cents-floor(exact_cents) desc,priority) as rank,
        net_profit*100 - sum(floor(exact_cents)) over() as remaining from shares
    ), amounts as (
      select party,(floor(exact_cents) + case when rank <= remaining then 1 else 0 end)/100 as amount from ranked
    ) select max(amount) filter(where party='home'),max(amount) filter(where party='rahaf'),
      max(amount) filter(where party='mira') into home_amount,rahaf_amount,mira_amount from amounts;
  end if;
  insert into public.order_profit_accruals
    (order_id,order_name,configured,purchase_value,spent,postal_fee,marketing_fee,distributable_profit,
      home_percent,rahaf_percent,mira_percent,home_profit,rahaf_profit,mira_profit)
  values (o.id,o.order_name,o.home_profit_percent is not null,purchase_total,coalesce(o.spent_amount,0),
    coalesce(o.postal_fee,0),o.marketing_fee,net_profit,o.home_profit_percent,o.rahaf_profit_percent,
    o.mira_profit_percent,home_amount,rahaf_amount,mira_amount)
  on conflict(order_id) do update set order_name=excluded.order_name,configured=excluded.configured,
    purchase_value=excluded.purchase_value,spent=excluded.spent,postal_fee=excluded.postal_fee,
    marketing_fee=excluded.marketing_fee,distributable_profit=excluded.distributable_profit,
    home_percent=excluded.home_percent,rahaf_percent=excluded.rahaf_percent,mira_percent=excluded.mira_percent,
    home_profit=excluded.home_profit,rahaf_profit=excluded.rahaf_profit,mira_profit=excluded.mira_profit,
    updated_at=clock_timestamp();
end;
$$;

create or replace function public.sync_order_profit()
returns trigger language plpgsql security definer set search_path = '' as $$
begin
  if tg_table_name = 'orders' then
    perform public.refresh_order_profit(new.id);
  else
    if tg_op <> 'INSERT' then perform public.refresh_order_profit(old.order_id); end if;
    if tg_op <> 'DELETE' then
      if tg_op = 'INSERT' then perform public.refresh_order_profit(new.order_id);
      elsif new.order_id is distinct from old.order_id then perform public.refresh_order_profit(new.order_id); end if;
    end if;
  end if;
  return null;
end;
$$;
revoke all on function public.refresh_order_profit(uuid),public.sync_order_profit() from public,anon,authenticated;

drop trigger if exists orders_profit_accrual on public.orders;
create trigger orders_profit_accrual after insert or update of order_name,spent_amount,postal_fee,marketing_fee,
  home_profit_percent,rahaf_profit_percent,mira_profit_percent on public.orders
  for each row execute function public.sync_order_profit();
drop trigger if exists purchases_profit_accrual on public.purchases;
create trigger purchases_profit_accrual after insert or delete or update of price,paid_price,order_id
  on public.purchases for each row execute function public.sync_order_profit();

do $$ declare item record; begin
  for item in select id from public.orders loop perform public.refresh_order_profit(item.id); end loop;
end $$;

create or replace function public.get_profit_balances()
returns table(party text,earned numeric,paid numeric,balance numeric)
language sql stable security invoker set search_path = '' as $$
  with earnings as (
    select coalesce(sum(home_profit),0) as home,coalesce(sum(rahaf_profit),0) as rahaf,
      coalesce(sum(mira_profit),0) as mira from public.order_profit_accruals
  ), parties as (
    select 'home' as party,home as earned from earnings union all
    select 'rahaf',rahaf from earnings union all select 'mira',mira from earnings
  ) select p.party,p.earned,coalesce(sum(t.amount),0),p.earned-coalesce(sum(t.amount),0)
  from parties p left join public.profit_payouts t on t.party=p.party group by p.party,p.earned;
$$;

create or replace function public.record_profit_payout(p_request_id uuid,p_party text,p_amount numeric,p_deduct_home_cash boolean)
returns public.profit_payouts language plpgsql security definer set search_path = '' as $$
declare
  saved public.profit_payouts%rowtype;
  available numeric;
  cash_available numeric;
  cash_id uuid;
  email text;
begin
  if not coalesce(public.is_admin(),false) then raise exception 'لا توجد صلاحية' using errcode='42501'; end if;
  if p_request_id is null or p_party is null or p_party not in ('home','rahaf','mira') or p_amount is null
    or p_amount <= 0 or p_amount <> round(p_amount,2) or p_amount::text in ('NaN','Infinity','-Infinity')
    or p_deduct_home_cash is null then raise exception 'بيانات التسليم غير صحيحة'; end if;
  perform 1 from public.home_cash_account where id=1 for update;
  if not found then raise exception 'تتبع مصاري البيت غير متاح'; end if;
  select * into saved from public.profit_payouts where id=p_request_id;
  if found then
    if saved.party=p_party and saved.amount=p_amount and saved.deduct_home_cash=p_deduct_home_cash then return saved; end if;
    raise exception 'هذا التسليم محفوظ مسبقاً بقيم مختلفة';
  end if;
  select balance into available from public.get_profit_balances() where party=p_party;
  if p_amount > available then raise exception 'المبلغ أكبر من رصيد أرباح الطرف'; end if;
  select u.email into email from auth.users u where u.id=auth.uid();
  if p_deduct_home_cash then
    select coalesce(sum(amount),0) into cash_available from public.home_cash_transactions;
    if p_amount > cash_available then raise exception 'المبلغ أكبر من الرصيد المتاح في البيت'; end if;
    insert into public.home_cash_transactions(kind,amount,source,customer_name,note,created_by,actor_email)
    values ('profit_payout',-p_amount,'home',case p_party when 'home' then 'البيت' when 'rahaf' then 'رهف' else 'ميرا' end,
      'تسليم أرباح',auth.uid(),email) returning id into cash_id;
  end if;
  insert into public.profit_payouts(id,party,amount,deduct_home_cash,cash_transaction_id,created_by,actor_email)
  values (p_request_id,p_party,p_amount,p_deduct_home_cash,cash_id,auth.uid(),email) returning * into saved;
  return saved;
end;
$$;

create or replace function public.get_home_cash_summary()
returns jsonb language sql stable security invoker set search_path = '' as $$
  select jsonb_build_object('started_at',a.started_at,'balance',coalesce(sum(t.amount),0),
    'income',coalesce(sum(t.amount) filter(where t.kind in ('receipt','receipt_adjustment')),0),
    'expenses',-coalesce(sum(t.amount) filter(where t.kind in ('expense','expense_reversal','profit_payout')),0))
  from public.home_cash_account a left join public.home_cash_transactions t on true where a.id=1 group by a.started_at;
$$;
revoke all on function public.get_profit_balances(),public.record_profit_payout(uuid,text,numeric,boolean) from public,anon,authenticated;
grant execute on function public.get_profit_balances(),public.record_profit_payout(uuid,text,numeric,boolean) to authenticated;

notify pgrst,'reload schema';
commit;
