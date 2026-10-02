begin;

alter table public.orders
  add column if not exists home_profit_deduction numeric(14,2) not null default 0,
  add column if not exists rahaf_profit_deduction numeric(14,2) not null default 0,
  add column if not exists mira_profit_deduction numeric(14,2) not null default 0;
alter table public.orders drop constraint if exists orders_profit_deductions_check;
alter table public.orders add constraint orders_profit_deductions_check check (
  home_profit_deduction >= 0 and home_profit_deduction::text not in ('NaN','Infinity','-Infinity')
  and rahaf_profit_deduction >= 0 and rahaf_profit_deduction::text not in ('NaN','Infinity','-Infinity')
  and mira_profit_deduction >= 0 and mira_profit_deduction::text not in ('NaN','Infinity','-Infinity')
);
alter table public.order_profit_accruals
  add column if not exists home_deduction numeric(14,2) not null default 0,
  add column if not exists rahaf_deduction numeric(14,2) not null default 0,
  add column if not exists mira_deduction numeric(14,2) not null default 0;

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
  perform 1 from public.home_cash_account where id = 1 for update;
  select * into o from public.orders where id = p_order_id;
  if not found then return; end if;
  select round(coalesce(sum(coalesce(paid_price,price,0)),0),2) into purchase_total
    from public.purchases where order_id = p_order_id;
  net_profit := greatest(0, purchase_total + round(coalesce(o.postal_fee,0),2) - round(coalesce(o.spent_amount,0),2) - o.marketing_fee);
  if o.home_profit_percent is not null then
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
  -- Deductions reduce only their own party's share, never other shares or physical cash.
  home_amount := greatest(0,home_amount - o.home_profit_deduction);
  rahaf_amount := greatest(0,rahaf_amount - o.rahaf_profit_deduction);
  mira_amount := greatest(0,mira_amount - o.mira_profit_deduction);
  insert into public.order_profit_accruals
    (order_id,order_name,configured,purchase_value,spent,postal_fee,marketing_fee,distributable_profit,
      home_percent,rahaf_percent,mira_percent,home_profit,rahaf_profit,mira_profit,
      home_deduction,rahaf_deduction,mira_deduction)
  values (o.id,o.order_name,o.home_profit_percent is not null,purchase_total,coalesce(o.spent_amount,0),
    coalesce(o.postal_fee,0),o.marketing_fee,net_profit,o.home_profit_percent,o.rahaf_profit_percent,
    o.mira_profit_percent,home_amount,rahaf_amount,mira_amount,
    o.home_profit_deduction,o.rahaf_profit_deduction,o.mira_profit_deduction)
  on conflict(order_id) do update set order_name=excluded.order_name,configured=excluded.configured,
    purchase_value=excluded.purchase_value,spent=excluded.spent,postal_fee=excluded.postal_fee,
    marketing_fee=excluded.marketing_fee,distributable_profit=excluded.distributable_profit,
    home_percent=excluded.home_percent,rahaf_percent=excluded.rahaf_percent,mira_percent=excluded.mira_percent,
    home_profit=excluded.home_profit,rahaf_profit=excluded.rahaf_profit,mira_profit=excluded.mira_profit,
    home_deduction=excluded.home_deduction,rahaf_deduction=excluded.rahaf_deduction,mira_deduction=excluded.mira_deduction,
    updated_at=clock_timestamp();
end;
$$;
revoke all on function public.refresh_order_profit(uuid) from public,anon,authenticated;

drop trigger if exists orders_profit_accrual on public.orders;
create trigger orders_profit_accrual after insert or update of order_name,spent_amount,postal_fee,marketing_fee,
  home_profit_percent,rahaf_profit_percent,mira_profit_percent,
  home_profit_deduction,rahaf_profit_deduction,mira_profit_deduction on public.orders
  for each row execute function public.sync_order_profit();

do $$ declare item record; begin
  for item in select id from public.orders loop perform public.refresh_order_profit(item.id); end loop;
end $$;

notify pgrst,'reload schema';
commit;
