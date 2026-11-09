begin;

alter table public.profit_payouts add column if not exists home_recipient text;
alter table public.profit_payouts drop constraint if exists profit_payouts_home_recipient_check;
alter table public.profit_payouts add constraint profit_payouts_home_recipient_check check (
  home_recipient is null or (party='home' and home_recipient in ('reem','rawand','mama'))
);

create or replace function public.record_profit_payout(
  p_request_id uuid,p_party text,p_amount numeric,p_deduct_home_cash boolean,p_home_recipient text
)
returns public.profit_payouts language plpgsql security definer set search_path='' as $$
declare saved public.profit_payouts%rowtype; available numeric; cash_available numeric;
  cash_id uuid; actor text; recipient text;
begin
  if not coalesce(public.is_admin(),false) then raise exception 'لا توجد صلاحية' using errcode='42501'; end if;
  if p_request_id is null or p_party is null or p_party not in ('home','rahaf','mira') or p_amount is null
    or p_amount<=0 or p_amount<>round(p_amount,2) or p_amount::text in ('NaN','Infinity','-Infinity')
    or p_deduct_home_cash is null then raise exception 'بيانات التسليم غير صحيحة'; end if;
  perform 1 from public.home_cash_account where id=1 for update;
  if not found then raise exception 'تتبع مصاري البيت غير متاح'; end if;
  select * into saved from public.profit_payouts where id=p_request_id;
  if found then
    if saved.party=p_party and saved.amount=p_amount and saved.deduct_home_cash=p_deduct_home_cash
      and saved.home_recipient is not distinct from p_home_recipient then return saved; end if;
    raise exception 'هذا التسليم محفوظ مسبقاً بقيم مختلفة';
  end if;
  if (p_party='home' and (p_home_recipient is null or p_home_recipient not in ('reem','rawand','mama')))
    or (p_party<>'home' and p_home_recipient is not null) then raise exception 'اختاري مستلم أرباح البيت'; end if;
  select balance into available from public.get_profit_balances() where party=p_party;
  if p_amount>available then raise exception 'المبلغ أكبر من رصيد أرباح الطرف'; end if;
  select email into actor from auth.users where id=auth.uid();
  recipient:=case when p_party='home' then case p_home_recipient when 'reem' then 'ريم' when 'rawand' then 'روند' else 'ماما' end
    when p_party='rahaf' then 'رهف' else 'ميرا' end;
  if p_deduct_home_cash then
    select coalesce(sum(amount),0) into cash_available from public.home_cash_transactions;
    if p_amount>cash_available then raise exception 'المبلغ أكبر من الرصيد المتاح في البيت'; end if;
    insert into public.home_cash_transactions(kind,amount,source,customer_name,note,created_by,actor_email)
      values('profit_payout',-p_amount,'home',recipient,'تسليم أرباح',auth.uid(),actor) returning id into cash_id;
  end if;
  insert into public.profit_payouts(id,party,home_recipient,amount,deduct_home_cash,cash_transaction_id,created_by,actor_email)
    values(p_request_id,p_party,p_home_recipient,p_amount,p_deduct_home_cash,cash_id,auth.uid(),actor) returning * into saved;
  return saved;
end;
$$;
revoke all on function public.record_profit_payout(uuid,text,numeric,boolean,text) from public,anon,authenticated;
grant execute on function public.record_profit_payout(uuid,text,numeric,boolean,text) to authenticated;

-- Keep the old signature for non-Home callers and idempotent legacy retries.
create or replace function public.record_profit_payout(p_request_id uuid,p_party text,p_amount numeric,p_deduct_home_cash boolean)
returns public.profit_payouts language sql security invoker set search_path='' as $$
  select public.record_profit_payout(p_request_id,p_party,p_amount,p_deduct_home_cash,null);
$$;

create or replace function public.get_home_cash_activity_view(p_view text,p_offset integer default 0,p_limit integer default 26)
returns table(id uuid,kind text,amount numeric,category text,quantity integer,source text,
  order_name text,customer_name text,note text,actor_email text,created_at timestamptz,
  balance_after numeric,is_voided boolean,pickup_actor_email text)
language sql stable security invoker set search_path='' as $$
  with ledger as (
    select t.*,sum(t.amount) over(order by t.created_at,t.id rows unbounded preceding) as running_balance
    from public.home_cash_transactions t
  )
  select t.id,t.kind,t.amount,t.category,t.quantity,t.source,t.order_name,t.customer_name,t.note,t.actor_email,t.created_at,
    t.running_balance,exists(select 1 from public.home_cash_transactions r where r.reverses_id=t.id),
    t.pickup_actor_email
  from ledger t
  where (p_view='receipts' and t.kind in ('receipt','receipt_adjustment'))
    or (p_view='manual' and t.kind not in ('receipt','receipt_adjustment'))
  order by t.created_at desc,t.id desc
  limit least(greatest(p_limit,1),100) offset greatest(p_offset,0);
$$;
revoke all on function public.get_home_cash_activity_view(text,integer,integer) from public,anon,authenticated;
grant execute on function public.get_home_cash_activity_view(text,integer,integer) to authenticated;

notify pgrst,'reload schema';
commit;
