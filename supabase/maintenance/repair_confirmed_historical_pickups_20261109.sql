-- User-confirmed full historical settlement. This is NOT a new cash receipt.
-- Run the whole file in Supabase SQL Editor as the database administrator.
-- Source numbers refer to the 58-row export reviewed on 2026-11-09.
-- Includes #34 and #35 subsequently confirmed by the user; 31 purchases total.
begin;

create temporary table confirmed_pickup_repairs (
  number integer primary key,
  purchase_id uuid unique not null,
  customer_name text not null,
  order_name text not null,
  pickup_point text not null,
  amount numeric(14,2) not null
) on commit drop;

insert into confirmed_pickup_repairs values
  (23,'e6420fce-28b3-45ee-9d2a-ad4c46f438b1','فاتنة 2','rahafshein333@outlook.sa','من نقطة الاستلام',227),
  (25,'9649fa41-01bf-4386-b653-ab43be5feb26','إسراء كتانه','mirashein14','من البيت',22),
  (26,'64a44350-09fb-4597-8578-aed89733e64f','حلا موسى','mirashein14','من البيت',19),
  (27,'d0443957-8057-4672-be1a-881ee60da031','رهف','mirashein14','من البيت',20),
  (28,'2069e691-9446-4bf0-939b-eeb0fa520417','لجين','mirashein14','من البيت',28),
  (29,'18a1ef20-aac5-4472-b476-f5d03168ea70','لين ابو صفية','mirashein14','من البيت',84),
  (30,'af3bc995-1cd9-49cb-abcb-a4e761520718','مرح ابو صفية','mirashein14','من البيت',114),
  (31,'d6d6dfe8-4a79-46d6-a9b7-a4a9318082d5','مي','mirashein14','من البيت',500),
  (32,'551bb13c-af22-42d5-b396-705d7736aa65','ميرا','mirashein14','من البيت',0),
  (33,'5c4a4b16-c86a-4ccc-a965-3edb6de6a7e5','ندين','mirashein14','من البيت',7),
  (34,'83994d6e-e962-41ad-8cf1-ede25051fb48','نور جلاد','mirashein14','من البيت',70),
  (35,'5ae9bfa2-7d31-4e1f-87e2-06f5be786793','حلا موسى','rahafshein07@outlook.com','من البيت',45),
  (36,'46b47cc2-192c-475c-9977-f7a62fc4c646','دانا ثروت','rahafshein07@outlook.com','من البيت',132),
  (37,'a4688eff-bfb0-4d16-8b93-a9d5885c16dc','رهف','rahafshein07@outlook.com','من البيت',0),
  (38,'11b52200-2d01-4a62-89c7-9625fe170324','ريما مدني','rahafshein07@outlook.com','من البيت',9),
  (39,'931f21e1-ebb1-4221-9682-666f8664a142','صبا كرمي','rahafshein07@outlook.com','من البيت',7),
  (40,'6916a055-d4b4-411d-92f5-d1e5fbb2a996','مرح ابو صفية','rahafshein07@outlook.com','من البيت',27),
  (41,'f7a0bb28-22be-4c7d-b5bf-11a5f50f42c8','ميرا','rahafshein07@outlook.com','من البيت',34),
  (42,'928b5033-bc8e-41aa-a887-411918735bff','يانا حجازي','rahafshein07@outlook.com','من البيت',294),
  (43,'9dd653ab-d236-4e5e-b6f1-c12024de6bba','يمنى','rahafshein07@outlook.com','من البيت',55),
  (44,'8f7e21c5-d3dd-4925-8c30-c48a3b60271c','ريما مدني','rahafshein6@outlook.com','من البيت',51),
  (45,'2937142a-ef36-4219-91e1-c3b1572dce2c','سالي صياغ','mirashein13@outlook.com','من البيت',200),
  (46,'3bb487b9-8f7b-49b9-ba7e-4cd9a1e89f99','ميرا','mirashein13@outlook.com','من البيت',24),
  (48,'68741110-ec5c-44e9-8364-343796e9c8a6','مرح ابو صفية','mirashein8@outlook.com','من البيت',88),
  (49,'9bef8adf-7aa2-49e9-8fc5-41a55dd32d85','مي','mirashein8@outlook.com','من البيت',69),
  (50,'df9932be-b387-48b5-88d5-d4773759d44a','ميرا','mirashein8@outlook.com','من البيت',23),
  (54,'45e6c985-8589-4fbd-beee-4ba834937b0c','اسماء مريم','rahafayman297@hotmail.com','توصيل',277),
  (55,'1c7ff488-70c6-4d9d-9616-e052163cc0cd','عميد حمد','rahafayman297@hotmail.com','توصيل',940),
  (56,'05fb9ceb-9a45-4f89-82f8-7d5b588b501b','اسماء مريم','Rahafayman296','توصيل',332),
  (57,'a52ada2d-786e-46c1-8683-a989e5708cca','الهام مناع','Rahafayman296','توصيل',576),
  (58,'0e3a19a6-ad27-484b-9b80-7b0bd019cc95','ميرا','طلبية rahafayman286','توصيل',32);

-- Snapshot the finance sources, not just their grand totals. No tables are disabled.
create temporary view confirmed_repair_finance_snapshot as
select jsonb_build_object(
  'orders',(select coalesce(jsonb_agg(to_jsonb(t) order by id),'[]') from public.orders t),
  'cash',(select coalesce(jsonb_agg(to_jsonb(t) order by id),'[]') from public.home_cash_transactions t),
  'cash_account',(select coalesce(jsonb_agg(to_jsonb(t) order by id),'[]') from public.home_cash_account t),
  'receipts',(select coalesce(jsonb_agg(to_jsonb(t) order by purchase_id),'[]') from public.home_cash_receipts t),
  'instant_receipts',(select coalesce(jsonb_agg(to_jsonb(t) order by instant_pickup_id),'[]') from public.instant_pickup_cash_receipts t),
  'accruals',(select coalesce(jsonb_agg(to_jsonb(t) order by order_id),'[]') from public.order_profit_accruals t),
  'payouts',(select coalesce(jsonb_agg(to_jsonb(t) order by id),'[]') from public.profit_payouts t),
  'profit_balances',(select coalesce(jsonb_agg(to_jsonb(t) order by party),'[]') from public.get_profit_balances() t),
  'purchase_details',(select coalesce(jsonb_agg(to_jsonb(t) - array[
      'picked_up','picked_up_at','collected','collected_at','received_amount','remitted_amount',
      'picked_up_by','picked_up_by_email','payment_model_ready'] order by id),'[]') from public.purchases t),
  'other_payments',(select coalesce(jsonb_agg(to_jsonb(t) order by id),'[]') from public.purchase_payments t
    where not exists(select 1 from pg_temp.confirmed_pickup_repairs r where r.purchase_id=t.purchase_id)),
  'finance_collection_totals',(select coalesce(jsonb_agg(to_jsonb(t) order by id),'[]') from (
    select o.id,coalesce(sum(coalesce(p.paid_price,p.price,0)) filter(where p.collected),0) as collected
    from public.orders o left join public.purchases p on p.order_id=o.id
    where o.created_at >= timestamptz '2026-10-01 00:00:00+03' group by o.id
  ) t)
) as value;

do $$
declare repair record; saved_purchase public.purchases%rowtype; history public.purchase_payments%rowtype;
  history_count integer; finance_before jsonb; finance_after jsonb;
begin
  if current_user in ('anon','authenticated') or auth.uid() is not null then
    raise exception 'Run this historical correction in SQL Editor without an application user session';
  end if;
  perform 1 from public.home_cash_account where id=1 for update;
  if not found then raise exception 'Home cash account is missing; nothing changed'; end if;
  perform 1 from public.purchases p join pg_temp.confirmed_pickup_repairs r on r.purchase_id=p.id
    order by p.id for update of p;
  if (select count(*) from public.purchases p join pg_temp.confirmed_pickup_repairs r on r.purchase_id=p.id)<>31 then
    raise exception 'Some selected purchases are missing; nothing changed';
  end if;
  select value into finance_before from pg_temp.confirmed_repair_finance_snapshot;

  for repair in select * from pg_temp.confirmed_pickup_repairs order by number loop
    select * into strict saved_purchase from public.purchases where id=repair.purchase_id;
    if saved_purchase.customer_name is distinct from repair.customer_name or saved_purchase.pickup_point is distinct from repair.pickup_point
      or coalesce(saved_purchase.paid_price,saved_purchase.price,0)<>repair.amount
      or not exists(select 1 from public.orders where id=saved_purchase.order_id and order_name=repair.order_name) then
      raise exception 'Purchase #% changed since the reviewed export; nothing changed',repair.number;
    end if;
    select count(*) into history_count from public.purchase_payments where purchase_id=saved_purchase.id;
    if history_count>0 then
      select * into history from public.purchase_payments where purchase_id=saved_purchase.id order by id limit 1;
      -- An exact previous run may be retried; any other payment history requires review.
      if history_count=1 and history.id=saved_purchase.id and history.legacy and history.amount=repair.amount
        and history.remitted_at is not null and history.cancelled_at is null
        and history.cash_recognized_amount=0 and history.cash_transaction_id is null
        and history.cancellation_cash_id is null and history.pickup_point=saved_purchase.pickup_point
        and saved_purchase.picked_up and saved_purchase.collected and saved_purchase.received_amount=repair.amount and saved_purchase.remitted_amount=repair.amount then
        continue;
      end if;
      raise exception 'Purchase #% has payment history that needs review; nothing changed',repair.number;
    end if;
    if saved_purchase.picked_up or saved_purchase.collected or saved_purchase.received_amount<>0 or saved_purchase.remitted_amount<>0 then
      raise exception 'Purchase #% status changed since the export; nothing changed',repair.number;
    end if;
    if exists(select 1 from public.home_cash_transactions where purchase_id=saved_purchase.id) then
      raise exception 'Purchase #% now has cash history that needs review; nothing changed',repair.number;
    end if;

    -- Historical baseline: paid and already remitted, with ZERO new cash recognized.
    -- Recording time is today; the original payment date/account remain unknown.
    insert into public.purchase_payments(id,purchase_id,amount,pickup_point,legacy,remitted_at,cash_recognized_amount)
      values(saved_purchase.id,saved_purchase.id,repair.amount,saved_purchase.pickup_point,true,clock_timestamp(),0);
    update public.purchases set received_amount=repair.amount,remitted_amount=repair.amount,picked_up=true
      where id=saved_purchase.id;
    if not exists(select 1 from public.purchases where id=saved_purchase.id and picked_up and collected
      and received_amount=repair.amount and remitted_amount=repair.amount) then
      raise exception 'Purchase #% could not be verified as settled; nothing changed',repair.number;
    end if;
  end loop;

  select value into finance_after from pg_temp.confirmed_repair_finance_snapshot;
  if finance_after is distinct from finance_before then
    raise exception 'Finance protection check failed; ALL corrections rolled back';
  end if;
end;
$$;

select r.number,p.id,p.customer_name,o.order_name,p.picked_up,p.collected,p.received_amount,p.remitted_amount,
  'unchanged'::text as home_cash_and_finance
from pg_temp.confirmed_pickup_repairs r join public.purchases p on p.id=r.purchase_id
join public.orders o on o.id=p.order_id order by r.number;

drop view pg_temp.confirmed_repair_finance_snapshot;
commit;
