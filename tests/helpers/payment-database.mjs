import { PGlite } from "@electric-sql/pglite";
import { readFile } from "node:fs/promises";

export const id = (n) => `00000000-0000-4000-8000-${String(n).padStart(12,"0")}`;
export const HOME="من البيت", MARY="من نقطة الاستلام", NABLUS="نقطة استلام نابلس", DELIVERY="توصيل";
export const sqlFile = (name) => readFile(new URL(`../../supabase/migrations/${name}`,import.meta.url),"utf8");
export async function paymentDatabase({ trackedBeforeUpgrade = false } = {}) {
  const db=new PGlite();
  await db.exec(`
    create role anon; create role authenticated; create schema auth;
    create table auth.users(id uuid primary key,email text);
    create function auth.uid() returns uuid language sql stable as $$ select nullif(current_setting('request.jwt.claim.sub',true),'')::uuid $$;
    create table public.user_roles(user_id uuid primary key,role text);
    create function public.is_admin() returns boolean language sql stable as $$ select exists(select 1 from public.user_roles where user_id=auth.uid() and role='admin') $$;
    create table public.orders(id uuid primary key,order_name text not null,spent_amount numeric default 0,postal_fee numeric default 0,
      total_profit numeric,mira_profit numeric,rahaf_profit numeric,created_at timestamptz default now(),order_date date default current_date,
      arrived boolean default true,placed_at_pickup boolean default true);
    create table public.purchases(id uuid primary key,order_id uuid references public.orders on delete cascade,customer_name text,
      price numeric,paid_price numeric,pickup_point text,picked_up boolean default false,picked_up_at timestamptz,
      collected boolean default false,collected_at timestamptz,ready_for_pickup boolean default true,ready_for_pickup_at timestamptz,created_at timestamptz default now());
    insert into auth.users values('${id(1)}','rahaf@she-store.com'),('${id(2)}','reem@she-store.com'),
      ('${id(3)}','rawand@she-store.com'),('${id(4)}','maryamti@she-store.com'),('${id(5)}','nablus@she-store.com'),
      ('${id(6)}','laaura@she-store.com'),('${id(7)}','not-reem@she-store.com');
    insert into public.user_roles values('${id(1)}','admin'),('${id(2)}','viewer'),('${id(3)}','viewer'),('${id(4)}','pickup'),
      ('${id(5)}','pickup'),('${id(6)}','pickup'),('${id(7)}','viewer');
    grant usage on schema public,auth to authenticated,anon;
    grant select on public.user_roles to authenticated;
    grant select,insert,update,delete on public.purchases to authenticated;
    alter table public.purchases enable row level security;
    create policy purchase_read on public.purchases for select to authenticated using(true);
    create policy purchase_admin on public.purchases for all to authenticated using(public.is_admin()) with check(public.is_admin());
    create policy pickup_update on public.purchases for update to authenticated using(exists(select 1 from public.user_roles where user_id=auth.uid() and role='pickup'));
    select set_config('request.jwt.claim.sub','${id(1)}',false);
    insert into public.orders(id,order_name,spent_amount,postal_fee) values('${id(10)}','Payment Order',20,5);
    insert into public.purchases(id,order_id,customer_name,price,paid_price,pickup_point) values
      ('${id(100)}','${id(10)}','Tala',50,42,'${HOME}'),('${id(101)}','${id(10)}','Nablus customer',100,100,'${NABLUS}'),
      ('${id(102)}','${id(10)}','Delivery customer',80,80,'${DELIVERY}'),('${id(103)}','${id(10)}','Maryamti customer',60,60,'${MARY}'),
      ('${id(104)}','${id(10)}','Historical Home',30,30,'${HOME}'),('${id(105)}','${id(10)}','Historical Nablus',20,20,'${NABLUS}');
    update public.purchases set picked_up=true,picked_up_at=now() where id in ('${id(104)}','${id(105)}');
  `);
  for(const file of ['20261002010000_add_home_cash_ledger.sql','20261002020000_add_instant_pickups.sql',
    '20261002030000_add_instant_pickup_delete.sql','20261002040000_add_profit_distribution.sql',
    '20261003000000_add_party_profit_deductions.sql','20261021000000_add_purchase_receipt_rpc.sql']) await db.exec(await sqlFile(file));
  await db.exec(`update public.orders set home_profit_percent=30,rahaf_profit_percent=30,mira_profit_percent=40;
    select public.create_instant_pickup('${id(200)}','Instant Home',70,'${HOME}');
    select public.create_instant_pickup('${id(201)}','Instant Nablus',70,'${NABLUS}');
    select public.create_instant_pickup('${id(202)}','Instant Mary',70,'${MARY}');`);
  if (trackedBeforeUpgrade) await db.exec(`update public.purchases set picked_up=true,picked_up_at=now() where id='${id(100)}';
    select public.update_instant_pickup('${id(200)}','receive');`);
  const migration=await sqlFile('20261021010000_add_purchase_payments.sql');
  await db.exec(migration);
  const staffMigration=await sqlFile('20261022000000_restore_pickup_staff_receipts.sql');
  await db.exec(staffMigration);
  const identity=(n)=>db.exec(`reset role; select set_config('request.jwt.claim.sub','${id(n)}',false); set role authenticated;`);
  return {db,migration,staffMigration,identity};
}
