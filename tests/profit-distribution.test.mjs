import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { PGlite } from "@electric-sql/pglite";

const id = (n) => `00000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const sqlFile = (name) => readFile(new URL(`../supabase/migrations/${name}`, import.meta.url), "utf8");

test("Profit distribution and partial payouts", async (t) => {
  const db = new PGlite();
  t.after(() => db.close());
  await db.exec(`
    create role anon; create role authenticated; create schema auth;
    create table auth.users(id uuid primary key,email text);
    create function auth.uid() returns uuid language sql stable as $$ select nullif(current_setting('request.jwt.claim.sub',true),'')::uuid $$;
    create table public.user_roles(user_id uuid primary key,role text);
    create function public.is_admin() returns boolean language sql stable as $$ select exists(select 1 from public.user_roles where user_id=auth.uid() and role='admin') $$;
    create table public.orders(id uuid primary key,order_name text not null,spent_amount numeric default 0,postal_fee numeric default 0,
      total_profit numeric,mira_profit numeric,rahaf_profit numeric);
    create table public.purchases(id uuid primary key,order_id uuid references public.orders on delete cascade,customer_name text,
      price numeric,paid_price numeric,pickup_point text,picked_up boolean default false,collected boolean default false);
    insert into auth.users values('${id(1)}','rahaf@she-store.com'),('${id(2)}','reem@she-store.com');
    insert into public.user_roles values('${id(1)}','admin'),('${id(2)}','viewer');
    grant usage on schema public,auth to authenticated;
    grant select on public.user_roles to authenticated;
    select set_config('request.jwt.claim.sub','${id(1)}',false);
    insert into public.orders values('${id(10)}','Order A',200,20,100,50,50);
    insert into public.purchases(id,order_id,customer_name,price,paid_price,pickup_point)
      values('${id(11)}','${id(10)}','Customer',1200,1000,'من البيت');
  `);
  await db.exec(await sqlFile("20261002010000_add_home_cash_ledger.sql"));
  const migration = await sqlFile("20261002040000_add_profit_distribution.sql");
  await db.exec(migration);
  const balance = async (party) => (await db.query("select * from public.get_profit_balances() where party=$1", [party])).rows[0];
  const cash = async () => (await db.query("select public.get_home_cash_summary() as value")).rows[0].value;
  const payout = (n, party, amount, deduct = false) => db.query("select * from public.record_profit_payout($1,$2,$3,$4)", [id(n), party, amount, deduct]);
  const accrual = async (n = 10) => (await db.query("select * from public.order_profit_accruals where order_id=$1", [id(n)])).rows[0];

  await t.test("zeros legacy values once and leaves all percentages unconfigured", async () => {
    const order = (await db.query("select * from public.orders")).rows[0];
    assert.equal(Number(order.total_profit),0);
    assert.equal(Number(order.mira_profit),0);
    assert.equal(Number(order.rahaf_profit),0);
    assert.equal(Number(order.marketing_fee),0);
    assert.equal(order.home_profit_percent,null);
    assert.equal((await accrual()).configured,false);
    assert.equal(Number((await balance("mira")).balance),0);
  });
  await t.test("requires three percentages totaling 100 and valid marketing", async () => {
    await assert.rejects(db.exec("update public.orders set mira_profit_percent=50"));
    await assert.rejects(db.exec("update public.orders set home_profit_percent=30,rahaf_profit_percent=30,mira_profit_percent=30"));
    for(const amount of [-1,"'NaN'","'Infinity'"]) await assert.rejects(db.exec(`update public.orders set marketing_fee=${amount}`));
    await db.exec("update public.orders set home_profit_percent=25,rahaf_profit_percent=25,mira_profit_percent=50,marketing_fee=20");
  });
  await t.test("earns immediately from paid price plus postal fee minus spending and marketing", async () => {
    const entry=await accrual();
    assert.equal(Number(entry.purchase_value),1000);
    assert.equal(Number(entry.distributable_profit),800);
    assert.equal(Number(entry.home_profit),200);
    assert.equal(Number(entry.rahaf_profit),200);
    assert.equal(Number(entry.mira_profit),400);
    assert.equal((await cash()).balance,0);
  });
  await t.test("400 earnings minus a 250 payment leaves 150; no cash deduction when declined", async () => {
    await db.exec("set role authenticated;");
    await payout(20,"mira",250,false);
    const mira=await balance("mira");
    assert.equal(Number(mira.earned),400);
    assert.equal(Number(mira.paid),250);
    assert.equal(Number(mira.balance),150);
    assert.equal((await cash()).balance,0);
    await payout(20,"mira",250,false);
    assert.equal(Number((await balance("mira")).balance),150);
    await assert.rejects(payout(20,"mira",200,false));
  });
  await t.test("payouts are atomic, validate inputs, and cannot exceed profit or cash balances", async () => {
    for(const amount of [-1,0,0.001,"NaN","Infinity",151]) await assert.rejects(payout(21,"mira",amount));
    await assert.rejects(payout(21,"unknown",1));
    await assert.rejects(payout(21,"mira",1,null));
    await assert.rejects(payout(21,"mira",100,true));
    assert.equal(Number((await balance("mira")).balance),150);
    assert.equal((await db.query("select count(*) as n from public.profit_payouts")).rows[0].n,1);
  });
  await t.test("optional cash deduction is logged once, including for the separate Home party", async () => {
    await db.exec("reset role; update public.purchases set picked_up=true; set role authenticated;");
    assert.equal((await cash()).balance,1000);
    await payout(22,"mira",100,true);
    await payout(22,"mira",100,true);
    assert.equal((await cash()).balance,900);
    assert.equal((await cash()).expenses,100);
    assert.equal(Number((await balance("mira")).balance),50);
    await payout(23,"home",50,false);
    assert.equal((await cash()).balance,900);
    await payout(24,"home",50,true);
    assert.equal((await cash()).balance,850);
    assert.equal(Number((await balance("home")).balance),100);
    const logs=(await db.query("select * from public.get_home_cash_activity(0,100)")).rows;
    assert.equal(logs.filter(row=>row.kind==="profit_payout").length,2);
    assert(logs.some(row=>row.kind==="profit_payout" && row.customer_name==="ميرا" && row.actor_email==="rahaf@she-store.com"));
  });
  await t.test("all settings and purchase edits recalculate earnings without resetting payouts", async () => {
    await db.exec("reset role; update public.orders set marketing_fee=120;");
    assert.equal(Number((await accrual()).distributable_profit),700);
    assert.equal(Number((await balance("mira")).balance),0);
    await db.exec("update public.orders set marketing_fee=20;");
    await db.query("insert into public.purchases(id,order_id,price,paid_price) values($1,$2,200,0)",[id(12),id(10)]);
    assert.equal(Number((await accrual()).purchase_value),1000);
    await db.query("update public.purchases set paid_price=null where id=$1",[id(12)]);
    assert.equal(Number((await accrual()).purchase_value),1200);
    await db.query("delete from public.purchases where id=$1",[id(12)]);
    assert.equal(Number((await accrual()).purchase_value),1000);
    assert.equal(Number((await balance("mira")).paid),350);
  });
  await t.test("negative profit earns zero; prior payments remain visible as an overpayment", async () => {
    await db.exec("update public.orders set marketing_fee=2000;");
    assert.equal(Number((await accrual()).distributable_profit),0);
    assert.equal(Number((await accrual()).mira_profit),0);
    assert.equal(Number((await balance("mira")).balance),-350);
    await assert.rejects(payout(25,"mira",1));
    await db.exec("update public.orders set marketing_fee=20;");
  });
  await t.test("zero marketing, price corrections, and moving purchases recalculate both orders", async () => {
    await db.exec("begin; update public.orders set marketing_fee=0;");
    assert.equal(Number((await accrual()).distributable_profit),820);
    await db.query("update public.purchases set paid_price=900 where id=$1",[id(11)]);
    assert.equal(Number((await accrual()).distributable_profit),720);
    await db.query("insert into public.orders(id,order_name,home_profit_percent,rahaf_profit_percent,mira_profit_percent) values($1,'Destination',0,100,0)",[id(31)]);
    await db.query("update public.purchases set order_id=$2 where id=$1",[id(11),id(31)]);
    assert.equal(Number((await accrual()).purchase_value),0);
    assert.equal(Number((await accrual()).distributable_profit),0);
    assert.equal(Number((await accrual(31)).rahaf_profit),900);
    await db.exec("rollback;");
  });
  await t.test("rounding conserves every cent, including a party with a zero percentage", async () => {
    await db.query("insert into public.orders(id,order_name,postal_fee,home_profit_percent,rahaf_profit_percent,mira_profit_percent) values($1,'Rounding',0.01,0,50,50)",[id(30)]);
    let entry=await accrual(30);
    assert.equal(Number(entry.home_profit),0);
    assert.equal(Number(entry.rahaf_profit),0.01);
    assert.equal(Number(entry.mira_profit),0);
    await db.query("update public.orders set postal_fee=0.05,home_profit_percent=33.34,rahaf_profit_percent=33.33,mira_profit_percent=33.33 where id=$1",[id(30)]);
    entry=await accrual(30);
    assert.equal(Number(entry.home_profit)+Number(entry.rahaf_profit)+Number(entry.mira_profit),0.05);
    await db.exec("begin;");
    await db.query("insert into public.orders(id,order_name,home_profit_percent,rahaf_profit_percent,mira_profit_percent) values($1,'Legacy decimals',0,100,0)",[id(32)]);
    await db.query("insert into public.purchases(id,order_id,price) values($1,$3,0.005),($2,$3,0.005)",[id(33),id(34),id(32)]);
    const legacy=await accrual(32);
    assert.equal(Number(legacy.purchase_value),0.01);
    assert.equal(Number(legacy.distributable_profit),0.01);
    assert.equal(Number(legacy.rahaf_profit),0.01);
    await db.exec("rollback;");
  });
  await t.test("anonymous/viewers cannot read or mutate; admins cannot rewrite payout records", async () => {
    await db.exec(`select set_config('request.jwt.claim.sub','${id(2)}',false); set role authenticated;`);
    assert.equal((await db.query("select * from public.profit_payouts")).rows.length,0);
    assert.equal((await db.query("select * from public.order_profit_accruals")).rows.length,0);
    assert((await db.query("select * from public.get_profit_balances()")).rows.every(row=>Number(row.balance)===0));
    await assert.rejects(payout(40,"mira",1));
    await db.exec("reset role; set role anon;");
    await assert.rejects(payout(40,"mira",1));
    await assert.rejects(db.query("select * from public.profit_payouts"));
    await db.exec(`reset role; select set_config('request.jwt.claim.sub','${id(1)}',false); set role authenticated;`);
    await assert.rejects(db.exec("delete from public.profit_payouts"));
    await assert.rejects(db.exec("update public.order_profit_accruals set mira_profit=99999"));
    await assert.rejects(db.query("select public.refresh_order_profit($1)",[id(10)]));
    await db.exec("reset role;");
  });
  await t.test("migration reruns preserve configured earnings, payouts and cash", async () => {
    await db.exec("update public.orders set total_profit=77;");
    const before=await balance("mira");
    await db.exec(migration);
    assert.deepEqual(await balance("mira"),before);
    assert.equal(Number((await db.query("select total_profit from public.orders where id=$1",[id(10)])).rows[0].total_profit),77);
    assert.equal((await cash()).balance,850);
  });
  await t.test("instant receipts increase cash without entering order earnings", async () => {
    await db.exec(await sqlFile("20261002020000_add_instant_pickups.sql"));
    const before=await balance("mira");
    await db.exec("begin;");
    await db.query("select public.create_instant_pickup($1,'Standalone',7,'من البيت')",[id(50)]);
    await db.query("select public.update_instant_pickup($1,'receive')",[id(50)]);
    assert.equal((await cash()).balance,857);
    assert.deepEqual(await balance("mira"),before);
    await db.exec("rollback;");
  });
  await t.test("order deletion preserves the earned snapshot and immutable payout history", async () => {
    const before=await balance("mira");
    await db.query("delete from public.orders where id=$1",[id(10)]);
    assert.deepEqual(await balance("mira"),before);
    assert.equal(Number((await accrual()).mira_profit),400);
    assert.equal((await db.query("select count(*) as n from public.profit_payouts")).rows[0].n,4);
  });
  await t.test("party deductions default to zero and reduce only that party's share", async () => {
    const deductionsMigration=await sqlFile("20261003000000_add_party_profit_deductions.sql");
    const before=await balance("mira");
    await db.exec(deductionsMigration);
    assert.deepEqual(await balance("mira"),before);
    await db.query("insert into public.orders(id,order_name,postal_fee,home_profit_percent,rahaf_profit_percent,mira_profit_percent) values($1,'Deductions',900,25,25,50)",[id(60)]);
    assert.equal(Number((await accrual(60)).mira_profit),450);
    assert.equal(Number((await accrual(60)).mira_deduction),0);
    await db.query("update public.orders set mira_profit_deduction=50,home_profit_deduction=25,rahaf_profit_deduction=10 where id=$1",[id(60)]);
    let entry=await accrual(60);
    assert.equal(Number(entry.mira_profit),400);
    assert.equal(Number(entry.home_profit),200);
    assert.equal(Number(entry.rahaf_profit),215);
    assert.equal((await cash()).balance,850);
    for(const amount of [-1,"'NaN'","'Infinity'"]) await assert.rejects(db.exec(`update public.orders set mira_profit_deduction=${amount} where id='${id(60)}'`));
    await payout(61,"mira",100,false);
    const paid=Number((await balance("mira")).paid);
    await db.query("update public.orders set mira_profit_deduction=500 where id=$1",[id(60)]);
    assert.equal(Number((await accrual(60)).mira_profit),0);
    assert.equal(Number((await balance("mira")).paid),paid);
    await db.query("update public.orders set mira_profit_deduction=50 where id=$1",[id(60)]);
    await db.query("insert into public.purchases(id,order_id,price) values($1,$2,100)",[id(62),id(60)]);
    assert.equal(Number((await accrual(60)).mira_profit),450);
    const balancesBefore=await balance("mira");
    await db.exec(deductionsMigration);
    assert.deepEqual(await balance("mira"),balancesBefore);
    assert.equal(Number((await accrual(60)).mira_deduction),50);
    assert.equal((await cash()).balance,850);
  });
});
