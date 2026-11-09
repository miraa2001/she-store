import assert from "node:assert/strict";
import test from "node:test";
import { canCollectPayment } from "../src/lib/payments.js";
import { paymentDatabase, id, sqlFile, HOME } from "./helpers/payment-database.mjs";

test("Home payments enter cash only on explicit collection without rewriting historical finances", async t => {
  const { db, identity } = await paymentDatabase();
  t.after(() => db.close());
  for (const file of ['20261106000000_location_readiness_and_pickup_actor.sql','20261109000000_add_full_pickup_confirmation.sql']) await db.exec(await sqlFile(file));
  const confirm = async (request, purchase, instant = false) => (await db.query('select public.confirm_full_purchase_receipt($1,$2,$3) as value', [id(request),id(purchase),instant])).rows[0].value;
  const record = async (request, purchase, amount, receive = false) => (await db.query('select public.record_purchase_payment($1,$2,$3,$4,false) as value',[id(request),id(purchase),amount,receive])).rows[0].value;
  const collect = async (purchases, instant = false) => Number((await db.query('select public.collect_purchase_payments($1,$2) as value',[purchases.map(id),instant])).rows[0].value);
  const cash = async () => Number((await db.query('select public.get_home_cash_summary() as value')).rows[0].value.balance);
  const purchase = async (n, instant = false) => (await db.query(`select * from public.${instant ? 'instant_pickups' : 'purchases'} where id=$1`,[id(n)])).rows[0];
  const snapshot = async () => {
    await db.exec('reset role;');
    const result = {};
    for (const table of ['purchases','instant_pickups','purchase_payments','orders','order_profit_accruals','profit_payouts','home_cash_transactions','home_cash_receipts','home_cash_account']) result[table]=(await db.query(`select * from public.${table} order by 1`)).rows;
    await identity(1);
    return result;
  };
  await identity(1);
  await confirm(800,100);
  assert.equal(await cash(),42);
  assert.equal((await purchase(100)).collected,true);
  await db.exec(`reset role; insert into public.purchases(id,order_id,customer_name,price,paid_price,pickup_point) values
    ('${id(120)}','${id(10)}','New full',60,50,'${HOME}'),
    ('${id(121)}','${id(10)}','New partial',100,100,'${HOME}'),
    ('${id(122)}','${id(10)}','Free Home',0,0,'${HOME}'),
    ('${id(123)}','${id(10)}','Cancel before collection',10,10,'${HOME}');`);
  await identity(1);
  const migration = await sqlFile('20261109020000_defer_home_cash_until_collection.sql');
  await t.test("migration changes no existing purchase, payment, cash, payout or profit record",async()=>{
    const before=await snapshot();await db.exec('reset role;');await db.exec(migration);
    assert.deepEqual(await snapshot(),before);
  });
  await t.test("Home one-tick receipt and retries stay paid/uncollected and do not add cash",async()=>{
    await identity(2);const result=await confirm(801,120);
    assert.equal(result.purchase.picked_up,true);assert.equal(result.purchase.collected,false);
    assert.equal(result.purchase.collected_at,null);assert.equal(Number(result.purchase.received_amount),50);
    assert.equal(Number(result.purchase.remitted_amount),0);assert.equal(result.purchase.picked_up_by_email,'reem@she-store.com');
    assert.equal((await confirm(801,120)).changed_receipt,false);assert.equal((await confirm(802,120)).changed_receipt,false);
    await assert.rejects(collect([120]),/رهف/);await identity(1);assert.equal(await cash(),42);
  });
  await t.test("explicit collection posts once, completes receipt and preserves original pickup actor",async()=>{
    assert.equal(await collect([120]),50);assert.equal(await cash(),92);
    const p=await purchase(120);assert.equal(p.collected,true);assert(p.collected_at);assert.equal(p.picked_up_by_email,'reem@she-store.com');
    assert.equal(await collect([120]),0);assert.equal(await cash(),92);
    const entry=(await db.query('select * from public.purchase_payments where purchase_id=$1',[id(120)])).rows[0];
    assert.equal(entry.remitted_by,id(1));assert.equal(Number(entry.cash_recognized_amount),50);
  });
  await t.test("each Home installment waits for collection; cancellations reverse only already collected cash",async()=>{
    await record(803,121,20,true);assert.equal(await cash(),92);assert.equal((await purchase(121)).collected,false);
    assert.equal(await collect([121]),20);assert.equal(await cash(),112);assert.equal((await purchase(121)).collected,false);
    await record(804,121,30);assert.equal(await cash(),112);
    await db.query('select public.cancel_purchase_payment($1,$2)',[id(804),'Not collected']);assert.equal(await cash(),112);
    await record(805,121,80);assert.equal(await cash(),112);assert.equal(await collect([121]),80);assert.equal(await cash(),192);
    assert.equal((await purchase(121)).collected,true);
    await db.query('select public.cancel_purchase_payment($1,$2)',[id(803),'Collected payment cancellation']);assert.equal(await cash(),172);
    assert.equal((await purchase(121)).collected,false);
  });
  await t.test("instant Home pickup uses the same deferred cash rules and never changes order profits",async()=>{
    const before=(await snapshot()).order_profit_accruals;
    await confirm(806,200,true);assert.equal((await purchase(200,true)).collected,false);assert.equal(await cash(),172);
    assert.equal(await collect([200],true),70);assert.equal(await cash(),242);assert.equal((await purchase(200,true)).collected,true);
    assert.deepEqual((await snapshot()).order_profit_accruals,before);
  });
  await t.test("Nablus staff and collection keep their existing deferred behavior",async()=>{
    await identity(5);await confirm(807,101);assert.equal((await purchase(101)).collected,false);
    await identity(1);assert.equal(await cash(),242);assert.equal(await collect([101]),100);assert.equal(await cash(),342);
  });
  await t.test("free Home purchases need explicit collection but never invent money",async()=>{
    await confirm(808,122);assert.equal((await purchase(122)).collected,false);assert.equal(await cash(),342);
    await assert.rejects(db.exec(`update public.purchases set collected_at=now(),collected=true where id='${id(122)}'`),/الإجراءات/);
    assert.equal(await collect([122]),0);assert.equal((await purchase(122)).collected,true);assert.equal(await cash(),342);
    assert.equal(await collect([122]),0);
    await db.query('select public.create_instant_pickup($1,$2,0,$3)',[id(203),'Free instant',HOME]);
    await confirm(809,203,true);assert.equal((await purchase(203,true)).collected,false);
    assert.equal(await collect([203],true),0);assert.equal((await purchase(203,true)).collected,true);assert.equal(await cash(),342);
  });
  await t.test("uncollected cancellation and failed batches add no cash; existing collected Home receipts stay collected",async()=>{
    await record(810,123,10,true);assert.equal(await cash(),342);
    await assert.rejects(collect([123,999]));assert.equal(await cash(),342);assert.equal((await purchase(123)).collected,false);
    await db.query('select public.cancel_purchase_payment($1,$2)',[id(810),'Customer cancelled']);assert.equal(await cash(),342);
    assert.equal((await purchase(123)).picked_up,false);
    assert.equal((await purchase(100)).collected,true);assert.equal(await collect([100]),0);assert.equal(await cash(),342);
  });
  await t.test("rerunning preserves history and totals, and anonymous callers cannot execute mutations",async()=>{
    const before=await snapshot();await db.exec('reset role;');await db.exec(migration);assert.deepEqual(await snapshot(),before);
    await db.exec('reset role;set role anon;');await assert.rejects(collect([120]),/permission denied/);
  });
});

test("collection controls include received free items but exclude unreceived and completed items",()=>{
  const p={picked_up:true,collected:false,paid_price:50,received_amount:50,remitted_amount:0};
  assert.equal(canCollectPayment(p),true);
  assert.equal(canCollectPayment({...p,picked_up:false}),false);
  assert.equal(canCollectPayment({...p,collected:true}),false);
  assert.equal(canCollectPayment({...p,remitted_amount:50}),false);
  assert.equal(canCollectPayment({...p,paid_price:0,received_amount:0}),true);
});

test("a pre-migration Home installment is never counted again when later installments are collected",async t=>{
  const {db,identity}=await paymentDatabase();t.after(()=>db.close());await identity(1);
  await db.query('select public.record_purchase_payment($1,$2,10,true,false)',[id(850),id(100)]);
  const balance=async()=>Number((await db.query('select public.get_home_cash_summary() as value')).rows[0].value.balance);
  assert.equal(await balance(),10);
  await db.exec('reset role;');await db.exec(await sqlFile('20261109020000_defer_home_cash_until_collection.sql'));await identity(1);
  await db.query('select public.record_purchase_payment($1,$2,32,false,false)',[id(851),id(100)]);
  assert.equal(await balance(),10);
  const collected=await db.query('select public.collect_purchase_payments($1,false) as value',[[id(100)]]);
  assert.equal(Number(collected.rows[0].value),32);assert.equal(await balance(),42);
  await db.query('select public.collect_purchase_payments($1,false)',[[id(100)]]);assert.equal(await balance(),42);
});
