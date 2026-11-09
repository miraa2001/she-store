import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { paymentDatabase, id, sqlFile, HOME, NABLUS } from "./helpers/payment-database.mjs";
import { confirmFullPickup } from "../src/lib/payments.js";

test("one-tick full pickup uses existing payments/cash rules and never settles an existing partial receipt", async t => {
  const { db, identity } = await paymentDatabase();
  t.after(() => db.close());
  const migration=await sqlFile('20261109000000_add_full_pickup_confirmation.sql');
  const rows=async table=>(await db.query(`select * from public.${table} order by 1`)).rows;
  const before=await rows('purchase_payments'), initialCash=await rows('home_cash_transactions');
  await db.exec(migration);
  assert.deepEqual(await rows('purchase_payments'),before);
  assert.deepEqual(await rows('home_cash_transactions'),initialCash);
  const confirm=async(request,n,instant=false)=>(await db.query('select public.confirm_full_purchase_receipt($1,$2,$3) as value',[id(request),id(n),instant])).rows[0].value;
  const cash=async()=>{await identity(1);return Number((await db.query('select public.get_home_cash_summary() as value')).rows[0].value.balance);};

  await t.test("Reem's one-tick Home receipt pays the full agreed price, not original price",async()=>{
    await identity(2);
    const result=await confirm(800,100);
    assert.equal(result.changed_receipt,true);assert.equal(result.purchase.picked_up,true);
    assert.equal(Number(result.purchase.received_amount),42);assert.equal(result.purchase.collected,true);
    assert.equal(await cash(),42);
  });
  await t.test("retries, duplicate clicks and different request IDs never duplicate payment",async()=>{
    await identity(2);
    assert.equal((await confirm(800,100)).changed_receipt,false);
    assert.equal((await confirm(801,100)).changed_receipt,false);
    assert.equal((await rows('purchase_payments')).filter(row=>row.purchase_id===id(100)).length,1);
    await assert.rejects(confirm(800,102),/مستخدم/);
    assert.equal(await cash(),42);
  });
  await t.test("staff keep their own locations and full-payment-only permissions",async()=>{
    await identity(4);await assert.rejects(confirm(802,101),/الخاصة/);
    const mary=await confirm(802,103);assert.equal(Number(mary.purchase.received_amount),60);assert.equal(mary.purchase.collected,false);
    await identity(5);const nab=await confirm(803,101);assert.equal(Number(nab.purchase.received_amount),100);assert.equal(nab.purchase.collected,false);
    await assert.rejects(confirm(804,102),/الخاصة/);
    assert.equal(await cash(),42);
  });
  await t.test("Delivery full receipt waits for collection before crediting Home cash",async()=>{
    await identity(2);const delivery=await confirm(805,102);
    assert.equal(Number(delivery.purchase.received_amount),80);assert.equal(Number(delivery.purchase.remitted_amount),0);
    assert.equal(await cash(),42);
  });
  await t.test("a stale unchecked client cannot fully pay an already received partial purchase",async()=>{
    await db.exec(`reset role;insert into public.purchases(id,order_id,customer_name,price,pickup_point)
      values('${id(120)}','${id(10)}','Partial',100,'${HOME}');`);
    await identity(1);
    await db.query('select public.record_purchase_payment($1,$2,5,true,false)',[id(806),id(120)]);
    const result=await confirm(807,120);
    assert.equal(result.changed_receipt,false);assert.equal(Number(result.purchase.received_amount),5);
    assert.equal(await cash(),47);
  });
  await t.test("zero-price ordinary and instant purchases work without invented payments",async()=>{
    await db.exec(`reset role;insert into public.purchases(id,order_id,customer_name,price,pickup_point) values
      ('${id(121)}','${id(10)}','Free Home',0,'${HOME}'),('${id(122)}','${id(10)}','Free Nablus',0,'${NABLUS}');
      select public.create_instant_pickup('${id(220)}','Free Instant',0,'${HOME}');`);
    const count=(await rows('purchase_payments')).length;
    await identity(2);
    assert.equal((await confirm(808,121)).purchase.picked_up,true);
    assert.equal((await confirm(809,220,true)).purchase.picked_up,true);
    await identity(5);assert.equal((await confirm(810,122)).purchase.picked_up,true);
    await identity(1);
    assert.equal((await rows('purchase_payments')).length,count);
    assert.equal(await cash(),47);
  });
  await t.test("unauthorized, anonymous, null and not-ready requests remain blocked",async()=>{
    await db.exec(`reset role;insert into public.purchases(id,order_id,customer_name,price,pickup_point,ready_for_pickup)
      values('${id(123)}','${id(10)}','Not ready',50,'${HOME}',false);`);
    await identity(1);await assert.rejects(confirm(811,123),/جاهز/);
    for(const actor of [3,6,7]){await identity(actor);await assert.rejects(confirm(812,123));}
    await identity(2);await assert.rejects(confirm(812,103));
    await assert.rejects(db.query('select public.confirm_full_purchase_receipt(null,$1,false)',[id(123)]));
    await db.exec('reset role;set role anon;');await assert.rejects(confirm(812,123));
    await identity(1);
  });
  await t.test("function migration reruns change no existing finance, cash, payment or pickup records",async()=>{
    const tables=['purchase_payments','purchases','instant_pickups','home_cash_transactions','profit_payouts','order_profit_accruals'];
    const snapshot=Object.fromEntries(await Promise.all(tables.map(async table=>[table,await rows(table)])));
    await db.exec('reset role;');await db.exec(migration);
    for(const table of tables)assert.deepEqual(await rows(table),snapshot[table]);
  });
});

test("full receipt client verifies responses and sends no caller-provided payment amount",async()=>{
  const purchase={id:id(100),paid_price:42,picked_up:true,picked_up_at:'2026-11-09T10:00:00Z',received_amount:42};
  let args;
  await confirmFullPickup({rpc:async(name,params)=>{args={name,params};return{data:{purchase,changed_receipt:true}};}},{requestId:id(800),purchase});
  assert.deepEqual(args,{name:'confirm_full_purchase_receipt',params:{p_request_id:id(800),p_id:id(100),p_instant:false}});
  for(const result of [null,{}, {purchase:{...purchase,picked_up:false},changed_receipt:true}, {purchase:{...purchase,received_amount:5},changed_receipt:true}]) {
    await assert.rejects(confirmFullPickup({rpc:async()=>({data:result})},{requestId:id(800),purchase}));
  }
  await assert.rejects(confirmFullPickup({rpc:async()=>({error:{code:'PGRST202'}})},{requestId:id(800),purchase}),/الاستلام الكامل/);
});

test("quick tick saves once, notifies only on new receipt and keeps the retry ID after failure",async()=>{
  const source=await readFile(new URL('../src/components/pickup/PurchasePaymentControl.jsx',import.meta.url),'utf8');
  const start=source.indexOf('  async function tickPickup() {'),end=source.indexOf('\n  function openPartialPayment',start);
  for(const mode of ['success','unchanged','error','partial','already-picked']) {
    const calls=[];let error='';
    const context={busyRef:{current:false},canPay:true,canConfirmPickup:false,partialPayment:mode==='partial',
      purchase:{id:id(100),picked_up:mode==='already-picked',customer_name:'Tala',paid_price:42},instant:false,sb:{},
      fullRequestRef:{current:id(800)},setBusy:()=>{},setError:value=>{error=value;},setOpen:()=>calls.push('open'),
      confirmFullPickup:async()=>{calls.push('save');if(mode==='error')throw Error('Save failed');return{purchase:{paid_price:42},changed_receipt:mode==='success'};},
      agreedAmount:row=>row.paid_price,buildPickupStatusMessage:args=>args,notifyPickupStatus:async()=>calls.push('notify'),onSaved:async()=>calls.push('refresh')};
    await vm.runInNewContext(`${source.slice(start,end)};tickPickup`,context)();
    if(['partial','already-picked'].includes(mode)){assert.deepEqual(calls,['open']);continue;}
    assert.equal(calls.filter(c=>c==='save').length,1);
    assert.equal(calls.filter(c=>c==='notify').length,mode==='success'?1:0);
    assert.equal(error,mode==='error'?'Save failed':'');
    assert.equal(context.fullRequestRef.current,mode==='error'?id(800):'');
  }
  for(const name of ['HomePickupPage','PickupPointPage']) {
    assert.match(await readFile(new URL(`../src/pages/${name}.jsx`,import.meta.url),'utf8'),/\[showAllOrdersMode, setShowAllOrdersMode\] = useState\(true\)/);
  }
});
