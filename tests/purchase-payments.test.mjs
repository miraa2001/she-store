import assert from "node:assert/strict";
import test from "node:test";
import vm from "node:vm";
import { readFile } from "node:fs/promises";
import { paymentDatabase,id,HOME,NABLUS } from "./helpers/payment-database.mjs";
import { paidAmount,remainingAmount,collectibleAmount,canRecordPayment,recordPayment,collectPayments } from "../src/lib/payments.js";

test("partial purchase payments, cash and permission boundaries",async(t)=>{
  const {db,migration,identity}=await paymentDatabase();
  t.after(()=>db.close());
  const pay=async(request,n,amount,receive=false,instant=false)=>(await db.query('select public.record_purchase_payment($1,$2,$3,$4,$5) as result',[id(request),id(n),amount,receive,instant])).rows[0].result;
  const state=async(n,instant=false)=>(await db.query('select public.get_purchase_payment_state($1,$2) as result',[id(n),instant])).rows[0].result;
  const collect=async(ns,instant=false)=>(await db.query('select public.collect_purchase_payments($1,$2) as result',[ns.map(id),instant])).rows[0].result;
  const cancel=async(n)=>(await db.query('select public.cancel_purchase_payment($1,$2) as result',[id(n),'Entry mistake'])).rows[0].result;
  const cash=async()=>{await identity(1); return Number((await db.query('select public.get_home_cash_summary() as result')).rows[0].result.balance);};
  const profit=async()=>{await identity(1);return (await db.query('select purchase_value,distributable_profit from public.order_profit_accruals')).rows[0];};
  let startingProfit=await profit();

  await t.test("migration baselines received purchases without adding cash",async()=>{
    assert.equal(await cash(),0);
    assert.equal(Number((await state(104)).purchase.received_amount),30);
    assert.equal((await state(104)).purchase.collected,true);
    assert.equal(Number((await state(105)).purchase.received_amount),20);
    assert.equal(Number((await state(105)).purchase.remitted_amount),0);
  });
  await t.test("Home first installment marks pickup and adds only actual payment",async()=>{
    await identity(2);
    const result=await pay(300,100,12.25,true);
    assert.equal(result.changed_receipt,true);
    assert.equal(result.purchase.picked_up,true); assert.equal(result.purchase.collected,false);
    assert.equal(Number(result.purchase.received_amount),12.25);
    assert.equal(Number(result.purchase.remitted_amount),12.25);
    assert.equal(await cash(),12.25);
    assert.deepEqual(await profit(),startingProfit);
  });
  await t.test("retry does not double cash and differing idempotency payload fails",async()=>{
    await identity(2);
    assert.equal((await pay(300,100,12.25,true)).changed_receipt,false);
    await assert.rejects(pay(300,100,10,true),/مختلفة/);
    assert.equal(await cash(),12.25);
  });
  await t.test("later Home payment completes the purchase with cent precision",async()=>{
    await identity(2);
    const result=await pay(301,100,29.75);
    assert.equal(result.purchase.collected,true);
    assert.equal(Number(result.purchase.received_amount),42);
    assert.equal(await cash(),42);
  });
  await t.test("Nablus installment can be collected while customer still owes money",async()=>{
    await identity(2); await pay(302,101,25,true);
    assert.equal(await cash(),42);
    assert.equal(Number(await collect([101])),25);
    assert.equal(Number(await collect([101])),0);
    const target=(await state(101)).purchase;
    assert.equal(target.collected,false);assert.equal(Number(target.remitted_amount),25);
    assert.equal(await cash(),67);
    await identity(2); await pay(303,101,75);
    assert.equal((await state(101)).purchase.collected,false);
    assert.equal(await cash(),67);
    await collect([101]); assert.equal((await state(101)).purchase.collected,true);
    assert.equal(await cash(),142);
  });
  await t.test("delivery and Maryamti also defer cash until collection",async()=>{
    await identity(2); await pay(304,102,10,true);
    await assert.rejects(pay(305,103,10,true),/متاح/);
    await identity(1); await pay(305,103,10,true);
    assert.equal(await cash(),142);
    assert.equal(Number(await collect([102,103])),20);
    assert.equal(await cash(),162);
  });
  await t.test("only Rahaf/Reem record; only Rahaf collects/cancels",async()=>{
    for(const actor of [3,4,5,6,7]){
      await identity(actor); await assert.rejects(pay(320+actor,102,1),/متاح/);
      await assert.rejects(collect([102]),/رهف/);await assert.rejects(cancel(304),/رهف/);
    }
    await identity(2);await assert.rejects(collect([102]),/رهف/);await assert.rejects(cancel(304),/رهف/);
  });
  await t.test("zero, negative, fractional-cent, excessive and premature payments fail",async()=>{
    await identity(2);
    for(const amount of [0,-1,1.001,'NaN','Infinity']) await assert.rejects(pay(350,102,amount));
    await assert.rejects(pay(350,102,71),/المتبقي/);
    await identity(1);
    await db.exec(`insert into public.purchases(id,order_id,customer_name,price,pickup_point,ready_for_pickup) values('${id(110)}','${id(10)}','Not ready',10,'${HOME}',false),('${id(111)}','${id(10)}','Unreceived',10,'${HOME}',true)`);
    const afterInsert=await profit();
    assert.equal(Number(afterInsert.purchase_value),Number(startingProfit.purchase_value)+20);
    startingProfit=afterInsert;
    await assert.rejects(pay(350,110,1,true),/جاهز/);
    await assert.rejects(pay(350,111,1,false),/أول دفعة/);
  });
  await t.test("old checkbox and collection paths cannot invent payments or cash",async()=>{
    await identity(2); await assert.rejects(db.query('select public.set_purchase_received($1,true)',[id(111)]),/دفعة/);
    await identity(5); await assert.rejects(db.query('update public.purchases set picked_up=false where id=$1',[id(101)]));
    await identity(1); await db.query('update public.purchases set collected=true where id=$1',[id(102)]);
    assert.equal((await state(102)).purchase.collected,false);
    assert.equal(await cash(),162);
  });
  await t.test("transfers, price reductions, readiness resets and deletion protect payment history",async()=>{
    await identity(1);
    await assert.rejects(db.query('update public.purchases set pickup_point=$2 where id=$1',[id(102),HOME]),/نقل/);
    await assert.rejects(db.query('update public.purchases set paid_price=5 where id=$1',[id(102)]),/السعر/);
    await assert.rejects(db.query('update public.purchases set ready_for_pickup=false where id=$1',[id(102)]),/الجاهزية/);
    await assert.rejects(db.query('delete from public.purchases where id=$1',[id(102)]),/حذف/);
  });
  await t.test("instant pickup uses identical payment, collection and transfer rules",async()=>{
    await identity(2); await pay(360,200,20,true,true);
    assert.equal(await cash(),182);
    await identity(2); await pay(361,201,30,true,true);
    await assert.rejects(pay(362,202,10,true,true));
    await identity(1);await assert.rejects(db.query("select public.update_instant_pickup($1,'transfer',$2)",[id(201),HOME]),/نقل/);
    await assert.rejects(db.query("select public.update_instant_pickup($1,'undo_receive')",[id(200)]));
    assert.equal(Number(await collect([201],true)),30);
    assert.equal((await state(201,true)).purchase.collected,false);
    assert.equal(await cash(),212);
  });
  await t.test("cancellation preserves audit history, reverses only recognized cash, and retries safely",async()=>{
    await identity(1);const result=await cancel(361);
    assert.equal(Number(result.received_amount),0); assert.equal(result.picked_up,false);
    assert.equal(await cash(),182);
    await cancel(361);assert.equal(await cash(),182);
    const history=(await state(201,true)).payments;
    assert.equal(history.length,1);assert(history[0].cancelled_at);assert.equal(history[0].cancellation_reason,'Entry mistake');
    await assert.rejects(db.query("select public.update_instant_pickup($1,'transfer',$2)",[id(201),HOME]),/نقل/);
    await identity(2);await pay(370,102,5);
    await identity(1);await cancel(370);assert.equal(await cash(),182);
    await cancel(300);assert.equal(await cash(),169.75);
  });
  await t.test("historical excluded cash is not reversed; old uncollected money can still be collected",async()=>{
    await identity(1);
    const legacy=(await state(104)).payments[0];
    await db.query('select public.cancel_purchase_payment($1,$2)',[legacy.id,'Historical correction']);
    assert.equal(await cash(),169.75);
    assert.equal(Number(await collect([105])),20);assert.equal(await cash(),189.75);
  });
  await t.test("RLS/private functions block ledger forgery and cross-location access",async()=>{
    await identity(2);await assert.rejects(db.query('select public.get_purchase_payment_state($1,false)',[id(103)]));
    await assert.rejects(db.exec("update public.purchase_payments set amount=1"));
    await assert.rejects(db.query('select public.refresh_purchase_payment_totals($1,false,true)',[id(111)]));
    await assert.rejects(db.query('update public.purchases set received_amount=1 where id=$1',[id(111)]).then(r=>{if(r.affectedRows===0) throw Error('RLS filtered');}));
    await db.exec('reset role; set role anon;');await assert.rejects(pay(380,111,1,true));
  });
  await t.test("rerun never resets cash/payments and installments never change profit",async()=>{
    await db.exec('reset role;');await db.exec(migration);
    assert.equal(await cash(),189.75);
    assert.deepEqual(await profit(),startingProfit);
  });
});

test("payment client validates cents and confirms results",async()=>{
  const row={id:id(100),paid_price:42,received_amount:12.25,remitted_amount:5};
  assert.equal(paidAmount(row),12.25);assert.equal(remainingAmount(row),29.75);assert.equal(collectibleAmount(row),7.25);
  assert(canRecordPayment('reem'));assert(!canRecordPayment('nablus'));
  for(const amount of ['0','-1','1.001','Infinity','']) await assert.rejects(recordPayment({}, {purchase:row,amount}));
  const client={rpc:async()=>({data:{purchase:row,changed_receipt:false}})};
  assert.equal((await recordPayment(client,{requestId:id(900),purchase:row,amount:'9.76'})).purchase.id,row.id);
  await assert.rejects(recordPayment({rpc:async()=>({data:{}})},{purchase:row,amount:'1'}));
  assert.equal(await collectPayments({rpc:async()=>({data:0})},[row.id]),0);
});

test("upgrade preserves money already counted under old receipt triggers",async(t)=>{
  const {db,identity}=await paymentDatabase({trackedBeforeUpgrade:true});
  t.after(()=>db.close());await identity(1);
  const balance=async()=>Number((await db.query('select public.get_home_cash_summary() as value')).rows[0].value.balance);
  assert.equal(await balance(),112);
  const legacy=(await db.query('select public.get_purchase_payment_state($1,false) as value',[id(100)])).rows[0].value.payments[0];
  assert.equal(Number(legacy.cash_recognized_amount),42);
  await db.query('select public.cancel_purchase_payment($1,$2)',[legacy.id,'Old receipt was a mistake']);
  assert.equal(await balance(),70);
});

test("payment dialog notifies only after a confirmed first receipt, and retains retry IDs on failure",async()=>{
  const source=await readFile(new URL('../src/components/pickup/PurchasePaymentControl.jsx',import.meta.url),'utf8');
  const start=source.indexOf('  async function save(event) {');
  const end=source.indexOf('\n  async function cancelPayment',start);
  for(const mode of ['success','unchanged','error']) {
    const calls=[];let error='',busy=false;
    const context={
      state:{},canPay:true,busyRef:{current:false},requestRef:{current:id(900)},sb:{},instant:false,
      active:{id:id(100),picked_up:false,customer_name:'Tala',pickup_point:HOME,paid_price:42},amount:'12.25',
      setBusy:value=>{busy=value;},setError:value=>{error=value;},setOpen:value=>calls.push(['open',value]),
      agreedAmount:row=>row.paid_price,buildPickupStatusMessage:args=>args,
      recordPayment:async()=>{calls.push(['save']);if(mode==='error')throw Error('Save failed');return{purchase:{id:id(100)},changed_receipt:mode==='success'};},
      notifyPickupStatus:async()=>calls.push(['notify']),onSaved:async()=>calls.push(['refresh'])
    };
    const save=vm.runInNewContext(`${source.slice(start,end)};save`,context);
    await save({preventDefault(){}});
    assert.equal(calls.filter(c=>c[0]==='notify').length,mode==='success'?1:0);
    assert.equal(calls.filter(c=>c[0]==='refresh').length,mode==='error'?0:1);
    assert.equal(calls[0][0],'save');assert.equal(error,mode==='error'?'Save failed':'');assert.equal(busy,false);
    assert.equal(context.requestRef.current,id(900));
  }
});
