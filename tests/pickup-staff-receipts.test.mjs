import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { paymentDatabase,id } from "./helpers/payment-database.mjs";
import { canConfirmStaffPickup,canRecordPayment,confirmStaffPickup } from "../src/lib/payments.js";

test("staff ordinary pickups remain separate from custom payment permissions",async(t)=>{
  const {db,identity,staffMigration}=await paymentDatabase();
  t.after(()=>db.close());
  const confirm=async(request,n,instant=false)=>(await db.query('select public.confirm_pickup_staff_receipt($1,$2,$3) as value',[id(request),id(n),instant])).rows[0].value;
  const row=async(n,instant=false)=>(await db.query('select public.get_purchase_payment_state($1,$2) as value',[id(n),instant])).rows[0].value;
  const cash=async()=>{await identity(1);return Number((await db.query('select public.get_home_cash_summary() as value')).rows[0].value.balance);};
  let profit;
  await identity(1);profit=(await db.query('select purchase_value,distributable_profit from public.order_profit_accruals')).rows[0];

  await t.test("Maryamti can confirm full paid pickup and it persists, without posting Home cash",async()=>{
    await identity(4);
    const result=await confirm(800,103);
    assert.equal(result.changed_receipt,true);
    assert.equal(result.purchase.picked_up,true);assert(result.purchase.picked_up_at);
    assert.equal(Number(result.purchase.received_amount),60);
    assert.equal(Number(result.purchase.remitted_amount),0);assert.equal(result.purchase.collected,false);
    const state=await row(103);assert.equal(state.purchase.picked_up,true);
    assert.equal(state.payments.length,1);assert.equal(state.payments[0].created_by,id(4));
    assert.equal(await cash(),0);
  });
  await t.test("Nablus can confirm full paid pickup at Nablus",async()=>{
    await identity(5);
    const result=await confirm(801,101);
    assert.equal(Number(result.purchase.received_amount),100);
    assert.equal(result.changed_receipt,true);assert.equal(result.purchase.collected,false);
    assert.equal(await cash(),0);
  });
  await t.test("retries and another click never duplicate a payment",async()=>{
    await identity(5);
    assert.equal((await confirm(801,101)).changed_receipt,false);
    assert.equal((await confirm(802,101)).changed_receipt,false);
    assert.equal((await row(101)).payments.length,1);
    await assert.rejects(confirm(801,201,true),/مستخدم/);
  });
  await t.test("staff cannot receive other locations or use a different role's endpoint",async()=>{
    await identity(4);
    for(const n of [100,101,102])await assert.rejects(confirm(810+n,n),/الخاصة/);
    await identity(5);
    for(const n of [100,102,103])await assert.rejects(confirm(810+n,n),/الخاصة/);
    for(const actor of [1,2,3,6,7]){await identity(actor);await assert.rejects(confirm(850+actor,103),/الخاصة/);}
  });
  await t.test("custom installment entry and cancellation stay unavailable to staff",async()=>{
    for(const actor of [4,5]){
      await identity(actor);
      await assert.rejects(db.query('select public.record_purchase_payment($1,$2,1,true,false)',[id(870+actor),id(actor===4?103:101)]),/متاح/);
      await assert.rejects(db.query('select public.cancel_purchase_payment($1,$2)',[id(801),'Wrong click']),/رهف/);
      await assert.rejects(db.query('select public.confirm_pickup_staff_receipt($1,$2,false,1)',[id(870+actor),id(103)]));
    }
  });
  await t.test("instant pickups use the same location and full-payment rules",async()=>{
    await identity(4);const mary=await confirm(880,202,true);
    assert.equal(Number(mary.purchase.received_amount),70);assert.equal(mary.purchase.picked_up,true);
    await assert.rejects(confirm(881,201,true),/الخاصة/);
    await identity(5);const nab=await confirm(881,201,true);
    assert.equal(Number(nab.purchase.received_amount),70);
    assert.equal((await confirm(881,201,true)).changed_receipt,false);
    assert.equal(await cash(),0);
  });
  await t.test("staff cannot undo a paid handover or silently settle a previously partial purchase",async()=>{
    await identity(4);await assert.rejects(db.query('select public.set_purchase_received($1,false)',[id(103)]));
    await assert.rejects(db.query("select public.update_instant_pickup($1,'undo_receive')",[id(202)]));
    await identity(1);
    await db.exec(`insert into public.purchases(id,order_id,customer_name,price,pickup_point) values('${id(120)}','${id(10)}','Partial',50,'نقطة استلام نابلس');`);
    await db.query('select public.record_purchase_payment($1,$2,10,true,false)',[id(890),id(120)]);
    await identity(5);const result=await confirm(891,120);
    assert.equal(result.changed_receipt,false);assert.equal(Number(result.purchase.received_amount),10);
    assert.equal((await row(120)).payments.length,1);
  });
  await t.test("not-ready, missing, null and anonymous requests cannot receive",async()=>{
    await identity(1);
    await db.exec(`insert into public.purchases(id,order_id,customer_name,price,pickup_point,ready_for_pickup) values('${id(121)}','${id(10)}','Not ready',10,'نقطة استلام نابلس',false);`);
    await identity(5);await assert.rejects(confirm(900,121),/جاهز/);
    await assert.rejects(confirm(900,999));
    await assert.rejects(db.query('select public.confirm_pickup_staff_receipt(null,$1,false)',[id(121)]));
    await db.exec('reset role;set role anon;');await assert.rejects(confirm(900,121));
  });
  await t.test("zero-price handovers work without inventing a payment",async()=>{
    await identity(1);
    await db.exec(`insert into public.purchases(id,order_id,customer_name,price,pickup_point) values('${id(122)}','${id(10)}','Free',0,'نقطة استلام نابلس');`);
    await identity(5);const result=await confirm(901,122);
    assert.equal(result.purchase.picked_up,true);assert.equal(result.purchase.collected,true);
    assert.equal(Number(result.purchase.received_amount),0);assert.equal((await row(122)).payments.length,0);
    assert.equal((await confirm(901,122)).changed_receipt,false);
  });
  await t.test("Rahaf collection posts each full payment once and migration reruns preserve it",async()=>{
    await identity(1);
    assert.equal(Number((await db.query('select public.collect_purchase_payments($1,false) as value',[[id(101),id(103)]])).rows[0].value),160);
    assert.equal(Number((await db.query('select public.collect_purchase_payments($1,true) as value',[[id(201),id(202)]])).rows[0].value),140);
    assert.equal(await cash(),300);
    await db.exec('reset role;');await db.exec(staffMigration);await identity(4);
    assert.equal((await confirm(800,103)).changed_receipt,false);assert.equal((await row(103)).payments.length,1);
    assert.equal(await cash(),300);
    const after=(await db.query('select purchase_value,distributable_profit from public.order_profit_accruals')).rows[0];
    // The test inserted two additional nonzero purchases; pickup/payment itself changes no profits.
    assert.equal(Number(after.purchase_value),Number(profit.purchase_value)+60);
  });
});

test("staff client enables only staff receipt confirmation and rejects unverified saves",async()=>{
  assert(canConfirmStaffPickup('maryamti'));assert(canConfirmStaffPickup('nablus'));
  assert(!canConfirmStaffPickup('reem'));assert(!canRecordPayment('maryamti'));assert(!canRecordPayment('nablus'));
  const purchase={id:id(101),paid_price:100,picked_up:true,picked_up_at:'2026-10-22T10:00:00Z',received_amount:100};
  let params;
  const valid={purchase,changed_receipt:true};
  assert.equal(await confirmStaffPickup({rpc:async(name,args)=>{params={name,args};return{data:valid};}},{requestId:id(800),purchase}),valid);
  assert.deepEqual(params,{name:'confirm_pickup_staff_receipt',args:{p_request_id:id(800),p_id:id(101),p_instant:false}});
  assert(!Object.hasOwn(params.args,'p_amount'));
  await assert.rejects(confirmStaffPickup({rpc:async()=>({error:{code:'PGRST202'}})},{requestId:id(800),purchase}),/مريمتي ونابلس/);
  for(const result of [null,{}, {...valid,purchase:{...purchase,picked_up:false}}, {...valid,purchase:{...purchase,received_amount:20}}, {...valid,purchase:{...purchase,id:id(102)}}, {...valid,purchase:{...purchase,picked_up_at:null}}]){
    await assert.rejects(confirmStaffPickup({rpc:async()=>({data:result})},{requestId:id(800),purchase}));
  }
});

test("staff notification is sent only for a verified new receipt",async()=>{
  const source=await readFile(new URL('../src/components/pickup/PurchasePaymentControl.jsx',import.meta.url),'utf8');
  const start=source.indexOf('  async function saveStaffPickup() {');
  const end=source.indexOf('\n  async function cancelPayment',start);
  for(const mode of ['success','unchanged','error']){
    const calls=[];let error='',busy=false;
    const purchase={id:id(101),customer_name:'Customer',pickup_point:'نقطة استلام نابلس',paid_price:100,picked_up:false};
    const context={state:{},busyRef:{current:false},requestRef:{current:id(800)},canConfirmPickup:true,active:purchase,instant:false,sb:{},
      setBusy:value=>{busy=value;},setError:value=>{error=value;},setOpen:()=>calls.push('close'),
      confirmStaffPickup:async()=>{calls.push('save');if(mode==='error')throw Error('Save failed');return{purchase:{...purchase,picked_up:true},changed_receipt:mode==='success'};},
      agreedAmount:row=>row.paid_price,buildPickupStatusMessage:args=>args,
      notifyPickupStatus:async()=>calls.push('notify'),onSaved:async()=>calls.push('refresh')};
    const save=vm.runInNewContext(`${source.slice(start,end)};saveStaffPickup`,context);
    await save();
    assert.equal(calls[0],'save');assert.equal(calls.filter(c=>c==='notify').length,mode==='success'?1:0);
    assert.equal(calls.filter(c=>c==='refresh').length,mode==='error'?0:1);
    assert.equal(error,mode==='error'?'Save failed':'');assert.equal(busy,false);
    assert.equal(context.requestRef.current,id(800));
  }
});
