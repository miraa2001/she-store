import assert from "node:assert/strict";
import test from "node:test";
import { paymentDatabase, id, sqlFile, HOME, NABLUS } from "./helpers/payment-database.mjs";

test("abandoned inventory preserves original finances, restricts access and supports whole-purchase resale", async t => {
  const { db, identity }=await paymentDatabase();
  t.after(()=>db.close());
  await db.exec(`create table public.purchase_images(id uuid primary key,purchase_id uuid references public.purchases on delete cascade,storage_path text);
    create table public.purchase_links(id uuid primary key,purchase_id uuid references public.purchases on delete cascade,url text);
    grant select,insert,update,delete on public.purchase_images,public.purchase_links to authenticated;
    grant select on public.orders to authenticated;
    insert into public.purchase_images values('${id(400)}','${id(100)}','original/photo.jpg');
    insert into public.purchase_links values('${id(401)}','${id(100)}','https://example.com/item');`);
  for(const name of ['20261106000000_location_readiness_and_pickup_actor.sql','20261109000000_add_full_pickup_confirmation.sql'])await db.exec(await sqlFile(name));
  const migration=await sqlFile('20261109010000_add_abandoned_purchase_archive.sql');
  const rows=async table=>(await db.query(`select * from public.${table} order by 1`)).rows;
  const finance=async()=>{
    const role=(await db.query('select current_user as role')).rows[0].role;
    await db.exec('reset role;');
    try {
      const snapshot={};
      for(const table of ['orders','order_profit_accruals','profit_payouts','home_cash_transactions','home_cash_receipts'])snapshot[table]=await rows(table);
      return snapshot;
    } finally { if(role==='authenticated')await db.exec('set role authenticated;'); }
  };
  const initial=await finance();await db.exec(migration);assert.deepEqual(await finance(),initial);
  const archive=async(request,ids)=>(await db.query('select public.archive_abandoned_purchases($1,$2) as result',[id(request),ids.map(id)])).rows[0].result;
  const restore=async(request,ids)=>(await db.query('select public.restore_abandoned_purchases($1,$2) as result',[id(request),ids])).rows[0].result;
  const resell=async(request,archiveId,location=HOME,price=25)=>(await db.query('select public.resell_abandoned_purchase($1,$2,$3,$4,$5) as result',[id(request),archiveId,'New customer',price,location])).rows[0].result;
  const purchase=async n=>(await db.query('select * from public.purchases where id=$1',[id(n)])).rows[0];
  let first;
  await t.test("only Rahaf can archive, restore, resell or read private archive/audit data",async()=>{
    for(const actor of [2,3,4,5,6,7]){await identity(actor);await assert.rejects(archive(800,[100]),/رهف/);await assert.rejects(restore(800,[id(999)]),/رهف/);await assert.rejects(resell(800,id(999)),/رهف/);}
    await identity(1);
  });
  await t.test("a batch is atomic and rejects received, paid, missing and unassigned purchases",async()=>{
    const before=await finance();
    await assert.rejects(archive(800,[100,104]),/غير المستلمة/);
    await assert.rejects(archive(800,[100,999]));
    await db.exec(`reset role;insert into public.purchases(id,order_id,customer_name,price,pickup_point)
      values('${id(130)}',null,'Unassigned',10,'${HOME}');`);
    await identity(1);await assert.rejects(archive(800,[130]));
    assert.equal((await rows('abandoned_purchase_archives')).length,0);
    assert.deepEqual(await finance(),before);
  });
  await t.test("archiving preserves source IDs, order, price, links, images and all cash/profits",async()=>{
    const before=await finance();
    first=await archive(800,[100,101]);assert.equal(first.purchase_count,2);
    for(const n of [100,101]){assert((await purchase(n)).abandoned_at);assert.equal((await purchase(n)).ready_for_pickup,false);}
    const a=(await rows('abandoned_purchase_archives')).find(a=>a.purchase_id===id(100));
    assert.equal(a.purchase_snapshot.customer_name,'Tala');assert.equal(a.purchase_snapshot.order_id,id(10));
    assert.equal(a.images[0].storage_path,'original/photo.jpg');assert.equal(a.links[0].url,'https://example.com/item');
    assert.deepEqual(await finance(),before);
    assert.deepEqual(await archive(800,[101,100]),first);
    await assert.rejects(archive(800,[102]),/مختلفة/);
    await assert.rejects(archive(801,[100]),/الأرشفة/);
  });
  await t.test("archived sources cannot be received, changed, deleted or stripped of metadata",async()=>{
    await assert.rejects(db.query('select public.confirm_full_purchase_receipt($1,$2,false)',[id(802),id(100)]),/جاهز/);
    for(const sql of ["update public.purchases set abandoned_at=null", "update public.purchases set paid_price=100", "update public.purchases set customer_name='Other'", "delete from public.purchases"]){
      await assert.rejects(db.exec(`${sql} where id='${id(100)}'`));
    }
    await assert.rejects(db.exec(`delete from public.purchase_images where purchase_id='${id(100)}'`));
    await assert.rejects(db.exec(`update public.purchase_links set url='https://other.test' where purchase_id='${id(100)}'`));
    await identity(2);assert.equal((await rows('abandoned_purchase_archives')).length,0);assert.equal((await rows('abandoned_purchase_operations')).length,0);await identity(1);
    await assert.rejects(db.exec(`delete from public.abandoned_purchase_archives where purchase_id='${id(100)}'`),/permission denied/);
    await assert.rejects(db.exec(`update public.abandoned_purchase_archives set status='restored' where purchase_id='${id(100)}'`),/permission denied/);
  });
  await t.test("partial and cancelled payment histories cannot be archived even when balances return to zero",async()=>{
    await db.query('select public.record_purchase_payment($1,$2,$3,true,false)',[id(820),id(102),10]);
    await assert.rejects(archive(821,[102]),/غير المستلمة/);
    const payment=(await db.query('select id from public.purchase_payments where purchase_id=$1 and cancelled_at is null',[id(102)])).rows[0];
    await db.query('select public.cancel_purchase_payment($1,$2)',[payment.id,'Test cancellation']);
    assert.equal(Number((await purchase(102)).received_amount),0);
    await assert.rejects(archive(822,[102]),/سجل دفعات/);
  });
  await t.test("undo uses the current order status and supports safe repeated archive cycles",async()=>{
    const a=(await rows('abandoned_purchase_archives')).find(a=>a.purchase_id===id(101));
    await db.exec(`reset role;update public.orders set arrived=false,placed_at_pickup=false where id='${id(10)}';`);
    assert.equal((await purchase(100)).ready_for_pickup,false);
    await identity(1);await restore(803,[a.id]);
    assert.equal((await purchase(101)).abandoned_at,null);assert.equal((await purchase(101)).ready_for_pickup,false);
    await restore(803,[a.id]);
    await archive(800,[100,101]);assert.equal((await purchase(101)).abandoned_at,null);
    const again=await archive(804,[101]);assert.notEqual(again.archive_ids[0],a.id);
    await db.exec(`reset role;update public.orders set arrived=true,placed_at_pickup=true where id='${id(10)}';`);await identity(1);
    assert.equal((await purchase(100)).ready_for_pickup,false);
    assert.equal((await purchase(101)).ready_for_pickup,false);
  });
  await t.test("resale creates a separate ready/unreceived instant pickup and does not enter order profits",async()=>{
    const a=(await rows('abandoned_purchase_archives')).find(a=>a.purchase_id===id(100)&&a.status==='available');
    const before=await finance();const sale=await resell(805,a.id);
    assert.equal(sale.customer_name,'New customer');assert.equal(Number(sale.price),25);assert.equal(sale.picked_up,false);assert.equal(sale.collected,false);
    assert.equal((await purchase(100)).customer_name,'Tala');assert.equal((await purchase(100)).order_id,id(10));
    assert.deepEqual(await finance(),before);
    assert.equal((await resell(805,a.id)).id,sale.id);
    await assert.rejects(resell(806,a.id));await assert.rejects(restore(807,[a.id]));
    await db.query('select public.confirm_full_purchase_receipt($1,$2,true)',[id(808),sale.id]);
    const after=await finance();
    assert.deepEqual(after.order_profit_accruals,before.order_profit_accruals);
    const cash=Number((await db.query('select public.get_home_cash_summary() as value')).rows[0].value.balance);assert.equal(cash,25);
    await assert.rejects(db.query('select public.delete_instant_pickup($1)',[sale.id]));
  });
  await t.test("deleting an unpaid/unreceived resale returns inventory; future Nablus resale defers cash",async()=>{
    const a=(await rows('abandoned_purchase_archives')).find(a=>a.purchase_id===id(101)&&a.status==='available');
    const sale=await resell(809,a.id,NABLUS,20);
    await db.query('select public.delete_instant_pickup($1)',[sale.id]);
    const available=(await rows('abandoned_purchase_archives')).find(row=>row.id===a.id);assert.equal(available.status,'available');assert.equal(available.instant_pickup_id,null);
    assert((await rows('abandoned_purchase_operations')).some(op=>op.kind==='cancel_resale'));
    const second=await resell(810,a.id,NABLUS,20);
    await identity(5);await db.query('select public.confirm_full_purchase_receipt($1,$2,true)',[id(811),second.id]);
    await identity(1);
    assert.equal(Number((await db.query('select public.get_home_cash_summary() as value')).rows[0].value.balance),25);
    await db.query('select public.collect_purchase_payments($1,true)',[[second.id]]);
    assert.equal(Number((await db.query('select public.get_home_cash_summary() as value')).rows[0].value.balance),45);
  });
  await t.test("migration reruns do not change existing records, snapshots, payments, cash or profits",async()=>{
    const tables=['purchases','instant_pickups','purchase_payments','abandoned_purchase_archives','abandoned_purchase_operations'];
    const before=await finance(),snap={};for(const table of tables)snap[table]=await rows(table);
    await db.exec('reset role;');await db.exec(migration);
    assert.deepEqual(await finance(),before);for(const table of tables)assert.deepEqual(await rows(table),snap[table]);
  });
});
