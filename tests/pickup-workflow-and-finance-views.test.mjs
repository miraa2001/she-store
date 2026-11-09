import assert from "node:assert/strict";
import test from "node:test";
import { paymentDatabase, id, sqlFile, DELIVERY } from "./helpers/payment-database.mjs";

const migrationNames = ["20261106000000_location_readiness_and_pickup_actor.sql", "20261106010000_home_recipients_and_cash_log_views.sql"];

test("SQL Editor readiness migration skips unassigned drafts without weakening their guard", async t => {
  const { db, identity } = await paymentDatabase();
  t.after(() => db.close());
  await db.exec(`create table public.purchase_links(id uuid primary key,purchase_id uuid);
    create table public.purchase_images(id uuid primary key,purchase_id uuid);
    update public.orders set arrived=true,placed_at_pickup=false;`);
  await db.exec(await sqlFile("20261025010000_add_unassigned_purchases.sql"));
  await db.exec(`insert into public.purchases(id,order_id,customer_name,price,pickup_point,ready_for_pickup)
    values('${id(300)}',null,'Unassigned draft',25,'من البيت',false);`);
  const draft = async () => (await db.query('select id,order_id,customer_name,price,pickup_point,ready_for_pickup,ready_for_pickup_at,received_amount,remitted_amount from public.purchases where id=$1', [id(300)])).rows[0];
  const before = await draft();
  const cash = (await db.query('select * from public.home_cash_transactions order by id')).rows;
  await db.exec("select set_config('request.jwt.claim.sub','',false);");
  // SQL Editor has no application identity, even when its database role is postgres.
  await assert.rejects(db.exec('update public.purchases set ready_for_pickup=ready_for_pickup;'), /المشتريات غير المعينة متاحة لرهف فقط/);
  const migration = await sqlFile(migrationNames[0]);
  await db.exec(migration);
  await db.exec(migration);
  assert.deepEqual(await draft(), before);
  assert.deepEqual((await db.query('select * from public.home_cash_transactions order by id')).rows
    .map(({ pickup_actor_email, ...entry }) => entry), cash);
  assert.equal((await db.query('select ready_for_pickup from public.purchases where id=$1', [id(100)])).rows[0].ready_for_pickup, true);
  assert.equal((await db.query('select ready_for_pickup from public.purchases where id=$1', [id(101)])).rows[0].ready_for_pickup, false);
  await identity(2);
  assert.equal((await db.query('select id from public.purchases where id=$1', [id(300)])).rows.length, 0);
  await db.exec('reset role;');
  await assert.rejects(db.query("update public.purchases set customer_name='Blocked' where id=$1", [id(300)]), /المشتريات غير المعينة متاحة لرهف فقط/);
});

test("location readiness, pickup audit, Home recipients and filtered cash logs", async t => {
  const { db, identity } = await paymentDatabase();
  t.after(() => db.close());
  await db.exec(`select public.record_profit_payout('${id(700)}','home',1,false);`);
  const migrations = await Promise.all(migrationNames.map(sqlFile));
  for (const sql of migrations) await db.exec(sql);
  await db.exec('grant select,update on public.orders to authenticated;');
  const purchase = async n => (await db.query('select * from public.purchases where id=$1', [id(n)])).rows[0];
  const query = (sql, args=[]) => db.query(sql,args);
  const payment = (request,n,amount,instant=false) => query('select public.record_purchase_payment($1,$2,$3,true,$4) as value',[id(request),id(n),amount,instant]);
  const activity = async view => (await query('select * from public.get_home_cash_activity_view($1,0,100)',[view])).rows;
  const payout = (request,party,recipient,amount=1,cash=false) => query('select public.record_profit_payout($1,$2,$3,$4,$5)',[id(request),party,amount,cash,recipient]);

  await t.test("pending hides unpaid purchases; arrival readies Home/Delivery only; placed readies all", async () => {
    await db.exec(`update public.orders set arrived=false,placed_at_pickup=false where id='${id(10)}';`);
    for (const n of [100,101,102,103]) assert.equal((await purchase(n)).ready_for_pickup,false);
    await assert.rejects(payment(710,100,1), /جاهز/);
    await db.exec(`update public.orders set arrived=true where id='${id(10)}';`);
    assert.equal((await purchase(100)).ready_for_pickup,true);
    assert.equal((await purchase(102)).ready_for_pickup,true);
    for (const n of [101,103]) {
      assert.equal((await purchase(n)).ready_for_pickup,false);
      await assert.rejects(payment(710,n,1), /جاهز/);
    }
    await db.exec(`update public.orders set placed_at_pickup=true where id='${id(10)}';`);
    for (const n of [100,101,102,103]) assert.equal((await purchase(n)).ready_for_pickup,true);
    await db.exec(`update public.orders set placed_at_pickup=false where id='${id(10)}';`);
    assert.equal((await purchase(103)).ready_for_pickup,false);
    await db.exec(`insert into public.purchases(id,order_id,customer_name,price,pickup_point) values
      ('${id(110)}','${id(10)}','Added after arrival',10,'${DELIVERY}');`);
    assert.equal((await purchase(110)).ready_for_pickup,true);
  });

  await t.test("receipt actor is distinct from collection actor, cannot be spoofed, and survives later installments", async () => {
    await identity(2);
    await payment(711,100,10);
    assert.equal((await purchase(100)).picked_up_by_email,'reem@she-store.com');
    await identity(1);
    await payment(712,100,5);
    assert.equal((await purchase(100)).picked_up_by_email,'reem@she-store.com');
    await db.exec(`update public.purchases set picked_up_by_email='spoofed',picked_up_by='${id(1)}' where id='${id(100)}';`);
    assert.equal((await purchase(100)).picked_up_by_email,'reem@she-store.com');
    await db.exec(`update public.orders set placed_at_pickup=true where id='${id(10)}';`);
    await identity(5);
    await query('select public.confirm_pickup_staff_receipt($1,$2,false)',[id(713),id(101)]);
    assert.equal((await purchase(101)).picked_up_by_email,'nablus@she-store.com');
    await identity(1);
    await query('select public.collect_purchase_payments($1,false)',[[id(101)]]);
    const log=(await activity('receipts')).find(row=>row.customer_name==='Nablus customer');
    assert.equal(log.actor_email,'rahaf@she-store.com');
    assert.equal(log.pickup_actor_email,'nablus@she-store.com');
    const homeLog=(await activity('receipts')).find(row=>Number(row.amount)===10);
    assert.equal(homeLog.pickup_actor_email,'reem@she-store.com');
    await db.exec(`update public.orders set arrived=false,placed_at_pickup=false where id='${id(10)}';`);
    assert.equal((await purchase(100)).ready_for_pickup,true);
    assert.equal((await purchase(110)).ready_for_pickup,false);
    await identity(2);
    await payment(714,200,10,true);
    const instant=(await query('select picked_up_by_email from public.instant_pickups where id=$1',[id(200)])).rows[0];
    assert.equal(instant.picked_up_by_email,'reem@she-store.com');
    await identity(1);
  });

  await t.test("Home recipients use one shared balance, keep legacy payouts, and validate retries", async () => {
    const balance=async()=>Number((await query("select balance from public.get_profit_balances() where party='home'")).rows[0].balance);
    const before=await balance();
    for (const [index,recipient] of ['reem','rawand','mama'].entries()) await payout(720+index,'home',recipient);
    assert.equal(await balance(),before-3);
    assert.equal((await query('select home_recipient from public.profit_payouts where id=$1',[id(700)])).rows[0].home_recipient,null);
    await query('select public.record_profit_payout($1,$2,$3,$4)',[id(700),'home',1,false]);
    await payout(720,'home','reem');
    assert.equal(await balance(),before-3);
    await assert.rejects(payout(720,'home','mama'),/مختلفة/);
    for (const recipient of [null,'unknown']) await assert.rejects(payout(730,'home',recipient),/مستلم/);
    await assert.rejects(payout(730,'mira','reem'),/مستلم/);
    await payout(723,'home','mama',1,true);
    const log=(await activity('manual')).find(row=>row.kind==='profit_payout');
    assert.equal(log.customer_name,'ماما');
    await identity(2);
    await assert.rejects(payout(730,'home','reem'),/صلاحية/);
    await identity(1);
  });

  await t.test("pickup cancellation clears the current actor but preserves historic cash actors", async () => {
    const original=(await activity('receipts')).find(row=>Number(row.amount)===10 && row.customer_name==='Tala');
    await query('select public.cancel_purchase_payment($1,$2)',[id(711),'Cancelled first payment']);
    assert.equal((await purchase(100)).picked_up_by_email,'reem@she-store.com');
    await query('select public.cancel_purchase_payment($1,$2)',[id(712),'Cancelled remaining payment']);
    assert.equal((await purchase(100)).picked_up_by_email,null);
    assert.equal((await activity('receipts')).find(row=>row.id===original.id).pickup_actor_email,'reem@she-store.com');
    await db.exec(`update public.orders set arrived=true where id='${id(10)}';`);
    await payment(750,100,1);
    assert.equal((await purchase(100)).picked_up_by_email,'rahaf@she-store.com');
    assert.equal((await activity('receipts')).find(row=>row.id===original.id).pickup_actor_email,'reem@she-store.com');
  });

  await t.test("two filtered views cover all events and paginate after global running balances", async () => {
    await query("select public.record_home_cash_expense($1,'bags',2,1,'Test')",[id(740)]);
    await query('select public.cancel_home_cash_expense($1)',[id(740)]);
    const receipts=await activity('receipts'), manual=await activity('manual');
    assert(receipts.every(row=>['receipt','receipt_adjustment'].includes(row.kind)));
    assert(manual.every(row=>!['receipt','receipt_adjustment'].includes(row.kind)));
    assert(manual.some(row=>row.kind==='expense_reversal'));
    assert(manual.some(row=>row.kind==='tracking_started'));
    const global=(await query('select * from public.get_home_cash_activity(0,100)')).rows;
    assert.equal(receipts.length+manual.length,global.length);
    for (const row of [...receipts,...manual]) assert.equal(row.balance_after,global.find(item=>item.id===row.id).balance_after);
    const second=(await query("select * from public.get_home_cash_activity_view('receipts',1,1)")).rows;
    assert.equal(second[0].id,receipts[1].id);
    await identity(5);
    assert.equal((await activity('manual')).length,0);
    await identity(1);
  });

  await t.test("migration reruns preserve actors, cash, payouts and the existing balance function", async () => {
    const actorBefore=(await purchase(100)).picked_up_by_email;
    const before=(await query('select * from public.get_profit_balances()')).rows;
    const cash=(await query('select * from public.home_cash_transactions order by id')).rows;
    const payouts=(await query('select * from public.profit_payouts order by id')).rows;
    await db.exec('reset role;');
    for (const sql of migrations) await db.exec(sql);
    assert.deepEqual((await query('select * from public.get_profit_balances()')).rows,before);
    assert.deepEqual((await query('select * from public.home_cash_transactions order by id')).rows,cash);
    assert.deepEqual((await query('select * from public.profit_payouts order by id')).rows,payouts);
    assert.equal((await purchase(100)).picked_up_by_email,actorBefore);
  });
});
