import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { paymentDatabase, sqlFile } from "./helpers/payment-database.mjs";

test("the 31 confirmed historical settlements never change cash or finance", async t => {
  const { db } = await paymentDatabase();
  t.after(() => db.close());
  for (const file of ["20261025040000_add_order_misc_deduction.sql", "20261104000000_apply_full_party_profit_deductions.sql",
    "20261106000000_location_readiness_and_pickup_actor.sql", "20261106010000_home_recipients_and_cash_log_views.sql"]) {
    await db.exec(await sqlFile(file));
  }
  const script = await readFile(new URL("../supabase/maintenance/repair_confirmed_historical_pickups_20261109.sql", import.meta.url), "utf8");
  // Let PostgreSQL parse the target manifest, avoiding a second hand-maintained list.
  await db.exec(script.slice(0, script.indexOf("-- Snapshot the finance sources")));
  const targets = (await db.query('select * from pg_temp.confirmed_pickup_repairs order by number')).rows;
  await db.exec('rollback;');
  assert.equal(targets.length, 31);
  assert.deepEqual(targets.map(r => r.number), [23,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,48,49,50,54,55,56,57,58]);
  assert.equal(targets.reduce((sum,r) => sum + Number(r.amount),0), 4306);
  for (const row of targets) {
    await db.query(`insert into public.orders(id,order_name,created_at,home_profit_percent,rahaf_profit_percent,mira_profit_percent)
      values(md5($1)::uuid,$1,'2026-06-01',35,30,35) on conflict do nothing`, [row.order_name]);
    await db.query(`insert into public.purchases(id,order_id,customer_name,price,paid_price,pickup_point)
      values($1,md5($2)::uuid,$3,$4,$5,$6)`, [row.purchase_id,row.order_name,row.customer_name,Number(row.amount)+10,row.amount,row.pickup_point]);
  }
  await db.exec("select set_config('request.jwt.claim.sub','',false);");
  const state = async () => {
    const result = {};
    for (const [table, sort] of [["orders","id"],["purchases","id"],["purchase_payments","id"],
      ["home_cash_transactions","id"],["profit_payouts","id"],["order_profit_accruals","order_id"]]) {
      result[table] = (await db.query(`select * from public.${table} order by ${sort}`)).rows;
    }
    return result;
  };
  const rejected = async pattern => { try { await assert.rejects(db.exec(script),pattern); } finally { await db.exec('rollback;'); } };

  await t.test("a changed final purchase rolls back all earlier corrections", async () => {
    await db.query('update public.purchases set paid_price=paid_price+1 where id=$1',[targets.at(-1).purchase_id]);
    const before=await state();
    await rejected(/Purchase #58 changed/);
    assert.deepEqual(await state(),before);
    await db.query('update public.purchases set paid_price=$1 where id=$2',[targets.at(-1).amount,targets.at(-1).purchase_id]);
  });
  await t.test("an unexpected cash-writing trigger aborts and rolls back the entire repair", async () => {
    await db.exec(`create function public.test_unsafe_cash_trigger() returns trigger language plpgsql as $$
      begin insert into public.home_cash_transactions(kind,amount,source,note) values('receipt',1,'home','Unexpected trigger'); return new; end; $$;
      create trigger test_unsafe_cash after update of picked_up on public.purchases
      for each row when (new.picked_up and not old.picked_up) execute function public.test_unsafe_cash_trigger();`);
    const before=await state();
    await rejected(/Finance protection check failed/);
    assert.deepEqual(await state(),before);
    await db.exec('drop trigger test_unsafe_cash on public.purchases; drop function public.test_unsafe_cash_trigger();');
  });
  await t.test("a change to collection totals for an October-onward finance order also aborts", async () => {
    await db.query("update public.orders set created_at='2026-10-15' where id=md5($1)::uuid",[targets.at(-1).order_name]);
    const before=await state();
    await rejected(/Finance protection check failed/);
    assert.deepEqual(await state(),before);
    await db.query("update public.orders set created_at='2026-06-01' where id=md5($1)::uuid",[targets.at(-1).order_name]);
  });
  await t.test("all 31 settle with matching historical payments, including two zero-price purchases", async () => {
    const before=await state();
    await db.exec(script);
    const after=await state();
    for (const table of ['orders','home_cash_transactions','profit_payouts','order_profit_accruals']) assert.deepEqual(after[table],before[table]);
    const selected=new Set(targets.map(r=>r.purchase_id));
    assert.deepEqual(after.purchases.filter(p=>!selected.has(p.id)),before.purchases.filter(p=>!selected.has(p.id)));
    for (const target of targets) {
      const p=after.purchases.find(p=>p.id===target.purchase_id);
      assert.equal(p.picked_up,true);assert.equal(p.collected,true);
      assert.equal(Number(p.received_amount),Number(target.amount));assert.equal(Number(p.remitted_amount),Number(target.amount));
      assert.equal(p.picked_up_at,null);assert.equal(p.picked_up_by_email,null);
      const payments=after.purchase_payments.filter(payment=>payment.purchase_id===p.id);
      assert.equal(payments.length,1);
      assert.equal(payments[0].legacy,true);assert(payments[0].remitted_at);
      assert.equal(Number(payments[0].cash_recognized_amount),0);assert.equal(payments[0].cash_transaction_id,null);
    }
  });
  await t.test("rerunning is an exact no-op, with no new payments or changed dates", async () => {
    const before=await state();
    await db.exec(script);
    assert.deepEqual(await state(),before);
  });
});
