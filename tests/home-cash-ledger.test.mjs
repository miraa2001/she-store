import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { PGlite } from "@electric-sql/pglite";

const uuid = (n) => `00000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const HOME = "\u0645\u0646 \u0627\u0644\u0628\u064a\u062a";
const MARYAMTI = "\u0645\u0646 \u0646\u0642\u0637\u0629 \u0627\u0644\u0627\u0633\u062a\u0644\u0627\u0645 - \u0645\u0631\u064a\u0645\u062a\u064a";
const NABLUS = "\u0646\u0642\u0637\u0629 \u0627\u0633\u062a\u0644\u0627\u0645 \u0646\u0627\u0628\u0644\u0633";
const DELIVERY = "\u062a\u0648\u0635\u064a\u0644";

test("Home cash migration and ledger", async (t) => {
  const db = new PGlite();
  t.after(() => db.close());
  const migration = await readFile(new URL("../supabase/migrations/20261002010000_add_home_cash_ledger.sql", import.meta.url), "utf8");
  // Minimal Supabase schemas let the production SQL run in isolated PostgreSQL.
  await db.exec(`
    create role anon;
    create role authenticated;
    create schema auth;
    create table auth.users (id uuid primary key, email text);
    create function auth.uid() returns uuid language sql stable as $$
      select nullif(current_setting('request.jwt.claim.sub', true), '')::uuid;
    $$;
    create table public.user_roles (user_id uuid primary key references auth.users(id), role text);
    create function public.is_admin() returns boolean language sql stable as $$
      select exists (select 1 from public.user_roles where user_id = auth.uid() and role = 'admin');
    $$;
    create table public.orders (id uuid primary key, order_name text);
    create table public.purchases (id uuid primary key, order_id uuid references public.orders on delete cascade,
      customer_name text, pickup_point text, price numeric, paid_price numeric,
      picked_up boolean default false, collected boolean default false);
    insert into auth.users values ('${uuid(1)}', 'rahaf@she-store.com'), ('${uuid(2)}', 'pickup@she-store.com');
    insert into public.user_roles values ('${uuid(1)}', 'admin'), ('${uuid(2)}', 'pickup');
    grant usage on schema auth, public to authenticated;
    grant select on public.user_roles to authenticated;
    grant select, insert, update, delete on public.purchases to authenticated;
    select set_config('request.jwt.claim.sub', '${uuid(1)}', false);
    insert into public.orders values ('${uuid(10)}', 'October Order');
  `);
  const add = (n, location, paid, picked = false, collected = false, price = 999) => db.query(
    "insert into public.purchases (id,order_id,customer_name,pickup_point,price,paid_price,picked_up,collected) values ($1,$2,$3,$4,$5,$6,$7,$8)",
    [uuid(n), uuid(10), `Customer ${n}`, location, price, paid, picked, collected]);
  const summary = async () => (await db.query("select public.get_home_cash_summary() as value")).rows[0].value;
  const spend = (n, category, amount, quantity = 1, note = null) => db.query(
    "select public.record_home_cash_expense($1,$2,$3,$4,$5)", [uuid(n), category, amount, quantity, note]);

  await t.test("starts at zero and excludes historical receipts and their later corrections", async () => {
    await add(11, HOME, 200, true);
    await add(19, NABLUS, 300, true, true);
    await db.exec(migration);
    assert.equal((await summary()).balance, 0);
    await db.query("update public.purchases set collected = true, paid_price = 250 where id = $1", [uuid(11)]);
    await db.query("update public.purchases set paid_price = 320 where id = $1", [uuid(19)]);
    assert.equal((await summary()).balance, 0);
  });

  await t.test("Home receipt adds paid price once, including a later collection confirmation", async () => {
    await add(12, HOME, 100);
    await db.query("update public.purchases set picked_up = true where id = $1", [uuid(12)]);
    assert.equal((await summary()).balance, 100);
    await db.query("update public.purchases set collected = true where id = $1", [uuid(12)]);
    assert.equal((await summary()).balance, 100);
    assert.equal((await db.query("select count(*) as count from public.home_cash_transactions where purchase_id = $1", [uuid(12)])).rows[0].count, 2);
  });

  await t.test("Maryamti, Nablus, and Delivery add cash only on collection", async () => {
    for (const [n, point] of [[13, MARYAMTI], [14, NABLUS], [15, DELIVERY]]) {
      const before = (await summary()).balance;
      await add(n, point, 50);
      await db.query("update public.purchases set picked_up = true where id = $1", [uuid(n)]);
      assert.equal((await summary()).balance, before);
      await db.query("update public.purchases set collected = true where id = $1", [uuid(n)]);
      assert.equal((await summary()).balance, before + 50);
    }
    assert.equal((await summary()).balance, 250);
  });

  await t.test("preserves zero paid price and falls back only for null paid price", async () => {
    await add(16, HOME, 0, true);
    assert.equal((await summary()).balance, 250);
    await add(17, HOME, null, true, false, 30);
    assert.equal((await summary()).balance, 280);
  });

  await t.test("price corrections and receipt reversals adjust cash; transfers do not", async () => {
    await db.query("update public.purchases set paid_price = 110 where id = $1", [uuid(12)]);
    assert.equal((await summary()).balance, 290);
    await db.query("update public.purchases set pickup_point = $2 where id = $1", [uuid(12), DELIVERY]);
    assert.equal((await summary()).balance, 290);
    await db.query("update public.purchases set collected = false where id = $1", [uuid(15)]);
    assert.equal((await summary()).balance, 240);
    await db.query("update public.purchases set collected = true where id = $1", [uuid(15)]);
    assert.equal((await summary()).balance, 290);
    await db.query("update public.purchases set picked_up = false where id = $1", [uuid(17)]);
    assert.equal((await summary()).balance, 260);
    await db.query("update public.purchases set picked_up = true where id = $1", [uuid(17)]);
    assert.equal((await summary()).balance, 290);
  });

  await t.test("expense totals are independent of quantity and retries are idempotent", async () => {
    await spend(100, "bags", 30, 50, "Bags");
    await spend(100, "bags", 30, 50, "Bags");
    assert.equal((await summary()).balance, 260);
    await assert.rejects(spend(100, "bags", 31, 50, "Bags"));
    for (const [n, category] of [[101, "pins"], [102, "name_stickers"], [103, "postal_delivery"], [104, "atm_deposit"]]) {
      await spend(n, category, 10);
    }
    assert.equal((await summary()).balance, 220);
    assert.equal((await summary()).expenses, 70);
  });

  await t.test("rejects overspending and invalid amount, quantity, category, and notes", async () => {
    for (const [category, amount, quantity, note] of [
      ["bags", 1000, 1, null], ["bags", 0, 1, null], ["bags", -5, 1, null],
      ["bags", 1.001, 1, null], ["bags", "NaN", 1, null], ["bags", "Infinity", 1, null],
      ["bags", 5, 0, null], ["other", 5, 1, null], ["bags", 5, 1, "x".repeat(501)]
    ]) await assert.rejects(spend(110, category, amount, quantity, note));
    assert.equal((await summary()).balance, 220);
  });

  await t.test("cancellation restores cash once and retains an audit reversal", async () => {
    await db.query("select public.cancel_home_cash_expense($1)", [uuid(100)]);
    await db.query("select public.cancel_home_cash_expense($1)", [uuid(100)]);
    assert.equal((await summary()).balance, 250);
    assert.equal((await summary()).expenses, 40);
    const entries = (await db.query("select * from public.get_home_cash_activity(0,100)")).rows;
    assert(entries.find((entry) => entry.id === uuid(100)).is_voided);
    assert.equal(entries.filter((entry) => entry.kind === "expense_reversal").length, 1);
    assert.equal(Number(entries[0].balance_after), 250);
    assert(entries.some((entry) => entry.actor_email === "rahaf@she-store.com"));
  });

  await t.test("preserves cash on deletion and migration reruns, with full-history pagination", async () => {
    await db.query("delete from public.purchases where id = $1", [uuid(13)]);
    assert.equal((await summary()).balance, 250);
    await db.exec(migration);
    assert.equal((await summary()).balance, 250);
    for (let n = 200; n < 230; n++) await spend(n, "postal_delivery", 1);
    const recent = (await db.query("select * from public.get_home_cash_activity(0,25)")).rows;
    const older = (await db.query("select * from public.get_home_cash_activity(25,25)")).rows;
    assert.equal(recent.length, 25);
    assert.equal(Number(recent[0].balance_after), 220);
    assert.equal(Number(older[0].balance_after), 245);
    assert(!recent.some((entry) => older.some((old) => old.id === entry.id)));
    assert.equal((await db.query("select count(*) as count from public.home_cash_transactions where kind = 'tracking_started'")).rows[0].count, 1);
  });

  await t.test("admins can read and use expense RPCs but cannot rewrite audit entries directly", async () => {
    await db.exec("set role authenticated;");
    assert.equal((await summary()).balance, 220);
    await spend(400, "pins", 1);
    assert.equal((await summary()).balance, 219);
    await assert.rejects(db.query("update public.home_cash_transactions set amount = 1000 where id = $1", [uuid(400)]));
    await assert.rejects(db.query("delete from public.home_cash_transactions where id = $1", [uuid(400)]));
    await db.query("select public.cancel_home_cash_expense($1)", [uuid(400)]);
    assert.equal((await summary()).balance, 220);
    await db.exec("reset role;");
  });

  await t.test("staff cannot read or write the ledger, but their receipts still add cash", async () => {
    await db.exec(`set role authenticated; select set_config('request.jwt.claim.sub', '${uuid(2)}', false);`);
    assert.equal(await summary(), null);
    assert.equal((await db.query("select * from public.get_home_cash_activity()")).rows.length, 0);
    await assert.rejects(spend(300, "bags", 5));
    await assert.rejects(db.query("select public.cancel_home_cash_expense($1)", [uuid(101)]));
    await assert.rejects(db.query("insert into public.home_cash_transactions (kind, amount) values ('receipt',1000)"));
    await assert.rejects(db.query("select * from public.home_cash_receipts"));
    await add(20, DELIVERY, 15, true);
    await db.query("update public.purchases set collected = true where id = $1", [uuid(20)]);
    await db.exec(`reset role; select set_config('request.jwt.claim.sub', '${uuid(1)}', false);`);
    assert.equal((await summary()).balance, 235);
    const latest = (await db.query("select * from public.get_home_cash_activity(0,1)")).rows[0];
    assert.equal(latest.actor_email, "pickup@she-store.com");
    assert.equal(latest.source, "delivery");
    await db.exec("set role anon;");
    await assert.rejects(db.query("select public.get_home_cash_summary()"));
    await db.exec("reset role;");
  });
});
