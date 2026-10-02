import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { PGlite } from "@electric-sql/pglite";

const uuid = (n) => `00000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const HOME = "\u0645\u0646 \u0627\u0644\u0628\u064a\u062a";
const MARYAMTI = "\u0645\u0646 \u0646\u0642\u0637\u0629 \u0627\u0644\u0627\u0633\u062a\u0644\u0627\u0645";
const NABLUS = "\u0646\u0642\u0637\u0629 \u0627\u0633\u062a\u0644\u0627\u0645 \u0646\u0627\u0628\u0644\u0633";
const DELIVERY = "\u062a\u0648\u0635\u064a\u0644";

test("Standalone instant pickups and Home Cash", async (t) => {
  const db = new PGlite();
  t.after(() => db.close());
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
    create table public.purchases (id uuid primary key, order_id uuid not null references public.orders on delete cascade,
      customer_name text, pickup_point text, price numeric, paid_price numeric,
      picked_up boolean default false, collected boolean default false);
    insert into auth.users values ('${uuid(1)}', 'rahaf@she-store.com'), ('${uuid(2)}', 'maryamti@she-store.com'),
      ('${uuid(3)}', 'NABLUS@she-store.com'), ('${uuid(4)}', 'reem@she-store.com'),
      ('${uuid(5)}', 'rawand@she-store.com'), ('${uuid(6)}', 'laaura@she-store.com');
    insert into public.user_roles values ('${uuid(1)}', 'admin'), ('${uuid(2)}', 'pickup'),
      ('${uuid(3)}', 'pickup'), ('${uuid(4)}', 'viewer'), ('${uuid(5)}', 'viewer'), ('${uuid(6)}', 'pickup');
    grant usage on schema auth, public to authenticated;
    grant select on public.user_roles to authenticated;
    select set_config('request.jwt.claim.sub', '${uuid(1)}', false);
    insert into public.orders values ('${uuid(10)}', 'Existing order');
    insert into public.purchases (id,order_id,customer_name,pickup_point,price) values ('${uuid(11)}','${uuid(10)}','Existing customer','${HOME}',100);
  `);
  const cashMigration = await readFile(new URL("../supabase/migrations/20261002010000_add_home_cash_ledger.sql", import.meta.url), "utf8");
  const instantMigration = await readFile(new URL("../supabase/migrations/20261002020000_add_instant_pickups.sql", import.meta.url), "utf8");
  await db.exec(cashMigration);
  await db.exec(instantMigration);
  const deleteMigration = await readFile(new URL("../supabase/migrations/20261002030000_add_instant_pickup_delete.sql", import.meta.url), "utf8");
  await db.exec(deleteMigration);
  const summary = async () => (await db.query("select public.get_home_cash_summary() as value")).rows[0].value;
  const create = (n, location, price = 10, name = `Instant ${n}`) => db.query(
    "select * from public.create_instant_pickup($1,$2,$3,$4)", [uuid(n), name, price, location]);
  const act = (n, action, point = null) => db.query("select * from public.update_instant_pickup($1,$2,$3)", [uuid(n), action, point]);
  const collect = (ids) => db.query("select public.collect_instant_pickups($1) as count", [ids.map(uuid)]);
  const identity = (n) => db.exec(`reset role; select set_config('request.jwt.claim.sub', '${uuid(n)}', false); set role authenticated;`);

  await t.test("creates ready standalone records without changing orders or purchases", async () => {
    for (const [n, point] of [[20, HOME], [21, MARYAMTI], [22, NABLUS], [23, DELIVERY]]) {
      const saved = (await create(n, point)).rows[0];
      assert.equal(saved.picked_up, false);
      assert.equal(saved.collected, false);
      assert.equal(saved.pickup_point, point);
      assert.equal(saved.created_by, uuid(1));
      assert(!Object.hasOwn(saved, "order_id"));
    }
    assert.equal((await db.query("select count(*) as count from public.orders")).rows[0].count, 1);
    assert.equal((await db.query("select count(*) as count from public.purchases")).rows[0].count, 1);
    assert.equal((await summary()).balance, 0);
  });

  await t.test("creation retries are idempotent and input validation rejects invalid entries", async () => {
    await create(20, HOME);
    assert.equal((await db.query("select count(*) as count from public.instant_pickups")).rows[0].count, 4);
    await assert.rejects(create(20, HOME, 11));
    for (const [point, price, name] of [[HOME, -1, "Customer"], [HOME, 1.001, "Customer"], [HOME, "NaN", "Customer"],
      [HOME, 1, " "], [HOME, 1, "x".repeat(201)], ["unknown", 1, "Customer"]]) {
      await assert.rejects(create(99, point, price, name));
    }
    await create(24, MARYAMTI, 0);
    assert.equal((await summary()).balance, 0);
  });

  await t.test("Home receipt credits cash immediately but later collection does not double count", async () => {
    await act(20, "receive");
    assert.equal((await summary()).balance, 10);
    await act(20, "receive");
    assert.equal((await summary()).balance, 10);
    await act(20, "collect");
    await act(20, "collect");
    assert.equal((await summary()).balance, 10);
    await assert.rejects(act(20, "undo_receive"));
    await assert.rejects(act(20, "transfer", NABLUS));
  });

  await t.test("other locations add cash on collection only and collection requires receipt", async () => {
    await assert.rejects(act(21, "collect"));
    for (const n of [21, 22, 23, 24]) await act(n, "receive");
    assert.equal((await summary()).balance, 10);
    assert.equal((await collect([21, 22, 23, 24])).rows[0].count, 4);
    assert.equal((await summary()).balance, 40);
    assert.equal((await collect([21, 22, 23, 24])).rows[0].count, 0);
    assert.equal((await summary()).balance, 40);
  });

  await t.test("bulk collection is atomic when any selected purchase is not received or missing", async () => {
    await create(25, DELIVERY, 15);
    await create(26, DELIVERY, 20);
    await act(25, "receive");
    await assert.rejects(collect([25, 26]));
    await assert.rejects(collect([25, 999]));
    await assert.rejects(collect([]));
    assert.equal((await db.query("select collected from public.instant_pickups where id = $1", [uuid(25)])).rows[0].collected, false);
    assert.equal((await summary()).balance, 40);
  });

  await t.test("transfers retain actual cash, receipt reversal adjusts it, and zero prices stay zero", async () => {
    await create(27, HOME, 12);
    await act(27, "receive");
    assert.equal((await summary()).balance, 52);
    await act(27, "undo_receive");
    assert.equal((await summary()).balance, 40);
    await act(27, "receive");
    await act(27, "transfer", MARYAMTI);
    assert.equal((await summary()).balance, 52);
    await act(27, "collect");
    assert.equal((await summary()).balance, 52);
    await act(26, "transfer", NABLUS);
    assert.equal((await summary()).balance, 52);
  });

  await t.test("Maryamti and Nablus staff see and receive only their own location", async () => {
    await identity(2);
    const maryamtiRows = (await db.query("select * from public.instant_pickups")).rows;
    assert(maryamtiRows.every((row) => row.pickup_point === MARYAMTI));
    await assert.rejects(act(26, "receive"));
    await assert.rejects(create(50, MARYAMTI));
    await assert.rejects(act(27, "collect"));
    await assert.rejects(act(27, "transfer", HOME));
    await assert.rejects(db.query("update public.instant_pickups set price = 999 where id = $1", [uuid(27)]));
    await identity(3);
    const nablusRows = (await db.query("select * from public.instant_pickups")).rows;
    assert(nablusRows.length > 0 && nablusRows.every((row) => row.pickup_point === NABLUS));
    await act(26, "receive");
    await act(26, "undo_receive");
    await act(26, "receive");
    await assert.rejects(act(25, "receive"));
    await assert.rejects(collect([26]));
    await identity(1);
    assert.equal((await summary()).balance, 52);
    await collect([26]);
    assert.equal((await summary()).balance, 72);
  });

  await t.test("viewers retain their existing locations and cannot mutate, blocked staff and anonymous cannot access", async () => {
    await identity(4);
    const reemRows = (await db.query("select * from public.instant_pickups")).rows;
    assert(reemRows.length > 0 && reemRows.every((row) => [HOME, NABLUS, DELIVERY].includes(row.pickup_point)));
    await assert.rejects(act(25, "receive"));
    await assert.rejects(create(51, HOME));
    await identity(5);
    const rawandRows = (await db.query("select * from public.instant_pickups")).rows;
    assert(rawandRows.length > 0 && rawandRows.every((row) => [HOME, DELIVERY].includes(row.pickup_point)));
    await identity(6);
    assert.equal((await db.query("select * from public.instant_pickups")).rows.length, 0);
    await assert.rejects(act(25, "receive"));
    await db.exec("reset role; set role anon;");
    await assert.rejects(db.query("select * from public.instant_pickups"));
    await assert.rejects(create(52, HOME));
    await identity(1);
  });

  await t.test("existing Home Cash expenses and reversals work with standalone income", async () => {
    const expense = (await db.query("select * from public.record_home_cash_expense($1,$2,$3,$4,$5)",
      [uuid(70), "bags", 10, 20, "Standalone income expense"])).rows[0];
    assert.equal(expense.instant_pickup_id, null);
    assert.equal((await summary()).balance, 62);
    await db.query("select public.cancel_home_cash_expense($1)", [uuid(70)]);
    assert.equal((await summary()).balance, 72);
  });

  await t.test("a pickup-role account cannot gain location access by including its name in an unrelated email", async () => {
    await db.exec("reset role;");
    await db.query("update auth.users set email = $2 where id = $1", [uuid(6), "nablus@unrelated.example"]);
    await identity(6);
    assert.equal((await db.query("select * from public.instant_pickups")).rows.length, 0);
    await assert.rejects(act(26, "receive"));
    await identity(1);
  });

  await t.test("only admins delete pending entries; retries are safe and received cash is protected", async () => {
    const remove = (n) => db.query("select public.delete_instant_pickup($1) as deleted", [uuid(n)]);
    await identity(1);
    for (const [n, point] of [[80, HOME], [81, MARYAMTI], [82, NABLUS], [83, DELIVERY]]) {
      await create(n, point);
    }
    for (const n of [2, 3, 4, 5, 6]) {
      await identity(n);
      await assert.rejects(remove(80), /صلاحية/);
      await assert.rejects(remove(999), /صلاحية/);
    }
    await db.exec("reset role; set role anon;");
    await assert.rejects(remove(80));
    await identity(1);
    await assert.rejects(db.query("delete from public.instant_pickups where id = $1", [uuid(80)]));
    await assert.rejects(db.query("select public.delete_instant_pickup(null)"));
    await act(80, "receive");
    assert.equal((await summary()).balance, 82);
    await assert.rejects(remove(80), /استلامه/);
    await act(81, "receive");
    await assert.rejects(remove(81), /استلامه/);
    await assert.rejects(remove(21), /استلامه/);
    await act(80, "undo_receive");
    await act(81, "undo_receive");
    assert.equal((await summary()).balance, 72);
    for (const n of [80, 81, 82, 83]) {
      assert.equal((await remove(n)).rows[0].deleted, true);
      assert.equal((await remove(n)).rows[0].deleted, false);
      assert.equal((await db.query("select * from public.instant_pickups where id = $1", [uuid(n)])).rows.length, 0);
    }
    const history = (await db.query("select * from public.home_cash_transactions where customer_name = 'Instant 80'")).rows;
    assert.equal(history.length, 2);
    assert(history.every((row) => row.instant_pickup_id === null));
    assert.equal((await summary()).balance, 72);
    await db.exec("reset role;");
    await db.exec(deleteMigration);
    await identity(1);
    assert.equal((await remove(80)).rows[0].deleted, false);
  });

  await t.test("cash audit entries identify instant pickups and survive migration reruns and deletion", async () => {
    const history = (await db.query("select * from public.get_home_cash_activity(0,100)")).rows;
    assert(history.some((row) => row.customer_name === "Instant 26" && row.order_name === "\u0627\u0633\u062a\u0644\u0627\u0645 \u0641\u0648\u0631\u064a"));
    assert.equal(Number(history[0].balance_after), 72);
    await db.exec("reset role;");
    await db.exec(instantMigration);
    assert.equal((await summary()).balance, 72);
    await db.query("delete from public.instant_pickups where id = $1", [uuid(20)]);
    assert.equal((await summary()).balance, 72);
    assert((await db.query("select * from public.home_cash_transactions where customer_name = 'Instant 20'")).rows.every((row) => row.instant_pickup_id === null));
    await db.exec(cashMigration);
    assert.equal((await summary()).balance, 72);
    assert.equal((await db.query("select count(*) as count from public.orders")).rows[0].count, 1);
    assert.equal((await db.query("select count(*) as count from public.purchases")).rows[0].count, 1);
  });
});
