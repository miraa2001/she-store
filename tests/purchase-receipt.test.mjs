import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { PGlite } from "@electric-sql/pglite";
import { savePurchaseReceipt } from "../src/lib/purchaseReceipt.js";

const id = (n) => `00000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const migrationUrl = new URL("../supabase/migrations/20261021000000_add_purchase_receipt_rpc.sql", import.meta.url);

test("receipt RPC permissions, persistence and cash integration", async (t) => {
  const db = new PGlite();
  t.after(() => db.close());
  await db.exec(`
    create role anon; create role authenticated;
    create schema auth;
    create table auth.users(id uuid primary key, email text);
    create function auth.uid() returns uuid language sql stable as $$
      select nullif(current_setting('request.jwt.claim.sub', true), '')::uuid;
    $$;
    create table public.user_roles(user_id uuid primary key, role text);
    create function public.is_admin() returns boolean language sql stable as $$
      select exists(select 1 from public.user_roles where user_id=auth.uid() and role='admin');
    $$;
    create table public.orders(id uuid primary key, order_name text);
    create table public.purchases(id uuid primary key, order_id uuid references public.orders,
      customer_name text, pickup_point text, price numeric, paid_price numeric,
      ready_for_pickup boolean default true, picked_up boolean default false,
      picked_up_at timestamptz, collected boolean default false);
    insert into public.orders values ('${id(10)}','Receipt order');
    insert into auth.users values ('${id(1)}','rahaf@she-store.com'), ('${id(2)}','reem@she-store.com'),
      ('${id(3)}','rawand@she-store.com'), ('${id(4)}','maryamti@she-store.com'),
      ('${id(5)}','nablus@she-store.com'), ('${id(6)}','laaura@she-store.com'),
      ('${id(7)}','not-reem@she-store.com');
    insert into public.user_roles values ('${id(1)}','admin'), ('${id(2)}','viewer'),
      ('${id(3)}','viewer'), ('${id(4)}','pickup'), ('${id(5)}','pickup'),
      ('${id(6)}','pickup'), ('${id(7)}','viewer');
    grant usage on schema public,auth to anon,authenticated;
    grant select on public.user_roles to authenticated;
    grant select,update on public.purchases to authenticated;
    alter table public.purchases enable row level security;
    create policy purchases_select_auth on public.purchases for select to authenticated using(true);
    create policy purchases_write_admin on public.purchases for all to authenticated using(public.is_admin()) with check(public.is_admin());
    create policy pickup_update on public.purchases for update to authenticated using (
      exists(select 1 from public.user_roles where user_id=auth.uid() and role in ('admin','pickup')));
    insert into public.purchases(id,order_id,customer_name,pickup_point,price,paid_price) values
      ('${id(100)}','${id(10)}','Tala','من البيت',50,42),
      ('${id(101)}','${id(10)}','Nablus','نقطة استلام نابلس',20,20),
      ('${id(102)}','${id(10)}','Delivery','توصيل',30,30),
      ('${id(103)}','${id(10)}','Maryamti','من نقطة الاستلام - مريمتي',40,40),
      ('${id(104)}','${id(10)}','Not ready','من البيت',10,10),
      ('${id(105)}','${id(10)}','Collected','من البيت',10,10);
    update public.purchases set ready_for_pickup=false where id='${id(104)}';
    update public.purchases set picked_up=true,collected=true where id='${id(105)}';
  `);
  await db.exec(await readFile(new URL("../supabase/migrations/20261002010000_add_home_cash_ledger.sql", import.meta.url), "utf8"));
  const migration = await readFile(migrationUrl, "utf8");
  await db.exec(migration);
  const identity = (n) => db.exec(`reset role; select set_config('request.jwt.claim.sub','${id(n)}',false); set role authenticated;`);
  const receive = async (n, checked = true) => (await db.query("select public.set_purchase_received($1,$2) as result", [id(n), checked])).rows[0].result;
  const balance = async () => {
    await identity(1);
    return Number((await db.query("select public.get_home_cash_summary() as summary")).rows[0].summary.balance);
  };

  await t.test("reproduces viewer silent zero-row update, while receipt RPC persists", async () => {
    await identity(2);
    assert.equal((await db.query("update public.purchases set picked_up=true where id=$1 returning id", [id(100)])).rows.length, 0);
    const result = await receive(100);
    assert.equal(result.picked_up, true);
    assert(result.picked_up_at);
    assert.equal(result.changed, true);
    assert.equal((await db.query("select picked_up from public.purchases where id=$1", [id(100)])).rows[0].picked_up, true);
    assert.equal(await balance(), 42);
  });
  await t.test("retries are idempotent and undo reverses Home cash", async () => {
    await identity(2);
    assert.equal((await receive(100)).changed, false);
    assert.equal(await balance(), 42);
    await identity(2);
    assert.equal((await receive(100, false)).picked_up_at, null);
    assert.equal(await balance(), 0);
  });
  await t.test("Reem can receive Nablus and Delivery, but not Maryamti", async () => {
    await identity(2);
    await receive(101); await receive(102);
    await assert.rejects(receive(103), /صلاحية/);
    assert.equal(await balance(), 0);
  });
  await t.test("viewers cannot edit prices, collection, location, or names directly", async () => {
    await identity(2);
    const result = await db.query("update public.purchases set paid_price=1,collected=true,pickup_point='توصيل',customer_name='Changed' where id=$1 returning id", [id(100)]);
    assert.equal(result.rows.length, 0);
    const row = (await db.query("select * from public.purchases where id=$1", [id(100)])).rows[0];
    assert.equal(Number(row.paid_price), 42); assert.equal(row.collected, false);
    assert.equal(row.customer_name, "Tala"); assert.equal(row.pickup_point, "من البيت");
  });
  await t.test("other viewers and blocked staff gain no receipt access", async () => {
    for (const actor of [3,6,7]) {
      await identity(actor);
      await assert.rejects(receive(100), /صلاحية/);
    }
  });
  await t.test("pickup staff use only their own location; admin can use both", async () => {
    await identity(4); await receive(103);
    await assert.rejects(receive(101, false), /صلاحية/);
    await identity(5); await receive(101, false);
    await assert.rejects(receive(103, false), /صلاحية/);
    await identity(1); await receive(103, false); await receive(101);
  });
  await t.test("collected, not-ready, missing, null and anonymous requests fail", async () => {
    await identity(2);
    await assert.rejects(receive(104), /جاهز/);
    await assert.rejects(receive(105, false), /تحصيله/);
    await assert.rejects(receive(999), /موجود/);
    await assert.rejects(receive(100, null), /صالحة/);
    await db.exec("reset role; set role anon;");
    await assert.rejects(receive(100));
    await db.exec("reset role; select set_config('request.jwt.claim.sub','',false); set role authenticated;");
    await assert.rejects(receive(100), /تسجيل الدخول/);
  });
  await t.test("migration reruns preserve receipt status and cash history", async () => {
    await db.exec("reset role;");
    await db.exec(migration);
    assert.equal((await db.query("select picked_up from public.purchases where id=$1", [id(101)])).rows[0].picked_up, true);
    assert.equal(await balance(), 0);
  });
});

test("client rejects missing, mismatched and unverified RPC results", async () => {
  const saved = { id: id(100), picked_up: true, picked_up_at: "2026-10-21T12:00:00Z", changed: true };
  let args;
  const client = { async rpc(name, params) { args = { name, params }; return { data: saved, error: null }; } };
  assert.equal(await savePurchaseReceipt(client, id(100), true), saved);
  assert.deepEqual(args, { name: "set_purchase_received", params: { p_purchase_id: id(100), p_picked_up: true } });
  for (const data of [null, [], { ...saved, id: id(200) }, { ...saved, picked_up: false }, { ...saved, picked_up_at: null }, { ...saved, changed: undefined }]) {
    await assert.rejects(savePurchaseReceipt({ rpc: async () => ({ data }) }, id(100), true));
  }
  await assert.rejects(savePurchaseReceipt({ rpc: async () => ({ error: { code: "PGRST202" } }) }, id(100), true), /Supabase/);
});

test("pickup pages use the shared payment control for receipt changes",async()=>{
  for(const page of ["HomePickupPage","PickupPointPage"]) {
    const source=await readFile(new URL(`../src/pages/${page}.jsx`,import.meta.url),"utf8");
    assert(source.includes("<PurchasePaymentControl"));
    assert(!source.includes("async function togglePicked("));
  }
});
