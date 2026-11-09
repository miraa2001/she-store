begin;

alter table public.purchases add column if not exists picked_up_by uuid,
  add column if not exists picked_up_by_email text;
alter table public.instant_pickups add column if not exists picked_up_by uuid,
  add column if not exists picked_up_by_email text;
alter table public.home_cash_transactions add column if not exists pickup_actor_email text;

-- Only verified, non-legacy receipt payments can identify an older pickup account.
update public.purchases p set picked_up_by=r.created_by,picked_up_by_email=u.email
from public.purchase_payments r join auth.users u on u.id=r.created_by
where p.id=r.purchase_id and p.order_id is not null and p.picked_up and p.picked_up_by is null
  and r.received_on_pickup and not r.legacy and r.cancelled_at is null
  and not exists(select 1 from public.purchase_payments earlier where earlier.purchase_id=p.id
    and earlier.received_on_pickup and not earlier.legacy and earlier.cancelled_at is null
    and (earlier.created_at,earlier.id)<(r.created_at,r.id));
update public.instant_pickups p set picked_up_by=r.created_by,picked_up_by_email=u.email
from public.purchase_payments r join auth.users u on u.id=r.created_by
where p.id=r.instant_pickup_id and p.picked_up and p.picked_up_by is null
  and r.received_on_pickup and not r.legacy and r.cancelled_at is null
  and not exists(select 1 from public.purchase_payments earlier where earlier.instant_pickup_id=p.id
    and earlier.received_on_pickup and not earlier.legacy and earlier.cancelled_at is null
    and (earlier.created_at,earlier.id)<(r.created_at,r.id));

create or replace function public.capture_pickup_actor()
returns trigger language plpgsql security definer set search_path='' as $$
begin
  if new.picked_up and (tg_op='INSERT' or not old.picked_up) then
    new.picked_up_by := auth.uid();
    select email into new.picked_up_by_email from auth.users where id=auth.uid();
  elsif not new.picked_up then
    new.picked_up_by := null;
    new.picked_up_by_email := null;
  else
    new.picked_up_by := old.picked_up_by;
    new.picked_up_by_email := old.picked_up_by_email;
  end if;
  return new;
end;
$$;
revoke all on function public.capture_pickup_actor() from public,anon,authenticated;
drop trigger if exists purchases_01_pickup_actor on public.purchases;
create trigger purchases_01_pickup_actor before insert or update on public.purchases
  for each row execute function public.capture_pickup_actor();
drop trigger if exists instant_01_pickup_actor on public.instant_pickups;
create trigger instant_01_pickup_actor before insert or update on public.instant_pickups
  for each row execute function public.capture_pickup_actor();

create or replace function public.apply_location_pickup_readiness()
returns trigger language plpgsql security definer set search_path='' as $$
declare o public.orders%rowtype; eligible boolean; protected boolean;
begin
  select * into o from public.orders where id=new.order_id for share;
  eligible := found and coalesce(o.arrived,false) and (
    public.home_cash_receipt_source(new.pickup_point) in ('home','delivery') or coalesce(o.placed_at_pickup,false));
  protected := new.picked_up or new.collected or exists(select 1 from public.purchase_payments where purchase_id=new.id);
  new.ready_for_pickup := coalesce(eligible,false) or coalesce(protected,false);
  new.ready_for_pickup_at := case when new.ready_for_pickup
    then coalesce(case when tg_op='UPDATE' then old.ready_for_pickup_at end,new.ready_for_pickup_at,clock_timestamp())
    else null end;
  return new;
end;
$$;
revoke all on function public.apply_location_pickup_readiness() from public,anon,authenticated;
drop trigger if exists purchases_00_location_readiness on public.purchases;
create trigger purchases_00_location_readiness before insert or update of order_id,pickup_point,ready_for_pickup on public.purchases
  for each row execute function public.apply_location_pickup_readiness();

create or replace function public.sync_location_pickup_readiness()
returns trigger language plpgsql security definer set search_path='' as $$
begin
  update public.purchases set ready_for_pickup=ready_for_pickup where order_id=new.id;
  return new;
end;
$$;
revoke all on function public.sync_location_pickup_readiness() from public,anon,authenticated;
drop trigger if exists orders_location_readiness on public.orders;
create trigger orders_location_readiness after update of arrived,placed_at_pickup on public.orders
  for each row execute function public.sync_location_pickup_readiness();

create or replace function public.capture_cash_pickup_actor()
returns trigger language plpgsql security definer set search_path='' as $$
declare receiver text; received boolean;
begin
  if new.kind not in ('receipt','receipt_adjustment') then new.pickup_actor_email:=null; return new; end if;
  if new.purchase_id is not null then
    select picked_up_by_email,picked_up into receiver,received from public.purchases where id=new.purchase_id;
  elsif new.instant_pickup_id is not null then
    select picked_up_by_email,picked_up into receiver,received from public.instant_pickups where id=new.instant_pickup_id;
  end if;
  -- The first Home payment is posted immediately before the pickup flag is saved.
  if not coalesce(received,false) and new.kind='receipt' and new.source='home' then
    select email into receiver from auth.users where id=auth.uid();
  end if;
  new.pickup_actor_email:=receiver;
  return new;
end;
$$;
revoke all on function public.capture_cash_pickup_actor() from public,anon,authenticated;
drop trigger if exists home_cash_pickup_actor on public.home_cash_transactions;
create trigger home_cash_pickup_actor before insert on public.home_cash_transactions
  for each row execute function public.capture_cash_pickup_actor();

-- Unassigned drafts have no order status and must never enter this backfill.
update public.purchases set ready_for_pickup=ready_for_pickup where order_id is not null;
notify pgrst,'reload schema';
commit;
