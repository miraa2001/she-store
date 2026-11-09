begin;

-- No data backfill: existing purchases and finance records are untouched.
alter table public.purchases add column if not exists abandoned_at timestamptz;
create table if not exists public.abandoned_purchase_archives (
  id uuid primary key default gen_random_uuid(),
  purchase_id uuid not null references public.purchases(id) on delete restrict,
  status text not null default 'available' check(status in ('available','restored','resold')),
  purchase_snapshot jsonb not null,
  order_snapshot jsonb not null,
  images jsonb not null default '[]',
  links jsonb not null default '[]',
  archived_at timestamptz not null default clock_timestamp(),
  archived_by uuid,
  actor_email text,
  restored_at timestamptz,
  resold_at timestamptz,
  instant_pickup_id uuid unique references public.instant_pickups(id) on delete restrict
);
create unique index if not exists abandoned_purchase_active_idx
  on public.abandoned_purchase_archives(purchase_id) where status<>'restored';
create index if not exists abandoned_purchase_status_idx on public.abandoned_purchase_archives(status,archived_at desc,id);
create table if not exists public.abandoned_purchase_operations (
  request_id uuid primary key,
  kind text not null check(kind in ('archive','restore','resell','cancel_resale')),
  payload jsonb not null,
  result jsonb not null,
  created_at timestamptz not null default clock_timestamp(),
  created_by uuid,
  actor_email text
);
alter table public.abandoned_purchase_archives enable row level security;
alter table public.abandoned_purchase_operations enable row level security;
revoke all on public.abandoned_purchase_archives,public.abandoned_purchase_operations from anon,authenticated;
grant select on public.abandoned_purchase_archives,public.abandoned_purchase_operations to authenticated;
drop policy if exists abandoned_archives_admin_read on public.abandoned_purchase_archives;
create policy abandoned_archives_admin_read on public.abandoned_purchase_archives for select to authenticated using(public.is_admin());
drop policy if exists abandoned_operations_admin_read on public.abandoned_purchase_operations;
create policy abandoned_operations_admin_read on public.abandoned_purchase_operations for select to authenticated using(public.is_admin());

create or replace function public.guard_abandoned_purchase()
returns trigger language plpgsql set search_path='' as $$
begin
  if tg_op='DELETE' then
    if old.abandoned_at is not null then raise exception 'أعيدي المشترى من الأرشيف قبل الحذف'; end if;
    return old;
  end if;
  -- Only the archive RPCs can set/clear the flag, even for the admin client.
  if current_user in ('authenticated','anon') and (
    (tg_op='INSERT' and new.abandoned_at is not null)
    or (tg_op='UPDATE' and new.abandoned_at is distinct from old.abandoned_at)
  ) then raise exception 'استخدمي أرشيف المشتريات المتروكة'; end if;
  if new.abandoned_at is not null then
    if not exists(select 1 from public.abandoned_purchase_archives a where a.purchase_id=new.id and a.status in ('available','resold')) then
      raise exception 'استخدمي أرشيف المشتريات المتروكة';
    end if;
    if new.picked_up or new.collected or new.received_amount<>0 or new.remitted_amount<>0
      or exists(select 1 from public.purchase_payments where purchase_id=new.id) then
      raise exception 'لا يمكن أرشفة مشترى مستلم أو له سجل دفعات';
    end if;
    if tg_op='UPDATE' and (to_jsonb(new)-array['abandoned_at','ready_for_pickup','ready_for_pickup_at','payment_model_ready'])
      is distinct from (to_jsonb(old)-array['abandoned_at','ready_for_pickup','ready_for_pickup_at','payment_model_ready']) then
      raise exception 'أعيدي المشترى من الأرشيف قبل تعديله';
    end if;
    new.ready_for_pickup:=false;
    new.ready_for_pickup_at:=null;
  elsif tg_op='UPDATE' and old.abandoned_at is not null and exists(
    select 1 from public.abandoned_purchase_archives where purchase_id=new.id and status<>'restored'
  ) then
    raise exception 'استخدمي التراجع في أرشيف المشتريات المتروكة';
  end if;
  return new;
end;
$$;
revoke all on function public.guard_abandoned_purchase() from public,anon,authenticated;
drop trigger if exists purchases_02_abandoned_guard on public.purchases;
create trigger purchases_02_abandoned_guard before insert or update or delete on public.purchases
  for each row execute function public.guard_abandoned_purchase();

create or replace function public.guard_abandoned_metadata()
returns trigger language plpgsql security definer set search_path='' as $$
begin
  if (tg_op<>'INSERT' and exists(select 1 from public.purchases where id=old.purchase_id and abandoned_at is not null))
    or (tg_op<>'DELETE' and exists(select 1 from public.purchases where id=new.purchase_id and abandoned_at is not null)) then
    raise exception 'صور وروابط المشترى محفوظة في الأرشيف؛ أعيديه قبل التعديل';
  end if;
  if tg_op='DELETE' then return old; end if;
  return new;
end;
$$;
revoke all on function public.guard_abandoned_metadata() from public,anon,authenticated;
drop trigger if exists images_abandoned_guard on public.purchase_images;
create trigger images_abandoned_guard before insert or update or delete on public.purchase_images
  for each row execute function public.guard_abandoned_metadata();
drop trigger if exists links_abandoned_guard on public.purchase_links;
create trigger links_abandoned_guard before insert or update or delete on public.purchase_links
  for each row execute function public.guard_abandoned_metadata();

create or replace function public.archive_abandoned_purchases(p_request_id uuid,p_purchase_ids uuid[])
returns jsonb language plpgsql security definer set search_path='' as $$
declare ids uuid[]; payload jsonb; prior public.abandoned_purchase_operations%rowtype;
  source_purchase public.purchases%rowtype; archive_id uuid; archived uuid[]:='{}'; actor text; result jsonb;
begin
  if not coalesce(public.is_admin(),false) then raise exception 'متاح لرهف فقط' using errcode='42501'; end if;
  if p_request_id is null or cardinality(p_purchase_ids) is null or cardinality(p_purchase_ids)=0
    or array_position(p_purchase_ids,null) is not null then raise exception 'اختاري المشتريات'; end if;
  select array_agg(distinct item order by item) into ids from unnest(p_purchase_ids) item;
  payload:=to_jsonb(ids);
  perform 1 from public.home_cash_account where id=1 for update;
  if not found then raise exception 'تتبع مصاري البيت غير متاح'; end if;
  select * into prior from public.abandoned_purchase_operations where request_id=p_request_id;
  if found then
    if prior.kind='archive' and prior.payload=payload then return prior.result; end if;
    raise exception 'معرف العملية مستخدم بقيم مختلفة';
  end if;
  perform 1 from public.purchases where id=any(ids) order by id for update;
  if (select count(*) from public.purchases where id=any(ids))<>cardinality(ids) then raise exception 'تغيرت المشتريات؛ أعيدي التحميل'; end if;
  if exists(select 1 from public.purchases p where p.id=any(ids) and (
    p.order_id is null or p.abandoned_at is not null or p.picked_up or p.collected or p.received_amount<>0 or p.remitted_amount<>0
    or exists(select 1 from public.purchase_payments pp where pp.purchase_id=p.id)
  )) then raise exception 'الأرشفة للمشتريات المرتبطة بطلب وغير المستلمة وبدون سجل دفعات فقط'; end if;
  select email into actor from auth.users where id=auth.uid();
  for source_purchase in select * from public.purchases where id=any(ids) order by id loop
    insert into public.abandoned_purchase_archives(purchase_id,purchase_snapshot,order_snapshot,images,links,archived_by,actor_email)
    select source_purchase.id,to_jsonb(source_purchase),to_jsonb(o),
      (select coalesce(jsonb_agg(to_jsonb(i) order by i.id),'[]') from public.purchase_images i where i.purchase_id=source_purchase.id),
      (select coalesce(jsonb_agg(to_jsonb(l) order by l.id),'[]') from public.purchase_links l where l.purchase_id=source_purchase.id),auth.uid(),actor
    from public.orders o where o.id=source_purchase.order_id returning id into archive_id;
    update public.purchases set abandoned_at=clock_timestamp(),ready_for_pickup=false,ready_for_pickup_at=null where id=source_purchase.id;
    archived:=array_append(archived,archive_id);
  end loop;
  result:=jsonb_build_object('archive_ids',archived,'purchase_count',cardinality(ids));
  insert into public.abandoned_purchase_operations(request_id,kind,payload,result,created_by,actor_email)
    values(p_request_id,'archive',payload,result,auth.uid(),actor);
  return result;
end;
$$;

create or replace function public.restore_abandoned_purchases(p_request_id uuid,p_archive_ids uuid[])
returns jsonb language plpgsql security definer set search_path='' as $$
declare ids uuid[]; payload jsonb; prior public.abandoned_purchase_operations%rowtype;
  a public.abandoned_purchase_archives%rowtype; result jsonb; actor text; eligible boolean;
begin
  if not coalesce(public.is_admin(),false) then raise exception 'متاح لرهف فقط' using errcode='42501'; end if;
  if p_request_id is null or cardinality(p_archive_ids) is null or cardinality(p_archive_ids)=0
    or array_position(p_archive_ids,null) is not null then raise exception 'اختاري المشتريات'; end if;
  select array_agg(distinct item order by item) into ids from unnest(p_archive_ids) item;
  payload:=to_jsonb(ids);
  perform 1 from public.home_cash_account where id=1 for update;
  select * into prior from public.abandoned_purchase_operations where request_id=p_request_id;
  if found then
    if prior.kind='restore' and prior.payload=payload then return prior.result; end if;
    raise exception 'معرف العملية مستخدم بقيم مختلفة';
  end if;
  perform 1 from public.abandoned_purchase_archives where id=any(ids) order by id for update;
  if (select count(*) from public.abandoned_purchase_archives where id=any(ids) and status='available')<>cardinality(ids) then
    raise exception 'التراجع متاح للمشتريات المتاحة التي لم يتم بيعها فقط'; end if;
  for a in select * from public.abandoned_purchase_archives where id=any(ids) order by id loop
    select coalesce(o.arrived,false) and (public.home_cash_receipt_source(p.pickup_point) in ('home','delivery') or coalesce(o.placed_at_pickup,false))
      into eligible from public.purchases p join public.orders o on o.id=p.order_id where p.id=a.purchase_id for update of p;
    update public.abandoned_purchase_archives set status='restored',restored_at=clock_timestamp() where id=a.id;
    update public.purchases set abandoned_at=null,ready_for_pickup=coalesce(eligible,false),
      ready_for_pickup_at=case when eligible then coalesce((a.purchase_snapshot->>'ready_for_pickup_at')::timestamptz,clock_timestamp()) end
      where id=a.purchase_id;
  end loop;
  result:=jsonb_build_object('purchase_count',cardinality(ids));
  select email into actor from auth.users where id=auth.uid();
  insert into public.abandoned_purchase_operations(request_id,kind,payload,result,created_by,actor_email)
    values(p_request_id,'restore',payload,result,auth.uid(),actor);
  return result;
end;
$$;

create or replace function public.resell_abandoned_purchase(p_request_id uuid,p_archive_id uuid,p_customer_name text,p_price numeric,p_pickup_point text)
returns jsonb language plpgsql security definer set search_path='' as $$
declare a public.abandoned_purchase_archives%rowtype; prior public.abandoned_purchase_operations%rowtype;
  payload jsonb; result jsonb; sale public.instant_pickups%rowtype; actor text;
begin
  if not coalesce(public.is_admin(),false) then raise exception 'متاح لرهف فقط' using errcode='42501'; end if;
  if p_request_id is null or p_archive_id is null then raise exception 'المشترى غير موجود'; end if;
  payload:=jsonb_build_object('archive_id',p_archive_id,'customer_name',btrim(p_customer_name),'price',p_price,'pickup_point',p_pickup_point);
  perform 1 from public.home_cash_account where id=1 for update;
  select * into prior from public.abandoned_purchase_operations where request_id=p_request_id;
  if found then
    if prior.kind='resell' and prior.payload=payload then return prior.result; end if;
    raise exception 'معرف العملية مستخدم بقيم مختلفة';
  end if;
  select * into a from public.abandoned_purchase_archives where id=p_archive_id for update;
  if not found or a.status<>'available' then raise exception 'المشترى غير متاح لإعادة البيع'; end if;
  if exists(select 1 from public.instant_pickups where id=p_request_id) then raise exception 'معرف العملية مستخدم مسبقاً'; end if;
  select * into sale from public.create_instant_pickup(p_request_id,p_customer_name,p_price,p_pickup_point);
  update public.abandoned_purchase_archives set status='resold',resold_at=clock_timestamp(),instant_pickup_id=sale.id where id=a.id;
  result:=to_jsonb(sale);
  select email into actor from auth.users where id=auth.uid();
  insert into public.abandoned_purchase_operations(request_id,kind,payload,result,created_by,actor_email)
    values(p_request_id,'resell',payload,result,auth.uid(),actor);
  return result;
end;
$$;
revoke all on function public.archive_abandoned_purchases(uuid,uuid[]),public.restore_abandoned_purchases(uuid,uuid[]),
  public.resell_abandoned_purchase(uuid,uuid,text,numeric,text) from public,anon,authenticated;
grant execute on function public.archive_abandoned_purchases(uuid,uuid[]),public.restore_abandoned_purchases(uuid,uuid[]),
  public.resell_abandoned_purchase(uuid,uuid,text,numeric,text) to authenticated;

create or replace function public.release_abandoned_resale()
returns trigger language plpgsql security definer set search_path='' as $$
declare a public.abandoned_purchase_archives%rowtype; actor text;
begin
  select * into a from public.abandoned_purchase_archives where instant_pickup_id=old.id for update;
  if not found then return old; end if;
  if not coalesce(public.is_admin(),false) or old.picked_up or old.collected
    or exists(select 1 from public.purchase_payments where instant_pickup_id=old.id) then
    raise exception 'لا يمكن إلغاء إعادة بيع مستلمة أو لها سجل دفعات';
  end if;
  update public.abandoned_purchase_archives set status='available',resold_at=null,instant_pickup_id=null where id=a.id;
  select email into actor from auth.users where id=auth.uid();
  insert into public.abandoned_purchase_operations(request_id,kind,payload,result,created_by,actor_email)
    values(gen_random_uuid(),'cancel_resale',jsonb_build_object('archive_id',a.id,'cancelled_sale',to_jsonb(old)),
      jsonb_build_object('status','available'),auth.uid(),actor);
  return old;
end;
$$;
revoke all on function public.release_abandoned_resale() from public,anon,authenticated;
drop trigger if exists instant_02_release_abandoned on public.instant_pickups;
create trigger instant_02_release_abandoned before delete on public.instant_pickups
  for each row execute function public.release_abandoned_resale();

notify pgrst,'reload schema';
commit;
