begin;

alter table public.purchases alter column order_id drop not null;
create index if not exists purchases_unassigned_created_idx
  on public.purchases(created_at desc, id) where order_id is null;

drop policy if exists purchases_unassigned_admin_only on public.purchases;
create policy purchases_unassigned_admin_only on public.purchases as restrictive
  for select to authenticated using (order_id is not null or public.is_admin());

-- Metadata follows the visibility of its purchase, including unassigned drafts.
drop policy if exists links_purchase_visibility on public.purchase_links;
create policy links_purchase_visibility on public.purchase_links as restrictive
  for select to authenticated using (
    exists(select 1 from public.purchases p where p.id = purchase_id)
  );
drop policy if exists images_purchase_visibility on public.purchase_images;
create policy images_purchase_visibility on public.purchase_images as restrictive
  for select to authenticated using (
    exists(select 1 from public.purchases p where p.id = purchase_id)
  );

create or replace function public.guard_unassigned_purchase()
returns trigger language plpgsql set search_path = '' as $$
begin
  if tg_op = 'UPDATE' and new.order_id is distinct from old.order_id
    and (new.order_id is null or old.order_id is null)
    and current_user in ('authenticated', 'anon') then
    raise exception 'استخدمي نقل المشتريات لتغيير الطلب';
  end if;
  if new.order_id is null then
    if not coalesce(public.is_admin(), false) then
      raise exception 'المشتريات غير المعينة متاحة لرهف فقط' using errcode = '42501';
    end if;
    if new.ready_for_pickup or new.picked_up or new.collected
      or coalesce(new.received_amount,0) <> 0 or coalesce(new.remitted_amount,0) <> 0
      or exists(select 1 from public.purchase_payments where purchase_id = new.id) then
      raise exception 'لا يمكن فصل مشترى جاهز للاستلام أو مستلم أو له سجل دفعات';
    end if;
  end if;
  return new;
end;
$$;
drop trigger if exists purchases_unassigned_guard on public.purchases;
create trigger purchases_unassigned_guard before insert or update on public.purchases
  for each row execute function public.guard_unassigned_purchase();

create or replace function public.transfer_order_purchases(
  p_source_order_id uuid, p_target_order_id uuid, p_purchase_ids uuid[]
)
returns integer language plpgsql security definer set search_path = '' as $$
declare
  requested integer;
  matched integer;
begin
  if not coalesce(public.is_admin(), false) then
    raise exception 'لا توجد صلاحية' using errcode = '42501';
  end if;
  if p_source_order_id is not distinct from p_target_order_id then
    raise exception 'اختاري وجهة أخرى للنقل';
  end if;
  select count(distinct id) into requested from unnest(p_purchase_ids) as ids(id);
  if requested = 0 or array_position(p_purchase_ids, null) is not null then
    raise exception 'اختاري المشتريات المراد نقلها';
  end if;

  perform 1 from public.home_cash_account where id = 1 for update;
  perform 1 from public.orders where id in (p_source_order_id, p_target_order_id) order by id for update;
  select count(*) into matched from public.orders where id in (p_source_order_id, p_target_order_id);
  if matched <> (case when p_source_order_id is null then 0 else 1 end
    + case when p_target_order_id is null then 0 else 1 end) then
    raise exception 'الطلب غير موجود';
  end if;
  if exists(select 1 from public.orders where id in (p_source_order_id, p_target_order_id)
    and (coalesce(arrived,false) or coalesce(placed_at_pickup,false))) then
    raise exception 'النقل متاح للطلبات قيد الانتظار فقط';
  end if;
  perform 1 from public.purchases where id = any(p_purchase_ids) order by id for update;
  select count(*) into matched from public.purchases
    where id = any(p_purchase_ids) and order_id is not distinct from p_source_order_id;
  if matched <> requested then raise exception 'تغيرت المشتريات، أعيدي تحميل القائمة'; end if;
  if exists(select 1 from public.purchases p where p.id = any(p_purchase_ids)
    and (p.ready_for_pickup or p.picked_up or p.collected
      or exists(select 1 from public.purchase_payments pay where pay.purchase_id = p.id))) then
    raise exception 'لا يمكن نقل مشتريات جاهزة للاستلام أو مستلمة أو لها سجل دفعات';
  end if;

  update public.purchases set order_id = p_target_order_id where id = any(p_purchase_ids);
  if p_source_order_id is not null then perform public.refresh_order_profit(p_source_order_id); end if;
  if p_target_order_id is not null then perform public.refresh_order_profit(p_target_order_id); end if;
  return requested;
end;
$$;
revoke all on function public.transfer_order_purchases(uuid,uuid,uuid[]) from public,anon,authenticated;
grant execute on function public.transfer_order_purchases(uuid,uuid,uuid[]) to authenticated;
notify pgrst, 'reload schema';
commit;
