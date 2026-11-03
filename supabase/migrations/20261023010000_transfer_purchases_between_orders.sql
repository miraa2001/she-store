begin;

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
  if p_source_order_id is null or p_target_order_id is null
    or p_source_order_id = p_target_order_id then
    raise exception 'اختاري طلباً آخر للنقل';
  end if;
  select count(distinct id) into requested from unnest(p_purchase_ids) as ids(id);
  if requested = 0 or array_position(p_purchase_ids, null) is not null then
    raise exception 'اختاري المشتريات المراد نقلها';
  end if;

  -- Match the cash/payment lock order before locking purchases and refreshing profits.
  perform 1 from public.home_cash_account where id = 1 for update;
  perform 1 from public.orders where id in (p_source_order_id, p_target_order_id) order by id for update;
  select count(*) into matched from public.orders
    where id in (p_source_order_id, p_target_order_id);
  if matched <> 2 then raise exception 'الطلب غير موجود'; end if;
  if exists(select 1 from public.orders where id in (p_source_order_id, p_target_order_id)
    and placed_at_pickup) then
    raise exception 'النقل متاح قبل وضع الطلبات في نقطة الاستلام فقط';
  end if;
  perform 1 from public.purchases where id = any(p_purchase_ids) order by id for update;
  select count(*) into matched from public.purchases
    where id = any(p_purchase_ids) and order_id = p_source_order_id;
  if matched <> requested then raise exception 'تغيرت المشتريات، أعيدي تحميل الطلب'; end if;
  if exists(select 1 from public.purchases p where p.id = any(p_purchase_ids)
    and (p.ready_for_pickup or p.picked_up or p.collected
      or exists(select 1 from public.purchase_payments pay where pay.purchase_id = p.id))) then
    raise exception 'لا يمكن نقل مشتريات جاهزة للاستلام أو مستلمة أو لها سجل دفعات';
  end if;

  update public.purchases set order_id = p_target_order_id where id = any(p_purchase_ids);
  perform public.refresh_order_profit(p_source_order_id);
  perform public.refresh_order_profit(p_target_order_id);
  return requested;
end;
$$;

revoke all on function public.transfer_order_purchases(uuid,uuid,uuid[]) from public,anon,authenticated;
grant execute on function public.transfer_order_purchases(uuid,uuid,uuid[]) to authenticated;
notify pgrst, 'reload schema';
commit;
