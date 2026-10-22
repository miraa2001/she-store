begin;

-- Allow receipt changes without giving viewers general UPDATE access to purchases.
create or replace function public.set_purchase_received(p_purchase_id uuid, p_picked_up boolean)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare
  actor_role text;
  actor_email text;
  source text;
  saved public.purchases%rowtype;
  changed boolean;
begin
  if auth.uid() is null then raise exception 'يجب تسجيل الدخول أولاً'; end if;
  if p_purchase_id is null or p_picked_up is null then raise exception 'بيانات الاستلام غير صالحة'; end if;
  select r.role, lower(u.email) into actor_role, actor_email
    from public.user_roles r join auth.users u on u.id = r.user_id
    where r.user_id = auth.uid();

  select * into saved from public.purchases where id = p_purchase_id for update;
  if not found then raise exception 'المشترى غير موجود'; end if;
  source := public.home_cash_receipt_source(saved.pickup_point);
  if not coalesce(
    actor_role = 'admin'
    or (actor_role = 'viewer' and actor_email = 'reem@she-store.com' and source in ('home', 'nablus', 'delivery'))
    or (actor_role = 'pickup' and actor_email = 'maryamti@she-store.com' and source = 'maryamti')
    or (actor_role = 'pickup' and actor_email = 'nablus@she-store.com' and source = 'nablus'), false
  ) then raise exception 'ليس لديك صلاحية تغيير الاستلام في هذا المكان'; end if;
  if saved.collected then raise exception 'لا يمكن تغيير استلام مشترى تم تحصيله'; end if;
  if not coalesce(saved.ready_for_pickup, false) then raise exception 'المشترى غير جاهز للاستلام'; end if;

  changed := saved.picked_up is distinct from p_picked_up;
  if changed then
    update public.purchases set picked_up = p_picked_up,
      picked_up_at = case when p_picked_up then clock_timestamp() else null end
      where id = saved.id returning * into saved;
    if not found or saved.picked_up is distinct from p_picked_up then
      raise exception 'تعذر تأكيد حفظ حالة الاستلام';
    end if;
  end if;
  return jsonb_build_object('id', saved.id, 'picked_up', saved.picked_up,
    'picked_up_at', saved.picked_up_at, 'changed', changed);
end;
$$;

revoke all on function public.set_purchase_received(uuid, boolean) from public, anon, authenticated;
grant execute on function public.set_purchase_received(uuid, boolean) to authenticated;

commit;
