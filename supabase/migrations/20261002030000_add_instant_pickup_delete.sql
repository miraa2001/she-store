begin;

create or replace function public.delete_instant_pickup(p_id uuid)
returns boolean language plpgsql security definer set search_path = '' as $$
declare
  saved public.instant_pickups%rowtype;
begin
  if not coalesce(public.is_admin(), false) then
    raise exception 'لا توجد صلاحية' using errcode = '42501';
  end if;
  if p_id is null then raise exception 'المستلم غير صحيح'; end if;
  select * into saved from public.instant_pickups where id = p_id for update;
  -- Retrying a successful deletion is harmless; cash history is never deleted.
  if not found then return false; end if;
  if saved.picked_up or saved.collected then
    raise exception 'لا يمكن حذف مستلم تم استلامه أو تحصيله. ألغِ الاستلام أولاً إن لم يتم التحصيل.';
  end if;
  delete from public.instant_pickups where id = p_id;
  return true;
end;
$$;

revoke all on function public.delete_instant_pickup(uuid) from public, anon, authenticated;
grant execute on function public.delete_instant_pickup(uuid) to authenticated;

notify pgrst, 'reload schema';
commit;
