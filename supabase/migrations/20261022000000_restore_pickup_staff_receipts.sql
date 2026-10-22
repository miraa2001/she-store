begin;

-- Staff confirm ordinary, fully-paid handovers. They cannot enter custom installments.
create or replace function public.confirm_pickup_staff_receipt(p_request_id uuid,p_id uuid,p_instant boolean default false)
returns jsonb language plpgsql security definer set search_path='' as $$
declare
  actor_role text;
  actor_email text;
  source text;
  target jsonb;
  prior public.purchase_payments%rowtype;
  agreed numeric;
  amount_due numeric;
begin
  if auth.uid() is null or p_request_id is null or p_id is null or p_instant is null then
    raise exception 'بيانات الاستلام غير صالحة';
  end if;
  select r.role,lower(u.email) into actor_role,actor_email
    from public.user_roles r join auth.users u on u.id=r.user_id where r.user_id=auth.uid();
  perform 1 from public.home_cash_account where id=1 for update;
  target := public.payment_target(p_id,p_instant);
  source := public.home_cash_receipt_source(target->>'pickup_point');
  if not coalesce(actor_role='pickup' and (
    (actor_email='maryamti@she-store.com' and source='maryamti')
    or (actor_email='nablus@she-store.com' and source='nablus')
  ),false) then raise exception 'يمكن تأكيد الاستلام في نقطة الاستلام الخاصة بحسابك فقط'; end if;

  select * into prior from public.purchase_payments where id=p_request_id;
  if found then
    if prior.created_by is distinct from auth.uid() or not prior.received_on_pickup or prior.legacy
      or (case when p_instant then prior.instant_pickup_id is distinct from p_id else prior.purchase_id is distinct from p_id end)
      then raise exception 'معرّف العملية مستخدم مسبقاً'; end if;
    return jsonb_build_object('purchase',target,'changed_receipt',false);
  end if;
  -- Never silently settle an existing partial-payment purchase or duplicate a handover.
  if (target->>'picked_up')::boolean then
    return jsonb_build_object('purchase',target,'changed_receipt',false);
  end if;
  if (target->>'collected')::boolean then raise exception 'المشترى مكتمل'; end if;
  if not p_instant and not coalesce((target->>'ready_for_pickup')::boolean,false) then
    raise exception 'المشترى غير جاهز للاستلام'; end if;
  agreed := round(case when p_instant then (target->>'price')::numeric
    else coalesce((target->>'paid_price')::numeric,(target->>'price')::numeric,0) end,2);
  amount_due := agreed-(target->>'received_amount')::numeric;
  if amount_due<0 then raise exception 'مبلغ الاستلام غير صحيح'; end if;
  if amount_due>0 then
    insert into public.purchase_payments(id,purchase_id,instant_pickup_id,amount,pickup_point,received_on_pickup,created_by)
    values(p_request_id,case when not p_instant then p_id end,case when p_instant then p_id end,
      amount_due,target->>'pickup_point',true,auth.uid());
  end if;
  target := public.refresh_purchase_payment_totals(p_id,p_instant,true);
  if not (target->>'picked_up')::boolean or (target->>'received_amount')::numeric<agreed then
    raise exception 'تعذر تأكيد حفظ الاستلام والدفع'; end if;
  return jsonb_build_object('purchase',target,'changed_receipt',true);
end;
$$;

revoke all on function public.confirm_pickup_staff_receipt(uuid,uuid,boolean) from public,anon,authenticated;
grant execute on function public.confirm_pickup_staff_receipt(uuid,uuid,boolean) to authenticated;
notify pgrst,'reload schema';
commit;
