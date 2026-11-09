begin;

-- New endpoint only: no historical backfill and no changes to cash/profit algorithms.
create or replace function public.confirm_full_purchase_receipt(p_request_id uuid,p_id uuid,p_instant boolean default false)
returns jsonb language plpgsql security definer set search_path='' as $$
declare target jsonb; prior public.purchase_payments%rowtype; due numeric;
begin
  if auth.uid() is null or p_request_id is null or p_id is null or p_instant is null then
    raise exception 'بيانات الاستلام غير صالحة';
  end if;
  perform 1 from public.home_cash_account where id=1 for update;
  target:=public.payment_target(p_id,p_instant);
  if not public.can_record_purchase_payments(target->>'pickup_point') then
    return public.confirm_pickup_staff_receipt(p_request_id,p_id,p_instant);
  end if;
  select * into prior from public.purchase_payments where id=p_request_id;
  if found then
    if prior.created_by is distinct from auth.uid() or not prior.received_on_pickup or prior.legacy
      or prior.cancelled_at is not null
      or (case when p_instant then prior.instant_pickup_id is distinct from p_id else prior.purchase_id is distinct from p_id end) then
      raise exception 'معرّف العملية مستخدم مسبقاً';
    end if;
    return jsonb_build_object('purchase',target,'changed_receipt',false);
  end if;
  -- A stale unchecked UI must never silently settle a purchase already received partially.
  if (target->>'picked_up')::boolean then
    return jsonb_build_object('purchase',target,'changed_receipt',false);
  end if;
  if (target->>'collected')::boolean then raise exception 'المشترى مكتمل'; end if;
  if not p_instant and not coalesce((target->>'ready_for_pickup')::boolean,false) then
    raise exception 'المشترى غير جاهز للاستلام';
  end if;
  due:=round(case when p_instant then (target->>'price')::numeric
    else coalesce((target->>'paid_price')::numeric,(target->>'price')::numeric,0) end,2)
    - (target->>'received_amount')::numeric;
  if due<0 then raise exception 'مبلغ الاستلام غير صحيح'; end if;
  if due>0 then
    return public.record_purchase_payment(p_request_id,p_id,due,true,p_instant);
  end if;
  -- Free purchases can be received without inventing a payment or cash entry.
  target:=public.refresh_purchase_payment_totals(p_id,p_instant,true);
  return jsonb_build_object('purchase',target,'changed_receipt',true);
end;
$$;
revoke all on function public.confirm_full_purchase_receipt(uuid,uuid,boolean) from public,anon,authenticated;
grant execute on function public.confirm_full_purchase_receipt(uuid,uuid,boolean) to authenticated;

notify pgrst,'reload schema';
commit;
