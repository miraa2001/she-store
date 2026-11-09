begin;

-- Future payments only. Existing receipt, cash and profit records are not rewritten.
create or replace function public.record_purchase_payment(p_request_id uuid,p_id uuid,p_amount numeric,p_receive boolean,p_instant boolean default false)
returns jsonb language plpgsql security definer set search_path='' as $$
declare target jsonb; prior public.purchase_payments%rowtype; agreed numeric; source text; changed boolean;
begin
  perform 1 from public.home_cash_account where id=1 for update;
  target := public.payment_target(p_id,p_instant);
  if not public.can_record_purchase_payments(target->>'pickup_point') then raise exception 'تسجيل الدفعات متاح لرهف وريم فقط ضمن أماكن الاستلام المسموحة'; end if;
  if p_request_id is null or p_instant is null or p_amount is null or p_amount<=0 or p_amount<>round(p_amount,2)
    or p_amount::text in ('NaN','Infinity','-Infinity') or p_receive is null then raise exception 'أدخلي مبلغاً صحيحاً أكبر من صفر بمنزلتين عشريتين'; end if;
  select * into prior from public.purchase_payments where id=p_request_id;
  if found then
    if prior.amount<>p_amount or prior.received_on_pickup<>p_receive or prior.created_by is distinct from auth.uid()
      or (case when p_instant then prior.instant_pickup_id is distinct from p_id else prior.purchase_id is distinct from p_id end)
      then raise exception 'معرّف الدفعة مستخدم بقيم مختلفة'; end if;
    return jsonb_build_object('purchase',target,'changed_receipt',false);
  end if;
  if (target->>'collected')::boolean then raise exception 'المشترى مكتمل الدفع والتحصيل'; end if;
  if not p_instant and not coalesce((target->>'ready_for_pickup')::boolean,false) then raise exception 'المشترى غير جاهز للاستلام'; end if;
  if not (target->>'picked_up')::boolean and not p_receive then raise exception 'يجب تسجيل أول دفعة مع الاستلام'; end if;
  agreed := case when p_instant then (target->>'price')::numeric else coalesce((target->>'paid_price')::numeric,(target->>'price')::numeric,0) end;
  if p_amount>round(agreed,2)-(target->>'received_amount')::numeric then raise exception 'الدفعة أكبر من المبلغ المتبقي'; end if;
  source := public.home_cash_receipt_source(target->>'pickup_point');
  if source is null then raise exception 'مكان الاستلام غير صحيح'; end if;
  -- Customer payment and business collection are separate, including at Home.
  insert into public.purchase_payments(id,purchase_id,instant_pickup_id,amount,pickup_point,received_on_pickup,created_by)
    values(p_request_id,case when not p_instant then p_id end,case when p_instant then p_id end,p_amount,target->>'pickup_point',p_receive,auth.uid());
  changed := p_receive and not (target->>'picked_up')::boolean;
  target := public.refresh_purchase_payment_totals(p_id,p_instant,p_receive);
  return jsonb_build_object('purchase',target,'changed_receipt',changed);
end;
$$;

create or replace function public.guard_purchase_payments()
returns trigger language plpgsql set search_path='' as $$
declare instant boolean := tg_table_name='instant_pickups'; agreed numeric; has_history boolean;
begin
  if tg_op='DELETE' then
    if exists(select 1 from public.purchase_payments where purchase_id=old.id or instant_pickup_id=old.id) then
      raise exception 'لا يمكن حذف مشترى له سجل دفعات';
    end if;
    return old;
  end if;
  if current_user in ('authenticated','anon') then
    if tg_op='INSERT' then
      if new.received_amount<>0 or new.remitted_amount<>0 or new.picked_up or new.collected or new.collected_at is not null then
        raise exception 'يجب تسجيل الاستلام من نافذة الدفعات'; end if;
    elsif new.received_amount is distinct from old.received_amount or new.remitted_amount is distinct from old.remitted_amount
      or new.picked_up is distinct from old.picked_up or new.picked_up_at is distinct from old.picked_up_at
      or new.collected is distinct from old.collected or new.collected_at is distinct from old.collected_at
      or new.payment_model_ready is distinct from old.payment_model_ready then
      raise exception 'يجب تسجيل الاستلام والتحصيل من الإجراءات المخصصة';
    end if;
  end if;
  if tg_op='UPDATE' then
    if new.picked_up is distinct from old.picked_up and new.received_amount=old.received_amount and new.received_amount>0 then
      raise exception 'للتراجع عن الاستلام يجب إلغاء الدفعات أولاً'; end if;
    if not instant and coalesce((to_jsonb(old)->>'ready_for_pickup')::boolean,false)
      and not coalesce((to_jsonb(new)->>'ready_for_pickup')::boolean,false)
      and exists(select 1 from public.purchase_payments where purchase_id=old.id) then
      raise exception 'لا يمكن إلغاء الجاهزية لمشترى له سجل دفعات'; end if;
    if new.pickup_point is distinct from old.pickup_point then
      select exists(select 1 from public.purchase_payments where purchase_id=old.id or instant_pickup_id=old.id) into has_history;
      if has_history then raise exception 'لا يمكن نقل مشترى بعد تسجيل دفعة'; end if;
    end if;
  end if;
  agreed := case when instant then (to_jsonb(new)->>'price')::numeric
    else coalesce((to_jsonb(new)->>'paid_price')::numeric,(to_jsonb(new)->>'price')::numeric,0) end;
  if new.received_amount<0 or new.remitted_amount<0 or new.remitted_amount>new.received_amount or new.received_amount>round(agreed,2) then
    raise exception 'السعر لا يمكن أن يكون أقل من الدفعات المسجلة'; end if;
  if new.picked_up and new.received_amount<=0 and agreed>0 then raise exception 'يجب تسجيل دفعة قبل الاستلام'; end if;
  new.payment_model_ready := true;
  new.collected := new.picked_up and new.received_amount>=round(agreed,2) and new.remitted_amount>=new.received_amount;
  -- Zero-value Home handovers still require the separate collection action.
  if public.home_cash_receipt_source(new.pickup_point)='home' and agreed=0 then
    new.collected := new.collected and new.collected_at is not null;
  end if;
  new.collected_at := case when new.collected then coalesce(new.collected_at,clock_timestamp()) else null end;
  return new;
end;
$$;

create or replace function public.collect_purchase_payments(p_ids uuid[],p_instant boolean default false)
returns numeric language plpgsql security definer set search_path='' as $$
declare item uuid; target jsonb; entry public.purchase_payments%rowtype; cash_id uuid; total numeric:=0; agreed numeric; source text;
begin
  if not coalesce(public.is_admin(),false) then raise exception 'التحصيل متاح لرهف فقط'; end if;
  if p_instant is null then raise exception 'نوع المشترى غير صحيح'; end if;
  if p_ids is null or cardinality(p_ids)=0 or array_position(p_ids,null) is not null then raise exception 'اختاري مشتريات للتحصيل'; end if;
  perform 1 from public.home_cash_account where id=1 for update;
  for item in select distinct x from unnest(p_ids) x order by x loop
    target := public.payment_target(item,p_instant);
    source := public.home_cash_receipt_source(target->>'pickup_point');
    for entry in select * from public.purchase_payments where cancelled_at is null and remitted_at is null
      and (case when p_instant then instant_pickup_id=item else purchase_id=item end) order by id for update loop
      cash_id := public.payment_cash_entry(target,p_instant,entry.amount,
        case when source='home' then 'تحصيل دفعة - البيت' else 'تحصيل دفعة من نقطة الاستلام أو التوصيل' end);
      update public.purchase_payments set remitted_at=clock_timestamp(),remitted_by=auth.uid(),
        cash_recognized_amount=amount,cash_transaction_id=cash_id where id=entry.id;
      total := total+entry.amount;
    end loop;
    agreed := case when p_instant then (target->>'price')::numeric
      else coalesce((target->>'paid_price')::numeric,(target->>'price')::numeric,0) end;
    if source='home' and agreed=0 then
      if (target->>'picked_up')::boolean and not (target->>'collected')::boolean then
        execute format('update public.%I set collected_at=clock_timestamp() where id=$1',
          case when p_instant then 'instant_pickups' else 'purchases' end) using item;
      end if;
    else
      perform public.refresh_purchase_payment_totals(item,p_instant);
    end if;
  end loop;
  return total;
end;
$$;

revoke all on function public.record_purchase_payment(uuid,uuid,numeric,boolean,boolean),
  public.guard_purchase_payments(),public.collect_purchase_payments(uuid[],boolean) from public,anon,authenticated;
grant execute on function public.record_purchase_payment(uuid,uuid,numeric,boolean,boolean),
  public.collect_purchase_payments(uuid[],boolean) to authenticated;
notify pgrst,'reload schema';
commit;
