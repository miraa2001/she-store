begin;

create table if not exists public.purchase_comparison_commits (
  id uuid primary key,
  payload jsonb not null,
  result jsonb not null,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default clock_timestamp()
);
alter table public.purchase_comparison_commits enable row level security;
revoke all on public.purchase_comparison_commits from anon,authenticated;
grant select on public.purchase_comparison_commits to authenticated;
drop policy if exists comparison_admin_read on public.purchase_comparison_commits;
create policy comparison_admin_read on public.purchase_comparison_commits for select to authenticated using(public.is_admin());

create or replace function public.apply_purchase_comparison(p_request_id uuid,p_groups jsonb)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare
  previous public.purchase_comparison_commits%rowtype;
  requested_ids uuid[];
  group_ids uuid[];
  seen_orders uuid[] := '{}';
  group_data jsonb;
  item jsonb;
  purchase public.purchases%rowtype;
  destination public.orders%rowtype;
  category text;
  subtype text;
  destination_id uuid;
  title text;
  target numeric;
  quantity integer;
  total numeric;
  matched integer;
  result_orders jsonb := '[]';
  response jsonb;
begin
  if not coalesce(public.is_admin(),false) then raise exception 'لا توجد صلاحية' using errcode='42501'; end if;
  if p_request_id is null or p_groups is null or jsonb_typeof(p_groups) <> 'array' then
    raise exception 'بيانات المقارنة غير صحيحة';
  end if;
  if jsonb_array_length(p_groups) < 1 or jsonb_array_length(p_groups) > 20 then raise exception 'عدد الطلبات غير صالح'; end if;

  -- Match the cash/payment lock order and serialize idempotent retries.
  perform 1 from public.home_cash_account where id=1 for update;
  if not found then raise exception 'حساب مصاري البيت غير متاح'; end if;
  select * into previous from public.purchase_comparison_commits where id=p_request_id;
  if found then
    if previous.payload <> p_groups or previous.created_by is distinct from auth.uid() then
      raise exception 'طلب الحفظ مستخدم ببيانات مختلفة';
    end if;
    return previous.result;
  end if;

  for group_data in select value from jsonb_array_elements(p_groups) loop
    if jsonb_typeof(group_data->'items') is distinct from 'array' then raise exception 'مشتريات الطلب غير صحيحة'; end if;
    if jsonb_array_length(group_data->'items') < 1 then raise exception 'لا يمكن تأكيد طلب فارغ'; end if;
  end loop;
  select array_agg((entry->>'id')::uuid) into requested_ids
    from jsonb_array_elements(p_groups) g cross join lateral jsonb_array_elements(g->'items') entry;
  if cardinality(requested_ids) > 2000 or array_position(requested_ids,null) is not null
    or (select count(distinct value) from unnest(requested_ids) value) <> cardinality(requested_ids) then
    raise exception 'قائمة المشتريات مكررة أو غير صالحة';
  end if;

  perform 1 from public.orders where id in (
    select nullif(value->>'order_id','')::uuid from jsonb_array_elements(p_groups)
  ) order by id for update;
  perform 1 from public.purchases where id=any(requested_ids) order by id for update;
  select count(*) into matched from public.purchases p where p.id=any(requested_ids)
    and p.order_id is null and not coalesce(p.ready_for_pickup,false) and not p.picked_up and not p.collected
    and not exists(select 1 from public.purchase_payments pay where pay.purchase_id=p.id);
  if matched <> cardinality(requested_ids) then raise exception 'تغيرت المشتريات. أعيدي تحميل القائمة والمقارنة'; end if;

  for group_data in select value from jsonb_array_elements(p_groups) loop
    category := group_data->>'category';
    subtype := case when category='fast' then 'normal' else category end;
    if category is null or category not in ('fast','makeup','electronics') then raise exception 'تصنيف غير صالح'; end if;
    target := (group_data->>'target')::numeric;
    if target is null or target <= 0 or target > 1000000000 or target <> round(target,2)
      or target::text in ('NaN','Infinity','-Infinity') then raise exception 'السعر المستهدف غير صالح'; end if;

    select array_agg((value->>'id')::uuid) into group_ids from jsonb_array_elements(group_data->'items');
    for item in select value from jsonb_array_elements(group_data->'items') loop
      select * into purchase from public.purchases where id=(item->>'id')::uuid;
      if purchase.price is distinct from (item->>'price')::numeric
        or purchase.qty is distinct from (item->>'qty')::integer
        or to_jsonb(purchase.purchase_categories) is distinct from item->'categories'
        or purchase.price < 0 or purchase.price::text in ('NaN','Infinity','-Infinity') then
        raise exception 'تغير السعر أو عدد القطع أو التصنيف. أعيدي المقارنة';
      end if;
      if cardinality(purchase.purchase_categories) > 0 and not (category=any(purchase.purchase_categories)) then
        raise exception 'لا يمكن خلط تصنيفات المشتريات';
      end if;
    end loop;
    select sum(qty),sum(round(price,2)) into quantity,total from public.purchases where id=any(group_ids);
    if quantity > 100 or quantity < 1 then raise exception 'الحد الأقصى للشحن 100 قطعة لكل طلب'; end if;
    destination_id := nullif(group_data->>'order_id','')::uuid;
    if destination_id is null then
      title := btrim(coalesce(group_data->>'name',''));
      if length(title) < 1 or length(title) > 200 then raise exception 'أدخلي اسم الطلب الجديد'; end if;
      destination_id := gen_random_uuid();
      insert into public.orders(id,order_name,order_date,order_type,shein_type,arrived,placed_at_pickup,
        spent_amount,postal_fee,marketing_fee)
        values(destination_id,title,(clock_timestamp() at time zone 'Asia/Hebron')::date,'shein',subtype,false,false,0,0,0);
    else
      if destination_id=any(seen_orders) then raise exception 'اختاري طلباً مختلفاً لكل مجموعة'; end if;
      select * into destination from public.orders where id=destination_id;
      if not found then raise exception 'الطلب غير موجود'; end if;
      if coalesce(destination.arrived,true) or coalesce(destination.placed_at_pickup,true)
        or destination.order_type is distinct from 'shein' or destination.shein_type is distinct from subtype
        or exists(select 1 from public.purchases where order_id=destination_id) then
        raise exception 'الوجهة يجب أن تكون طلباً فارغاً قيد الانتظار ومن نفس التصنيف';
      end if;
      title := destination.order_name;
    end if;
    seen_orders := array_append(seen_orders,destination_id);
    update public.purchases set order_id=destination_id where id=any(group_ids);
    perform public.refresh_order_profit(destination_id);
    result_orders := result_orders || jsonb_build_array(jsonb_build_object('order_id',destination_id,'name',title,
      'category',category,'target',target,'total',total,'qty',quantity,'count',cardinality(group_ids)));
  end loop;
  response := jsonb_build_object('orders',result_orders,'purchase_count',cardinality(requested_ids));
  insert into public.purchase_comparison_commits(id,payload,result,created_by) values(p_request_id,p_groups,response,auth.uid());
  return response;
end;
$$;
revoke all on function public.apply_purchase_comparison(uuid,jsonb) from public,anon,authenticated;
grant execute on function public.apply_purchase_comparison(uuid,jsonb) to authenticated;
notify pgrst,'reload schema';
commit;
