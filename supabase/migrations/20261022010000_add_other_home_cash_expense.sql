begin;

alter table public.home_cash_transactions
  drop constraint if exists home_cash_transactions_category_check;
alter table public.home_cash_transactions
  add constraint home_cash_transactions_category_check
  check (category in ('bags','pins','name_stickers','postal_delivery','atm_deposit','other'));

create or replace function public.record_home_cash_expense(
  p_request_id uuid, p_category text, p_amount numeric, p_quantity integer, p_note text default null
)
returns public.home_cash_transactions language plpgsql security definer set search_path = '' as $$
declare
  existing public.home_cash_transactions%rowtype;
  saved public.home_cash_transactions%rowtype;
  available numeric;
  cash_actor_email text;
begin
  if not coalesce(public.is_admin(), false) then raise exception 'لا توجد صلاحية' using errcode = '42501'; end if;
  if p_request_id is null or p_amount is null or p_amount <= 0 or p_amount <> round(p_amount, 2)
    or p_amount::text in ('NaN', 'Infinity', '-Infinity') or p_quantity is null or p_quantity <= 0
    or p_category is null or p_category not in ('bags', 'pins', 'name_stickers', 'postal_delivery', 'atm_deposit', 'other')
    or length(coalesce(p_note, '')) > 500 then raise exception 'بيانات المصروف غير صحيحة'; end if;

  perform 1 from public.home_cash_account where id = 1 for update;
  if not found then raise exception 'تتبع مصاري البيت غير متاح'; end if;
  select * into existing from public.home_cash_transactions where id = p_request_id;
  if found then
    if existing.kind = 'expense' and existing.amount = -p_amount and existing.category = p_category
      and existing.quantity = p_quantity and coalesce(existing.note, '') = coalesce(btrim(p_note), '') then return existing; end if;
    raise exception 'هذا المصروف محفوظ مسبقاً بقيم مختلفة';
  end if;
  select coalesce(sum(amount), 0) into available from public.home_cash_transactions;
  if available < p_amount then raise exception 'المبلغ أكبر من الرصيد المتاح في البيت'; end if;
  select email into cash_actor_email from auth.users where id = auth.uid();
  insert into public.home_cash_transactions (id, kind, amount, category, quantity, note, created_by, actor_email)
  values (p_request_id, 'expense', -p_amount, p_category, p_quantity, nullif(btrim(p_note), ''), auth.uid(), cash_actor_email)
  returning * into saved;
  return saved;
end;
$$;

revoke all on function public.record_home_cash_expense(uuid,text,numeric,integer,text) from public,anon,authenticated;
grant execute on function public.record_home_cash_expense(uuid,text,numeric,integer,text) to authenticated;
notify pgrst,'reload schema';
commit;
