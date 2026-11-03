begin;

alter table public.purchases
  add column if not exists purchase_categories text[] not null default '{}';
alter table public.purchases drop constraint if exists purchases_categories_check;
alter table public.purchases add constraint purchases_categories_check check (
  purchase_categories <@ array['fast','makeup','electronics']::text[]
  and cardinality(purchase_categories) <= 3
  and coalesce(array_ndims(purchase_categories),1) = 1
  and array_position(purchase_categories,null) is null
);

notify pgrst, 'reload schema';
commit;
