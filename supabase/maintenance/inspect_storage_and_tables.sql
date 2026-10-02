-- Read-only diagnostics. Run in the Supabase SQL Editor; share results before dropping anything.
select schemaname, relname as table_name,
       pg_size_pretty(pg_total_relation_size(relid)) as total_size,
       pg_size_pretty(pg_relation_size(relid)) as data_size,
       pg_size_pretty(pg_indexes_size(relid)) as index_size,
       n_live_tup as estimated_rows, n_dead_tup as estimated_dead_rows
from pg_stat_user_tables
where schemaname = 'public'
order by pg_total_relation_size(relid) desc;

select tablename, indexname, indexdef
from pg_indexes
where schemaname = 'public'
order by tablename, indexname;

-- These are review candidates, NOT instructions to delete files.
-- Storage objects must be deleted through the Storage API or Dashboard, never SQL.
select o.name, o.created_at, o.metadata ->> 'size' as bytes
from storage.objects o
where o.bucket_id = 'purchase-images'
  and not exists (
    select 1 from public.purchase_images i where i.storage_path = o.name
  )
order by o.created_at;
