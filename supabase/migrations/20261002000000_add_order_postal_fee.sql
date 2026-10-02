alter table public.orders
  add column if not exists postal_fee numeric not null default 0
    constraint orders_postal_fee_nonnegative check (postal_fee >= 0);
