-- Allow duplicate emails; retain store/address conflict target for CSV imports.
drop index if exists public.customers_unique_email;

create unique index if not exists customers_unique_store_address
on public.customers (store_name, address, suburb, state, postcode);
