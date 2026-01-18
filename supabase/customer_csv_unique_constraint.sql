-- Ensure deterministic upsert conflicts for CSV imports (per-user scope)
create unique index if not exists customers_unique_store_address
on public.customers (user_id, store_name, address, suburb, state, postcode);
