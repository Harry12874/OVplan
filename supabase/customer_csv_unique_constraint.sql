-- Ensure deterministic upsert conflicts for CSV imports (shared scope)
create unique index if not exists customers_unique_email
on public.customers (email)
where email is not null;

create unique index if not exists customers_unique_store_address
on public.customers (store_name, address, suburb, state, postcode);
