-- Customer scoping + safe wipe support (run in Supabase SQL editor)

-- 1) customers: add user_id if missing
alter table if exists public.customers
add column if not exists user_id uuid;

-- backfill (only if you have single-user data; otherwise skip and do manual mapping)
-- update public.customers set user_id = auth.uid() where user_id is null;

-- 2) schedule_events: add user_id + customer_id FK if missing
alter table if exists public.schedule_events
add column if not exists user_id uuid;

alter table if exists public.schedule_events
add column if not exists customer_id uuid;

alter table if exists public.schedule_events
drop constraint if exists schedule_events_customer_id_fkey;

alter table if exists public.schedule_events
add constraint schedule_events_customer_id_fkey
foreign key (customer_id) references public.customers(id)
on delete cascade;

-- 3) Helpful indexes
create index if not exists idx_customers_user_id on public.customers(user_id);
create index if not exists idx_schedule_events_user_id on public.schedule_events(user_id);
create index if not exists idx_schedule_events_customer_id on public.schedule_events(customer_id);

-- 4) Enable RLS + policies
alter table public.customers enable row level security;
alter table public.schedule_events enable row level security;

drop policy if exists "customers_select_own" on public.customers;
create policy "customers_select_own"
on public.customers for select
using (auth.uid() = user_id);

drop policy if exists "customers_insert_own" on public.customers;
create policy "customers_insert_own"
on public.customers for insert
with check (auth.uid() = user_id);

drop policy if exists "customers_update_own" on public.customers;
create policy "customers_update_own"
on public.customers for update
using (auth.uid() = user_id);

drop policy if exists "customers_delete_own" on public.customers;
create policy "customers_delete_own"
on public.customers for delete
using (auth.uid() = user_id);

drop policy if exists "schedule_select_own" on public.schedule_events;
create policy "schedule_select_own"
on public.schedule_events for select
using (auth.uid() = user_id);

drop policy if exists "schedule_insert_own" on public.schedule_events;
create policy "schedule_insert_own"
on public.schedule_events for insert
with check (auth.uid() = user_id);

drop policy if exists "schedule_update_own" on public.schedule_events;
create policy "schedule_update_own"
on public.schedule_events for update
using (auth.uid() = user_id);

drop policy if exists "schedule_delete_own" on public.schedule_events;
create policy "schedule_delete_own"
on public.schedule_events for delete
using (auth.uid() = user_id);
