create extension if not exists "pgcrypto";

create table if not exists sticky_notes (
  id uuid primary key default gen_random_uuid(),
  customer_id text not null,
  customer_name text not null,
  text text not null,
  priority text not null check (priority in ('low', 'medium', 'high', 'urgent')),
  status text not null default 'open' check (status in ('open', 'done')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists sticky_notes_customer_id_idx on sticky_notes (customer_id);
create index if not exists sticky_notes_status_idx on sticky_notes (status);

create or replace function set_sticky_notes_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

do $$
begin
  if not exists (
    select 1 from pg_trigger where tgname = 'sticky_notes_updated_at'
  ) then
    create trigger sticky_notes_updated_at
      before update on sticky_notes
      for each row
      execute function set_sticky_notes_updated_at();
  end if;
end;
$$;
