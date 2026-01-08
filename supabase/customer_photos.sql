create extension if not exists "pgcrypto";

create table if not exists customer_photos (
  id uuid primary key default gen_random_uuid(),
  customer_id uuid not null references customers (id) on delete cascade,
  image_path text not null,
  caption text,
  store_location text,
  created_by uuid,
  created_at timestamptz not null default now()
);

create index if not exists customer_photos_customer_id_idx on customer_photos (customer_id);
create index if not exists customer_photos_created_at_idx on customer_photos (created_at desc);
