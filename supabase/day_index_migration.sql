-- Optional one-time migration
-- Use ONLY if existing data is stored as Sun0 (0=Sun..6=Sat).
-- This converts day arrays to Mon0 (0=Mon..6=Sun) using (old + 6) % 7.

update customers
set
  delivery_days = coalesce(
    (
      select array_agg((d + 6) % 7 order by (d + 6) % 7)
      from unnest(coalesce(delivery_days, '{}'::int[])) as d
    ),
    '{}'::int[]
  ),
  packing_days = coalesce(
    (
      select array_agg((d + 6) % 7 order by (d + 6) % 7)
      from unnest(coalesce(packing_days, '{}'::int[])) as d
    ),
    '{}'::int[]
  ),
  order_days = coalesce(
    (
      select array_agg((d + 6) % 7 order by (d + 6) % 7)
      from unnest(coalesce(order_days, '{}'::int[])) as d
    ),
    '{}'::int[]
  );
