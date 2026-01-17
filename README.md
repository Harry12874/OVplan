# Orchard Valley Planner

Local-first planning app for Orchard Valley Australia. Track orders from order → pack → deliver, keep cadence expectations, and export delivery routes.

## Getting started

### Option A: open directly
Open `index.html` in Chrome/Edge.

### Option B: local server (recommended)
```bash
python -m http.server 8000
```
Then visit `http://localhost:8000`.

## Data storage
All data is stored locally in your browser (IndexedDB with localStorage fallback). Clearing browser data will remove records.
Sticky notes are scoped per user via Supabase RLS; `user_id` defaults to `auth.uid()` on insert.

## Backup & restore
Use the **Backup & Restore** tab to download a JSON backup and restore it later. Consider running a backup weekly.

## CSV export
The **Export** tab lets you:
- Select date range and rep
- Include or exclude completed deliveries
- Configure output columns using the column mapping editor

Presets are saved in your browser and can be edited anytime.

## Customer CSV import
Use **Customers → Upload CSV** to import `ov_customers.csv`-style files. The importer:
- Previews the first 20 rows with validation errors.
- Upserts by email (preferred) or store name + address + postcode.
- Normalizes day arrays to weekdays only (Mon–Fri).

**Accepted header aliases (case/spacing/underscore-insensitive)**
- Store name: `store_name`, `storeName`, `storename`, `store`
- Address: `fullAddress` (preferred), `address`, `address1`
- Suburb: `suburb`, `suburb1`
- State: `state`, `state1`
- Postcode: `postcode`, `postcode1`
- Contact name: `contact_name`, `contactName`
- Delivery terms: `delivery_terms`, `deliveryTerms`
- Order source: `order_source`, `orderChannel`
- Rep name: `rep_name`, `assignedRepName`
- Days: `order_days`/`schedule_customerOrderDays`, `packing_days`/`schedule_packDays`, `delivery_days`/`schedule_deliverDays`
- Notes: `deliveryNotes`, `customerNotes`, `notes`
- Extra fields JSON: `extraFields_json`

**CSV → DB mapping**
- `store_name` → `customers.store_name` / `storeName`
- `contact_name` → `customers.contact_name` / `contactName`
- `phone` → `customers.phone` / `phone`
- `email` → `customers.email` / `email`
- `address` → `customers.address` / `fullAddress` + `address1`
- `suburb` → `customers.suburb` / `suburb1`
- `state` → `customers.state` / `state1`
- `postcode` → `customers.postcode` / `postcode1`
- `order_source` → `customers.order_source` / `orderChannel`
- `delivery_terms` → `customers.delivery_terms` / `deliveryTerms`
- `notes` → `customerNotes`
- `packing_days` → `customers.packing_days` / `schedule.packDays`
- `delivery_days` → `customers.delivery_days` / `schedule.deliverDays`
- `order_days` → `customers.order_days` / `schedule.customerOrderDays`
- `rep_name` → `customers.rep_name` / `repName` (matched against rep name for default rep selection)
- `extraFields_json` → `extraFields` (raw JSON string)
- Unknown columns → merged into `extraFields`

## Supabase migrations
Run `supabase/customer_data_scope.sql` in the Supabase SQL editor to add `user_id`, RLS policies, and customer scoping.

## Sample data
Use **Load sample data** to populate demo reps, customers, and orders.

## Melbourne timezone
Dates are computed using Australia/Melbourne local time, with due dates stored as date-only strings.
