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

## Backup & restore
Use the **Backup & Restore** tab to download a JSON backup and restore it later. Consider running a backup weekly.

## CSV export
The **Export** tab lets you:
- Select date range and rep
- Include or exclude completed deliveries
- Configure output columns using the column mapping editor

Presets are saved in your browser and can be edited anytime.

## Sample data
Use **Load sample data** to populate demo reps, customers, and orders.

## Melbourne timezone
Dates are computed using Australia/Melbourne local time, with due dates stored as date-only strings.
