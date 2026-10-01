# Flamingo Weather Intelligence

Classification: INTERNAL

Static HTML dashboard, split from the single-file export `Flamingo_Weather_Dashboard.html`.

## Structure

| Path | Content |
|---|---|
| `index.html` | Page markup (header, tabs, panels) |
| `css/app.css` | Dashboard styles (Flamingo palette) |
| `css/gridjs.css` | Grid.js styles (vendor) |
| `js/app.js` | Dashboard logic |
| `js/vendor/chart.js` | Chart.js 4.5.0 (vendor) |
| `js/vendor/gridjs.js` | Grid.js (vendor) |
| `data/locations_weather.js` | `DATA`: locations, observations, forecasts |
| `data/alerts.js` | `ALERTS`: weather alerts snapshot (17/09/2026) |
| `data/accuracy.js` | `ACC`: forecast accuracy (04/06/2026 to 06/07/2026) |
| `data/crops.js` | `CROP`: crop risk data |
| `assets/logo.png` | Header logo |

## Running

Open `index.html` in a browser. No build step.

## Known limits

- Data files are static snapshots embedded in the export. Nothing refreshes them in this repository.
- `/api/whoami` and `/api/actioned` (Admin and Actioned tabs) and the Power BI refresh button need the original hosting environment. They fail when opened from disk.
