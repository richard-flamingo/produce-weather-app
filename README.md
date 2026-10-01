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
| `data/locations_weather.json` | `DATA`: locations, observations, forecasts |
| `data/alerts.json` | `ALERTS`: weather alerts snapshot (17/09/2026) |
| `data/accuracy.json` | `ACC`: forecast accuracy (04/06/2026 to 06/07/2026) |
| `data/crops.json` | `CROP`: crop risk data |
| `js/bootstrap.js` | Fetches the JSON files, then starts `app.js` |
| `assets/logo.png` | Header logo |

## Running

Serve the folder over HTTP, then open `index.html` (for example `http://flm-c2l-apps:5000/`). No build step.

```
python3 -m http.server 8000
```

Opening the file directly from disk does not work: browsers block JSON loading from `file://`.

## Known limits

- Data files are static snapshots embedded in the export. Nothing refreshes them in this repository.
- There is no access control. The sign-in check and the Admin tab were removed, so anyone who can reach the site sees every tab. Restrict access at the web server or network level.
- `/api/actioned` (Actioned tab) and the Power BI refresh button need the original hosting environment. They fail on a plain static host.
