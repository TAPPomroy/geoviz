# GeoViz — Project Guide

## What This Is

An Excel VBA workbook that reads a company table (name, address, ISP, attributes), geocodes addresses via Nominatim (caching lat/lon back in the sheet), and generates a standalone HTML/Leaflet.js map opened in the default browser. Primary use: click a company during an ISP outage, see geographically nearby companies on the same ISP.

## Stack

- **Runtime**: Pure VBA (Excel 2016+ on Windows) — no Python, no add-ins
- **HTTP**: `WinHttp.WinHttpRequest.5.1` for Nominatim geocoding
- **File output**: `ADODB.Stream` with `Charset = "utf-8"` (never FSO)
- **Map**: Leaflet.js 1.9.x loaded from CDN
- **Basemap**: CartoDB Positron (muted background, better marker legibility)
- **Geocoding**: Nominatim (OpenStreetMap) — 1 req/sec throttle, User-Agent required

## Project Structure

```
.planning/
  PROJECT.md        — project context and decisions
  REQUIREMENTS.md   — v1 requirements with REQ-IDs
  ROADMAP.md        — 3-phase execution plan
  STATE.md          — current phase and status
  config.json       — workflow settings
  research/         — stack, features, architecture, pitfalls research
```

## GSD Workflow

This project uses the GSD (Get Shit Done) workflow:

```
/gsd:discuss-phase N    — gather context before planning
/gsd:plan-phase N       — create PLAN.md for a phase
/gsd:execute-phase N    — execute the plan
/gsd:verify-work N      — verify phase deliverables
/gsd:progress           — check current status
```

## Phases

| Phase | Goal | Requirements |
|-------|------|--------------|
| 1 — Data Pipeline | Read sheet, geocode, cache | DATA-01–04 |
| 2 — Map Rendering | Generate HTML, markers, button | MAP-01–05 |
| 3 — Interactivity | Click, radius, highlight, filter | INT-01–04, FILT-01–02 |

## Key Constraints

- **ADODB.Stream only** for file output — FSO defaults to ANSI and breaks accented characters
- **Haversine distance** — not flat-earth; meaningful difference at 50–100 km scale
- **Nominatim rate limit** — 1 req/sec hard limit; `Application.Wait` between each call
- **Mark of the Web** — workbook must be unblocked (right-click → Properties → Unblock) before macros run
- **file:// CORS** — tile layers may not load when HTML opened directly; test in Chrome early
- **7-color ISP palette** — deterministic (alphabetical sort), "Other" gray for overflow; never exceed 7 named colors

## Critical Pitfalls

1. Nominatim without User-Agent header → 403 or silent failures
2. FSO for file output → ANSI encoding → garbled company names
3. Not caching lat/lon → re-geocodes everything on every run → hits rate limit
4. Flat-earth proximity → 1–3% error at regional distances → use haversine
5. WebBrowser ActiveX control → broken in current Excel/Windows → open in browser instead
