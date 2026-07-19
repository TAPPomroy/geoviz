# GeoViz — Company Map Explorer

## What This Is

A self-contained Excel workbook with a VBA macro that reads a company table (addresses + attributes), geocodes addresses once and caches lat/lon back in the sheet, then generates a standalone HTML/Leaflet map opened in the browser. Primary use case: during an ISP outage, select an affected company and instantly see which nearby companies share the same ISP provider.

## Core Value

Click any company on the map and immediately see which neighboring companies share its ISP — with an adjustable proximity radius.

## Requirements

### Validated

(None yet — ship to validate)

### Active

- [ ] Read company table from Excel sheet (name, address, ISP, other attributes)
- [ ] Geocode addresses via Nominatim API, store lat/lon back in Excel (cached — skips already-geocoded rows)
- [ ] Generate a self-contained HTML file with Leaflet.js interactive map
- [ ] Open generated HTML in default browser from Excel macro button
- [ ] Display company markers colored by ISP (or selected attribute)
- [ ] Click company → popup showing company name, ISP, and other attributes
- [ ] Click company → highlight nearby companies within a user-adjustable radius
- [ ] Highlight same-ISP neighbors distinctly from other nearby companies
- [ ] Attribute filter: switch coloring/filtering between ISP and other columns
- [ ] Proximity radius adjustable by slider or input on the map UI

### Out of Scope

- Python runtime dependency — pure VBA + HTML/JS, zero install required
- Excel add-in — VBA macro embedded in the workbook only
- Real-time data sync — map is generated on demand from current sheet state
- Mobile/tablet support — desktop browser only

## Context

- Platform: Windows, Excel (any recent version supporting VBA)
- Dataset scale: dozens of companies (< 100 rows) — performance is not a concern
- Geocoding: Nominatim (OpenStreetMap free API) — no API key required
- Map library: Leaflet.js loaded from CDN (or inlined for offline use)
- Primary attribute: ISP provider; designed to accommodate additional attributes as needs grow
- The "geoviz" project name reflects the geographic visualization focus

## Constraints

- **Tech stack**: Pure VBA + self-contained HTML/JavaScript — no Python, no add-ins, no external installs
- **Geocoding API**: Nominatim rate limit (1 req/sec) — fine for < 100 companies, must throttle requests
- **Offline use**: Generated HTML depends on Leaflet CDN unless tiles/library are inlined
- **Excel version**: Must work with Excel 2016+ on Windows

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Pure VBA over Python | Zero dependencies — runs on any Windows Excel without setup | — Pending |
| Nominatim for geocoding | Free, no API key, sufficient for small datasets | — Pending |
| Leaflet.js for map | Lightweight, well-documented, works in standalone HTML | — Pending |
| Cache lat/lon in Excel | Avoid re-geocoding on every run; preserves API rate limits | — Pending |

## Evolution

This document evolves at phase transitions and milestone boundaries.

**After each phase transition** (via `/gsd-transition`):
1. Requirements invalidated? → Move to Out of Scope with reason
2. Requirements validated? → Move to Validated with phase reference
3. New requirements emerged? → Add to Active
4. Decisions to log? → Add to Key Decisions
5. "What This Is" still accurate? → Update if drifted

**After each milestone** (via `/gsd:complete-milestone`):
1. Full review of all sections
2. Core Value check — still the right priority?
3. Audit Out of Scope — reasons still valid?
4. Update Context with current state

---
*Last updated: 2026-07-18 after initialization*
