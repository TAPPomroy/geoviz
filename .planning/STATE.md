# GeoViz — Project State

## Current Status

Phase: Phase 2 Complete

## Project Reference

See: .planning/PROJECT.md

**Core value:** Click any company on the map and immediately see which neighboring companies share its ISP — with an adjustable proximity radius.

**Current focus:** Phase 3

---

## Current Position

| Field | Value |
|-------|-------|
| Phase | 2 — Map Rendering |
| Plan | 02-02-PLAN.md |
| Status | Complete |
| Progress | 2/3 phases complete |

Progress: [x] Phase 1 / [x] Phase 2 / [ ] Phase 3

---

## Performance Metrics

| Metric | Value |
|--------|-------|
| Phases defined | 3 |
| Requirements mapped | 15/15 |
| Plans created | 8 |
| Plans complete | 8 |

---

## Accumulated Context

### Key Decisions

| Decision | Rationale |
|----------|-----------|
| Pure VBA + HTML/JS | Zero dependencies — runs on any Windows Excel without setup |
| Nominatim geocoding | Free, no API key, sufficient for < 100 companies |
| Leaflet.js via CDN | Lightweight, well-documented, standalone HTML |
| Cache lat/lon in sheet | Avoid re-geocoding; preserve API rate limits |
| WinHttp over MSXML2 | Better TLS on corporate machines |
| ADODB.Stream for file write | Correct UTF-8 output; FSO writes ANSI |

### Critical Pitfalls (from research)

- Nominatim rate limit: enforce 1 req/sec via Application.Wait; check HTTP status
- User-Agent required: `GeoViz/1.0 (pomroyanalytics@gmail.com)` on every request
- Mark of the Web: document Unblock step for distributed workbook
- JSON injection: apply JsonEscape() to every string field before embedding
- Cache overwrite: skip rows where lat/lon is already populated
- OneDrive path: `ThisWorkbook.Path` returns `https://d.docs.live.net/...` URL when synced via OneDrive — use `LocalWorkbookPath()` in mod_Macros (resolves via `Environ("OneDriveConsumer")`) for any file write operations

### Open Questions

- Tile loading behavior on target machine (corporate proxy unknown)
- ~~Whether ThisWorkbook.Path resolves to a OneDrive path~~ — **Resolved:** it does return a URL; fixed via `LocalWorkbookPath()` using `Environ("OneDriveConsumer")`
- MOTW distribution scenario (email/download vs. local copy)
- Inline Leaflet JS/CSS vs. CDN reliance

### Todos

- None yet

### Blockers

- None

---

## Session Continuity

*Last updated: 2026-07-19 — Phase 2 complete; map generates and opens in browser*
