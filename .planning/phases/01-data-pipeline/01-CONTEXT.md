# Phase 1: Data Pipeline - Context

**Gathered:** 2026-07-18
**Status:** Ready for planning

<domain>
## Phase Boundary

VBA reads the company table from the Excel workbook, geocodes uncached addresses via Nominatim at 1 req/sec, writes lat/lon back to the sheet, and marks failures. No HTML generation, no browser launch — that is Phase 2. Phase 1 ends when the sheet has lat/lon populated for every resolvable address.

</domain>

<decisions>
## Implementation Decisions

### Sheet Structure
- **D-01:** Data must be in a named Excel Table (Insert → Table) named `CompanyData` on a worksheet named `GeoViz`. The macro locates the table by name — not by range address or row scan.
- **D-02:** Required column headers: `Name`, `Address`, `ISP` (case-insensitive match). Lat/Lon cache columns are also part of this table (headers TBD by planner — e.g., `Lat`, `Lon`, `GeocodedAt`).
- **D-03:** Any additional columns in the table are treated as user-defined attributes and passed through without modification.

### Progress Feedback
- **D-04:** During geocoding, update `Application.StatusBar` with each row: e.g., `"Geocoding 5/42 — Acme Corp..."`. Non-intrusive, visible without interrupting the user.
- **D-05:** On completion, rows where geocoding failed (`GEOCODE_FAILED`) are highlighted in red (interior color). No modal dialog unless a hard error occurs (e.g., HTTP auth failure).
- **D-06:** Status bar is cleared (`Application.StatusBar = False`) after the run.

### JSON Shape
- **D-07:** Each company object in the JSON array includes ALL columns from `CompanyData` dynamically — every column header becomes a JSON key. This is required for Phase 3's attribute filter (FILT-01) to work without code changes.
- **D-08:** Lat/Lon are included as numeric values (not strings). All other values are strings with JSON-safe escaping (`JsonEscape()` applied to every string field).
- **D-09:** JSON is built in memory (VBA String variable) and returned from a function — no intermediate file or hidden sheet cell. Phase 2's `GenerateMap` macro calls this function directly.

### Macro Entry Points
- **D-10:** Two separate public macros: `GeocodeSheet` (this phase) and `GenerateMap` (Phase 2). User can re-geocode without launching the browser, and regenerate the map without re-geocoding.
- **D-11:** `GeocodeSheet` is the only Phase 1 entry point. It handles DATA-01 through DATA-04: read table, geocode uncached rows, cache lat/lon, mark failures.

### Claude's Discretion
- Names for the Lat/Lon/timestamp cache columns (planner should pick something clear, e.g., `Lat`, `Lon`, `GeocodedAt`).
- Whether to use a single module or split into `modGeocoding` + `modData` — planner decides based on size.

</decisions>

<canonical_refs>
## Canonical References

**Downstream agents MUST read these before planning or implementing.**

### Requirements
- `.planning/REQUIREMENTS.md` — DATA-01 through DATA-04 are the Phase 1 requirements; read the full file for context on Phase 2/3 integration points
- `.planning/ROADMAP.md` — Phase 1 success criteria (4 items); the JSON array output is verified by pasting into a browser console

### Architecture & Stack
- `CLAUDE.md` — Critical constraints: ADODB.Stream for file output, WinHttp.WinHttpRequest.5.1 for HTTP, Nominatim User-Agent requirement, Mark of the Web note

### Key Pitfalls (from research — see STATE.md)
- Nominatim User-Agent: `GeoViz/1.0 (pomroyanalytics@gmail.com)` on every request
- Rate limit: `Application.Wait Now + TimeValue("0:00:01")` between each geocode call
- Cache guard: skip rows where `Lat` column is already non-empty (and not `GEOCODE_FAILED`)
- JSON injection: `JsonEscape()` must be applied to every string before embedding in JSON

</canonical_refs>

<code_context>
## Existing Code Insights

### Reusable Assets
- None — this is a greenfield workbook. No existing VBA modules.

### Established Patterns
- None yet — Phase 1 establishes the patterns (named table access, WinHttp wrapper, ADODB.Stream writer) that later phases will follow.

### Integration Points
- `GeocodeSheet` populates `Lat`, `Lon`, `GeocodedAt` columns in `CompanyData`. Phase 2's `GenerateMap` reads these same columns to build the JSON payload.
- The `BuildCompanyJson()` function (or equivalent) is the seam between Phase 1 and Phase 2 — it reads `CompanyData` and returns a JSON string.

</code_context>

<specifics>
## Specific Ideas

- The JSON array should be verifiable by copy-pasting into a browser console (`JSON.parse(...)` succeeds with no errors) — this is the Phase 1 acceptance test per ROADMAP.md.
- Failure highlight: use a distinct red fill (e.g., Excel color index for bright red) on the entire row or just the Lat/Lon cells — planner's call on scope of highlight.

</specifics>

<deferred>
## Deferred Ideas

- OneDrive path resolution for `ThisWorkbook.Path` — relevant to Phase 2 (where the HTML file is written to disk).
- Corporate proxy / tile loading behavior — Phase 2 concern (CDN tile loading).
- Inline Leaflet JS/CSS for offline use — v2 requirement (OFF-V2-01), not Phase 1.
- MsgBox summary of failed addresses — considered, user preferred row highlighting instead. Could be added later if desired.

</deferred>

---

*Phase: 1-Data Pipeline*
*Context gathered: 2026-07-18*
