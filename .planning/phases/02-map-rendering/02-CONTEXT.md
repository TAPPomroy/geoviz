# Phase 2: Map Rendering - Context

**Gathered:** 2026-07-18
**Status:** Ready for planning

<domain>
## Phase Boundary

Phase 2 implements `GenerateMap()` in VBA: it reads the company table (via `BuildCompanyJson()`), builds a standalone HTML file embedding that JSON, a Leaflet.js map with ISP-colored CircleMarkers, MarkerCluster clustering, a CartoDB Positron basemap, and an ISP color legend — then writes the file to `ThisWorkbook.Path\geoviz_map.html` and opens it in the default browser via `Shell "explorer.exe"`.

No click interactions, radius circle, slider, or attribute filtering — those are Phase 3. Phase 2 ends when the HTML opens in the browser and all companies appear as correctly colored, clustered markers with a legend.

</domain>

<decisions>
## Implementation Decisions

### HTML File Output
- **D-01:** Write `geoviz_map.html` to `ThisWorkbook.Path` (same folder as the .xlsm). Simple and predictable; OneDrive sync is not a concern since the file is opened immediately after write.
- **D-02:** Open the file with `Shell "explorer.exe """ & filePath & """"`. Handles spaces in paths when properly quoted; uses the OS default browser handler.

### ISP Color Palette
- **D-03:** Use a standard qualitative palette (ColorBrewer Set1 or equivalent) — 7 high-contrast, colorblind-friendly colors suitable for CartoDB Positron's light gray background. Planner picks specific hex values from Set1 or a comparable 7-color qualitative set.
- **D-04:** ISPs beyond 7 (alphabetically) use `#999999` neutral gray as the "Other" color. Visually recedes behind the 7 named ISP colors without disappearing.
- **D-05:** ISP-to-color assignment is deterministic: sort all unique ISP values alphabetically, assign colors by index (1→color1, 2→color2, ..., 8+→#999999). Consistent across map regenerations.

### Marker Visual Style
- **D-06:** Use Leaflet `CircleMarker` for all company markers. No image assets needed; color is set via `fillColor` property.
- **D-07:** Marker style: `radius: 8, fillOpacity: 0.85, weight: 1, color: "#ffffff"` (white border). Visible at county/state zoom levels; white border separates adjacent same-ISP markers.
- **D-08:** Legend position: Leaflet bottom-right control. Lists each ISP with a color swatch and its name; "Other" gray shown if any ISPs overflow.

### VBA Module Structure
- **D-09:** Add a new `mod_MapBuilder` module for Phase 2 HTML generation. Mirrors Phase 1 pattern (one module per concern). `mod_Macros.GenerateMap()` calls `mod_MapBuilder.BuildMapHtml()` and handles file write + Shell open.
- **D-10:** HTML is built via string concatenation with a private `AppendLine(sb, line)` helper Sub inside `mod_MapBuilder`. The helper appends `line & vbLf` to the `sb` string variable — idiomatic VBA, no external dependencies.
- **D-11:** The company JSON array is embedded directly into the HTML as a JavaScript variable (`var companies = [...]`), inlined by calling `BuildCompanyJson()` from Phase 1. No separate data file.

### Claude's Discretion
- Specific ColorBrewer Set1 hex values (or equivalent 7-color qualitative palette) — planner picks values that render well on CartoDB Positron.
- Whether `mod_MapBuilder` exports `BuildMapHtml()` as a Function returning a String, or writes the file internally — planner decides based on testability preference.
- MarkerCluster CDN URL and version to use.

</decisions>

<canonical_refs>
## Canonical References

**Downstream agents MUST read these before planning or implementing.**

### Requirements
- `.planning/REQUIREMENTS.md` — MAP-01 through MAP-05 are the Phase 2 requirements; read Phase 3 requirements for integration awareness (INT-01, FILT-01, FILT-02 affect what Phase 2 must leave room for)
- `.planning/ROADMAP.md` — Phase 2 success criteria (4 items); Phase 1 success criteria shows what BuildCompanyJson() already delivers

### Architecture & Stack
- `CLAUDE.md` — Critical constraints: ADODB.Stream for file output (never FSO), WinHttp for HTTP, CartoDB Positron basemap, 7-color ISP palette rule, file:// CORS note for tile loading
- `.planning/phases/01-data-pipeline/01-CONTEXT.md` — Phase 1 decisions; especially D-07 through D-09 defining the JSON shape that mod_MapBuilder must consume

### Existing Code (Phase 1)
- `src/mod_JsonBuilder.bas` — `BuildCompanyJson(data As Collection)` returns the JSON array string; mod_MapBuilder calls this
- `src/mod_DataReader.bas` — `ReadCompanyData()` returns the Collection; GenerateMap() needs to call this before BuildCompanyJson()
- `src/mod_Macros.bas` — `GenerateMap()` stub is here; Phase 2 implements it
- `src/mod_Config.bas` — shared constants (HDR_LAT, HDR_LON, GEOCODE_FAILED, etc.) that mod_MapBuilder may reference

### Leaflet Resources
- Leaflet.js 1.9.x CDN (https://unpkg.com/leaflet@1.9.x) — per CLAUDE.md stack
- Leaflet.MarkerCluster plugin — required for MAP-03 (cluster markers in dense areas)
- CartoDB Positron tile URL — `https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png`

</canonical_refs>

<code_context>
## Existing Code Insights

### Reusable Assets
- `mod_JsonBuilder.BuildCompanyJson(data As Collection) As String` — returns JSON array of all successfully-geocoded companies with all columns. Phase 2 embeds this directly into the HTML.
- `mod_DataReader.ReadCompanyData() As Collection` — returns all company rows. GenerateMap() calls this first, passes result to BuildCompanyJson().
- `mod_Config.bas` constants — HDR_LAT, HDR_LON, GEOCODE_FAILED, TABLE_NAME, SHEET_NAME. mod_MapBuilder should import/reference rather than redefine.

### Established Patterns
- Module-per-concern: mod_DataReader (read), mod_Geocoder (geocode), mod_JsonBuilder (serialize), mod_Macros (entry points). mod_MapBuilder follows this as the HTML generation concern.
- Entry point in mod_Macros calls into specialist modules — `GenerateMap()` calls `mod_DataReader.ReadCompanyData()`, `mod_JsonBuilder.BuildCompanyJson()`, `mod_MapBuilder.*`.
- ADODB.Stream with `Charset = "utf-8"` for all file writes — MUST be used for HTML output (not FSO).
- `Application.StatusBar` for progress feedback during long operations.

### Integration Points
- `GenerateMap()` in mod_Macros is the public entry point — Phase 2 implements its body.
- The JSON string from `BuildCompanyJson()` is the seam between Phase 1 and Phase 2 — Phase 2 embeds it as `var companies = <json>;` in the HTML script block.
- Phase 3 will add click handlers and radius logic to the same HTML — Phase 2 should structure the JS in the HTML to make marker references accessible (e.g., store markers in an array `var markers = []`).

</code_context>

<specifics>
## Specific Ideas

- Phase 3 will need to iterate over markers to apply radius highlighting. Phase 2 should build the markers array (`var markers = []`) alongside the Leaflet layer so Phase 3 can reference it without restructuring the HTML.
- The legend should show ISP name + color swatch for each named ISP, plus an "Other" row if any companies use the overflow gray. Legend entries don't need to be clickable in Phase 2 (that's FILT-02 in Phase 3).
- CartoDB Positron attribution is required by their tile license — include it in the Leaflet map attribution string.

</specifics>

<deferred>
## Deferred Ideas

- Inline Leaflet JS/CSS for offline use — v2 requirement (OFF-V2-01); Phase 2 uses CDN only.
- Corporate proxy / tile loading failure handling — noted as an open question in STATE.md; Phase 2 accepts CDN dependency, document as a known limitation for users behind strict proxies.
- Legend entries clickable to show/hide ISP markers — Phase 3 (FILT-02).
- Attribute filter dropdown to switch coloring between columns — Phase 3 (FILT-01).
- `MsgBox` or status bar confirmation after file write — Claude's call (low value given the browser opens immediately).

</deferred>

---

*Phase: 2-Map Rendering*
*Context gathered: 2026-07-18*
