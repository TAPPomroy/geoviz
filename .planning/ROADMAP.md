# Roadmap: GeoViz

**Defined:** 2026-07-18
**Granularity:** Standard
**Coverage:** 15/15 v1 requirements mapped

---

## Phases

- [x] **Phase 1: Data Pipeline** - VBA reads the company sheet, geocodes addresses via Nominatim, and caches results
- [x] **Phase 2: Map Rendering** - HTML builder produces a Leaflet map with colored markers, clustering, and a macro button
- [x] **Phase 3: Interactivity** - Click interactions, radius highlighting, attribute filtering, and ISP legend controls (completed 2026-07-19)

---

## Phase Details

### Phase 1: Data Pipeline
**Goal:** The macro reliably reads the company table, geocodes uncached addresses at 1 req/sec, and writes lat/lon back to the sheet so subsequent runs skip already-geocoded rows.
**Depends on:** Nothing
**Requirements:** DATA-01, DATA-02, DATA-03, DATA-04
**Success Criteria:**
1. Running the macro on a fresh sheet populates lat/lon and a timestamp for every valid address row within the Nominatim rate limit (no HTTP 429 errors).
2. Re-running the macro on a fully-geocoded sheet completes instantly — zero new API calls — because cached rows are skipped.
3. Rows with unresolvable addresses receive the "GEOCODE_FAILED" sentinel in the lat/lon cells and are not retried on subsequent runs.
4. The macro produces a valid JSON array of company objects (verifiable by pasting into a browser console) with no character-encoding corruption in names or addresses.
**Plans:** `.planning/phases/01-data-pipeline/PLAN.md` (6 tasks)
**UI hint:** no

### Phase 2: Map Rendering
**Goal:** A single macro button generates a standalone HTML file and opens it in the browser, showing all companies as ISP-colored, clustered markers on a muted basemap with a legend.
**Depends on:** Phase 1
**Requirements:** MAP-01, MAP-02, MAP-03, MAP-04, MAP-05
**Success Criteria:**
1. Clicking the Excel macro button generates `geoviz_map.html` and opens it in the default browser without any manual steps.
2. Every company appears as a marker colored consistently by ISP (up to 7 named colors; additional ISPs show neutral gray), and the legend correctly labels each color.
3. Dense clusters of markers collapse into count bubbles that expand on click, preventing marker overlap.
4. The map renders on the CartoDB Positron basemap and is fully pannable and zoomable without page reloads.
**Plans:** 2 plans
Plans:
- [x] 02-01-PLAN.md — Create mod_MapBuilder.bas (HTML builder: BuildMapHtml + WriteUtf8File)
- [x] 02-02-PLAN.md — Implement GenerateMap() in mod_Macros + add Excel button + human verify
**UI hint:** yes

### Phase 3: Interactivity
**Goal:** Clicking a company on the map triggers a popup, a live radius circle, and neighbor highlighting; a slider adjusts the radius; attribute filter controls and a legend toggle let users focus on specific groups; a search control navigates to any company by name; and name labels appear automatically when zoomed in.
**Depends on:** Phase 2
**Requirements:** INT-01, INT-02, INT-03, INT-04, FILT-01, FILT-02, INT-05, INT-06
**Success Criteria:**
1. Clicking any marker opens a popup listing all company fields except Lat, Lon, GeocodedAt, and Column1; no VBA changes are needed to support new attribute columns.
2. After clicking a marker, a radius circle appears immediately and all companies within the radius are highlighted; companies sharing the clicked company's current color-by value are visually distinct from other neighbors.
3. Dragging the radius slider updates the circle boundary and neighbor highlights in real time without a page reload.
4. Using the attribute dropdown switches marker coloring to the selected column, and clicking a legend entry shows or hides matching markers immediately.
5. Selecting a company from the "Find company" dropdown flies the map to that company at zoom 15 (~2 mile view) and activates its full selection state.
6. When the visible map width is 5 miles or less, all company name labels appear to the right of their markers in the current field color; labels remain visible at full opacity when a company is selected.
**Plans:** 2/2 plans complete
Plans:
- [x] 03-01-PLAN.md — Marker click interaction, radius circle, neighbor highlighting, radius slider (INT-01 through INT-04)
- [x] 03-02-PLAN.md — Interactive legend with toggle, attribute filter dropdown, dynamic recoloring (FILT-01, FILT-02)
- [x] Post-plan — Find Company search control (INT-05), company name labels (INT-06), UX fixes
**UI hint:** yes

---

## Progress

| Phase | Plans Complete | Status | Completed |
|-------|----------------|--------|-----------|
| 1. Data Pipeline | 6/6 planned | Complete | 2026-07-18 |
| 2. Map Rendering | 2/2 planned | Complete | 2026-07-19 |
| 3. Interactivity | 2/2 | Complete   | 2026-07-19 |
