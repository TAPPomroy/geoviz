# Roadmap: GeoViz

**Defined:** 2026-07-18
**Granularity:** Standard
**Coverage:** 15/15 v1 requirements mapped

---

## Phases

- [ ] **Phase 1: Data Pipeline** - VBA reads the company sheet, geocodes addresses via Nominatim, and caches results
- [ ] **Phase 2: Map Rendering** - HTML builder produces a Leaflet map with colored markers, clustering, and a macro button
- [ ] **Phase 3: Interactivity** - Click interactions, radius highlighting, attribute filtering, and ISP legend controls

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
**Plans:** TBD
**UI hint:** yes

### Phase 3: Interactivity
**Goal:** Clicking a company on the map triggers a popup, a live radius circle, and neighbor highlighting; a slider adjusts the radius; and attribute filter controls and an ISP legend toggle let users focus on specific groups.
**Depends on:** Phase 2
**Requirements:** INT-01, INT-02, INT-03, INT-04, FILT-01, FILT-02
**Success Criteria:**
1. Clicking any marker opens a popup listing company name, ISP, address, and all attribute columns; no VBA changes are needed to support new attribute columns.
2. After clicking a marker, a radius circle appears immediately and all companies within the radius are highlighted; companies sharing the clicked company's ISP are visually distinct from other neighbors.
3. Dragging the radius slider updates the circle boundary and neighbor highlights in real time without a page reload.
4. Using the attribute dropdown switches marker coloring to the selected column, and clicking an ISP legend entry shows or hides that ISP's markers immediately.
**Plans:** TBD
**UI hint:** yes

---

## Progress

| Phase | Plans Complete | Status | Completed |
|-------|----------------|--------|-----------|
| 1. Data Pipeline | 6/6 planned | Ready to execute | - |
| 2. Map Rendering | 0/? | Not started | - |
| 3. Interactivity | 0/? | Not started | - |
