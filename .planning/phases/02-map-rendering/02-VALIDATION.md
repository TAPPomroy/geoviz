# Phase 2: Map Rendering — Validation Strategy

**Phase:** 02-map-rendering
**Date:** 2026-07-18
**Validation mode:** Manual only (no automated VBA test framework in use)

---

## Wave 0: Automated Tests

**Wave 0: not applicable.** The GeoViz project has no automated test framework. All validation is manual, performed by the developer in the VBA IDE and browser. This is consistent with the project stack (pure VBA, Excel macro, standalone HTML output).

No automated test files to create or run.

---

## Validation Requirements

| REQ-ID | Requirement | Validation Method | Manual Steps |
|--------|-------------|-------------------|--------------|
| MAP-01 | Leaflet map renders, pannable and zoomable | Visual (browser) | Open `geoviz_map.html` in Chrome; drag to pan, scroll to zoom — map tiles load and viewport moves |
| MAP-02 | ISP-colored CircleMarkers, 7-color palette, legend | Visual (browser) | Inspect markers — each ISP gets a distinct color from #e41a1c/#377eb8/#4daf4a/#984ea3/#ff7f00/#a65628/#f781bf; legend lists each ISP with matching swatch |
| MAP-03 | Dense markers cluster into count bubbles | Visual (browser) | Zoom out to national level — overlapping markers collapse into numbered bubbles; click bubble to expand |
| MAP-04 | CartoDB Positron basemap | Visual (browser) | Gray/white tile basemap loads; street names visible; CartoDB attribution present in bottom-right |
| MAP-05 | Single button generates HTML and opens browser | Functional (Excel) | Click macro button on GeoViz sheet (or run GenerateMap from Macro dialog); `geoviz_map.html` appears in workbook folder; default browser opens automatically |

---

## Source Assumptions (from RESEARCH.md)

| ID | Assumption | Risk if Wrong |
|----|------------|---------------|
| A3 | `file://` HTML can load HTTPS CartoDB tiles in Chrome/Edge | MAP-04 fails — blank basemap |
| A4 | `Shell "explorer.exe ..."` invokes default browser | MAP-05 fails — file opens in wrong app |

Both assumptions must be verified during the MAP-05 manual test step. If A3 fails, document as a known limitation and advise users to open the file from a web server or use Edge's local file mode.

---

## Acceptance Threshold

All 5 requirements (MAP-01 through MAP-05) must pass manual verification before the phase is marked complete.

---

## Verification Results

**Date:** 2026-07-19  
**Result:** PASSED — all 5 requirements verified

| REQ-ID | Result | Notes |
|--------|--------|-------|
| MAP-01 | ✓ PASS | Map renders, pans, and zooms in browser |
| MAP-02 | ✓ PASS | ISP-colored markers and legend display correctly |
| MAP-03 | ✓ PASS | Dense markers cluster into count bubbles |
| MAP-04 | ✓ PASS | CartoDB Positron basemap loads via file:// in Chrome |
| MAP-05 | ✓ PASS | Button click generates `geoviz_map.html` and opens browser |

**Assumption outcomes:**
- A3 (`file://` → HTTPS tiles): **Confirmed working** in Chrome
- A4 (`Shell "explorer.exe"`): **Confirmed** opens default browser

**Deviations from plan:**
- `ThisWorkbook.Path` returned OneDrive URL (`https://d.docs.live.net/...`) — not anticipated; resolved by adding `LocalWorkbookPath()` to mod_Macros using `Environ("OneDriveConsumer")`
- ADODB.Stream required `Position = 0` before `SaveToFile` — added to WriteUtf8File

---

*Validation strategy: 2026-07-18*
*Verified: 2026-07-19*
*Phase: 2 — Map Rendering*
