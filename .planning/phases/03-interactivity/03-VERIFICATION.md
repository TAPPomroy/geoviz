---
phase: 03-interactivity
verified: 2026-07-19T00:00:00Z
status: human_needed
score: 16/16 must-haves verified
overrides_applied: 0
human_verification:
  - test: "Run GenerateMap macro, open geoviz_map.html in browser. Click any company marker."
    expected: "Popup appears with all company fields except Lat and Lon. Blue semi-transparent circle draws around the marker. Non-neighbor markers fade to low opacity. Same-field neighbors show thicker border. Top-left slider shows 'Radius: 0.5 mi'."
    why_human: "Requires running Excel macro and visual browser inspection — cannot be verified by static code analysis."
  - test: "With a marker selected, drag the radius slider from 0.5 to 3."
    expected: "Circle expands live. Neighbor highlighting updates to include/exclude markers as they enter/leave the new radius. Label text changes to match slider value."
    why_human: "Live DOM event behavior requires browser interaction."
  - test: "Click the map background (not on a marker or control)."
    expected: "Circle disappears. All markers restore to default style and opacity. Popup closes. Hint text reappears."
    why_human: "Requires observing map background click propagation behavior in a live browser."
  - test: "Click the same marker a second time while it is selected."
    expected: "Selection clears (toggle deselect). Circle and popup disappear."
    why_human: "Requires live browser interaction."
  - test: "Open the 'Color by' dropdown (top-right). Verify it lists all columns except Lat, Lon, GeocodedAt. Select a non-ISP column."
    expected: "All markers recolor using the new column's values. Legend header changes to the selected column name. Legend swatches reflect the new column's unique values."
    why_human: "Requires a real dataset and browser rendering to confirm dropdown options and color output."
  - test: "Click a legend entry to hide its markers. Then click within the radius that includes some hidden markers."
    expected: "Hidden markers do not appear and are not highlighted as neighbors. Cluster counts decrease visually when markers are hidden."
    why_human: "Requires observing cluster layer behavior and hiddenValues interaction in a running browser."
  - test: "Hide some legend entries, then switch the 'Color by' column."
    expected: "Previously hidden markers reappear. Legend entries all restore to full opacity. No strikethrough on any entry."
    why_human: "Requires browser interaction to verify visibility reset across filter switch."
---

# Phase 3: Interactivity Verification Report

**Phase Goal:** Add marker click interactivity, proximity radius highlighting, and attribute filter — the core use case of clicking a company and seeing nearby ISP neighbors.
**Verified:** 2026-07-19
**Status:** human_needed
**Re-verification:** No — initial verification

## Goal Achievement

### Observable Truths

| # | Truth | Status | Evidence |
|---|-------|--------|----------|
| 1 | Clicking a marker opens a popup table of all fields except Lat and Lon | VERIFIED | `buildPopupHtml` uses `POPUP_EXCLUDE = ['Lat','Lon']`, renders HTML table rows with `escHtml`; `applySelection` calls `L.popup().setContent(buildPopupHtml(company)).openOn(map)` — line 200 |
| 2 | A blue semi-transparent circle appears centered on the clicked marker, sized by the slider value | VERIFIED | `radiusCircle = L.circle(marker.getLatLng(), { radius: radiusMeters, color: '#0078ff', weight: 1, fillOpacity: 0.05, interactive: false }).addTo(map)` — line 205; `radiusMeters = sliderVal * 1609.34` |
| 3 | Non-neighbor markers fade to fillOpacity 0.15; neighbor markers remain at 0.85 | VERIFIED | Non-neighbor path: `m.setStyle({ fillOpacity: 0.15 })` — line 215; neighbor path: `fillOpacity: 0.85` — line 212 |
| 4 | Same-ISP neighbors show white border weight 3; other-ISP neighbors show weight 1 | VERIFIED | `weight: String(companies[i][currentField]) === String(company[currentField]) ? 3 : 1` — line 212; `color: '#ffffff'` for both |
| 5 | The selected marker expands to radius 12, white fill, ISP color border, weight 2 | VERIFIED | `marker.setRadius(12)` — line 201; `marker.setStyle({ fillColor: '#ffffff', color: getFieldColor(company[currentField]), weight: 2, fillOpacity: 0.85 })` — line 202 |
| 6 | Dragging the radius slider updates the label text and redraws the circle + highlights live | VERIFIED | Slider `input` listener updates `radiusLabel.textContent` and calls `refreshRadius()` — lines 349-352; `refreshRadius()` redraws circle and iterates markers — lines 322-345 |
| 7 | Clicking the map background resets all markers, removes the circle, and closes the popup | VERIFIED | `map.on('click', clearSelection)` — line 365; `clearSelection` removes `radiusCircle`, resets all marker styles via `forEach`, calls `map.closePopup()` — lines 178-189 |
| 8 | Clicking the currently-selected marker again clears selection (toggle deselect) | VERIFIED | `applySelection` first check: `if (selectedMarker === marker) { clearSelection(); return; }` — line 194 |
| 9 | The filter dropdown lists all column names except Lat, Lon, and GeocodedAt | VERIFIED | `SYSTEM_FIELDS = ['Lat','Lon','GeocodedAt']` checked in `filterControl.onAdd` — lines 301-306; options built from `Object.keys(companies[0])` excluding these |
| 10 | Changing the dropdown re-colors all visible markers using the selected column's values | VERIFIED | `applyFilter` calls `m.setStyle({ fillColor: getFieldColor(v) })` for all markers — lines 268-270 |
| 11 | The legend header updates to show the selected column name | VERIFIED | `buildLegend` sets `html = '<strong>' + currentField + '</strong><br>'` — line 224; `applyFilter` sets `currentField = field` then calls `buildLegend()` |
| 12 | The legend swatches and labels reflect the selected column's unique values (same alphabetical-sort + ISP_COLORS algorithm) | VERIFIED | `buildColorMap(field)` collects unique values, sorts (`fieldNames.sort()`), assigns `ISP_COLORS` by index — lines 159-170; `buildLegend` renders `fieldNames` with `fieldColorMap` colors |
| 13 | Clicking a legend entry hides all markers matching that value by removing them from clusterGroup | VERIFIED | `toggleValue` hide branch: `clusterGroup.removeLayer(m)` — lines 250-254; wired via event delegation on `legendDiv` using `.legend-item` and `data-val` attribute — lines 98-101 |
| 14 | Clicking a hidden legend entry restores those markers by adding them back to clusterGroup | VERIFIED | `toggleValue` show branch: `clusterGroup.addLayer(m)` — lines 243-246 |
| 15 | Hidden markers remain hidden even when inside the radius circle (hiddenValues takes precedence) | VERIFIED | Both `applySelection` and `refreshRadius` check `if (hiddenValues[String(companies[i][currentField])]) return;` before highlighting — lines 208, 332 |
| 16 | Switching the filter column resets all hidden values and restores all previously-hidden markers | VERIFIED | `applyFilter` sets `hiddenValues = {}` and calls `clusterGroup.addLayer(m)` for every marker — lines 265, 270 |

**Score:** 16/16 truths verified

### Required Artifacts

| Artifact | Expected | Status | Details |
|----------|----------|--------|---------|
| `src/mod_MapBuilder.bas` | BuildMapHtml() with interactivity JS block containing `function haversine` | VERIFIED | `function haversine` present at line 116; uses `var R = 3958.8` |
| `src/mod_MapBuilder.bas` | radius slider Leaflet control (`radiusControl`) | VERIFIED | `var radiusControl = L.control({position:'topleft'})` at line 282 |
| `src/mod_MapBuilder.bas` | `clearSelection` and `applySelection` functions | VERIFIED | Both defined at lines 178 and 193 |
| `src/mod_MapBuilder.bas` | `buildColorMap`, `getFieldColor`, `buildLegend`, `toggleValue` | VERIFIED | All four defined at lines 159, 172, 222, 240 |
| `src/mod_MapBuilder.bas` | `applyFilter` function | VERIFIED | Defined at line 262 |
| `src/mod_MapBuilder.bas` | interactive legend replacing static legend (`var legendDiv`) | VERIFIED | `var legendDiv = null` at line 87; `legend.onAdd` stores `legendDiv = div` at line 93 without setting innerHTML |

### Key Link Verification

| From | To | Via | Status | Details |
|------|----|-----|--------|---------|
| `markers[i]` | `companies[i]` | parallel array — forEach with index `i` | VERIFIED | `markers.forEach(function(m, i) { ... companies[i] ... }` throughout applySelection, refreshRadius, toggleValue |
| `L.circle` | slider value | `sliderVal * 1609.34` (miles to meters) | VERIFIED | `var radiusMeters = sliderVal * 1609.34` in applySelection (line 204) and refreshRadius (line 327) |
| `marker.setRadius(12)` | `marker.setStyle` | called separately — not `setStyle({radius:12})` | VERIFIED | Lines 201-202 show separate calls |
| `buildLegend` | `legendDiv` | `legendDiv.innerHTML = html` | VERIFIED | Line 237: `legendDiv.innerHTML = html`; `legendDiv` captured in `legend.onAdd` |
| `toggleValue` | `clusterGroup` | `clusterGroup.removeLayer / addLayer` | VERIFIED | Lines 245-246 (addLayer) and 251-252 (removeLayer) |
| `applyFilter` | `clearSelection` | `clearSelection()` as first statement | VERIFIED | Line 263: `clearSelection();` is first line of applyFilter |
| `fieldColorMap` | `clearSelection` | `getFieldColor(companies[i][currentField])` | VERIFIED | Line 182 in clearSelection: `fillColor: getFieldColor(companies[i][currentField])` |
| `legendDiv` click | `toggleValue` | event delegation via `.legend-item` `.data-val` | VERIFIED | Lines 98-101: `legendDiv.addEventListener('click', ...)` with `e.target.closest('.legend-item')` and `el.dataset.val` |

### Data-Flow Trace (Level 4)

| Artifact | Data Variable | Source | Produces Real Data | Status |
|----------|---------------|--------|--------------------|--------|
| `buildPopupHtml` | `company` object | `companies[]` array injected from `jsonStr` VBA parameter | Yes — JSON from Excel sheet | FLOWING |
| `buildLegend` | `fieldNames`, `fieldColorMap` | `buildColorMap(field)` reading `companies[]` | Yes — derived from real company data | FLOWING |
| `applyFilter` | `currentField`, `fieldColorMap` | `buildColorMap(field)` then marker forEach | Yes — real data per company | FLOWING |
| `filterControl` dropdown options | `filterFields` | `Object.keys(companies[0])` excluding SYSTEM_FIELDS | Yes — column names from actual data | FLOWING |

### Behavioral Spot-Checks

Step 7b: SKIPPED — the deliverable is VBA source code (`mod_MapBuilder.bas`) that generates an HTML file at runtime via the Excel macro. No runnable entry point exists outside of Excel. Browser-based behaviors are routed to human verification.

### Probe Execution

Step 7c: SKIPPED — no probe scripts declared in PLANs and no `scripts/*/tests/probe-*.sh` found for this project.

### Requirements Coverage

| Requirement | Source Plan | Description | Status | Evidence |
|-------------|-------------|-------------|--------|----------|
| INT-01 | 03-01-PLAN.md | Clicking a company marker opens a popup showing all attribute columns | SATISFIED | `buildPopupHtml` renders all keys except `['Lat','Lon']`; wired via `applySelection` |
| INT-02 | 03-01-PLAN.md | Clicking draws a radius circle and highlights companies within the radius | SATISFIED | `L.circle` drawn in `applySelection`; neighbor loop with haversine distance check |
| INT-03 | 03-01-PLAN.md | Same-ISP neighbors are visually distinguished | SATISFIED | `weight: 3` for same-field-value neighbors vs `weight: 1` for others; now field-agnostic via `currentField` (correct generalization) |
| INT-04 | 03-01-PLAN.md | Radius slider lets user adjust proximity; circle and highlights update live | SATISFIED | `radiusControl` with range input (min 0.1, max 5, step 0.1); `refreshRadius()` called on `input` event |
| FILT-01 | 03-02-PLAN.md | Dropdown lets user switch marker coloring between attribute columns | SATISFIED | `filterControl` dropdown; `applyFilter` recolors all markers via `getFieldColor(companies[i][field])` |
| FILT-02 | 03-02-PLAN.md | Legend allows individual ISPs to be shown or hidden | SATISFIED | `toggleValue` uses `clusterGroup.removeLayer/addLayer`; legend items use event delegation with `data-val` |

### Anti-Patterns Found

| File | Line | Pattern | Severity | Impact |
|------|------|---------|----------|--------|
| — | — | No TODO/FIXME/XXX/TBD markers found | — | — |
| — | — | No stub returns (return null, return [], return {}) in rendering paths | — | — |

No anti-patterns found. All functions produce substantive output from real data.

### Notable Deviations (Non-Blocking)

**Legend event wiring changed from inline onclick to event delegation:**
The plan specified `onclick="toggleValue('...')"` inline in each legend item's HTML. The implementation uses `legendDiv.addEventListener('click', ...)` with `.closest('.legend-item')` and `data-val` attribute instead. This is a strictly better approach (avoids XSS via inline onclick and handles single-quote escaping cleanly). The SUMMARY comments reference this as WR-01/WR-02. Single-quote escaping is now handled via `escHtml(v)` on the `data-val` attribute. Functional outcome is identical.

**Slider calls `refreshRadius()` instead of `applySelection(selectedMarker, selectedCompany)`:**
The plan specified re-calling `applySelection` on slider drag. The implementation defines a separate `refreshRadius()` function that skips the toggle-deselect check (CR-01 in code comment). This prevents the slider from accidentally clearing selection. The behavior is correct and more robust than the plan spec.

### Human Verification Required

All automated code checks pass. The following behaviors require human testing in a running browser:

### 1. Marker Click — Popup and Visual Selection

**Test:** Run the GenerateMap macro. Open geoviz_map.html. Click any company marker.
**Expected:** Popup opens showing all company fields except Lat and Lon. Blue semi-transparent circle appears centered on the marker. Non-neighbor markers fade visually. Selected marker appears larger with white fill. Top-left slider panel shows "Radius: 0.5 mi".
**Why human:** Requires running Excel macro and observing live Leaflet rendering in a browser.

### 2. Slider Live Update

**Test:** With a marker selected, drag the radius slider from 0.5 to 3 mi.
**Expected:** Circle expands continuously. Neighbor set updates as markers enter/leave the radius. Label text tracks the slider value.
**Why human:** Live DOM event behavior and visual circle redraw require browser interaction.

### 3. Map Background Click — Clear Selection

**Test:** Click the map background (not on a marker or UI control).
**Expected:** Blue circle disappears. All markers restore to default style. Popup closes. "Click a company to see neighbors" hint text reappears.
**Why human:** Requires observing event propagation in a live browser.

### 4. Toggle Deselect

**Test:** Click the same marker a second time while it is selected.
**Expected:** Selection clears immediately — no new selection, circle removed, markers restore.
**Why human:** Requires live browser interaction.

### 5. Filter Dropdown — Recolor by Attribute

**Test:** Open the "Color by" dropdown (top-right). Verify the list contains all column names except Lat, Lon, GeocodedAt. Select a non-ISP column.
**Expected:** All markers recolor using the new column's palette. Legend header changes to the selected column name. Legend swatches and labels reflect that column's unique values in the same 7-color scheme.
**Why human:** Requires a real dataset loaded through the macro; requires browser rendering to confirm.

### 6. Legend Hide/Show with Radius Interaction

**Test:** Click a legend entry to hide its markers. Then click a company and observe neighbors in the radius.
**Expected:** Hidden markers are absent from the map (cluster counts decrease). Hidden markers are NOT highlighted as neighbors even if within the radius.
**Why human:** Requires observing cluster layer behavior and `hiddenValues` enforcement in a running browser.

### 7. Filter Column Switch Resets Hidden Markers

**Test:** Hide some legend entries, then select a different "Color by" column.
**Expected:** All previously hidden markers reappear. Legend shows no strikethrough entries. All markers recolored by the new column.
**Why human:** Requires browser interaction to confirm visibility reset and cluster group restoration.

---

_Verified: 2026-07-19_
_Verifier: Claude (gsd-verifier)_
