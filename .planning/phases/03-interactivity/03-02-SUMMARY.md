---
phase: 03-interactivity
plan: "02"
subsystem: mod_MapBuilder
tags: [interactivity, filter, legend, leaflet, javascript, vba]
dependency_graph:
  requires: [03-01-SUMMARY.md]
  provides: [dynamic-legend, attribute-filter-dropdown, buildColorMap, getFieldColor, buildLegend, toggleValue, applyFilter]
  affects: [src/mod_MapBuilder.bas]
tech_stack:
  added: []
  patterns: [L.control-topright, clusterGroup-removeLayer-addLayer, fieldColorMap-palette, onclick-single-quote-escape]
key_files:
  created: []
  modified:
    - src/mod_MapBuilder.bas
decisions:
  - "Single editing pass for both tasks — plan permitted combined pass; all JS written in correct dependency order"
  - "applySelection updated to use getFieldColor(company[currentField]) for selected marker border color (not ISP-specific)"
  - "hiddenValues check in applySelection updated from companies[i].ISP to companies[i][currentField] (correctness fix)"
  - "Same-value neighbor highlight weight now compares String(companies[i][currentField]) === String(company[currentField]) instead of ISP-specific comparison"
metrics:
  duration: "~8 minutes"
  completed: "2026-07-19"
---

# Phase 3 Plan 02: Dynamic Legend and Attribute Filter Dropdown Summary

**One-liner:** Interactive legend with per-entry click-to-hide (clusterGroup.removeLayer/addLayer) and a topright "Color by" dropdown recoloring all markers via buildColorMap/getFieldColor — FILT-01 and FILT-02 complete.

## Tasks Completed

| Task | Name | Commit | Files |
|------|------|--------|-------|
| 1 | buildColorMap, getFieldColor, buildLegend, toggleValue — dynamic legend | fd3c783 | src/mod_MapBuilder.bas |
| 2 | applyFilter function and filter dropdown Leaflet control | fd3c783 | src/mod_MapBuilder.bas |

Both tasks were implemented in a single editing pass (plan authorized combined pass).

## What Was Built

Rewrote `BuildMapHtml()` in `src/mod_MapBuilder.bas` with the following changes:

**Legend converted from static to dynamic:**
- `legend.onAdd` now only creates the div, applies CSS, stores `legendDiv = div`, and returns div — no innerHTML set
- `buildLegend()` called after `legend.addTo(map)` (and after `buildColorMap('ISP')`) to populate initial content
- `var legendDiv = null` declared alongside other state vars

**New JS functions added to generated HTML:**

- `buildColorMap(field)` — collects unique values for the given field from `companies[]`, sorts alphabetically, assigns `ISP_COLORS` palette (same algorithm as the Phase 2 `ispColorMap`), populates `fieldColorMap` and `fieldNames`
- `getFieldColor(value)` — returns `fieldColorMap[String(value)] || OTHER_COLOR`
- `buildLegend()` — rebuilds `legendDiv.innerHTML` with `currentField` as header; each entry has a colored swatch and a clickable span (strikethrough + 0.4 opacity when hidden); single-quotes in field values are escaped with `.replace(/'/g, "\\'")` (T-03-02A mitigation)
- `toggleValue(val)` — toggles `hiddenValues[val]`; uses `clusterGroup.removeLayer(m)` to hide and `clusterGroup.addLayer(m)` to restore; calls `clearSelection()` when hiding the currently-selected marker; calls `buildLegend()` to update visual state
- `applyFilter(field)` — calls `clearSelection()` first, sets `currentField = field`, resets `hiddenValues = {}`, calls `buildColorMap(field)`, recolors all markers via `m.setStyle({ fillColor: getFieldColor(v) })`, calls `clusterGroup.addLayer(m)` for all markers (restore all to visible, D-13), calls `buildLegend()`
- `filterControl` — `L.control({position:'topright'})` with a "Color by" label and `<select id="fieldSelect">` dropdown; options built from `Object.keys(companies[0])` excluding `SYSTEM_FIELDS = ['Lat','Lon','GeocodedAt']`; ISP option carries `selected` attribute as default; `disableClickPropagation` and `disableScrollPropagation` applied
- `fieldSelect` change event calls `applyFilter(this.value)`
- Initialization: `buildColorMap('ISP'); buildLegend();` called after all function definitions

**Updated existing functions:**
- `clearSelection()` — `fillColor` restore line changed from `ispColorMap[companies[i].ISP]` to `getFieldColor(companies[i][currentField])`
- `applySelection()` — selected marker border color changed from `ispColorMap[company.ISP]` to `getFieldColor(company[currentField])`; neighbor `fillColor` changed to `getFieldColor(companies[i][currentField])`; same-value neighbor weight comparison changed from `companies[i].ISP === company.ISP` to `String(companies[i][currentField]) === String(company[currentField])`; `hiddenValues` check changed from `companies[i].ISP` to `String(companies[i][currentField])`

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 1 - Bug] Updated applySelection to use getFieldColor and currentField throughout**
- **Found during:** Task 1 (single-pass rewrite)
- **Issue:** Plan 02 only explicitly called out updating `clearSelection`, but `applySelection` also used `ispColorMap[company.ISP]` for the selected marker border color, `ispColorMap[companies[i].ISP]` for neighbor fill colors, `companies[i].ISP === company.ISP` for same-ISP neighbor weight, and `hiddenValues[companies[i].ISP]` for the hidden-values guard. After Plan 02 introduces `currentField`, all these would show stale ISP colors when a different filter column is active.
- **Fix:** Changed all four usages in `applySelection` to use `getFieldColor(...)` and `companies[i][currentField]` as described above.
- **Files modified:** src/mod_MapBuilder.bas
- **Commit:** fd3c783

## Known Stubs

None — all filter and legend functionality is fully wired to the `companies[]` data.

## Threat Flags

None beyond what is documented in the plan's threat model (T-03-02A mitigated via single-quote escape; T-03-02B accepted).

## Self-Check

- [x] src/mod_MapBuilder.bas modified with all AppendLine blocks
- [x] Commit fd3c783 exists
- [x] `fieldNames.sort()` present in buildColorMap
- [x] `legendDiv = div` in legend.onAdd (not div.innerHTML)
- [x] `buildLegend()` called after `legend.addTo(map)`
- [x] `clusterGroup.removeLayer(m)` in toggleValue hide branch
- [x] `clusterGroup.addLayer(m)` in toggleValue restore branch and in applyFilter
- [x] `clearSelection()` as first statement in applyFilter
- [x] `hiddenValues = {}` in applyFilter
- [x] `SYSTEM_FIELDS = ['Lat','Lon','GeocodedAt']` in filterControl.onAdd
- [x] `disableClickPropagation` called on filterControl div
- [x] `buildColorMap('ISP'); buildLegend();` initialization calls present
- [x] Single-quote escape `.replace(/'/g, "\\'")` in buildLegend onclick attribute

## Self-Check: PASSED
