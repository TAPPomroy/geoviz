---
phase: 03-interactivity
plan: "01"
subsystem: mod_MapBuilder
tags: [interactivity, leaflet, javascript, vba, haversine]
dependency_graph:
  requires: [02-01-SUMMARY.md]
  provides: [marker-click-interactivity, radius-slider, neighbor-highlighting]
  affects: [src/mod_MapBuilder.bas]
tech_stack:
  added: []
  patterns: [haversine-distance, leaflet-circleMarker-setStyle, L.control-topleft, DomEvent-stopPropagation]
key_files:
  created: []
  modified:
    - src/mod_MapBuilder.bas
decisions:
  - "Single editing pass for Task 1 and Task 2 JS blocks — plan explicitly permitted this for cleanliness"
  - "setHint() defined before clearSelection/applySelection so both can call it at parse time"
  - "hiddenValues guard in applySelection neighbor loop ensures legend-hidden markers are skipped (D-13)"
metrics:
  duration: "~10 minutes"
  completed: "2026-07-19"
---

# Phase 3 Plan 01: Marker Click Interactivity and Radius Slider Summary

**One-liner:** Haversine-based marker click with blue radius circle, neighbor highlighting, and a Leaflet top-left slider control — all wired via stopPropagation and disableClickPropagation guards.

## Tasks Completed

| Task | Name | Commit | Files |
|------|------|--------|-------|
| 1 | State vars, haversine, clearSelection, applySelection, buildPopupHtml | 715799a | src/mod_MapBuilder.bas |
| 2 | Radius slider Leaflet control and event bindings | 715799a | src/mod_MapBuilder.bas |

Both tasks were implemented in a single editing pass as permitted by the plan.

## What Was Built

Added ~90 lines of VBA AppendLine calls to `BuildMapHtml()` in `src/mod_MapBuilder.bas`, injected between the existing `legend.addTo(map);` line and the closing `</script></body></html>` tag.

**JavaScript added to the generated HTML:**

- `haversine(lat1, lon1, lat2, lon2)` — Earth radius 3958.8 miles, returns distance in miles
- `buildPopupHtml(company)` — renders an HTML table of all company fields excluding Lat and Lon
- `setHint(show)` — toggles the "Click a company to see neighbors" hint div
- `clearSelection()` — removes radius circle, resets all markers to default style (radius 8, fillOpacity 0.85), closes popup, nulls state vars, shows hint
- `applySelection(marker, company)` — toggle-deselect if same marker; otherwise clears previous, opens popup, styles selected marker (radius 12, white fill, ISP-colored border weight 2), draws blue L.circle at sliderVal * 1609.34 meters, iterates all markers to apply neighbor/non-neighbor styles
- `radiusControl` — L.control at topleft with label, range input (min 0.1, max 5, step 0.1, default 0.5), and hint div; disableClickPropagation and disableScrollPropagation applied
- Slider input listener — updates label text and calls applySelection if marker is selected
- Per-marker click handlers — stopPropagation then applySelection
- `map.on('click', clearSelection)` — background click resets all state

**Neighbor highlighting logic:**
- Same-ISP neighbor: fillColor from ispColorMap, color '#ffffff', weight 3, fillOpacity 0.85
- Other-ISP neighbor: fillColor from ispColorMap, color '#ffffff', weight 1, fillOpacity 0.85
- Non-neighbor: setStyle fillOpacity 0.15 (retains existing fillColor)
- hiddenValues[companies[i].ISP] guard skips legend-hidden markers

## Deviations from Plan

None — plan executed exactly as written. Tasks 1 and 2 were combined into one editing pass as the plan explicitly authorized.

## Self-Check

- [x] src/mod_MapBuilder.bas modified with all AppendLine blocks
- [x] Commit 715799a exists: `git log --oneline` confirms
- [x] haversine uses R = 3958.8
- [x] applySelection calls setRadius(12) and setStyle separately
- [x] radiusCircle = L.circle with radius: sliderVal * 1609.34
- [x] disableClickPropagation present in onAdd
- [x] hiddenValues guard in neighbor loop
- [x] buildPopupHtml excludes 'Lat' and 'Lon'

## Self-Check: PASSED
