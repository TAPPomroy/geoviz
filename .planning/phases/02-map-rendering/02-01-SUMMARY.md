---
plan: "02-01"
phase: "02-map-rendering"
status: complete
self_check: PASSED
---

# Plan 02-01 Summary: Create mod_MapBuilder.bas

## What Was Built

`src/mod_MapBuilder.bas` — the HTML generation module for Phase 2.

Two public exports:
- `BuildMapHtml(jsonStr As String) As String` — returns a complete standalone HTML document embedding a Leaflet 1.9.4 map with ISP-colored CircleMarkers, MarkerCluster 1.5.3 clustering, CartoDB Positron basemap, and a bottom-right ISP color legend.
- `WriteUtf8File(filePath As String, content As String)` — writes to disk via ADODB.Stream with `Charset = "utf-8"` (never FSO).

## Key Implementation Details

- **ISP palette**: ColorBrewer Set1 (7 colors: `#e41a1c` through `#f781bf`), overflow gray `#999999`, determined by alphabetical sort of ISP names.
- **Marker style**: `radius:8, fillOpacity:0.85, weight:1, color:"#ffffff"` (white border).
- **CDN load order**: Leaflet CSS → MarkerCluster.css → MarkerCluster.Default.css → Leaflet JS → MarkerCluster JS (both CSS files required — Pitfall 1).
- **Basemap**: CartoDB Positron with full OSM/CARTO attribution (license requirement).
- **fitBounds guard**: `if(markers.length > 0)` prevents error on empty dataset.
- **Phase 3 compatibility**: `var markers = []` parallel array populated alongside the cluster group.
- **SRI integrity**: Leaflet 1.9.4 JS and CSS include `integrity=` hashes; MarkerCluster 1.5.3 pinned by version URL.

## Acceptance Criteria Verification

All acceptance criteria from the plan pass:
- `Attribute VB_Name = "mod_MapBuilder"` and `Option Explicit` present ✓
- `AppendLine` private helper appends `line & vbLf` ✓
- `WriteUtf8File` uses `ADODB.Stream`, `Charset = "utf-8"`, `SaveToFile ... 2` ✓
- `BuildMapHtml` function declared and implemented ✓
- `FileSystemObject` does not appear in the file ✓
- All CDN links present in correct order ✓
- `var markers = []` present ✓
- `L.markerClusterGroup()`, `ispNames.sort()`, `bottomright`, `cartocdn.com/light_all` all present ✓

## Files Created

| File | Purpose |
|------|---------|
| `src/mod_MapBuilder.bas` | HTML generation module (BuildMapHtml + WriteUtf8File) |

## Commits

- `feat(02-01): create mod_MapBuilder.bas with BuildMapHtml and WriteUtf8File`
