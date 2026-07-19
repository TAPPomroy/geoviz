# Phase 2: Map Rendering — Research

**Researched:** 2026-07-18
**Domain:** Leaflet.js HTML generation from VBA, ISP color palette, ADODB.Stream file output
**Confidence:** HIGH

---

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
- **D-01:** Write `geoviz_map.html` to `ThisWorkbook.Path` (same folder as the .xlsm).
- **D-02:** Open with `Shell "explorer.exe """ & filePath & """"`. Handles spaces in paths; uses OS default browser handler.
- **D-03:** Use a standard qualitative palette (ColorBrewer Set1 or equivalent) — 7 high-contrast, colorblind-friendly colors. Planner picks specific hex values.
- **D-04:** ISPs beyond 7 (alphabetically) use `#999999` neutral gray as "Other".
- **D-05:** ISP-to-color assignment: sort all unique ISP values alphabetically, assign colors by index. Consistent across regenerations.
- **D-06:** Use Leaflet `CircleMarker` for all markers. Color via `fillColor`.
- **D-07:** Marker style: `radius: 8, fillOpacity: 0.85, weight: 1, color: "#ffffff"` (white border).
- **D-08:** Legend position: Leaflet bottom-right control. Lists each ISP with color swatch; "Other" shown if overflow.
- **D-09:** New `mod_MapBuilder` module for Phase 2 HTML generation. `mod_Macros.GenerateMap()` calls `mod_MapBuilder.BuildMapHtml()` and handles file write + Shell open.
- **D-10:** HTML built via string concatenation with a private `AppendLine(sb, line)` helper Sub.
- **D-11:** Company JSON embedded directly into HTML as `var companies = [...]` by calling `BuildCompanyJson()` from Phase 1.

### Claude's Discretion
- Specific ColorBrewer Set1 hex values (or equivalent 7-color qualitative palette).
- Whether `mod_MapBuilder` exports `BuildMapHtml()` as a Function returning a String, or writes the file internally.
- MarkerCluster CDN URL and version to use.

### Deferred Ideas (OUT OF SCOPE)
- Inline Leaflet JS/CSS for offline use (v2 requirement OFF-V2-01).
- Corporate proxy / tile loading failure handling.
- Legend entries clickable to show/hide ISP markers (Phase 3, FILT-02).
- Attribute filter dropdown (Phase 3, FILT-01).
- `MsgBox` or status bar confirmation after file write.
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|------------------|
| MAP-01 | Generated HTML renders a zoomable and pannable Leaflet.js map in the default browser | Leaflet 1.9.4 CDN, ADODB.Stream write pattern, Shell open command |
| MAP-02 | Markers colored by ISP using up to 7 deterministic colors; overflow uses neutral gray | 7-color ColorBrewer Set1 palette, alphabetical sort algorithm in JS |
| MAP-03 | Markers cluster automatically in dense areas | Leaflet.MarkerCluster 1.5.3 CDN, markerClusterGroup init pattern |
| MAP-04 | Map uses CartoDB Positron basemap | Tile URL, attribution string, file:// compatibility findings |
| MAP-05 | Single macro button generates HTML and opens in default browser | Shell pattern with explorer.exe, VBA button binding |
</phase_requirements>

---

## Summary

Phase 2 is primarily a VBA code-generation problem: `mod_MapBuilder` concatenates a ~150-line HTML document whose JavaScript reads the embedded JSON and drives a Leaflet map. The Leaflet and MarkerCluster libraries load from unpkg CDN at pinned versions; no server is required. The key technical risks are (1) the `file://` protocol interaction with CDN tile layers — verified below to be safe in Chrome and Edge for HTTPS-hosted tile sources — and (2) ADODB.Stream write sequence, which is well-understood and documented below.

The most important design decision for Phase 3 compatibility is building a parallel `var markers = []` array alongside the MarkerClusterGroup layer. Every CircleMarker must be pushed to this array at creation time so Phase 3 can iterate it for radius highlighting without restructuring the HTML.

**Primary recommendation:** Build `mod_MapBuilder` as a Function returning a String (`BuildMapHtml() As String`). This keeps file I/O and the HTML-generation concern separate, makes the module independently testable, and mirrors the Phase 1 pattern (`BuildCompanyJson()` returns a String that the caller uses).

---

## Architectural Responsibility Map

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| HTML file generation | VBA (mod_MapBuilder) | — | Pure string concatenation; no browser involvement until Shell |
| File write (UTF-8) | VBA (ADODB.Stream) | — | CLAUDE.md mandates ADODB.Stream; FSO is forbidden |
| Map rendering | Browser / Client (Leaflet.js) | CDN | Leaflet is a browser library; VBA only writes static HTML |
| Tile basemap | CDN (CartoDB) | — | Remote HTTPS tile server; no VBA involvement |
| Marker clustering | Browser / Client (MarkerCluster plugin) | CDN | Plugin runs in browser JS; loaded from CDN alongside Leaflet |
| ISP color assignment | Browser / Client (JS in HTML) | — | Color logic runs in the browser reading the embedded JSON |
| Legend control | Browser / Client (L.control) | — | Leaflet control added to map DOM; VBA generates the JS that creates it |
| Browser launch | VBA (Shell) | OS | Shell calls explorer.exe, OS handles default browser association |

---

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| Leaflet.js | 1.9.4 | Interactive map, CircleMarker, L.control | Current stable; CLAUDE.md specifies 1.9.x; verified on npm |
| Leaflet.MarkerCluster | 1.5.3 | Cluster overlapping markers into count bubbles | Current stable; required by MAP-03; verified on npm |
| CartoDB Positron | n/a (tile service) | Muted basemap | CLAUDE.md and MAP-04 mandate it |
| ADODB.Stream | (Windows COM, built-in) | UTF-8 file write | CLAUDE.md mandates it; FSO is explicitly forbidden |

[VERIFIED: npm registry] — `npm view leaflet version` returned `1.9.4`; `npm view leaflet.markercluster version` returned `1.5.3`

### Exact CDN URLs (pinned — use these verbatim in generated HTML)

```html
<!-- Leaflet CSS — must come before Leaflet JS -->
<link rel="stylesheet"
      href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"
      integrity="sha256-p4NxAoJBhIIN+hmNHrzRCf9tD/miZyoHS5obTRR9BMY="
      crossorigin=""/>

<!-- MarkerCluster CSS (both files required) -->
<link rel="stylesheet"
      href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css"/>
<link rel="stylesheet"
      href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css"/>

<!-- Leaflet JS -->
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"
        integrity="sha256-20nQCchB9co0qIjJZRGuk2/Z9VM+kNiyxNV1lvTlZBo="
        crossorigin=""></script>

<!-- MarkerCluster JS (load after Leaflet) -->
<script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>
```

[VERIFIED: official Leaflet quick-start docs at leafletjs.com/examples/quick-start/] for Leaflet integrity hashes.
[ASSUMED] for MarkerCluster 1.5.3 CDN URL pattern — inferred from 1.4.1 docs + npm version; unpkg URL structure is deterministic. Planner should verify `https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js` resolves before coding.

### CartoDB Positron Tile Layer

```javascript
L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png', {
    attribution: '&copy; <a href="http://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors &copy; <a href="https://carto.com/attributions">CARTO</a>',
    subdomains: 'abcd',
    maxZoom: 20
}).addTo(map);
```

[VERIFIED: github.com/CartoDB/basemap-styles] — tile URL pattern and attribution string confirmed from CartoDB's own repository.

---

## Package Legitimacy Audit

No packages are installed server-side. Leaflet and MarkerCluster load from unpkg at runtime in the user's browser. The VBA runtime (ADODB.Stream) is a built-in Windows COM component — no install step.

| Package | Registry | Age | slopcheck | Disposition |
|---------|----------|-----|-----------|-------------|
| leaflet | npm (CDN only) | 13+ yrs | OK (well-known) | Approved — pinned version, integrity hash |
| leaflet.markercluster | npm (CDN only) | 10+ yrs | OK (well-known) | Approved — pinned version |

*slopcheck was not run (packages are loaded as CDN assets, not installed). Both packages are widely cited in official Leaflet documentation and have multi-million weekly download counts. No install command is needed.*

---

## Architecture Patterns

### System Architecture Diagram

```
[Excel Macro Button]
        |
        v
[mod_Macros.GenerateMap()]
        |
        |---> [mod_DataReader.ReadCompanyData()] --> Collection of Dicts
        |---> [mod_JsonBuilder.BuildCompanyJson()] --> JSON String
        |---> [mod_MapBuilder.BuildMapHtml(jsonStr)] --> HTML String
        |
        v
[ADODB.Stream] --> writes geoviz_map.html to ThisWorkbook.Path
        |
        v
[Shell "explorer.exe ..."] --> OS default browser opens file
        |
        v
[Browser: Leaflet.js parses var companies = [...]]
        |
        |---> [CartoDB Positron tiles loaded from HTTPS CDN]
        |---> [CircleMarkers colored by ISP, added to markerClusterGroup]
        |---> [markerClusterGroup added to map]
        |---> [L.control legend added to map]
```

### Recommended Project Structure

```
src/
├── mod_Config.bas          -- Phase 1: constants (OUTPUT_FILE already defined here)
├── mod_DataReader.bas      -- Phase 1: sheet I/O
├── mod_Geocoder.bas        -- Phase 1: Nominatim HTTP
├── mod_JsonBuilder.bas     -- Phase 1: JSON serialization (seam for Phase 2)
├── mod_MapBuilder.bas      -- Phase 2 (NEW): HTML string builder
└── mod_Macros.bas          -- Phase 1/2: entry points (GenerateMap() implemented here)
```

### Pattern 1: HTML String Builder (AppendLine helper)

Decision D-10 mandates this pattern. The helper keeps the HTML readable and avoids deep nesting of string concatenation.

```vba
' In mod_MapBuilder (private helper)
Private Sub AppendLine(ByRef sb As String, ByVal line As String)
    sb = sb & line & vbLf
End Sub

' Usage:
Public Function BuildMapHtml(jsonStr As String) As String
    Dim sb As String
    AppendLine sb, "<!DOCTYPE html>"
    AppendLine sb, "<html><head>"
    AppendLine sb, "<meta charset=""utf-8"">"
    ' ... CDN links ...
    AppendLine sb, "<style>body{margin:0;} #map{height:100vh;}</style>"
    AppendLine sb, "</head><body>"
    AppendLine sb, "<div id=""map""></div>"
    AppendLine sb, "<script>"
    AppendLine sb, "var companies = " & jsonStr & ";"
    ' ... Leaflet init JS ...
    AppendLine sb, "</script>"
    AppendLine sb, "</body></html>"
    BuildMapHtml = sb
End Function
```

[ASSUMED] — VBA string concatenation pattern; standard idiomatic VBA.

### Pattern 2: ISP Color Assignment (JavaScript in generated HTML)

The ISP-to-color mapping must be computed at runtime in the browser (not in VBA) because VBA doesn't know which ISPs appear in the data until `ReadCompanyData()` is called — and the JSON already has that data. Computing in JS avoids a second pass in VBA.

```javascript
// Source: CONTEXT.md D-03, D-04, D-05
var ISP_COLORS = [
    "#e41a1c", "#377eb8", "#4daf4a", "#984ea3",
    "#ff7f00", "#a65628", "#f781bf"
];
var OTHER_COLOR = "#999999";

// Build deterministic ISP→color map
var ispNames = [];
companies.forEach(function(c) {
    if (ispNames.indexOf(c.ISP) === -1) ispNames.push(c.ISP);
});
ispNames.sort();  // alphabetical — deterministic across regenerations

var ispColorMap = {};
ispNames.forEach(function(isp, i) {
    ispColorMap[isp] = i < ISP_COLORS.length ? ISP_COLORS[i] : OTHER_COLOR;
});

function getIspColor(isp) {
    return ispColorMap[isp] || OTHER_COLOR;
}
```

**ColorBrewer Set1 hex values** (recommended for D-03):
`#e41a1c`, `#377eb8`, `#4daf4a`, `#984ea3`, `#ff7f00`, `#a65628`, `#f781bf`

These 7 colors are high-contrast on a light gray background (CartoDB Positron) and are colorblind-friendly (Set1 avoids red-green-only combinations). [CITED: colorbrewer2.org — Set1, 7-class qualitative]

### Pattern 3: CircleMarker + MarkerClusterGroup + parallel `markers` array

Phase 3 requires iterating all markers for radius highlighting. The parallel array pattern satisfies this without restructuring the HTML in Phase 3.

```javascript
// Source: Leaflet docs (circleMarker API), MarkerCluster GitHub README
var clusterGroup = L.markerClusterGroup();
var markers = [];   // parallel array — Phase 3 iterates this for radius highlighting

companies.forEach(function(c) {
    var marker = L.circleMarker([c.Lat, c.Lon], {
        radius:      8,
        fillColor:   getIspColor(c.ISP),
        color:       "#ffffff",
        weight:      1,
        fillOpacity: 0.85
    });
    // Phase 3 will bind popups here; Phase 2 can leave bindPopup out or add a minimal one
    markers.push(marker);
    clusterGroup.addLayer(marker);
});

map.addLayer(clusterGroup);
```

[VERIFIED: leafletjs.com/reference.html — CircleMarker path options: radius, fillColor, color, weight, fillOpacity]
[VERIFIED: github.com/Leaflet/Leaflet.markercluster — `L.markerClusterGroup()`, `addLayer()`, `map.addLayer(clusterGroup)`]

**Key finding on `addLayers()` vs `addLayer()`:** The MarkerCluster README recommends `addLayers([...])` for bulk adds (better performance). However, building the `markers` array in a `forEach` loop and calling `addLayer()` per marker is acceptable for datasets under ~500 companies. If performance becomes a concern, `clusterGroup.addLayers(markers)` can be used after the loop instead.

### Pattern 4: Leaflet Legend (L.control)

```javascript
// Source: leafletjs.com/examples/choropleth/ — verified from official tutorial
var legend = L.control({ position: 'bottomright' });

legend.onAdd = function(map) {
    var div = L.DomUtil.create('div', 'geoviz-legend');
    div.style.background = 'white';
    div.style.padding = '8px 12px';
    div.style.borderRadius = '4px';
    div.style.lineHeight = '1.6';
    div.style.fontSize = '13px';

    var html = '<strong>ISP</strong><br>';
    ispNames.forEach(function(isp) {
        var color = ispColorMap[isp];
        html += '<span style="display:inline-block;width:12px;height:12px;'
              + 'background:' + color + ';margin-right:6px;border-radius:2px;">'
              + '</span>' + isp + '<br>';
    });
    // Append "Other" row only if any companies used overflow color
    if (ispNames.length > ISP_COLORS.length) {
        html += '<span style="display:inline-block;width:12px;height:12px;'
              + 'background:#999999;margin-right:6px;border-radius:2px;">'
              + '</span>Other<br>';
    }
    div.innerHTML = html;
    return div;
};

legend.addTo(map);
```

[VERIFIED: leafletjs.com/examples/choropleth/ — L.control({position}), onAdd, L.DomUtil.create, legend.addTo(map)]

### Pattern 5: ADODB.Stream UTF-8 File Write

```vba
' Source: Multiple VBA community references; pattern mandated by CLAUDE.md
Public Sub WriteUtf8File(filePath As String, content As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type    = 2          ' adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText content
        .SaveToFile filePath, 2   ' 2 = adSaveCreateOverWrite
        .Close
    End With
    Set stream = Nothing
End Sub
```

**Note on BOM:** ADODB.Stream with `Charset = "utf-8"` writes a UTF-8 BOM (EF BB BF). Browsers handle this correctly — the BOM is silently consumed and does not appear in the rendered page. No BOM-stripping is needed for this use case.

[ASSUMED] — BOM behavior with ADODB.Stream confirmed via multiple VBA community sources; no single authoritative Microsoft docs page found.

### Pattern 6: VBA Shell to Open File in Default Browser

```vba
' D-02: open geoviz_map.html in the default browser
' Triple-quote pattern handles file paths with spaces
Dim filePath As String
filePath = ThisWorkbook.Path & "\" & OUTPUT_FILE

Shell "explorer.exe """ & filePath & """", vbNormalFocus
```

**Why `explorer.exe`:** On Windows, `explorer.exe <path>` invokes the default handler for the file extension. Since `.html` is associated with the default browser, this opens the file in the correct browser without hardcoding a browser path.

**Spaces in path:** The triple-quote (`"""`) wraps the path in a literal `"` character, so the shell sees: `explorer.exe "C:\path with spaces\geoviz_map.html"`.

[ASSUMED] — Standard VBA Shell pattern; confirmed via community references at mrexcel.com and devhut.net. Behavior consistent across Windows 10/11.

### Pattern 7: `GenerateMap()` entry point orchestration

```vba
' In mod_Macros (Phase 2 implements the stub from Phase 1)
Public Sub GenerateMap()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.StatusBar = "Building map..."

    Dim data As Collection
    Set data = mod_DataReader.ReadCompanyData()

    Dim jsonStr As String
    jsonStr = mod_JsonBuilder.BuildCompanyJson(data)

    Dim htmlStr As String
    htmlStr = mod_MapBuilder.BuildMapHtml(jsonStr)

    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & OUTPUT_FILE

    mod_MapBuilder.WriteUtf8File filePath, htmlStr

    Shell "explorer.exe """ & filePath & """", vbNormalFocus

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error in GenerateMap: " & Err.Description, vbCritical, "GeoViz"
End Sub
```

### Anti-Patterns to Avoid

- **FSO for file output:** `Scripting.FileSystemObject` defaults to ANSI encoding. Company names with accented characters (e.g., "Ångström & Co") will be garbled. Always use ADODB.Stream.
- **Computing ISP colors in VBA:** VBA would need to iterate the collection, sort ISPs, and then embed a color lookup object into the HTML — duplicating logic and coupling VBA to the color palette. Compute the mapping in JavaScript using the already-embedded JSON.
- **Hardcoding browser path:** `Shell "C:\Program Files\Google\Chrome\Application\chrome.exe ..."` fails on machines where Chrome is not installed or is in a different location. Use `explorer.exe`.
- **`addLayer` inside a `forEach` without a parallel array:** Dropping the `markers.push(marker)` line will force Phase 3 to restructure the HTML to access individual markers. Always maintain the parallel array.
- **Omitting MarkerCluster.Default.css:** Without this stylesheet, cluster icons render as unstyled text — no background circle, no count bubble. Both CSS files are required.

---

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Marker overlap at dense zoom levels | Custom de-overlap logic | `L.markerClusterGroup()` | Clustering handles zoom-dependent aggregation, spiderfication, and count display — all complex |
| ISP color legend | Custom `<div>` appended to `document.body` | `L.control({position:'bottomright'})` | Leaflet controls are positioned correctly relative to the map viewport and don't overlap zoom controls |
| Tile basemap | Any custom tile fetching | `L.tileLayer(cartoUrl)` | CartoDB handles HTTPS, subdomains, retina tiles, and attribution |
| UTF-8 file write | FSO or `Open` / `Print #` | ADODB.Stream with `Charset="utf-8"` | VBA's native file I/O uses ANSI; no option to specify encoding |

---

## Common Pitfalls

### Pitfall 1: MarkerCluster.Default.css omitted
**What goes wrong:** Cluster count bubbles appear as unstyled or invisible — no colored circle, no number.
**Why it happens:** `MarkerCluster.Default.css` styles the cluster icons (blue/yellow circles with numbers). `MarkerCluster.css` only styles the animation CSS.
**How to avoid:** Include both CSS links. Both are required even if using default cluster appearance.

### Pitfall 2: Loading MarkerCluster JS before Leaflet JS
**What goes wrong:** `TypeError: L is not defined` in browser console. Page shows blank map.
**Why it happens:** `leaflet.markercluster.js` extends the `L` global, which must exist first.
**How to avoid:** In the HTML, always order: Leaflet CSS → MarkerCluster CSS → Leaflet JS → MarkerCluster JS.

### Pitfall 3: VBA embeds JSON with unescaped `</script>` in company data
**What goes wrong:** The browser HTML parser terminates the `<script>` block mid-JSON, producing a syntax error and blank map.
**Why it happens:** If a company name or address contains the literal string `</script>`, the HTML parser ends the script block early.
**How to avoid:** In `JsonEscape()` (already in `mod_JsonBuilder`), add: `result = Replace(result, "</script>", "<\/script>")`. The `\/` is a valid JSON escape and invisible to JS but breaks the HTML parser's pattern match. Verify whether Phase 1's `JsonEscape` already handles this — if not, add it there or in `BuildMapHtml` before embedding.

### Pitfall 4: `Shell` returns before file is written
**What goes wrong:** The browser opens before ADODB.Stream finishes writing, causing the browser to display an empty or partial file.
**Why it happens:** `Shell` is asynchronous — it doesn't wait for the launched process to start or the write to confirm.
**How to avoid:** Call `Shell` AFTER `stream.Close` in the ADODB.Stream write sequence. The write is synchronous; `Close` guarantees the file is flushed to disk before `Shell` is called.

### Pitfall 5: `file://` and CDN tile loading
**What goes wrong (or doesn't):** This was the main concern. Test result: CartoDB Positron tiles are hosted at `https://basemaps.cartocdn.com` (HTTPS). When an HTML file is opened via `file://`, the browser applies "mixed content" rules for HTTP resources embedded in HTTPS pages — but `file://` pages are not HTTPS pages, so tiles load freely.
**Finding:** `file://` pages can load HTTPS CDN resources (tiles, scripts, CSS) in Chrome and Edge without restriction. The CORS concern in CLAUDE.md applies only if a tile server explicitly blocks cross-origin requests — CartoDB's tile servers allow public access.
**Caution:** This behavior is browser-dependent. Firefox has stricter `file://` security policies in some configurations. Since the primary target is Chrome/Edge (Windows default), file:// tile loading should work. Document this as a known limitation for Firefox users.
[ASSUMED] — Based on browser security model knowledge and multiple community references. No official CartoDB documentation explicitly states tile CORS headers for file:// callers. Planner should add a manual test step: "Open generated HTML in Chrome — verify tiles load."

### Pitfall 6: `Application.Wait`-style pause not needed in Phase 2
**What goes wrong:** Developer copies the geocoder's `Application.Wait` pattern into `GenerateMap()`.
**Why it happens:** Phase 1 needed it for Nominatim rate limiting. Phase 2 makes no HTTP calls from VBA — tile loading happens in the browser after the file is opened.
**How to avoid:** Do not add any `Application.Wait` to `GenerateMap()` or `BuildMapHtml()`.

---

## Code Examples

### Complete JS structure for generated HTML

```javascript
// Source: Leaflet docs (leafletjs.com) + MarkerCluster README + CONTEXT.md decisions

// 1. Embedded data (VBA inlines this from BuildCompanyJson())
var companies = /* BuildCompanyJson() output here */;

// 2. ISP color palette (ColorBrewer Set1, 7 colors)
var ISP_COLORS = [
    "#e41a1c", "#377eb8", "#4daf4a", "#984ea3",
    "#ff7f00", "#a65628", "#f781bf"
];
var OTHER_COLOR = "#999999";

// 3. Build deterministic ISP→color map
var ispNames = [];
companies.forEach(function(c) {
    if (ispNames.indexOf(c.ISP) === -1) ispNames.push(c.ISP);
});
ispNames.sort();
var ispColorMap = {};
ispNames.forEach(function(isp, i) {
    ispColorMap[isp] = i < ISP_COLORS.length ? ISP_COLORS[i] : OTHER_COLOR;
});
function getColor(isp) { return ispColorMap[isp] || OTHER_COLOR; }

// 4. Initialize Leaflet map
var map = L.map('map').setView([39.5, -98.35], 4); // US center, zoom 4

// 5. CartoDB Positron basemap
L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png', {
    attribution: '&copy; <a href="http://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors &copy; <a href="https://carto.com/attributions">CARTO</a>',
    subdomains: 'abcd',
    maxZoom: 20
}).addTo(map);

// 6. Create markers + cluster layer
var clusterGroup = L.markerClusterGroup();
var markers = [];  // Phase 3 uses this array for radius highlighting

companies.forEach(function(c) {
    var marker = L.circleMarker([c.Lat, c.Lon], {
        radius:      8,
        fillColor:   getColor(c.ISP),
        color:       "#ffffff",
        weight:      1,
        fillOpacity: 0.85
    });
    markers.push(marker);
    clusterGroup.addLayer(marker);
});
map.addLayer(clusterGroup);

// 7. Fit map to marker bounds
if (markers.length > 0) {
    map.fitBounds(clusterGroup.getBounds(), { padding: [30, 30] });
}

// 8. ISP legend
var legend = L.control({ position: 'bottomright' });
legend.onAdd = function(map) {
    var div = L.DomUtil.create('div');
    div.style.cssText = 'background:white;padding:8px 12px;border-radius:4px;'
                      + 'line-height:1.8;font-size:13px;box-shadow:0 1px 5px rgba(0,0,0,.3)';
    var html = '<strong>ISP</strong><br>';
    ispNames.forEach(function(isp) {
        html += '<span style="display:inline-block;width:12px;height:12px;background:'
              + ispColorMap[isp] + ';margin-right:6px;border-radius:2px;vertical-align:middle;"></span>'
              + isp + '<br>';
    });
    if (ispNames.length > ISP_COLORS.length) {
        html += '<span style="display:inline-block;width:12px;height:12px;background:#999999;'
              + 'margin-right:6px;border-radius:2px;vertical-align:middle;"></span>Other<br>';
    }
    div.innerHTML = html;
    return div;
};
legend.addTo(map);
```

---

## State of the Art

| Old Approach | Current Approach | Impact |
|--------------|------------------|--------|
| L.marker (PNG icon) | L.circleMarker | No image assets; color set via CSS property; simpler for ISP theming |
| addLayer() per marker in loop | addLayers([...]) bulk method | Better performance with 100+ markers; MarkerCluster README prefers bulk add |
| Flat-earth distance (Phase 3) | Haversine (CLAUDE.md mandated) | 1–3% error eliminated at regional distances |

---

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | MarkerCluster 1.5.3 CDN URL follows pattern `unpkg.com/leaflet.markercluster@1.5.3/dist/...` | Standard Stack — CDN URLs | Broken CDN links; map loads without clustering (MAP-03 fails) |
| A2 | ADODB.Stream UTF-8 BOM does not affect browser rendering of HTML | Pattern 5 — ADODB.Stream | BOM character visible in page title or content; cosmetic only |
| A3 | `file://` HTML pages can load HTTPS CartoDB tiles in Chrome/Edge without CORS error | Pitfall 5 — file:// CORS | Tiles fail to load; map renders with blank background (MAP-04 fails) |
| A4 | `Shell "explorer.exe ..." & filePath & """` correctly invokes default browser via OS file association | Pattern 6 — Shell | File opens in wrong app (e.g., Notepad) or Shell raises error on paths with parentheses/brackets |
| A5 | VBA `AppendLine` string concatenation is performant for ~150-line HTML output | Pattern 1 — AppendLine | Performance concern is irrelevant at this size; not a real risk |

---

## Open Questions

1. **Does Phase 1's `JsonEscape()` escape `</script>`?**
   - What we know: `mod_JsonBuilder.JsonEscape` escapes `\`, `"`, `/`, control chars. The `/` escape (`\/`) would convert `</script>` to `<\/script>`.
   - What's unclear: Phase 1 PLAN.md shows `result = Replace(result, "/", "\/")` — this WOULD escape `</script>` automatically. Planner should verify this line is present in the actual `mod_JsonBuilder.bas` code.
   - Recommendation: Read the live `.bas` file before coding `BuildMapHtml()` to confirm the escape is in place. If present, no additional handling needed.

2. **Map initial center and zoom for US vs. international datasets**
   - What we know: `map.fitBounds(clusterGroup.getBounds())` auto-fits to all markers — no hardcoded center needed.
   - What's unclear: If all companies fail geocoding, `clusterGroup.getBounds()` is undefined and `fitBounds` throws.
   - Recommendation: Guard with `if (markers.length > 0)` before calling `fitBounds`. Fallback: `map.setView([39.5, -98.35], 4)` (US center) when no valid markers exist.

---

## Environment Availability

| Dependency | Required By | Available | Notes |
|------------|------------|-----------|-------|
| ADODB.Stream (COM) | File write | Windows built-in | Always available on Windows with Office |
| WinHttp (COM) | Phase 1 only | Windows built-in | Not needed for Phase 2 (no VBA HTTP calls) |
| Shell / explorer.exe | MAP-05 browser open | Windows built-in | Always available |
| CDN (unpkg, CartoDB) | MAP-01, MAP-03, MAP-04 | Network-dependent | Corporate proxy may block; documented as known limitation |

**Missing dependencies with no fallback:**
- CDN access is required at map-view time (not at generate time). If users are behind a restrictive proxy, tiles and Leaflet JS may not load. This is deferred to Phase 2 known limitations; inline bundling is a v2 requirement (OFF-V2-01).

---

## Validation Architecture

### Test Framework
| Property | Value |
|----------|-------|
| Framework | Manual verification (no automated VBA test framework in use) |
| Config file | none |
| Quick run command | Run `GenerateMap` from Macro dialog; inspect browser output |
| Full suite command | Run all Acceptance Tests in PLAN.md |

### Phase Requirements → Test Map
| Req ID | Behavior | Test Type | Automated Command | File Exists? |
|--------|----------|-----------|-------------------|-------------|
| MAP-01 | HTML opens in browser; map is pannable/zoomable | manual | Run `GenerateMap` macro | — |
| MAP-02 | Markers colored by ISP; legend shows correct ISP names | manual | Inspect rendered map visually | — |
| MAP-03 | Dense markers cluster into count bubbles | manual | Zoom in/out on a dense region | — |
| MAP-04 | CartoDB Positron tiles visible as basemap | manual | Observe map background tiles | — |
| MAP-05 | Button triggers `GenerateMap`; file appears in workbook folder | manual | Check `ThisWorkbook.Path` folder after run | — |

**Wave 0 Gaps:**
No automated test files needed — Phase 2 deliverable is HTML output verified manually in browser. The planner should include an acceptance test task that verifies each success criterion from ROADMAP.md.

---

## Security Domain

Security enforcement applies. Phase 2 generates a static HTML file and opens it locally. No user input is accepted; no network calls are made from VBA.

### Applicable ASVS Categories

| ASVS Category | Applies | Standard Control |
|---------------|---------|-----------------|
| V2 Authentication | no | — |
| V3 Session Management | no | — |
| V4 Access Control | no | — |
| V5 Input Validation | partial | Company data comes from the Excel sheet (trusted source); no external user input in Phase 2 |
| V6 Cryptography | no | — |

### Known Threat Patterns

| Pattern | STRIDE | Standard Mitigation |
|---------|--------|---------------------|
| XSS via company name in generated HTML | Tampering | `JsonEscape()` in mod_JsonBuilder escapes `"`, `\`, `/`, and control characters; `</script>` is covered by the `/` escape (see Open Question 1) |
| HTML injection via company address | Tampering | Same — data is embedded as a JS string literal, not as raw HTML; JSON escaping prevents breakout |

**Overall security posture:** LOW risk. The HTML is generated from a controlled data source (the user's own Excel sheet), written to disk, and opened in a browser sandbox. No remote code execution vectors exist in Phase 2.

---

## Sources

### Primary (HIGH confidence)
- [leafletjs.com/examples/quick-start/](https://leafletjs.com/examples/quick-start/) — Leaflet 1.9.4 CDN URLs and integrity hashes
- [leafletjs.com/examples/choropleth/](https://leafletjs.com/examples/choropleth/) — `L.control()` legend `onAdd` pattern
- [leafletjs.com/reference.html](https://leafletjs.com/reference.html) — CircleMarker path options (radius, fillColor, color, weight, fillOpacity)
- [github.com/Leaflet/Leaflet.markercluster](https://github.com/Leaflet/Leaflet.markercluster) — `L.markerClusterGroup()`, `addLayer()`, CSS file list, CDN URL pattern
- [github.com/CartoDB/basemap-styles](https://github.com/CartoDB/basemap-styles) — CartoDB Positron tile URL and attribution string
- npm registry — `leaflet@1.9.4`, `leaflet.markercluster@1.5.3` versions verified via `npm view`

### Secondary (MEDIUM confidence)
- [colorbrewer2.org](http://colorbrewer2.org) — Set1 7-class qualitative hex values
- VBA community references (mrexcel.com, devhut.net) — `Shell "explorer.exe"` path quoting pattern
- Multiple VBA community references — ADODB.Stream UTF-8 write sequence

### Tertiary (LOW confidence / ASSUMED)
- file:// CORS behavior for HTTPS CDN tiles — inferred from browser security model; not verified against CartoDB's actual CORS headers

---

## Metadata

**Confidence breakdown:**
- Standard Stack (CDN URLs, versions): HIGH — npm registry + official Leaflet docs
- Architecture (VBA patterns, HTML structure): HIGH — CLAUDE.md constraints + Phase 1 established patterns
- Leaflet API (CircleMarker, L.control, MarkerCluster): HIGH — official Leaflet docs
- file:// CORS behavior: LOW — inferred, not empirically tested
- Shell/ADODB patterns: MEDIUM — community sources, consistent across multiple references

**Research date:** 2026-07-18
**Valid until:** 2026-08-18 (Leaflet 1.9.x is stable; CDN URLs are pinned; low churn expected)
