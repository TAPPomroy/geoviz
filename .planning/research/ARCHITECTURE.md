# Architecture Research: GeoViz

**Researched:** 2026-07-18
**Confidence:** HIGH (Leaflet official docs via Context7, Nominatim policy, verified VBA patterns)

---

## Component Overview

The system has three discrete phases of execution that share no runtime — they run sequentially
in the same VBA process and hand off data through the worksheet and the file system.

```
[Excel Worksheet]
       |
       v
[VBA: mod_DataReader]  -- reads address table, detects missing/stale geocodes
       |
       v
[VBA: mod_Geocoder]    -- calls Nominatim, writes lat/lon back to sheet cache columns
       |
       v
[VBA: mod_HtmlBuilder] -- serialises data to JSON, builds full HTML string
       |
       v
[VBA: mod_FileOutput]  -- writes .html file, launches default browser]
       |
       v
[Browser: standalone HTML with embedded Leaflet + embedded JSON data]
       |
       +-- [JS: map init, marker rendering]
       +-- [JS: click handler -> draw proximity circle]
       +-- [JS: slider input -> refilter markers by radius]
```

No server. No Python. No external runtime. The HTML file is fully self-contained and requires
only an internet connection for OpenStreetMap tile images (or a pre-downloaded tile set).

---

## VBA Module Structure

Use **standard modules** (.bas), not class modules, for this scale of project. Four modules
plus the optional ThisWorkbook event stub.

### `mod_Config` — constants only, no logic

```vba
' mod_Config
Public Const NOMINATIM_URL   As String = "https://nominatim.openstreetmap.org/search"
Public Const USER_AGENT      As String = "GeoViz/1.0 pomroyanalytics@gmail.com"
Public Const RATE_LIMIT_MS   As Long   = 1100   ' > 1 req/sec per Nominatim policy
Public Const DATA_SHEET      As String = "Locations"
Public Const COL_ADDRESS     As Long   = 1   ' A
Public Const COL_LAT         As Long   = 2   ' B  (cache)
Public Const COL_LON         As Long   = 3   ' C  (cache)
Public Const COL_GEOCODE_TS  As Long   = 4   ' D  (timestamp of last geocode)
Public Const COL_LABEL       As Long   = 5   ' E  (popup label)
Public Const DATA_FIRST_ROW  As Long   = 2   ' row 1 is header
Public Const OUTPUT_FILENAME As String = "geoviz_map.html"
```

Rationale: a single constants module prevents magic strings from scattering across all
other modules. Any config change is one-file-only.

### `mod_DataReader` — worksheet I/O

Responsibility: read the table into an array/collection; detect which rows need geocoding.

```vba
' Returns a Collection of Dictionaries (scripting.dictionary), one per data row.
' Rows with non-empty lat/lon that are not stale are returned as-is.
' Rows missing lat/lon, or with stale timestamps (> STALE_DAYS old), get lat="" lon=""
' so mod_Geocoder knows to fill them.
Public Function ReadLocations() As Collection
```

Keep worksheet access in one module. Every other module receives a Collection — it never
touches the sheet directly. This means unit-testing logic modules is possible in isolation.

### `mod_Geocoder` — HTTP + rate limiting + cache write-back

Responsibility: for each row where lat/lon is empty, call Nominatim, parse JSON/XML
response, write lat/lon and timestamp back to the sheet.

```vba
Public Sub GeocodeAll(locations As Collection)
' Loops collection; skips rows that already have lat/lon.
' Sleeps RATE_LIMIT_MS between requests (Application.Wait).
' Writes result back to sheet via COL_LAT, COL_LON, COL_GEOCODE_TS.
```

Use `MSXML2.XMLHTTP60` (late-bound: `CreateObject("MSXML2.XMLHTTP.6.0")`). Request
`format=json` from Nominatim — easier to parse with VBA-JSON or manual string slicing
than XML. Set the `User-Agent` header on every request (Nominatim policy requirement).

### `mod_HtmlBuilder` — data serialisation + HTML template

Responsibility: transform the Collection into a JSON literal; inject it and all static
assets into the HTML template string; return the complete HTML as a String.

```vba
Public Function BuildHtml(locations As Collection) As String
```

This is the largest module. Keep the HTML template as a series of `Const` string chunks
and concatenate with the dynamic JSON payload in the middle. Avoid building HTML line by
line in a loop — build once into a `StringBuilder`-style pattern using a `String` variable
with `& vbLf`.

### `mod_FileOutput` — write file + launch browser

Responsibility: accept an HTML string, write it to disk, open it.

```vba
Public Sub WriteAndOpen(htmlContent As String, outputPath As String)
    ' Scripting.FileSystemObject CreateTextFile (overwrite=True, unicode=False)
    ' Shell "explorer.exe """ & outputPath & """"
```

### `ThisWorkbook` — entry point only

```vba
' Optional: expose a ribbon/shortcut-triggered entry point
Public Sub RunGeoViz()
    Dim locs As Collection
    Set locs = mod_DataReader.ReadLocations()
    mod_Geocoder.GeocodeAll locs
    Dim html As String
    html = mod_HtmlBuilder.BuildHtml(locs)
    mod_FileOutput.WriteAndOpen html, mod_Config.GetOutputPath()
End Sub
```

---

## Data Embedding Strategy

**Recommended: JSON literal in a `<script>` tag — not a data URI, not a .js sidecar.**

Rationale:

| Option | Verdict | Reason |
|--------|---------|--------|
| JSON in `<script>` tag | **Use this** | Single file, trivially readable, debuggable in browser DevTools, no CORS issues |
| `data:` URI for whole file | Avoid | Base64 bloat, unreadable, breaks browser history/reload |
| Separate `.js` sidecar | Avoid | `file://` protocol blocks cross-origin script loads in Chrome/Edge; requires same-folder discipline |
| Fetch from URL | Avoid | Requires a server |

### Implementation

In `mod_HtmlBuilder`, emit a block like:

```html
<script id="gv-data" type="application/json">
[
  {"label":"Acme Corp","address":"123 Main St","lat":45.123,"lon":-93.456},
  ...
]
</script>
```

Then in the map JS, parse it:

```javascript
const locations = JSON.parse(
    document.getElementById('gv-data').textContent
);
```

Using `type="application/json"` on the script tag prevents the browser from trying to
execute it as JS. `textContent` is safe on all modern browsers.

---

## Geocoding Cache Design

### Column layout (in the data worksheet)

| Col | Name | Type | Notes |
|-----|------|------|-------|
| A | Address | String | Source — user-maintained |
| B | Lat | Double | Nominatim result; blank = needs geocoding |
| C | Lon | Double | Nominatim result; blank = needs geocoding |
| D | GeocodedAt | Date | Excel Date serial written by VBA |
| E | Label | String | Map popup text (may differ from address) |
| F+ | Any business columns | — | Passed through to JSON as extra properties |

Put cache columns to the **right** of the address column and to the **left** of business
data. Never interleave cache cols between business cols — it breaks `ListObject` table
sorting.

### Staleness detection

```vba
Const STALE_DAYS As Long = 90

Function NeedsGeocode(latVal As Variant, tsVal As Variant) As Boolean
    If IsEmpty(latVal) Or latVal = "" Then
        NeedsGeocode = True
        Exit Function
    End If
    If IsEmpty(tsVal) Or tsVal = "" Then
        NeedsGeocode = True   ' has lat but no timestamp — legacy row
        Exit Function
    End If
    NeedsGeocode = (Now() - CDate(tsVal)) > STALE_DAYS
End Function
```

Never delete the old lat/lon when a row is flagged stale — keep it as a fallback display
value while the new geocode request is in flight. Only overwrite on successful response.

### Failure handling

If Nominatim returns zero results or HTTP error: write `"GEOCODE_FAILED"` to the Lat column
and the current timestamp to GeocodedAt. This prevents repeated retries on every run for
bad addresses. The HTML builder skips rows where Lat = "GEOCODE_FAILED".

---

## Map JS Architecture

Embed all JS directly in the HTML `<body>` in a single `<script>` block (no modules,
no ES6 imports — `file://` protocol and some corp browsers block module loading).

### Structure inside the script block

```javascript
// 1. Parse embedded data
const locations = JSON.parse(document.getElementById('gv-data').textContent);

// 2. Initialise map — fitBounds to data extent, not hardcoded centre
const map = L.map('map');
const markerLayer = L.layerGroup().addTo(map);
let proximityCircle = null;  // the single drawn radius circle
let selectedLatLng = null;   // the clicked point

// 3. Tile layer (online) — see Offline section for alternative
L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '&copy; OpenStreetMap contributors',
    maxZoom: 19
}).addTo(map);

// 4. Render all markers, fit bounds
function renderMarkers(data) {
    markerLayer.clearLayers();
    const bounds = [];
    data.forEach(function(loc) {
        const m = L.marker([loc.lat, loc.lon])
            .bindPopup('<b>' + loc.label + '</b><br>' + loc.address);
        markerLayer.addLayer(m);
        bounds.push([loc.lat, loc.lon]);
    });
    if (bounds.length > 0) map.fitBounds(bounds, {padding: [30, 30]});
}
renderMarkers(locations);

// 5. Click on map -> draw radius circle, refilter markers
map.on('click', function(e) {
    selectedLatLng = e.latlng;
    updateProximity();
});

// 6. Slider -> refilter
document.getElementById('radiusSlider').addEventListener('input', function() {
    document.getElementById('radiusDisplay').textContent = this.value + ' km';
    if (selectedLatLng) updateProximity();
});

// 7. Proximity update — single function, called by both click and slider
function updateProximity() {
    const radiusKm = parseFloat(document.getElementById('radiusSlider').value);
    const radiusM  = radiusKm * 1000;

    if (proximityCircle) map.removeLayer(proximityCircle);
    proximityCircle = L.circle(selectedLatLng, {
        radius: radiusM, color: '#2563eb', fillOpacity: 0.08
    }).addTo(map);

    const nearby = locations.filter(function(loc) {
        return haversineKm(selectedLatLng.lat, selectedLatLng.lng,
                           loc.lat, loc.lon) <= radiusKm;
    });
    renderMarkers(nearby);
}

// 8. Haversine formula (pure JS, no Turf dependency needed at this scale)
function haversineKm(lat1, lon1, lat2, lon2) {
    const R = 6371;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
              Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
              Math.sin(dLon/2) * Math.sin(dLon/2);
    return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}
```

Key decisions:
- `markerLayer` is a `LayerGroup`, not added directly to map. `clearLayers()` + re-add
  is the correct Leaflet pattern for dynamic filtering — the `filter` option on GeoJSON
  does not re-evaluate already-added layers (Leaflet docs limitation).
- A single `updateProximity()` function handles both click and slider events, preventing
  logic drift between the two handlers.
- No Turf.js dependency — haversine inline keeps the file self-contained without a CDN call
  for Turf. Turf adds ~250 KB for math the haversine formula covers in 10 lines.

---

## Proximity Calculation Approach

The slider passes its value to `updateProximity()` via a shared DOM read — no global
variable needed for the radius itself. The `selectedLatLng` state variable is the only
persistent JS state.

### Reset behaviour

Add a "Clear" button that sets `selectedLatLng = null`, removes `proximityCircle`,
and calls `renderMarkers(locations)` to restore all markers. This avoids a confusing
"stuck filter" state when users navigate away and return.

### Slider UI placement

Place the slider as an overlay panel using absolute-positioned CSS inside the map container,
not outside it. This keeps the HTML output visually self-contained (no layout dependence on
body/container sizing) and works on any screen resolution.

---

## File Output Strategy

### File location

Write to the same directory as the workbook:

```vba
Function GetOutputPath() As String
    GetOutputPath = ThisWorkbook.Path & "\" & OUTPUT_FILENAME
End Function
```

Do not use `Environ("TEMP")`. Temp files are ephemeral — users lose them across reboots or
temp-cleanup tools. A named file next to the workbook is findable, shareable, and
re-openable. If `ThisWorkbook.Path` is empty (workbook never saved), fall back to
`Environ("USERPROFILE") & "\Desktop\"` and warn the user.

### File writing

```vba
Sub WriteAndOpen(htmlContent As String, outputPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim f As Object
    Set f = fso.CreateTextFile(outputPath, True, False)  ' overwrite=True, unicode=False (UTF-8 via explicit encoding)
    f.Write htmlContent
    f.Close
    ' Open in default browser
    CreateObject("WScript.Shell").Run """" & outputPath & """"
End Sub
```

Write the file as UTF-8 without BOM. Since `CreateTextFile` with `unicode=False` writes
ANSI, and VBA string manipulation is ANSI-safe for Latin characters, this is acceptable for
English addresses. If international addresses are needed, use `ADODB.Stream` with charset
`utf-8` instead of FSO.

### Overwrite strategy

Always overwrite — never append. The file represents the current state of the sheet.
If users want to preserve a snapshot, they copy the file manually.

---

## Offline Considerations

### Map tiles

Leaflet tiles are raster images fetched from `tile.openstreetmap.org`. They cannot be
inlined into a single HTML file without extraordinary effort (each zoom/x/y tile is a
separate PNG — thousands of files for any useful area).

**Practical options:**

| Option | When to Use | Complexity |
|--------|------------|------------|
| Online OSM tiles (default) | Any machine with internet | Zero |
| Cached browser tiles | Repeat viewing on same machine; browser caches tile requests automatically for ~1 week | Zero extra work |
| Local tile server (TileServer-GL + MBTiles) | True air-gap requirement | High — requires Node.js install, not VBA-compatible |
| Protomaps PMTiles + local HTTP | Offline requirement, technical users | Medium — single .pmtiles file, needs a tiny HTTP server |

**Recommendation for this project:** Use online OSM tiles. Document the internet
requirement. If offline is needed as a future milestone, add a config constant
`TILE_URL` in `mod_Config` that points to a local tile server URL, making the switch
a one-line change.

### Leaflet library itself

Leaflet's CSS and JS (~150 KB combined unminified) can be base64-embedded directly in
the HTML `<head>` using a `<style>` and `<script>` block. VBA can read the Leaflet files
from a known local path and embed them at HTML-build time. This makes the HTML file work
with no CDN dependency, even if tiles still need internet.

```vba
' In mod_HtmlBuilder:
Function ReadFileAsString(path As String) As String
    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(path, 1)  ' ForReading
    ReadFileAsString = f.ReadAll
    f.Close
End Function
```

Ship `leaflet.css` and `leaflet.js` (downloaded once from leafletjs.com/download) in a
`/assets/` subfolder next to the workbook. The builder reads them at run time and inlines
them. This is the most robust approach — no CDN, no network dependency for the library,
only tiles require internet.

---

## Build Order

Build in this sequence. Each step is testable independently before the next begins.

| Step | Module(s) | Deliverable | Test |
|------|-----------|-------------|------|
| 1 | `mod_Config` | All constants defined | Manual inspection |
| 2 | `mod_DataReader` | `ReadLocations()` returns collection from sheet | Print to Immediate Window |
| 3 | `mod_Geocoder` | Single-address geocode test against Nominatim | One row, check B/C/D columns |
| 4 | `mod_Geocoder` | Batch geocode with rate limiting and cache skip | All rows, verify only blanks are hit |
| 5 | `mod_HtmlBuilder` | JSON serialiser only — emit just the data block | Paste into browser console, verify parse |
| 6 | `mod_HtmlBuilder` | Full HTML with hardcoded Leaflet CDN links | Open file, verify map renders |
| 7 | `mod_HtmlBuilder` | Replace CDN with inlined local Leaflet assets | Open file offline, verify map renders |
| 8 | JS (in HTML template) | Marker rendering from embedded data | DevTools console |
| 9 | JS | Click handler + proximity circle | Manual click test |
| 10 | JS | Slider + `updateProximity()` + `renderMarkers()` filter | Drag slider, verify counts |
| 11 | `mod_FileOutput` | Write file + open browser | Full end-to-end run |
| 12 | `ThisWorkbook` | `RunGeoViz()` entry point wired to button/shortcut | Full user workflow |

---

## Sources

- Leaflet documentation (via Context7 `/websites/leafletjs`): initialization, LayerGroup, fitBounds, GeoJSON filter limitation
- [Nominatim Usage Policy](https://operations.osmfoundation.org/policies/nominatim/) — 1 req/sec limit, caching requirement, User-Agent requirement
- [FileSystemObject — Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object)
- [Leaflet Download Page](https://leafletjs.com/download.html) — local install instructions
- [leaflet.offline (GitHub)](https://github.com/allartk/leaflet.offline) — considered and deferred; adds complexity beyond project scope
- VBA module organisation: [Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/5268832/question-about-best-practices-for-vba-modules-1-wi), [MrExcel thread](https://www.mrexcel.com/board/threads/best-practices-for-organizing-subroutines-and-modules.798936/)
- [nominatim-excel (GitHub)](https://github.com/zbyna/nominatim-excel) — reference implementation reviewed for HTTP pattern
