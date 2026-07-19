# Stack Research: GeoViz

**Project:** VBA-driven Excel → standalone HTML map generator
**Researched:** 2026-07-18
**Constraint:** Pure VBA + HTML/JS, zero external installs, ~100 companies max

---

## Recommended Stack

| Layer | Technology | Version / CDN |
|-------|-----------|---------------|
| Map library | Leaflet.js | 1.9.x via jsDelivr CDN |
| Tile provider | OpenStreetMap (tile.openstreetmap.org) | Built-in, no key |
| Clustering | Leaflet.markercluster | 1.5.x via jsDelivr CDN |
| HTTP client (VBA) | WinHttp.WinHttpRequest.5.1 | Built-in Windows COM |
| Geocoding API | Nominatim (nominatim.openstreetmap.org) | Free, no key |
| File output | FileSystemObject (Scripting.FileSystemObject) | Built-in Windows COM |
| JSON parsing | Custom VBA parser (regex or Mid/InStr) | No external dependency |

All CDN references load at map-open time. The HTML file is fully self-describing with inline JS — no local server required.

---

## Map Library Decision

**Recommendation: Leaflet.js 1.9.x**

### Why Leaflet wins for this use case

- **Simplest API for raster tile + marker use case.** The target is plotting ~100 points with popups on an OSM basemap. Leaflet does this in ~20 lines of JS.
- **Smallest footprint.** Leaflet core is ~42 KB minified+gzipped. MapLibre GL is ~280 KB. For a standalone HTML file where the user's browser downloads assets fresh each time, this matters for offline/slow network scenarios.
- **No WebGL requirement.** Leaflet renders via SVG/Canvas. MapLibre GL requires WebGL — enterprise Windows machines with locked-down GPU drivers or older integrated graphics can silently fail.
- **Plugin ecosystem matches all needed features.** Clustering (markercluster), circle radius visualization (L.circle), custom popup styling — all covered with well-maintained plugins.
- **1.4M+ npm downloads/month in 2025.** Dominant community, extensive SO answers, no deprecation risk.

### Why not MapLibre GL

MapLibre is the right choice when you need vector tiles, 3D terrain, or dynamic style switching. None of those apply here. The WebGL dependency is a real operational risk on corporate machines. Overkill.

### Why not OpenLayers

OpenLayers shines for WMS/WFS layers, projection handling, and complex GIS. API surface is significantly larger. Initial learning curve and bundle size (~400 KB) are not justified for a points-on-a-map use case.

### Leaflet CDN snippet (inline in generated HTML)

```html
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.min.css"/>
<script src="https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.min.js"></script>
```

Use jsDelivr over unpkg — jsDelivr has better uptime SLA and is the Leaflet project's recommended CDN.

---

## Geocoding from VBA

**Recommendation: WinHttp.WinHttpRequest.5.1 against Nominatim**

### HTTP client: WinHttp over MSXML2.XMLHTTP

`WinHttp.WinHttpRequest.5.1` is the correct choice for this project:

- Runs at the Windows HTTP stack level — faster, more reliable TLS handling than MSXML2.
- Successfully resolves HTTPS endpoints where MSXML2 occasionally fails due to certificate chain issues on restrictive Windows policies.
- Ships on every Windows version from XP onward — zero deployment risk.
- Not XML-aware (fine here — you parse JSON yourself).

`MSXML2.XMLHTTP` is acceptable as a fallback if WinHttp is somehow blocked by GPO, but prefer WinHttp as primary.

### Nominatim usage: acceptable for 100 companies

The official Nominatim Usage Policy requires:
- Max **1 request per second** (hard limit)
- A descriptive `User-Agent` header identifying your application
- No bulk commercial use

100 companies takes ~100 seconds (1.7 minutes) at compliant rate. This is well within acceptable non-commercial use. Caching lat/lon back to the Excel sheet (as planned) means geocoding only runs once per address — subsequent map generations are instant.

### Rate-limiting pattern

```vba
' Core geocoding function skeleton
Function Geocode(address As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    Dim url As String
    url = "https://nominatim.openstreetmap.org/search?q=" & _
          EncodeURL(address) & "&format=json&limit=1"
    
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", "GeoViz-Excel/1.0 (pomroyanalytics@gmail.com)"
    http.Send
    
    If http.Status = 200 Then
        Geocode = http.ResponseText  ' parse lat/lon from JSON
    Else
        Geocode = ""
    End If
    
    Application.Wait Now + TimeValue("00:00:01")  ' 1-second rate limit
End Function
```

### JSON parsing (no library needed)

For Nominatim's response, the lat/lon fields are shallow and consistent. Use `InStr` + `Mid` extraction — no need for a full JSON parser at 100 records:

```vba
' Extract "lat":"VALUE" from JSON string
Function ExtractJSON(json As String, key As String) As String
    Dim pos As Long, endPos As Long
    pos = InStr(json, """" & key & """:""") + Len(key) + 4
    endPos = InStr(pos, json, """")
    ExtractJSON = Mid(json, pos, endPos - pos)
End Function
```

### Error handling requirements

- HTTP status != 200: log to a "GeocodeErrors" column, skip to next row
- Empty result array `[]`: address not found — flag cell in red
- Network timeout: set `http.SetTimeouts 5000, 5000, 10000, 10000` (connect/send/receive ms)
- Re-run protection: check if lat/lon columns already populated before calling API

---

## Tile Providers

**Recommendation: OpenStreetMap standard tiles (no API key, unlimited for non-commercial)**

### Tier 1 — No key required, works in any browser

| Provider | URL Pattern | Attribution Required | Notes |
|----------|------------|---------------------|-------|
| OpenStreetMap | `https://tile.openstreetmap.org/{z}/{x}/{y}.png` | Yes (OSM contributors) | Default choice. Reliable. |
| CartoDB Positron | `https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png` | CARTO + OSM | Clean, light-colored — good for colored markers |
| CartoDB Dark Matter | `https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png` | CARTO + OSM | High contrast for dense point sets |
| Stamen Toner Lite | Via Stadia CDN (requires free account) | Stamen + OSM | Minimalist, print-friendly |

CartoDB Positron is the best secondary option — the light grey background makes colored ISP-provider markers visually distinct without OSM's label noise.

### Tier 2 — Free tier with key (worth evaluating)

- **Maptiler** — 100K tiles/month free. Clean vector-style raster tiles, better label rendering than OSM standard.
- **Thunderforest** — 150K tiles/month free. Good terrain and transport styles.

For an internal company tool with ~100 uses/month, Tier 1 keyless providers are sufficient. Do not complicate the deployment with API key management.

### CORS/standalone HTML consideration

All Tier 1 providers listed above serve tiles with permissive CORS headers. Standalone `file://` HTML can load them without issue — browsers do not block cross-origin tile images (only XHR/fetch from `file://` is restricted, which is not used for tile loading).

---

## VBA HTML Generation

**Recommendation: FileSystemObject with a string-builder accumulator pattern**

### Approach: accumulate into a String variable, single-write at end

Do NOT write line-by-line to the file in a loop — each `TextStream.WriteLine` call is a separate disk write. For a 200–500 line HTML file, accumulate into a VBA `String` and write once:

```vba
Sub GenerateMap()
    Dim html As String
    html = html & "<!DOCTYPE html>" & vbLf
    html = html & "<html><head><meta charset='utf-8'>" & vbLf
    html = html & "<title>GeoViz Company Map</title>" & vbLf
    ' ... (Leaflet CSS/JS links, styles) ...
    html = html & "<script>" & vbLf
    html = html & "  var markers = " & BuildMarkersJSON() & ";" & vbLf
    html = html & "</script>" & vbLf
    ' ... rest of body/script ...
    
    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile(outputPath, True, False)  ' False = UTF-8 without BOM via ADODB trick
    f.Write html
    f.Close
    
    Shell "explorer.exe """ & outputPath & """", vbNormalFocus
End Sub
```

### UTF-8 without BOM (important for Leaflet/JS)

`FileSystemObject.CreateTextFile` with `Unicode:=False` writes ANSI, not UTF-8. Company names with accented characters or non-ASCII will corrupt. Use `ADODB.Stream` instead:

```vba
Dim stream As Object
Set stream = CreateObject("ADODB.Stream")
stream.Type = 2         ' text
stream.Charset = "UTF-8"
stream.Open
stream.WriteText html
stream.SaveToFile outputPath, 2  ' 2 = overwrite
stream.Close
```

This is the correct pattern — ADODB.Stream is available on all Windows machines with Office installed.

### Data injection pattern

Emit the company data as a JS array literal directly in the HTML, not via a fetch() call (which would fail from `file://`):

```javascript
var companies = [
  { name: "Acme Corp", lat: 45.123, lon: -93.456, isp: "Comcast", address: "123 Main St" },
  // ...
];
```

VBA builds this array string by looping the Excel table. This pattern avoids any need for a local server or XHR.

### File placement

Write to the same directory as the workbook (`ThisWorkbook.Path`) so the path is always writable and findable. Offer the user a Save As dialog via `Application.GetSaveAsFilename` if customization is needed.

---

## Recommended Plugins

All via jsDelivr CDN — no npm, no build step.

### 1. Leaflet.markercluster — INCLUDE

**Why:** At 100 companies there may be geographic clusters (e.g., multiple companies in a metro area). Without clustering, overlapping markers make popups unclickable.

**Behavior:** Groups nearby markers into numbered circle clusters that expand on click. Spiderfies at max zoom. Works out of the box with zero configuration.

```html
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/leaflet.markercluster@1.5.3/dist/MarkerCluster.min.css"/>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.min.css"/>
<script src="https://cdn.jsdelivr.net/npm/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.min.js"></script>
```

### 2. L.circle for radius visualization — USE BUILT-IN

Leaflet's built-in `L.circle(latlng, {radius: meters})` draws a circle that scales with zoom. This is the right primitive for "coverage area" or "signal radius" visualization if needed. No plugin required.

### 3. Custom popup styling — CSS ONLY, no plugin needed

Leaflet popups accept arbitrary HTML. Override the default white box with CSS in the `<style>` block of the generated HTML:

```css
.leaflet-popup-content-wrapper {
    border-radius: 4px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.3);
}
.leaflet-popup-content {
    font-family: Segoe UI, sans-serif;
    font-size: 13px;
    line-height: 1.5;
}
```

No plugin needed for styled popups. Full HTML (tables, links, colored badges) renders inside popup content.

### 4. leaflet-sidebar-v2 — OPTIONAL, include if detail panel needed

If company attributes are numerous (ISP, contract dates, contacts, etc.) and the popup becomes cluttered, `leaflet-sidebar-v2` provides a slide-in side panel triggered by marker click. Adds ~15 KB. Only include if popup space is genuinely insufficient.

CDN: `https://cdn.jsdelivr.net/npm/leaflet-sidebar-v2@3.2.3/src/`

### 5. Leaflet.heat — SKIP

A heatmap layer would be misleading at 100 points (implies density data you don't have). Clustering + individual markers is more honest and more useful.

### 6. Leaflet.fullscreen — NICE TO HAVE

Adds a fullscreen button (top-right corner). 3 KB. Worth including for usability:

```html
<script src="https://cdn.jsdelivr.net/npm/leaflet.fullscreen@3.0.0/Control.FullScreen.min.js"></script>
```

---

## Confidence Levels

| Area | Confidence | Basis |
|------|------------|-------|
| Leaflet vs MapLibre vs OL decision | HIGH | Official docs, download stats, multiple comparison sources verified 2025 |
| WinHttp vs MSXML2 recommendation | HIGH | Microsoft docs + VBA-tools GitHub issue discussions |
| Nominatim rate limit (1 req/sec) | HIGH | Official OSM Foundation policy page |
| Nominatim acceptability for 100 addrs | MEDIUM | Policy text is ambiguous on "non-commercial small use" — compliant behavior (1 req/sec + User-Agent) is well established |
| CDN tile provider CORS behavior | MEDIUM | OSM and CARTO are well-documented; edge cases on corporate proxy setups possible |
| ADODB.Stream UTF-8 pattern | HIGH | Standard pattern documented on Microsoft Learn and widely used in VBA community |
| Leaflet.markercluster stability | HIGH | Official Leaflet org repo, 1.5.3 released 2022, no breaking changes since |
| leaflet-sidebar-v2 maintenance | MEDIUM | Last release 2021; works with Leaflet 1.9.x but watch for Leaflet 2.x compatibility |

---

## Sources

- [Geoapify: Map Libraries Comparison 2025](https://www.geoapify.com/map-libraries-comparison-leaflet-vs-maplibre-gl-vs-openlayers-trends-and-statistics/)
- [Jawg Blog: MapLibre GL vs Leaflet](https://blog.jawg.io/maplibre-gl-vs-leaflet-choosing-the-right-tool-for-your-interactive-map/)
- [Web Maps in 2025 - Grendelman.net](https://www.grendelman.net/wp/web-maps-in-2025/)
- [Nominatim Official Usage Policy](https://operations.osmfoundation.org/policies/nominatim/)
- [OSM Wiki: Nominatim Usage Policy](https://wiki.openstreetmap.org/wiki/Nominatim_usage_policy)
- [Raster Tile Providers - OSM Wiki](https://wiki.openstreetmap.org/wiki/Raster_tile_providers)
- [Leaflet-providers demo (all free CDN providers)](https://leaflet-extras.github.io/leaflet-providers/preview/)
- [WinHttp vs MSXML2 VBA-tools discussion](https://github.com/VBA-tools/VBA-Web/issues/420)
- [the-automator: XMLHTTPRequest vs WinHTTPRequest](https://www.the-automator.com/xmlhttprequest-vs-winhttprequest/)
- [Microsoft Learn: FileSystemObject](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object)
- [Leaflet.markercluster GitHub](https://github.com/leaflet/leaflet.markercluster)
- [leaflet-sidebar-v2 GitHub](https://github.com/noerw/leaflet-sidebar-v2)
- [Astro-Geo-GIS: Nominatim Geocoding in Excel](https://astro-geo-gis.com/the-costless-way-to-geocoding-addresses-in-excel-part-3-bulk-data-geocoding-with-nominatim-and-others-geocoding-tools/)
