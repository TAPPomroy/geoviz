# Research Summary: GeoViz

**Synthesized:** 2026-07-18
**Sources:** STACK.md, FEATURES.md, ARCHITECTURE.md, PITFALLS.md

---

## Recommended Stack

| Layer | Choice | Rationale |
|-------|--------|-----------|
| Map library | Leaflet.js 1.9.x via jsDelivr | Smallest footprint (~42 KB), no WebGL, dominant community |
| Tile provider | OpenStreetMap standard tiles | No API key, permissive CORS on tile images |
| Clustering | Leaflet.markercluster 1.5.x | Handles overlapping pins at metro density; include from start |
| Geocoding | Nominatim | Free, no key, acceptable for 100 addresses at 1 req/sec |
| HTTP client | WinHttp.WinHttpRequest.5.1 | Better TLS handling than MSXML2 on corporate machines |
| JSON parsing | Custom VBA (InStr/Mid) | No library needed for Nominatim shallow response |
| File output | ADODB.Stream (Charset=UTF-8) | FSO CreateTextFile writes ANSI; ADODB.Stream is correct |
| Proximity math | Haversine inline JS | 10 lines; no Turf.js dependency at this scale |

Do not use: MapLibre GL (WebGL risk on corporate hardware), OpenLayers (overbuilt), MSXML2.XMLHTTP as primary (SSL edge cases), FSO for file writing (ANSI encoding).

---

## Table Stakes Features

The core loop is non-negotiable: click company -> see popup -> see radius circle -> see neighbors highlighted.

1. Pan and zoom -- Leaflet default, zero cost
2. Clickable markers with popup -- company name, address, ISP; readable without scrolling
3. ISP-based color coding -- deterministic (alphabetical sort -> fixed palette), 5-7 colors max with Other bucket; colorblind-safe palette required
4. Legend -- static HTML div generated alongside map; colors are meaningless without it
5. Radius circle on marker click -- L.circle drawn immediately on click, before slider interaction
6. Neighbor highlighting -- dim non-neighbors to ~30% opacity, never hide; geographic context must be preserved
7. Neighbor count in popup header -- N companies within X km on [ISP] as first line
8. Geocode cache written back to sheet -- lat/lon + timestamp + display_name; skip already-geocoded rows on re-run

---

## Architecture Highlights

Pipeline is sequential, stateless between phases. VBA runs DataReader -> Geocoder -> HtmlBuilder -> FileOutput.
The worksheet is the only persistent state. The HTML file is the only output. No server, no Python, no Node.

Four VBA modules with strict boundaries:
- mod_Config: constants only, no logic; single source of truth for column indices, URLs, filenames
- mod_DataReader: only module that touches the worksheet; returns a Collection; all others never access the sheet
- mod_Geocoder: HTTP calls + cache write-back; Application.Wait enforces the rate limit
- mod_HtmlBuilder: largest module; builds full HTML as a single String; never writes to disk
- mod_FileOutput: writes via ADODB.Stream UTF-8; launches browser via WScript.Shell

Data embedding: JSON literal in a script tag with type=application/json and id=gv-data. Parsed in JS via JSON.parse.
Single file, DevTools-debuggable, no CORS issues, no sidecar files required.

All interactivity in JavaScript, not VBA. VBA serializes data; JS handles filtering, Haversine distance,
color assignment, and opacity changes. Map is debuggable in browser DevTools without touching Excel.

LayerGroup pattern for filtering: markerLayer.clearLayers() + re-add. A single updateProximity() handles
both click and slider events to prevent logic drift between handlers.

File location: ThisWorkbook.Path + geoviz_map.html. Fixed filename so browser reloads rather than accumulating
stale tabs. Fall back to Desktop if workbook is unsaved.

---

## Top Pitfalls to Avoid

**1. Nominatim rate limit violation (CRITICAL)**
A bare VBA loop fires 100 requests in seconds. Results return blank or HTTP 429 with no obvious indication
it was a rate error. Prevention: Application.Wait 1 second after every call; check HTTP status; write
GEOCODE_FAILED sentinel (not blank) on errors to prevent retry loops.

**2. Missing User-Agent causes silent blocking (CRITICAL)**
Nominatim policy requires an identifying User-Agent. Generic strings are blocked without a clear error.
Prevention: always set User-Agent: GeoViz/1.0 (pomroyanalytics@gmail.com) on every request.

**3. Mark of the Web blocks macro execution (CRITICAL)**
Since 2022, Excel blocks macros in files received by email or download. The tool silently does nothing
for first-time users. Prevention: document right-click > Properties > Unblock prominently; sign the macro
or use a Trusted Location for distribution.

**4. JSON injection breaks the map (HIGH)**
Company names with quotes, backslashes, or angle brackets produce malformed JavaScript and a blank map.
Prevention: apply a JsonEscape() VBA function to every string field before HTML embedding.

**5. UTF-8 encoding corruption (HIGH)**
FSO CreateTextFile(unicode:=False) writes ANSI. Accented characters in company names corrupt silently.
Prevention: use ADODB.Stream with Charset=utf-8 exclusively for file output.

**6. Geocode cache overwrite destroys validated data (HIGH)**
A re-run that overwrites good lat/lon with a wrong Nominatim result silently degrades data quality.
Prevention: skip rows where lat/lon is populated and not stale; only overwrite on successful response.

**7. OneDrive sync delay (MEDIUM)**
ThisWorkbook.Path may point to a OneDrive-synced directory; Shell opening the file immediately after
write can hit a race condition. Prevention: detect OneDrive paths and warn user, or write to TEMP instead.

**8. Tile CORS on Chrome (MEDIUM)**
Corporate proxy configurations can block OSM tile loading from the file:// origin.
Prevention: test specifically on Chrome on the target machine; document internet requirement.

---

## Phase Implications

Three build phases emerge naturally from the research. Each is independently testable before the next begins.

### Phase 1 -- Data Pipeline (VBA Core)

Build mod_Config, mod_DataReader, mod_Geocoder, and the file-write portion of mod_FileOutput. Deliver a
macro that geocodes uncached rows, writes lat/lon + timestamp + display_name back to the sheet, and emits
a valid JSON array. Testable by pasting JSON output into a browser console.

Must solve: rate limiting, User-Agent header, cache skip logic, GEOCODE_FAILED sentinel, JsonEscape()
function, ADODB.Stream UTF-8 output, late binding for all COM objects.

### Phase 2 -- Map Rendering (Static HTML)

Build mod_HtmlBuilder and wire ThisWorkbook.RunGeoViz() to a button. Deliver a generated HTML file with
all companies as ISP-colored markers, popups, legend, and markercluster. No interactivity yet.

Must solve: map container explicit pixel height, Leaflet init order (script at bottom of body), CDN version
lock (1.9.4 not latest), file path double-quoting in Shell, MOTW documentation for distribution.

### Phase 3 -- Interactivity (JS Features)

Add the full interaction loop in the JS template: click handler producing radius circle + neighbor
highlighting + neighbor count; radius slider wired to updateProximity(); ISP filter toggles; same-ISP mode
toggle; Clear button; clipboard export. Developed and tested in browser DevTools against Phase 2 HTML --
no VBA changes required.

Must solve: dim (not hide) non-neighbors to 30% opacity; single updateProximity() shared by click and slider;
colorblind-safe palette with Other bucket for more than 7 ISPs.

Defer to v2: bidirectional sidebar + company list sync (genuinely fiddly event wiring), sidebar name search,
mobile-responsive layout.

### Research Flags

| Phase | Needs research pass? | Reason |
|-------|---------------------|--------|
| Phase 1 | No | VBA HTTP and cache patterns fully documented and verified |
| Phase 2 | No | Leaflet initialization and HTML generation are standard patterns |
| Phase 3 | No | Haversine, opacity toggling, and filter patterns are all straightforward |

### Gaps to resolve before Phase 1 begins

- Confirm tile loading behavior on the target machine (corporate proxy presence unknown)
- Confirm whether ThisWorkbook.Path resolves to a OneDrive path (affects output file strategy)
- Confirm MOTW distribution scenario (email/download vs. local copy)
- Decide: inline Leaflet JS/CSS at build time for offline resilience, or rely on jsDelivr CDN for simplicity
