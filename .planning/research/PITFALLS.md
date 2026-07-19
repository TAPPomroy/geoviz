# Pitfalls Research: GeoViz

**Domain:** VBA-based geospatial company mapping tool
**Researched:** 2026-07-18
**Stack:** Excel VBA + Nominatim geocoding + Leaflet standalone HTML
**Scale:** Single user, ~100 companies

---

## Geocoding Pitfalls

### CRITICAL: Nominatim Rate Limit Enforcement

**What goes wrong:** The public Nominatim API enforces an absolute hard limit of 1 request per second. Scripts exceeding this receive `"Usage limit reached"` errors or get silently blocked. Because VBA loops run faster than expected, a tight geocoding loop for 100 companies can fire 100 requests in seconds and hit the wall immediately.

**Why it happens:** No built-in throttling in VBA. A bare `For Each` loop with an HTTP call completes each request in ~50-200ms, well under the 1-second floor.

**Consequences:** Requests return HTTP 429 or a Nominatim error JSON. If not handled, the loop silently writes empty lat/lon to the cache sheet, and the failure looks like a missing result rather than a rate error.

**Prevention:**
- Insert `Application.Wait Now + TimeValue("00:00:01")` after every geocoding call.
- Add a secondary check: if the response body contains `"error"` or is empty, log the failure and do NOT write blank coordinates to cache.
- Check HTTP status code; treat anything other than 200 as a failure.

**Detection:** Results come back blank for all companies after the first few. Check the raw response string in a debug variable.

---

### CRITICAL: Missing or Wrong User-Agent Causes Silent Blocking

**What goes wrong:** Nominatim policy explicitly requires a valid, identifying `User-Agent` or `HTTP Referer` header. The default `Msxml2.XMLHTTP` User-Agent string is generic and non-identifying. Requests without a proper User-Agent may be blocked without a clear error message.

**Why it happens:** VBA developers set no custom headers; the library sends a generic WinHTTP or MSXML agent string.

**Prevention:**
```vba
xhr.setRequestHeader "User-Agent", "GeoViz/1.0 (pomroyanalytics@gmail.com)"
```
Use a string that identifies your application and includes contact info. Nominatim's policy specifically states stock User-Agents will not suffice.

---

### Ambiguous Address Returns Wrong Location

**What goes wrong:** Nominatim returns the first result from a potentially large list. "Springfield, USA" could match any of dozens of Springfields. A company address with only a city and state — no street — routinely resolves to the geographic center of that city, not the company location.

**Why it happens:** Nominatim picks the highest-importance OSM match. For partial addresses, that is often the most famous place with that name (e.g., Springfield, IL rather than Springfield, OR).

**Consequences:** Markers appear in wrong states or countries. Proximity calculations are skewed. Problems may not be noticed until a user reports a clearly wrong pin location.

**Prevention:**
- Always include full address: street number, street name, city, state, postal code.
- Log the `display_name` field from the response alongside lat/lon in the cache sheet so you can visually audit what was actually matched.
- For ambiguous results, check the `importance` score in the response (lower = less confident match).
- Consider caching the full response JSON or at least the matched `display_name` to enable auditing.

---

### International Address Failures

**What goes wrong:** Non-US addresses frequently fail or geocode imprecisely. Address format assumptions (street number before street name) do not hold worldwide. Some countries have poor OSM coverage.

**Why it happens:** Nominatim is strong for Europe and North America, inconsistent for Asia, Africa, and South America. Address component order varies by country.

**Prevention:**
- Do not assume any particular address structure for international entries.
- Pass addresses as a single `q=` freeform query rather than structured fields.
- Flag international entries in the cache sheet for manual review.
- Consider falling back to country-level coordinates if city-level fails.

---

### Cache Sheet Corruption Overwrites Good Data

**What goes wrong:** A re-geocoding run (triggered accidentally or for a partial update) may overwrite previously validated lat/lon values with new, incorrect geocoder results.

**Prevention:**
- Add a "Geocoded" boolean column. Skip any row that is already marked geocoded.
- Only geocode rows with empty lat/lon, or rows explicitly flagged for refresh.
- Never auto-clear the cache before a run.

---

## VBA HTTP Pitfalls

### SSL Certificate Errors with Msxml2.XMLHTTP

**What goes wrong:** `Msxml2.XMLHTTP` can fail on HTTPS endpoints if there is any certificate issue (expired cert, unknown CA, hostname mismatch). This manifests as a VBA runtime error rather than a meaningful HTTP error message.

**Why it happens:** MSXML validates SSL certificates strictly by default, and corporate environments sometimes have proxy-injected certificates not trusted by the MSXML certificate store.

**Prevention:**
- Use `Msxml2.ServerXMLHTTP.6.0` instead of `Msxml2.XMLHTTP.6.0`. ServerXMLHTTP is designed for server-side/scripted use and handles certificates more robustly.
- To explicitly ignore certificate errors (use only for internal/trusted endpoints):
  ```vba
  xhr.setOption 2, 13056  ' SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
  ```
  Values: 256 = unknown CA, 4096 = CN mismatch, 8192 = date invalid, 13056 = ignore all.

---

### Timeout Hangs the UI

**What goes wrong:** If Nominatim is slow or unreachable, the default XMLHTTP timeout is either infinite or very long. Excel freezes with no feedback to the user.

**Prevention:**
```vba
xhr.setTimeouts 5000, 10000, 10000, 10000  ' resolve, connect, send, receive (ms)
```
Wrap all HTTP calls in error handling. If the call times out, log the failure and move on rather than halting the loop.

---

### Corporate Proxy Blocks Outbound HTTP

**What goes wrong:** On corporate networks, outbound HTTPS may be routed through a proxy. `Msxml2.XMLHTTP` does not automatically pick up the system proxy. Requests silently fail or return proxy error pages instead of JSON.

**Why it happens:** `Msxml2.XMLHTTP` uses WinInet (browser proxy settings). `Msxml2.ServerXMLHTTP` uses WinHTTP (system proxy, configured separately via `netsh winhttp`). These are different proxy stacks.

**Prevention:**
- Test on the target machine from Excel before distributing.
- If proxies are a concern, use `Msxml2.XMLHTTP` (WinInet-based, inherits browser proxy) rather than ServerXMLHTTP.
- Alternatively: `WinHttpRequest` object with explicit `.SetProxy` call.

---

### UTF-8 Response Decoding

**What goes wrong:** Nominatim returns JSON encoded as UTF-8. VBA strings are internally UTF-16. If the response contains accented characters (common in company names, international cities), the string will appear garbled when written to the sheet or embedded into the HTML file.

**Why it happens:** `XMLHTTP.responseText` does attempt UTF-8 decoding, but edge cases around character set detection exist. Company names with accents, umlauts, or non-Latin characters are at risk.

**Prevention:**
- Read response as `xhr.responseText` (which MSXML decodes from UTF-8 to VBA's internal unicode).
- When writing the HTML file, use `ADODB.Stream` with `Charset = "utf-8"` rather than VBA's native file I/O (`Open filename For Output`), which defaults to ANSI.
- Test with a company name containing characters like `é`, `ü`, `ñ` before shipping.

---

### JSON Injection via Company Name

**What goes wrong:** Company names containing double quotes, backslashes, or angle brackets break the JSON literal you embed in the HTML file.

**Example:** A company named `Smith & "Associates"` generates malformed JavaScript: `name: "Smith & "Associates""`.

**Prevention:**
- Write a dedicated JSON-escape function in VBA that replaces `\` with `\\`, `"` with `\"`, and newlines with `\n` before embedding any string into the HTML output.
- Also HTML-escape `&`, `<`, `>` in popup content to prevent broken HTML rendering.

---

## HTML/Browser Pitfalls

### CRITICAL: Tile Layers Do Not Load from file:// Protocol

**What goes wrong:** Modern browsers (Chrome, Edge, Firefox) apply strict CORS rules to `file://` origins. Fetching external tile URLs (OpenStreetMap, CartoDB, etc.) from a `file://` document generates CORS errors, and tiles fail to load — leaving a grey map with markers but no basemap.

**Why it happens:** `file://` origins are treated as `null` origin by browsers. Tile servers return their tiles without the `Access-Control-Allow-Origin: *` header in all cases, and some CDN tile configurations fail on null origins.

**Current status:** As of 2025, this remains a real issue in Chrome and Edge. Firefox is more permissive about file:// cross-origin requests.

**Prevention — ordered by preference:**
1. **Best:** Open the HTML via a local HTTP server. VBA can launch one: `Shell "python -m http.server 8080 --directory """ & outputPath & """"` then `Shell "start http://localhost:8080/map.html"`. Requires Python on the target machine.
2. **Acceptable:** Use a tile provider that explicitly supports CORS on null origins, or use a CDN that does so.
3. **Fallback:** Embed a minimal tile workaround — use Leaflet's `L.tileLayer` with `crossOrigin: true`, which sometimes resolves the issue with certain tile providers.
4. **Avoid:** Do not assume it will just work — test on Chrome specifically, which is strictest.

---

### Windows File Path Issues

**What goes wrong:** `Shell "start " & filePath` fails if the file path contains spaces. The generated HTML file may not open if the path is not quoted.

**Prevention:**
```vba
Shell "cmd /c start """" """ & filePath & """"
```
Double-quoting in the `Shell` call. Also test with OneDrive paths, which often contain spaces and sometimes non-ASCII characters.

---

### OneDrive Sync Timing

**What goes wrong:** If the output HTML is written to an OneDrive-synced folder, the file may not be locally available immediately. `Shell "start"` on it can open a stale cached version or fail to find the file briefly after write.

**Prevention:** Write to a local temp path (e.g., `Environ("TEMP") & "\geoviz_map.html"`) rather than a synced folder. Open from temp, not from OneDrive.

---

### Browser Opens Wrong File

**What goes wrong:** Clicking "Generate Map" twice creates a new file and opens a browser tab. The previous tab still shows old data and is not refreshed. Users may be confused about which tab is current.

**Prevention:** Always write to the same fixed filename. The browser will reload it on next open rather than accumulating stale tabs.

---

## Leaflet Pitfalls

### CDN Dependency Requires Internet

**What goes wrong:** Standard Leaflet initialization references `https://unpkg.com/leaflet@...` or `https://cdnjs.cloudflare.com/...` for JS and CSS. If the user's machine is offline or the CDN is unavailable, the map renders as a blank page with broken layout.

**Why it happens:** Standalone HTML files commonly load Leaflet from CDN for simplicity. File:// documents can still fetch external URLs; it is CORS that blocks tiles, not CDN JS/CSS.

**Prevention — choose one:**
- **Inline Leaflet:** At build time, fetch the Leaflet JS and CSS, minify, and embed directly in the HTML. VBA can do this once. Eliminates all CDN dependency.
- **Lock a specific version:** Reference `leaflet@1.9.4` (or current stable) explicitly, not `latest`, so a CDN update does not break the map.
- **Bundle locally:** Copy `leaflet.js` and `leaflet.css` to the same output folder and reference with relative paths.

For a single-user internal tool, inlining is simplest and most reliable.

---

### Map Container Sizing Issues

**What goes wrong:** The map renders as zero height, showing nothing. This happens when the `<div id="map">` has no explicit height, or its parent has `height: 100%` but the `<html>` and `<body>` elements do not.

**Prevention:**
```css
html, body { height: 100%; margin: 0; padding: 0; }
#map { height: 100%; width: 100%; }
```
Or use an explicit pixel height: `#map { height: 600px; }`. The explicit pixel approach is more reliable in standalone files.

---

### Popup Content Overflow

**What goes wrong:** Company popups with many fields (name, address, phone, notes) overflow the default popup size. Leaflet's default max popup width is 300px. Long company names or addresses get clipped or cause layout breaks.

**Prevention:**
```javascript
L.marker([lat, lng]).bindPopup(content, { maxWidth: 400, maxHeight: 300 });
```
Also test popups at the edge of the map — Leaflet auto-pans to fit popups in view, but if the popup is taller than the map, it pans infinitely.

---

### Marker Z-Index Conflicts at Same Location

**What goes wrong:** Two companies at the exact same address (e.g., shared office building) stack markers directly on top of each other. Only one is clickable. No indication that multiple companies share the location.

**Prevention:**
- Check for duplicate coordinates after geocoding. If duplicates exist, slightly offset markers (e.g., add ±0.0001 degrees jitter) or use marker clustering.
- `Leaflet.markercluster` handles this gracefully at scale; for 100 companies it may be overkill but worth considering.

---

### Leaflet Map Initialization Order

**What goes wrong:** Calling `L.map('map')` before the DOM element exists throws a silent error and the map never renders.

**Prevention:** Place all Leaflet initialization code in a `<script>` tag at the bottom of `<body>`, after the `<div id="map">`, or wrap in `document.addEventListener('DOMContentLoaded', ...)`.

---

## Excel/Macro Pitfalls

### CRITICAL: Mark of the Web Blocks Macro Execution

**What goes wrong:** Since 2022, Microsoft blocks VBA macros in files downloaded from the internet or received by email. The file is tagged with a "Mark of the Web" (MOTW) zone identifier. The user sees "Microsoft has blocked macros from running because the source of this file is untrusted" with no obvious fix.

**Why it happens:** Files from the internet or email attachments carry NTFS zone ID metadata. Excel now blocks macros in these files by default, regardless of Trust Center settings.

**Consequences:** The tool does not work at all for first-time users who received the file by email or download.

**Prevention:**
- Distribute the .xlsm file with instructions to right-click > Properties > check "Unblock" before opening.
- Alternatively, place the file in a trusted location (Trust Center > Trusted Locations) on the target machine.
- For broader distribution, sign the macro with a code-signing certificate and instruct users to trust the publisher.
- Document this step prominently in any user guide.

---

### .xlsm Extension Requirement

**What goes wrong:** Saving an Excel file with macros as .xlsx silently strips all VBA code. User opens the file, macros are gone, buttons do nothing.

**Prevention:** Always save as .xlsm (Excel Macro-Enabled Workbook). Add a check on workbook open that verifies the expected macros exist.

---

### Late Binding vs. Early Binding References

**What goes wrong:** If code uses early binding (e.g., `Dim xhr As MSXML2.XMLHTTP60`) and the MSXML 6.0 library is not registered on the target machine, the workbook opens with a reference error and all macros fail.

**Why it happens:** Library availability varies between Office versions and Windows configurations. Office 365 generally includes MSXML 6.0, but older installs or stripped-down environments may not.

**Prevention:**
- Use late binding: `Dim xhr As Object: Set xhr = CreateObject("Msxml2.XMLHTTP.6.0")`.
- Late binding resolves at runtime and gives a clear error message rather than a compile-time failure that breaks all macros.

---

### Worksheet Name Changes Break Code

**What goes wrong:** VBA code that references sheets by name (`Sheets("GeoCache")`) breaks silently if a user renames the sheet.

**Prevention:** Reference sheets by CodeName (the internal VBA name set in the Properties window, e.g., `Sheet1.Range(...)`) rather than tab display name. CodeNames are not visible to users and cannot be changed from the sheet tab.

---

## Data Quality Pitfalls

### Missing Addresses Produce Null Islands

**What goes wrong:** If the address column is blank and no guard is in place, Nominatim may return a result for an empty query (sometimes the center of a country or a zero coordinate), and a marker appears at (0, 0) — "null island" in the ocean.

**Prevention:**
- Skip geocoding if the address cell is empty. Log the company name as "skipped — no address."
- After geocoding, validate that lat is between -90 and 90, and lon between -180 and 180. Reject (0, 0) explicitly.

---

### Encoding Problems in Company Names

**What goes wrong:** Company names imported from external sources (CSV, copy-paste) may carry Windows-1252 encoding that looks fine in Excel but produces garbled characters when embedded in the HTML file.

**Why it happens:** Excel internally uses UTF-16 but displays many encodings correctly. Writing to a file with ANSI or without proper charset specification breaks non-ASCII characters.

**Prevention:**
- Always write the HTML output file using `ADODB.Stream` with `Charset = "utf-8"`.
- Test with company names containing `&`, `'`, `<`, `>`, `"`, and common accented characters.

---

### Address Data Inconsistency

**What goes wrong:** Addresses in different formats across rows. Some have postal codes, some do not. Some use abbreviations ("St" vs "Street"), some have suite numbers in unexpected formats ("Ste 200" vs "#200" vs "Suite 200").

**Consequences:** Nominatim matches at different precision levels. Some companies geocode to rooftop level, others to city-block or city level. The map looks uniform but accuracy is inconsistent.

**Prevention:**
- Standardize address format in the Excel sheet before geocoding.
- Log the Nominatim `display_name` in the cache to allow post-geocoding audit.
- Consider a pre-flight data quality check that flags rows with obviously incomplete addresses.

---

### Special Characters Break JavaScript in HTML Output

**What goes wrong:** Company names or addresses with `"`, `\`, or `<script>` sequences break the JavaScript data array embedded in the HTML file. The map either fails to load or opens a blank page with a console error.

**Prevention:**
Write a VBA escape function:
```vba
Function JsonEscape(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(13), "\r")
    s = Replace(s, Chr(10), "\n")
    JsonEscape = s
End Function
```
Apply this to every string field (name, address, notes) before embedding in the HTML output.

---

## Prevention Strategies

### Geocoding Safeguards

| Risk | Strategy |
|------|----------|
| Rate limit exceeded | `Application.Wait` 1 second between requests; check HTTP status |
| Missing User-Agent | Always set a custom `User-Agent` header identifying the app |
| Bad geocode result | Log `display_name` from response; validate coordinate bounds |
| Overwriting good data | Check "Geocoded" flag; only geocode empty rows |
| Silent HTTP failure | Wrap all HTTP calls in error handling; log failures to a separate sheet |

### HTML Generation Safeguards

| Risk | Strategy |
|------|----------|
| Garbled characters | Use `ADODB.Stream` with `Charset = "utf-8"` for file output |
| JSON injection | Apply JSON-escape function to all string fields |
| Tile CORS failure | Test on Chrome; use local HTTP server or inline Leaflet |
| Path with spaces | Double-quote file paths in `Shell` commands |
| OneDrive sync delay | Write output to `Environ("TEMP")`, not synced folder |

### VBA/Excel Safeguards

| Risk | Strategy |
|------|----------|
| Macro blocked by MOTW | Document "Unblock" step; use trusted location |
| Sheet name change | Reference sheets by CodeName, not display name |
| Missing library reference | Use late binding for all COM objects |
| .xlsm saved as .xlsx | Document save format requirement; add open-time check |

### Proximity Calculation

**Flat-earth vs. Haversine:** For distances under ~20 km, flat-earth approximation is acceptable (errors < 1%). For distances up to 100 km, flat-earth error grows to 1-3%. For a 100-company dataset where proximity comparisons may span regional distances, **use haversine**. It is only a few lines of VBA and eliminates the question entirely.

```vba
Function Haversine(lat1 As Double, lon1 As Double, lat2 As Double, lon2 As Double) As Double
    Const R As Double = 6371  ' Earth radius km
    Dim dLat As Double, dLon As Double, a As Double
    dLat = (lat2 - lat1) * Application.Pi() / 180
    dLon = (lon2 - lon1) * Application.Pi() / 180
    a = Sin(dLat / 2) ^ 2 + Cos(lat1 * Application.Pi() / 180) * _
        Cos(lat2 * Application.Pi() / 180) * Sin(dLon / 2) ^ 2
    Haversine = R * 2 * Atn(Sqr(a) / Sqr(1 - a))
End Function
```

### Testing Checklist Before Shipping

- [ ] Test geocoding with a blank address row — does it skip gracefully?
- [ ] Test with a company name containing `"`, `&`, `é`, `\`
- [ ] Test HTML open in Chrome (strictest CORS) — do tiles load?
- [ ] Test with file path containing spaces
- [ ] Confirm macro runs after downloading and unblocking the file on a fresh machine
- [ ] Verify 1-second delay between geocoding calls is enforced
- [ ] Confirm lat/lon cache is not overwritten for already-geocoded rows

---

**Sources:**
- Nominatim Usage Policy: https://operations.osmfoundation.org/policies/nominatim/
- Nominatim rate limit enforcement: https://community.openstreetmap.org/t/help-nominatim-returns-usage-limit-reached/103102
- Microsoft Macro Blocking (MOTW): https://learn.microsoft.com/en-us/microsoft-365-apps/security/internet-macros-blocked
- VBA UTF-8 file writing: https://www.codestudy.net/blog/save-text-file-utf-8-encoded-with-vba/
- VBA XMLHTTP SSL options: https://www.normanbauer.com/2011/02/10/certificate-problems-with-vbscript-and-xml-http-calls/
- Haversine formula reference: https://www.movable-type.co.uk/scripts/latlong.html
- Leaflet CORS issue tracker: https://github.com/Leaflet/Leaflet/issues/5692
- localStorage on file:// protocol: https://www.xjavascript.com/blog/does-localstorage-in-firefox-only-work-when-the-page-is-online/
- VBA-JSON escaping issues: https://github.com/VBA-tools/VBA-JSON/issues/37
