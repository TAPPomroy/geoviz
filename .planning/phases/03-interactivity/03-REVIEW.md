---
phase: 03-interactivity
reviewed: 2026-07-19T00:00:00Z
depth: standard
files_reviewed: 1
files_reviewed_list:
  - src/mod_MapBuilder.bas
findings:
  critical: 2
  warning: 3
  info: 1
  total: 6
status: issues_found
---

# Phase 03: Code Review Report

**Reviewed:** 2026-07-19
**Depth:** standard
**Files Reviewed:** 1
**Status:** issues_found

## Summary

`mod_MapBuilder.bas` generates a standalone Leaflet HTML map with marker clustering, click-to-select proximity highlighting, a radius slider, a color-by-field dropdown, and a toggleable legend. The VBA structure and ADODB.Stream file-writing are sound. Two blockers were found: a logic bug that causes the radius slider to clear the current selection instead of updating the radius circle, and an XSS vulnerability from raw HTML injection of company data values into the popup. Three additional warnings cover related XSS surfaces and incomplete onclick escaping in the legend.

## Critical Issues

### CR-01: Radius slider clears selection instead of updating radius

**File:** `src/mod_MapBuilder.bas:306`

**Issue:** The slider `input` event handler calls `applySelection(selectedMarker, selectedCompany)` to refresh the radius after the user drags the slider. However, `applySelection` begins with a toggle-off guard: `if (selectedMarker === marker) { clearSelection(); return; }`. When the slider fires, `marker` receives the value of `selectedMarker`, so the condition is always true — the function immediately calls `clearSelection()` and returns without updating the radius circle or neighbour highlights. Every slider drag clears the selection rather than refreshing it.

**Fix:** Replace the slider callback's call to `applySelection` with a dedicated helper that performs only the radius update, bypassing the toggle-off guard:

```javascript
// Add a dedicated helper (generated inside the <script> block)
function refreshRadius() {
  if (!selectedMarker) return;
  if (radiusCircle) { map.removeLayer(radiusCircle); radiusCircle = null; }
  var sliderVal = parseFloat(document.getElementById('radiusSlider').value);
  var radiusMeters = sliderVal * 1609.34;
  radiusCircle = L.circle(selectedMarker.getLatLng(),
    { radius: radiusMeters, color: '#0078ff', weight: 1, fillOpacity: 0.05, interactive: false }
  ).addTo(map);
  markers.forEach(function(m, i) {
    if (m === selectedMarker) return;
    if (hiddenValues[String(companies[i][currentField])]) return;
    var dist = haversine(selectedCompany.Lat, selectedCompany.Lon, companies[i].Lat, companies[i].Lon);
    if (dist <= sliderVal) {
      m.setRadius(8);
      m.setStyle({ fillColor: getFieldColor(companies[i][currentField]),
                   color: '#ffffff',
                   weight: String(companies[i][currentField]) === String(selectedCompany[currentField]) ? 3 : 1,
                   fillOpacity: 0.85 });
    } else {
      m.setRadius(8);
      m.setStyle({ fillOpacity: 0.15 });
    }
  });
}

// Slider handler becomes:
document.getElementById('radiusSlider').addEventListener('input', function() {
  document.getElementById('radiusLabel').textContent = this.value;
  if (selectedMarker) { refreshRadius(); }
});
```

---

### CR-02: XSS via raw company field values injected into popup HTML

**File:** `src/mod_MapBuilder.bas:126`

**Issue:** `buildPopupHtml` concatenates company property values directly into an HTML string that is set as popup content via `L.popup().setContent(...)`. A company field value containing `<img src=x onerror=alert(document.cookie)>` or any HTML/script would execute in the browser. Although the data currently comes from a trusted Excel sheet, any future import mechanism (CSV paste, bulk load) widens this surface, and the generated HTML file persists on disk.

**Fix:** Add a minimal HTML-escaping function and use it for every key and value in the popup:

```javascript
function escHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function buildPopupHtml(company) {
  var POPUP_EXCLUDE = ['Lat','Lon'];
  var rows = '';
  Object.keys(company).forEach(function(k) {
    if (POPUP_EXCLUDE.indexOf(k) === -1) {
      rows += '<tr><td style="font-weight:bold;padding-right:8px">' + escHtml(k)
            + '</td><td>' + escHtml(company[k]) + '</td></tr>';
    }
  });
  return '<table style="border-collapse:collapse;font-size:13px">' + rows + '</table>';
}
```

Apply `escHtml` to every user-controlled value written into `innerHTML` (see WR-01 and WR-02 below as well).

---

## Warnings

### WR-01: XSS in `buildLegend` — field values inserted into `innerHTML` unescaped

**File:** `src/mod_MapBuilder.bas:215`

**Issue:** The legend HTML string is built by concatenating `v` (a field value from company data) directly into the `<span>` content and into the `onclick` attribute string. HTML-special characters in `v` (e.g., `<`, `>`, `&`) break the surrounding markup. A value like `AT&T` renders as `AT` due to the unescaped `&`, and a value like `</span><script>…</script>` injects HTML.

**Fix:** Apply `escHtml` (defined in CR-02 fix) when emitting the display text, and use `encodeURIComponent` or data attributes instead of an inline onclick string for the toggle handler:

```javascript
// Safer pattern: use a data attribute and a delegated listener instead of inline onclick
html += '<span data-val="' + escHtml(v) + '" class="legend-item" style="cursor:pointer;' + strike + '">'
      + escHtml(v) + '</span><br>';
```

Then bind a single delegated click listener on `legendDiv` that reads `dataset.val`.

---

### WR-02: Incomplete escaping in `toggleValue` onclick — backslash and `</span>` bypass

**File:** `src/mod_MapBuilder.bas:214`

**Issue:** The inline `onclick` handler is constructed with `v.replace(/'/g, "\\'")`. This escapes single quotes but does not handle:
- Backslashes in `v` (e.g., `UPC\Broadband` → `toggleValue('UPC\Broadband')` — `\B` is not an escape sequence but `\n`, `\r`, `\\` would corrupt the call)
- Closing `</span>` in `v` which terminates the surrounding tag and can inject new HTML before `onclick` is reached

**Fix:** Replace inline onclick with a `data-val` attribute approach as described in WR-01. This eliminates the need for any ad-hoc string escaping inside event handler strings.

---

### WR-03: Filter dropdown option values not HTML-escaped

**File:** `src/mod_MapBuilder.bas:292`

**Issue:** Column header names from the company data are embedded directly into `<option value="...">` strings without HTML escaping. A column named `foo"onmouseover="alert(1)` would break the attribute boundary. While column names are controlled by the workbook author, this is still a correctness issue if headers contain `&`, `<`, `>`, or `"`.

**Fix:**
```javascript
// Use escHtml() on both the value attribute and the display text
return '<option value="' + escHtml(f) + '"' + (f === 'ISP' ? ' selected' : '') + '>' + escHtml(f) + '</option>';
```

---

## Info

### IN-01: Magic number 1609.34 (miles-to-meters) appears in two separate places

**File:** `src/mod_MapBuilder.bas:186` (and the refreshRadius helper introduced by CR-01 fix)

**Issue:** The miles-to-meters conversion factor `1609.34` is hard-coded inline in `applySelection` at line 186. When `refreshRadius()` is extracted as per the CR-01 fix, this constant appears a second time. A typo in either copy would silently produce a wrong radius circle while the correct distance comparison still uses `sliderVal` directly (haversine returns miles, comparison is done in miles).

**Fix:** Define a named constant at the top of the `<script>` block:
```javascript
var METERS_PER_MILE = 1609.34;
```
Use `METERS_PER_MILE` everywhere the conversion appears.

---

_Reviewed: 2026-07-19_
_Reviewer: Claude (gsd-code-reviewer)_
_Depth: standard_
