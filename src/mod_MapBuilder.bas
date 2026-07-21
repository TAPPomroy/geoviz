Attribute VB_Name = "mod_MapBuilder"
Option Explicit

' Appends a line of text (with trailing newline) to a string buffer.
Private Sub AppendLine(ByRef sb As String, ByVal line As String)
    sb = sb & line & vbLf
End Sub

' Writes content to filePath as UTF-8 text via ADODB.Stream.
' Never use FSO here -- it defaults to ANSI and corrupts non-ASCII characters.
Public Sub WriteUtf8File(ByVal filePath As String, ByVal content As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2          ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText content
    stream.Position = 0             ' rewind before SaveToFile -- required by ADODB.Stream
    stream.SaveToFile filePath, 2   ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing
End Sub

' Returns a complete standalone HTML document as a String.
' jsonStr must be a JSON array string (output of mod_JsonBuilder.BuildCompanyJson).
' JsonEscape in mod_JsonBuilder already converts "/" to "\/" so "</script>" is safe.
Public Function BuildMapHtml(ByVal jsonStr As String) As String
    Dim sb As String

    ' --- DOCTYPE and head ---
    AppendLine sb, "<!DOCTYPE html>"
    AppendLine sb, "<html><head><meta charset=""utf-8""><title>GeoViz Map</title>"
    AppendLine sb, "<style>body{margin:0;padding:0;} #map{height:100vh;width:100%;}</style>"

    ' CSS load order: Leaflet CSS first, then both MarkerCluster CSS files (Pitfall 1 & 2)
    AppendLine sb, "<link rel=""stylesheet"" href=""https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"" crossorigin="""" integrity=""sha256-p4NxAoJBhIIN+hmNHrzRCf9tD/miZyoHS5obTRR9BMY=""/>"
    AppendLine sb, "<link rel=""stylesheet"" href=""https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css""/>"
    AppendLine sb, "<link rel=""stylesheet"" href=""https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css""/>"

    ' JS load order: Leaflet JS first, then MarkerCluster JS (Pitfall 2)
    AppendLine sb, "<script src=""https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"" crossorigin="""" integrity=""sha256-20nQCchB9co0qIjJZRGuk2/Z9VM+kNiyxNV1lvTlZBo=""></script>"
    AppendLine sb, "<script src=""https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js""></script>"

    AppendLine sb, "</head><body><div id=""map""></div>"
    AppendLine sb, "<script>"

    ' Embedded company data (jsonStr already JSON-escaped by mod_JsonBuilder)
    AppendLine sb, "var companies = " & jsonStr & ";"

    ' ColorBrewer Set1 palette -- 7 high-contrast colors, colorblind-friendly (D-03, D-04)
    AppendLine sb, "var ISP_COLORS = [""#e41a1c"",""#377eb8"",""#4daf4a"",""#984ea3"",""#ff7f00"",""#a65628"",""#f781bf""];"
    AppendLine sb, "var OTHER_COLOR = ""#999999"";"

    ' Deterministic ISP->color map via alphabetical sort (D-05) -- used for initial marker creation
    AppendLine sb, "var ispNames = [];"
    AppendLine sb, "companies.forEach(function(c){ if(ispNames.indexOf(c.ISP)===-1) ispNames.push(c.ISP); });"
    AppendLine sb, "ispNames.sort();"
    AppendLine sb, "var ispColorMap = {};"
    AppendLine sb, "ispNames.forEach(function(isp,i){ ispColorMap[isp] = i < ISP_COLORS.length ? ISP_COLORS[i] : OTHER_COLOR; });"
    AppendLine sb, "function getColor(isp){ return ispColorMap[isp] || OTHER_COLOR; }"

    ' Leaflet map -- US center/zoom4 as fallback before fitBounds
    AppendLine sb, "var map = L.map('map').setView([39.5,-98.35],4);"

    ' CartoDB Positron basemap (MAP-04) -- attribution required by CartoDB license
    AppendLine sb, "L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png',{"
    AppendLine sb, "  attribution:'&copy; <a href=""http://www.openstreetmap.org/copyright"">OpenStreetMap</a> contributors &copy; <a href=""https://carto.com/attributions"">CARTO</a>',"
    AppendLine sb, "  subdomains:'abcd',maxZoom:20}).addTo(map);"

    ' MarkerCluster group + parallel markers array (D-06, D-07; markers[] needed by Phase 3)
    AppendLine sb, "var clusterGroup = L.markerClusterGroup();"
    AppendLine sb, "var markers = [];"
    AppendLine sb, "companies.forEach(function(c){"
    AppendLine sb, "  var marker = L.circleMarker([c.Lat,c.Lon],{"
    AppendLine sb, "    radius:8, fillColor:getColor(c.ISP), color:""#ffffff"", weight:1, fillOpacity:0.85"
    AppendLine sb, "  });"
    AppendLine sb, "  markers.push(marker);"
    AppendLine sb, "  clusterGroup.addLayer(marker);"
    AppendLine sb, "});"
    AppendLine sb, "map.addLayer(clusterGroup);"

    ' Fit bounds only when there are markers (guard against empty dataset)
    AppendLine sb, "if(markers.length > 0){ map.fitBounds(clusterGroup.getBounds(),{padding:[30,30]}); }"

    ' Dynamic legend in bottom-right corner (D-08). legendDiv stores the div reference
    ' so buildLegend() can update innerHTML without re-querying the DOM.
    AppendLine sb, "var legendDiv = null;"
    AppendLine sb, "var legend = L.control({position:'bottomright'});"
    AppendLine sb, "legend.onAdd = function(map){"
    AppendLine sb, "  var div = L.DomUtil.create('div');"
    AppendLine sb, "  div.style.cssText = 'background:white;padding:8px 12px;border-radius:4px;line-height:1.8;font-size:13px;box-shadow:0 1px 5px rgba(0,0,0,.3);max-height:60vh;overflow-y:auto';"
    AppendLine sb, "  legendDiv = div;"
    AppendLine sb, "  return div;"
    AppendLine sb, "};"
    AppendLine sb, "legend.addTo(map);"
    ' Delegated click handler on legendDiv -- avoids per-item inline onclick (WR-01, WR-02).
    ' legendDiv is guaranteed to exist after legend.addTo(map) runs onAdd.
    AppendLine sb, "legendDiv.addEventListener('click', function(e) {"
    AppendLine sb, "  var el = e.target.closest('.legend-item');"
    AppendLine sb, "  if (el && el.dataset.val !== undefined) { toggleValue(el.dataset.val); }"
    AppendLine sb, "});"
    AppendLine sb, ""

    ' --- State variables (Plan 01 + Plan 02) ---
    AppendLine sb, "var selectedMarker = null;"
    AppendLine sb, "var selectedCompany = null;"
    AppendLine sb, "var selectedISP = null;"
    AppendLine sb, "var radiusCircle = null;"
    AppendLine sb, "var currentField = 'ISP';"
    AppendLine sb, "var fieldColorMap = {};"
    AppendLine sb, "var fieldNames = [];"
    AppendLine sb, "var hiddenValues = {};"
    AppendLine sb, ""

    ' --- Helper: haversine distance in miles ---
    AppendLine sb, "function haversine(lat1, lon1, lat2, lon2) {"
    AppendLine sb, "  var R = 3958.8;"
    AppendLine sb, "  var dLat = (lat2-lat1)*Math.PI/180;"
    AppendLine sb, "  var dLon = (lon2-lon1)*Math.PI/180;"
    AppendLine sb, "  var a = Math.sin(dLat/2)*Math.sin(dLat/2) +"
    AppendLine sb, "          Math.cos(lat1*Math.PI/180)*Math.cos(lat2*Math.PI/180)*"
    AppendLine sb, "          Math.sin(dLon/2)*Math.sin(dLon/2);"
    AppendLine sb, "  return R*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));"
    AppendLine sb, "}"
    AppendLine sb, ""

    ' --- Helper: HTML-escape to prevent XSS (CR-02) ---
    AppendLine sb, "function escHtml(s) {"
    AppendLine sb, "  return String(s)"
    AppendLine sb, "    .replace(/&/g, '&amp;')"
    AppendLine sb, "    .replace(/</g, '&lt;')"
    AppendLine sb, "    .replace(/>/g, '&gt;')"
    AppendLine sb, "    .replace(/""/g, '&quot;')"
    AppendLine sb, "    .replace(/'/g, '&#39;');"
    AppendLine sb, "}"
    AppendLine sb, ""

    ' --- Helper: popup HTML table ---
    AppendLine sb, "function buildPopupHtml(company) {"
    AppendLine sb, "  var POPUP_EXCLUDE = ['Lat','Lon'];"
    AppendLine sb, "  var rows = '';"
    AppendLine sb, "  Object.keys(company).forEach(function(k) {"
    AppendLine sb, "    if (POPUP_EXCLUDE.indexOf(k) === -1) {"
    AppendLine sb, "      rows += '<tr><td style=""font-weight:bold;padding-right:8px"">' + escHtml(k) + '</td><td>' + escHtml(company[k]) + '</td></tr>';"
    AppendLine sb, "    }"
    AppendLine sb, "  });"
    AppendLine sb, "  return '<table style=""border-collapse:collapse;font-size:13px"">' + rows + '</table>';"
    AppendLine sb, "}"
    AppendLine sb, ""

    ' --- Helper: radius hint visibility ---
    AppendLine sb, "function setHint(show) {"
    AppendLine sb, "  var h = document.getElementById('radiusHint');"
    AppendLine sb, "  if (h) h.style.display = show ? 'block' : 'none';"
    AppendLine sb, "}"
    AppendLine sb, ""

    ' --- Plan 02: buildColorMap and getFieldColor (defined before clearSelection) ---
    AppendLine sb, "function buildColorMap(field) {"
    AppendLine sb, "  fieldNames = [];"
    AppendLine sb, "  fieldColorMap = {};"
    AppendLine sb, "  companies.forEach(function(c) {"
    AppendLine sb, "    var v = c[field] !== undefined ? String(c[field]) : '';"
    AppendLine sb, "    if (fieldNames.indexOf(v) === -1) { fieldNames.push(v); }"
    AppendLine sb, "  });"
    AppendLine sb, "  fieldNames.sort();"
    AppendLine sb, "  fieldNames.forEach(function(v, i) {"
    AppendLine sb, "    fieldColorMap[v] = i < ISP_COLORS.length ? ISP_COLORS[i] : OTHER_COLOR;"
    AppendLine sb, "  });"
    AppendLine sb, "}"
    AppendLine sb, ""
    AppendLine sb, "function getFieldColor(value) {"
    AppendLine sb, "  return fieldColorMap[String(value)] || OTHER_COLOR;"
    AppendLine sb, "}"
    AppendLine sb, ""

    ' --- clearSelection: uses getFieldColor(companies[i][currentField]) (Plan 02 update) ---
    AppendLine sb, "function clearSelection() {"
    AppendLine sb, "  if (radiusCircle) { map.removeLayer(radiusCircle); radiusCircle = null; }"
    AppendLine sb, "  markers.forEach(function(m, i) {"
    AppendLine sb, "    m.setRadius(8);"
    AppendLine sb, "    m.setStyle({ fillColor: getFieldColor(companies[i][currentField]), color: '#ffffff', weight: 1, fillOpacity: 0.85 });"
    AppendLine sb, "  });"
    AppendLine sb, "  map.closePopup();"
    AppendLine sb, "  selectedMarker = null;"
    AppendLine sb, "  selectedCompany = null;"
    AppendLine sb, "  selectedISP = null;"
    AppendLine sb, "  setHint(true);"
    AppendLine sb, "}"
    AppendLine sb, ""

    ' --- applySelection: uses getFieldColor and currentField throughout (Plan 02 update) ---
    AppendLine sb, "function applySelection(marker, company) {"
    AppendLine sb, "  if (selectedMarker === marker) { clearSelection(); return; }"
    AppendLine sb, "  clearSelection();"
    AppendLine sb, "  setHint(false);"
    AppendLine sb, "  selectedMarker = marker;"
    AppendLine sb, "  selectedCompany = company;"
    AppendLine sb, "  selectedISP = company.ISP;"
    AppendLine sb, "  L.popup().setContent(buildPopupHtml(company)).setLatLng(marker.getLatLng()).openOn(map);"
    AppendLine sb, "  marker.setRadius(12);"
    AppendLine sb, "  marker.setStyle({ fillColor: '#ffffff', color: getFieldColor(company[currentField]), weight: 2, fillOpacity: 0.85 });"
    AppendLine sb, "  var sliderVal = parseFloat(document.getElementById('radiusSlider').value);"
    AppendLine sb, "  var radiusMeters = sliderVal * 1609.34;"
    AppendLine sb, "  radiusCircle = L.circle(marker.getLatLng(), { radius: radiusMeters, color: '#0078ff', weight: 1, fillOpacity: 0.05, interactive: false }).addTo(map);"
    AppendLine sb, "  markers.forEach(function(m, i) {"
    AppendLine sb, "    if (m === marker) return;"
    AppendLine sb, "    if (hiddenValues[String(companies[i][currentField])]) return;"
    AppendLine sb, "    var dist = haversine(company.Lat, company.Lon, companies[i].Lat, companies[i].Lon);"
    AppendLine sb, "    if (dist <= sliderVal) {"
    AppendLine sb, "      m.setRadius(8);"
    AppendLine sb, "      m.setStyle({ fillColor: getFieldColor(companies[i][currentField]), color: '#ffffff', weight: String(companies[i][currentField]) === String(company[currentField]) ? 3 : 1, fillOpacity: 0.85 });"
    AppendLine sb, "    } else {"
    AppendLine sb, "      m.setRadius(8);"
    AppendLine sb, "      m.setStyle({ fillOpacity: 0.15 });"
    AppendLine sb, "    }"
    AppendLine sb, "  });"
    AppendLine sb, "}"
    AppendLine sb, ""

    ' --- Plan 02 Task 1: buildLegend and toggleValue ---
    AppendLine sb, "function buildLegend() {"
    AppendLine sb, "  if (!legendDiv) return;"
    AppendLine sb, "  var html = '<strong>' + currentField + '</strong><br>';"
    AppendLine sb, "  fieldNames.forEach(function(v) {"
    AppendLine sb, "    var color = fieldColorMap[v] || OTHER_COLOR;"
    AppendLine sb, "    var hidden = hiddenValues[v] ? true : false;"
    AppendLine sb, "    var opacity = hidden ? '0.4' : '1';"
    AppendLine sb, "    var strike = hidden ? 'text-decoration:line-through;' : '';"
    AppendLine sb, "    html += '<span style=""display:inline-block;width:12px;height:12px;background:' + color"
    AppendLine sb, "          + ';margin-right:6px;border-radius:2px;vertical-align:middle;opacity:' + opacity + '""></span>'"
    AppendLine sb, "          + '<span data-val=""' + escHtml(v) + '"" class=""legend-item"" style=""cursor:pointer;' + strike + '"">' + escHtml(v) + '</span><br>';"
    AppendLine sb, "  });"
    AppendLine sb, "  if (fieldNames.length > ISP_COLORS.length) {"
    AppendLine sb, "    html += '<span style=""display:inline-block;width:12px;height:12px;background:#999999;margin-right:6px;border-radius:2px;vertical-align:middle""></span>Other<br>';"
    AppendLine sb, "  }"
    AppendLine sb, "  legendDiv.innerHTML = html;"
    AppendLine sb, "}"
    AppendLine sb, ""
    AppendLine sb, "function toggleValue(val) {"
    AppendLine sb, "  if (hiddenValues[val]) {"
    AppendLine sb, "    delete hiddenValues[val];"
    AppendLine sb, "    markers.forEach(function(m, i) {"
    AppendLine sb, "      if (String(companies[i][currentField]) === val) {"
    AppendLine sb, "        clusterGroup.addLayer(m);"
    AppendLine sb, "      }"
    AppendLine sb, "    });"
    AppendLine sb, "  } else {"
    AppendLine sb, "    hiddenValues[val] = true;"
    AppendLine sb, "    markers.forEach(function(m, i) {"
    AppendLine sb, "      if (String(companies[i][currentField]) === val) {"
    AppendLine sb, "        clusterGroup.removeLayer(m);"
    AppendLine sb, "        if (m === selectedMarker) { clearSelection(); }"
    AppendLine sb, "      }"
    AppendLine sb, "    });"
    AppendLine sb, "  }"
    AppendLine sb, "  buildLegend();"
    AppendLine sb, "}"
    AppendLine sb, ""

    ' --- Plan 02 Task 2: applyFilter ---
    AppendLine sb, "function applyFilter(field) {"
    AppendLine sb, "  clearSelection();"
    AppendLine sb, "  currentField = field;"
    AppendLine sb, "  hiddenValues = {};"
    AppendLine sb, "  buildColorMap(field);"
    AppendLine sb, "  markers.forEach(function(m, i) {"
    AppendLine sb, "    var v = companies[i][field] !== undefined ? String(companies[i][field]) : '';"
    AppendLine sb, "    m.setStyle({ fillColor: getFieldColor(v) });"
    AppendLine sb, "    clusterGroup.addLayer(m);"
    AppendLine sb, "  });"
    AppendLine sb, "  buildLegend();"
    AppendLine sb, "}"
    AppendLine sb, ""

    ' --- Initialize default color map and legend ---
    AppendLine sb, "buildColorMap('ISP');"
    AppendLine sb, "buildLegend();"
    AppendLine sb, ""

    ' --- Radius slider Leaflet control (Plan 01 Task 2) ---
    AppendLine sb, "var radiusControl = L.control({position:'topleft'});"
    AppendLine sb, "radiusControl.onAdd = function(map) {"
    AppendLine sb, "  var div = L.DomUtil.create('div', 'radius-control');"
    AppendLine sb, "  div.style.cssText = 'background:white;padding:8px 12px;border-radius:4px;font-size:13px;box-shadow:0 1px 5px rgba(0,0,0,.3);min-width:160px';"
    AppendLine sb, "  div.innerHTML = '<label style=""display:block;margin-bottom:6px""><strong>Radius: <span id=""radiusLabel"">0.5</span> mi</strong></label>'"
    AppendLine sb, "             + '<input id=""radiusSlider"" type=""range"" min=""0.1"" max=""5"" step=""0.1"" value=""0.5"" style=""width:140px"">'"
    AppendLine sb, "             + '<div id=""radiusHint"" style=""margin-top:6px;font-size:11px;color:#888"">Click a company to see neighbors</div>';"
    AppendLine sb, "  L.DomEvent.disableClickPropagation(div);"
    AppendLine sb, "  L.DomEvent.disableScrollPropagation(div);"
    AppendLine sb, "  return div;"
    AppendLine sb, "};"
    AppendLine sb, "radiusControl.addTo(map);"
    AppendLine sb, ""

    ' --- Plan 02 Task 2: Filter dropdown Leaflet control ---
    AppendLine sb, "var filterControl = L.control({position:'topright'});"
    AppendLine sb, "filterControl.onAdd = function(map) {"
    AppendLine sb, "  var div = L.DomUtil.create('div', 'filter-control');"
    AppendLine sb, "  div.style.cssText = 'background:white;padding:8px 12px;border-radius:4px;font-size:13px;box-shadow:0 1px 5px rgba(0,0,0,.3)';"
    AppendLine sb, "  var SYSTEM_FIELDS = ['Lat','Lon','GeocodedAt','Address'];"
    AppendLine sb, "  var filterFields = [];"
    AppendLine sb, "  if (companies.length > 0) {"
    AppendLine sb, "    Object.keys(companies[0]).forEach(function(k) {"
    AppendLine sb, "      if (SYSTEM_FIELDS.indexOf(k) === -1) { filterFields.push(k); }"
    AppendLine sb, "    });"
    AppendLine sb, "  }"
    AppendLine sb, "  var opts = filterFields.map(function(f) {"
    AppendLine sb, "    return '<option value=""' + escHtml(f) + '""' + (f === 'ISP' ? ' selected' : '') + '>' + escHtml(f) + '</option>';"
    AppendLine sb, "  }).join('');"
    AppendLine sb, "  div.innerHTML = '<label style=""display:block;margin-bottom:4px""><strong>Color by</strong></label>'"
    AppendLine sb, "              + '<select id=""fieldSelect"" style=""font-size:13px;width:100%"">' + opts + '</select>';"
    AppendLine sb, "  L.DomEvent.disableClickPropagation(div);"
    AppendLine sb, "  L.DomEvent.disableScrollPropagation(div);"
    AppendLine sb, "  return div;"
    AppendLine sb, "};"
    AppendLine sb, "filterControl.addTo(map);"
    AppendLine sb, ""

    ' --- refreshRadius: updates radius circle and neighbour highlights WITHOUT toggle-off check ---
    ' Called by the slider so dragging never clears the current selection (CR-01).
    AppendLine sb, "function refreshRadius() {"
    AppendLine sb, "  if (!selectedMarker) return;"
    AppendLine sb, "  if (radiusCircle) { map.removeLayer(radiusCircle); radiusCircle = null; }"
    AppendLine sb, "  var sliderVal = parseFloat(document.getElementById('radiusSlider').value);"
    AppendLine sb, "  var radiusMeters = sliderVal * 1609.34;"
    AppendLine sb, "  radiusCircle = L.circle(selectedMarker.getLatLng(),"
    AppendLine sb, "    { radius: radiusMeters, color: '#0078ff', weight: 1, fillOpacity: 0.05, interactive: false }"
    AppendLine sb, "  ).addTo(map);"
    AppendLine sb, "  markers.forEach(function(m, i) {"
    AppendLine sb, "    if (m === selectedMarker) return;"
    AppendLine sb, "    if (hiddenValues[String(companies[i][currentField])]) return;"
    AppendLine sb, "    var dist = haversine(selectedCompany.Lat, selectedCompany.Lon, companies[i].Lat, companies[i].Lon);"
    AppendLine sb, "    if (dist <= sliderVal) {"
    AppendLine sb, "      m.setRadius(8);"
    AppendLine sb, "      m.setStyle({ fillColor: getFieldColor(companies[i][currentField]),"
    AppendLine sb, "                   color: '#ffffff',"
    AppendLine sb, "                   weight: String(companies[i][currentField]) === String(selectedCompany[currentField]) ? 3 : 1,"
    AppendLine sb, "                   fillOpacity: 0.85 });"
    AppendLine sb, "    } else {"
    AppendLine sb, "      m.setRadius(8);"
    AppendLine sb, "      m.setStyle({ fillOpacity: 0.15 });"
    AppendLine sb, "    }"
    AppendLine sb, "  });"
    AppendLine sb, "}"
    AppendLine sb, ""

    ' --- Company search Leaflet control (topleft, below radius slider) ---
    AppendLine sb, "var searchControl = L.control({position:'topleft'});"
    AppendLine sb, "searchControl.onAdd = function(map) {"
    AppendLine sb, "  var div = L.DomUtil.create('div', 'search-control');"
    AppendLine sb, "  div.style.cssText = 'background:white;padding:8px 12px;border-radius:4px;font-size:13px;box-shadow:0 1px 5px rgba(0,0,0,.3);min-width:160px';"
    AppendLine sb, "  var opts = companies.map(function(c, i) {"
    AppendLine sb, "    return '<option value=""' + i + '"">' + escHtml(c.Name) + '</option>';"
    AppendLine sb, "  }).join('');"
    AppendLine sb, "  div.innerHTML = '<label style=""display:block;margin-bottom:4px""><strong>Find company</strong></label>'"
    AppendLine sb, "              + '<select id=""companySearch"" style=""font-size:13px;width:100%"">'"
    AppendLine sb, "              + '<option value="" selected>Select Company</option>' + opts + '</select>';"
    AppendLine sb, "  L.DomEvent.disableClickPropagation(div);"
    AppendLine sb, "  L.DomEvent.disableScrollPropagation(div);"
    AppendLine sb, "  return div;"
    AppendLine sb, "};"
    AppendLine sb, "searchControl.addTo(map);"
    AppendLine sb, ""

    ' --- Event bindings ---
    AppendLine sb, "document.getElementById('radiusSlider').addEventListener('input', function() {"
    AppendLine sb, "  document.getElementById('radiusLabel').textContent = this.value;"
    AppendLine sb, "  if (selectedMarker) { refreshRadius(); }"
    AppendLine sb, "});"
    AppendLine sb, ""
    AppendLine sb, "document.getElementById('companySearch').addEventListener('change', function() {"
    AppendLine sb, "  var idx = parseInt(this.value, 10);"
    AppendLine sb, "  if (isNaN(idx)) return;"
    AppendLine sb, "  var m = markers[idx];"
    AppendLine sb, "  var c = companies[idx];"
    AppendLine sb, "  map.setView(m.getLatLng(), Math.max(map.getZoom(), 12));"
    AppendLine sb, "  applySelection(m, c);"
    AppendLine sb, "  this.value = '';"
    AppendLine sb, "});"
    AppendLine sb, ""
    AppendLine sb, "document.getElementById('fieldSelect').addEventListener('change', function() {"
    AppendLine sb, "  applyFilter(this.value);"
    AppendLine sb, "});"
    AppendLine sb, ""
    AppendLine sb, "markers.forEach(function(marker, i) {"
    AppendLine sb, "  marker.on('click', function(e) {"
    AppendLine sb, "    L.DomEvent.stopPropagation(e);"
    AppendLine sb, "    applySelection(marker, companies[i]);"
    AppendLine sb, "  });"
    AppendLine sb, "});"
    AppendLine sb, ""
    AppendLine sb, "map.on('click', clearSelection);"

    AppendLine sb, "</script></body></html>"

    BuildMapHtml = sb
End Function
