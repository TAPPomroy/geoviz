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

    ' Deterministic ISP->color map via alphabetical sort (D-05)
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

    ' ISP legend in bottom-right corner (D-08)
    AppendLine sb, "var legend = L.control({position:'bottomright'});"
    AppendLine sb, "legend.onAdd = function(map){"
    AppendLine sb, "  var div = L.DomUtil.create('div');"
    AppendLine sb, "  div.style.cssText = 'background:white;padding:8px 12px;border-radius:4px;line-height:1.8;font-size:13px;box-shadow:0 1px 5px rgba(0,0,0,.3)';"
    AppendLine sb, "  var html = '<strong>ISP</strong><br>';"
    AppendLine sb, "  ispNames.forEach(function(isp){"
    AppendLine sb, "    html += '<span style=""display:inline-block;width:12px;height:12px;background:'"
    AppendLine sb, "          + ispColorMap[isp] + ';margin-right:6px;border-radius:2px;vertical-align:middle;""></span>'"
    AppendLine sb, "          + isp + '<br>';"
    AppendLine sb, "  });"
    AppendLine sb, "  if(ispNames.length > ISP_COLORS.length){"
    AppendLine sb, "    html += '<span style=""display:inline-block;width:12px;height:12px;background:#999999;margin-right:6px;border-radius:2px;vertical-align:middle;""></span>Other<br>';"
    AppendLine sb, "  }"
    AppendLine sb, "  div.innerHTML = html;"
    AppendLine sb, "  return div;"
    AppendLine sb, "};"
    AppendLine sb, "legend.addTo(map);"

    AppendLine sb, "</script></body></html>"

    BuildMapHtml = sb
End Function
