Attribute VB_Name = "mod_Config"
Option Explicit

Public Const NOMINATIM_URL  As String = "https://nominatim.openstreetmap.org/search"
Public Const USER_AGENT     As String = "GeoViz/1.0 (pomroyanalytics@gmail.com)"
Public Const RATE_LIMIT_MS  As Long   = 1100
Public Const TABLE_NAME     As String = "CompanyData"
Public Const SHEET_NAME     As String = "GeoViz"
Public Const OUTPUT_FILE    As String = "geoviz_map.html"
Public Const GEOCODE_FAILED As String = "GEOCODE_FAILED"

Public Const HDR_NAME       As String = "Name"
Public Const HDR_ADDRESS    As String = "Address"
Public Const HDR_ISP        As String = "ISP"
Public Const HDR_LAT        As String = "Lat"
Public Const HDR_LON        As String = "Lon"
Public Const HDR_GEOCODED   As String = "GeocodedAt"
