# Phase 1: Data Pipeline — PLAN.md

**Phase:** 1 — Data Pipeline
**Goal:** The macro reliably reads the company table, geocodes uncached addresses at 1 req/sec, and writes lat/lon back to the sheet so subsequent runs skip already-geocoded rows.
**Created:** 2026-07-18
**Status:** Ready to execute

---

## Requirements Covered

| Req | Description |
|-----|-------------|
| DATA-01 | VBA reads company table from `CompanyData` table on `GeoViz` sheet |
| DATA-02 | Geocodes via Nominatim, 1 req/sec, with User-Agent header |
| DATA-03 | Caches lat/lon + timestamp; skips already-geocoded rows |
| DATA-04 | Writes `GEOCODE_FAILED` for unresolvable addresses; no retry on next run |

---

## Deliverables

1. Excel workbook (`.xlsm`) with VBA modules: `mod_Config`, `mod_DataReader`, `mod_Geocoder`, `mod_JsonBuilder`
2. `GeocodeSheet` public macro runnable from Developer tab or macro dialog
3. `BuildCompanyJson()` function returning a valid JSON array string (seam for Phase 2)
4. Sample `CompanyData` table on `GeoViz` sheet with 3–5 test rows for verification

---

## Sheet Contract

### Worksheet: `GeoViz`
### Table name: `CompanyData`

| Column | Header | Type | Notes |
|--------|--------|------|-------|
| Required | `Name` | String | Company name |
| Required | `Address` | String | Full address for geocoding |
| Required | `ISP` | String | ISP provider |
| Cache (auto-added) | `Lat` | Double / "GEOCODE_FAILED" | Written by `GeocodeSheet` |
| Cache (auto-added) | `Lon` | Double / "GEOCODE_FAILED" | Written by `GeocodeSheet` |
| Cache (auto-added) | `GeocodedAt` | Date string | Written by `GeocodeSheet` |
| Optional | any additional columns | Variant | Passed through to JSON |

**Cache columns are added to the table automatically on first run if they don't exist.**
Cache columns go to the right of ISP; additional user attribute columns go to the right of cache columns.

---

## Tasks

### Task 1 — Create workbook and `mod_Config`

**File:** `GeoViz.xlsm` (new workbook, save as macro-enabled)

Create a standard VBA module `mod_Config` with all project constants:

```vba
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
```

**Verify:** Module compiles with no errors (Debug → Compile VBAProject).

---

### Task 2 — Create `GeoViz` sheet and sample `CompanyData` table

In `GeoViz.xlsm`:

1. Rename `Sheet1` to `GeoViz`.
2. In cell A1 type the headers: `Name`, `Address`, `ISP` (in columns A–C).
3. Insert sample data rows 2–5 (3–4 realistic addresses in the same region, mix of ISPs).
4. Select A1:C5 → Insert → Table → check "My table has headers" → name the table `CompanyData`.

**Sample data (adjust city to a real one for geocoding tests):**

| Name | Address | ISP |
|------|---------|-----|
| Acme Corp | 350 5th Ave, New York, NY 10118 | Comcast |
| Globex | 1600 Pennsylvania Ave NW, Washington, DC 20500 | AT&T |
| Initech | 233 S Wacker Dr, Chicago, IL 60606 | Comcast |
| Umbrella Co | 1 Infinite Loop, Cupertino, CA 95014 | Verizon |

**Verify:** `CompanyData` appears in the Name Box dropdown; table has 3 columns.

---

### Task 3 — Create `mod_DataReader`

Responsibilities: locate `CompanyData`, ensure cache columns exist, read all rows into a `Collection` of `Scripting.Dictionary` objects.

```vba
Option Explicit

' Returns a Collection of Scripting.Dictionary, one per data row.
' Each dict has keys matching every column header.
' Rows that already have a non-empty Lat that is not GEOCODE_FAILED are
' returned as-is; rows needing geocoding have Lat="" and Lon="".
Public Function ReadCompanyData() As Collection
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim result As New Collection
    Dim row As ListRow
    Dim col As ListColumn
    Dim dict As Object
    Dim hdr As String
    Dim latVal As Variant

    Set ws = GetGeoVizSheet()
    Set tbl = GetCompanyTable(ws)

    EnsureCacheColumns tbl

    For Each row In tbl.ListRows
        Set dict = CreateObject("Scripting.Dictionary")
        For Each col In tbl.ListColumns
            hdr = col.Name
            dict(hdr) = row.Range.Cells(1, col.Index).Value
        Next col
        result.Add dict
    Next row

    Set ReadCompanyData = result
End Function

' Returns the GeoViz worksheet; raises error if not found.
Private Function GetGeoVizSheet() As Worksheet
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1, "mod_DataReader", _
            "Sheet '" & SHEET_NAME & "' not found. Please create it."
    End If
    Set GetGeoVizSheet = ws
End Function

' Returns the CompanyData ListObject; raises error if not found.
Private Function GetCompanyTable(ws As Worksheet) As ListObject
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = ws.ListObjects(TABLE_NAME)
    On Error GoTo 0
    If tbl Is Nothing Then
        Err.Raise vbObjectError + 2, "mod_DataReader", _
            "Table '" & TABLE_NAME & "' not found on sheet '" & SHEET_NAME & "'."
    End If
    Set GetCompanyTable = tbl
End Function

' Adds Lat, Lon, GeocodedAt columns to the table if they do not exist.
Private Sub EnsureCacheColumns(tbl As ListObject)
    Dim needed(2) As String
    needed(0) = HDR_LAT
    needed(1) = HDR_LON
    needed(2) = HDR_GEOCODED
    Dim i As Integer
    Dim col As ListColumn
    Dim found As Boolean
    Dim n As String

    For i = 0 To 2
        n = needed(i)
        found = False
        For Each col In tbl.ListColumns
            If LCase(col.Name) = LCase(n) Then found = True: Exit For
        Next col
        If Not found Then
            tbl.ListColumns.Add.Name = n
        End If
    Next i
End Sub

' Helper: get the ListRow in tbl that corresponds to a Collection item's row number.
' (Used by mod_Geocoder to write back to the sheet.)
Public Function GetTableRow(tbl As ListObject, rowIndex As Long) As ListRow
    Set GetTableRow = tbl.ListRows(rowIndex)
End Function

' Helper: write geocode result back to the sheet row.
Public Sub WriteGeocode(tbl As ListObject, rowIndex As Long, _
                         latVal As Variant, lonVal As Variant)
    Dim r As ListRow
    Set r = tbl.ListRows(rowIndex)
    r.Range.Cells(1, tbl.ListColumns(HDR_LAT).Index).Value = latVal
    r.Range.Cells(1, tbl.ListColumns(HDR_LON).Index).Value = lonVal
    r.Range.Cells(1, tbl.ListColumns(HDR_GEOCODED).Index).Value = _
        Format(Now(), "yyyy-mm-dd hh:mm:ss")
End Sub

' Highlights failed rows in red interior fill.
Public Sub HighlightFailedRows(tbl As ListObject)
    Dim row As ListRow
    Dim latIdx As Long
    latIdx = tbl.ListColumns(HDR_LAT).Index
    For Each row In tbl.ListRows
        If CStr(row.Range.Cells(1, latIdx).Value) = GEOCODE_FAILED Then
            row.Range.Interior.Color = RGB(255, 102, 102)
        End If
    Next row
End Sub
```

**Verify:** In Immediate Window: `?mod_DataReader.ReadCompanyData().Count` returns 4 (or however many sample rows).

---

### Task 4 — Create `mod_Geocoder`

Responsibilities: for each row with an empty or stale `Lat`, call Nominatim, parse the first result's lat/lon, write back via `mod_DataReader`, enforce 1 req/sec.

```vba
Option Explicit

' Geocodes all rows in the collection that need it.
' Updates the sheet directly via mod_DataReader helpers.
' Updates Application.StatusBar during processing.
Public Sub GeocodeAll(data As Collection)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim dict As Object
    Dim addr As String
    Dim latResult As Double, lonResult As Double
    Dim success As Boolean
    Dim total As Long

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    Set tbl = ws.ListObjects(TABLE_NAME)
    total = data.Count

    For i = 1 To total
        Set dict = data(i)
        addr = CStr(dict(HDR_ADDRESS))

        ' Skip already-geocoded rows
        Dim latVal As Variant
        latVal = dict(HDR_LAT)
        If NeedsGeocode(latVal) = False Then GoTo NextRow

        Application.StatusBar = "Geocoding " & i & "/" & total & " — " & dict(HDR_NAME) & "..."

        success = CallNominatim(addr, latResult, lonResult)

        If success Then
            mod_DataReader.WriteGeocode tbl, i, latResult, lonResult
            dict(HDR_LAT) = latResult
            dict(HDR_LON) = lonResult
        Else
            mod_DataReader.WriteGeocode tbl, i, GEOCODE_FAILED, GEOCODE_FAILED
            dict(HDR_LAT) = GEOCODE_FAILED
            dict(HDR_LON) = GEOCODE_FAILED
        End If

        ' Rate limit: wait > 1 second between requests (Nominatim policy)
        Application.Wait Now + TimeSerial(0, 0, 0) + (RATE_LIMIT_MS / 86400000#)

NextRow:
    Next i

    Application.StatusBar = False
    mod_DataReader.HighlightFailedRows tbl
End Sub

' Returns True if row needs geocoding (Lat is empty, or is GEOCODE_FAILED).
Private Function NeedsGeocode(latVal As Variant) As Boolean
    If IsEmpty(latVal) Or latVal = "" Then
        NeedsGeocode = True
    ElseIf CStr(latVal) = GEOCODE_FAILED Then
        NeedsGeocode = False  ' Already marked failed — do not retry
    Else
        NeedsGeocode = False  ' Has valid coords
    End If
End Function

' Calls Nominatim for a single address.
' Returns True and sets lat/lon on success; returns False on any failure.
Private Function CallNominatim(address As String, _
                                ByRef latOut As Double, _
                                ByRef lonOut As Double) As Boolean
    Dim http As Object
    Dim url As String
    Dim resp As String
    Dim latStr As String, lonStr As String

    url = NOMINATIM_URL & "?q=" & UrlEncode(address) & "&format=json&limit=1"

    On Error GoTo HttpError
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", USER_AGENT
    http.SetRequestHeader "Accept-Language", "en"
    http.Send

    If http.Status <> 200 Then GoTo HttpError

    resp = http.ResponseText

    ' Parse lat/lon from JSON: find first "lat":"..." and "lon":"..."
    latStr = ExtractJsonString(resp, """lat"":")
    lonStr = ExtractJsonString(resp, """lon"":")

    If latStr = "" Or lonStr = "" Then GoTo HttpError

    latOut = CDbl(latStr)
    lonOut = CDbl(lonStr)
    CallNominatim = True
    Exit Function

HttpError:
    CallNominatim = False
End Function

' Minimal JSON string extractor: finds key in JSON, returns the quoted value.
' Works for simple flat JSON — sufficient for Nominatim lat/lon fields.
Private Function ExtractJsonString(json As String, key As String) As String
    Dim pos As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim rawVal As String

    pos = InStr(json, key)
    If pos = 0 Then Exit Function

    startPos = pos + Len(key)
    ' Skip whitespace
    Do While Mid(json, startPos, 1) = " " Or Mid(json, startPos, 1) = Chr(9)
        startPos = startPos + 1
    Loop

    ' Value may be quoted ("45.123") or unquoted (45.123)
    If Mid(json, startPos, 1) = Chr(34) Then
        ' Quoted string — extract between quotes
        startPos = startPos + 1
        endPos = InStr(startPos, json, Chr(34))
        rawVal = Mid(json, startPos, endPos - startPos)
    Else
        ' Unquoted number — read until comma, ], or }
        endPos = startPos
        Do While endPos <= Len(json)
            Dim c As String
            c = Mid(json, endPos, 1)
            If c = "," Or c = "]" Or c = "}" Then Exit Do
            endPos = endPos + 1
        Loop
        rawVal = Trim(Mid(json, startPos, endPos - startPos))
    End If

    ExtractJsonString = rawVal
End Function

' URL-encodes a string for query parameters.
Private Function UrlEncode(s As String) As String
    Dim i As Integer
    Dim c As String
    Dim result As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case Asc(c)
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                result = result & c
            Case 32
                result = result & "+"
            Case Else
                result = result & "%" & Right("0" & Hex(Asc(c)), 2)
        End Select
    Next i
    UrlEncode = result
End Function
```

**Verify:** With 1 sample row uncached, run `GeocodeAll` against one address in Immediate Window — confirm Lat/Lon columns populate in the sheet and status bar clears.

---

### Task 5 — Create `mod_JsonBuilder`

Responsibilities: serialize the company data collection to a JSON array string. This is the Phase 1/Phase 2 seam — Phase 2 calls `BuildCompanyJson()` directly.

```vba
Option Explicit

' Returns a JSON array string of all companies that have valid coordinates.
' Rows with GEOCODE_FAILED in Lat are excluded from the output.
' All table columns are included dynamically as JSON keys.
Public Function BuildCompanyJson(data As Collection) As String
    Dim sb As String
    Dim dict As Object
    Dim first As Boolean
    Dim key As Variant
    Dim val As Variant
    Dim isFirst As Boolean

    sb = "[" & vbLf
    first = True

    Dim item As Object
    For Each item In data
        ' Skip failed geocodes
        If CStr(item(HDR_LAT)) = GEOCODE_FAILED Then GoTo NextItem
        If IsEmpty(item(HDR_LAT)) Or item(HDR_LAT) = "" Then GoTo NextItem

        If Not first Then sb = sb & "," & vbLf
        first = False

        sb = sb & "  {"
        isFirst = True
        For Each key In item.Keys
            If Not isFirst Then sb = sb & ","
            isFirst = False

            val = item(key)
            ' Lat and Lon are numeric — no quotes
            If LCase(CStr(key)) = LCase(HDR_LAT) Or _
               LCase(CStr(key)) = LCase(HDR_LON) Then
                If IsNumeric(val) Then
                    sb = sb & """" & JsonKey(CStr(key)) & """:" & CStr(val)
                Else
                    sb = sb & """" & JsonKey(CStr(key)) & """:null"
                End If
            Else
                sb = sb & """" & JsonKey(CStr(key)) & """:""" & JsonEscape(CStr(val)) & """"
            End If
        Next key
        sb = sb & "}"
NextItem:
    Next item

    sb = sb & vbLf & "]"
    BuildCompanyJson = sb
End Function

' Escapes a string for safe embedding in a JSON string value.
Private Function JsonEscape(s As String) As String
    Dim result As String
    result = s
    result = Replace(result, "\", "\\")
    result = Replace(result, Chr(34), "\" & Chr(34))
    result = Replace(result, "/", "\/")
    result = Replace(result, Chr(8), "\b")
    result = Replace(result, Chr(12), "\f")
    result = Replace(result, Chr(10), "\n")
    result = Replace(result, Chr(13), "\r")
    result = Replace(result, Chr(9), "\t")
    JsonEscape = result
End Function

' Sanitizes a column header for use as a JSON key.
Private Function JsonKey(s As String) As String
    JsonKey = JsonEscape(s)
End Function
```

**Verify:** In Immediate Window, after geocoding: `?mod_JsonBuilder.BuildCompanyJson(mod_DataReader.ReadCompanyData())` — paste the output into a browser console and run `JSON.parse(...)`. Must parse without errors.

---

### Task 6 — Create `GeocodeSheet` public macro entry point

In `ThisWorkbook` module (or a new module `mod_Macros`), add:

```vba
Option Explicit

' Public entry point: reads company table, geocodes uncached rows, caches results.
' Run from Developer → Macros, or assign to a button.
Public Sub GeocodeSheet()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim data As Collection
    Set data = mod_DataReader.ReadCompanyData()

    mod_Geocoder.GeocodeAll data

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error in GeocodeSheet: " & Err.Description, vbCritical, "GeoViz"
End Sub

' Stub for Phase 2 — GenerateMap will be implemented there.
' Declared here so Phase 1 testing can call it as a no-op.
Public Sub GenerateMap()
    MsgBox "GenerateMap not yet implemented (Phase 2).", vbInformation, "GeoViz"
End Sub
```

**Verify:** Run `GeocodeSheet` from Macro dialog. Status bar shows progress, Lat/Lon/GeocodedAt columns populate, failed rows (if any) go red. Re-run — completes instantly with no new HTTP calls.

---

## Acceptance Tests

Matching ROADMAP.md success criteria:

| # | Test | Pass Condition |
|---|------|----------------|
| SC-1 | Fresh sheet: run `GeocodeSheet` | Lat/Lon/GeocodedAt populated for all valid rows; no HTTP 429 in status |
| SC-2 | Re-run on fully-geocoded sheet | Completes in < 2 seconds, zero new API calls (add `Debug.Print` count as temp check) |
| SC-3 | Row with garbage address (e.g., `"zzzzz"`) | Lat = `GEOCODE_FAILED`, row highlighted red, not retried on re-run |
| SC-4 | JSON output | `JSON.parse(BuildCompanyJson(ReadCompanyData()))` succeeds in browser console; no garbled characters in Name/Address fields |

---

## Module Summary

| Module | Role | Phase |
|--------|------|-------|
| `mod_Config` | Constants (URLs, headers, table/column names) | 1 |
| `mod_DataReader` | Sheet I/O, cache column management, write-back | 1 |
| `mod_Geocoder` | Nominatim HTTP calls, rate limiting, failure handling | 1 |
| `mod_JsonBuilder` | JSON serialization of company collection | 1 (seam for Phase 2) |
| `ThisWorkbook` | `GeocodeSheet` entry point; `GenerateMap` stub | 1 |

---

## Critical Constraints (from CLAUDE.md)

- **HTTP client:** `WinHttp.WinHttpRequest.5.1` — not MSXML2 (better TLS on corporate machines)
- **File output:** `ADODB.Stream` with `Charset = "utf-8"` — Phase 2 concern but noted here
- **Rate limit:** `Application.Wait` must enforce ≥ 1100 ms between Nominatim requests
- **User-Agent:** `GeoViz/1.0 (pomroyanalytics@gmail.com)` on every HTTP request
- **Cache guard:** Skip rows where Lat is non-empty AND not `GEOCODE_FAILED`

---

*Plan created: 2026-07-18*
*Phase: 1 — Data Pipeline*
