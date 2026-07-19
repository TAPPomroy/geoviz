Attribute VB_Name = "mod_DataReader"
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
