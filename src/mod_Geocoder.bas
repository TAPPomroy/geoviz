Attribute VB_Name = "mod_Geocoder"
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
    Dim latVal As Variant

    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    Set tbl = ws.ListObjects(TABLE_NAME)
    total = data.Count

    For i = 1 To total
        Set dict = data(i)
        addr = CStr(dict(HDR_ADDRESS))

        ' Skip already-geocoded rows
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

        ' Rate limit: wait >= 1100 ms between Nominatim requests
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

' Minimal JSON string extractor: finds key in JSON, returns the value.
' Works for simple flat JSON — sufficient for Nominatim lat/lon fields.
Private Function ExtractJsonString(json As String, key As String) As String
    Dim pos As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim rawVal As String
    Dim c As String

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
