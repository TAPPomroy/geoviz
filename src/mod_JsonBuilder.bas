Attribute VB_Name = "mod_JsonBuilder"
Option Explicit

' Returns a JSON array string of all companies that have valid coordinates.
' Rows with GEOCODE_FAILED in Lat are excluded from the output.
' All table columns are included dynamically as JSON keys.
Public Function BuildCompanyJson(data As Collection) As String
    Dim sb As String
    Dim first As Boolean
    Dim key As Variant
    Dim val As Variant
    Dim isFirst As Boolean
    Dim item As Object

    sb = "[" & vbLf
    first = True

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
