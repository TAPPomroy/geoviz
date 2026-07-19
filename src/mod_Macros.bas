Attribute VB_Name = "mod_Macros"
Option Explicit

' Public entry point: reads company table, geocodes uncached rows, caches results.
' Run from Developer -> Macros, or assign to a button.
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

' Reads company data, builds the Leaflet map HTML, writes it to disk, and opens it in the browser.
' Assign to the "Generate Map" button on the GeoViz sheet.
Public Sub GenerateMap()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.StatusBar = "Building map..."

    Dim data As Collection
    Set data = mod_DataReader.ReadCompanyData()

    Dim jsonStr As String
    jsonStr = mod_JsonBuilder.BuildCompanyJson(data)

    Dim htmlStr As String
    htmlStr = mod_MapBuilder.BuildMapHtml(jsonStr)

    Dim filePath As String
    filePath = LocalWorkbookPath() & "\" & OUTPUT_FILE
    mod_MapBuilder.WriteUtf8File filePath, htmlStr

    ' Shell must run AFTER WriteUtf8File completes (Shell is async; write is synchronous)
    Shell "explorer.exe """ & filePath & """", vbNormalFocus

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error in GenerateMap: " & Err.Description, vbCritical, "GeoViz"
End Sub

' Returns the local filesystem path of the workbook directory.
' When the workbook is opened via OneDrive sync, ThisWorkbook.Path returns an
' https://d.docs.live.net/<id>/... URL that ADODB.Stream cannot write to.
' This function extracts the relative path segment and combines it with the
' local OneDrive root from the OneDriveConsumer environment variable.
Private Function LocalWorkbookPath() As String
    Dim wbPath As String
    wbPath = ThisWorkbook.Path

    If Left(wbPath, 8) <> "https://" Then
        LocalWorkbookPath = wbPath
        Exit Function
    End If

    ' Resolve local OneDrive root (personal sync folder)
    Dim odRoot As String
    odRoot = Environ("OneDriveConsumer")
    If odRoot = "" Then odRoot = Environ("OneDrive")
    If odRoot = "" Then
        Err.Raise vbObjectError + 1, "LocalWorkbookPath", _
            "Cannot resolve local OneDrive path. Save the workbook locally and retry."
    End If

    ' URL format: https://d.docs.live.net/<account-id>/<rel/path/segments>
    ' Split on "/" — parts(3) is the account ID, parts(4+) are the relative path.
    Dim parts() As String
    parts = Split(wbPath, "/")

    Dim relPath As String
    Dim i As Integer
    For i = 4 To UBound(parts)
        relPath = relPath & "\" & parts(i)
    Next i

    LocalWorkbookPath = odRoot & relPath
End Function
