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
    filePath = ThisWorkbook.Path & "\" & OUTPUT_FILE
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
