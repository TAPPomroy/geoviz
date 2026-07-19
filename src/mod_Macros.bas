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

' Stub for Phase 2 — GenerateMap will be implemented there.
' Declared here so Phase 1 testing can call it as a no-op.
Public Sub GenerateMap()
    MsgBox "GenerateMap not yet implemented (Phase 2).", vbInformation, "GeoViz"
End Sub
