' import_modules.vbs  (Step 2 of 2)
' Imports the VBA .bas modules into GeoViz.xlsm.
' Requires: Trust Center > Macro Settings > "Trust access to VBA project object model" ON.

Option Explicit

Dim fso, scriptDir, wbPath, srcDir
Dim xl, wb, vbProj, comp

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
wbPath = fso.BuildPath(scriptDir, "GeoViz.xlsm")
srcDir = fso.BuildPath(scriptDir, "src")

If Not fso.FileExists(wbPath) Then
    WScript.Echo "ERROR: GeoViz.xlsm not found. Run setup_workbook.vbs first."
    WScript.Quit 1
End If

Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

Set wb = xl.Workbooks.Open(wbPath)

Set vbProj = wb.VBProject

' Remove any default modules (Module1, Module2, etc.)
Dim toRemove()
Dim count
count = 0
ReDim toRemove(vbProj.VBComponents.Count)
For Each comp In vbProj.VBComponents
    If comp.Type = 1 Then  ' vbext_ct_StdModule = 1
        toRemove(count) = comp.Name
        count = count + 1
    End If
Next

Dim j
For j = 0 To count - 1
    vbProj.VBComponents.Remove vbProj.VBComponents(toRemove(j))
Next

' Import .bas files
vbProj.VBComponents.Import fso.BuildPath(srcDir, "mod_Config.bas")
vbProj.VBComponents.Import fso.BuildPath(srcDir, "mod_DataReader.bas")
vbProj.VBComponents.Import fso.BuildPath(srcDir, "mod_Geocoder.bas")
vbProj.VBComponents.Import fso.BuildPath(srcDir, "mod_JsonBuilder.bas")
vbProj.VBComponents.Import fso.BuildPath(srcDir, "mod_Macros.bas")

wb.Save
wb.Close False
xl.Quit

Set xl = Nothing
Set fso = Nothing

WScript.Echo "Done. VBA modules imported and workbook saved."
WScript.Echo ""
WScript.Echo "Verification steps:"
WScript.Echo "  1. Open GeoViz.xlsm in Excel."
WScript.Echo "  2. Alt+F11 to open VBA editor — confirm 5 modules are present."
WScript.Echo "  3. Debug > Compile VBAProject — should complete with no errors."
WScript.Echo "  4. Run macro: Developer > Macros > GeocodeSheet"
