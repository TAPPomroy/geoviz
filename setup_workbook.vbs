' setup_workbook.vbs  (Step 1 of 2)
' Creates GeoViz.xlsm with the GeoViz sheet and CompanyData table.
' Does NOT require VBA project trust — run this first.
' After this runs, open GeoViz.xlsm in Excel, then run import_modules.vbs
' (which requires Trust Center > Macro Settings > Trust access to VBA project object model).

Option Explicit

Dim fso, scriptDir, wbPath
Dim xl, wb, ws, tbl, rng

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
wbPath = fso.BuildPath(scriptDir, "GeoViz.xlsm")

If fso.FileExists(wbPath) Then
    WScript.Echo "GeoViz.xlsm already exists — delete it first if you want to recreate it."
    WScript.Quit 1
End If

Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

Set wb = xl.Workbooks.Add

' ----- Sheet setup -----
Set ws = wb.Sheets(1)
ws.Name = "GeoViz"

' Delete extra sheets (Excel adds 3 by default)
Dim i
For i = wb.Sheets.Count To 2 Step -1
    wb.Sheets(i).Delete
Next

' Write headers A1:C1
ws.Cells(1, 1).Value = "Name"
ws.Cells(1, 2).Value = "Address"
ws.Cells(1, 3).Value = "ISP"

' Sample data rows 2-5
ws.Cells(2, 1).Value = "Acme Corp"
ws.Cells(2, 2).Value = "350 5th Ave, New York, NY 10118"
ws.Cells(2, 3).Value = "Comcast"

ws.Cells(3, 1).Value = "Globex"
ws.Cells(3, 2).Value = "1600 Pennsylvania Ave NW, Washington, DC 20500"
ws.Cells(3, 3).Value = "AT&T"

ws.Cells(4, 1).Value = "Initech"
ws.Cells(4, 2).Value = "233 S Wacker Dr, Chicago, IL 60606"
ws.Cells(4, 3).Value = "Comcast"

ws.Cells(5, 1).Value = "Umbrella Co"
ws.Cells(5, 2).Value = "1 Infinite Loop, Cupertino, CA 95014"
ws.Cells(5, 3).Value = "Verizon"

' Auto-fit columns A-C
ws.Columns("A:C").AutoFit

' Create Excel Table named CompanyData (xlSrcRange=1, xlYes=1)
Set rng = ws.Range("A1:C5")
Set tbl = ws.ListObjects.Add(1, rng, , 1)
tbl.Name = "CompanyData"

' Save as .xlsm (xlOpenXMLMacroEnabled = 52)
wb.SaveAs wbPath, 52

wb.Close False
xl.Quit

Set xl = Nothing
Set fso = Nothing

WScript.Echo "Step 1 done. GeoViz.xlsm created at: " & wbPath
WScript.Echo ""
WScript.Echo "Next steps:"
WScript.Echo "  1. In Excel: File > Options > Trust Center > Trust Center Settings"
WScript.Echo "     > Macro Settings > check 'Trust access to the VBA project object model'"
WScript.Echo "     > click OK twice."
WScript.Echo "  2. Run: cscript import_modules.vbs"
