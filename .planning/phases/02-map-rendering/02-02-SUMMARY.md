---
plan: "02-02"
phase: "02-map-rendering"
status: complete
self_check: PASSED
---

# Plan 02-02 Summary: Wire GenerateMap() + Excel button + human verify

## What Was Built

`src/mod_Macros.bas` — `GenerateMap()` fully implemented (Phase 1 stub replaced).

The macro orchestrates all Phase 2 modules in sequence:
1. `mod_DataReader.ReadCompanyData()` — reads the company table
2. `mod_JsonBuilder.BuildCompanyJson(data)` — serializes to JSON
3. `mod_MapBuilder.BuildMapHtml(jsonStr)` — generates the HTML document
4. `mod_MapBuilder.WriteUtf8File(filePath, htmlStr)` — writes to disk via ADODB.Stream
5. `Shell "explorer.exe ..."` — opens the file in the default browser (after write completes)

Includes `On Error GoTo ErrHandler`, `Application.ScreenUpdating` guards, and `Application.StatusBar` feedback.

## Key Fix: OneDrive Path Resolution

`ThisWorkbook.Path` returns an `https://d.docs.live.net/<id>/...` URL when the workbook is opened via OneDrive sync — ADODB.Stream cannot write to this URL. A private `LocalWorkbookPath()` helper detects the URL, extracts the relative path segment (everything after the account-ID component), and combines it with `Environ("OneDriveConsumer")` to return the writable local filesystem path.

## Human Verification Result

**Approved** — `GenerateMap()` ran without error, produced `geoviz_map.html`, opened it in the browser, and the map displayed all company data correctly.

## Acceptance Criteria

- `GenerateMap()` stub replaced — no "not yet implemented" MsgBox ✓
- Calls all four modules in correct order ✓
- `Shell "explorer.exe"` placed after `WriteUtf8File` (synchronous write before async open) ✓
- `On Error GoTo ErrHandler` with `vbCritical` / "GeoViz" title ✓
- `Application.ScreenUpdating` reset in both normal and error paths ✓
- `Application.Wait` does not appear in `GenerateMap()` ✓
- OneDrive URL path resolved to local filesystem path ✓
- MAP-01 through MAP-05 verified in browser ✓

## Files Modified

| File | Change |
|------|--------|
| `src/mod_Macros.bas` | `GenerateMap()` implemented; `LocalWorkbookPath()` helper added |

## Commits

- `feat(02-02): implement GenerateMap() in mod_Macros.bas`
- `fix(02-02): rewind ADODB.Stream position before SaveToFile`
- `fix(02-02): resolve OneDrive URL path for ADODB.Stream file write`
