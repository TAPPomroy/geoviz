# Phase 1: Data Pipeline — Discussion Log

**Date:** 2026-07-18
**Areas discussed:** Sheet structure, Progress feedback, JSON shape, Macro entry point

---

## Sheet Structure

**Q: How should the macro find your company data?**
Options: Named Excel Table / Fixed header row / Configurable via constants
**Selected:** Named Excel Table (ListObject)

**Q: What should the named table and worksheet be called?**
Options: Table: CompanyData, Sheet: GeoViz / Table: Companies, Sheet: Companies / You decide
**Selected:** Table: CompanyData, Sheet: GeoViz

**Q: What are the required column header names?**
Options: Name / Address / ISP / Company / FullAddress / ISPProvider / You decide
**Selected:** Name / Address / ISP

---

## Progress Feedback

**Q: What should the user see while geocoding is running?**
Options: Status bar updates / MsgBox at end only / Both status bar + MsgBox
**Selected:** Status bar updates

**Q: When geocoding finishes, what should happen?**
Options: Status bar summary / MsgBox with failure details / Highlight failed rows in red
**Selected:** Highlight failed rows in red

---

## JSON Shape

**Q: What fields should each company object include?**
Options: All table columns dynamically / Fixed fields only / You decide
**Selected:** All table columns dynamically

**Q: Where should the JSON array live between Phase 1 and Phase 2?**
Options: Built in memory / Hidden sheet cell / Written to .json file on disk
**Selected:** Built in memory, passed to HTML generator

---

## Macro Entry Point

**Q: How should the VBA macros be structured?**
Options: Separate macros: GeocodeSheet and GenerateMap / Single macro: RunGeoViz / You decide
**Selected:** Separate macros: GeocodeSheet and GenerateMap

---

## Deferred Ideas

- MsgBox failure summary (considered, user preferred row highlighting)
- OneDrive path resolution → Phase 2
- Corporate proxy / tile loading → Phase 2
- Offline Leaflet inlining → v2 (OFF-V2-01)
