# Requirements: GeoViz

**Defined:** 2026-07-18
**Core Value:** Click any company on the map and immediately see which neighboring companies share its ISP — with an adjustable proximity radius.

## v1 Requirements

### Data Pipeline

- [ ] **DATA-01**: VBA reads company table from a designated Excel worksheet (columns: name, address, ISP, plus any additional attribute columns)
- [ ] **DATA-02**: VBA geocodes addresses via Nominatim API, throttling to 1 request per second with a proper User-Agent header
- [ ] **DATA-03**: VBA caches geocoded lat/lon and a timestamp back into the Excel sheet, skipping rows that already have coordinates
- [ ] **DATA-04**: VBA writes "GEOCODE_FAILED" to the lat/lon cells for addresses that cannot be geocoded, preventing retry storms

### Map Rendering

- [ ] **MAP-01**: Generated HTML renders a zoomable and pannable Leaflet.js map in the default browser
- [ ] **MAP-02**: Company markers are colored by ISP using up to 7 deterministic colors (alphabetical ISP sort → fixed palette); ISPs beyond 7 use a neutral "Other" gray
- [ ] **MAP-03**: Markers cluster automatically in dense areas to prevent overlap
- [ ] **MAP-04**: Map uses CartoDB Positron as the basemap (muted background improves marker legibility)
- [ ] **MAP-05**: A single macro button in the Excel workbook generates the HTML file and opens it in the default browser

### Interaction

- [ ] **INT-01**: Clicking a company marker opens a popup showing: company name, ISP, address, and all attribute columns
- [ ] **INT-02**: Clicking a company marker draws a radius circle centered on that company and highlights all companies within the radius
- [ ] **INT-03**: Within the highlighted neighbors, companies sharing the same ISP as the selected company are visually distinguished (e.g., brighter color or different marker shape)
- [ ] **INT-04**: A radius slider on the map UI lets the user adjust the proximity distance; the circle and highlights update live

### Filtering & Legend

- [ ] **FILT-01**: A dropdown or toggle on the map lets the user switch marker coloring between available attribute columns (defaults to ISP)
- [ ] **FILT-02**: An ISP legend on the map allows individual ISPs to be shown or hidden by clicking their legend entry

## v2 Requirements

### Enhanced Interaction

- **INT-V2-01**: Neighbor count displayed in the popup header ("3 neighbors within 5 miles")
- **INT-V2-02**: "Same ISP only" mode toggle — hide all non-same-ISP markers globally when a company is selected

### Offline Support

- **OFF-V2-01**: Option to inline Leaflet.js CSS/JS into the HTML for fully offline use (no CDN dependency)

### Data Quality

- **DATA-V2-01**: Audit column showing what Nominatim actually matched (display_name field) for geocoded addresses

## Out of Scope

| Feature | Reason |
|---------|--------|
| Python runtime | Zero-dependency requirement — pure VBA only |
| Excel add-in | Workbook-embedded macro only; add-ins require IT deployment |
| Real-time data sync | Map is a snapshot generated on demand |
| Routing / directions | Not relevant to the ISP outage use case |
| Heatmaps | Misleading at this dataset density; adds complexity |
| Mobile / tablet | Desktop browser only; single user |
| WebBrowser ActiveX control | Broken in current Excel/Windows — open in browser instead |
| Paid geocoding APIs | Nominatim is sufficient for < 100 companies |
| Multi-workbook support | Single workbook, single sheet |

## Traceability

| Requirement | Phase | Status |
|-------------|-------|--------|
| DATA-01 | TBD | Pending |
| DATA-02 | TBD | Pending |
| DATA-03 | TBD | Pending |
| DATA-04 | TBD | Pending |
| MAP-01 | TBD | Pending |
| MAP-02 | TBD | Pending |
| MAP-03 | TBD | Pending |
| MAP-04 | TBD | Pending |
| MAP-05 | TBD | Pending |
| INT-01 | TBD | Pending |
| INT-02 | TBD | Pending |
| INT-03 | TBD | Pending |
| INT-04 | TBD | Pending |
| FILT-01 | TBD | Pending |
| FILT-02 | TBD | Pending |

**Coverage:**
- v1 requirements: 15 total
- Mapped to phases: 0 (pending roadmap)
- Unmapped: 15 ⚠️

---
*Requirements defined: 2026-07-18*
*Last updated: 2026-07-18 after initial definition*
