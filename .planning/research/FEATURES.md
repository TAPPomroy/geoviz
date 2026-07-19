# Features Research: GeoViz

**Domain:** Single-user Excel VBA tool generating a self-contained interactive HTML map
**Researched:** 2026-07-18
**Primary use case:** ISP outage triage — find geographically nearby companies sharing an affected ISP
**Secondary use case:** General proximity browsing for any selected company

---

## Table Stakes (must have)

These are the features users will immediately notice are missing. Without them the tool feels
broken, not merely incomplete.

| Feature | Why Expected | Complexity | Notes |
|---------|--------------|------------|-------|
| Pan and zoom | Core map navigation; every map tool has this | Low | Leaflet default behavior; zero implementation cost |
| Clickable markers | Users expect clicking a pin to show company details | Low | Leaflet `bindPopup`; the primary interaction model |
| Popup with key fields | Company name, address, ISP at minimum | Low | Must be readable without scrolling inside the popup |
| ISP-based color coding | The primary analytical question is "which ISP?" — visual encoding answers it instantly | Low-Med | Use a fixed palette of 5-7 colors; assign ISPs deterministically at generation time |
| Radius circle on selection | When a company is selected, draw a visible circle showing the search boundary | Low | Leaflet `L.circle`; visually confirms what "nearby" means |
| Highlight nearby companies | After selection, visually differentiate neighbors from the rest of the map | Med | Dim or gray out non-neighbors; do not hide them entirely (context matters) |
| Legend | Map colors mean nothing without a legend showing ISP-to-color mapping | Low | Static HTML element; generate alongside the map |
| Stable, offline HTML output | File must open in any modern browser with no internet dependency | Med | Bundle Leaflet JS/CSS locally or use a CDN with a reliable fallback; tile layer is the one exception (OpenStreetMap) |

**The non-negotiable interaction loop:**
Click a company → see its details → see a radius circle → see highlighted neighbors using the same ISP.
Every table-stakes feature serves this loop.

---

## Differentiators (nice to have)

Features that meaningfully improve the tool without bloating it. Each one earns its complexity.

| Feature | Value Proposition | Complexity | Notes |
|---------|-------------------|------------|-------|
| Adjustable radius slider | Users can tighten or broaden the neighbor search interactively; critical for dense vs. sparse geographies | Med | HTML range input; recalculate on change using Haversine; no server call needed |
| ISP filter toggle (show/hide) | During an outage, suppress all companies NOT on the affected ISP to reduce noise | Low-Med | Checkbox list in sidebar; toggle marker visibility; works well for small datasets |
| Sidebar company list | A scrollable list of markers synced to map view — clicking a list item flies to the marker | Med | Enables fast lookup when markers are dense or overlapping; bidirectional sync (map click updates list, list click pans map) |
| Search/filter by company name | Type to narrow both the sidebar list and visible markers | Low-Med | Client-side substring filter on the static dataset; no backend needed |
| "Same ISP only" mode toggle | One-click to show only companies sharing the currently selected company's ISP | Low | Simple visibility filter; distinct from the radius feature |
| Neighbor count badge | After selection, display "X companies within Y km using [ISP]" | Low | Computed at click time; displayed in popup or sidebar header |
| Export selection to clipboard | Copy the list of nearby companies (name, address, ISP) for pasting into a ticket or email | Low | `navigator.clipboard.writeText`; high operational value during an outage |
| Marker clustering at low zoom | Prevent marker overlap when zoomed out; cluster dissolves on zoom in | Med | Leaflet.markercluster plugin; necessary if dataset exceeds ~50 companies |

**Priority order for differentiators:**
Adjustable radius > ISP filter toggle > "same ISP" mode > neighbor count > sidebar list > name search > clustering > export.
The first three cost little and directly serve the outage triage workflow.

---

## Anti-Features (deliberately exclude)

These features are tempting but wrong for a single-user, VBA-generated, static HTML tool.

| Anti-Feature | Why to Exclude | What to Do Instead |
|--------------|---------------|-------------------|
| Real-time data refresh | VBA generates a snapshot; adding live refresh requires a backend, authentication, and network calls — a completely different product | Regenerate the HTML from Excel when data changes; this takes seconds |
| Heatmaps | Heatmaps are for density analysis of large datasets (thousands of points); ISP outage triage needs individual company identity, not density aggregates | ISP-colored markers with dimming achieve the same situational awareness |
| Routing / driving directions | Adds Google Maps API dependency, API key management, and usage costs; outage triage is about understanding who is affected, not how to drive to them | Link to Google Maps in the popup for the one case where navigation is wanted |
| User accounts / saved states | Single user, local file — there is no meaningful "state" to save across sessions | The Excel workbook IS the state; regenerate when data changes |
| Drawing tools (custom polygons) | High implementation cost; the adjustable radius circle covers the geographic selection use case | Radius slider |
| Timeline / historical playback | Requires storing timestamped snapshots; out of scope for a VBA workbook | Not in scope |
| Choropleth / admin boundary overlays | Meaningful only for aggregated regional data; individual company points don't benefit | Not in scope |
| Mobile-responsive layout | Single user opening a local HTML file on a desktop/laptop; over-engineering the viewport handling adds complexity for zero operational benefit | Basic responsive CSS only; don't optimize for touch |
| Multi-language / i18n | Internal tool; English only is correct | Not in scope |
| Embed in Excel (WebBrowser control) | The Excel WebBrowser ActiveX control is broken in current Excel/Windows versions | Open in default browser via `Shell` command; this is the reliable path |

---

## UX Patterns Worth Adopting

Patterns from the broader interactive map and data exploration literature that apply cleanly to this tool's scope.

**One popup open at a time**
Leaflet's `map.openPopup()` enforces this automatically. Do not fight it. Multiple open popups create clutter; the sidebar handles persistent multi-company comparison if needed.

**Dim rather than hide non-selected markers**
When a company is selected and neighbors are highlighted, reduce opacity of non-neighbors to ~30% rather than hiding them. Users retain geographic context and can see the overall network distribution. Hiding creates disorientation ("where did everything go?").

**Deterministic ISP color assignment**
Sort ISP names alphabetically at generation time and assign colors from a fixed palette in that order. This ensures the same ISP always gets the same color across different map generations. Using a hashed or random assignment breaks user memory between sessions.

**Fixed 5-7 color palette, with a fallback**
Limit distinct ISP colors to 7. If the dataset has more ISPs than colors, group low-frequency ISPs into an "Other" category with a neutral gray. Exceeding 7 colors destroys discrimination, especially for users with color vision deficiency. Use a colorblind-safe palette (ColorBrewer qualitative schemes work well).

**Radius circle as confirmation, not just decoration**
Draw the radius circle immediately on marker click, before the user adjusts the slider. This confirms "these are your neighbors" visually and makes the slider adjustment feel like a refinement rather than a setup step.

**Sidebar list for dense datasets**
When markers overlap at the working zoom level, a sidebar list provides an alternative selection mechanism. List items should show: company name, ISP (colored dot), and distance from currently selected company (if one is selected). Clicking a list item should pan and zoom the map to that marker and open its popup.

**Search filters the list, not the map directly**
Typing in the search box should filter the sidebar list and pulse/highlight matching markers — do not hide non-matching markers, as this again destroys geographic context. The user is looking for a specific company; once found, they click it and the full map context remains visible.

**Neighbor count in the popup header**
Put "N neighbors within X km on [ISP]" as the first line of the selected company's popup. This is the answer to the primary question and should not require the user to count highlights.

---

## Feature Complexity Notes

Context on implementation effort to inform roadmap phase planning.

**What is genuinely simple in a static HTML/Leaflet context:**
- Marker rendering with color from a lookup table: trivial
- Popup HTML with arbitrary fields: trivial
- Legend: a static `<div>` generated alongside the map data
- Circle overlay: `L.circle(latlng, {radius: meters})` — three lines of JS
- Marker opacity changes: `marker.setOpacity(value)` — one line per marker
- Client-side name search: iterate an array, toggle a CSS class

**What requires care but is not complex:**
- Haversine distance calculation: ~10 lines of JS; math is well-known
- Radius slider wired to neighbor recalculation: event listener + loop over markers
- ISP filter toggles: build a visibility state object, iterate markers on change
- Marker clustering: one additional JS library (leaflet.markercluster, ~150KB)

**What is genuinely complex for this context:**
- Bidirectional map-sidebar sync: requires maintaining shared state object and careful event wiring; testable but fiddly
- Generating the full self-contained HTML from VBA: string concatenation of JS/CSS/data is error-prone; use a template file approach rather than building HTML in VBA string literals

**The VBA boundary: do heavy work in HTML/JS, not VBA**
VBA should only: read the Excel data, serialize it as a JSON array embedded in the HTML, and write the file. All filtering, distance calculation, coloring, and interaction logic belongs in the generated JavaScript. This separation makes the map debuggable in a browser dev tools without touching Excel.

**Dataset size assumptions**
This tool is designed for a small dataset — likely 20-200 companies. Clustering and performance optimization are differentiators, not requirements. At fewer than 200 markers, Leaflet renders instantly with no special handling. Do not add complexity to solve a scale problem that does not exist.

---

## Sources

- [Map UI Design Patterns — UXPin](https://www.uxpin.com/studio/blog/map-ui/)
- [Attribute Filter Pattern — Map UI Patterns](https://mapuipatterns.com/attribute-filter/)
- [5 Map UI Design Patterns That Elevate UX — Bricxlabs](https://bricxlabs.com/blogs/map-ui-design-patterns-examples)
- [Working with Clusters in Leaflet — Digital Geography](https://digital-geography.com/working-with-clusters-in-leaflet-increasing-useability/)
- [Leaflet Quick Start Guide](https://leafletjs.com/examples/quick-start/)
- [Filter UI Design: Sidebar vs Top Bar — Setproduct](https://www.setproduct.com/blog/filter-ui-design)
- [ISPBox Network Map Features](https://ispbox.net/feature/map)
- [7 Best Interactive Map Designs — Map Library](https://www.maplibrary.org/10196/7-ways-to-balance-ux-and-data-in-interactive-maps/)
