# Phase 2: Map Rendering - Discussion Log

> **Audit trail only.** Do not use as input to planning, research, or execution agents.
> Decisions are captured in CONTEXT.md — this log preserves the alternatives considered.

**Date:** 2026-07-18
**Phase:** 2-Map Rendering
**Areas discussed:** HTML file path, ISP color palette, Marker visual style, VBA module structure

---

## HTML File Path

### File location

| Option | Description | Selected |
|--------|-------------|----------|
| ThisWorkbook.Path | Same folder as .xlsm. Simple, predictable. OneDrive sync fine since file opens immediately. | ✓ |
| Environ("TEMP") | Always writable, no OneDrive interference. File accumulates across runs. | |
| Hardcoded path | Predictable but requires folder to exist; not portable. | |

**User's choice:** ThisWorkbook.Path

### Browser open method

| Option | Description | Selected |
|--------|-------------|----------|
| Shell + explorer.exe | `Shell "explorer.exe """ & filePath & """"`. Handles paths with spaces when quoted. | ✓ |
| Shell + cmd /c start | More fragile with spaces in paths. | |
| Workbooks.Open | Opens in Excel — wrong tool for HTML. | |

**User's choice:** Shell + explorer.exe

---

## ISP Color Palette

### Palette source

| Option | Description | Selected |
|--------|-------------|----------|
| Standard qualitative palette | ColorBrewer Set1 or similar — 7 high-contrast, colorblind-friendly colors. Planner picks hex values. | ✓ |
| Custom hex values | User specifies exact colors. | |
| You decide | Claude picks colors for CartoDB Positron background. | |

**User's choice:** Standard qualitative palette (ColorBrewer Set1)

### "Other" color

| Option | Description | Selected |
|--------|-------------|----------|
| #999999 neutral gray | Visually recedes behind named ISP colors without disappearing. | ✓ |
| #CCCCCC light gray | May be too faint on Positron's light background. | |
| You decide | Planner picks a suitable gray. | |

**User's choice:** #999999 neutral gray

---

## Marker Visual Style

### Marker type

| Option | Description | Selected |
|--------|-------------|----------|
| Leaflet CircleMarker | Filled circle, ISP color fill, white border. Lightweight, no image assets. | ✓ |
| Custom L.divIcon | HTML/CSS circle. More control but heavier. | |
| Default Leaflet pin | Doesn't support per-ISP color without custom images. | |

**User's choice:** Leaflet CircleMarker

### Marker size and opacity

| Option | Description | Selected |
|--------|-------------|----------|
| radius: 8, fillOpacity: 0.85, weight: 1, color: white | Visible at county/state zoom; white border separates adjacent markers. | ✓ |
| radius: 6, fillOpacity: 0.7, weight: 1, color: white | Smaller, may be hard to click on dense map. | |
| You decide | Planner picks. | |

**User's choice:** radius: 8, fillOpacity: 0.85, weight: 1, color: white

### Legend position

| Option | Description | Selected |
|--------|-------------|----------|
| Bottom-right corner | Standard Leaflet control placement. Out of the way of map content. | ✓ |
| Top-right corner | More prominent but overlaps zoom controls. | |
| Bottom-left corner | Less conventional for legends. | |

**User's choice:** Bottom-right corner

---

## VBA Module Structure

### Module organization

| Option | Description | Selected |
|--------|-------------|----------|
| New mod_MapBuilder module | Mirrors Phase 1 pattern. mod_Macros.GenerateMap() calls it. | ✓ |
| Extend mod_Macros | Simpler but mixes concerns; makes that module large. | |
| You decide | Planner picks based on code size. | |

**User's choice:** New mod_MapBuilder module

### HTML construction method

| Option | Description | Selected |
|--------|-------------|----------|
| String concatenation with AppendLine helper | Private Sub AppendLine(sb, line) — idiomatic VBA, no dependencies. | ✓ |
| Array of strings joined at end | Join(lines, vbLf). Clean but slight VBA idiom mismatch. | |
| You decide | Planner picks. | |

**User's choice:** String concatenation with AppendLine helper

---

## Claude's Discretion

- Specific ColorBrewer Set1 hex values (or equivalent 7-color qualitative palette)
- Whether mod_MapBuilder exports BuildMapHtml() as a Function returning a String, or handles file write internally
- MarkerCluster CDN URL and version

## Deferred Ideas

- Inline Leaflet JS/CSS for offline use (OFF-V2-01 — v2 requirement)
- Corporate proxy / tile loading failure handling (open question in STATE.md)
- Clickable legend entries (Phase 3, FILT-02)
- Attribute filter dropdown (Phase 3, FILT-01)
