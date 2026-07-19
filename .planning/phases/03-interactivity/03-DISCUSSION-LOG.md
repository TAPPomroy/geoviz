# Phase 3: Interactivity - Discussion Log

> **Audit trail only.** Do not use as input to planning, research, or execution agents.
> Decisions are captured in CONTEXT.md — this log preserves the alternatives considered.

**Date:** 2026-07-19
**Phase:** 3-Interactivity
**Areas discussed:** Radius unit & defaults, Highlighting visual style, Multi-click & clear behavior, Attribute filter & legend interaction

---

## Radius Unit & Defaults

| Option | Description | Selected |
|--------|-------------|----------|
| Miles | More natural for US-based ISP outage scenarios | ✓ (selected) |
| Kilometers | Standard SI unit | |
| Both — user toggles mi/km | Toggle button next to slider | |

**User's choice:** Miles

---

| Option | Description | Selected |
|--------|-------------|----------|
| Default 25 mi, range 1–100 mi | City-to-regional scale | |
| Default 10 mi, range 1–50 mi | Tighter focus, dense metro | |
| Default 0.5 mi, range 0.1–5 mi | User-specified tight range | ✓ |

**User's choice:** Free text — "this needs to be a tighter focus. default radius 0.5 miles, ranged from 0.1 miles to 5 miles"
**Notes:** Companies are in a dense urban/campus environment, not regional ISP coverage zones. The preset options were all too wide.

---

| Option | Description | Selected |
|--------|-------------|----------|
| Top-left panel | Standard Leaflet map control placement | ✓ |
| Bottom-left panel | Can conflict with attribution | |
| Floating above map center | Can obscure markers | |

**User's choice:** Top-left panel

---

## Highlighting Visual Style

| Option | Description | Selected |
|--------|-------------|----------|
| Larger radius + white fill | radius: 12, fillColor white, ISP color as border | ✓ |
| Gold/yellow border highlight | Keep size/color, gold border, weight 3 | |
| Pulsing animation ring | CSS keyframes ripple | |

**User's choice:** Larger radius + white fill

---

| Option | Description | Selected |
|--------|-------------|----------|
| Outside = faded (low opacity), inside = normal | Non-neighbors get fillOpacity: 0.15 | ✓ |
| Inside = brighter, outside = unchanged | Neighbors get 1.0, others stay 0.85 | |
| Outside = hidden, inside = normal | Non-neighbors fully hidden | |

**User's choice:** Outside = faded (low opacity), inside = normal

---

| Option | Description | Selected |
|--------|-------------|----------|
| Bright white border, increased weight | color: '#ffffff', weight: 3 for same-ISP | ✓ |
| Star/diamond marker shape | Switch from CircleMarker (complex) | |
| Pulsing border animation | CSS keyframes on same-ISP neighbors | |

**User's choice:** Bright white border, increased weight

---

## Multi-click & Clear Behavior

| Option | Description | Selected |
|--------|-------------|----------|
| Replace selection | Clear old, activate new. One at a time. | ✓ |
| Compare mode — show both circles | Two circles simultaneously | |

**User's choice:** Replace selection

---

| Option | Description | Selected |
|--------|-------------|----------|
| Click anywhere on the map background | Leaflet map 'click' event | ✓ |
| Click selected marker again to deselect | Toggle behavior | |
| Only via a 'Clear' button in the top-left panel | Explicit button only | |

**User's choice:** Click anywhere on the map background

---

## Attribute Filter & Legend Interaction

| Option | Description | Selected |
|--------|-------------|----------|
| All non-system columns | Exclude Lat, Lon, timestamp. Include everything else. | ✓ |
| Only ISP + manually-tagged attribute columns | Requires Excel convention | |
| All columns including Lat/Lon | Include everything | |

**User's choice:** All non-system columns

---

| Option | Description | Selected |
|--------|-------------|----------|
| Legend updates to show new column's values + colors | Header and swatches update | ✓ |
| Legend stays as ISP regardless of filter | Simpler but misleading | |
| Legend hides when non-ISP column selected | Avoids confusion, removes info | |

**User's choice:** Legend updates to show the new column's unique values + colors

---

| Option | Description | Selected |
|--------|-------------|----------|
| No — hidden ISPs stay hidden even inside radius | Consistent with filter intent | ✓ |
| Yes — clicking reveals all neighbors regardless of legend | Radius overrides legend | |

**User's choice:** No — hidden ISPs stay hidden even inside the radius

---

## Claude's Discretion

- Exact CSS styling of the top-left control panel
- Step size for the radius slider (suggested 0.1 mi)
- Exact color and opacity of the radius circle overlay
- Whether to show a "Click a company to see neighbors" hint in the panel when no selection is active
- Toggle-deselect: whether clicking the selected marker again also clears (in addition to map background click)

## Deferred Ideas

- Neighbor count in popup header ("3 neighbors within 0.5 miles") — INT-V2-01
- "Same ISP only" global toggle mode — INT-V2-02
- Inline Leaflet JS/CSS for offline use — OFF-V2-01
- Manually-tagged attribute columns convention — deferred to v2
