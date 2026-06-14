# Smart Material Estimator & Planner — UI Technical Specification
**Version 3 · Generated 2026-06-14**

> **Purpose of this document:** A complete, designer-ready blueprint of every screen, widget, layout structure, state variable, and data flow in the application. Nothing is omitted. A UI designer should be able to reconstruct the full functional experience from this document alone.

---

## Table of Contents

1. [Application Overview](#application-overview)
2. [Pre-Auth: Login Gate](#pre-auth-login-gate)
3. [Global Elements & Sidebar](#global-elements--sidebar)
4. [Sticky Header](#sticky-header)
5. [Tab 0 · Dashboard](#tab-0--dashboard)
   - [Sub-view A: Project Overview](#sub-view-a--project-overview)
   - [Sub-view B: Material Requirement & Procurement](#sub-view-b--material-requirement--procurement)
6. [Tab 1 · Selective Equipment Entry](#tab-1--selective-equipment-entry)
7. [Tab 2 · Session Order Report](#tab-2--session-order-report)
8. [Tab 3 · Location Report](#tab-3--location-report)
   - [Sub-view: Location Based](#sub-view-location-based)
   - [Sub-view: All Equipment](#sub-view-all-equipment)
9. [Tab 4 · Execution Plan](#tab-4--execution-plan)
   - [Sub-view: Execution Plan](#sub-view-execution-plan-1)
   - [Sub-view: Progress List](#sub-view-progress-list)
10. [Tab 5 (tab_consume) · Inventory](#tab-5-tab_consume--inventory)
    - [Sub-mode: Main Inventory](#sub-mode-main-inventory)
    - [Sub-mode: Consumption](#sub-mode-consumption)
    - [Sub-mode: Receipts](#sub-mode-receipts)
    - [Sub-mode: Ordered](#sub-mode-ordered)
11. [Tab 6 (tab5) · Total Overview](#tab-6-tab5--total-overview)
12. [Tab 7 (tab_master) · Master Data](#tab-7-tab_master--master-data)
13. [Data & Backend Flow](#data--backend-flow)
14. [Reusable Component Library](#reusable-component-library)
15. [Design Token Reference](#design-token-reference)

---

## Application Overview

| Property | Value |
|---|---|
| Framework | Streamlit (Python) |
| Page layout | `wide` |
| Sidebar initial state | `expanded` |
| Page title | "Smart Material Estimator & Planner" |
| Page icon | (default) |
| Primary font | `Inter` (body) + `JetBrains Mono` (metrics, tabs, monospaced labels) |
| Theme | Adaptive — fully supports both Streamlit Light and Dark mode via CSS variables |
| Database | SQLite (`sme_database.db`). Falls back to Excel files if DB absent |
| Locations | Three hardcoded: `Brown Field`, `TRAIN J`, `TRAIN K` |

The app runs as a single-page Streamlit app with **8 tabs**, a persistent **sidebar**, a **sticky fixed header**, and a **login gate** before any content renders.

---

## Pre-Auth: Login Gate

**Triggered when:** `st.session_state["_authenticated"]` is `False` or missing.

**Layout:**

```
Row:  [1 col spacer] [1 col center content] [1 col spacer]
```

**Content (center column):**

1. `st.image(LOGO_PATH, width=200)` — company logo (conditional on file existing)
2. HTML block: App title ("🏗 Smart Material Estimator") + subtitle ("Please log in to continue")
3. Three-column layout (`[1, 1, 1]`) — form rendered in the middle column:
   - `st.text_input("Username", key="_login_user", placeholder="Enter username")`
   - `st.text_input("Password", type="password", key="_login_pass", placeholder="Enter password")`
   - `st.button("🔐  Login", use_container_width=True, key="_login_btn")`

**Conditional Logic:**

- On button click: checks `user == "admin"` and `pwd == "admin2026"`.
- **Success:** sets `st.session_state["_authenticated"] = True` → `st.rerun()`.
- **Failure:** `st.error("❌ Invalid credentials. Please try again.")`

**State written:** `_authenticated` (bool)

After login, `st.stop()` prevents the rest of the app from rendering. All subsequent content only renders when `_authenticated == True`.

---

## Global Elements & Sidebar

### Sidebar

The sidebar is persistent across all tabs and renders `st.sidebar` context.

**Layout (top to bottom):**

1. **Logo:** `st.image(LOGO_PATH, width=140)` — conditional on file existing.

2. **App Title Block (HTML):**
   - Large text: "🏗 SME" in amber (`#F59E0B`)
   - Subtitle: "Smart Material Estimator v3" (uppercase, spaced, muted)

3. **Section: 📍 Project Overview**
   - Custom `sec-hdr` divider label
   - Iterates over `LOCATION_ORDER = ["Brown Field", "TRAIN J", "TRAIN K"]`
   - For each location: displays a color-coded location badge (`.loc-bf`, `.loc-tj`, `.loc-tk`) + equipment count (e.g., "14 equip.") in a flex row

4. **Section: 📦 Inventory**
   - Custom `sec-hdr` divider
   - `st.caption(f"📦 {len(inv)} materials  ·  ⚠️ {(inv['Available_Qty']==0).sum()} at zero stock")`
   - Displays total material count and zero-stock count inline

5. **Section: 📋 Session**
   - Custom `sec-hdr` divider
   - If `session_tags` is not empty: lists each tag with `st.caption(f"  · {t}")` 
   - `st.button("🗑 Clear Session", key="clear_sidebar")` — clears `session_tags` and reruns
   - If empty: `st.caption("No equipment added yet.")`

6. **Color Legend (HTML):**
   - "🟢 100%  🟠 90–99%  🟡 80–89%  🔴 <80%"
   - Explains the fulfillment color-coding used throughout the app

**State read:** `st.session_state.session_tags`
**State written:** `st.session_state.session_tags = []` (Clear Session button)

---

### Sticky Header

A fixed-position HTML block injected via `st.markdown(..., unsafe_allow_html=True)` that stays at the top of the viewport during scrolling.

**Content:**
- Left: Company logo (base64 inline image, 34px height) + App name ("Smart Material Estimator & Planner") + tagline ("System-code level · Cascading allocation · Priority-based")
- Right: Version badge ("v3" in amber pill style) + green status dot (always "online")

**CSS behavior:**
- `position: fixed; top: 0; left: 0; right: 0; z-index: 999990`
- Shifts right when sidebar is open (uses CSS `:has()` selector targeting sidebar `aria-expanded`)
- Amber gradient top border (2px)
- Backdrop blur effect

**Tab bar behavior:**
- The tab list (`[data-testid="stTabs"] > div:first-of-type`) is `position: sticky; top: 78px` so it stays just below the sticky header while scrolling.

**Custom Hamburger (JS injected):**
- A JS script injects a custom `☰` button (40×40px, amber border) fixed at `top: 14px; left: 16px` that proxies clicks to Streamlit's native sidebar toggle button.
- Survives Streamlit re-renders via a `MutationObserver`.

---

### Global Session State Variables

| Key | Type | Default | Description |
|---|---|---|---|
| `_authenticated` | bool | `False` | Login gate state |
| `session_tags` | list[str] | `[]` | Equipment tags added to the current session (shared across tabs 1, 2, 4) |
| `_session_key` | str (UUID4) | auto-generated | Unique identifier for the current browser session; used to partition `draft_consumption` DB rows |
| `loc_order` | dict[str, list[str]] | `{}` | Per-location equipment priority orders (used in Tab 3) |
| `all_eq_order` | list[str] | all tags | Global order for "All Equipment" mode in Tab 3 |
| `_ord_draft` | list[dict] | — | Temporary order draft records pending submission (Tab 5 / Ordered) |
| `_ord_source_detail` | str | — | Label of the source selection for the current order draft |

---

## Tab 0 · Dashboard

**Tab label:** `📊  Dashboard`
**Purpose:** High-level project health overview with filtering. Two sub-views switchable via a radio button.

### Shared Filter Bar (applies to both sub-views)

Located immediately below the sub-view radio, above all content.

**Layout:** `sec-hdr` label + 4-column row (`st.columns(4)`)

| Column | Widget | Key | Options | Default |
|---|---|---|---|---|
| 1 | `st.multiselect(" Location")` | `dash_loc` | `["Brown Field", "TRAIN J", "TRAIN K"]` | All 3 selected |
| 2 | `st.multiselect(" Type")` | `dash_type` | All unique `Type` values from `eq_master` (sorted) | All selected |
| 3 | `st.multiselect(" System Code")` | `dash_code` | All unique `(Lining_System_Code, Lining_System_Short_Name)` pairs formatted as `"Code N – Name"` | All selected |
| 4 | `st.multiselect(" Substrate")` | `dash_substrate` | All unique `Substrate` values from `eq_master` (sorted, dropna) | All selected |

**Sub-view toggle:**
- `st.radio("View", ["📈 Project Overview", "🛒 Material Requirement & Procurement"], horizontal=True, key="dash_view", label_visibility="collapsed")`
- Placed above the filter bar
- `st.markdown("<hr>")` separator below radio

---

### Sub-view A · Project Overview

**Purpose:** Bird's-eye KPI + coverage charts for all filtered equipment.

#### KPI Strip — 7 Clickable Metric Popover Cards

Layout: `st.columns(7)`

Each uses the `dbl_click_metric()` component — a styled `st.popover` button that looks like a metric card. Clicking opens a drill-down `st.dataframe` inside the popover.

| Column | Label | Value | Popover Title | Drill-down Data |
|---|---|---|---|---|
| k1 | Equipment | Count of filtered tags | "Equipment List" | `[Equipment_Tag_No., Name, Location, Type, Substrate]` |
| k2 | Total SQM | Sum of remaining SQM | "SQM by Equipment & System Code" | `[Equipment_Tag_No., Lining_System_Code, Total_SQM]` sorted desc |
| k3 | Available Coverage SQM | `Total_SQM × (Coverage% / 100)` | "Coverable SQM by Equipment & System Code" | Per-tag/code SQM coverage breakdown |
| k4 | SQM Deficit | `Total_SQM − Coverage_SQM` | "SQM Deficit by Equipment & System Code" | Rows where SQM Deficit > 0, sorted desc |
| k5 | Overall Coverage | `Allocated / Demand × 100 %` | "Coverable SQM by Equipment & System Code" | Same as k3 drill-down + delta string shown |
| k6 | Shortfall SQM | Same as k4 value | "SQM Deficit by Equipment & System Code" | Same as k4 drill-down |
| k7 | Critical (<50%) | Count of materials where Coverage < 50% | "Critical Materials (Coverage < 50%)" | Materials with Coverage_Pct < 50, sorted asc |

#### Row 1 — 2-Column Layout (`[1, 1.6]` ratio, gap="large")

**Left column (row1a):**

1. `sec-hdr`: "🎯 Overall Coverage"
2. `st.plotly_chart` — Plotly `go.Indicator` gauge chart:
   - Mode: `gauge+number+delta`
   - Value: Overall coverage % (0–100)
   - Delta reference: 100 (shows deficit/surplus)
   - Gauge color: green ≥100%, orange ≥90%, yellow ≥80%, red <80%
   - Background zones: 4 colored step ranges (0–50 red tint, 50–80 yellow, 80–90 orange, 90–100 green)
   - Title text: `"Coverage  ·  {can_sqm} / {proj_sqm} SQM Available Material Coverage"`
   - Height: 240px
   - Key: `"dash_gauge"`
3. `st.plotly_chart` — Mini stacked horizontal bar chart:
   - Two bars: "Available" (green) + "Shortfall" (red), stacked vertically
   - Orientation: vertical (y-axis = `["Inventory"]`)
   - Shows aggregate `f_total_avail` and `f_total_short` values as text labels
   - Height: 120px
   - Key: `"dash_dmini"`

**Right column (row1b):**

1. `sec-hdr`: "📍 Coverage by Location (SQM)"
2. `st.plotly_chart` — Grouped/stacked vertical bar chart (one bar group per location):
   - For each location: "Can Do" bar (location color) + "Deficit" bar (red)
   - Inside text: "Coverage%\nX SQM" for can-do, "X SQM deficit" for deficit
   - Height: 220px
   - Key: `"dash_loc_bar"`
3. Location stat mini-cards: `st.columns(len(loc_rows))` — one card per location showing:
   - Status dot emoji (🟢/🟠/🟡/🔴)
   - Location name in amber
   - Coverage % in large bold text
   - "Can / Total SQM" fraction
   - Equipment count

#### Row 2 — 2-Column Layout (equal, gap="large")

**Left column (row2a):**

1. `sec-hdr`: " Coverage by System Code (SQM)"
2. `st.plotly_chart` — Horizontal bar chart:
   - Y-axis: `"Code N  ShortName"` labels
   - X-axis: Coverage % (0–115 range)
   - Color coded: green/orange/yellow/red by %
   - Text inside bars: `"Coverage%  (Can SQM/Total SQM)"`
   - Dashed vertical line at x=100
   - Height: dynamic, `max(220, len(sc_df) * 42)`
   - Key: `"dash_sc_bar"`
3. `st.dataframe` — System code summary table:
   - Columns: Code, Short Name, SQM Total, Available Material Coverage (SQM), SQM Deficit, Coverage %
   - `use_container_width=True, hide_index=True`
   - Key: `"dash_sc_tbl"`

**Right column (row2b):**

1. `sec-hdr`: "🧪 Coverage by Material"
2. `st.plotly_chart` — Horizontal bar chart:
   - Y-axis: `"Material_Code  Material_Name[:18]"` labels
   - X-axis: Coverage % (0–115 range)
   - Color coded per row
   - Hover template shows Available, Demand, Shortfall with UOM
   - Height: dynamic, `max(340, len(mat_rows_d) * 34)`
   - Key: `"dash_mat_bar"`

#### Full-Width Material Balance Table

1. `sec-hdr`: "📋 Full Material Balance"
2. `st.dataframe` — Color-coded styled dataframe:
   - Columns: Code, Material Name, UOM, Available, On Order, Total Demand, Shortfall, Net Shortfall, Coverage %
   - Row color: green/orange/yellow/red by Coverage %, Coverage % cell text also colored
   - Sorted ascending by Coverage %
   - Height: `50 + len(tbl_show) * 35`
   - Key: `"dash_mat_tbl"`

#### Stock-Only Materials Table (conditional)

Shown only when materials in inventory have no matching recipe demand:

1. `sec-hdr`: "📦 Stock-Only Materials (No Demand in Any System Code)"
2. `st.caption`: explanatory text
3. `st.dataframe` — Columns: Code, Material Name, UOM, Available, On Order
   - Key: `"dash_stock_only"`

#### Download Button

- `st.columns(2)` — button in left column only:
  - `st.download_button("⬇ Download Material Balance", ...)` → `dashboard_material_balance.xlsx` (color scheme: "dashboard")

---

### Sub-view B · Material Requirement & Procurement

**Purpose:** Procurement-oriented breakdown, organized by location then by system code, with per-expander material tables and a grand total procurement list.

#### KPI Strip — 6 Clickable Metric Popover Cards

Layout: `st.columns(6)`

| Column | Label | Value | Popover Drill-down |
|---|---|---|---|
| p1 | Equipment | Filtered tag count | Equipment list |
| p2 | Total SQM | Remaining SQM | SQM per tag/code |
| p3 | Available Coverage SQM | Coverable SQM | Per-tag/code SQM coverage |
| p4 | SQM Deficit | SQM uncoverable | Per-tag/code SQM deficit |
| p5 | Shortfall Units | Sum of all material shortfall units | Materials with Shortfall > 0 |
| p6 | After Orders (Net) | Sum of net shortfall (after subtracting Ordered_Qty) | Materials with Net Shortfall > 0 |

#### Location Sections (one per selected location, iterated)

For each location (`Brown Field`, `TRAIN J`, `TRAIN K`):

1. **Location row (HTML):** Status dot + color-coded badge + `"N equip · X/Y SQM · Z%"` metadata line

2. **Per System Code Expanders** (one per system code within that location):
   - Label: `"{dot}  Code N – Name  ·  X/Y SQM  ·  Z%"` 
   - Default: `expanded=False`
   - **Inside expander — 5 metrics (`st.columns(5)`):**
     - `st.metric("System Code", f"Code {code}")`
     - `st.metric("Short Name", sname)`
     - `st.metric("SQM Total", f"{code_sqm:,.2f}")`
     - `st.metric("Available Material Coverage (SQM)", f"{c_can_sqm:,.2f}")`
     - `st.metric("SQM Deficit", f"{c_sh_sqm:,.2f}")`
   - **Styled material table (`st.dataframe`):**
     - Columns: Code, Material Name, UOM, Available, On Order, Demand, Shortfall, Net Shortfall (After Orders), Fulfil %
     - Row coloring by Fulfil %, Fulfil % cell text also colored
     - Height: `65 + len(tbl_proc) * 35`
     - Key: `f"proc_{loc}_{code}"`

3. Thin horizontal divider between locations

#### Grand Total Procurement Table

1. `sec-hdr`: "📦 Grand Total — All Selected Equipment"
2. `st.dataframe` — styled table (same structure as per-code tables above, but aggregated across all locations/codes):
   - Columns: Code, Material Name, UOM, Available, On Order, Demand, Shortfall, Net Shortfall, Coverage %
   - Height: `50 + len(grand_show) * 35`
   - Key: `"proc_grand_tbl"`

#### Download Buttons — `st.columns(2)`

- Left: `st.download_button("⬇ Download Grand Procurement Table", ...)` → `procurement_grand_total.xlsx`
- Right (conditional — only if Net Shortfall rows exist): `st.download_button("⬇ Net Order List Only", ...)` → `net_order_list.xlsx`

---

## Tab 1 · Selective Equipment Entry

**Tab label:** `🔍  Selective Equipment Entry`
**Purpose:** Search for and select individual equipment tags, view their material requirements by system code, and build a prioritized session list.

**Layout:** Two-column split — `st.columns([1, 1.65], gap="large")`

---

### Left Panel

#### Filter Equipment Section

`sec-hdr`: "🎛 Filter Equipment"

| Widget | Key | Options | Default |
|---|---|---|---|
| `st.multiselect(" Location")` | `t1_loc` | `LOCATION_ORDER` | `[]` (empty — shows all) |
| `st.multiselect(" Type")` | `t1_type` | All unique Type values (sorted) | `[]` |
| `st.multiselect(" System Code")` | `t1_code` | All unique Lining_System_Code values (sorted as int), displayed as `"Code N – Name"` | `[]` |

Filters are AND-combined. Empty = no filter applied for that dimension.

#### Find Equipment Section

`sec-hdr`: "🔍 Find Equipment"

- `st.selectbox("tag_search", ...)` — key: `"tag_select"`, label hidden
  - Options: `[""] + filtered_tags_t1`
  - `format_func`: shows `"{tag}  —  {Name}"` for non-empty selections
  - Default: empty string `""`
- `st.columns([2, 1])`:
  - Left: `st.button("＋ Add to Session", key="add_btn", disabled=(tag=="" or tag already in session))`
  - Right: Shows `"✓ In session"` (green HTML) if tag is already in session

**On "Add" click:** Appends `selected_tag` to `st.session_state.session_tags` → `st.rerun()`

#### Session Priority List Section

`sec-hdr`: "📋 Session Priority List"

**When session is empty:** `st.info("Add equipment tags above to build your session.")`

**When session has items:**

1. `st.caption("⠿ Drag to reorder — order is applied instantly.")`
2. **`sort_items()` drag-and-drop component** (from `streamlit_sortables`):
   - Items: `["#{i+1}  ||  {tag}  ||  {Name[:28]}  ||  {Location}"]` per tag
   - Direction: `"vertical"`
   - Key: `"sess_sort_" + "_".join(session_tags)` — regenerated when items change
   - Reordering auto-applies via `st.rerun()` when order changes
3. **Per-tag rows** (iterated over `new_order`):
   - `st.columns([9, 1])` per tag:
     - Left (9): HTML card showing — tag (with status dot prefix), name, system code chips (`"C{N}"`), location badge, SQM fraction (`"can_sqm/total_sqm SQM"`), fulfillment pill
     - Right (1): `st.button("✕", key=f"rm_{tag}_{idx}")` — removes tag from session
4. `st.button("🗑 Clear All", key="clear_all")` — clears entire session

**State read/written:** `st.session_state.session_tags`

---

### Right Panel

**When no tag selected:** Centered placeholder message "SELECT AN EQUIPMENT TAG TO VIEW DETAILS"

**When tag selected:** Equipment info card + per-system-code material tables

#### Equipment Info Card (amber left-border card, HTML)

Displays:
- "Equipment Tag" label + tag value (amber, large)
- Location badge (top-right)
- Equipment Name (large)
- 3-column metadata grid: Type, Substrate, Material Spec.
- Full-width "Lining Systems" text (multiline)

#### System Code Material Requirements Section

`sec-hdr`: "⚗️ System Code Material Requirements"

For each system code the equipment has (sorted numerically):

**`st.expander(f"System Code {code}  ·  {sname}  ·  {sqm} SQM  ·  Coverage: {pct:.1f}%", expanded=False)`**

Inside each expander — **4 clickable metric cards (`st.columns(4)`):**

| Card | Label | Value | Popover Shows |
|---|---|---|---|
| mi1 | System Code | Code N | Material breakdown |
| mi2 | Short Name | Short name string | Material breakdown |
| mi3 | Surface Area | `"{sqm:,.2f} SQM"` | Material breakdown |
| mi4 | Coverage | `"{pct:.1f}%"` | Coverage detail |

Then: **`plotly_mat_table()`** color-coded material table:
- Columns: Code, Material Name, UOM, Demand, Allocated, Shortfall, Fulfil %
- Row coloring by Fulfil %
- Height: `65 + len(mat_rows) * 30`
- Key: `f"entry_{selected_tag}_{code}"`

#### Equipment Grand Total Box (amber gradient box, HTML)

A `grand-box` styled div showing a 4-cell grid:
- System Codes count
- Total Demand (summed units, no UOM label)
- Total Shortfall (red if > 0, green if 0)
- Coverage % (color-coded by threshold)

---

## Tab 2 · Session Order Report

**Tab label:** `📦  Session Order Report`
**Purpose:** Full cascading allocation report for the session's equipment list, with drag-to-reorder priority, per-equipment breakdowns, combined procurement list, and export.

**When session is empty:** `st.info("Add equipment tags in the Entry tab to generate a report.")`

---

### KPI Strip — 5 Clickable Metric Popover Cards

Layout: `st.columns(5)`

| Col | Label | Value | Popover Drill-down |
|---|---|---|---|
| k1 | Equipment | Count of session tags | Session equipment list (Tag, Name, Location) |
| k2 | Materials | Unique material count across session | Material Demand Summary (all materials) |
| k3 | Need to Order | Materials with Shortfall > 0 | Materials to Procure (shortfall > 0 only) |
| k4 | Total Shortfall | Sum of all shortfall units | Same as k3 |
| k5 | Overall Coverage | `Allocated / Demand × 100%` | Coverage by Material (sorted asc) |

---

### Priority Drag-to-Reorder List

1. `sec-hdr`: "⠿ Drag to Reorder Priority — changes reflect everywhere"
2. `st.caption`: instruction text
3. **`sort_items()` component** — same structure as Tab 1 but keyed `"t2_sort_" + "_".join(session_tags)`
   - Reordering writes to `st.session_state.session_tags` → `st.rerun()`

---

### Per-Equipment Expanders

`sec-hdr`: "Per-Equipment System Code Breakdown"

For each tag in session (in current priority order):

**`st.expander(f"{dot}  #{i+1}  {tag}  ·  {name}  ·  {type|substrate}  ·  {loc}  ·  {can_sqm}/{sqm} SQM  ·  {pct:.1f}%", expanded=False)`**

Inside:
- **4-column metadata strip** (`st.columns(4)`):
  - `m1.markdown(f"**Type:** {type}")`
  - `m2.markdown(f"**Substrate:** {substrate}")`
  - `m3.markdown(f"**Material Spec.:** {spec}")`
  - `m4.markdown(f"**Total SQM:** `{sqm:,.2f}`")`
- `st.caption(f"**Lining Systems:** ...")` (pipe-separated)
- `st.markdown("---")` divider

**Per System Code (within expander, iterated):**

1. HTML `syscode-block` header row: dot + `code-badge` + short name + `"can/total SQM"` + fulfillment pill (float right)
2. **4 clickable metric cards (`st.columns(4)`):**
   - Demand (total units for this code): popover shows material detail
   - Allocated: same popover
   - Shortfall (conditional — only shown if > 0.001): popover shows shortfall-only rows
   - SQM Deficit (conditional — same condition): same popover
3. **`plotly_mat_table()` with `show_sqm=True`:**
   - Columns: Code, Material Name, UOM, Demand, Allocated, Shortfall, Fulfil %, SQM Total, SQM Done, SQM Deficit, SQM Done %
   - Height: `65 + len(code_alloc) * 30`
   - Key: `f"rep_{tag}_{code}"`

**Equipment Grand Total Bar (HTML, bottom of each expander):**
- Amber-tinted div: `"GRAND TOTAL — {tag}"` + Demand + Allocated + Shortfall (conditional) + fulfillment pill

---

### Combined Procurement List

`sec-hdr`: "🛒 Combined Procurement List"

**Shortage stacked bar chart** (shown only when materials have shortfall):
- `st.plotly_chart` — horizontal stacked bar:
  - Y-axis: `"Material_Code  Material_Name[:22]"` per material with shortfall
  - Two traces: "Available" (green) + "To Order" (red)
  - Height: `max(300, len(shortage_only) * 42)`
  - Key: `"session_bar"`

**`plotly_mat_table(combined)`** — full combined table with SQM columns:
- Columns: Code, Material Name, UOM, Demand, Allocated, Shortfall, Fulfil %, SQM Total, SQM Done, SQM Deficit, SQM Done %
- Sorted by Fulfillment_Pct ascending
- Height: `90 + len(combined) * 30`
- Key: `"session_combined"`

**Grand Total Box (HTML `grand-box`):**
5-cell grid: Equipment count, Materials count, Total Demand, To Procure count (red), Shortfall Units (red)

---

### Download Buttons — `st.columns(2)`

- Left: `st.download_button("⬇ Full Session Report", ...)` → `session_full_report.xlsx` (color: "session")
- Right (conditional — if shortage exists): `st.download_button("⬇ Order List Only", ...)` → `order_list.xlsx`

---

### Smart Reordering Suggestions Panel

Rendered at the bottom via `render_suggestion_panel(session_tags, "tab2")`.

- `sec-hdr`: "💡 Smart Reordering Suggestions"
- `st.caption`: explanation of what suggestions show
- Spinner during computation: `st.spinner("Analysing reorder scenarios…")`
- Hidden completely when no improvements found (> 0.4% gain threshold)
- **`st.columns(2, gap="large")`:**
  - **Left — "By Equipment":** Up to 5 suggestion cards (HTML `.card`), each showing tag + name, current→suggested position, current→new %, gain
  - **Right — "By System Code":** Up to 8 suggestion cards, each showing Code + short name + tag, move, gain, optional "100% COMPLETE" green badge

---

## Tab 3 · Location Report

**Tab label:** `📍  Location Report`
**Purpose:** View all equipment organized by location (or globally) with per-location, per-equipment, per-system-code cascading allocation, each with its own drag-to-reorder priority.

**Sub-view toggle (at top):**
```
st.radio("View Mode", ["📍 Location Based", "🌐 All Equipment"],
         horizontal=True, key="loc_report_mode", label_visibility="collapsed")
```
Then `st.markdown("<hr>")`.

---

### Sub-view: Location Based

**Purpose:** Shows equipment grouped by location. Each location has an independent drag-to-reorder list and its own cascade allocation.

`sec-hdr`: "📍 All Equipment by Location — Cascading Balance"
`st.caption`: drag instruction text

**For each location in `LOCATION_ORDER`:**

1. **Amber-colored location sub-header (HTML):** "📍 {loc} — Drag to set build priority"

2. **`sort_items()` drag-to-reorder** — per-location:
   - Items: `["#{i+1}  ||  {tag}  ||  {Name[:26]}  ||  {SQM:,.1f} SQM"]`
   - Key: `"loc_sort_{loc}_" + "_".join(loc_tags)`
   - On reorder: writes to `st.session_state.loc_order[loc]` → `st.rerun()`

3. **Location header row (HTML):**
   - Status dot + location color-badge + equipment count + `"can/total SQM"` + Coverage %

4. **Per-equipment `st.expander`s** (each collapsed by default):
   - Label: `"{dot}  {tag}  ·  {Name}  ·  {type|substrate}  ·  {can/total SQM}  ·  {pct:.1f}%"`
   - Inside:
     - 3-column metadata: `st.columns(3)` → Type, Substrate, Material Spec.
     - `st.caption("**Lining:** ...")` 
     - If 100%: green success banner ("✅ All materials fully covered — ready to proceed")
     - `st.markdown("---")`
     - Per system code: HTML `syscode-block` header + `plotly_mat_table(..., show_sqm=True, allocated_label="Available")`
     - Equipment grand total bar (HTML)
     - "Add to Session" button / "✓ In session" text

5. **`st.expander(f"📊 Show Shortfall Chart — {loc}", expanded=False)`:**
   - Inside: `st.plotly_chart` — horizontal stacked bar chart (Allocated + Shortfall per `tag/code` pair)
   - Y-axis labels: `"{tag}\nCode {code} ({sname})"`
   - Color: location accent color for allocated, red for shortfall
   - Height: `max(350, len(cdf) * 36)`
   - Key: `f"loc_chart_{loc}"`

6. Thin `div` bottom divider between locations

#### Downloads Section

`sec-hdr`: "📥 Download Report per Location"

`st.columns(len(LOCATION_ORDER))` — one download button per location:
- Each: `st.download_button(f"⬇ {loc}", ...)` → `location_report_{loc}.xlsx` (per-location color scheme)

Then: `st.download_button("⬇ All Locations — Combined (Multi-Sheet)", ...)` → `location_report_all_{date}.xlsx` (multi-sheet workbook)

#### Print Button (HTML)

```html
<button onclick="window.print()">🖨 Print / Save as PDF</button>
```
With `@media print` CSS that hides sidebar, header, sticky header, tab bar, and the button itself during printing.

#### Smart Reordering Suggestions (per location)

For each location with ≥ 2 equipment:
- `st.expander(f"💡 Smart Reordering Suggestions — {loc}", expanded=False)`
- Inside: `render_suggestion_panel(loc_tags, f"tab3_{loc}")`

---

### Sub-view: All Equipment

**Purpose:** Treat all equipment in the project as a single global pool with one shared cascade allocation.

`sec-hdr`: "🌐 All Equipment — Global Cascading Balance"
`st.caption`: "All equipment in file order. Inventory pool is shared globally across all locations."

1. `st.caption("⇅ Drag to reorder — order determines cascade priority.")`
2. **`sort_items()` global drag-to-reorder:**
   - Items: `["#{i+1}  ||  {tag}  ||  {Name[:26]}  ||  {Location}  ||  {SQM:,.1f} SQM"]`
   - Key: `"ae_sort_" + str(len(all_eq_tags))`
   - On reorder: writes to `st.session_state.all_eq_order` → `st.rerun()`

3. **Global KPI strip — `st.columns(5)` (plain `st.metric`, not popovers):**
   - `ae_k1.metric("Equipment", str(count))`
   - `ae_k2.metric("Total SQM", f"{sqm:,.1f}")`
   - `ae_k3.metric("Available Coverage SQM", f"{can_sqm:,.2f}")`
   - `ae_k4.metric("SQM Deficit", f"{deficit:,.2f}")`
   - `ae_k5.metric("Overall Coverage", f"{pct:.1f}%")`

4. `sec-hdr`: "Per-Equipment Detail"

5. **Per-equipment `st.expander`s** (same structure as Location Based mode but uses `ae_alloc`):
   - `plotly_mat_table(..., allocated_label="Available")`
   - Each expander has "Add to Session" button or "✓ In session" indicator

6. **`st.expander("💡 Smart Reordering Suggestions — All Equipment", expanded=False)`:**
   - `render_suggestion_panel(all_eq_tags, "tab3_all")`

7. `st.download_button("⬇ Download All Equipment Report", ...)` → `all_equipment_report_{date}.xlsx` (color: "overview")

---

## Tab 4 · Execution Plan

**Tab label:** `⚙️  Execution Plan`
**Purpose:** Critical system code analysis per equipment, showing what to order first; also includes a project progress list view.

**Sub-view toggle:**
```
st.radio("View", ["⚙️ Execution Plan", "📋 Progress List"],
         horizontal=True, key="exec_subview", label_visibility="collapsed")
```
Then `st.markdown("<hr>")`.

---

### Sub-view: Execution Plan

**When session is empty:** `st.info("Add equipment tags in the Entry tab first.")`

**When session has items:**

#### Equipment Selector

```python
st.selectbox("Select Equipment",
    options=session_tags,
    format_func=lambda t: f"{t}  —  {name}",
    key="exec_tag")
```

#### System Code Selector

```python
st.selectbox("Select Critical System Code",
    options=avail_codes,  # sorted from current equipment's allocation
    format_func=lambda c: f"Code {c}  —  {short_name}",
    key="exec_code")
```
**When no system codes exist:** `st.warning("No system code data for this equipment.")`

---

#### Critical System Code Card (amber left-border card, HTML)

Displays:
- "Critical System Code" label
- `code-badge` chip + short name + SQM + large coverage % (color-coded)
- Explanatory sentence:
  - If no shortfall: "✅ All materials for this system code are fully covered."
  - If shortfall: "⚠️ X units short across N material(s) — order these first to proceed."

#### All Materials Status Table

- `sec-hdr`: "All Materials — Status Overview"
- `st.caption`: showing row count, system code count, priority position
- `plotly_mat_table(tag_alloc, ...)`:
  - Columns: Code, Material Name, UOM, Demand, Allocated, Shortfall, Fulfil %
  - Height: `80 + len(tag_alloc) * 30`
  - Key: `f"exec_all_{sel_tag}"`

#### Procurement Order Priority Section

`sec-hdr`: "📋 Procurement Order Priority"

**1️⃣ Critical Code Block (red-bordered HTML div):**
- Label: "Order First — System Code {code} ({name}) · Critical Path"
- Text: either "No shortages on critical system code ✅" or "N material(s) need to be procured"
- If shortfalls: `plotly_mat_table(crit_short_df, ...)` — height: `65 + N*30`

**2️⃣+ Other System Codes (amber-bordered HTML div, one per other code):**
- Label: "Order Next — System Code {code} ({name}) · Coverage: {pct:.1f}%"
- Text: "All materials covered ✅" or "N material(s) short. Order after critical code is secured."
- If shortfalls: `plotly_mat_table(code_short_display, ...)` — height: `65 + N*30`

#### Execution Summary Box (amber gradient HTML `grand-box`)

Displays a textual execution summary:
- Which critical code, what % coverage
- How many critical materials to order and total units
- Additional other-code shortfall summary
- "Total to order for full completion: X units across N material(s)"

#### Download Button (conditional — only if shortfalls exist)

`st.download_button(f"⬇ Download Execution Order List — {sel_tag}", ...)` → `execution_plan_{tag}.xlsx` (color: "execution")

---

### Sub-view: Progress List

**Requires DB:** `st.warning("Database required for the Progress List.")` if DB not found.

#### Project Progress KPIs — `st.columns(4)` (plain `st.metric`)

- `pk1.metric("Total SQM", f"{tot_orig:,.2f}")`
- `pk2.metric("Done SQM", f"{tot_done:,.2f}")`
- `pk3.metric("Remaining SQM", f"{tot_rem:,.2f}")`
- `pk4.metric("Completion", f"{tot_pct:.1f}%")`

#### Filter Row — `st.columns(2)`

| Widget | Key | Options |
|---|---|---|
| `st.selectbox("Filter Location")` | `prog_loc_f` | `["All"] + sorted unique locations` |
| `st.selectbox("Filter Status")` | `prog_status_f` | `["All", "✅ Complete", "🔄 In Progress", "⏳ Not Started"]` |

Status is derived: ≥100% → "✅ Complete", >0% → "🔄 In Progress", 0% → "⏳ Not Started".

#### Progress Table

`st.dataframe` (styled):
- Columns: Equipment Tag, System Code, System Name, Location, Equipment Name, Total SQM, Done SQM, Remaining SQM, Completion %, Status
- Row coloring by Completion %: green ≥100%, amber >0%, red =0%
- `Completion %` cell text also colored
- Height: `min(700, 60 + N*35)`
- Key: `"prog_list_tbl"`

#### Download

`st.download_button("⬇ Download Progress List", ...)` → `progress_list_{date}.xlsx` (color: "overview")

---

## Tab 5 (tab_consume) · Inventory

**Tab label:** `📦  Inventory`
**Purpose:** Full ERP-style inventory management — view stock movements, record daily consumption, log material receipts, and manage purchase orders.

**DB guard:** If DB not found: `st.error("Database not found. Run python setup_db.py first.")` + `st.stop()`

**Sub-mode toggle:**
```
st.radio("Mode", ["📊 Main Inventory", "📅 Consumption", "📦 Receipts", "📋 Ordered"],
         horizontal=True, key="inv_mode", label_visibility="collapsed")
```
Then `st.markdown("<hr>")`.

---

### Sub-mode: Main Inventory

`sec-hdr`: "📊 Main Inventory Dashboard — Movements by Date Range"

#### Filter Row — `st.columns(3)`

| Widget | Key | Options | Default |
|---|---|---|---|
| `st.date_input("From Date")` | `mi_from` | Any date | Today − 30 days |
| `st.date_input("To Date")` | `mi_to` | Any date | Today |
| `st.radio("Group By", ["By Material", "By System Code"])` | `mi_view` | Two options | "By Material" |

#### By Material View

`st.dataframe` (styled):
- Columns: Code, Material Name, UOM, Total Receipts, Total Consumed, Current Stock, Ordered Qty
- Row + Current Stock cell coloring: red ≤0, amber <50, green ≥50
- All qty columns formatted `{:,.3f}`
- Key: `"mi_mat_tbl"`

`st.download_button("⬇ Download Inventory Dashboard", ...)` → `inventory_dashboard_{date}.xlsx`

#### By System Code View

For each system code (sorted as int):
- `st.expander(f"Code {code} — {name}  ·  {N} materials  ·  {pct:.0f}% stock coverage", expanded=False)`
- Inside: `st.dataframe` with same columns as By Material view
- Key: `f"mi_sc_{code}"`

---

### Sub-mode: Consumption

`sec-hdr`: "Step 1 — Select Work Location & Equipment"

#### Cascading Dropdowns — `st.columns(4)`

All four form a cascading chain (each filters the next):

| Col | Widget | Key | Options | Behavior |
|---|---|---|---|---|
| 1 | `st.selectbox(" Location")` | `ce_loc` | `[""] + LOCATION_ORDER` | Filters Types and Tags |
| 2 | `st.selectbox(" Type")` | `ce_type` | `[""] + types filtered by location` | Filters Tags |
| 3 | `st.selectbox(" Equipment Tag")` | `ce_tag` | `[""] + tags filtered by loc+type`, displayed as full label | Filters System Codes |
| 4 | `st.selectbox(" System Code")` | `ce_code` | `[""] + codes for selected tag`, displayed as `"Code N  –  Name"` | Unlocks Step 2 |

**When both `ce_tag` and `ce_code` are selected:** Steps 2–3 appear.

---

#### Step 2 — SQM Info Row (`st.columns(4)`, clickable metric popovers)

| Card | Label | Value | Popover shows |
|---|---|---|---|
| sc1 | System Code | `sname` | Recipe detail for this code |
| sc2 | Original SQM | `total_sqm_orig:,.2f` | Recipe detail |
| sc3 | Already Done SQM | `done_sqm_prev:,.2f` | Recipe with "Demand for Remaining SQM" column |
| sc4 | Remaining SQM | `remaining_sqm:,.2f` | Material needed for remaining SQM |

---

#### Step 2/3 — Consumption Form (`st.form(key="ce_form")`)

`sec-hdr`: "Step 2 — Enter SQM Completed Today"

**Two-column header row (`st.columns(2)`):**
- Left: `st.date_input("📅 Work Date", value=date.today(), key="form_ce_date")`
- Right: `st.number_input("SQM Completed Today", min_value=0.0, max_value=remaining_sqm, value=0.0, step=0.5, format="%.2f", key="form_ce_sqm")`

`sec-hdr`: "Step 3 — Material Quantities Consumed"
`st.caption`: guidance on defaults vs. actual override

**7-column header row (column labels, not widgets):**
Code | Material Name | UOM | Available | For 1 SQM | Actual Consumed | On Order

**Per material row (iterated, `st.columns([2,3,1,1.5,1.5,1.5,1.5])`):**
- `c1.markdown(f"<code>{mc}</code>")` — Material Code
- `c2.write(material_name)`
- `c3.write(uom)`
- `c4.write(f"{avail:,.3f}")` — read-only Available
- `c5.write(f"{for_1:,.3f}")` — read-only For 1 SQM
- `c6.number_input("qty", min_value=0.0, value=0.0, step=0.001, key=f"form_mat_{mc}")` — **editable actual consumed**
- `c7.write(f"{onord:,.3f}")` — read-only On Order

**Live variance preview** (outside form, above submit):
- If SQM > 0: for each material, computes expected vs. actual
  - Over consumption: `st.warning(f"⚠️ {mc}: Over Consumption (+{pct:.1f}%)")`
  - Less consumption: `st.info(f"ℹ️ {mc}: Less Consumption ({pct:.1f}%)")`

`st.text_area("📝 Notes (optional)", placeholder="Weather conditions, issues, remarks…", key="form_notes", height=70)`

`st.form_submit_button("➕ Add to Grid", use_container_width=False)`

**Outside form:**
`st.button("🧹 Clear Form", key="ce_clr_btn")` — clears all `form_ce_*`, `form_mat_*`, `form_notes` keys

---

#### Draft Consumption Grid

`sec-hdr`: "📋 Draft Consumption Grid — Pending Submission"

**When empty:** `st.info("No pending entries. Use 'Add to Grid' above to stage consumption data...")`

**When populated:** `st.data_editor(...)` — key: `"draft_ce_editor"`
- Columns: ☐ Del (checkbox), ID, entry_date, equipment_tag, lining_system_code, material_code, material_name, uom, expected_qty, effective_qty, variance_pct, variance_status, notes
- `num_rows="fixed"`, `hide_index=True`
- Row coloring by variance_status: amber = Over, blue = Less, green = OK
- Column configs: ID disabled, `☐ Del` = CheckboxColumn, `variance_status` = SelectboxColumn (OK/Over/Less)
- Height: `min(500, 55 + N*35)`

**Button row (`st.columns([2, 2, 4])`):**
- Left: `st.button("🗑️ Delete Selected", key="del_draft_rows")` — reads editor state for checked `☐ Del` rows → DELETE from DB
- Middle: `st.button("🧹 Clear All Draft", key="clear_draft_all")` — clears all draft rows for session key

`sec-hdr`: "✅ Submit All Draft Entries"

`st.button("✅ Submit Consumption", key="submit_draft_btn", type="primary")`:
- Writes all draft rows to `consumption_log`
- Deducts `effective_qty` from `inventory.available_qty`
- Increments `sqm_progress.done_sqm`
- Deletes draft rows
- Shows `st.success(f"✅ {N} entries submitted...")`
- Shows `st.download_button("⬇ Download Consumption Report", ...)` → `consumption_report_{date}.xlsx` (multi-sheet: "Full Day Summary" + one sheet per system code)

---

#### Consumption History Expander

`st.expander("📜 View & Edit Consumption History", expanded=False)`

Inside:
- `st.columns(2)` filter row:
  - `st.selectbox("Filter by Equipment", key="log_tag")` — "All" + unique equipment tags
  - `st.selectbox("Filter by Date", key="log_date")` — "All" + unique dates (newest first)
- `st.data_editor` — key: `"cons_log_editor"`:
  - Columns: Sl. No. (disabled), ☐ Select (checkbox), ID (disabled), Date (disabled), Equipment (disabled), Code (disabled), System, SQM Done, Material, Material Name, UOM, Expected Qty, Consumed Qty, Notes, Submitted At (disabled)
  - `num_rows="fixed"`, last 200 rows
  - Height: `min(600, 50 + N*35)`
- Button row (`st.columns([2, 2, 4])`):
  - `st.button("💾 Save Cell Edits", key="cons_log_save")` — UPDATE editable columns
  - `st.button("🗑️ Delete Selected", key="cons_log_del")` — DELETE checked rows
- `st.download_button("⬇ Download Consumption Log", ...)` → `consumption_log_{date}.xlsx`

---

### Sub-mode: Receipts

`sec-hdr`: "📦 Record Material Receipt — New Stock Received"
`st.caption`: guidance text

#### Order ID Pre-Selector

`st.selectbox("🔗 Link to Order ID / PR# (Optional — select first to filter materials)", key="rc_order_id")`:
- Options: `["— None —"] + list of open order IDs` (orders where `status != 'Fulfilled'` and pending qty > 0, ordered newest first)
- Selecting an order filters the Material selector below to that order's items

#### Material Selector

`st.selectbox("🧪 Select Material", key="rc_mat_sel")`:
- Options: filtered by selected order (if not "— None —") or all inventory materials
- `format_func`: `"{code}  —  {name[:35]}"`

**Auto-fill display (when material selected) — `st.columns(2)`):**
- `st.text_input("Current Available Qty", value=f"{avail:,.3f}", disabled=True, key="rc_avail_disp")`
- `st.text_input("Current Ordered Qty",   value=f"{ordered:,.3f}", disabled=True, key="rc_ord_disp")`

---

#### Receipt Form (`st.form(key="receipt_form")`)

**3-column row:**
- `st.date_input("📅 Receipt Date", value=date.today(), key="rc_date")`
- `st.number_input("Qty Received", min_value=0.0, value=0.0, step=1.0, format="%.3f", key="rc_qty")`
- `st.text_input("Notes / PO Ref.", placeholder="PO number, supplier…", key="rc_notes")`

**Dynamic extra fields** (from `receipt_log` DB schema PRAGMA, excluding fixed columns):
- Numeric keywords (qty/sqm/amount/value) → `st.number_input`
- Others → `st.text_input`
- Rendered in groups of 3 columns

`st.form_submit_button("✅  Record Receipt", use_container_width=False)`

**Outside form:** `st.button("🧹 Clear Form", key="rc_clr_btn")`

**On submit:**
- Inserts into `receipt_log`
- `UPDATE inventory SET available_qty += received, ordered_qty = MAX(0, ordered_qty - received)`
- If linked to order: updates `orders_log.fulfilled_qty` and `status` (Pending/Partial/Fulfilled)
- `st.success(f"✅ Received {qty} units of {name} added...")`

---

#### Receipt History Expander

`st.expander("📜 View & Edit Receipt History", expanded=False)`

Inside:
- `st.columns(2)` filter row:
  - `st.selectbox("Filter by Material", key="rlog_mat")`
  - `st.selectbox("Filter by Date", key="rlog_date")`
- `st.data_editor` — key: `"receipt_log_editor"`:
  - Columns: Sl. No. (disabled), ☐ Select (checkbox), ID (disabled), Date, Code (disabled), Material, UOM, Received Qty, Notes, Received At (disabled)
  - `num_rows="fixed"`, last 100 rows, `hide_index=True`
  - Height: `min(500, 50 + N*35)`
- Button row (`st.columns([2, 2, 4])`):
  - `st.button("💾 Save Cell Edits", key="rlog_save")` — UPDATE editable fields
  - `st.button("🗑️ Delete Selected", key="rlog_del")` — DELETE checked rows
- `st.download_button("⬇ Download Receipt Log", ...)` → `receipt_log_{date}.xlsx`

---

### Sub-mode: Ordered

`sec-hdr`: "📋 Order Management System"

**Sub-mode radio:**
```
st.radio("Generate From",
    ["📍 Location Needs", "⚙️ System Code Needs", "📜 View Past Orders"],
    horizontal=True, key="ord_mode")
```
Then `st.markdown("<hr>")`.

#### 📍 Location Needs

`st.selectbox("Select Location", ["All Locations"] + LOCATION_ORDER, key="ord_loc_sel")`

`st.button("🔍 Calculate Shortfalls", key="calc_ord_short")`:
- Runs `cascade_allocate()` for selected location's tags
- Computes shortfall per material
- Stores result in `st.session_state["_ord_draft"]` (list of dicts)
- Stores location label in `st.session_state["_ord_source_detail"]`

#### ⚙️ System Code Needs

`st.selectbox("Select System Code", key="ord_code_sel")`:
- Options: all unique lining system codes (sorted as int)
- `format_func`: `"Code N — Short Name"`

`st.button("🔍 Calculate Shortfalls", key="calc_ord_sc_short")`:
- Same logic as Location Needs but filtered to system code

#### 📜 View Past Orders

`st.selectbox("Select Order ID", key="view_past_ord")`:
- Options: all distinct order IDs from `orders_log` (newest first)

**When order selected:**
- PR# input row (`st.columns([3, 1])`):
  - `st.text_input("🔖 PR# (Purchase Request Number)", key="pr_input")` — prefilled with existing value
  - `st.button("💾 Update PR#", key="update_pr_btn")` — saves to `orders_log`
- `st.dataframe` (styled by Status) — key: `"ord_status_tbl"`:
  - Columns: Order ID, PR#, Material Code, Material Name, UOM, Ordered Qty, Fulfilled Qty, Pending Qty, Status, Notes
  - Row coloring: green = Fulfilled, amber = Partial, red = Pending
  - Height: `60 + N*35`
- `st.download_button("⬇ Download Order Status", ...)` → `order_status_{orderId}.xlsx`

---

#### Editable Order Draft Grid (Location Needs + System Code Needs only)

Shown when `st.session_state["_ord_draft"]` is populated:

`sec-hdr`: "✏️ Review & Edit Order List"
`st.caption`: edit instructions

`st.data_editor(...)` — key: `"ord_draft_editor"`:
- `num_rows="dynamic"` (rows can be added/deleted)
- Columns: Material Code, Material Name, UOM, Current Stock (disabled), Shortfall Qty (disabled), Order Qty (editable, numeric), Notes (editable), ☑ Remove (checkbox), Source (disabled)
- Height: `min(500, 60 + N*35)`

`st.button("📤 Submit Order", key="submit_ord_btn", type="primary")`:
- Reads final state (original rows + edits + added rows), excludes "☑ Remove" rows and rows with Order Qty ≤ 0
- Generates order ID: `ORD-{YYYYMMDD}-{NNN}`
- Inserts all rows into `orders_log` (status = "Pending")
- Shows `st.success(f"✅ Order {oid} submitted — N material(s).")`
- Shows `st.download_button(f"⬇ Download Order {oid}", ...)` → `order_{oid}.xlsx`
- Clears `_ord_draft` from session state

---

## Tab 6 (tab5) · Total Overview

**Tab label:** `📈  Total Overview`
**Purpose:** A single master filterable table covering every (Equipment × System Code) pair in the project, with SQM progress, demand/allocation, and material drill-downs.

`sec-hdr`: "📈 Total Overview — Master Equipment & Material Table"

---

### Filter Controls — `st.columns(4)`

| Widget | Key | Options | Default |
|---|---|---|---|
| `st.multiselect(" Location")` | `ov_loc` | `LOCATION_ORDER` | All |
| `st.multiselect(" Type")` | `ov_type` | All unique types (sorted, dropna) | All |
| `st.multiselect(" System Code")` | `ov_code` | All unique system codes (sorted as int) | All |
| `st.selectbox(" Status")` | `ov_status` | `["All", "Fully Ready (100%)", "Partial (50-99%)", "Blocked (<50%)"]` | "All" |

---

### KPI Strip — 6 Clickable Metric Popover Cards

Layout: `st.columns(6)`

| Card | Label | Value | Popover Drill-down |
|---|---|---|---|
| ov1 | No. of Items (filtered) | Count of filtered rows | All filtered rows basic view |
| ov2 | Total SQM | Sum of Total SQM | Sorted desc by Total SQM |
| ov3 | Already Done SQM | Sum of Done SQM | Sorted desc |
| ov4 | Remaining SQM | Sum of Remaining SQM | Sorted desc |
| ov5 | Shortfall SQM | SQM deficit weighted by Fulfil % | Rows with Fulfil % < 100, sorted desc by Shortfall SQM |
| ov6 | Avg Coverage | Mean Fulfil % across filtered rows | Sorted asc by Fulfil % |

---

### Master Table

`st.dataframe` (styled) — key: `"ov_master_tbl"`:
- 20 columns: S.No, Equipment No, Name, Substrate, Type, Location, Lining System+, System Code, System Name, Lining Type, Material Spec., Design, Total SQM, Lining Area SQM, Already Done SQM, Remaining SQM, Total Demand, Allocated, Shortfall Qty, Fulfil %
- Row coloring by Fulfil %, Fulfil % cell text also colored
- Numeric format: SQM cols `{:,.2f}`, Lining Area `{:,.3f}`, Demand/Allocated/Shortfall `{:,.3f}`, Fulfil % `{:.1f}%`
- Height: `min(700, 50 + N*35)`
- Sorted: as returned from master build (file order)

---

### Sub-metrics Strip — `st.columns(4)` (plain `st.metric`)

Appears below the master table:
- `_sm1.metric("Total SQM", ...)` — sum of Total SQM
- `_sm2.metric("Already Done SQM", ...)` — sum of Done SQM
- `_sm3.metric("Pending SQM", ...)` — sum of Remaining SQM
- `_sm4.metric("Completion", ...)` — Done / Total × 100%

---

### Material Detail by System Code

`sec-hdr`: "🔬 Material Detail by System Code"

For each unique System Code in filtered rows (sorted as int):

**`st.expander(f"{dot}  Code {code}  ·  {sname}  ·  {can}/{sqm} SQM  ·  Done: {done}  ·  {avg:.1f}%", expanded=False)`**

Inside:
- **5 clickable metric cards (`st.columns(5)`):**
  - System Code (drill-down: material breakdown)
  - Short Name (same)
  - Total SQM (same)
  - Already Done SQM (same)
  - Coverage SQM: `f"{sc_can:,.2f}  ({sc_cov_avg:.1f}%)"` (same)
- `st.dataframe` — material-level breakdown:
  - Columns: Code, Material Name, UOM, Available, Total Demand, Shortfall, Coverage %
  - Coloring by Coverage %
  - Key: `f"ov_det_{code}"`

---

### Download Buttons — `st.columns(2)`

- Left: `st.download_button("⬇ Download Filtered Master Table", ...)` → `total_overview_{date}.xlsx` (color: "overview")
- Right (conditional — DB only): `st.download_button("⬇ Download Full Consumption Log", ...)` → `consumption_log_full_{date}.xlsx`

---

## Tab 7 (tab_master) · Master Data

**Tab label:** `🗄️  Master Data`
**Purpose:** ERP-style CRUD interface for the three core database tables: Equipment, Recipe, and Inventory.

`sec-hdr`: "🗄️ Master Data — View, Add & Delete Records"

**DB guard:** `st.error("Database not found. Run python setup_db.py first.")` + `st.stop()` if DB absent.

---

### Table Selector

```python
st.radio("Select Table to Manage",
    options=["Equipment", "LINING SYSTEM MATERIAL CONSM", "Materials_DetailsAvailable_Qty"],
    key="md_table_radio", horizontal=True)
```
Then `st.markdown("<hr>")`.

Maps to DB tables: `equipment`, `recipe`, `inventory`.

---

### Add New Row Section

`sec-hdr`: "➕ Add New Row"

#### When "Equipment" is selected — Smart Entry Form

1. `st.multiselect("🔧 Select Lining System Code(s) *", options=code_opts, key="seq_codes_pre")`:
   - `code_opts` = all `"Code N — Short Name"` strings from `recipe` table
   - One DB row will be inserted per selected code

2. **When codes selected:** `st.form(key="smart_eq_form")` appears:

   - `sec-hdr`: "Equipment Identity"
   - `st.columns(2)`:
     - `st.text_input("🏷️ Equipment Tag No. *", placeholder="e.g. V-1001", key="seq_tag")`
     - `st.selectbox("📍 Location *", options=LOCATION_ORDER, key="seq_loc")`

   - `sec-hdr`: "Equipment Details (shared across all selected codes)"
   - Dynamic columns (3 per row) from DB PRAGMA (skipping: id, auto-filled fields, PK fields):
     - Numeric keywords (sqm/qty/for_1) → `st.number_input`
     - Others → `st.text_input`
     - Keys: `f"seq_sh_{col_name}"`

   - `sec-hdr`: "Per Lining System Code"
   - For each selected code:
     - `code-badge` chip + auto-filled short name
     - `st.columns(3)`: Short Name (disabled), Lining Type (disabled), Material Spec. (disabled)
     - `st.columns(2)`: Lining System+ (disabled), **`st.number_input(f"Surface Area SQM * (Code {code})", key=f"seq_sqm_{code}")`** — editable, required
   
   - `st.form_submit_button("💾 Save Equipment")`

3. **Outside form:** `st.button("🧹 Clear Equipment Form", key="seq_clr_btn")`

4. **When no codes selected:** `st.info("Select one or more Lining System Codes above to build the entry form.")`

**On submit:**
- Validates: tag required, codes required, all shared fields filled, all SQMs > 0
- For each code: INSERT into `equipment` + UPSERT into `sqm_progress` (original_sqm = entered SQM)

#### When "LINING SYSTEM MATERIAL CONSM" or "Materials_DetailsAvailable_Qty" is selected — Dynamic Form

- Reads schema from `PRAGMA table_info({db_table})`
- Skips: id, "sl. #" variants
- Renders columns in groups of 3:
  - Numeric keywords → `st.number_input(value=0.0, step=0.001, format="%.4f")`
  - Others → `st.text_input`
  - Keys: `f"dyn_{db_table}_{col_name}"`
- `st.form_submit_button(f"➕ Add Row to {display_name}")`
- Outside form: `st.button("🧹 Clear Form", key=f"dyn_clr_{db_table}")`

---

### View, Edit & Delete Section

`sec-hdr`: `f"📋 View, Edit & Delete — {selected_table}"`

#### Search Filter — `st.columns([2, 1])`

- `st.text_input("🔍 Search table...", key=f"md_search_{db_table}", placeholder="Type to filter…")` — full-text search
- `st.selectbox("in column", ["All columns"] + all_column_names, key=f"md_col_{db_table}")` — column scope selector

Search is case-insensitive, applied client-side (re-filters on every rerun).

#### Data Editor

`st.data_editor(...)` — key: `f"md_editor_{db_table}"`:
- Columns: ☐ Select (checkbox), Sl. No. (disabled), all DB columns (PK disabled)
- `num_rows="fixed"`, `hide_index=True`
- Height: `min(600, 50 + N*35)`
- `st.caption(f"Total entries: {count}")` below

#### Action Buttons — `st.columns([2, 2, 3])`

- Left: `st.button("💾 Save Cell Edits", type="primary", key=f"save_edits_{db_table}")`:
  - Reads `edited_rows` from editor state
  - Skips Sl. No., PK, and ☐ Select columns
  - Runs `UPDATE {table} SET ... WHERE "{pk}" = ?` for each changed row
  - `st.success(f"✅ {N} row(s) updated.")`

- Middle: `st.button("🗑️ Delete Selected Rows", type="secondary", key=f"del_checked_{db_table}")`:
  - Reads `edited_rows` for rows with `☐ Select == True`
  - For Equipment: also deletes matching `sqm_progress` row (cascade)
  - `st.success(f"✅ {N} row(s) deleted.")`
  - `st.warning` if no rows checked

`st.caption("⚠️ Deletion is permanent. Equipment rows also remove the matching sqm_progress record.")`

#### Download

`st.download_button(f"⬇ Download {table_name} Table", ...)` → `{db_table}_export_{date}.xlsx` (color scheme per table: brown_field / train_j / train_k)

---

## Data & Backend Flow

### Data Sources

```
SQLite (sme_database.db)  ─────►  load_all()  ──────────────────────────────────────►
                                  @st.cache_data                                     │
Excel Fallback (3 files)  ─────►  validate_data.py cleaners                         │
                                                                                     ▼
                            inv (DataFrame)      — material inventory + ordered qty
                            recipe (DataFrame)   — per-sqm material requirements
                            equip_sc (DataFrame) — equipment × system code + SQM progress
                            dm (DataFrame)       — demand matrix: Demand_Qty = For_1_SQM × remaining_SQM
                            eq_master (DataFrame)— one row per equipment tag (aggregated)
                            sqm_ref (DataFrame)  — SQM totals including done/remaining
```

### Cascading Allocation Engine

```
cascade_allocate(tag_order: list[str])
    ↓
_cached_cascade_allocate(tag_order_tuple)   ← @st.cache_data (keyed on exact tuple)
    ↓
For each tag in order:
  For each system code (numeric sort):
    For each material:
      alloc = min(demand, pool[material])
      pool[material] -= alloc
      shortfall = demand - alloc
    ↓
Returns alloc_df with: Equipment_Tag_No., Lining_System_Code, Material_Code,
                       Demand_Qty, Allocated_Qty, Shortfall_Qty, Fulfillment_Pct,
                       Pool_Before, Pool_After, Total_SQM
```

### DB Write Protocol

Every time a write operation (INSERT/UPDATE/DELETE) succeeds:
1. `conn.commit()`
2. `st.cache_data.clear()` — forces `load_all()` to re-fetch from DB on next run
3. `st.rerun()` — triggers full Streamlit re-render with fresh data

### Session-to-Tab Data Flow

```
session_tags (list in session_state)
    │
    ├── Tab 1: Equipment Entry ── adds/removes/reorders tags
    │
    ├── Tab 2: Session Report ── reads tags → cascade_allocate() → display
    │
    ├── Tab 3: Location Report ── loc_order[loc] (independent per-location) 
    │          all_eq_order (independent global order)
    │
    └── Tab 4: Execution Plan ── reads session_tags → cascade_allocate()
                                 → selectbox to pick tag and system code
```

---

## Reusable Component Library

### `dbl_click_metric(label, value, state_key, drilldown_title, drilldown_df, ...)`

A styled `st.popover` that renders as a metric card. Clicking opens a `st.dataframe` drill-down.

**Visual structure:** Amber left-border, background `var(--bg2)`, styled like `[data-testid="stMetric"]`.
**Popover content:** `st.subheader`, optional `st.caption`, `st.dataframe` (max 200 rows, height capped at 420px).

Used in: Dashboard (7 cards), Session Report (5), Entry Tab (per-SC), Inventory Tab (4), Total Overview (6), Execution Plan (per-SC).

---

### `plotly_mat_table(df, key_suffix, height, show_sqm, tag, code, allocated_label)`

A color-coded `st.dataframe` rendered with pandas `.style`.

**Base columns:** Code, Material Name, UOM, Demand, [On Order if available], Allocated/Available, Shortfall, Fulfil %
**SQM columns (if `show_sqm=True`):** SQM Total, SQM Done, SQM Deficit, SQM Done %

Row background color by Fulfil %:
- ≥100%: green `rgba(16,185,129, 0.12)`
- ≥90%: orange `rgba(249,115,22, 0.12)`
- ≥80%: yellow `rgba(234,179,8, 0.12)`
- <80%: red `rgba(239,68,68, 0.12)`

Fulfil % cell gets matching text color + `font-weight: 700`.

---

### `render_suggestion_panel(tag_list, panel_key)`

Runs `_run_suggestion_engine()` (with spinner), then renders two-column suggestion cards.

**Algorithm:** For each non-first tag, tries every earlier position and picks the one giving the best % gain. Threshold: >0.4%.

**Outputs:** Up to 5 "By Equipment" cards + up to 8 "By System Code" cards.

**Visual:** `.card` HTML div, amber `+N%` gain badge, green if would reach 100%.

---

### `generate_excel_report(df, title, add_grand_total, color_scheme)`

Generates a professionally formatted `.xlsx` file with:
- Rows 0–3: Logo image area (121×83 px, 96 DPI)
- Row 4: Title bar (merged, dark bg, white bold 13pt)
- Row 5: Column headers (navy bg, white bold 10pt, borders, AutoFilter)
- Row 6+: Data rows (9pt, borders)
- Last row: GRAND TOTAL (gold/colored bg, bold) if `add_grand_total=True`
- Top-right metadata: "Report Generated: YYYY-MM-DD HH:MM" + "Generated By: Smart Material Estimator"
- Auto-column widths (max 42 chars)

Color schemes:
| Key | Title Bg | Header Bg | Total Bg |
|---|---|---|---|
| dashboard | `#1A2A3A` | `#2D4A6A` | `#F0C040` |
| brown_field | `#0F2D52` | `#1E5799` | `#BDD7F0` |
| train_j | `#4A2E00` | `#A0620A` | `#FDE8A0` |
| train_k | `#0A2E1A` | `#1A6B48` | `#B3F0D8` |
| session | `#2D1A52` | `#5B2D8E` | `#E5D0F0` |
| execution | `#3A0A0A` | `#8E1A1A` | `#F5C6C6` |
| overview | `#0A2A2A` | `#0E7490` | `#A5F3FC` |

---

### Location Color Badges

Three classes defined in CSS, applied via `loc_badge(loc)` helper:

| Location | Class | Color |
|---|---|---|
| Brown Field | `.loc-bf` | Blue `#3B82F6` |
| TRAIN J | `.loc-tj` | Amber `#F59E0B` |
| TRAIN K | `.loc-tk` | Green `#10B981` |

---

### Fulfillment Pill

`fulfil_pill(pct)` — inline `<span class="pill pill-{color}">`:
- ≥100%: `.pill-g` (green)
- ≥90%: `.pill-o` (orange)
- ≥80%: `.pill-y` (yellow)
- <80%: `.pill-r` (red)

---

## Design Token Reference

All tokens are injected as CSS variables via `st.markdown(<style>)` to support both light and dark mode.

### Text Colors

| Token | Usage |
|---|---|
| `var(--t0)` | Primary text (full opacity) |
| `var(--t1)` | Standard body text |
| `var(--t2)` | 82% opacity — secondary |
| `var(--t3)` | 62% opacity — tertiary |
| `var(--t4)` | 45% opacity — metadata labels |
| `var(--t5)` | 30% opacity — very muted |

### Background Colors

| Token | Usage |
|---|---|
| `var(--bg0)` | Main page background |
| `var(--bg1)` | Sidebar background |
| `var(--bg2)` | Cards, metric tiles |
| `var(--bg3)` | Nested cards (syscode-block) |
| `var(--bg4)` | Deep nested backgrounds |

### Accent Colors

| Token | Hex | Usage |
|---|---|---|
| `var(--amber)` | `#F59E0B` | Primary brand accent, active tabs, buttons |
| `var(--amber2)` | `#FCD34D` | Amber hover state |
| `var(--amber3)` | `#D97706` | Amber dark (button gradient start) |
| `var(--amber-bg)` | `rgba(245,158,11, .11)` | Amber tinted backgrounds |
| `var(--green)` | `#10B981` | 100% fulfillment, success |
| `var(--red)` | `#EF4444` | Shortfall, error, blocked |
| `var(--orange)` | `#F97316` | 90–99% coverage (partial) |
| `var(--yellow)` | `#EAB308` | 80–89% coverage (caution) |
| `var(--blue)` | `#3B82F6` | Info, Brown Field accent |

### Border Radii

| Token | Value |
|---|---|
| `var(--r-sm)` | 6px |
| `var(--r-md)` | 10px |
| `var(--r-lg)` | 14px |

### Fulfillment Color Thresholds (universal across the app)

| Range | Color | Status |
|---|---|---|
| ≥ 100% | Green `#10B981` | Fully Ready / Complete |
| 90–99% | Orange `#F97316` | Near-complete / Warning |
| 80–89% | Yellow `#EAB308` | Partial / Caution |
| < 80% | Red `#EF4444` | Blocked / Critical |

---

*End of specification. Every tab, widget, layout element, state variable, and data flow in the Smart Material Estimator & Planner v3 application has been documented above.*
