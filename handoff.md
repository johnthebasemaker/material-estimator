# Smart Material Estimator — Developer Handoff

> Single-source-of-truth document for any developer (or new Claude Code chat)
> picking up this project. Read this top-to-bottom before touching code.
> Everything you need is here — quick start, architecture, schema, file map,
> conventions, and known issues.

---

## 1. What this project is

A **Streamlit desktop web app** for CNCEC's *Rubber Lining (RL)* and *Brick
Lining (BL)* material planning. It answers four core questions for the
site engineer:

1. **What can I build today** with the materials I currently have?
2. **What materials must I procure** to finish what I cannot build?
3. **How much have I actually consumed vs. expected** per equipment / system
   code?
4. **How many more days can I keep producing** at today's run-rate before the
   first material runs out?

It is **single-user / single-tenant**, runs locally
(`streamlit run app.py`), and persists everything in a local SQLite database
(`sme_database.db`).

---

## 2. Quick start

```bash
# 1. Place these files in one folder:
#    app.py · validate_data.py · allocation_engine.py · setup_db.py
#    requirements.txt · sme_database.db (optional — auto-created on first run)
#    Equipment.xlsx · For_1_SQM.xlsx · Materials_DetailsAvailable_Qty.xlsx
#    logo.png

# 2. Virtual env
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

# 3. Install
pip install -r requirements.txt

# 4. (Optional first-time) seed the DB from the three Excel files
python setup_db.py

# 5. Run
streamlit run app.py
# default browser opens http://localhost:8501
# If port 8501 is busy:
streamlit run app.py --server.port 8502
```

**Login credentials** — set as constants in `app.py`:

| Field      | Value      | Defined at         |
|------------|------------|--------------------|
| Username   | `admin`    | `_ADMIN_USER` (line 1252) |
| Password   | `admin2026`| `_ADMIN_PASS` (line 1253) |

There is **no** "download password" input on the login screen. Excel and PDF
downloads use two hardcoded passwords defined at the top of `app.py`:

| Constant         | Value       | Line |
|------------------|-------------|------|
| `_XLSX_PASSWORD` | `excel2026` | 39   |
| `_PDF_PASSWORD`  | `pdf2026`   | 40   |

Change those constants to rotate download protection. The user types the same
value when prompted to download.

---

## 3. File / module map

```
CNCEC RL and BL Material Prediction Project/
├── app.py                       8,500+ LOC — the entire Streamlit UI
├── allocation_engine.py         Pure algorithms: demand matrix, cascade
│                                allocation, feasibility, suggestion engine,
│                                procurement list builder (~460 LOC)
├── validate_data.py             Excel → cleaned DataFrames + join validation
├── setup_db.py                  One-shot loader: cleaned Excel → SQLite tables
├── requirements.txt             Pinned minimums (see §4)
├── sme_database.db              SQLite DB (auto-created on first run)
├── Equipment.xlsx               File C — equipment master with surface areas
├── For_1_SQM.xlsx               File B — lining-system recipes (qty per m²)
├── Materials_DetailsAvailable_Qty.xlsx
│                                File A — inventory snapshot + PO data
├── logo.png                     200px logo for header + login card
└── handoff.md                   ← this file
```

`app.py` is intentionally a monolith — every UI tab, helper, and data
transform lives here so the surface a developer needs to scan is **one file**.
Keep it that way unless a section is genuinely reusable.

---

## 4. Dependencies (`requirements.txt`)

```
streamlit>=1.35.0
pandas>=2.0.0
openpyxl>=3.1.0
xlsxwriter>=3.1.0
plotly>=5.18.0
streamlit-sortables>=0.3.0
reportlab>=4.0.0      # PDF report generation + AES encryption
pyzipper>=0.3.6       # AES-ZIP wrapper for Excel downloads
```

Both `reportlab` and `pyzipper` are imported behind `try/except` with
`_HAS_REPORTLAB` / `_HAS_PYZIPPER` flags so the app still runs (without
encryption / without PDF) if they are missing.

---

## 5. Data flow at a glance

```
Excel files ──► validate_data.clean_*()  ──► setup_db.dynamic_sync_table()
                                                       │
                                                       ▼
                                              sme_database.db
                                                       │
                                            app.load_all() [st.cache_data]
                                                       │
                                                       ▼
                ┌──────────────────────────────────────┴──────────────────────┐
                │                                                             │
                ▼                                                             ▼
   allocation_engine.cascade_allocate(tag_order)              Direct SQL reads
   (Tab 1 priority list drives this)                          (Inventory, Consumption,
                │                                              Orders, Receipts)
                ▼
   Per-tag fulfillment %, per-system-code SQM, shortfalls
                │
                ▼
    Renders into Tabs 0 / 2 / 3 / 4 / 5 / 6 / 7 / 8
```

**Cache invalidation rule:** *any* DB mutation (insert/update/delete) must
be followed by `st.cache_data.clear()` so `load_all()` re-reads. Search for
this pattern when editing.

---

## 6. SQLite schema (`sme_database.db`)

| Table | Purpose |
|---|---|
| `inventory` | Material master + `available_qty` + `ordered_qty`. PK = `material_code`. |
| `recipe` | Lining-system recipes from *For_1_SQM*. One row per (system_code × material_code) with `for_1_sqm`. |
| `equipment` | Equipment master from *Equipment.xlsx*. Includes `location`, `type`, `substrate`, `equipment_tag`, `surface_area_sqm`. Has duplicate original-spelling columns (`Ht. /W`, `Dia / L`, etc.) added by `dynamic_sync_table` — kept in DB for autofill but hidden in the UI. |
| `sqm_progress` | Running tally per (equipment_tag, system_code): `original_sqm`, `done_sqm`. Driven by consumption_log SUMs but cached here for fast Dashboard reads. |
| `consumption_log` | Append-only daily entries. (date × equipment_tag × system_code × material_code) with `sqm_completed`, `expected_qty`, `consumed_qty`, variance. |
| `draft_consumption` | Pre-submit staging area for the multi-row consumption form. Cleared on submit. |
| `receipt_log` | Goods receipt entries (incoming stock). `received_qty` increases `inventory.available_qty`. |
| `orders_log` | Procurement orders raised through the app — `order_id`, `ordered_qty`, `fulfilled_qty`, `status`. Linked back via `receipt_log.order_id`. |
| `locations` | **Dynamic** location list (seeded with Brown Field / TRAIN J / TRAIN K). Editable via *Master Data → ➕ Add Location*. Drives every Location dropdown in the app. |
| `types` | **Dynamic** equipment-type list (seeded with Vessel / Tank / Column / Pipe / Reactor + any types found on existing equipment rows). Editable via *Master Data → ➕ Add Type*. Drives every Type dropdown in the app. |

The schema is *evolving* — `setup_db.dynamic_sync_table()` will `ALTER
TABLE … ADD COLUMN` for any new column it sees in the source Excel. Don't
drop unknown columns; many are kept for downstream reporting and autofill
(e.g. the equipment table still carries `Lining_System`, `Material Spec.`,
`Lining_Area/location` even though their snake_case counterparts are the
canonical fields).

---

## 7. UI structure (tabs and sub-views)

The app shows **9 top-level tabs**, declared in one place:

```python
# app.py ~ line 3407
tab0, tab1, tab2, tab3, tab_eqrep, tab4, tab_consume, tab5, tab_master = st.tabs([
    "📊  Dashboard",
    "🔍  Selective Equipment Entry",
    "📦  Session Order Report",
    "📍  Location Report",
    "📋  Equipment Report",
    "⚙️  Execution Plan",
    "📦  Inventory",
    "📈  Total Overview",
    "🗄️  Master Data",
])
```

| # | Tab | Sub-views (radio) | Purpose |
|---|---|---|---|
| 0 | Dashboard | Project Overview · Material Requirement & Procurement | KPI strip + Material Balance + Stock-only materials + per-location procurement breakdown. |
| 1 | Selective Equipment Entry | – | Drag-drop priority ordering of equipment tags (via `streamlit-sortables`). Drives `cascade_allocate()`. |
| 2 | Session Order Report | – | The full computed picture for the current priority order — per-equipment fulfillment, materials, downloads. |
| 3 | Location Report | Location Based · All Equipment | Same data sliced by location. The drag-priority sortable is rendered in a narrow 2/5-width column. **No charts** — per-equipment material tables only. Per-location Excel + PDF downloads. |
| 4 | Equipment Report | – | Per-equipment deep-dive. **Column order:** Equipment Tag No. → Equipment Name → System Code → System Name → Total SQM. **Excel report sections (in order):** 1) Summary by Equipment, 2) Summary by System Code, 3) Detailed Table. |
| 5 | Execution Plan | ⚙️ Execution Plan · 📋 Progress List · 📊 Consumption Comparison | Day-by-day execution; numbered production-detail blocks; expected vs actual variance. |
| 6 | Inventory | Inventory Dashboard · Consumption · Order Status · New Order · Receipt Log · Consumption Log | Operational data entry — daily consumption, goods receipts, raising orders. Consumption submit triggers the **Days-of-Continuation** report inline (see §13). |
| 7 | Total Overview | – | Roll-up across all locations and codes. Independent of priority order. |
| 8 | Master Data | Equipment · LINING SYSTEM MATERIAL CONSM · Materials_DetailsAvailable_Qty · ➕ Add Location · ➕ Add Type | CRUD on the three source tables + location + type management. The Equipment view hides the original-spelling duplicate columns (`Ht. /W`, `Dia / L`, `Equipment Total SQM`, `Remaraks`, `Material Spec.`, `Lining_Area/location`, `Lining_System`, `Sl. #`, `Project`, `WBS #`, `IO#`, `Drawing #`). |

Each sub-view is rendered by an `if exec_subview == "..."` / `if
md_table_sel == "..."` branch inside the relevant `with tabX:` block.

---

## 8. Sticky / frozen header

Three elements are pinned at the top while the user scrolls:

| Element             | CSS technique                | Top offset |
|---------------------|------------------------------|-----------:|
| Title bar (`.sticky-header-wrap` — logo + project title + v3 chip) | `position: fixed; top: 0;` | 0 |
| Tab strip (`[data-testid="stTabs"] > div:first-of-type`) | `position: sticky;` | 76px |
| Active sub-view radio (the first horizontal stRadio in a tab panel, matched via `:has()`) | `position: sticky;` | 122px |

`overflow: visible !important` is applied to the intermediate ancestors
(`stAppViewBlockContainer`, `stMainBlockContainer`, the tabs containers, and
tab panels) so sticky positioning isn't broken by an `overflow:hidden`
ancestor. **Do not** add `overflow: visible` to `[data-testid="stAppViewContainer"]`
or `section.main` — those own the page scrollbar and forcing them to visible
freezes the whole page.

---

## 9. Critical algorithms (`allocation_engine.py`)

| Function | What it does |
|---|---|
| `build_demand_matrix(eq, recipe, sqm_progress)` | Joins equipment × recipe × remaining SQM → one row per (equipment_tag × material_code) with `Demand_Qty`. |
| `allocate_sequential(demand_df, inv_df, tag_order)` | Walks `tag_order` top-to-bottom. For each tag, allocates available stock to each material until depleted. Output: `Allocated_Qty`, `Shortfall`, `Coverage_Pct` per (tag × material). |
| `compute_feasibility(alloc_df)` | Aggregates per equipment_tag → `Equipment_Fulfillment_Pct` + status label (✅ Fully Ready / 🟡 Partial / 🔴 Blocked). |
| `run_suggestion_engine(...)` | Tries pausing each tag one-by-one to see if downstream tags become fully buildable. Returns the **best pause scenario**. |
| `build_procurement_list(alloc_df, inv_df)` | Output: per-material `Shortfall` (after `Available_Qty`), `Net_Shortfall` (after `Ordered_Qty`), used by Tab 0 → Procurement view. |

Inside `app.py` the same engine is wrapped in:

```python
@st.cache_data(show_spinner=False)
def _cached_cascade_allocate(tag_order_tuple: tuple) -> pd.DataFrame: ...

def cascade_allocate(tag_order: list[str]) -> pd.DataFrame:
    return _cached_cascade_allocate(tuple(tag_order))
```

— so flipping the priority order is instant after the first compute.

---

## 10. Cross-cutting helpers in `app.py`

| Helper | Line | What |
|---|---:|---|
| `_xlsx_password()` / `_pdf_password()` | 43 / 47 | Return the hardcoded download passwords. |
| `_current_username()` | 51 | Username captured at login (`st.session_state["_login_username"]`), used in download filenames. |
| `_safe_for_filename(text)` | 56 | Strip path-unfriendly characters. |
| `_standard_filename(stem, ext)` | 66 | Apply project filename convention: `<ReportName>_<username>_<YYYY-MM-DD>.<ext>` (date-only — no time). |
| `_rename_to_standard(file_name, …)` | 77 | Rewrite any pre-existing filename to the standard convention. |
| `_encrypt_xlsx_bytes(raw, pwd, …)` | 91 | Wraps any `.xlsx` bytes in an AES-encrypted ZIP via `pyzipper`. |
| `_pdf_from_sheets([{name, df, title}, ...], pwd)` | 124 | Builds a multi-sheet landscape A4 PDF via ReportLab with `encrypt=pwd`. |
| `_pdf_from_df(df, title, pwd)` | 187 | Thin wrapper for single-sheet PDFs. |
| `_trigger_browser_download(bytes, name, mime)` | 198 | Injects a hidden `<a download>` into the parent document via a `components.html` iframe and `click()`s it. Used to auto-fire downloads after the password popover accepts the input — see §11. |
| `_orig_download_button` | 231 | Saved reference to the unpatched `st.download_button` (used internally and by `_pdf_download_button`). |
| `_secure_download_button(...)` | 232 | **Monkey-patches `st.download_button`.** For `.xlsx`/`.xlsm` it: AES-encrypts the bytes, renames to standard, then wraps in a `st.popover` with a password input. When the entered value matches `_XLSX_PASSWORD` the browser download fires automatically (no second click). For all other extensions it just renames and passes through. |
| `_pdf_download_button(label, …)` | 317 | PDF companion. Same popover + password flow but the password must match `_PDF_PASSWORD`. PDF bytes are generated lazily — only after the password matches — so changing the password input doesn't re-render the PDF on every keystroke. |
| `_show_login()` | 1255 | Plain Streamlit columns + form (no custom card div — that wrapper was previously closed prematurely by Streamlit's element wrappers, which hid the form). Username + password only; the legacy "Download Password" field has been removed. Enter submits. |
| `_DEFAULT_LOCATIONS` | 1324 | Seed list for the `locations` table. |
| `_ensure_locations_table()` / `_refresh_location_order()` | 1331 / 1353 | Boot-time: create + seed `locations`, refresh the mutable module-level `LOCATION_ORDER` list. |
| `_DEFAULT_TYPES` | 1375 | Seed list for the `types` table (`Vessel`, `Tank`, `Column`, `Pipe`, `Reactor`). |
| `_ensure_types_table()` / `_refresh_type_order()` | 1385 / 1425 | Boot-time: create + seed `types`, refresh `TYPE_ORDER`. Backfills the table with any distinct `equipment.type` values not in the defaults. |
| `_get_all_types(eq_master_df)` | 1442 | Authoritative Type list: union of `TYPE_ORDER` (registered) + any types actually present on equipment. Used by every Type dropdown. |
| `load_all()` | 1476 | `@st.cache_data` — single source of truth for in-memory DataFrames. Returns inv, recipe, equipment, sqm_progress, etc. |
| `cascade_allocate` / `tag_fulfillment` / `syscode_fulfillment` / `sqm_can_do` | 1671 / 1679 / 1686 / 1693 | Thin wrappers around the allocation engine. |
| `dbl_click_metric(label, value, key, drill_title, drill_df)` | 1714 | KPI tile with click-to-expand drill-down DataFrame. Plain `st.metric` is used where no drill-down should appear (e.g. inside Session Order Report and Total Overview). |
| `generate_excel_report(df, title, color_scheme)` | 1758 | xlsxwriter formatting with color schemes. |
| `generate_multi_sheet_excel(sheets)` | 1889 | Multi-sheet Excel — one DataFrame per sheet, each with its own header strip. |
| `_equipment_report_excel(...)` | 1996 | Per-equipment / per-location Excel report. **Sections written in this order:** title → 1) Summary by Equipment (cols: Equipment Tag No., Equipment Name, System Name, Total SQM) → 2) Summary by System Code → 3) Detailed Table. The legacy `Equipment No.` column name is auto-renamed to `Equipment Tag No.` for backwards-compat. |
| `_location_report_excel(*, sheets)` | 2372 | Big formatted multi-block per-location report (alloc matrix + 3 summary blocks). |
| `_run_suggestion_engine(...)` / `render_suggestion_panel(...)` | 2826 / 2900 | The "best pause" predictive panel. |
| `loc_badge(loc)` / `status_dot(pct)` / `fulfil_pill(pct)` | 1618 / 1622 / 1628 | Small HTML pill renderers. |
| `plotly_mat_table(...)` | 2992 | Plotly table renderer for per-(tag, code) material breakdowns. |
| `_apply_theme_attr()` | 3086 | Injects `data-theme="dark|light"` so dark/light toggle works without re-rendering Plotly. |

---

## 11. Download / encryption model

Every report download must come in **two side-by-side buttons**:

```python
c_xlsx, c_pdf = st.columns(2)
with c_xlsx:
    st.download_button(
        "⬇ Excel — <Title>",
        data=generate_excel_report(df, "<Title>"),
        file_name="<stem>.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with c_pdf:
    _pdf_download_button(
        "⬇ PDF — <Title>",
        df=df,                       # or sheets=[{name, df, title}, ...]
        title="<Title>",
        file_stem="<stem>",
        key="pdf_<unique>",
        use_container_width=True,
    )
```

What happens behind the scenes:

1. **Both buttons render as `st.popover`s** (Excel and PDF). Clicking opens a
   small panel with a password input.
2. The user types the password (`excel2026` for Excel, `pdf2026` for PDF).
3. **As soon as the entered value matches** the corresponding constant, the
   file is auto-downloaded via `_trigger_browser_download()` — no second
   click required.
4. A per-button session-state flag (`_xlsx_dl_fired__<key>` /
   `_pdf_dl_fired__<key>`) prevents the download from firing on every rerun
   while the password remains in the field. The popover then shows
   `✓ Download started` plus a `↻ Download again` button to reset and re-fire.
5. The Excel file is also AES-encrypted at rest (the bytes are wrapped in a
   `.protected.zip` via `pyzipper`). The PDF is also password-protected via
   ReportLab's `encrypt=` option. Both use the same constants. So even after
   the UI gate, opening the downloaded file in Excel / a PDF reader prompts
   for the same password — that is intentional defense-in-depth.

**Filename convention** — every download is rewritten to:

```
<ReportName>_<username>_<YYYY-MM-DD>.<ext>
```

For Excel that becomes `<ReportName>_<username>_<YYYY-MM-DD>.protected.zip`
(the encrypted wrapper). The inner archive contains `report.xlsx`.

When adding a new download, **always provide both buttons.** Search the
codebase for `_pdf_download_button` for ~30 reference call sites.

---

## 12. Consumption entry — gotchas

The Consumption sub-tab under **Inventory** is the most behaviorally
complex area. Key rules already wired:

1. **Multi-code selection** uses a multiselect with a sentinel option
   `_SELECT_ALL = "✨ Select All"`. When picked, all available codes are
   used.
2. **Separate SQM input per code** — `f"form_ce_sqm__{tag}__{code}"`.
3. **SQM cap** = `min(remaining_sqm, stock_coverage_sqm)` where
   `stock_coverage_sqm = min(Available / For_1_SQM)` across all recipe
   materials. The `st.number_input` enforces `max_value=`.
4. **Required Qty** column is **live** — recomputed as `for_1 × sqm_today`
   on every rerun. No "For 1 SQM" column is shown to the user anymore.
5. **Actual Consumed** is capped at `Available_Qty` via `max_value=`.
6. **Available = 0** rows render in **red bold**.
7. **Live shortfall warnings** show below the materials grid as soon as
   the user types an SQM that exceeds any material's available stock.
8. **Widget-key uniqueness** is non-negotiable. The key for each material
   row is `f"form_mat__{tag}__{code}__{row_idx}__{material_code}"` — the
   `row_idx` is required because a code can have duplicated material
   codes in its recipe; without it Streamlit raises
   `StreamlitDuplicateElementKey`.

---

## 13. Days-of-Continuation report (post-submit)

After **Submit Consumption** succeeds, a `📆 Days of Continuation` block
renders inline below the success message (see `app.py` ~ line 6549).

Logic:

```
daily_consumption_per_material = sum(effective_qty in this submission)
days_remaining                 = available_qty / daily_consumption_per_material
days_remaining_with_po         = (available_qty + ordered_qty) / daily_consumption
bottleneck                     = min(finite days_remaining values)
```

The block:

- Sorts most-critical material first.
- Colour-codes rows: red < 3 days, amber 3–7 days, green ≥ 7 days, infinity
  for materials with no daily burn.
- Headlines the bottleneck (`"available stock will sustain this output for
  X day(s) before the first material runs out"`).
- Renders an Excel + PDF download pair (same password-popover flow as every
  other report).

Important: `st.rerun()` is intentionally **not** called after submit, so the
success message, the consumption report downloads, and the
Days-of-Continuation block all stay visible together. `st.cache_data.clear()`
is called so the next user action sees the post-deduction inventory.

---

## 14. Theming & styling

- Theme palette is CSS-variable based, declared in one big `<style>` block
  near the top of `app.py` (~line 530 onwards).
- `_apply_theme_attr()` injects `data-theme="dark|light"` so dark/light
  toggling works without re-rendering Plotly charts.
- The sidebar has a "theme toggle pill" styled as `.sme-theme-toggle`.
- Color tokens used everywhere: `var(--t1..t5)` for text, `var(--bg1..bg3)`
  for backgrounds, `var(--amber)` for the brand accent.
- **KPI cards (`stMetric` / `stPopover`) have no hover float or transition** —
  click-only popover drilldowns. Belt-and-braces `transition / animation /
  transform: none !important` is applied across every state and on the
  `:first-child` / `:last-child` columns explicitly.
- **Input dimensions** are unified via CSS: `min-height: 38px` on every
  `stTextInput`/`stNumberInput`/`stDateInput`/`stSelectbox` (no fixed height
  — that previously collapsed padded children like the login form).
  Multiselect uses `min-height` only so chips can wrap.

When adding UI: prefer existing tokens over hex literals.

---

## 15. Login flow (current implementation)

```
_show_login() renders centered inside a [1, 1.6, 1] column.
  ↓
st.form("_login_form")        ← Enter key submits the form
  ├ Username
  ├ Password
  └ st.form_submit_button("🔐  Login", type="primary")
        ↓ on success
   session_state["_authenticated"]  = True
   session_state["_login_username"] = <username>   (used in download filenames)
   st.rerun()
```

**Admin credentials are hardcoded** at the top of `_show_login()` (line 1252):

```python
_ADMIN_USER = "admin"
_ADMIN_PASS = "admin2026"
```

This is intentional for a desktop-only deployment. If multi-user is ever
needed, replace with a `users` table + bcrypt — but **keep the download
passwords as separate hardcoded constants** at the top of `app.py`. They are
intentionally independent of any user account.

---

## 16. Dropdowns

- A small JS component injected via `streamlit.components.v1.html` (~line
  3344) listens for clicks on `li[role="option"]` and dispatches `Escape`
  to close the menu after each pick — fixes Streamlit's default behavior
  of leaving multiselect menus open.
- All Type dropdowns (Dashboard filter, Session-Order filter, Consumption
  form, Total Overview filter, Equipment add form) call `_get_all_types(...)`
  so a newly registered Type from Master Data appears immediately, even
  before any equipment uses it.
- All Location dropdowns reference the mutable module-level `LOCATION_ORDER`
  list — populated from the `locations` table at boot, refreshed in place
  by `_refresh_location_order()`.

---

## 17. Adding a new location / type end-to-end

**Add Location:** Master Data → **➕ Add Location** → fill the form.

What happens automatically afterward:

1. `_refresh_location_order()` reloads the in-memory `LOCATION_ORDER`
   list (it's a mutable module-level list — never reassigned, always
   `.clear()`+`.extend()`'d so every reference sees the update).
2. Every Location dropdown picks it up immediately.
3. To add equipment under the new location, switch the same Master Data
   tab to *Equipment* and use the Add-Row form — `location` will be in
   the dropdown.

Default locations (`Brown Field`, `TRAIN J`, `TRAIN K`) are
**delete-protected**. Custom locations can be removed only if no
equipment row references them.

**Add Type:** Master Data → **➕ Add Type** → fill the form. Mirror behavior:
`_refresh_type_order()` refreshes `TYPE_ORDER`; every Type dropdown picks it
up. Default types (`Vessel`, `Tank`, `Column`, `Pipe`, `Reactor`) are
delete-protected. Custom types can be removed only if no equipment row
references them.

---

## 18. Conventions to keep

- **One file, one source of truth.** Don't split `app.py` into modules
  just for readability — the search-with-grep workflow depends on it.
- **`st.cache_data.clear()` after every DB write.** This is the
  invalidation contract.
- **All downloads come in pairs (Excel + PDF).** No exceptions. Both pairs
  go through the password-popover gate.
- **Widget keys must be globally unique** and prefixed by their parent
  context (e.g. `form_mat__{tag}__{code}__{idx}__{mc}`). If you reuse
  a key, Streamlit will silently share state across widgets.
- **No bare hex colors.** Use CSS variables.
- **`LOCATION_ORDER` and `TYPE_ORDER` are mutable.** Never do
  `LOCATION_ORDER = [...]` — always `LOCATION_ORDER.clear();
  LOCATION_ORDER.extend(...)`.
- **`dbl_click_metric` vs `st.metric`** — use the latter where a popover
  drill-down would be noise (Session Order Report, Total Overview System
  Code tiles).
- **Master Data → Equipment** display drops a known-duplicate column list
  (see `_EQUIP_DUPLICATE_COLS` in the view section). If you add new
  cleaner-kept columns, add the original-spelling counterpart to that set
  so the grid stays de-duplicated.
- **Sticky positioning** depends on `:has()` CSS (Chromium 105+, Safari
  15.4+, Firefox 121+). If a target browser is older, the sub-view radio
  won't stick — fall back to scrolling.

---

## 19. Troubleshooting cheat sheet

| Symptom | Likely cause | Fix |
|---|---|---|
| `NameError: _dd_equip_df` (or similar `_dd_*_df`) on Dashboard | Variable defined inside one `if dash_view == ...` branch but referenced in the other | Hoist the assignment above the `if/else`. |
| `StreamlitDuplicateElementKey` on Consumption Select-All | Two material rows produced the same widget key | Confirm key includes `row_idx`: `f"form_mat__{tag}__{code}__{row_idx}__{mc}"`. |
| Excel download arrives as `.protected.zip` | This is **expected** — that's the AES-encrypted wrapper. Open with `excel2026`. |
| Inside the .protected.zip the file is called `report.xlsx`, not the report name | Cosmetic — the outer zip filename carries the report name + username + date. Inner name is fixed by `_encrypt_xlsx_bytes`. |
| PDF is empty or unopenable | `reportlab` not installed → `_HAS_REPORTLAB = False`; install with `pip install reportlab>=4.0.0`. |
| Port 8501 already in use | Another Streamlit app is running | `streamlit run app.py --server.port 8502`. |
| Master Data edits don't reflect on Dashboard | Missing `st.cache_data.clear()` after the mutation | Add it immediately after the `commit()`. |
| New location/type doesn't show in dropdowns | `_refresh_location_order()` / `_refresh_type_order()` not called after insert | Call it right after committing the new row. |
| Login button does nothing | You're outside `st.form` — `st.form_submit_button` only works inside one. |
| Whole page un-scrollable | Someone added `overflow: visible !important` to `[data-testid="stAppViewContainer"]` or `section.main`. Those own the scrollbar — leave them alone. |
| Login form invisible / fields collapsed | A blanket `height: 38px !important` rule was added on `[data-baseweb="input"]`. Use `min-height` only. |
| Download fires repeatedly on every keystroke | `_xlsx_dl_fired__<key>` / `_pdf_dl_fired__<key>` flag isn't being set before `_trigger_browser_download`. |

---

## 20. Where to start when picking this up cold

1. Run the app (§2), log in with `admin / admin2026`.
2. Click through every tab once. Note which sub-views exist.
3. Open `app.py` and `Cmd-F` for the section banner of the tab you want
   to change (e.g. `TAB 5 · EXECUTION PLAN`).
4. For algorithm changes, open `allocation_engine.py` — it's pure pandas
   and easy to unit-test in a notebook.
5. For schema changes, edit `setup_db.py` AND add the column in the
   corresponding `validate_data.clean_*` cleaner so a fresh-DB rebuild
   still works.
6. Before committing: run the app, confirm login, dashboard, one report
   download (Excel + PDF — type the passwords), one consumption entry +
   verify the Days-of-Continuation block renders. That's the minimum
   smoke test.

---

## 21. Known issues / open items

> **Note for the next developer / new chat:** when you fix any of these,
> **delete that bullet from this list** so the section stays accurate.
> If this list is empty, leave the section header but write `_None at the
> moment._` under it so future readers know it's been swept.

- **Drag-priority sortable doesn't always feel "compact".** The
  `streamlit-sortables` widget renders inside its own iframe; the CSS that
  trims width (`.sme-compact-sortable + …`) targets the outer wrapper but
  doesn't fully reach inside the iframe DOM. Items still look uniform but
  the panel itself can be slightly wider than 460px on some Streamlit
  versions.
- **Multiselect auto-close JS depends on BaseWeb DOM markers.** The script
  listens for clicks on `li[role="option"]` or `[data-baseweb="menu"] li`.
  If Streamlit ships a new dropdown implementation that changes these
  selectors, the auto-close behavior silently breaks (selectboxes will
  still close on selection — that's native browser behavior).
- **Auto-download via injected `<a download>`** depends on the browser
  honouring data-URI downloads triggered from a same-origin iframe in the
  current user-gesture context. Modern Chrome/Edge/Safari/Firefox all
  allow it. Very old browsers, or aggressive ad-blockers that nuke
  third-party iframe scripts, may block the click. There is no automatic
  fallback — the user would see "Downloading…" but no file would arrive.
- **No automated tests.** The whole codebase relies on manual QA. Any
  refactor should add at least smoke tests around
  `allocation_engine.allocate_sequential` and `cascade_allocate`.
- **Single-user only.** No row-level locking on `consumption_log` /
  `inventory`. Two concurrent users would race.
- **Excel column drift.** `setup_db.dynamic_sync_table` will add new
  columns but never rename or drop them. Stale columns accumulate over
  time. The Master Data Equipment view filters a known duplicate set
  (`_EQUIP_DUPLICATE_COLS`) — extend that set if more duplicates appear.
- **PDF page width.** Wide tables get column-truncated to 12 cols
  (`_df_for_pdf`). For very wide reports prefer multi-sheet Excel.
- **Streamlit hot-reload caveat.** Edits to imports or top-level
  decorators sometimes require a full process restart (`Ctrl-C` →
  `streamlit run app.py`) — not just the in-browser "Rerun".

— End of handoff —
