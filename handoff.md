# Smart Material Estimator — Developer Handoff

> Single-source-of-truth document for any developer (or new Claude Code chat)
> picking up this project. Read this top-to-bottom before touching code.

---

## 1. What this project is

A **Streamlit desktop web app** for CNCEC's *Rubber Lining (RL)* and *Brick
Lining (BL)* material planning. It answers three core questions for the
site engineer:

1. **What can I build today** with the materials I currently have?
2. **What materials must I procure** to finish what I cannot build?
3. **How much have I actually consumed vs. expected** per equipment / system code?

It is **single-user / single-tenant**, runs locally (`streamlit run app.py`),
and persists everything in a local SQLite database (`sme_database.db`).

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

**Login credentials (demo)** — set inside `app.py`:

| Field | Value |
|---|---|
| Username | `admin` |
| Password | `admin2026` |
| Download Password | *any string ≥ 4 chars* — encrypts every download this session |

The Download Password is **session-scoped** (`st.session_state["_dl_pwd"]`).
It encrypts both PDF (ReportLab `encrypt=`) and Excel (AES ZIP via
`pyzipper` → `*.protected.zip`).

---

## 3. File / module map

```
CNCEC RL and BL Material Prediction Project/
├── app.py                       7,700+ LOC — the entire Streamlit UI
├── allocation_engine.py         Pure algorithms: demand matrix, cascade
│                                allocation, feasibility, suggestion engine,
│                                procurement list builder
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
transform lives here so the surface a developer needs to scan is **one
file**. Keep it that way unless a section is genuinely reusable.

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
| `equipment` | Equipment master from *Equipment.xlsx*. Includes `location`, `type`, `substrate`, `equipment_tag`, `surface_area_sqm`. |
| `sqm_progress` | Running tally per (equipment_tag, system_code): `original_sqm`, `done_sqm`. Driven by consumption_log SUMs but cached here for fast Dashboard reads. |
| `consumption_log` | Append-only daily entries. (date × equipment_tag × system_code × material_code) with `sqm_completed`, `expected_qty`, `consumed_qty`, variance. |
| `draft_consumption` | Pre-submit staging area for the multi-row consumption form. Cleared on submit. |
| `receipt_log` | Goods receipt entries (incoming stock). `received_qty` increases `inventory.available_qty`. |
| `orders_log` | Procurement orders raised through the app — `order_id`, `ordered_qty`, `fulfilled_qty`, `status`. Linked back via `receipt_log.order_id`. |
| `locations` | **Dynamic** location list (seeded with Brown Field / TRAIN J / TRAIN K). Editable via *Master Data → Add Location*. Drives every Location dropdown in the app. |

The schema is *evolving* — `setup_db.dynamic_sync_table()` will `ALTER
TABLE … ADD COLUMN` for any new column it sees in the source Excel. Don't
drop unknown columns; many are kept for downstream reporting.

---

## 7. UI structure (tabs and sub-views)

The app shows **9 top-level tabs**, declared in one place:

```python
# app.py ~ line 2861
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
| 3 | Location Report | All Equipment · Location Based | Same data sliced by location with per-location Excel + PDF download. |
| 4 | Equipment Report | – | Per-equipment deep-dive sheet — surface area, system codes, demand, fulfillment, materials. |
| 5 | Execution Plan | ⚙️ Execution Plan · 📋 Progress List · 📊 Consumption Comparison | Day-by-day execution; numbered production-detail blocks; expected vs actual variance. |
| 6 | Inventory | Inventory Dashboard · Consumption · Order Status · New Order · Receipt Log · Consumption Log | Operational data entry — daily consumption, goods receipts, raising orders. |
| 7 | Total Overview | – | Roll-up across all locations and codes. Independent of priority order. |
| 8 | Master Data | Equipment · LINING SYSTEM MATERIAL CONSM · Materials_DetailsAvailable_Qty · ➕ Add Location | CRUD on the three source tables + location management. |

Each sub-view is rendered by an `if exec_subview == "..."` / `if
md_table_sel == "..."` branch inside the relevant `with tabX:` block.

---

## 8. Critical algorithms (`allocation_engine.py`)

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

## 9. Cross-cutting helpers in `app.py`

| Helper | Where | What |
|---|---|---|
| `_dl_password()` | top of file | Returns the session download password, default `"smartmaterial"`. |
| `_encrypt_xlsx_bytes(raw, pwd)` | top of file | Wraps any `.xlsx` bytes in an AES-encrypted ZIP via `pyzipper`. |
| `_pdf_from_sheets([{name, df, title}, ...], pwd)` | top of file | Builds a multi-sheet landscape A4 PDF via ReportLab with `encrypt=pwd`. |
| `_secure_download_button(...)` | top of file | **Monkey-patches `st.download_button`.** Intercepts every `.xlsx` download and swaps it for `*.protected.zip`. `_orig_download_button` keeps an unpatched reference for the PDF helper. |
| `_pdf_download_button(...)` | top of file | Side-by-side PDF companion to every Excel download. Uses `_orig_download_button` so the PDF bytes don't get re-wrapped. |
| `_show_login()` | ~line 859 | Centered login card. Inputs wrapped in `st.form(...)` so **Enter** submits. |
| `_ensure_locations_table()` / `_refresh_location_order()` | ~line 935 | Boot-time: creates `locations` table + seeds defaults, refreshes the mutable module-level `LOCATION_ORDER` list. |
| `load_all()` | ~line 993 | `@st.cache_data` — single source of truth for in-memory DataFrames. Returns inv, recipe, equipment, sqm_progress, etc. |
| `cascade_allocate` / `tag_fulfillment` / `syscode_fulfillment` / `sqm_can_do` | ~line 1188 | Thin wrappers around the allocation engine. |
| `dbl_click_metric(label, value, key, drill_title, drill_df)` | ~line 1231 | KPI tile with click-to-expand drill-down DataFrame. Plain `st.metric` is used where no drill-down should appear (e.g. inside Session Order Report and Total Overview). |
| `generate_excel_report(df, title, color_scheme)` | ~line 1275 | xlsxwriter formatting with color schemes. |
| `generate_multi_sheet_excel(sheets)` | ~line 1406 | Multi-sheet Excel — one DataFrame per sheet, each with its own header strip. |
| `_equipment_report_excel(...)` / `_location_report_excel(...)` | ~line 1513 / 1863 | Big formatted multi-block per-equipment / per-location reports. |
| `_run_suggestion_engine(...)` / `render_suggestion_panel(...)` | ~line 2317 | The "best pause" predictive panel. |
| `loc_badge(loc)` / `status_dot(pct)` / `fulfil_pill(pct)` | ~line 1135 | Small HTML pill renderers. |

---

## 10. Theming & styling

- Theme palette is CSS-variable based, declared in one big `<style>` block
  near the top of `app.py` (~line 213 onwards).
- `_apply_theme_attr()` injects `data-theme="dark|light"` so dark/light
  toggling works without re-rendering Plotly charts.
- The sidebar has a "theme toggle pill" styled as `.sme-theme-toggle`.
- Color tokens used everywhere: `var(--t1..t5)` for text, `var(--bg1..bg3)`
  for backgrounds, `var(--amber)` for the brand accent.

When adding UI: prefer existing tokens over hex literals.

---

## 11. Login flow (current implementation)

```
_show_login() renders inside a horizontally-centered column [1, 1.4, 1]
  ↓
st.form("_login_form")        ← Enter key submits the form
  ├ Username
  ├ Password
  ├ Download Password  (≥ 4 chars; encrypts all downloads this session)
  └ st.form_submit_button("🔐  Login", type="primary")
        ↓ on success
   session_state["_authenticated"] = True
   session_state["_dl_pwd"]        = <download password>
   st.rerun()
```

**Admin credentials are hardcoded** at the top of `_show_login()`:

```python
_ADMIN_USER = "admin"
_ADMIN_PASS = "admin2026"
```

This is intentional for a desktop-only deployment. If multi-user is ever
needed, replace with a `users` table + bcrypt — but **keep the download
password as a separate, session-only secret** (it's never persisted).

---

## 12. Download / encryption model

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

- The Excel side is **automatically encrypted** by the monkey-patched
  `st.download_button` → user actually downloads `<stem>.protected.zip`
  containing `report.xlsx`. Password = the session download password.
- The PDF side uses `_pdf_download_button(...)` which calls
  `_orig_download_button` directly so the PDF bytes are not
  double-encrypted. ReportLab's `encrypt=` already password-protects the
  PDF.
- For multi-sheet exports pass `sheets=[{name, df, title}, ...]` instead
  of `df=`.

When adding a new download, **always provide both buttons.** Search the
codebase for `_pdf_download_button` for ~30 reference call sites.

---

## 13. Consumption entry — gotchas

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

## 14. Adding a new location end-to-end

Master Data → ➕ Add Location lets you add a row to the `locations` table.

What happens automatically afterward:

1. `_refresh_location_order()` reloads the in-memory `LOCATION_ORDER`
   list (it's a mutable module-level list — never reassigned, always
   `.clear()`+`.extend()`'d so every reference sees the update).
2. Every Location dropdown (Dashboard filter, Selective Equipment Entry,
   Location Report, Total Overview filter, Add-Equipment form) picks it
   up immediately.
3. To add equipment under the new location, switch the same Master Data
   tab to *Equipment* and use the Add-Row form — `location` will be in
   the dropdown.

Default locations (`Brown Field`, `TRAIN J`, `TRAIN K`) are
**delete-protected**. Custom locations can be removed only if no
equipment row references them.

---

## 15. Conventions to keep

- **One file, one source of truth.** Don't split `app.py` into modules
  just for readability — the search-with-grep workflow depends on it.
- **`st.cache_data.clear()` after every DB write.** This is the
  invalidation contract.
- **All downloads come in pairs (Excel + PDF).** No exceptions.
- **Widget keys must be globally unique** and prefixed by their parent
  context (e.g. `form_mat__{tag}__{code}__{idx}__{mc}`). If you reuse
  a key, Streamlit will silently share state across widgets.
- **No bare hex colors.** Use CSS variables.
- **`LOCATION_ORDER` is mutable.** Never do `LOCATION_ORDER = [...]` —
  always `LOCATION_ORDER.clear(); LOCATION_ORDER.extend(...)`.
- **`dbl_click_metric` vs `st.metric`** — use the latter where a popover
  drill-down would be noise (Session Order Report, Total Overview System
  Code tiles).

---

## 16. Known limitations / future work

- **Single-user only.** No row-level locking on `consumption_log` /
  `inventory`. Two concurrent users would race.
- **Hardcoded admin credentials.** See §11.
- **Excel column drift.** `setup_db.dynamic_sync_table` will add new
  columns but never rename or drop them. Stale columns accumulate over
  time — periodically inspect with `.schema` and clean up manually.
- **PDF page width.** Wide tables get column-truncated to 12 cols
  (`_df_for_pdf`). For very wide reports prefer multi-sheet Excel.
- **No automated tests.** The whole codebase relies on manual QA. Any
  refactor should add at least smoke tests around
  `allocation_engine.allocate_sequential` and `cascade_allocate`.
- **Streamlit hot-reload caveat.** Edits to imports or top-level
  decorators sometimes require a full process restart (`Ctrl-C` →
  `streamlit run app.py`) — not just the in-browser "Rerun".

---

## 17. Troubleshooting cheat sheet

| Symptom | Likely cause | Fix |
|---|---|---|
| `NameError: _dd_equip_df` (or similar `_dd_*_df`) on Dashboard | Variable defined inside one `if dash_view == ...` branch but referenced in the other | Hoist the assignment above the `if/else`. |
| `StreamlitDuplicateElementKey` on Consumption Select-All | Two material rows produced the same widget key | Confirm key includes `row_idx`: `f"form_mat__{tag}__{code}__{row_idx}__{mc}"`. |
| Downloads come out as `.protected.zip` instead of `.xlsx` | This is **expected** — that's the AES-encrypted wrapper. Open with the session download password. |
| PDF is empty or unopenable | `reportlab` not installed → `_HAS_REPORTLAB = False`; install with `pip install reportlab>=4.0.0`. |
| Port 8501 already in use | Another Streamlit app is running (e.g. another project) | `streamlit run app.py --server.port 8502`. |
| Master Data edits don't reflect on Dashboard | Missing `st.cache_data.clear()` after the mutation | Add it immediately after the `commit()`. |
| New location doesn't show in any dropdown | `_refresh_location_order()` not called after insert | Call it right after committing the new row. |
| Login button does nothing | You're outside `st.form` — `st.form_submit_button` only works inside one. |

---

## 18. Where to start when picking this up cold

1. Run the app (§2), log in with `admin / admin2026 / smart1234`.
2. Click through every tab once. Note which sub-views exist.
3. Open `app.py` and `Cmd-F` for the section banner of the tab you want
   to change (e.g. `TAB 5 · EXECUTION PLAN`).
4. For algorithm changes, open `allocation_engine.py` — it's pure pandas
   and easy to unit-test in a notebook.
5. For schema changes, edit `setup_db.py` AND add the column in the
   corresponding `validate_data.clean_*` cleaner so a fresh-DB rebuild
   still works.
6. Before committing: run the app, confirm login, dashboard, one report
   download (Excel + PDF), one consumption entry. That's the minimum
   smoke test.

— End of handoff —
