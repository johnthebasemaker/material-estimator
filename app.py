"""
Smart Material Estimator · app.py (v4 — redesigned)

Run: streamlit run app.py

v4 changes vs v3:
  • New IA: 6 flat tabs (was 8). Reports merges Session + Location + Total Overview.
  • Reference-aligned visual system: dark navy, amber accent, JetBrains Mono numerics,
    SVG charts (replaces Plotly), HTML mat-tables with row-tinted Fulfil%.
  • Subtle glassmorphism: 14px backdrop-blur, semi-transparent cards over dark base,
    barrier layer for WCAG-compliant text contrast.
  • Backend preserved: load_all, cascade_allocate, DB write protocol unchanged.
  • Old v3 backup at app_v3_backup.py.
"""
import io, os, sys, sqlite3, base64, uuid
from datetime import date, datetime
import pandas as pd
import numpy as np
import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# ───────────────────────────────────────────────────────────────────────────────
# CONSTANTS & PATHS
# ───────────────────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, "sme_database.db")
LOGO_PATH = os.path.join(BASE_DIR, "logo.png")
PATH_A   = os.path.join(BASE_DIR, "Materials_DetailsAvailable_Qty.xlsx")
PATH_B   = os.path.join(BASE_DIR, "For_1_SQM.xlsx")
PATH_C   = os.path.join(BASE_DIR, "Equipment.xlsx")
SHEET_A, SHEET_B, SHEET_C = "Materials", "LINING SYSTEM MATERIAL CONSM", "Data Input"
LOCATION_ORDER = ["Brown Field", "TRAIN J", "TRAIN K"]

_ADMIN_USER = "admin"
_ADMIN_PASS = "admin2026"

# ───────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ───────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Smart Material Estimator & Planner",
    page_icon="🏗",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ───────────────────────────────────────────────────────────────────────────────
# DESIGN TOKENS + INJECTED CSS (subtle glassmorphism over dark navy)
# ───────────────────────────────────────────────────────────────────────────────
GLASS_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500;600;700;800&display=swap');

:root {
  --bg-base: #0F172A;
  --bg-deep: #0B1220;
  --bg-card: rgba(30,41,59,0.55);
  --bg-card-solid: #1E293B;
  --bg-inset: #0F172A;
  --bg-nested: #162032;

  --border-soft: rgba(148,163,184,0.10);
  --border-mid:  rgba(148,163,184,0.18);
  --border-strong: rgba(148,163,184,0.28);

  --t0:#F8FAFC;
  --t1:#CBD5E1;
  --t2:#94A3B8;
  --t3:#64748B;
  --t4:#475569;
  --t5:#334155;

  --amber:#F59E0B; --amber2:#FCD34D; --amber3:#D97706;
  --amber-bg:rgba(245,158,11,0.15); --amber-soft:rgba(245,158,11,0.10);
  --green:#10B981; --green-bg:rgba(16,185,129,0.15); --green-soft:rgba(16,185,129,0.08);
  --red:#EF4444;   --red-bg:rgba(239,68,68,0.15);    --red-soft:rgba(239,68,68,0.08);
  --orange:#F97316;--orange-bg:rgba(249,115,22,0.15);
  --yellow:#EAB308;--yellow-bg:rgba(234,179,8,0.15);
  --blue:#3B82F6;  --blue-bg:rgba(59,130,246,0.15); --blue-soft:rgba(59,130,246,0.08);

  --r-sm:6px; --r-md:8px; --r-lg:12px; --r-xl:16px;
}

/* ── Base + Streamlit overrides ── */
html, body, [class*="css"], .stApp {
  background: var(--bg-base) !important;
  color: var(--t1);
  font-family: 'Inter', -apple-system, sans-serif !important;
}
.stApp::before {
  content:''; position:fixed; inset:0; z-index:0; pointer-events:none;
  background:
    radial-gradient(circle at 12% 10%, rgba(245,158,11,0.06) 0%, transparent 38%),
    radial-gradient(circle at 88% 80%, rgba(59,130,246,0.05) 0%, transparent 42%);
}
.main .block-container {
  padding: 0.5rem 1.5rem 2rem 1.5rem !important;
  max-width: 1600px !important;
  position: relative; z-index: 1;
}
.stApp > header { background: transparent !important; }
#MainMenu, footer, [data-testid="stStatusWidget"] { visibility: hidden; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
  background: var(--bg-deep) !important;
  border-right: 1px solid var(--border-soft) !important;
}
[data-testid="stSidebar"]::before {
  content:''; display:block; height:3px;
  background: linear-gradient(90deg, var(--amber3), var(--amber), var(--amber2));
}
[data-testid="stSidebar"] *, [data-testid="stSidebar"] .stMarkdown {
  font-family:'Inter',sans-serif !important; color: var(--t1);
}

/* ── Scrollbars ── */
::-webkit-scrollbar { width:6px; height:6px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: var(--border-strong); border-radius: 99px; }
::-webkit-scrollbar-thumb:hover { background: var(--amber); }

/* ── Glass card primitive ── */
.glass {
  background: var(--bg-card);
  backdrop-filter: blur(14px) saturate(120%);
  -webkit-backdrop-filter: blur(14px) saturate(120%);
  border: 1px solid var(--border-mid);
  border-radius: var(--r-lg);
  padding: 18px;
  position: relative;
}
.glass-flat {
  background: var(--bg-card-solid);
  border: 1px solid var(--border-mid);
  border-radius: var(--r-lg);
  padding: 18px;
}
.glass-tight { padding: 14px; }
.glass-amber-bar { border-left: 3px solid var(--amber); }

/* ── Section header (orange tick + uppercase label) ── */
.sec-hdr {
  display:flex; align-items:center; gap:8px; margin: 4px 0 14px 0;
}
.sec-hdr-bar { width:3px; height:13px; background: var(--amber); border-radius: 2px; flex-shrink: 0; }
.sec-hdr-txt {
  font-size: 11px; font-weight: 700; color: var(--t2);
  text-transform: uppercase; letter-spacing: 0.8px;
}

/* ── KPI tile ── */
.kpi {
  background: var(--bg-card);
  backdrop-filter: blur(10px);
  border: 1px solid var(--border-mid);
  border-radius: 10px;
  padding: 14px 16px;
  transition: all .15s;
}
.kpi:hover { border-color: rgba(245,158,11,0.5); box-shadow: 0 0 0 1px rgba(245,158,11,0.12); }
.kpi-label {
  font-size: 10px; font-weight: 600; color: var(--t3);
  text-transform: uppercase; letter-spacing: .7px; margin-bottom: 7px;
}
.kpi-value {
  font-size: 20px; font-weight: 800;
  font-family: 'JetBrains Mono', monospace; line-height: 1;
}
.kpi-sub { font-size: 10px; color: var(--t4); margin-top: 5px; }

/* ── Pills + badges ── */
.pill {
  display: inline-block; padding: 2px 9px; border-radius: 20px;
  font-size: 11px; font-weight: 700; font-family: 'JetBrains Mono', monospace;
  white-space: nowrap; border-width:1px; border-style:solid;
}
.pill-g { background: rgba(16,185,129,.13); color: var(--green);  border-color: rgba(16,185,129,.35); }
.pill-o { background: rgba(249,115,22,.13); color: var(--orange); border-color: rgba(249,115,22,.35); }
.pill-y { background: rgba(234,179,8,.13);  color: var(--yellow); border-color: rgba(234,179,8,.35); }
.pill-r { background: rgba(239,68,68,.13);  color: var(--red);    border-color: rgba(239,68,68,.35); }

.loc-badge {
  display: inline-block; padding: 2px 8px; border-radius: 6px;
  font-size: 10px; font-weight: 700; white-space: nowrap;
  border-width:1px; border-style:solid;
}
.loc-bf { background: rgba(59,130,246,.13);  color: var(--blue);  border-color: rgba(59,130,246,.35); }
.loc-tj { background: rgba(245,158,11,.13);  color: var(--amber); border-color: rgba(245,158,11,.35); }
.loc-tk { background: rgba(16,185,129,.13);  color: var(--green); border-color: rgba(16,185,129,.35); }

.code-badge {
  display: inline-block; padding: 2px 7px; border-radius: 4px;
  font-size: 11px; font-weight: 700; font-family: 'JetBrains Mono', monospace;
  background: var(--amber-bg); color: var(--amber);
  border: 1px solid rgba(245,158,11,0.3);
}

.dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; vertical-align: middle; }
.dot-g { background: var(--green); }
.dot-o { background: var(--orange); }
.dot-y { background: var(--yellow); }
.dot-r { background: var(--red); }

/* ── Material table (HTML) ── */
.mat-tbl-wrap {
  overflow-x: auto;
  border-radius: 8px;
  border: 1px solid var(--border-mid);
  background: var(--bg-card-solid);
}
table.mat-tbl {
  width: 100%; border-collapse: collapse; font-size: 12px;
  font-family: 'Inter', sans-serif;
}
table.mat-tbl th {
  padding: 9px 12px;
  background: var(--bg-inset);
  color: var(--t3);
  font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;
  border-bottom: 1px solid var(--border-mid);
  white-space: nowrap; text-align: left;
}
table.mat-tbl th.num { text-align: right; }
table.mat-tbl td {
  padding: 8px 12px;
  border-bottom: 1px solid rgba(51,65,85,0.4);
  color: var(--t1);
}
table.mat-tbl td.code { color: var(--t2); font-family: 'JetBrains Mono', monospace; font-size: 11px; }
table.mat-tbl td.uom  { color: var(--t3); font-size: 11px; }
table.mat-tbl td.num  { font-family: 'JetBrains Mono', monospace; font-size: 11px; text-align: right; }
table.mat-tbl td.num-red   { color: var(--red); }
table.mat-tbl td.num-green { color: var(--green); }
table.mat-tbl tr.row-g td:last-child { color: var(--green);  font-weight: 700; }
table.mat-tbl tr.row-o td:last-child { color: var(--orange); font-weight: 700; }
table.mat-tbl tr.row-y td:last-child { color: var(--yellow); font-weight: 700; }
table.mat-tbl tr.row-r td:last-child { color: var(--red);    font-weight: 700; }
table.mat-tbl tr.row-g { background: rgba(16,185,129,.045); }
table.mat-tbl tr.row-o { background: rgba(249,115,22,.045); }
table.mat-tbl tr.row-y { background: rgba(234,179,8,.045); }
table.mat-tbl tr.row-r { background: rgba(239,68,68,.05); }

/* ── Sticky 3-tier: header (z=100) + tab bar (z=99) + sub-mode radio (z=98) ── */
.sme-header-sticky {
  position: sticky; top: 0; z-index: 100;
  background: var(--bg-base); padding: 10px 0 0 0;
  margin: -8px 0 0 0;
  border-bottom: 1px solid var(--border-mid);
}
[data-testid="stTabs"] [data-baseweb="tab-list"] {
  background: var(--bg-base) !important;
  border-bottom: 1px solid var(--border-mid);
  gap: 4px; padding: 0 4px;
  position: sticky; top: 72px; z-index: 99;
  backdrop-filter: blur(10px);
  -webkit-backdrop-filter: blur(10px);
}
/* The submode marker is an invisible element-container we drop right before
   each radio sub-mode toggle. Using :has() to find its sibling element-container
   that holds the radio, then making that sibling sticky. */
.sme-submode-marker { display:none; }
[data-testid="stElementContainer"]:has(.sme-submode-marker) + [data-testid="stElementContainer"] {
  position: sticky;
  top: 116px;
  z-index: 98;
  background: var(--bg-base);
  padding: 6px 0 2px 0;
  border-bottom: 1px solid var(--border-soft);
  backdrop-filter: blur(8px);
  -webkit-backdrop-filter: blur(8px);
}
[data-testid="stTabs"] [data-baseweb="tab"] {
  background: transparent !important;
  color: var(--t3) !important;
  font-family: 'Inter', sans-serif !important;
  font-size: 12.5px !important; font-weight: 500 !important;
  padding: 10px 16px !important;
  border-bottom: 2px solid transparent !important;
  border-radius: 0 !important;
  transition: all .15s;
}
[data-testid="stTabs"] [data-baseweb="tab"]:hover { color: var(--t1) !important; }
[data-testid="stTabs"] [data-baseweb="tab"][aria-selected="true"] {
  color: var(--amber) !important;
  border-bottom-color: var(--amber) !important;
  font-weight: 600 !important;
}
[data-testid="stTabs"] [data-baseweb="tab-highlight"] { display: none; }
[data-testid="stTabs"] [data-baseweb="tab-panel"] { padding-top: 22px; }

/* ── Buttons ── */
.stButton > button, .stDownloadButton > button {
  background: transparent !important;
  border: 1px solid var(--border-strong) !important;
  border-radius: 8px !important;
  color: var(--t2) !important;
  font-family: 'Inter', sans-serif !important;
  font-size: 12px !important; font-weight: 500 !important;
  padding: 6px 14px !important;
  transition: all .15s;
}
.stButton > button:hover, .stDownloadButton > button:hover {
  background: var(--amber-soft) !important;
  border-color: rgba(245,158,11,0.4) !important;
  color: var(--amber) !important;
}
.stButton > button[kind="primary"] {
  background: linear-gradient(135deg, var(--amber3), var(--amber)) !important;
  border: none !important; color: #fff !important; font-weight: 700 !important;
}

/* ── Inputs / selects ── */
.stTextInput input, .stNumberInput input, .stDateInput input, .stTextArea textarea {
  background: var(--bg-inset) !important;
  border: 1px solid var(--border-mid) !important;
  color: var(--t0) !important;
  border-radius: 8px !important;
  font-family: 'Inter', sans-serif !important;
}
.stTextInput input:focus, .stNumberInput input:focus, .stDateInput input:focus {
  border-color: var(--amber) !important;
  box-shadow: 0 0 0 2px rgba(245,158,11,0.15) !important;
}
.stSelectbox > div > div, .stMultiSelect > div > div {
  background: var(--bg-inset) !important;
  border: 1px solid var(--border-mid) !important;
  border-radius: 8px !important;
}

/* ── Compact multiselect: cap to one line, show chip-count on overflow ── */
[data-testid="stMultiSelect"] [data-baseweb="select"] > div:first-child {
  max-height: 40px !important;
  overflow: hidden !important;
  flex-wrap: nowrap !important;
  position: relative;
}
[data-testid="stMultiSelect"] [data-baseweb="tag"] {
  flex-shrink: 0 !important;
  max-width: 130px !important;
}
[data-testid="stMultiSelect"] [data-baseweb="tag"] span {
  white-space: nowrap !important;
  overflow: hidden !important;
  text-overflow: ellipsis !important;
}
/* Subtle right-edge fade hint that more chips are hidden */
[data-testid="stMultiSelect"] [data-baseweb="select"] > div:first-child::after {
  content: '';
  position: absolute; top: 0; right: 0; width: 40px; height: 100%;
  background: linear-gradient(90deg, transparent, var(--bg-inset) 70%);
  pointer-events: none;
}
.stRadio > div { gap: 6px; flex-wrap: wrap; }
.stRadio label, .stCheckbox label { color: var(--t1) !important; font-family: 'Inter' !important; }
.stRadio label p, .stCheckbox label p { color: var(--t1) !important; }

/* ── Metric (native fallback if used) ── */
[data-testid="stMetric"] {
  background: var(--bg-card-solid);
  border: 1px solid var(--border-mid);
  border-radius: 10px;
  padding: 12px 14px;
}
[data-testid="stMetricLabel"] {
  color: var(--t3) !important; font-size: 10px !important;
  text-transform: uppercase; letter-spacing: .7px;
}
[data-testid="stMetricValue"] {
  color: var(--amber) !important;
  font-family: 'JetBrains Mono', monospace !important;
  font-size: 20px !important; font-weight: 800 !important;
}

/* ── Dataframe polish ── */
[data-testid="stDataFrame"] {
  background: var(--bg-card-solid) !important;
  border: 1px solid var(--border-mid) !important;
  border-radius: 10px !important;
}

/* ── Expander ── */
[data-testid="stExpander"] {
  background: var(--bg-card-solid) !important;
  border: 1px solid var(--border-mid) !important;
  border-radius: 10px !important;
  margin-bottom: 8px;
}
[data-testid="stExpander"] summary { font-family: 'Inter', sans-serif !important; }
[data-testid="stExpander"] summary p { color: var(--t1) !important; }

/* ── Alert blocks ── */
[data-testid="stAlert"] {
  background: var(--bg-card-solid) !important;
  border: 1px solid var(--border-mid) !important;
  border-radius: 8px !important;
}

/* ── Misc ── */
hr { border-color: var(--border-mid) !important; }
.grand-box {
  background: linear-gradient(135deg, rgba(245,158,11,.10), rgba(245,158,11,.03));
  border: 1px solid rgba(245,158,11,.25);
  border-radius: 12px; padding: 15px;
}
.label-row { font-size: 10px; color: var(--t3); margin-bottom: 4px; text-transform: uppercase; letter-spacing: .5px; }
.value-mono { font-family: 'JetBrains Mono', monospace; }

@keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.35} }
.live-dot {
  display:inline-block; width:8px; height:8px; border-radius:50%;
  background: var(--green); box-shadow: 0 0 8px var(--green);
  animation: pulse 2.5s ease infinite; margin-right:5px;
}

/* Hide raw markdown bullets next to our HTML chips */
.stMarkdown ul { list-style: none; padding-left: 0; }

@media print {
  [data-testid="stSidebar"], .no-print,
  .sme-header-sticky, [data-testid="stTabs"] [data-baseweb="tab-list"] {
    display: none !important;
  }
  .main .block-container { max-width: 100% !important; padding: 0 !important; }
  .stApp::before { display: none !important; }
  .glass, .glass-flat { break-inside: avoid; }
  body { background: white !important; color: black !important; }
}
</style>
"""
st.markdown(GLASS_CSS, unsafe_allow_html=True)

# ───────────────────────────────────────────────────────────────────────────────
# LOGO
# ───────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def _logo_b64() -> str:
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""

# ───────────────────────────────────────────────────────────────────────────────
# DB HELPERS
# ───────────────────────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def db_available() -> bool:
    return os.path.exists(DB_PATH)

def _next_order_id(conn) -> str:
    today = date.today().strftime("%Y%m%d")
    prefix = f"ORD-{today}-"
    row = conn.execute(
        "SELECT order_id FROM orders_log WHERE order_id LIKE ? ORDER BY id DESC LIMIT 1",
        (prefix + "%",),
    ).fetchone()
    n = int(row["order_id"].split("-")[-1]) + 1 if row else 1
    return f"{prefix}{n:03d}"

def _commit_and_refresh(conn):
    """Mandatory DB-write protocol: commit → cache clear → rerun."""
    conn.commit()
    st.cache_data.clear()
    st.rerun()

# ───────────────────────────────────────────────────────────────────────────────
# DATA LOADING (preserved from v3)
# ───────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Loading project data…")
def load_all():
    if db_available():
        conn = get_db()
        inv = pd.read_sql(
            "SELECT material_code AS Material_Code, material_name AS Material_Name, "
            "nature AS Nature, uom AS UOM, "
            "available_qty AS Available_Qty, ordered_qty AS Ordered_Qty "
            "FROM inventory", conn)
        recipe = pd.read_sql(
            "SELECT lining_system_code AS Lining_System_Code, "
            "lining_system_short_name AS Lining_System_Short_Name, "
            "lining_type AS Lining_Type, material_code AS Material_Code, "
            "material_description AS Material_Description, "
            "material_name AS Material_Name, for_1_sqm AS For_1_SQM, uom AS UOM "
            "FROM recipe", conn)
        equip_raw = pd.read_sql(
            "SELECT location AS Location, type AS Type, "
            "lining_system_code AS Lining_System_Code, "
            "lining_system_short_name AS Lining_System_Short_Name, "
            "lining_type AS Lining_Type, equipment_tag AS \"Equipment_Tag_No.\", "
            "name AS Name, substrate AS Substrate, "
            "material_spec AS \"Material Spec.\", design AS Design, "
            "surface_area_sqm AS Surface_Area_SQM, lining_systems AS \"Lining_System+\" "
            "FROM equipment", conn)
        sqm_prog = pd.read_sql(
            "SELECT equipment_tag AS \"Equipment_Tag_No.\", "
            "lining_system_code AS Lining_System_Code, "
            "original_sqm, done_sqm, "
            "(original_sqm - done_sqm) AS remaining_sqm "
            "FROM sqm_progress", conn)
        conn.close()

        equip_sc = (equip_raw
            .groupby(["Equipment_Tag_No.", "Lining_System_Code", "Lining_System_Short_Name"],
                     as_index=False)["Surface_Area_SQM"].sum()
            .rename(columns={"Surface_Area_SQM": "Total_SQM_Original"}))
        equip_sc = equip_sc.merge(
            sqm_prog[["Equipment_Tag_No.", "Lining_System_Code", "remaining_sqm", "done_sqm"]],
            on=["Equipment_Tag_No.", "Lining_System_Code"], how="left")
        equip_sc["remaining_sqm"] = equip_sc["remaining_sqm"].fillna(equip_sc["Total_SQM_Original"])
        equip_sc["done_sqm"]      = equip_sc["done_sqm"].fillna(0)
        equip_sc["Total_SQM"]     = equip_sc["remaining_sqm"]
    else:
        from validate_data import clean_inventory, clean_recipe, clean_equipment
        df_a_raw = pd.read_excel(PATH_A, sheet_name=SHEET_A)
        df_b_raw = pd.read_excel(PATH_B, sheet_name=SHEET_B)
        df_c_raw = pd.read_excel(PATH_C, sheet_name=SHEET_C)
        inv    = clean_inventory(df_a_raw)
        recipe = clean_recipe(df_b_raw)
        equip_raw = clean_equipment(df_c_raw)
        inv_full = df_a_raw.copy(); inv_full.columns = inv_full.columns.str.strip()
        inv_full["Material_Code"] = inv_full["Material_Code"].astype(str).str.strip()
        ordered_col = next((c for c in inv_full.columns
                            if c.strip() in ("Ordered_Qty","Balance To Be Received")), None)
        if ordered_col:
            inv_full[ordered_col] = pd.to_numeric(inv_full[ordered_col], errors="coerce").fillna(0)
            inv_ordered = inv_full.groupby("Material_Code", as_index=False).agg(Ordered_Qty=(ordered_col,"sum"))
            inv = inv.merge(inv_ordered, on="Material_Code", how="left")
            inv["Ordered_Qty"] = inv["Ordered_Qty"].fillna(0)
        else:
            inv["Ordered_Qty"] = 0.0
        equip_sc = (equip_raw
            .groupby(["Equipment_Tag_No.","Lining_System_Code","Lining_System_Short_Name"],
                     as_index=False)["Surface_Area_SQM"].sum()
            .rename(columns={"Surface_Area_SQM":"Total_SQM_Original"}))
        equip_sc["done_sqm"] = 0.0
        equip_sc["remaining_sqm"] = equip_sc["Total_SQM_Original"]
        equip_sc["Total_SQM"] = equip_sc["Total_SQM_Original"]

    dm = equip_sc.merge(recipe, on="Lining_System_Code", suffixes=("_e","_r"))
    dm["Demand_Qty"] = dm["For_1_SQM"] * dm["Total_SQM"]
    if "Lining_System_Short_Name_e" in dm.columns:
        dm = dm.rename(columns={"Lining_System_Short_Name_e":"Lining_System_Short_Name"})
        dm.drop(columns=["Lining_System_Short_Name_r"], inplace=True, errors="ignore")
    dm = dm[["Equipment_Tag_No.","Lining_System_Code","Lining_System_Short_Name",
             "Total_SQM","Material_Code","Material_Name","UOM","Demand_Qty"]]

    eq_master = equip_raw.groupby("Equipment_Tag_No.", as_index=False).agg(
        Name          =("Name",          "first"),
        Substrate     =("Substrate",      "first"),
        Location      =("Location",      "first"),
        Type          =("Type",          "first"),
        Lining_Systems=("Lining_System+","first"),
        Lining_Type   =("Lining_Type",   "first"),
        Material_Spec =("Material Spec.","first"),
        Design        =("Design",        "first"),
        Total_SQM     =("Surface_Area_SQM","sum"),
    )
    eq_master["Location"] = eq_master["Location"].astype(str).str.strip()
    eq_master["Type"]     = eq_master["Type"].astype(str).str.strip()
    sqm_ref = equip_sc[["Equipment_Tag_No.","Lining_System_Code",
                         "Total_SQM","Total_SQM_Original","done_sqm"]].drop_duplicates()
    return inv, recipe, equip_sc, dm, eq_master, sqm_ref


# ───────────────────────────────────────────────────────────────────────────────
# AUTH GATE
# ───────────────────────────────────────────────────────────────────────────────
if "_authenticated" not in st.session_state:
    st.session_state["_authenticated"] = False

def show_login():
    logo = _logo_b64()
    logo_html = (f'<img src="data:image/png;base64,{logo}" style="height:64px;margin-bottom:14px;">'
                 if logo else '<div style="font-size:44px;margin-bottom:10px;">🏗</div>')
    st.markdown(f"""
    <div style="min-height:80vh;display:flex;align-items:center;justify-content:center;">
      <div style="width:420px;background:rgba(30,41,59,0.55);backdrop-filter:blur(20px) saturate(120%);
                  -webkit-backdrop-filter:blur(20px) saturate(120%);
                  border:1px solid var(--border-mid);border-radius:16px;padding:42px 38px;
                  box-shadow:0 32px 64px rgba(0,0,0,.55);position:relative;overflow:hidden;">
        <div style="position:absolute;top:0;left:0;right:0;height:3px;
                    background:linear-gradient(90deg,var(--amber3),var(--amber),var(--amber2));"></div>
        <div style="text-align:center;margin-bottom:30px;">
          {logo_html}
          <div style="font-size:21px;font-weight:800;color:var(--t0);letter-spacing:-.5px;">Smart Material Estimator</div>
          <div style="font-size:10px;color:var(--t4);margin-top:5px;
                      text-transform:uppercase;letter-spacing:1.8px;">Enterprise Platform · v4</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        user = st.text_input("Username", key="_login_user", placeholder="Enter username")
        pwd  = st.text_input("Password", key="_login_pass", type="password", placeholder="Enter password")
        if st.button("🔐  Login", use_container_width=True, key="_login_btn", type="primary"):
            if user == _ADMIN_USER and pwd == _ADMIN_PASS:
                st.session_state["_authenticated"] = True
                st.rerun()
            else:
                st.error("❌ Invalid credentials. Please try again.")
        st.caption("Demo: admin / admin2026")

if not st.session_state["_authenticated"]:
    show_login()
    st.stop()

# ───────────────────────────────────────────────────────────────────────────────
# LOAD DATA (after auth)
# ───────────────────────────────────────────────────────────────────────────────
inv, recipe, equip_sc, dm, eq_master, sqm_ref = load_all()
ALL_TAGS         = sorted(eq_master["Equipment_Tag_No."].tolist())
INV_POOL_INIT    = inv.set_index("Material_Code")["Available_Qty"].to_dict()
INV_ORDERED_INIT = inv.set_index("Material_Code")["Ordered_Qty"].to_dict()

# ── SESSION STATE init ─────────
for k, v in {
    "session_tags": [],
    "_session_key": str(uuid.uuid4()),
    "loc_order": {},
    "all_eq_order": list(ALL_TAGS),
    "_ord_draft": [],
    "_ord_source_detail": "",
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ───────────────────────────────────────────────────────────────────────────────
# COMPONENT HELPERS
# ───────────────────────────────────────────────────────────────────────────────
def fc(pct):
    if pct >= 100: return "#10B981"
    if pct >= 90:  return "#F97316"
    if pct >= 80:  return "#EAB308"
    return "#EF4444"

def row_cls(pct):
    if pct >= 100: return "row-g"
    if pct >= 90:  return "row-o"
    if pct >= 80:  return "row-y"
    return "row-r"

def lc(loc):
    return {"Brown Field":"#3B82F6","TRAIN J":"#F59E0B","TRAIN K":"#10B981"}.get(loc,"#3B82F6")

def fulfil_pill(pct):
    cls = "pill-g" if pct>=100 else "pill-o" if pct>=90 else "pill-y" if pct>=80 else "pill-r"
    return f'<span class="pill {cls}">{pct:.1f}%</span>'

def loc_badge(loc):
    cls = {"Brown Field":"loc-bf","TRAIN J":"loc-tj","TRAIN K":"loc-tk"}.get(loc,"loc-bf")
    return f'<span class="loc-badge {cls}">{loc}</span>'

def code_badge(code):
    return f'<span class="code-badge">C{code}</span>'

def status_dot(pct):
    cls = "dot-g" if pct>=100 else "dot-o" if pct>=90 else "dot-y" if pct>=80 else "dot-r"
    return f'<span class="dot {cls}"></span>'

def sec_header(text: str):
    st.markdown(
        f'<div class="sec-hdr"><div class="sec-hdr-bar"></div>'
        f'<span class="sec-hdr-txt">{text}</span></div>',
        unsafe_allow_html=True,
    )

def submode_radio(label: str, options: list[str], key: str) -> str:
    """A horizontal radio that becomes sticky under the tab bar.
       Drop a marker first; CSS :has() makes the very next element sticky."""
    st.markdown('<div class="sme-submode-marker"></div>', unsafe_allow_html=True)
    return st.radio(label, options, horizontal=True, key=key,
                     label_visibility="collapsed")

def popover_multiselect(label: str, options: list[str], default: list[str] | None,
                         key: str, help: str = "") -> list[str]:
    """Click-to-open popover that holds a multiselect.
       The trigger button shows 'Label · N selected'."""
    default = default if default is not None else list(options)
    state_key = f"_pms_{key}"
    if state_key not in st.session_state:
        st.session_state[state_key] = [d for d in default if d in options]
    sel = [s for s in st.session_state[state_key] if s in options]
    if len(options) == 0:
        st.markdown(
            f'<div class="label-row">{label}</div>'
            f'<div style="background:var(--bg-inset);border:1px solid var(--border-mid);'
            f'border-radius:8px;padding:9px 12px;color:var(--t4);font-size:12px;">— No options —</div>',
            unsafe_allow_html=True)
        return []
    summary = f"{label}  ·  {len(sel)} of {len(options)} selected"
    with st.popover(summary, use_container_width=True, help=help or None):
        # Quick action row
        ac1, ac2 = st.columns(2)
        with ac1:
            if st.button("Select all", key=f"{key}_all", use_container_width=True):
                st.session_state[state_key] = list(options)
                st.rerun()
        with ac2:
            if st.button("Clear", key=f"{key}_clear", use_container_width=True):
                st.session_state[state_key] = []
                st.rerun()
        picked = st.multiselect(
            label, options=options,
            default=sel,
            key=f"{key}_ms",
            label_visibility="collapsed",
        )
        if picked != st.session_state[state_key]:
            st.session_state[state_key] = picked
    return st.session_state[state_key]

def kpi_tile(label: str, value: str, sub: str = "", color: str = None) -> str:
    color = color or "var(--amber)"
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    return (
        f'<div class="kpi">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value" style="color:{color};">{value}</div>'
        f'{sub_html}</div>'
    )

def kpi_strip(items: list[tuple], cols: int = None):
    cols = cols or len(items)
    html = '<div style="display:grid;grid-template-columns:repeat(' + str(cols) + ',1fr);gap:10px;margin-bottom:18px;">'
    for it in items:
        label, value, sub, color = (it + (None,)*4)[:4]
        html += kpi_tile(label, value, sub or "", color)
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)

# ── Material table (HTML, color-coded rows) ──────────────────────────────
def mat_table_html(df: pd.DataFrame, columns: list[str] = None) -> str:
    """Render a material table as an HTML <table>. Columns can be any of:
       Code, Material Name, UOM, Available, On Order, Demand, Allocated,
       Shortfall, Net Shortfall, Fulfil %, Coverage %, Total Receipts,
       Total Consumed, Current Stock, Ordered Qty."""
    if df.empty:
        return '<div class="glass-flat" style="text-align:center;color:var(--t3);padding:18px;">No data.</div>'
    columns = columns or ["Code","Material Name","UOM","Available","Demand","Shortfall","Fulfil %"]
    numeric_cols = {"Available","On Order","Demand","Allocated","Shortfall","Net Shortfall",
                    "Fulfil %","Coverage %","Total Receipts","Total Consumed",
                    "Current Stock","Ordered Qty"}
    header = "".join(
        f'<th class="num">{c}</th>' if c in numeric_cols else f'<th>{c}</th>'
        for c in columns
    )
    body = []
    for _, r in df.iterrows():
        pct = r.get("Fulfil %", r.get("Coverage %", 100))
        try: pct = float(pct)
        except: pct = 0.0
        cls = row_cls(pct)
        cells = []
        for c in columns:
            v = r.get(c, "")
            if c == "Code":
                cells.append(f'<td class="code">{v}</td>')
            elif c == "Material Name":
                cells.append(f'<td>{v}</td>')
            elif c == "UOM":
                cells.append(f'<td class="uom">{v}</td>')
            elif c in ("Shortfall","Net Shortfall"):
                try: vf = float(v)
                except: vf = 0
                extra = "num-red" if vf > 0.001 else "num-green"
                cells.append(f'<td class="num {extra}">{vf:,.3f}</td>')
            elif c in ("Available","Demand","Allocated","On Order","Total Receipts",
                       "Total Consumed","Current Stock","Ordered Qty"):
                try: vf = float(v)
                except: vf = 0
                cells.append(f'<td class="num">{vf:,.3f}</td>')
            elif c in ("Fulfil %","Coverage %"):
                try: vf = float(v)
                except: vf = 0
                cells.append(f'<td class="num">{vf:.1f}%</td>')
            else:
                cells.append(f'<td>{v}</td>')
        body.append(f'<tr class="{cls}">' + "".join(cells) + "</tr>")
    return (
        '<div class="mat-tbl-wrap"><table class="mat-tbl">'
        f'<thead><tr>{header}</tr></thead>'
        f'<tbody>{"".join(body)}</tbody>'
        '</table></div>'
    )

# ── SVG gauge chart ───────────────────────────────────────────────────────
def gauge_svg(pct: float, demand_sqm: float, alloc_sqm: float) -> str:
    pct = max(0, min(100, pct))
    col = fc(pct)
    w, h, cx, cy, R = 300, 168, 150, 158, 115
    import math
    sA = -math.pi
    def arc(r, s2, e2):
        x1 = cx + r*math.cos(s2); y1 = cy + r*math.sin(s2)
        x2 = cx + r*math.cos(e2); y2 = cy + r*math.sin(e2)
        large = 1 if (e2 - s2) > math.pi else 0
        return f"M {x1:.1f} {y1:.1f} A {r} {r} 0 {large} 1 {x2:.1f} {y2:.1f}"
    vA = sA + (pct/100) * math.pi
    return f"""<svg width="100%" height="{h}" viewBox="0 0 {w} {h}" style="display:block;">
      <path d="{arc(R, sA, sA+.5*3.14159)}" fill="none" stroke="rgba(239,68,68,.18)" stroke-width="22"/>
      <path d="{arc(R, sA+.5*3.14159, sA+.7*3.14159)}" fill="none" stroke="rgba(234,179,8,.18)" stroke-width="22"/>
      <path d="{arc(R, sA+.7*3.14159, sA+.85*3.14159)}" fill="none" stroke="rgba(249,115,22,.18)" stroke-width="22"/>
      <path d="{arc(R, sA+.85*3.14159, 0)}" fill="none" stroke="rgba(16,185,129,.18)" stroke-width="22"/>
      <path d="{arc(R, sA, 0)}" fill="none" stroke="#1E293B" stroke-width="20"/>
      {'<path d="' + arc(R, sA, vA) + f'" fill="none" stroke="{col}" stroke-width="20" stroke-linecap="round"/>' if pct > 0 else ''}
      <text x="{cx-R+2}" y="{cy+20}" fill="#475569" font-size="10" font-family="JetBrains Mono">0%</text>
      <text x="{cx+R-22}" y="{cy+20}" fill="#475569" font-size="10" font-family="JetBrains Mono">100%</text>
      <text x="{cx}" y="{cy-20}" text-anchor="middle" fill="{col}" font-size="32" font-weight="800" font-family="JetBrains Mono">{pct:.1f}%</text>
      <text x="{cx}" y="{cy-2}" text-anchor="middle" fill="#64748B" font-size="11" font-family="Inter">Overall Coverage</text>
      <text x="{cx}" y="{cy+14}" text-anchor="middle" fill="#475569" font-size="10" font-family="JetBrains Mono">{alloc_sqm:,.1f} / {demand_sqm:,.1f} SQM</text>
    </svg>"""

# ── SVG horizontal bar chart ──────────────────────────────────────────────
def hbar_svg(data: list[tuple], label_w: int = 150) -> str:
    """data = [(label, value_0_100), ...]"""
    if not data: return "<div></div>"
    w, bH, gap = 480, 24, 7
    padL, padR = label_w, 60
    iW = w - padL - padR
    totalH = len(data)*(bH+gap) + 16
    rows = []
    for i, (label, val) in enumerate(data):
        y = i*(bH+gap) + 6
        col = fc(val)
        bW = max(2, (val/100) * iW)
        lbl = (label[:24] + "…") if len(label) > 25 else label
        rows.append(
            f'<g><text x="{padL-7}" y="{y+bH/2+4}" text-anchor="end" fill="#94A3B8" font-size="11" font-family="Inter">{lbl}</text>'
            f'<rect x="{padL}" y="{y}" width="{iW}" height="{bH}" rx="4" fill="#0F172A" opacity=".5"/>'
            f'<rect x="{padL}" y="{y}" width="{bW:.1f}" height="{bH}" rx="4" fill="{col}" opacity=".85"/>'
            f'<text x="{padL+bW+5:.1f}" y="{y+bH/2+4}" fill="#94A3B8" font-size="11" font-family="JetBrains Mono">{val:.1f}%</text></g>'
        )
    return f'<svg width="100%" height="{totalH}" viewBox="0 0 {w} {totalH}" style="display:block;">{"".join(rows)}</svg>'

# ── SVG vertical bars per location ─────────────────────────────────────────
def loc_vbar_svg(loc_stats: list[dict]) -> str:
    if not loc_stats: return ""
    w, h = 520, 195
    bars = []
    for p in (0, 25, 50, 75, 100):
        y = 175 - p*1.45
        bars.append(f'<line x1="70" y1="{y}" x2="500" y2="{y}" stroke="#334155" stroke-width=".5" stroke-dasharray="4,4"/>')
        bars.append(f'<text x="62" y="{y+4}" text-anchor="end" fill="#475569" font-size="10">{p}%</text>')
    for i, ls in enumerate(loc_stats):
        x = 95 + i*140
        bW = 52
        col = lc(ls["loc"])
        canH = (ls["cov"]/100)*145
        bars.append(f'<rect x="{x}" y="{175-canH:.1f}" width="{bW}" height="{canH:.1f}" rx="4" fill="{col}" opacity=".85"/>')
        bars.append(f'<text x="{x+bW/2}" y="{max(170-canH, 14):.1f}" text-anchor="middle" fill="{col}" font-size="11" font-weight="700">{ls["cov"]:.1f}%</text>')
        bars.append(f'<text x="{x+bW/2}" y="192" text-anchor="middle" fill="#94A3B8" font-size="10">{ls["loc"].split()[0]}</text>')
    return f'<svg width="100%" height="{h}" viewBox="0 0 {w} {h}" style="display:block;">{"".join(bars)}</svg>'

# ───────────────────────────────────────────────────────────────────────────────
# CASCADE ALLOCATION (preserved from v3)
# ───────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def _cached_cascade_allocate(tag_order_tuple: tuple) -> pd.DataFrame:
    pool = dict(INV_POOL_INIT)
    rows = []
    for tag in tag_order_tuple:
        tag_dm = dm[dm["Equipment_Tag_No."] == tag].copy()
        for code in sorted(tag_dm["Lining_System_Code"].unique(), key=lambda x: int(x)):
            code_rows = tag_dm[tag_dm["Lining_System_Code"] == code]
            for _, r in code_rows.iterrows():
                mat    = r["Material_Code"]
                demand = r["Demand_Qty"]
                before = pool.get(mat, 0.0)
                alloc  = min(demand, before)
                short  = demand - alloc
                after  = max(0.0, before - alloc)
                pool[mat] = after
                rows.append({
                    "Equipment_Tag_No.":         tag,
                    "Lining_System_Code":        code,
                    "Lining_System_Short_Name":  r["Lining_System_Short_Name"],
                    "Total_SQM":                 r["Total_SQM"],
                    "Material_Code":             mat,
                    "Material_Name":             r["Material_Name"],
                    "UOM":                       r["UOM"],
                    "Demand_Qty":                round(demand, 4),
                    "Allocated_Qty":             round(alloc, 4),
                    "Shortfall_Qty":             round(short, 4),
                    "Pool_Before":               round(before, 4),
                    "Pool_After":                round(after, 4),
                })
    result = pd.DataFrame(rows)
    if not result.empty:
        result["Fulfillment_Pct"] = (
            result["Allocated_Qty"] / result["Demand_Qty"].replace(0, np.nan) * 100
        ).fillna(100).clip(0, 100).round(2)
    return result

def cascade_allocate(tag_order: list[str]) -> pd.DataFrame:
    return _cached_cascade_allocate(tuple(tag_order))


# ───────────────────────────────────────────────────────────────────────────────
# SMART REORDERING SUGGESTIONS (ported from v3)
# ───────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def _run_suggestion_engine(tag_tuple: tuple) -> dict:
    """For each non-first tag, try every earlier position. Pick the one giving
       the best % gain. Threshold: > 0.4% absolute improvement."""
    tags = list(tag_tuple)
    if len(tags) < 2:
        return {"by_tag": [], "by_code": []}
    base_alloc = _cached_cascade_allocate(tag_tuple)
    base_dem  = base_alloc["Demand_Qty"].sum()
    base_act  = base_alloc["Allocated_Qty"].sum()
    base_pct  = (base_act/base_dem*100) if base_dem > 0 else 100

    by_tag = []
    by_code = []
    for i, tag in enumerate(tags):
        if i == 0: continue
        best = None
        for j in range(0, i):
            new_order = list(tags)
            new_order.insert(j, new_order.pop(i))
            ca = _cached_cascade_allocate(tuple(new_order))
            d = ca["Demand_Qty"].sum(); a = ca["Allocated_Qty"].sum()
            new_pct = (a/d*100) if d > 0 else 100
            gain = new_pct - base_pct
            if gain > 0.4 and (best is None or gain > best["gain"]):
                # Per-tag fulfillment improvement
                tag_old_pct = base_alloc[base_alloc["Equipment_Tag_No."]==tag]
                tag_new_pct = ca[ca["Equipment_Tag_No."]==tag]
                tdo = tag_old_pct["Demand_Qty"].sum(); tao = tag_old_pct["Allocated_Qty"].sum()
                tdn = tag_new_pct["Demand_Qty"].sum(); tan = tag_new_pct["Allocated_Qty"].sum()
                old_tp = (tao/tdo*100) if tdo > 0 else 100
                new_tp = (tan/tdn*100) if tdn > 0 else 100
                best = {"tag":tag, "from":i+1, "to":j+1, "gain":gain,
                         "old_pct":old_tp, "new_pct":new_tp}
        if best:
            by_tag.append(best)

    # By system-code suggestion — pick top per-code improvements
    base_by_code = base_alloc.groupby("Lining_System_Code").agg(
        d=("Demand_Qty","sum"), a=("Allocated_Qty","sum"),
        sname=("Lining_System_Short_Name","first")).reset_index()
    base_by_code["pct"] = (base_by_code["a"]/base_by_code["d"].replace(0,np.nan)*100).fillna(100)

    for bt in by_tag[:8]:
        new_order = list(tags)
        idx = new_order.index(bt["tag"])
        new_order.insert(bt["to"]-1, new_order.pop(idx))
        ca = _cached_cascade_allocate(tuple(new_order))
        ca_by_code = ca.groupby("Lining_System_Code").agg(
            d=("Demand_Qty","sum"), a=("Allocated_Qty","sum"),
            sname=("Lining_System_Short_Name","first")).reset_index()
        ca_by_code["pct"] = (ca_by_code["a"]/ca_by_code["d"].replace(0,np.nan)*100).fillna(100)
        merged = base_by_code.merge(ca_by_code[["Lining_System_Code","pct"]],
                                     on="Lining_System_Code", suffixes=("_old","_new"))
        merged["delta"] = merged["pct_new"] - merged["pct_old"]
        for _, mr in merged.sort_values("delta", ascending=False).head(2).iterrows():
            if mr["delta"] > 0.4:
                by_code.append({"code":mr["Lining_System_Code"], "sname":mr["sname"],
                                  "tag":bt["tag"], "from":bt["from"], "to":bt["to"],
                                  "old_pct":mr["pct_old"], "new_pct":mr["pct_new"],
                                  "gain":mr["delta"]})
    # Dedup by_code
    seen = set(); dedup = []
    for c in by_code:
        k = (c["code"], c["tag"])
        if k in seen: continue
        seen.add(k); dedup.append(c)
    return {"by_tag": by_tag[:5], "by_code": dedup[:8]}


def render_suggestion_panel(tag_list: list[str], panel_key: str) -> None:
    if len(tag_list) < 2:
        st.caption("Add at least 2 tags to surface reorder suggestions.")
        return
    with st.spinner("Analysing reorder scenarios…"):
        res = _run_suggestion_engine(tuple(tag_list))
    if not res["by_tag"] and not res["by_code"]:
        st.caption("✅ Current order is already optimal (no gain > 0.4%).")
        return
    st.caption("Suggestions that improve overall coverage by more than 0.4%. Apply by manually reordering.")
    c1, c2 = st.columns(2, gap="large")
    with c1:
        st.markdown('<div class="label-row">By Equipment</div>', unsafe_allow_html=True)
        for s in res["by_tag"]:
            reach_100 = s["new_pct"] >= 100 and s["old_pct"] < 100
            badge = ('<span style="background:rgba(16,185,129,.2);color:var(--green);'
                     'padding:2px 7px;border-radius:4px;font-size:10px;font-weight:700;'
                     'margin-left:6px;">→ 100%</span>') if reach_100 else ""
            st.markdown(f"""
            <div style="background:var(--bg-card-solid);border:1px solid var(--border-mid);
                        border-radius:8px;padding:11px 13px;margin-bottom:7px;">
              <div style="font-family:JetBrains Mono;font-size:12px;color:var(--amber);
                          font-weight:700;">{s["tag"]} {badge}</div>
              <div style="font-size:11px;color:var(--t2);margin-top:4px;">
                Move position {s["from"]} → {s["to"]}
              </div>
              <div style="display:flex;align-items:center;gap:8px;margin-top:6px;">
                <span style="font-size:11px;color:var(--t3);font-family:JetBrains Mono;">
                  {s["old_pct"]:.1f}% → {s["new_pct"]:.1f}%
                </span>
                <span style="margin-left:auto;background:rgba(245,158,11,.18);color:var(--amber);
                            padding:2px 7px;border-radius:4px;font-size:11px;font-weight:700;
                            font-family:JetBrains Mono;">+{s["gain"]:.1f}%</span>
              </div>
            </div>
            """, unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="label-row">By System Code</div>', unsafe_allow_html=True)
        for s in res["by_code"]:
            reach_100 = s["new_pct"] >= 100 and s["old_pct"] < 100
            badge = ('<span style="background:rgba(16,185,129,.2);color:var(--green);'
                     'padding:2px 7px;border-radius:4px;font-size:10px;font-weight:700;'
                     'margin-left:6px;">→ 100%</span>') if reach_100 else ""
            st.markdown(f"""
            <div style="background:var(--bg-card-solid);border:1px solid var(--border-mid);
                        border-radius:8px;padding:11px 13px;margin-bottom:7px;">
              <div style="display:flex;align-items:center;gap:8px;">
                {code_badge(s["code"])}
                <span style="font-size:12px;color:var(--t1);">{s["sname"]}</span>{badge}
              </div>
              <div style="font-size:11px;color:var(--t2);margin-top:5px;">
                via <span style="font-family:JetBrains Mono;color:var(--amber);">{s["tag"]}</span>
                : pos {s["from"]} → {s["to"]}
              </div>
              <div style="display:flex;align-items:center;gap:8px;margin-top:6px;">
                <span style="font-size:11px;color:var(--t3);font-family:JetBrains Mono;">
                  {s["old_pct"]:.1f}% → {s["new_pct"]:.1f}%
                </span>
                <span style="margin-left:auto;background:rgba(245,158,11,.18);color:var(--amber);
                            padding:2px 7px;border-radius:4px;font-size:11px;font-weight:700;
                            font-family:JetBrains Mono;">+{s["gain"]:.1f}%</span>
              </div>
            </div>
            """, unsafe_allow_html=True)

def alloc_to_mat_df(alloc: pd.DataFrame, columns: list[str] = None) -> pd.DataFrame:
    """Aggregate cascade allocation to material level for table display."""
    if alloc.empty: return pd.DataFrame()
    agg = alloc.groupby(["Material_Code","Material_Name","UOM"], as_index=False).agg(
        Demand=("Demand_Qty","sum"),
        Allocated=("Allocated_Qty","sum"),
        Shortfall=("Shortfall_Qty","sum"),
    )
    agg["Available"] = agg["Material_Code"].map(INV_POOL_INIT).fillna(0)
    agg["On Order"] = agg["Material_Code"].map(INV_ORDERED_INIT).fillna(0)
    agg["Net Shortfall"] = (agg["Shortfall"] - agg["On Order"]).clip(lower=0)
    agg["Fulfil %"] = (agg["Allocated"]/agg["Demand"].replace(0, np.nan)*100).fillna(100).clip(0,100)
    agg = agg.rename(columns={"Material_Code":"Code","Material_Name":"Material Name"})
    return agg

def sc_alloc_to_mat_df(alloc: pd.DataFrame) -> pd.DataFrame:
    """Per-row mat table (no aggregation, retains system-code context)."""
    if alloc.empty: return pd.DataFrame()
    out = alloc.copy()
    out["Available"] = out["Material_Code"].map(INV_POOL_INIT).fillna(0)
    out["Fulfil %"] = out["Fulfillment_Pct"]
    out = out.rename(columns={"Material_Code":"Code","Material_Name":"Material Name",
                              "Demand_Qty":"Demand","Allocated_Qty":"Allocated",
                              "Shortfall_Qty":"Shortfall"})
    return out

# ───────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT (simplified vs v3 — keeps headers, color schemes)
# ───────────────────────────────────────────────────────────────────────────────
EXCEL_SCHEMES = {
    "dashboard":   {"title":"#1A2A3A", "head":"#2D4A6A", "total":"#F0C040"},
    "brown_field": {"title":"#0F2D52", "head":"#1E5799", "total":"#BDD7F0"},
    "train_j":     {"title":"#4A2E00", "head":"#A0620A", "total":"#FDE8A0"},
    "train_k":     {"title":"#0A2E1A", "head":"#1A6B48", "total":"#B3F0D8"},
    "session":     {"title":"#2D1A52", "head":"#5B2D8E", "total":"#E5D0F0"},
    "execution":   {"title":"#3A0A0A", "head":"#8E1A1A", "total":"#F5C6C6"},
    "overview":    {"title":"#0A2A2A", "head":"#0E7490", "total":"#A5F3FC"},
}

def excel_bytes_multi(sheets: list[dict]) -> bytes:
    """Multi-sheet workbook. Each sheet dict: {name, df, title, color_scheme}."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        for s in sheets:
            scheme = EXCEL_SCHEMES.get(s.get("color_scheme","dashboard"), EXCEL_SCHEMES["dashboard"])
            df    = s["df"]
            name  = (s.get("name") or "Sheet")[:31]
            title = s.get("title", name)
            df.to_excel(writer, index=False, sheet_name=name, startrow=5, header=False)
            ws = writer.sheets[name]
            title_fmt = wb.add_format({"bold":True,"font_size":13,"bg_color":scheme["title"],
                                       "font_color":"#FFFFFF","align":"center","valign":"vcenter",
                                       "border":1})
            hdr_fmt   = wb.add_format({"bold":True,"font_size":10,"bg_color":scheme["head"],
                                       "font_color":"#FFFFFF","align":"center","valign":"vcenter",
                                       "border":1})
            total_fmt = wb.add_format({"bold":True,"bg_color":scheme["total"],"border":1,"font_size":10})
            if os.path.exists(LOGO_PATH):
                try:
                    from PIL import Image as _PIL
                    _img = _PIL.open(LOGO_PATH)
                    _w, _h = _img.size
                    scale = 70 / _h
                    ws.insert_image("A1", LOGO_PATH,
                        {"x_scale": scale, "y_scale": scale,
                         "x_offset": 4, "y_offset": 4, "object_position": 3})
                except Exception: pass
            ncols = max(1, len(df.columns))
            ws.merge_range(4, 0, 4, ncols-1, title, title_fmt)
            for i, col in enumerate(df.columns):
                ws.write(5, i, col, hdr_fmt)
                maxlen = max([len(str(col))] + [len(str(v)) for v in df[col].astype(str).head(60)])
                ws.set_column(i, i, min(42, maxlen + 2))
            if s.get("add_grand_total", True) and len(df) > 0:
                tr = 6 + len(df)
                numeric = df.select_dtypes(include=[np.number]).columns
                for i, col in enumerate(df.columns):
                    if col in numeric:
                        ws.write(tr, i, df[col].sum(), total_fmt)
                    else:
                        ws.write(tr, i, "TOTAL" if i == 0 else "", total_fmt)
            ws.autofilter(5, 0, 5 + len(df), ncols-1)
    return buf.getvalue()


def shortfall_chart_svg(chart_data: list[dict], accent_col: str = "#F59E0B") -> str:
    """Horizontal stacked bar — Allocated (accent) + Shortfall (red) per (tag, code)."""
    if not chart_data: return ""
    w = 720; bH = 26; gap = 7
    padL, padR = 220, 60
    iW = w - padL - padR
    totalH = len(chart_data)*(bH+gap) + 22
    max_d = max(d["Demand"] for d in chart_data) or 1
    rows = []
    for i, d in enumerate(chart_data):
        y = i*(bH+gap) + 14
        alloc_w = (d["Allocated"]/max_d)*iW
        short_w = (d["Shortfall"]/max_d)*iW
        lbl = d["Label"][:36]
        rows.append(
            f'<text x="{padL-7}" y="{y+bH/2+4}" text-anchor="end" fill="#94A3B8" font-size="10" font-family="Inter">{lbl}</text>'
            f'<rect x="{padL}" y="{y}" width="{iW}" height="{bH}" rx="3" fill="#0F172A" opacity=".4"/>'
            f'<rect x="{padL}" y="{y}" width="{alloc_w:.1f}" height="{bH}" rx="3" fill="{accent_col}" opacity=".85"/>'
            f'<rect x="{padL+alloc_w:.1f}" y="{y}" width="{short_w:.1f}" height="{bH}" rx="3" fill="#EF4444" opacity=".85"/>'
            + (f'<text x="{padL+alloc_w/2:.1f}" y="{y+bH/2+4}" text-anchor="middle" fill="#fff" font-size="9" font-family="JetBrains Mono">{d["Allocated"]:,.0f}</text>' if alloc_w > 30 else '')
            + (f'<text x="{padL+alloc_w+short_w/2:.1f}" y="{y+bH/2+4}" text-anchor="middle" fill="#fff" font-size="9" font-family="JetBrains Mono">{d["Shortfall"]:,.0f}</text>' if short_w > 30 else '')
        )
    # Legend
    rows.append(
        f'<rect x="{padL}" y="0" width="10" height="10" fill="{accent_col}" opacity=".85"/>'
        f'<text x="{padL+14}" y="9" fill="#94A3B8" font-size="10" font-family="Inter">Allocated</text>'
        f'<rect x="{padL+80}" y="0" width="10" height="10" fill="#EF4444" opacity=".85"/>'
        f'<text x="{padL+94}" y="9" fill="#94A3B8" font-size="10" font-family="Inter">Shortfall</text>'
    )
    return f'<svg width="100%" height="{totalH}" viewBox="0 0 {w} {totalH}" style="display:block;">{"".join(rows)}</svg>'


def excel_bytes(df: pd.DataFrame, title: str, color_scheme: str = "dashboard",
                add_grand_total: bool = True) -> bytes:
    scheme = EXCEL_SCHEMES.get(color_scheme, EXCEL_SCHEMES["dashboard"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Report", startrow=5, header=False)
        wb  = writer.book
        ws  = writer.sheets["Report"]
        title_fmt  = wb.add_format({"bold":True,"font_size":13,"bg_color":scheme["title"],
                                    "font_color":"#FFFFFF","align":"center","valign":"vcenter",
                                    "border":1})
        hdr_fmt    = wb.add_format({"bold":True,"font_size":10,"bg_color":scheme["head"],
                                    "font_color":"#FFFFFF","align":"center","valign":"vcenter",
                                    "border":1})
        meta_fmt   = wb.add_format({"font_size":9,"italic":True,"font_color":"#666666","align":"right"})
        total_fmt  = wb.add_format({"bold":True,"bg_color":scheme["total"],"border":1,"font_size":10})
        ws.set_row(0, 18); ws.set_row(1, 18); ws.set_row(2, 18); ws.set_row(3, 18)
        # Resize logo down to a small badge (~120x80 px) so it doesn't span the page
        if os.path.exists(LOGO_PATH):
            try:
                from PIL import Image as _PIL
                _img = _PIL.open(LOGO_PATH)
                _w, _h = _img.size
                target_h = 70
                scale = target_h / _h
                ws.insert_image("A1", LOGO_PATH,
                    {"x_scale": scale, "y_scale": scale,
                     "x_offset": 4, "y_offset": 4,
                     "object_position": 3})  # 3 = don't move or size with cells
            except Exception:
                pass
        ncols = max(1, len(df.columns))
        ws.merge_range(4, 0, 4, ncols-1, title, title_fmt)
        ws.write(0, ncols-1, f"Generated: {datetime.now():%Y-%m-%d %H:%M}", meta_fmt)
        ws.write(1, ncols-1, "Smart Material Estimator v4", meta_fmt)
        for i, col in enumerate(df.columns):
            ws.write(5, i, col, hdr_fmt)
            maxlen = max([len(str(col))] + [len(str(v)) for v in df[col].astype(str).head(60)])
            ws.set_column(i, i, min(42, maxlen + 2))
        if add_grand_total and len(df) > 0:
            tr = 6 + len(df)
            numeric = df.select_dtypes(include=[np.number]).columns
            for i, col in enumerate(df.columns):
                if col in numeric:
                    ws.write(tr, i, df[col].sum(), total_fmt)
                else:
                    ws.write(tr, i, "TOTAL" if i == 0 else "", total_fmt)
        ws.autofilter(5, 0, 5 + len(df), ncols-1)
    return buf.getvalue()

# ───────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ───────────────────────────────────────────────────────────────────────────────
def render_sidebar():
    with st.sidebar:
        logo = _logo_b64()
        if logo:
            st.markdown(f'<div style="text-align:center;padding:14px 0 6px 0;">'
                        f'<img src="data:image/png;base64,{logo}" style="height:54px;"></div>',
                        unsafe_allow_html=True)
        st.markdown("""
        <div style="padding:6px 4px 18px 4px;">
          <div style="font-size:24px;font-weight:900;color:var(--amber);
                      font-family:'JetBrains Mono',monospace;line-height:1.1;">🏗 SME</div>
          <div style="font-size:9px;color:var(--t4);text-transform:uppercase;
                      letter-spacing:1.8px;margin-top:4px;">Smart Material Estimator v4</div>
        </div>
        """, unsafe_allow_html=True)

        sec_header("📍 Project Overview")
        for loc in LOCATION_ORDER:
            n = (eq_master["Location"] == loc).sum()
            st.markdown(
                f'<div style="display:flex;align-items:center;justify-content:space-between;padding:4px 0;">'
                f'{loc_badge(loc)}'
                f'<span style="font-size:11px;color:var(--t3);font-family:JetBrains Mono;">{n} equip.</span></div>',
                unsafe_allow_html=True
            )

        st.markdown("<hr>", unsafe_allow_html=True)
        sec_header("📦 Inventory")
        zero_n = (inv["Available_Qty"] == 0).sum()
        st.markdown(
            f'<div style="font-size:11px;color:var(--t3);">'
            f'📦 {len(inv)} materials &nbsp;·&nbsp; ⚠️ {zero_n} at zero stock</div>',
            unsafe_allow_html=True
        )

        st.markdown("<hr>", unsafe_allow_html=True)
        sec_header(f"📋 Session ({len(st.session_state.session_tags)} tags)")
        if st.session_state.session_tags:
            for t in st.session_state.session_tags:
                st.markdown(
                    f'<div style="padding:3px 0;font-size:11px;color:var(--t2);'
                    f'font-family:JetBrains Mono;">·  {t}</div>',
                    unsafe_allow_html=True
                )
            if st.button("🗑 Clear Session", key="clear_sidebar", use_container_width=True):
                st.session_state.session_tags = []
                st.rerun()
        else:
            st.caption("No equipment added yet.")

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown(
            '<div style="font-size:10px;color:var(--t4);line-height:2;">'
            '🟢 100%  &nbsp; 🟠 90–99%  &nbsp; 🟡 80–89%  &nbsp; 🔴 &lt;80%</div>',
            unsafe_allow_html=True
        )
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("↩  Logout", use_container_width=True, key="_logout_btn"):
            st.session_state["_authenticated"] = False
            st.rerun()

# ───────────────────────────────────────────────────────────────────────────────
# STICKY HEADER
# ───────────────────────────────────────────────────────────────────────────────
def render_header():
    logo = _logo_b64()
    logo_html = (f'<img src="data:image/png;base64,{logo}" style="height:34px;">'
                 if logo else '<span style="font-size:24px;">🏗</span>')
    st.markdown(f"""
    <div class="sme-header-sticky">
      <div style="display:flex;align-items:center;gap:14px;padding:8px 4px 10px 4px;">
        {logo_html}
        <div>
          <div style="font-size:15px;font-weight:700;color:var(--t0);letter-spacing:-.2px;">
            Smart Material Estimator &amp; Planner
          </div>
          <div style="font-size:10px;color:var(--t4);letter-spacing:.5px;">
            System-code level · Cascading allocation · Priority-based
          </div>
        </div>
        <div style="flex:1;"></div>
        <span class="pill" style="background:var(--amber-bg);color:var(--amber);
              border-color:rgba(245,158,11,.3);">v4</span>
        <span style="display:flex;align-items:center;gap:5px;">
          <span class="live-dot"></span>
          <span style="font-size:11px;color:var(--green);">Online</span>
        </span>
      </div>
    </div>
    """, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════
def tab_dashboard():
    # ── Sub-view toggle (sticky under tab bar) ─────────
    view = submode_radio("dash_view",
                         ["📈 Project Overview", "🛒 Material Requirement & Procurement"],
                         key="dash_view")

    # ── Cascading Filters via popovers: Type → Location → Substrate → System Code ─────
    sec_header("🎛 Filters")
    fc1, fc2, fc3, fc4 = st.columns(4)
    em = eq_master.copy()

    with fc1:
        type_opts = sorted(em["Type"].dropna().unique())
        f_type = popover_multiselect("Type", type_opts, type_opts, key="dash_type")
    em = em[em["Type"].isin(f_type)] if f_type else em.iloc[0:0]

    with fc2:
        loc_opts = [l for l in LOCATION_ORDER if l in em["Location"].unique()]
        f_loc = popover_multiselect("Location", loc_opts, loc_opts, key="dash_loc")
    em = em[em["Location"].isin(f_loc)] if f_loc else em.iloc[0:0]

    with fc3:
        sub_opts = sorted(em["Substrate"].dropna().unique())
        f_sub = popover_multiselect("Substrate", sub_opts, sub_opts, key="dash_sub")
    if sub_opts:
        em = em[em["Substrate"].isin(f_sub) | em["Substrate"].isna()] if f_sub else em.iloc[0:0]

    # System Codes — narrow to those owned by surviving tags
    surviving_tags = em["Equipment_Tag_No."].tolist()
    sc_avail = equip_sc[equip_sc["Equipment_Tag_No."].isin(surviving_tags)]
    with fc4:
        code_pairs = (sc_avail[["Lining_System_Code","Lining_System_Short_Name"]]
                      .drop_duplicates().sort_values("Lining_System_Code", key=lambda s: s.astype(int)))
        code_opts = [f'Code {r["Lining_System_Code"]} – {r["Lining_System_Short_Name"]}'
                     for _, r in code_pairs.iterrows()]
        f_code_lbls = popover_multiselect("System Code", code_opts, code_opts, key="dash_code")
        f_codes = [s.split(" ")[1] for s in f_code_lbls]

    sc_filtered = sc_avail[sc_avail["Lining_System_Code"].isin(f_codes)]
    em = em[em["Equipment_Tag_No."].isin(sc_filtered["Equipment_Tag_No."])]
    tags = em["Equipment_Tag_No."].tolist()

    if not tags:
        st.warning("No equipment matches the current filters.")
        return

    alloc = cascade_allocate(tags)

    # ── Compute KPIs ─────────
    total_sqm    = sc_filtered[sc_filtered["Equipment_Tag_No."].isin(tags)]["Total_SQM"].sum()
    total_demand = alloc["Demand_Qty"].sum() if not alloc.empty else 0
    total_alloc  = alloc["Allocated_Qty"].sum() if not alloc.empty else 0
    total_short  = alloc["Shortfall_Qty"].sum() if not alloc.empty else 0
    cov_pct      = (total_alloc/total_demand*100) if total_demand > 0 else 100
    can_sqm      = total_sqm * cov_pct / 100
    sqm_deficit  = max(0, total_sqm - can_sqm)
    mat_agg      = alloc_to_mat_df(alloc) if not alloc.empty else pd.DataFrame()
    crit_n       = ((mat_agg["Fulfil %"] < 50).sum()) if not mat_agg.empty else 0

    # ────────── SUB-VIEW A: PROJECT OVERVIEW ──────────
    if view.startswith("📈"):
        kpi_strip([
            ("Equipment",            f"{len(tags)} tags",          f"{len(tags)} total"),
            ("Total SQM",            f"{total_sqm:,.1f}",          "remaining"),
            ("Available Coverage",   f"{can_sqm:,.1f}",            "SQM backed"),
            ("SQM Deficit",          f"{sqm_deficit:,.1f}",        "uncoverable",        "#EF4444"),
            ("Overall Coverage",     f"{cov_pct:.1f}%",            "alloc/demand",       fc(cov_pct)),
            ("Critical (<50%)",      f"{crit_n}",                  "materials",
                                     "#EF4444" if crit_n>0 else "#10B981"),
        ], cols=6)

        # ── Row 1: Gauge + Location bars ─────────
        c1, c2 = st.columns([1, 1.6], gap="large")
        with c1:
            st.markdown('<div class="glass">', unsafe_allow_html=True)
            sec_header("🎯 Overall Coverage")
            st.markdown(gauge_svg(cov_pct, total_sqm, can_sqm), unsafe_allow_html=True)
            # availability bar
            avail_w = max(0, min(100, cov_pct))
            short_w = 100 - avail_w
            st.markdown(f"""
            <div style="margin-top:14px;height:22px;border-radius:6px;overflow:hidden;display:flex;">
              <div style="width:{avail_w}%;background:#10B981;display:flex;align-items:center;
                          justify-content:flex-end;padding-right:7px;">
                <span style="font-size:10px;color:#fff;font-family:JetBrains Mono;font-weight:700;">{can_sqm:,.0f}</span>
              </div>
              <div style="width:{short_w}%;background:rgba(239,68,68,.75);display:flex;
                          align-items:center;padding-left:6px;">
                <span style="font-size:10px;color:#fff;font-family:JetBrains Mono;font-weight:700;">{sqm_deficit:,.0f}</span>
              </div>
            </div>
            <div style="display:flex;gap:18px;margin-top:6px;font-size:10px;color:var(--t4);">
              <span>■ Available: {can_sqm:,.0f} SQM</span>
              <span>■ Shortfall: {sqm_deficit:,.0f} SQM</span>
            </div>
            </div>
            """, unsafe_allow_html=True)

        with c2:
            st.markdown('<div class="glass">', unsafe_allow_html=True)
            sec_header("📍 Coverage by Location (SQM)")
            loc_stats = []
            for loc in LOCATION_ORDER:
                ltags = em[em["Location"] == loc]["Equipment_Tag_No."].tolist()
                if not ltags: continue
                lsqm = sc_filtered[sc_filtered["Equipment_Tag_No."].isin(ltags)]["Total_SQM"].sum()
                lalloc = cascade_allocate(ltags)
                ldem = lalloc["Demand_Qty"].sum() if not lalloc.empty else 0
                lact = lalloc["Allocated_Qty"].sum() if not lalloc.empty else 0
                lcov = (lact/ldem*100) if ldem > 0 else 100
                lcan = lsqm * lcov / 100
                loc_stats.append({"loc":loc, "sqm":lsqm, "can":lcan, "cov":lcov, "count":len(ltags)})
            st.markdown(loc_vbar_svg(loc_stats), unsafe_allow_html=True)
            # Per-location stat tiles
            cards = '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:9px;margin-top:10px;">'
            for ls in loc_stats:
                cards += (f'<div style="background:var(--bg-inset);border-radius:8px;padding:10px 12px;'
                          f'border:1px solid {lc(ls["loc"])}33;">'
                          f'<div style="font-size:11px;color:{lc(ls["loc"])};font-weight:700;margin-bottom:4px;">{ls["loc"]}</div>'
                          f'<div style="font-size:20px;font-weight:900;color:{fc(ls["cov"])};'
                          f'font-family:JetBrains Mono;">{ls["cov"]:.1f}%</div>'
                          f'<div style="font-size:10px;color:var(--t3);margin-top:2px;">'
                          f'{ls["can"]:,.0f} / {ls["sqm"]:,.0f} SQM · {ls["count"]} eq.</div></div>')
            cards += "</div></div>"
            st.markdown(cards, unsafe_allow_html=True)

        # ── Row 2: Coverage by System Code + Coverage by Material ─────────
        c1, c2 = st.columns(2, gap="large")
        with c1:
            st.markdown('<div class="glass">', unsafe_allow_html=True)
            sec_header("🔬 Coverage by System Code (SQM)")
            sc_data = []
            for code in sorted(set(f_codes), key=lambda x: int(x)):
                sub = alloc[alloc["Lining_System_Code"] == code]
                if sub.empty: continue
                d = sub["Demand_Qty"].sum(); a = sub["Allocated_Qty"].sum()
                pct = (a/d*100) if d > 0 else 100
                sname = sub["Lining_System_Short_Name"].iloc[0]
                sc_data.append((f"C{code} – {sname}", pct))
            st.markdown(hbar_svg(sc_data, label_w=140), unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with c2:
            st.markdown('<div class="glass">', unsafe_allow_html=True)
            sec_header("🧪 Coverage by Material")
            mat_data = []
            if not mat_agg.empty:
                for _, r in mat_agg.sort_values("Fulfil %").head(10).iterrows():
                    nm = r["Material Name"][:18]
                    mat_data.append((f'{r["Code"]}  {nm}', float(r["Fulfil %"])))
            st.markdown(hbar_svg(mat_data, label_w=180), unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        # ── Full Material Balance ─────────
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        sec_header("📋 Full Material Balance")
        if not mat_agg.empty:
            mat_show = mat_agg.sort_values("Fulfil %").copy()
            st.markdown(mat_table_html(mat_show, columns=[
                "Code","Material Name","UOM","Available","On Order",
                "Demand","Shortfall","Net Shortfall","Fulfil %"]),
                unsafe_allow_html=True)
            st.download_button(
                "⬇ Download Material Balance",
                excel_bytes(mat_show, "Full Material Balance", "dashboard"),
                file_name=f"material_balance_{date.today():%Y%m%d}.xlsx", key="dash_dl_a",
            )
        st.markdown("</div>", unsafe_allow_html=True)

    # ────────── SUB-VIEW B: MATERIAL REQUIREMENT & PROCUREMENT ──────────
    else:
        kpi_strip([
            ("Equipment",       f"{len(tags)}",              "tags"),
            ("Total SQM",       f"{total_sqm:,.1f}",         "remaining"),
            ("Available Cov.",  f"{can_sqm:,.1f}",           "SQM backed"),
            ("SQM Deficit",     f"{sqm_deficit:,.1f}",       "uncoverable",   "#EF4444"),
        ], cols=4)

        # Per location
        for loc in LOCATION_ORDER:
            ltags = em[em["Location"] == loc]["Equipment_Tag_No."].tolist()
            if not ltags: continue
            lalloc = cascade_allocate(ltags)
            lsqm = sc_filtered[sc_filtered["Equipment_Tag_No."].isin(ltags)]["Total_SQM"].sum()
            ldem = lalloc["Demand_Qty"].sum() if not lalloc.empty else 0
            lact = lalloc["Allocated_Qty"].sum() if not lalloc.empty else 0
            lcov = (lact/ldem*100) if ldem > 0 else 100
            lcan = lsqm * lcov / 100
            color = lc(loc)
            st.markdown(f"""
            <div style="display:flex;align-items:center;gap:10px;margin:18px 0 10px 0;
                        padding:9px 14px;background:{color}11;
                        border:1px solid {color}33;border-radius:8px;">
              <span class="dot" style="background:{color};"></span>
              {loc_badge(loc)}
              <span style="font-size:12px;color:var(--t2);">{len(ltags)} equip · {lcan:,.0f}/{lsqm:,.0f} SQM</span>
              {fulfil_pill(lcov)}
            </div>
            """, unsafe_allow_html=True)

            # Per system code expander
            for code in sorted(lalloc["Lining_System_Code"].unique(), key=lambda x: int(x)):
                sub = lalloc[lalloc["Lining_System_Code"] == code]
                if sub.empty: continue
                sname = sub["Lining_System_Short_Name"].iloc[0]
                csqm = sub["Total_SQM"].drop_duplicates().sum() if "Total_SQM" in sub.columns else 0
                cdem = sub["Demand_Qty"].sum(); cact = sub["Allocated_Qty"].sum()
                ccov = (cact/cdem*100) if cdem > 0 else 100
                ccan = csqm * ccov / 100
                with st.expander(f"Code {code} – {sname}  ·  {ccan:,.0f}/{csqm:,.0f} SQM  ·  {ccov:.1f}%", expanded=False):
                    # 5-card metric strip
                    m1, m2, m3, m4, m5 = st.columns(5)
                    m1.metric("System Code", f"Code {code}")
                    m2.metric("Short Name", sname)
                    m3.metric("SQM Total", f"{csqm:,.2f}")
                    m4.metric("Coverage SQM", f"{ccan:,.2f}")
                    m5.metric("SQM Deficit", f"{csqm-ccan:,.2f}")
                    sub_mat = alloc_to_mat_df(sub)
                    st.markdown(mat_table_html(sub_mat, columns=[
                        "Code","Material Name","UOM","Available","On Order",
                        "Demand","Shortfall","Net Shortfall","Fulfil %"]),
                        unsafe_allow_html=True)

        # Grand total
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        sec_header("📦 Grand Total — All Selected Equipment")
        if not mat_agg.empty:
            st.markdown(mat_table_html(
                mat_agg.sort_values("Fulfil %"),
                columns=["Code","Material Name","UOM","Available","On Order",
                         "Demand","Shortfall","Net Shortfall","Fulfil %"]),
                unsafe_allow_html=True)
            d1, d2 = st.columns(2)
            with d1:
                st.download_button(
                    "⬇ Grand Procurement Table",
                    excel_bytes(mat_agg.sort_values("Fulfil %"),
                                "Grand Procurement Table", "dashboard"),
                    file_name=f"procurement_grand_{date.today():%Y%m%d}.xlsx", key="dash_dl_b1",
                )
            with d2:
                net = mat_agg[mat_agg["Net Shortfall"] > 0]
                if not net.empty:
                    st.download_button(
                        "⬇ Net Order List Only",
                        excel_bytes(net, "Net Order List", "execution"),
                        file_name=f"net_order_list_{date.today():%Y%m%d}.xlsx", key="dash_dl_b2",
                    )
        st.markdown("</div>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — ENTRY
# ═══════════════════════════════════════════════════════════════════════════════
def tab_entry():
    left, right = st.columns([1, 1.65], gap="large")

    # ── Left: cascading filters + multi-select + session list ─────────
    with left:
        st.markdown('<div class="glass-flat">', unsafe_allow_html=True)
        sec_header("🎛 Filter Equipment")
        em = eq_master.copy()

        # Cascading: Type → Location → Substrate → System Code (same as Dashboard)
        type_opts = sorted(em["Type"].dropna().unique())
        type_f = popover_multiselect("Type", type_opts, type_opts, key="t1_type")
        em = em[em["Type"].isin(type_f)] if type_f else em.iloc[0:0]

        loc_opts = [l for l in LOCATION_ORDER if l in em["Location"].unique()]
        loc_f = popover_multiselect("Location", loc_opts, loc_opts, key="t1_loc")
        em = em[em["Location"].isin(loc_f)] if loc_f else em.iloc[0:0]

        sub_opts = sorted(em["Substrate"].dropna().unique())
        sub_f = popover_multiselect("Substrate", sub_opts, sub_opts, key="t1_sub")
        if sub_opts:
            em = em[em["Substrate"].isin(sub_f) | em["Substrate"].isna()] if sub_f else em.iloc[0:0]

        # System Codes narrow to surviving tags
        surviving = em["Equipment_Tag_No."].tolist()
        sc_avail = equip_sc[equip_sc["Equipment_Tag_No."].isin(surviving)]
        code_pairs = (sc_avail[["Lining_System_Code","Lining_System_Short_Name"]]
                      .drop_duplicates().sort_values("Lining_System_Code",
                                                       key=lambda s: s.astype(int)))
        code_opts = [f'Code {r["Lining_System_Code"]} – {r["Lining_System_Short_Name"]}'
                     for _, r in code_pairs.iterrows()]
        code_f_lbls = popover_multiselect("System Code", code_opts, code_opts, key="t1_code")
        codes = [s.split(" ")[1] for s in code_f_lbls]
        sc_filt = sc_avail[sc_avail["Lining_System_Code"].isin(codes)]
        em = em[em["Equipment_Tag_No."].isin(sc_filt["Equipment_Tag_No."])]
        filtered_tags = em["Equipment_Tag_No."].tolist()
        st.markdown(
            f'<div style="font-size:11px;color:var(--t3);margin-top:6px;">'
            f'{len(filtered_tags)} equipment match these filters.</div>',
            unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Multi-equipment picker ─────────
        st.markdown('<div class="glass-flat">', unsafe_allow_html=True)
        sec_header("🔍 Select Equipment to Add")
        not_in_session = [t for t in filtered_tags if t not in st.session_state.session_tags]
        in_session_count = sum(1 for t in filtered_tags if t in st.session_state.session_tags)
        if not_in_session:
            picked = st.multiselect(
                "Pick tags to add",
                options=not_in_session,
                format_func=lambda t: f'{t}  —  {eq_master[eq_master["Equipment_Tag_No."]==t]["Name"].iloc[0]}',
                key="t1_picker", label_visibility="collapsed",
                placeholder=f"Pick from {len(not_in_session)} available tags…",
            )
            ac1, ac2 = st.columns([2, 1])
            with ac1:
                if st.button(
                    f"＋ Add {len(picked)} Selected to Session" if picked
                    else "＋ Add Selected to Session",
                    disabled=not picked, key="add_picked", use_container_width=True):
                    for t in picked:
                        if t not in st.session_state.session_tags:
                            st.session_state.session_tags.append(t)
                    st.rerun()
            with ac2:
                if st.button(f"＋ ALL {len(not_in_session)}", key="add_all_filtered",
                              use_container_width=True,
                              help="Add every tag matching the current filters"):
                    for t in not_in_session:
                        st.session_state.session_tags.append(t)
                    st.rerun()
        else:
            if filtered_tags:
                st.success(f"✓ All {len(filtered_tags)} filtered equipment already in session.")
            else:
                st.info("No equipment matches the current filters.")
        if in_session_count:
            st.caption(f"✓ {in_session_count} of these tags already in session.")
        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # For per-tag detail on the right pane: pick one to inspect
        tag_sel = ""
        if filtered_tags:
            tag_sel = st.selectbox(
                "Inspect equipment",
                options=[""] + filtered_tags,
                format_func=lambda t: "— Click to inspect detail —" if t == "" else
                                       f'{t}  —  {eq_master[eq_master["Equipment_Tag_No."]==t]["Name"].iloc[0]}',
                key="tag_select", label_visibility="collapsed",
            )

        st.markdown('<div class="glass-flat">', unsafe_allow_html=True)
        sec_header(f"📋 Session Priority List ({len(st.session_state.session_tags)})")
        if not st.session_state.session_tags:
            st.info("Add equipment tags above to build your session.")
        else:
            st.caption("⬆ ⬇ Reorder priority — order is applied instantly.")
            for idx, tag in enumerate(list(st.session_state.session_tags)):
                row = eq_master[eq_master["Equipment_Tag_No."] == tag]
                if row.empty: continue
                row = row.iloc[0]
                # Per-tag cascading alloc fragment (not session-aware here; quick view)
                tag_alloc = cascade_allocate([tag])
                d = tag_alloc["Demand_Qty"].sum() if not tag_alloc.empty else 0
                a = tag_alloc["Allocated_Qty"].sum() if not tag_alloc.empty else 0
                pct = (a/d*100) if d > 0 else 100
                codes = sorted(equip_sc[equip_sc["Equipment_Tag_No."] == tag]["Lining_System_Code"].unique(),
                               key=lambda x: int(x))
                code_chips = "".join(code_badge(c) + " " for c in codes)
                st.markdown(f"""
                <div style="display:flex;align-items:center;gap:9px;padding:9px 12px;
                            background:var(--bg-inset);border-radius:8px;margin-bottom:6px;
                            border:1px solid var(--border-mid);">
                  <div style="width:22px;height:22px;border-radius:50%;background:rgba(245,158,11,.2);
                              color:var(--amber);font-size:10px;font-weight:800;display:flex;
                              align-items:center;justify-content:center;flex-shrink:0;">{idx+1}</div>
                  <div style="flex:1;min-width:0;">
                    <div style="display:flex;align-items:center;gap:6px;margin-bottom:3px;">
                      <span style="font-size:12px;font-weight:800;color:var(--amber);
                                   font-family:JetBrains Mono;">{tag}</span>
                      {loc_badge(row["Location"])}
                    </div>
                    <div style="font-size:11px;color:var(--t2);overflow:hidden;
                                text-overflow:ellipsis;white-space:nowrap;">{row["Name"]}</div>
                    <div style="display:flex;gap:4px;margin-top:4px;align-items:center;">
                      {code_chips}{fulfil_pill(pct)}
                    </div>
                  </div>
                </div>
                """, unsafe_allow_html=True)
                bc1, bc2, bc3 = st.columns([1, 1, 4])
                with bc1:
                    if idx > 0 and st.button("↑", key=f"up_{tag}_{idx}", use_container_width=True):
                        s = st.session_state.session_tags
                        s[idx-1], s[idx] = s[idx], s[idx-1]
                        st.rerun()
                with bc2:
                    if idx < len(st.session_state.session_tags)-1 and st.button("↓", key=f"dn_{tag}_{idx}", use_container_width=True):
                        s = st.session_state.session_tags
                        s[idx+1], s[idx] = s[idx], s[idx+1]
                        st.rerun()
                with bc3:
                    if st.button("✕  Remove", key=f"rm_{tag}_{idx}", use_container_width=True):
                        st.session_state.session_tags.remove(tag)
                        st.rerun()
            if st.button("🗑 Clear All", key="clear_all", use_container_width=True):
                st.session_state.session_tags = []
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    # ── Right: equipment detail ─────────
    with right:
        if not tag_sel:
            st.markdown('<div style="display:flex;align-items:center;justify-content:center;'
                        'height:400px;color:var(--t4);font-size:14px;text-transform:uppercase;'
                        'letter-spacing:2px;">Select an Equipment Tag to View Details</div>',
                        unsafe_allow_html=True)
            return
        row = eq_master[eq_master["Equipment_Tag_No."] == tag_sel].iloc[0]
        # Hero card
        st.markdown(f"""
        <div class="glass glass-amber-bar" style="margin-bottom:16px;">
          <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;">
            <div>
              <div class="label-row">Equipment Tag</div>
              <div style="font-size:28px;font-weight:900;color:var(--amber);
                          font-family:JetBrains Mono;">{tag_sel}</div>
            </div>
            {loc_badge(row["Location"])}
          </div>
          <div style="font-size:17px;font-weight:600;color:var(--t0);margin-bottom:14px;">{row["Name"]}</div>
          <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:9px;">
            <div style="background:var(--bg-inset);border-radius:8px;padding:9px 11px;">
              <div class="label-row">Type</div>
              <div style="font-size:13px;color:var(--t1);font-weight:500;">{row.get("Type","—")}</div>
            </div>
            <div style="background:var(--bg-inset);border-radius:8px;padding:9px 11px;">
              <div class="label-row">Substrate</div>
              <div style="font-size:13px;color:var(--t1);font-weight:500;">{row.get("Substrate","—")}</div>
            </div>
            <div style="background:var(--bg-inset);border-radius:8px;padding:9px 11px;">
              <div class="label-row">Material Spec.</div>
              <div style="font-size:13px;color:var(--t1);font-weight:500;">{row.get("Material_Spec","—")}</div>
            </div>
          </div>
          <div style="margin-top:11px;font-size:12px;color:var(--t3);">
            <span style="color:var(--t2);">Lining Systems: </span>{row.get("Lining_Systems","—")}
          </div>
        </div>
        """, unsafe_allow_html=True)

        sec_header("⚗️ System Code Material Requirements")
        tag_alloc = cascade_allocate([tag_sel])
        for code in sorted(tag_alloc["Lining_System_Code"].unique(), key=lambda x: int(x)):
            sub = tag_alloc[tag_alloc["Lining_System_Code"] == code]
            sname = sub["Lining_System_Short_Name"].iloc[0]
            csqm = sub["Total_SQM"].iloc[0]
            d = sub["Demand_Qty"].sum(); a = sub["Allocated_Qty"].sum()
            pct = (a/d*100) if d > 0 else 100
            with st.expander(f"Code {code} – {sname}  ·  {csqm:,.2f} SQM  ·  {pct:.1f}%", expanded=False):
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("System Code", f"Code {code}")
                m2.metric("Short Name", sname)
                m3.metric("Surface Area", f"{csqm:,.2f} SQM")
                m4.metric("Coverage", f"{pct:.1f}%")
                st.markdown(mat_table_html(sc_alloc_to_mat_df(sub),
                    columns=["Code","Material Name","UOM","Demand","Allocated","Shortfall","Fulfil %"]),
                    unsafe_allow_html=True)

        # Grand total box
        total_d = tag_alloc["Demand_Qty"].sum() if not tag_alloc.empty else 0
        total_a = tag_alloc["Allocated_Qty"].sum() if not tag_alloc.empty else 0
        total_s = tag_alloc["Shortfall_Qty"].sum() if not tag_alloc.empty else 0
        pct = (total_a/total_d*100) if total_d > 0 else 100
        n_codes = tag_alloc["Lining_System_Code"].nunique() if not tag_alloc.empty else 0
        st.markdown(f"""
        <div class="grand-box" style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;">
          <div><div class="label-row">System Codes</div>
               <div style="font-size:20px;font-weight:900;color:var(--amber);font-family:JetBrains Mono;">{n_codes}</div></div>
          <div><div class="label-row">Total Demand</div>
               <div style="font-size:20px;font-weight:900;color:var(--amber);font-family:JetBrains Mono;">{total_d:,.1f}</div></div>
          <div><div class="label-row">Total Shortfall</div>
               <div style="font-size:20px;font-weight:900;color:{'#10B981' if total_s<0.01 else '#EF4444'};
                           font-family:JetBrains Mono;">{total_s:,.1f}</div></div>
          <div><div class="label-row">Coverage</div>
               <div style="font-size:20px;font-weight:900;color:{fc(pct)};font-family:JetBrains Mono;">{pct:.1f}%</div></div>
        </div>
        """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — REPORTS (Session + Location + Total Overview merged)
# ═══════════════════════════════════════════════════════════════════════════════
def tab_reports():
    mode = submode_radio(
        "Report mode",
        ["📦 Session Report", "📍 Location Report", "📈 Total Overview", "📅 Equipment Progress Report"],
        key="reports_mode",
    )

    if   mode.startswith("📦"): _reports_session()
    elif mode.startswith("📍"): _reports_location()
    elif mode.startswith("📈"): _reports_overview()
    else:                       _reports_progress()

def _reports_session():
    s = st.session_state.session_tags
    if not s:
        st.info("📦 Add equipment tags in the Entry tab to generate a report.")
        return
    alloc = cascade_allocate(s)
    mat_agg = alloc_to_mat_df(alloc)
    total_d = alloc["Demand_Qty"].sum()
    total_a = alloc["Allocated_Qty"].sum()
    total_s = alloc["Shortfall_Qty"].sum()
    pct = (total_a/total_d*100) if total_d > 0 else 100
    need_order = int((mat_agg["Shortfall"] > 0).sum()) if not mat_agg.empty else 0

    kpi_strip([
        ("Equipment",        f"{len(s)}",            "session"),
        ("Materials",        f"{len(mat_agg)} unique","across SCs"),
        ("Need to Order",    f"{need_order}",        "materials",   "#EF4444" if need_order>0 else "#10B981"),
        ("Total Shortfall",  f"{total_s:,.1f}",      "mixed units", "#EF4444"),
        ("Overall Coverage", f"{pct:.1f}%",          "alloc/demand",fc(pct)),
    ], cols=5)

    # Reorder list
    st.markdown('<div class="glass-flat">', unsafe_allow_html=True)
    sec_header("⬆⬇ Reorder Priority — changes reflect everywhere")
    for idx, tag in enumerate(list(s)):
        row = eq_master[eq_master["Equipment_Tag_No."] == tag]
        if row.empty: continue
        row = row.iloc[0]
        tag_alloc = alloc[alloc["Equipment_Tag_No."] == tag]
        d = tag_alloc["Demand_Qty"].sum(); a = tag_alloc["Allocated_Qty"].sum()
        tpct = (a/d*100) if d > 0 else 100
        rc1, rc2, rc3 = st.columns([7, 1, 1])
        with rc1:
            st.markdown(f"""
            <div style="display:flex;align-items:center;gap:9px;padding:8px 11px;
                        background:var(--bg-inset);border-radius:8px;
                        border:1px solid var(--border-mid);">
              <div style="width:22px;height:22px;border-radius:50%;background:rgba(245,158,11,.2);
                          color:var(--amber);font-size:10px;font-weight:800;display:flex;
                          align-items:center;justify-content:center;">{idx+1}</div>
              <span style="font-family:JetBrains Mono;font-size:13px;font-weight:800;color:var(--amber);">{tag}</span>
              <span style="font-size:12px;color:var(--t2);">{row["Name"]}</span>
              {loc_badge(row["Location"])}
              <span style="margin-left:auto;">{fulfil_pill(tpct)}</span>
            </div>
            """, unsafe_allow_html=True)
        with rc2:
            if idx > 0 and st.button("↑", key=f"t2up_{tag}_{idx}", use_container_width=True):
                st.session_state.session_tags[idx-1], st.session_state.session_tags[idx] = \
                    st.session_state.session_tags[idx], st.session_state.session_tags[idx-1]
                st.rerun()
        with rc3:
            if idx < len(s)-1 and st.button("↓", key=f"t2dn_{tag}_{idx}", use_container_width=True):
                st.session_state.session_tags[idx+1], st.session_state.session_tags[idx] = \
                    st.session_state.session_tags[idx], st.session_state.session_tags[idx+1]
                st.rerun()
    st.markdown("</div><br>", unsafe_allow_html=True)

    # Per-equipment expanders
    sec_header("Per-Equipment System Code Breakdown")
    for idx, tag in enumerate(s):
        row = eq_master[eq_master["Equipment_Tag_No."] == tag]
        if row.empty: continue
        row = row.iloc[0]
        tag_alloc = alloc[alloc["Equipment_Tag_No."] == tag]
        d = tag_alloc["Demand_Qty"].sum(); a = tag_alloc["Allocated_Qty"].sum()
        tpct = (a/d*100) if d > 0 else 100
        sqm = tag_alloc[["Lining_System_Code","Total_SQM"]].drop_duplicates()["Total_SQM"].sum()
        can_sqm = sqm * tpct / 100
        with st.expander(f"#{idx+1}  {tag}  ·  {row['Name']}  ·  {row['Location']}  ·  {can_sqm:,.0f}/{sqm:,.0f} SQM  ·  {tpct:.1f}%",
                          expanded=False):
            for code in sorted(tag_alloc["Lining_System_Code"].unique(), key=lambda x: int(x)):
                sub = tag_alloc[tag_alloc["Lining_System_Code"] == code]
                sname = sub["Lining_System_Short_Name"].iloc[0]
                cd = sub["Demand_Qty"].sum(); ca = sub["Allocated_Qty"].sum()
                cpct = (ca/cd*100) if cd > 0 else 100
                st.markdown(f"""
                <div style="background:var(--bg-nested);border-radius:8px;padding:13px;margin-bottom:8px;">
                  <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px;">
                    {code_badge(code)}
                    <span style="font-size:12px;color:var(--t2);">{sname}  ·  {sub["Total_SQM"].iloc[0]:,.0f} SQM</span>
                    {fulfil_pill(cpct)}
                  </div>
                </div>
                """, unsafe_allow_html=True)
                st.markdown(mat_table_html(sc_alloc_to_mat_df(sub),
                    columns=["Code","Material Name","UOM","Demand","Allocated","Shortfall","Fulfil %"]),
                    unsafe_allow_html=True)

    # Combined procurement
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="glass">', unsafe_allow_html=True)
    sec_header("🛒 Combined Procurement List")
    if not mat_agg.empty:
        st.markdown(mat_table_html(mat_agg.sort_values("Fulfil %"),
            columns=["Code","Material Name","UOM","Available","Demand","Allocated","Shortfall","Fulfil %"]),
            unsafe_allow_html=True)
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "⬇ Full Session Report",
                excel_bytes(mat_agg, "Session Report", "session"),
                file_name=f"session_report_{date.today():%Y%m%d}.xlsx", key="rep_sess_dl_a",
            )
        with d2:
            short = mat_agg[mat_agg["Shortfall"] > 0]
            if not short.empty:
                st.download_button(
                    "⬇ Order List Only",
                    excel_bytes(short, "Order List", "execution"),
                    file_name=f"order_list_{date.today():%Y%m%d}.xlsx", key="rep_sess_dl_b",
                )
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Smart Reordering Suggestions (session-scope) ────────────
    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("💡 Smart Reordering Suggestions — Session", expanded=False):
        render_suggestion_panel(s, "sug_session")


def _render_equipment_block(loc: str, tag: str, tag_alloc: pd.DataFrame, key_prefix: str):
    """Single-equipment expander used in both Location Based and All Equipment modes.
       Includes metadata grid + per-SC blocks + grand-total bar + Add-to-Session."""
    row = eq_master[eq_master["Equipment_Tag_No."] == tag].iloc[0]
    d = tag_alloc["Demand_Qty"].sum(); a = tag_alloc["Allocated_Qty"].sum()
    s = tag_alloc["Shortfall_Qty"].sum()
    tpct = (a/d*100) if d > 0 else 100
    tsqm  = tag_alloc[["Lining_System_Code","Total_SQM"]].drop_duplicates()["Total_SQM"].sum()
    tcan  = tsqm * tpct / 100
    dot = "🟢" if tpct>=100 else "🟠" if tpct>=90 else "🟡" if tpct>=80 else "🔴"
    meta = " · ".join(p for p in [str(row.get("Type") or "").strip(),
                                    str(row.get("Substrate") or "").strip()]
                       if p and p not in ("nan","—","None"))
    with st.expander(
        f"{dot}  {tag}  ·  {row['Name']}  ·  {meta}  ·  "
        f"{tcan:,.1f}/{tsqm:,.1f} SQM  ·  {tpct:.1f}%",
        expanded=False,
    ):
        # Metadata grid
        m1, m2, m3 = st.columns(3)
        m1.markdown(f'**Type:** {row.get("Type","—") or "—"}')
        m2.markdown(f'**Substrate:** {row.get("Substrate","—") or "—"}')
        m3.markdown(f'**Material Spec.:** {row.get("Material_Spec","—") or "—"}')
        lining = str(row.get("Lining_Systems") or "").replace("\n", " | ").strip()
        if lining:
            st.caption(f"**Lining Systems:** {lining}")
        if tpct >= 100:
            st.markdown(
                '<div style="background:var(--green-bg);border:1px solid rgba(16,185,129,.3);'
                'border-radius:7px;padding:9px 13px;margin:8px 0;color:var(--green);'
                'font-size:12px;font-weight:600;">✅ All materials fully covered — ready to proceed</div>',
                unsafe_allow_html=True)
        st.markdown("<hr style='margin:8px 0;'>", unsafe_allow_html=True)

        # Per system code
        for code in sorted(tag_alloc["Lining_System_Code"].unique(), key=lambda x: int(x)):
            scd = tag_alloc[tag_alloc["Lining_System_Code"] == code]
            sname = scd["Lining_System_Short_Name"].iloc[0]
            csqm = scd["Total_SQM"].iloc[0]
            cd = scd["Demand_Qty"].sum(); ca = scd["Allocated_Qty"].sum()
            cpct = (ca/cd*100) if cd > 0 else 100
            ccan = csqm * cpct / 100
            c_dot = "🟢" if cpct>=100 else "🟠" if cpct>=90 else "🟡" if cpct>=80 else "🔴"
            st.markdown(f"""
            <div style="background:var(--bg-nested);border-radius:8px;padding:10px 14px;
                        margin-bottom:8px;display:flex;align-items:center;gap:10px;
                        flex-wrap:wrap;">
              <span style="font-size:14px;">{c_dot}</span>
              {code_badge(code)}
              <span style="font-size:12px;color:var(--t1);">{sname}</span>
              <span style="font-family:JetBrains Mono;font-size:11px;color:var(--t3);">
                {ccan:,.1f}/{csqm:,.1f} SQM
              </span>
              <span style="margin-left:auto;">{fulfil_pill(cpct)}</span>
            </div>
            """, unsafe_allow_html=True)
            st.markdown(mat_table_html(sc_alloc_to_mat_df(scd),
                columns=["Code","Material Name","UOM","Demand","Allocated","Shortfall","Fulfil %"]),
                unsafe_allow_html=True)

        # Per-equipment grand total bar
        st.markdown(f"""
        <div style="background:rgba(245,158,11,.05);border:1px solid rgba(245,158,11,.18);
                    border-radius:7px;padding:10px 14px;margin-top:10px;
                    display:flex;align-items:center;gap:14px;flex-wrap:wrap;">
          <span style="font-family:JetBrains Mono;color:var(--amber);font-weight:700;font-size:12px;">
            TOTAL — {tag}
          </span>
          <span style="font-size:11px;color:var(--t3);">Demand:
            <b style="color:var(--t1);font-family:JetBrains Mono;">{d:,.3f}</b></span>
          <span style="font-size:11px;color:var(--t3);">Allocated:
            <b style="color:var(--green);font-family:JetBrains Mono;">{a:,.3f}</b></span>
          {f'<span style="font-size:11px;color:var(--t3);">Shortfall: <b style="color:var(--red);font-family:JetBrains Mono;">{s:,.3f}</b></span>' if s > 0.001 else ''}
          <span style="margin-left:auto;">{fulfil_pill(tpct)}</span>
        </div>
        """, unsafe_allow_html=True)

        # Add to Session
        if tag in st.session_state.session_tags:
            st.markdown(
                '<div style="color:var(--green);font-size:12px;font-weight:600;margin-top:8px;">'
                '✓ Already in session</div>', unsafe_allow_html=True)
        else:
            if st.button(f"＋ Add {tag} to Session", key=f"{key_prefix}_{tag}"):
                st.session_state.session_tags.append(tag)
                st.rerun()


def _reports_location():
    sub = st.radio("Scope", ["📍 Location Based", "🌐 All Equipment"],
                    horizontal=True, key="loc_report_mode", label_visibility="collapsed")
    st.markdown("<hr>", unsafe_allow_html=True)

    if sub.startswith("📍"):
        sec_header("📍 All Equipment by Location — Cascading Balance")

        all_loc_sheets = []   # for multi-sheet combined Excel

        for loc in LOCATION_ORDER:
            ltags_seed = eq_master[eq_master["Location"] == loc]["Equipment_Tag_No."].tolist()
            if not ltags_seed: continue
            if loc not in st.session_state.loc_order:
                st.session_state.loc_order[loc] = list(ltags_seed)
            order = [t for t in st.session_state.loc_order[loc] if t in ltags_seed]
            for t in ltags_seed:
                if t not in order: order.append(t)
            st.session_state.loc_order[loc] = order

            lalloc = cascade_allocate(order)
            lsqm = equip_sc[equip_sc["Equipment_Tag_No."].isin(order)]["Total_SQM"].sum()
            ld = lalloc["Demand_Qty"].sum(); la = lalloc["Allocated_Qty"].sum()
            lpct = (la/ld*100) if ld > 0 else 100
            lcan = lsqm * lpct / 100
            color = lc(loc)

            # Location title + drag-reorder helper (↑↓ buttons inline)
            st.markdown(f"""
            <div style="margin-top:22px;padding:10px 14px;background:{color}11;
                        border:1px solid {color}33;border-radius:8px;
                        display:flex;align-items:center;gap:10px;">
              <span class="dot" style="background:{color};"></span>{loc_badge(loc)}
              <span style="font-size:12px;color:var(--t2);">{len(order)} equip  ·  {lcan:,.1f}/{lsqm:,.1f} SQM  ·  Coverage:</span>
              {fulfil_pill(lpct)}
            </div>
            """, unsafe_allow_html=True)

            # Reorder controls (collapsed by default to reduce noise)
            with st.expander(f"⇅ Reorder priority — {loc} ({len(order)} tags)", expanded=False):
                st.caption("Move tags up/down to change cascade priority for this location.")
                for idx, tag in enumerate(order):
                    name = eq_master.set_index("Equipment_Tag_No.")["Name"].get(tag, tag)
                    sqm = sqm_ref[sqm_ref["Equipment_Tag_No."]==tag]["Total_SQM"].sum()
                    rc1, rc2, rc3 = st.columns([7, 1, 1])
                    with rc1:
                        st.markdown(
                            f'<div style="padding:5px 9px;background:var(--bg-inset);'
                            f'border-radius:6px;border:1px solid var(--border-mid);'
                            f'display:flex;align-items:center;gap:9px;">'
                            f'<span style="color:{color};font-weight:800;font-family:JetBrains Mono;'
                            f'font-size:11px;min-width:18px;">{idx+1}</span>'
                            f'<span style="font-family:JetBrains Mono;font-size:12px;font-weight:800;'
                            f'color:var(--amber);">{tag}</span>'
                            f'<span style="font-size:11px;color:var(--t2);">{name[:30]}</span>'
                            f'<span style="margin-left:auto;font-family:JetBrains Mono;'
                            f'font-size:11px;color:var(--t3);">{sqm:,.1f} SQM</span>'
                            f'</div>', unsafe_allow_html=True)
                    with rc2:
                        if idx > 0 and st.button("↑", key=f"locup_{loc}_{tag}_{idx}", use_container_width=True):
                            st.session_state.loc_order[loc][idx-1], st.session_state.loc_order[loc][idx] = \
                                st.session_state.loc_order[loc][idx], st.session_state.loc_order[loc][idx-1]
                            st.rerun()
                    with rc3:
                        if idx < len(order)-1 and st.button("↓", key=f"locdn_{loc}_{tag}_{idx}", use_container_width=True):
                            st.session_state.loc_order[loc][idx+1], st.session_state.loc_order[loc][idx] = \
                                st.session_state.loc_order[loc][idx], st.session_state.loc_order[loc][idx+1]
                            st.rerun()

            # Per-equipment expanders
            for tag in order:
                tag_alloc = lalloc[lalloc["Equipment_Tag_No."] == tag]
                _render_equipment_block(loc, tag, tag_alloc, key_prefix=f"locadd_{loc}")

            # Per-location shortfall chart (collapsed)
            with st.expander(f"📊 Show Shortfall Chart — {loc}", expanded=False):
                chart_data = []
                for tag in order:
                    tag_alloc = lalloc[lalloc["Equipment_Tag_No."] == tag]
                    for code in sorted(tag_alloc["Lining_System_Code"].unique(), key=lambda x: int(x)):
                        ca = tag_alloc[tag_alloc["Lining_System_Code"] == code]
                        sname = ca["Lining_System_Short_Name"].iloc[0]
                        chart_data.append({
                            "Label":     f"{tag} · Code {code} ({sname})",
                            "Demand":    ca["Demand_Qty"].sum(),
                            "Allocated": ca["Allocated_Qty"].sum(),
                            "Shortfall": ca["Shortfall_Qty"].sum(),
                        })
                st.markdown(shortfall_chart_svg(chart_data, accent_col=color),
                            unsafe_allow_html=True)

            # Smart suggestions per location
            with st.expander(f"💡 Smart Reordering Suggestions — {loc}", expanded=False):
                render_suggestion_panel(order, f"sug_{loc}")

            # Per-equipment detail Excel rows (raw allocation, not aggregated)
            export_df = lalloc.merge(
                inv[["Material_Code","Available_Qty","Ordered_Qty"]],
                on="Material_Code", how="left"
            )[["Equipment_Tag_No.","Lining_System_Code","Lining_System_Short_Name",
                "Total_SQM","Material_Code","Material_Name","UOM",
                "Demand_Qty","Available_Qty","Ordered_Qty",
                "Allocated_Qty","Shortfall_Qty","Fulfillment_Pct"]].fillna(0)
            all_loc_sheets.append({
                "name":            loc[:31],
                "df":              export_df,
                "title":           f"Location Report — {loc}",
                "color_scheme":    {"Brown Field":"brown_field","TRAIN J":"train_j","TRAIN K":"train_k"}[loc],
                "add_grand_total": True,
            })

            st.markdown('<hr style="margin:18px 0;">', unsafe_allow_html=True)

        # ── Downloads per location + Combined ──────────
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        sec_header("📥 Download Report per Location")
        cols = st.columns(len(LOCATION_ORDER))
        for ic, loc in enumerate(LOCATION_ORDER):
            ltags = st.session_state.loc_order.get(loc, [])
            if not ltags:
                cols[ic].caption(f"— No equipment in {loc} —"); continue
            lalloc = cascade_allocate(ltags)
            export_df = lalloc.merge(
                inv[["Material_Code","Available_Qty","Ordered_Qty"]],
                on="Material_Code", how="left"
            )[["Equipment_Tag_No.","Lining_System_Code","Lining_System_Short_Name",
                "Total_SQM","Material_Code","Material_Name","UOM",
                "Demand_Qty","Available_Qty","Ordered_Qty",
                "Allocated_Qty","Shortfall_Qty","Fulfillment_Pct"]].fillna(0)
            scheme = {"Brown Field":"brown_field","TRAIN J":"train_j","TRAIN K":"train_k"}[loc]
            with cols[ic]:
                st.download_button(
                    f"⬇ {loc}",
                    excel_bytes(export_df, f"Location Report — {loc}", scheme),
                    file_name=f"location_{loc.replace(' ','_')}_{date.today():%Y%m%d}.xlsx",
                    key=f"loc_dl_{loc}", use_container_width=True,
                )
        # Combined multi-sheet
        if all_loc_sheets:
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                "⬇ All Locations — Combined (Multi-Sheet)",
                excel_bytes_multi(all_loc_sheets),
                file_name=f"location_report_all_{date.today():%Y%m%d}.xlsx",
                key="dl_loc_all", use_container_width=True,
            )
        # Print button
        st.markdown("""
        <div style="margin-top:14px;">
          <button class="no-print" onclick="window.print()" style="
            background:transparent;border:1px solid var(--border-strong);border-radius:8px;
            padding:8px 18px;color:var(--t2);font-family:Inter;font-size:12px;cursor:pointer;">
            🖨 Print / Save as PDF
          </button>
          <span style="margin-left:10px;font-size:11px;color:var(--t4);">
            (Sidebar &amp; header hide automatically during print)
          </span>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    else:
        # ── All Equipment mode ──
        sec_header("🌐 All Equipment — Global Cascading Balance")
        st.caption("All equipment in priority order. Inventory pool is shared globally across all locations.")
        # init / sync order
        cur = [t for t in st.session_state.all_eq_order if t in ALL_TAGS]
        for t in ALL_TAGS:
            if t not in cur: cur.append(t)
        st.session_state.all_eq_order = cur

        alloc_all = cascade_allocate(cur)
        sqm_all = equip_sc["Total_SQM"].sum()
        d = alloc_all["Demand_Qty"].sum(); a = alloc_all["Allocated_Qty"].sum()
        pct = (a/d*100) if d > 0 else 100
        can = sqm_all * pct / 100
        kpi_strip([
            ("Equipment",       f"{len(cur)}",       "all tags"),
            ("Total SQM",       f"{sqm_all:,.1f}",   ""),
            ("Available Cov.",  f"{can:,.1f}",       "SQM"),
            ("SQM Deficit",     f"{max(0,sqm_all-can):,.1f}",     "", "#EF4444"),
            ("Overall Cov.",    f"{pct:.1f}%",       "", fc(pct)),
        ], cols=5)

        # Reorder helper
        with st.expander(f"⇅ Reorder global priority ({len(cur)} tags)", expanded=False):
            for idx, tag in enumerate(cur):
                name = eq_master.set_index("Equipment_Tag_No.")["Name"].get(tag, tag)
                loc  = eq_master.set_index("Equipment_Tag_No.")["Location"].get(tag, "")
                rc1, rc2, rc3 = st.columns([7, 1, 1])
                with rc1:
                    st.markdown(
                        f'<div style="padding:5px 9px;background:var(--bg-inset);'
                        f'border-radius:6px;border:1px solid var(--border-mid);'
                        f'display:flex;align-items:center;gap:9px;">'
                        f'<span style="color:var(--amber);font-weight:800;font-family:JetBrains Mono;'
                        f'font-size:11px;min-width:18px;">{idx+1}</span>'
                        f'<span style="font-family:JetBrains Mono;font-size:12px;font-weight:800;'
                        f'color:var(--amber);">{tag}</span>'
                        f'<span style="font-size:11px;color:var(--t2);">{name[:30]}</span>'
                        f'{loc_badge(loc)}'
                        f'</div>', unsafe_allow_html=True)
                with rc2:
                    if idx > 0 and st.button("↑", key=f"aeup_{tag}_{idx}", use_container_width=True):
                        st.session_state.all_eq_order[idx-1], st.session_state.all_eq_order[idx] = \
                            st.session_state.all_eq_order[idx], st.session_state.all_eq_order[idx-1]
                        st.rerun()
                with rc3:
                    if idx < len(cur)-1 and st.button("↓", key=f"aedn_{tag}_{idx}", use_container_width=True):
                        st.session_state.all_eq_order[idx+1], st.session_state.all_eq_order[idx] = \
                            st.session_state.all_eq_order[idx], st.session_state.all_eq_order[idx+1]
                        st.rerun()

        sec_header("Per-Equipment Detail")
        for tag in cur:
            tag_alloc = alloc_all[alloc_all["Equipment_Tag_No."] == tag]
            _render_equipment_block("ALL", tag, tag_alloc, key_prefix="ae")

        with st.expander("💡 Smart Reordering Suggestions — All Equipment", expanded=False):
            render_suggestion_panel(cur, "sug_all")

        st.download_button(
            "⬇ Download All Equipment Report",
            excel_bytes(alloc_all, "All Equipment — Global Report", "overview"),
            file_name=f"all_equipment_report_{date.today():%Y%m%d}.xlsx",
            key="dl_ae_all",
        )


def _reports_overview():
    sec_header("📈 Total Overview — Master Equipment & Material Table")
    fc1, fc2, fc3, fc4 = st.columns(4)
    with fc1:
        f_loc = st.multiselect("Location", LOCATION_ORDER, default=LOCATION_ORDER, key="ov_loc")
    with fc2:
        t_opts = sorted(eq_master["Type"].dropna().unique())
        f_type = st.multiselect("Type", t_opts, default=t_opts, key="ov_type")
    with fc3:
        c_opts = sorted(equip_sc["Lining_System_Code"].unique(), key=lambda x: int(x))
        f_code = st.multiselect("System Code", c_opts, default=c_opts, key="ov_code")
    with fc4:
        f_status = st.selectbox("Status",
            ["All","Fully Ready (100%)","Partial (50-99%)","Blocked (<50%)"], key="ov_status")

    # Build master rows
    eq_f = eq_master[eq_master["Location"].isin(f_loc) & eq_master["Type"].isin(f_type)]
    tags_f = eq_f["Equipment_Tag_No."].tolist()
    sc_f = equip_sc[equip_sc["Equipment_Tag_No."].isin(tags_f) & equip_sc["Lining_System_Code"].isin(f_code)]
    if sc_f.empty:
        st.warning("No rows match the current filters.")
        return
    alloc = cascade_allocate(tags_f)

    rows = []
    for _, sr in sc_f.iterrows():
        sub = alloc[(alloc["Equipment_Tag_No."] == sr["Equipment_Tag_No."]) &
                    (alloc["Lining_System_Code"] == sr["Lining_System_Code"])]
        d = sub["Demand_Qty"].sum(); a = sub["Allocated_Qty"].sum()
        pct = (a/d*100) if d > 0 else 100
        em_r = eq_master[eq_master["Equipment_Tag_No."] == sr["Equipment_Tag_No."]].iloc[0]
        rows.append({
            "Equipment No": sr["Equipment_Tag_No."],
            "Name":         em_r["Name"],
            "Substrate":    em_r["Substrate"],
            "Type":         em_r["Type"],
            "Location":     em_r["Location"],
            "Code":         f"Code {sr['Lining_System_Code']}",
            "System Name":  sr["Lining_System_Short_Name"],
            "Total SQM":    sr["Total_SQM_Original"],
            "Done SQM":     sr.get("done_sqm", 0),
            "Remaining":    sr["Total_SQM"],
            "Demand":       d,
            "Allocated":    a,
            "Shortfall":    max(0, d - a),
            "Fulfil %":     pct,
        })
    df = pd.DataFrame(rows)
    if f_status == "Fully Ready (100%)":
        df = df[df["Fulfil %"] >= 100]
    elif f_status == "Partial (50-99%)":
        df = df[(df["Fulfil %"] >= 50) & (df["Fulfil %"] < 100)]
    elif f_status == "Blocked (<50%)":
        df = df[df["Fulfil %"] < 50]

    kpi_strip([
        ("No. of Rows",   f"{len(df)}",                 "filtered"),
        ("Total SQM",     f"{df['Total SQM'].sum():,.1f}", ""),
        ("Done SQM",      f"{df['Done SQM'].sum():,.1f}",  ""),
        ("Remaining SQM", f"{df['Remaining'].sum():,.1f}", ""),
        ("Shortfall SQM", f"{max(0,df['Total SQM'].sum()-df['Total SQM'].sum()*df['Fulfil %'].mean()/100):,.1f}",
                          "", "#EF4444"),
        ("Avg Coverage",  f"{df['Fulfil %'].mean():.1f}%" if len(df) else "—",
                          "", fc(df['Fulfil %'].mean() if len(df) else 100)),
    ], cols=6)

    # Color-style the master table via st.dataframe with row tint
    def _row_style(row):
        return [f'background-color: {fc(row["Fulfil %"])}1A' for _ in row]
    styled = (df.style.apply(_row_style, axis=1)
                .format({"Total SQM":"{:,.2f}","Done SQM":"{:,.2f}",
                         "Remaining":"{:,.2f}","Demand":"{:,.3f}",
                         "Allocated":"{:,.3f}","Shortfall":"{:,.3f}","Fulfil %":"{:.1f}%"}))
    st.dataframe(styled, use_container_width=True, hide_index=True,
                 height=min(700, 50 + len(df)*35))

    st.download_button(
        "⬇ Download Filtered Master Table",
        excel_bytes(df, "Total Overview", "overview"),
        file_name=f"total_overview_{date.today():%Y%m%d}.xlsx",
        key="ov_dl",
    )


def _reports_progress():
    """Per-equipment date-by-date progress timeline.
       Shows status (Not Started / In Progress / Complete), per-day work log
       (SQM done, system code, materials + qty consumed, notes)."""
    sec_header("📅 Equipment Progress Report — Date-by-Date Timeline")
    if not db_available():
        st.warning("Database required for the progress report.")
        return

    # ── Equipment selector with cascading Location → Type → Tag ────
    c1, c2, c3 = st.columns(3)
    with c1:
        f_loc = st.selectbox("📍 Location", ["All"] + LOCATION_ORDER, key="prog_rep_loc")
    em_f = eq_master.copy()
    if f_loc != "All": em_f = em_f[em_f["Location"] == f_loc]
    with c2:
        types = ["All"] + sorted(em_f["Type"].dropna().unique().tolist())
        f_type = st.selectbox("🔧 Type", types, key="prog_rep_type")
    if f_type != "All": em_f = em_f[em_f["Type"] == f_type]
    with c3:
        tag_list = em_f["Equipment_Tag_No."].tolist()
        if not tag_list:
            st.info("No equipment matches the filters."); return
        sel_tag = st.selectbox("🏷️ Equipment Tag", tag_list,
            format_func=lambda t: f"{t}  —  {em_f[em_f['Equipment_Tag_No.']==t]['Name'].iloc[0]}",
            key="prog_rep_tag")

    # ── Pull equipment SQM progress + consumption history ────
    em_row = eq_master[eq_master["Equipment_Tag_No."] == sel_tag].iloc[0]
    prog = sqm_ref[sqm_ref["Equipment_Tag_No."] == sel_tag].copy()
    if prog.empty:
        st.warning("No SQM data for this equipment."); return

    total_sqm = prog["Total_SQM_Original"].sum()
    done_sqm  = prog["done_sqm"].sum()
    pct = (done_sqm / total_sqm * 100) if total_sqm > 0 else 0
    if pct >= 100:
        status_txt = "✅ Complete"; status_col = "var(--green)"; status_bg = "rgba(16,185,129,.13)"
    elif pct > 0:
        status_txt = "🔄 In Progress"; status_col = "var(--amber)"; status_bg = "rgba(245,158,11,.13)"
    else:
        status_txt = "⏳ Not Started"; status_col = "var(--t3)"; status_bg = "rgba(100,116,139,.13)"

    # ── Status hero card ────
    st.markdown(f"""
    <div class="glass glass-amber-bar" style="margin-bottom:14px;">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px;">
        <div>
          <div class="label-row">Equipment Tag</div>
          <div style="font-size:26px;font-weight:900;color:var(--amber);
                      font-family:JetBrains Mono;">{sel_tag}</div>
          <div style="font-size:14px;color:var(--t1);margin-top:4px;">{em_row["Name"]}</div>
        </div>
        <div style="text-align:right;">
          {loc_badge(em_row["Location"])}
          <div style="margin-top:6px;display:inline-block;background:{status_bg};color:{status_col};
                       padding:5px 14px;border-radius:20px;font-size:13px;font-weight:700;">
            {status_txt}
          </div>
        </div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-top:10px;">
        <div><div class="label-row">Total SQM</div>
             <div style="font-size:18px;font-weight:800;color:var(--t0);font-family:JetBrains Mono;">{total_sqm:,.2f}</div></div>
        <div><div class="label-row">Done SQM</div>
             <div style="font-size:18px;font-weight:800;color:var(--green);font-family:JetBrains Mono;">{done_sqm:,.2f}</div></div>
        <div><div class="label-row">Remaining SQM</div>
             <div style="font-size:18px;font-weight:800;color:var(--amber);font-family:JetBrains Mono;">{total_sqm-done_sqm:,.2f}</div></div>
        <div><div class="label-row">Completion</div>
             <div style="font-size:18px;font-weight:800;color:{fc(pct)};font-family:JetBrains Mono;">{pct:.1f}%</div></div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Per system code progress strip ────
    sec_header("📐 Per System Code Progress")
    for _, pr in prog.iterrows():
        code = pr["Lining_System_Code"]
        sname_row = recipe[recipe["Lining_System_Code"] == code]
        sname = sname_row["Lining_System_Short_Name"].iloc[0] if not sname_row.empty else ""
        ssqm = pr["Total_SQM_Original"]
        sdone = pr["done_sqm"]
        sp = (sdone / ssqm * 100) if ssqm > 0 else 0
        st.markdown(f"""
        <div style="background:var(--bg-card-solid);border:1px solid var(--border-mid);
                    border-radius:10px;padding:11px 14px;margin-bottom:7px;
                    display:flex;align-items:center;gap:11px;">
          {code_badge(code)}
          <span style="font-size:13px;color:var(--t1);">{sname}</span>
          <span style="margin-left:auto;font-family:JetBrains Mono;font-size:12px;color:var(--t2);">
            {sdone:,.2f} / {ssqm:,.2f} SQM
          </span>
          {fulfil_pill(sp)}
        </div>
        """, unsafe_allow_html=True)

    # ── Date-by-date timeline ────
    conn = get_db()
    log = pd.read_sql(
        "SELECT entry_date, lining_system_code, lining_system_name, sqm_completed, "
        "material_code, material_name, uom, consumed_qty, notes, submitted_at "
        "FROM consumption_log WHERE equipment_tag = ? ORDER BY entry_date DESC, id DESC",
        conn, params=(sel_tag,))
    conn.close()

    sec_header("📅 Work-Day Timeline")
    if log.empty:
        st.info("No consumption entries logged yet for this equipment.")
        return

    # Group by date, then within date by system code
    for d in log["entry_date"].unique():
        day = log[log["entry_date"] == d]
        # SQM done per code on this date
        day_codes = day.groupby(["lining_system_code","lining_system_name"]).agg(
            sqm=("sqm_completed","max"),
            notes=("notes","first"),
        ).reset_index()
        total_day_sqm = day_codes["sqm"].sum()
        st.markdown(f"""
        <div style="background:var(--bg-card-solid);border-left:3px solid var(--amber);
                    border:1px solid var(--border-mid);border-radius:10px;padding:14px 16px;
                    margin-bottom:10px;">
          <div style="display:flex;align-items:center;gap:12px;margin-bottom:10px;">
            <span style="background:var(--amber-bg);color:var(--amber);padding:4px 12px;
                         border-radius:6px;font-family:JetBrains Mono;font-size:13px;font-weight:700;">
              📅 {d}
            </span>
            <span style="font-size:12px;color:var(--t2);">
              Total SQM done: <strong style="color:var(--green);font-family:JetBrains Mono;">{total_day_sqm:,.2f}</strong>
            </span>
            <span style="font-size:11px;color:var(--t4);margin-left:auto;">
              {len(day)} material entries
            </span>
          </div>
        </div>
        """, unsafe_allow_html=True)
        # Per system code within the day
        for _, dc in day_codes.iterrows():
            ccode = dc["lining_system_code"]; cname = dc["lining_system_name"] or ""
            csqm  = dc["sqm"]; cnotes = dc["notes"] or ""
            mats = day[day["lining_system_code"] == ccode][
                ["material_code","material_name","uom","consumed_qty"]
            ].rename(columns={"material_code":"Code","material_name":"Material Name",
                                "uom":"UOM","consumed_qty":"Allocated"})
            mats["Demand"] = mats["Allocated"]
            mats["Shortfall"] = 0
            mats["Fulfil %"] = 100
            st.markdown(f"""
            <div style="margin-left:24px;margin-bottom:10px;">
              <div style="display:flex;align-items:center;gap:10px;margin-bottom:7px;">
                {code_badge(ccode)}
                <span style="font-size:12px;color:var(--t2);">{cname}</span>
                <span style="font-size:11px;color:var(--t3);font-family:JetBrains Mono;">
                  · {csqm:,.2f} SQM done
                </span>
              </div>
            </div>
            """, unsafe_allow_html=True)
            st.markdown(mat_table_html(mats, columns=[
                "Code","Material Name","UOM","Allocated"]), unsafe_allow_html=True)
            if cnotes:
                st.markdown(
                    f'<div style="margin:6px 0 10px 24px;font-size:11px;color:var(--t3);'
                    f'font-style:italic;">📝 {cnotes}</div>',
                    unsafe_allow_html=True)

    # Download full progress log for this equipment
    st.download_button(
        f"⬇ Download Progress Log — {sel_tag}",
        excel_bytes(log, f"Progress {sel_tag}", "overview"),
        file_name=f"progress_{sel_tag}_{date.today():%Y%m%d}.xlsx",
        key="prog_rep_dl",
    )


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 4 — PLAN
# ═══════════════════════════════════════════════════════════════════════════════
def tab_plan():
    sub = submode_radio("View", ["⚙️ Execution Plan", "📋 Progress List"], key="plan_sub")

    if sub.startswith("⚙️"):
        if not st.session_state.session_tags:
            st.info("Add equipment tags in the Entry tab first.")
            return
        c1, c2 = st.columns(2)
        with c1:
            sel_tag = st.selectbox(
                "Select Equipment", st.session_state.session_tags,
                format_func=lambda t: f"{t}  —  {eq_master[eq_master['Equipment_Tag_No.']==t]['Name'].iloc[0]}",
                key="exec_tag")
        tag_alloc = cascade_allocate(st.session_state.session_tags)
        tag_alloc = tag_alloc[tag_alloc["Equipment_Tag_No."] == sel_tag]
        codes = sorted(tag_alloc["Lining_System_Code"].unique(), key=lambda x: int(x))
        if not codes:
            st.warning("No system code data for this equipment.")
            return
        with c2:
            sel_code = st.selectbox(
                "Select Critical System Code", codes,
                format_func=lambda c: f"Code {c}  —  {tag_alloc[tag_alloc['Lining_System_Code']==c]['Lining_System_Short_Name'].iloc[0]}",
                key="exec_code")

        sub_df = tag_alloc[tag_alloc["Lining_System_Code"] == sel_code]
        sname = sub_df["Lining_System_Short_Name"].iloc[0]
        sqm   = sub_df["Total_SQM"].iloc[0]
        d = sub_df["Demand_Qty"].sum(); a = sub_df["Allocated_Qty"].sum()
        pct = (a/d*100) if d > 0 else 100
        shortfalls = sub_df[sub_df["Shortfall_Qty"] > 0]

        st.markdown(f"""
        <div class="glass glass-amber-bar" style="margin-bottom:16px;">
          <div class="label-row">Critical System Code</div>
          <div style="display:flex;align-items:center;gap:11px;margin-bottom:9px;">
            {code_badge(sel_code)}
            <span style="font-size:14px;color:var(--t2);">{sname}</span>
            <span style="font-size:12px;color:var(--t3);">·  {sqm:,.2f} SQM</span>
            <span style="font-size:28px;font-weight:900;color:{fc(pct)};
                         font-family:JetBrains Mono;margin-left:auto;">{pct:.1f}%</span>
          </div>
          <div style="color:{'#10B981' if pct>=100 else '#EAB308'};font-size:13px;">
            {'✅ All materials for this system code are fully covered.' if pct>=100
             else f'⚠️ {len(shortfalls)} materials short — order these first to proceed.'}
          </div>
        </div>
        """, unsafe_allow_html=True)

        sec_header("All Materials — Status Overview")
        st.markdown(mat_table_html(sc_alloc_to_mat_df(tag_alloc),
            columns=["Code","Material Name","UOM","Demand","Allocated","Shortfall","Fulfil %"]),
            unsafe_allow_html=True)

        sec_header("📋 Procurement Order Priority")
        # Critical block
        st.markdown(f"""
        <div style="background:rgba(239,68,68,.06);border:2px solid rgba(239,68,68,.25);
                    border-radius:12px;padding:15px;margin-bottom:12px;">
          <div style="font-size:13px;font-weight:700;color:var(--red);margin-bottom:7px;">
            🔴 Order First — Code {sel_code} ({sname}) · Critical Path
          </div>
          <div style="font-size:12px;color:var(--t2);margin-bottom:10px;">
            {len(shortfalls)} material(s) need to be procured
          </div>
        </div>
        """, unsafe_allow_html=True)
        if not shortfalls.empty:
            st.markdown(mat_table_html(sc_alloc_to_mat_df(shortfalls),
                columns=["Code","Material Name","UOM","Demand","Allocated","Shortfall","Fulfil %"]),
                unsafe_allow_html=True)

        # Other codes
        for code in codes:
            if code == sel_code: continue
            csub = tag_alloc[tag_alloc["Lining_System_Code"] == code]
            cd = csub["Demand_Qty"].sum(); ca = csub["Allocated_Qty"].sum()
            cpct = (ca/cd*100) if cd > 0 else 100
            cn = csub["Lining_System_Short_Name"].iloc[0]
            cshort = csub[csub["Shortfall_Qty"] > 0]
            st.markdown(f"""
            <div style="background:rgba(245,158,11,.05);border:1px solid rgba(245,158,11,.2);
                        border-radius:12px;padding:15px;margin-bottom:12px;">
              <div style="font-size:13px;font-weight:700;color:var(--amber);margin-bottom:5px;">
                🟡 Order Next — Code {code} ({cn}) · Coverage: {cpct:.1f}%
              </div>
              <div style="font-size:12px;color:{'#10B981' if cshort.empty else 'var(--t2)'};">
                {'All materials covered ✅' if cshort.empty
                 else f'{len(cshort)} material(s) short. Order after critical code is secured.'}
              </div>
            </div>
            """, unsafe_allow_html=True)
            if not cshort.empty:
                st.markdown(mat_table_html(sc_alloc_to_mat_df(cshort),
                    columns=["Code","Material Name","UOM","Demand","Allocated","Shortfall","Fulfil %"]),
                    unsafe_allow_html=True)

        # Summary box
        all_short = tag_alloc[tag_alloc["Shortfall_Qty"] > 0]
        st.markdown(f"""
        <div class="grand-box">
          <div style="font-size:13px;font-weight:700;color:var(--amber);margin-bottom:9px;">
            📋 Execution Summary
          </div>
          <div style="font-size:12px;color:var(--t2);line-height:1.9;">
            <div>• Critical code: Code {sel_code} ({sname}) — {pct:.1f}% coverage</div>
            <div>• Critical materials to order: {len(shortfalls)} ({shortfalls["Shortfall_Qty"].sum():,.1f} units)</div>
            <div>• Total to order for completion: {all_short["Shortfall_Qty"].sum():,.1f} units across {len(all_short)} material(s)</div>
          </div>
        </div>
        """, unsafe_allow_html=True)
        if not all_short.empty:
            st.download_button(
                f"⬇ Download Execution Order List — {sel_tag}",
                excel_bytes(sc_alloc_to_mat_df(all_short), f"Execution Plan {sel_tag}", "execution"),
                file_name=f"execution_plan_{sel_tag}_{date.today():%Y%m%d}.xlsx",
                key="exec_dl",
            )

    else:
        # PROGRESS LIST
        if not db_available():
            st.warning("Database required for the Progress List.")
            return
        sec_header("📋 Project Progress")
        prog = sqm_ref.copy()
        prog["Equipment Tag"] = prog["Equipment_Tag_No."]
        prog = prog.merge(eq_master[["Equipment_Tag_No.","Name","Location"]],
                          left_on="Equipment_Tag_No.", right_on="Equipment_Tag_No.", how="left")
        prog["Completion %"] = (prog["done_sqm"]/prog["Total_SQM_Original"].replace(0,np.nan)*100).fillna(0).round(1)
        prog["Status"] = prog["Completion %"].apply(
            lambda p: "✅ Complete" if p>=100 else ("🔄 In Progress" if p>0 else "⏳ Not Started")
        )
        prog["System Code"] = "Code " + prog["Lining_System_Code"].astype(str)
        prog["Total SQM"]   = prog["Total_SQM_Original"]
        prog["Done SQM"]    = prog["done_sqm"]
        prog["Remaining"]   = prog["Total_SQM_Original"] - prog["done_sqm"]

        kpi_strip([
            ("Total SQM",     f"{prog['Total SQM'].sum():,.1f}", ""),
            ("Done SQM",      f"{prog['Done SQM'].sum():,.1f}", ""),
            ("Remaining SQM", f"{prog['Remaining'].sum():,.1f}", ""),
            ("Completion",    f"{(prog['Done SQM'].sum()/prog['Total SQM'].sum()*100) if prog['Total SQM'].sum() else 0:.1f}%",
                              "", fc(prog['Done SQM'].sum()/prog['Total SQM'].sum()*100 if prog['Total SQM'].sum() else 0)),
        ], cols=4)

        fc1, fc2 = st.columns(2)
        with fc1:
            f_loc = st.selectbox("Filter Location", ["All"] + LOCATION_ORDER, key="prog_loc_f")
        with fc2:
            f_st  = st.selectbox("Filter Status",
                ["All","✅ Complete","🔄 In Progress","⏳ Not Started"], key="prog_status_f")
        view = prog.copy()
        if f_loc != "All": view = view[view["Location"] == f_loc]
        if f_st  != "All": view = view[view["Status"]   == f_st]

        show = view[["Equipment Tag","System Code","Location","Name","Total SQM","Done SQM",
                     "Remaining","Completion %","Status"]]
        st.dataframe(show, use_container_width=True, hide_index=True,
                     height=min(700, 60 + len(show)*35))
        # Download exports the FILTERED `show` (not the full prog table)
        st.download_button(
            f"⬇ Download Filtered Progress List ({len(show)} rows)",
            excel_bytes(show, "Progress List (Filtered)", "overview"),
            file_name=f"progress_list_filtered_{date.today():%Y%m%d}.xlsx",
            key="prog_dl",
        )


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 5 — INVENTORY
# ═══════════════════════════════════════════════════════════════════════════════
def tab_inventory():
    if not db_available():
        st.error("Database not found. Run `python setup_db.py` first.")
        return
    mode = submode_radio(
        "Mode",
        ["📊 Main Inventory","📅 Consumption","📦 Receipts","📋 Ordered"],
        key="inv_mode",
    )
    if   mode.startswith("📊"): _inv_main()
    elif mode.startswith("📅"): _inv_consumption()
    elif mode.startswith("📦"): _inv_receipts()
    else:                       _inv_ordered()


def _inv_main():
    sec_header("📊 Main Inventory Dashboard — Movements by Date Range")
    c1, c2, c3 = st.columns(3)
    with c1:
        d_from = st.date_input("From Date", value=date.today().replace(day=1), key="mi_from")
    with c2:
        d_to   = st.date_input("To Date", value=date.today(), key="mi_to")
    with c3:
        view = st.radio("Group By", ["By Material","By System Code"],
                         horizontal=True, key="mi_view", label_visibility="collapsed")

    conn = get_db()
    cons = pd.read_sql(
        "SELECT material_code, SUM(consumed_qty) AS consumed "
        "FROM consumption_log WHERE entry_date BETWEEN ? AND ? GROUP BY material_code",
        conn, params=(d_from.isoformat(), d_to.isoformat()))
    rec = pd.read_sql(
        "SELECT material_code, SUM(received_qty) AS received "
        "FROM receipt_log WHERE entry_date BETWEEN ? AND ? GROUP BY material_code",
        conn, params=(d_from.isoformat(), d_to.isoformat()))
    conn.close()

    df = inv[["Material_Code","Material_Name","UOM","Available_Qty","Ordered_Qty"]].copy()
    df = df.merge(cons, left_on="Material_Code", right_on="material_code", how="left")
    df = df.merge(rec,  left_on="Material_Code", right_on="material_code", how="left")
    df["consumed"] = df["consumed"].fillna(0)
    df["received"] = df["received"].fillna(0)
    df = df.rename(columns={"Material_Code":"Code","Material_Name":"Material Name",
                            "received":"Total Receipts","consumed":"Total Consumed",
                            "Available_Qty":"Current Stock","Ordered_Qty":"Ordered Qty"})

    if view == "By Material":
        st.markdown(mat_table_html(df,
            columns=["Code","Material Name","UOM","Total Receipts","Total Consumed","Current Stock","Ordered Qty"]),
            unsafe_allow_html=True)
        st.download_button(
            "⬇ Download Inventory Dashboard",
            excel_bytes(df, "Inventory Dashboard", "dashboard"),
            file_name=f"inventory_dashboard_{date.today():%Y%m%d}.xlsx",
            key="mi_dl",
        )
    else:
        # Group by system code
        codes = sorted(recipe["Lining_System_Code"].unique(), key=lambda x: int(x))
        for code in codes:
            mats = recipe[recipe["Lining_System_Code"] == code]["Material_Code"].unique()
            sub_df = df[df["Code"].isin(mats)]
            if sub_df.empty: continue
            cnt = len(sub_df)
            zero = (sub_df["Current Stock"] <= 0).sum()
            sname = recipe[recipe["Lining_System_Code"] == code]["Lining_System_Short_Name"].iloc[0]
            with st.expander(f"Code {code} — {sname}  ·  {cnt} materials  ·  {zero} at zero", expanded=False):
                st.markdown(mat_table_html(sub_df,
                    columns=["Code","Material Name","UOM","Total Receipts","Total Consumed","Current Stock","Ordered Qty"]),
                    unsafe_allow_html=True)


def _inv_consumption():
    sec_header("Step 1 — Select Work Location & Equipment")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        loc = st.selectbox("📍 Location", [""] + LOCATION_ORDER, key="ce_loc")
    types = sorted(eq_master[eq_master["Location"] == loc]["Type"].unique()) if loc else []
    with c2:
        ttype = st.selectbox("🔧 Type", [""] + list(types), key="ce_type")
    tags_pool = (eq_master[(eq_master["Location"] == loc) & (eq_master["Type"] == ttype)]
                 if loc and ttype else pd.DataFrame(columns=eq_master.columns))
    tag_opts = tags_pool["Equipment_Tag_No."].tolist() if not tags_pool.empty else []
    with c3:
        tag = st.selectbox(
            "🏷️ Equipment Tag",
            [""] + tag_opts,
            format_func=lambda t: "— Select —" if t=="" else
                                  (f"{t}  —  {tags_pool[tags_pool['Equipment_Tag_No.']==t]['Name'].iloc[0]}"
                                   if t in tag_opts else t),
            key="ce_tag",
        )
    tag_codes = sorted(equip_sc[equip_sc["Equipment_Tag_No."] == tag]["Lining_System_Code"].unique(),
                       key=lambda x: int(x)) if tag else []
    with c4:
        code = st.selectbox(
            "⚗️ System Code",
            [""] + tag_codes,
            format_func=lambda c: "— Select —" if c=="" else
                                  f"Code {c} – {equip_sc[(equip_sc['Equipment_Tag_No.']==tag) & (equip_sc['Lining_System_Code']==c)]['Lining_System_Short_Name'].iloc[0]}" if c in tag_codes else c,
            key="ce_code",
        )

    if not (tag and code): return

    sc_row = equip_sc[(equip_sc["Equipment_Tag_No."]==tag) & (equip_sc["Lining_System_Code"]==code)].iloc[0]
    total_sqm = sc_row["Total_SQM_Original"]
    done_sqm  = sc_row["done_sqm"]
    remaining = total_sqm - done_sqm
    sname     = sc_row["Lining_System_Short_Name"]

    kpi_strip([
        ("System Code",    f"Code {code}", sname),
        ("Original SQM",   f"{total_sqm:,.2f}", "total"),
        ("Already Done",   f"{done_sqm:,.2f}", "previous"),
        ("Remaining SQM",  f"{remaining:,.2f}", "to complete"),
    ], cols=4)

    rec_mats = recipe[recipe["Lining_System_Code"] == code]
    st.markdown('<div class="glass-flat">', unsafe_allow_html=True)
    sec_header("Step 2 — Enter SQM Completed Today")
    with st.form("ce_form", clear_on_submit=False):
        h1, h2 = st.columns(2)
        with h1:
            work_date = st.date_input("📅 Work Date", value=date.today(), key="form_ce_date")
        with h2:
            sqm_today = st.number_input("SQM Completed Today", min_value=0.0,
                                         max_value=float(remaining), value=0.0, step=0.5,
                                         format="%.2f", key="form_ce_sqm")

        sec_header("Step 3 — Material Quantities Consumed")
        st.caption("Defaults are computed from the per-SQM recipe. Override with actual consumed values.")
        mat_inputs = {}
        for _, mr in rec_mats.iterrows():
            mc = mr["Material_Code"]
            for_1 = float(mr["For_1_SQM"] or 0)
            expected = for_1 * sqm_today
            avail = INV_POOL_INIT.get(mc, 0)
            on_ord = INV_ORDERED_INIT.get(mc, 0)
            cc = st.columns([2, 3, 1, 1.5, 1.5, 1.5, 1.5])
            cc[0].markdown(f'<code>{mc}</code>', unsafe_allow_html=True)
            cc[1].write(mr["Material_Name"])
            cc[2].write(mr["UOM"])
            cc[3].write(f"{avail:,.3f}")
            cc[4].write(f"{for_1:,.3f}")
            qty = cc[5].number_input("qty", min_value=0.0, value=float(expected),
                                       step=0.001, format="%.3f", label_visibility="collapsed",
                                       key=f"form_mat_{mc}")
            cc[6].write(f"{on_ord:,.3f}")
            mat_inputs[mc] = (qty, expected, mr["Material_Name"], mr["UOM"])
        notes = st.text_area("📝 Notes (optional)", placeholder="Weather, issues, remarks…",
                              key="form_notes", height=68)
        submit = st.form_submit_button("➕ Add to Grid")
    st.markdown("</div>", unsafe_allow_html=True)

    if submit and sqm_today > 0:
        conn = get_db()
        for mc, (qty, exp, mn, uom) in mat_inputs.items():
            variance = ((qty - exp) / exp * 100) if exp > 0 else 0
            status = "OK" if abs(variance) < 5 else ("Over" if variance > 0 else "Less")
            conn.execute(
                "INSERT INTO draft_consumption (session_key, entry_date, equipment_tag, "
                "lining_system_code, lining_system_name, sqm_completed, material_code, "
                "material_name, uom, expected_qty, actual_qty, effective_qty, "
                "variance_pct, variance_status, notes) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (st.session_state["_session_key"], work_date.isoformat(), tag, code, sname,
                 sqm_today, mc, mn, uom, round(exp, 4), round(qty, 4), round(qty, 4),
                 round(variance, 2), status, notes),
            )
        _commit_and_refresh(conn)

    # Draft grid + Submit
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="glass">', unsafe_allow_html=True)
    sec_header("📋 Draft Consumption Grid — Pending Submission")
    conn = get_db()
    draft = pd.read_sql(
        "SELECT id, entry_date, equipment_tag, lining_system_code, sqm_completed, "
        "material_code, material_name, uom, expected_qty, effective_qty, "
        "variance_pct, variance_status, notes "
        "FROM draft_consumption WHERE session_key = ? ORDER BY id",
        conn, params=(st.session_state["_session_key"],),
    )
    if draft.empty:
        st.info("No pending entries. Use 'Add to Grid' above to stage data.")
    else:
        st.dataframe(draft, use_container_width=True, hide_index=True,
                     height=min(500, 55 + len(draft)*35))
        b1, b2, _ = st.columns([2, 2, 4])
        with b1:
            if st.button("🧹 Clear All Draft", key="clear_draft_all", use_container_width=True):
                conn.execute("DELETE FROM draft_consumption WHERE session_key = ?",
                             (st.session_state["_session_key"],))
                _commit_and_refresh(conn)
        with b2:
            if st.button("✅ Submit Consumption", key="submit_draft_btn",
                          use_container_width=True, type="primary"):
                rows = conn.execute("SELECT * FROM draft_consumption WHERE session_key = ?",
                                     (st.session_state["_session_key"],)).fetchall()
                for r in rows:
                    conn.execute(
                        "INSERT INTO consumption_log (entry_date, equipment_tag, lining_system_code, "
                        "lining_system_name, sqm_completed, material_code, material_name, uom, "
                        "expected_qty, consumed_qty, variance_pct, variance_status, notes, submitted_at) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (r["entry_date"], r["equipment_tag"], r["lining_system_code"],
                         r["lining_system_name"], r["sqm_completed"], r["material_code"],
                         r["material_name"], r["uom"], r["expected_qty"], r["effective_qty"],
                         r["variance_pct"], r["variance_status"], r["notes"],
                         datetime.now().isoformat()))
                    conn.execute("UPDATE inventory SET available_qty = MAX(0, available_qty - ?) "
                                 "WHERE material_code = ?",
                                 (r["effective_qty"], r["material_code"]))
                # sqm_progress increment per (tag, code)
                grp = (pd.DataFrame([dict(r) for r in rows])
                         .groupby(["equipment_tag","lining_system_code"])["sqm_completed"].max())
                for (etag, ecode), sq in grp.items():
                    conn.execute("UPDATE sqm_progress SET done_sqm = done_sqm + ? "
                                 "WHERE equipment_tag = ? AND lining_system_code = ?",
                                 (float(sq), etag, ecode))
                conn.execute("DELETE FROM draft_consumption WHERE session_key = ?",
                             (st.session_state["_session_key"],))
                _commit_and_refresh(conn)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── Consumption History editor ───────────
    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("📜 View & Edit Consumption History", expanded=False):
        hist = pd.read_sql(
            "SELECT id, entry_date, equipment_tag, lining_system_code, lining_system_name, "
            "sqm_completed, material_code, material_name, uom, expected_qty, consumed_qty, "
            "variance_pct, variance_status, notes, submitted_at "
            "FROM consumption_log ORDER BY id DESC LIMIT 200", conn)
        if hist.empty:
            st.info("No consumption logged yet.")
        else:
            f1, f2 = st.columns(2)
            with f1:
                f_tg = st.selectbox("Filter by Equipment",
                                      ["All"] + sorted(hist["equipment_tag"].unique().tolist()),
                                      key="log_tag")
            with f2:
                f_dt = st.selectbox("Filter by Date",
                                      ["All"] + sorted(hist["entry_date"].unique().tolist(), reverse=True),
                                      key="log_date")
            view = hist.copy()
            if f_tg != "All": view = view[view["equipment_tag"] == f_tg]
            if f_dt != "All": view = view[view["entry_date"]    == f_dt]
            view.insert(0, "☐ Select", False)
            view.insert(1, "Sl. No.", range(1, len(view)+1))
            edited = st.data_editor(
                view, num_rows="fixed", hide_index=True, use_container_width=True,
                height=min(600, 50 + len(view)*35), key="cons_log_editor",
                column_config={
                    "☐ Select": st.column_config.CheckboxColumn(),
                    "Sl. No.":  st.column_config.NumberColumn(disabled=True),
                    "id":         st.column_config.NumberColumn(disabled=True),
                    "submitted_at": st.column_config.TextColumn(disabled=True),
                    "equipment_tag": st.column_config.TextColumn(disabled=True),
                    "lining_system_code": st.column_config.TextColumn(disabled=True),
                    "material_code": st.column_config.TextColumn(disabled=True),
                },
            )
            b1, b2, b3 = st.columns([2, 2, 4])
            with b1:
                if st.button("💾 Save Edits", key="cons_log_save", use_container_width=True):
                    for _, r in edited.iterrows():
                        conn.execute(
                            "UPDATE consumption_log SET entry_date=?, sqm_completed=?, "
                            "consumed_qty=?, notes=? WHERE id=?",
                            (r["entry_date"], float(r["sqm_completed"]),
                             float(r["consumed_qty"]), r["notes"], int(r["id"])))
                    _commit_and_refresh(conn)
            with b2:
                if st.button("🗑️ Delete Selected", key="cons_log_del", use_container_width=True):
                    checked = edited[edited["☐ Select"] == True]
                    for pkv in checked["id"]:
                        conn.execute("DELETE FROM consumption_log WHERE id = ?", (int(pkv),))
                    _commit_and_refresh(conn)
            with b3:
                st.download_button(
                    "⬇ Download Consumption Log",
                    excel_bytes(hist, "Consumption Log", "dashboard"),
                    file_name=f"consumption_log_{date.today():%Y%m%d}.xlsx",
                    key="cons_log_dl", use_container_width=True,
                )
    conn.close()


def _inv_receipts():
    sec_header("📦 Record Material Receipt — New Stock Received")

    # ── Flash success after rerun ─────────
    if st.session_state.get("_rc_flash"):
        st.success(st.session_state.pop("_rc_flash"))

    conn = get_db()
    open_orders = pd.read_sql(
        "SELECT DISTINCT order_id FROM orders_log WHERE status != 'Fulfilled' ORDER BY id DESC",
        conn)["order_id"].tolist()
    ord_id = st.selectbox("🔗 Link to Order ID / PR# (Optional)",
                           ["— None —"] + open_orders, key="rc_order_id")

    if ord_id != "— None —":
        ord_mats = pd.read_sql(
            "SELECT DISTINCT material_code, material_name FROM orders_log WHERE order_id = ?",
            conn, params=(ord_id,))
        mat_opts = ord_mats["material_code"].tolist()
        fmt = lambda c: f'{c}  —  {ord_mats[ord_mats["material_code"]==c]["material_name"].iloc[0][:35]}'
    else:
        mat_opts = inv["Material_Code"].tolist()
        fmt = lambda c: f'{c}  —  {inv[inv["Material_Code"]==c]["Material_Name"].iloc[0][:35]}'
    mat_code = st.selectbox("🧪 Select Material", [""] + mat_opts,
                             format_func=lambda c: "— Select Material —" if c=="" else fmt(c),
                             key="rc_mat_sel")

    # Read-only displays via st.markdown so they always reflect live DB
    if mat_code:
        mat_row = inv[inv["Material_Code"] == mat_code].iloc[0]
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(
                f'<div class="label-row">Current Available Qty</div>'
                f'<div style="background:var(--bg-inset);border:1px solid var(--border-mid);'
                f'border-radius:8px;padding:9px 12px;font-family:JetBrains Mono;'
                f'color:var(--green);font-size:13px;font-weight:700;">'
                f'{mat_row["Available_Qty"]:,.3f} {mat_row["UOM"]}</div>',
                unsafe_allow_html=True)
        with c2:
            st.markdown(
                f'<div class="label-row">Current Ordered Qty</div>'
                f'<div style="background:var(--bg-inset);border:1px solid var(--border-mid);'
                f'border-radius:8px;padding:9px 12px;font-family:JetBrains Mono;'
                f'color:var(--amber);font-size:13px;font-weight:700;">'
                f'{mat_row["Ordered_Qty"]:,.3f} {mat_row["UOM"]}</div>',
                unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

    with st.form("receipt_form", clear_on_submit=True):
        cc1, cc2, cc3 = st.columns(3)
        rdate = cc1.date_input("📅 Receipt Date", value=date.today())
        qty   = cc2.number_input("Qty Received", min_value=0.0, value=0.0,
                                  step=1.0, format="%.3f")
        notes = cc3.text_input("Notes / PO Ref.", placeholder="PO number, supplier…")
        submitted = st.form_submit_button("✅  Record Receipt", type="primary")

    if submitted:
        if not mat_code:
            st.error("Please select a material first.")
        elif qty <= 0:
            st.error("Quantity must be greater than 0.")
        else:
            mat_name = inv[inv["Material_Code"] == mat_code]["Material_Name"].iloc[0]
            mat_uom  = inv[inv["Material_Code"] == mat_code]["UOM"].iloc[0]
            conn.execute(
                "INSERT INTO receipt_log (entry_date, material_code, material_name, uom, "
                "received_qty, notes, order_id, submitted_at) "
                "VALUES (?,?,?,?,?,?,?,?)",
                (rdate.isoformat(), mat_code, mat_name, mat_uom, qty, notes,
                 ord_id if ord_id != "— None —" else None, datetime.now().isoformat()))
            conn.execute("UPDATE inventory SET available_qty = available_qty + ?, "
                         "ordered_qty = MAX(0, ordered_qty - ?) WHERE material_code = ?",
                         (qty, qty, mat_code))
            if ord_id != "— None —":
                conn.execute("UPDATE orders_log SET fulfilled_qty = fulfilled_qty + ?, "
                             "status = CASE WHEN fulfilled_qty + ? >= ordered_qty THEN 'Fulfilled' "
                             "ELSE 'Partial' END WHERE order_id = ? AND material_code = ?",
                             (qty, qty, ord_id, mat_code))
            st.session_state["_rc_flash"] = (
                f"✅ Recorded receipt: {qty:,.3f} {mat_uom} of {mat_code} ({mat_name})."
            )
            _commit_and_refresh(conn)

    # ── History editor ───────────
    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("📜 View & Edit Receipt History", expanded=False):
        hist = pd.read_sql(
            "SELECT id, entry_date, material_code, material_name, uom, received_qty, "
            "notes, order_id, submitted_at FROM receipt_log ORDER BY id DESC LIMIT 100", conn)
        if hist.empty:
            st.info("No receipts recorded yet.")
        else:
            f1, f2 = st.columns(2)
            with f1:
                f_mat = st.selectbox("Filter by Material",
                                       ["All"] + hist["material_code"].unique().tolist(),
                                       key="rlog_mat")
            with f2:
                f_dt = st.selectbox("Filter by Date",
                                      ["All"] + sorted(hist["entry_date"].unique().tolist(), reverse=True),
                                      key="rlog_date")
            view = hist.copy()
            if f_mat != "All": view = view[view["material_code"] == f_mat]
            if f_dt  != "All": view = view[view["entry_date"]    == f_dt]
            view.insert(0, "☐ Select", False)
            view.insert(1, "Sl. No.", range(1, len(view)+1))
            edited = st.data_editor(
                view, num_rows="fixed", hide_index=True, use_container_width=True,
                height=min(500, 50 + len(view)*35), key="receipt_log_editor",
                column_config={
                    "☐ Select": st.column_config.CheckboxColumn(),
                    "Sl. No.":  st.column_config.NumberColumn(disabled=True),
                    "id":        st.column_config.NumberColumn(disabled=True),
                    "submitted_at": st.column_config.TextColumn(disabled=True),
                    "material_code": st.column_config.TextColumn(disabled=True),
                },
            )
            b1, b2, b3 = st.columns([2, 2, 4])
            with b1:
                if st.button("💾 Save Edits", key="rlog_save", use_container_width=True):
                    for _, r in edited.iterrows():
                        conn.execute(
                            "UPDATE receipt_log SET entry_date=?, received_qty=?, notes=? WHERE id=?",
                            (r["entry_date"], float(r["received_qty"]), r["notes"], int(r["id"])))
                    _commit_and_refresh(conn)
            with b2:
                if st.button("🗑️ Delete Selected", key="rlog_del", use_container_width=True):
                    checked = edited[edited["☐ Select"] == True]
                    for pkv in checked["id"]:
                        conn.execute("DELETE FROM receipt_log WHERE id = ?", (int(pkv),))
                    _commit_and_refresh(conn)
            with b3:
                st.download_button(
                    "⬇ Download Receipt Log",
                    excel_bytes(hist, "Receipt Log", "dashboard"),
                    file_name=f"receipt_log_{date.today():%Y%m%d}.xlsx",
                    key="rlog_dl", use_container_width=True,
                )
    conn.close()


def _inv_ordered():
    sec_header("📋 Order Management System")

    # Flash after a successful submit/update so the user sees confirmation
    if st.session_state.get("_ord_flash"):
        st.success(st.session_state.pop("_ord_flash"))

    sub = st.radio("Generate From",
        ["📍 Location Needs", "⚙️ System Code Needs", "🔎 Lookup by PR#", "📜 View Past Orders"],
        horizontal=True, key="ord_mode", label_visibility="collapsed")
    st.markdown("<hr>", unsafe_allow_html=True)
    conn = get_db()

    if sub.startswith("📍"):
        loc = st.selectbox("Select Location", ["All Locations"] + LOCATION_ORDER, key="ord_loc_sel")
        if st.button("🔍 Calculate Shortfalls", key="calc_ord_short"):
            tags = (eq_master["Equipment_Tag_No."].tolist() if loc == "All Locations"
                    else eq_master[eq_master["Location"] == loc]["Equipment_Tag_No."].tolist())
            alloc = cascade_allocate(tags)
            mat_agg = alloc_to_mat_df(alloc)
            short = mat_agg[mat_agg["Shortfall"] > 0]
            st.session_state["_ord_draft"] = short.to_dict(orient="records")
            st.session_state["_ord_source_detail"] = loc

    elif sub.startswith("⚙️"):
        codes_avail = sorted(equip_sc["Lining_System_Code"].unique(), key=lambda x: int(x))
        ccode = st.selectbox("Select System Code", codes_avail,
            format_func=lambda c: f"Code {c} — {equip_sc[equip_sc['Lining_System_Code']==c]['Lining_System_Short_Name'].iloc[0]}",
            key="ord_code_sel")
        if st.button("🔍 Calculate Shortfalls", key="calc_ord_sc_short"):
            tags = equip_sc[equip_sc["Lining_System_Code"] == ccode]["Equipment_Tag_No."].unique().tolist()
            alloc = cascade_allocate(tags)
            alloc = alloc[alloc["Lining_System_Code"] == ccode]
            mat_agg = alloc_to_mat_df(alloc)
            short = mat_agg[mat_agg["Shortfall"] > 0]
            st.session_state["_ord_draft"] = short.to_dict(orient="records")
            st.session_state["_ord_source_detail"] = f"Code {ccode}"

    elif sub.startswith("🔎"):
        # ── Lookup by PR# ──────────────
        st.markdown('<div class="glass-flat">', unsafe_allow_html=True)
        pr_q = st.text_input("🔎 Lookup by PR# (Purchase Request Number)",
                              placeholder="Type a PR# to find its order…", key="pr_lookup")
        st.markdown("</div>", unsafe_allow_html=True)
        if pr_q:
            df = pd.read_sql(
                "SELECT order_id, pr_number, material_code, material_name, uom, ordered_qty, "
                "fulfilled_qty, (ordered_qty - fulfilled_qty) AS pending_qty, status, notes "
                "FROM orders_log WHERE pr_number LIKE ? ORDER BY id DESC",
                conn, params=(f"%{pr_q}%",))
            if df.empty:
                st.warning(f"No orders found matching PR# '{pr_q}'.")
            else:
                st.caption(f"Found {df['order_id'].nunique()} order(s), {len(df)} material line(s).")
                st.dataframe(df, use_container_width=True, hide_index=True,
                             height=min(500, 60 + len(df)*35))
                st.download_button(
                    f"⬇ Download Lookup Results",
                    excel_bytes(df, f"PR# Lookup {pr_q}", "execution"),
                    file_name=f"pr_lookup_{pr_q}_{date.today():%Y%m%d}.xlsx",
                    key="ord_dl_pr",
                )

    else:
        # ── View Past Orders (with PR# inline editor) ──────────────
        past_ids = pd.read_sql("SELECT DISTINCT order_id FROM orders_log ORDER BY id DESC",
                                conn)["order_id"].tolist()
        if past_ids:
            ord_id = st.selectbox("Select Order ID", past_ids, key="view_past_ord")
            # Inline PR# editor
            cur_pr_row = conn.execute(
                "SELECT pr_number FROM orders_log WHERE order_id = ? LIMIT 1", (ord_id,)
            ).fetchone()
            cur_pr = (cur_pr_row["pr_number"] or "") if cur_pr_row else ""
            pc1, pc2 = st.columns([3, 1])
            with pc1:
                new_pr = st.text_input("🔖 PR# (editable)", value=cur_pr,
                                         placeholder="Enter PR# for future reference…",
                                         key="pr_input")
            with pc2:
                st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
                if st.button("💾 Update PR#", key="update_pr_btn", use_container_width=True):
                    conn.execute("UPDATE orders_log SET pr_number = ? WHERE order_id = ?",
                                  (new_pr.strip() or None, ord_id))
                    st.session_state["_ord_flash"] = f"✅ PR# updated for {ord_id}."
                    _commit_and_refresh(conn)

            df = pd.read_sql(
                "SELECT order_id, pr_number, material_code, material_name, uom, ordered_qty, "
                "fulfilled_qty, (ordered_qty - fulfilled_qty) AS pending_qty, status, notes "
                "FROM orders_log WHERE order_id = ?", conn, params=(ord_id,))
            st.dataframe(df, use_container_width=True, hide_index=True,
                         height=min(500, 60 + len(df)*35))
            st.download_button(
                f"⬇ Download Order Status — {ord_id}",
                excel_bytes(df, f"Order {ord_id}", "execution"),
                file_name=f"order_status_{ord_id}.xlsx", key="ord_dl_past",
            )
        else:
            st.info("No past orders.")

    # Draft grid
    if sub.startswith(("📍","⚙️")) and st.session_state.get("_ord_draft"):
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        sec_header(f"✏️ Review & Edit Order List — {st.session_state['_ord_source_detail']}")
        draft_df = pd.DataFrame(st.session_state["_ord_draft"])
        edited = st.data_editor(
            draft_df[["Code","Material Name","UOM","Available","Shortfall","On Order"]].assign(
                **{"Order Qty": draft_df["Shortfall"], "Notes": "", "Remove": False}),
            num_rows="dynamic", use_container_width=True, hide_index=True,
            key="ord_draft_editor",
            column_config={"Remove": st.column_config.CheckboxColumn("☑ Remove"),
                            "Available": st.column_config.NumberColumn(disabled=True, format="%.3f"),
                            "Shortfall": st.column_config.NumberColumn(disabled=True, format="%.3f")},
            height=min(500, 60 + len(draft_df)*35),
        )
        # Optional PR# at submit time
        pr_at_submit = st.text_input(
            "🔖 PR# (optional — can also be added later via 'View Past Orders')",
            placeholder="Leave blank if PR# isn't issued yet…", key="ord_pr_at_submit"
        )
        if st.button("📤 Submit Order", key="submit_ord_btn", type="primary"):
            new_id = _next_order_id(conn)
            for _, r in edited.iterrows():
                if bool(r.get("Remove", False)) or float(r.get("Order Qty", 0)) <= 0: continue
                conn.execute(
                    "INSERT INTO orders_log (order_id, order_date, generated_by, "
                    "generation_source, source_detail, material_code, material_name, uom, "
                    "ordered_qty, fulfilled_qty, status, notes, submitted_at, pr_number) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (new_id, date.today().isoformat(), "admin",
                     "Location Needs" if st.session_state["_ord_source_detail"] in LOCATION_ORDER + ["All Locations"]
                     else "System Code Needs",
                     st.session_state["_ord_source_detail"],
                     r["Code"], r["Material Name"], r.get("UOM",""),
                     float(r["Order Qty"]), 0.0, "Pending", r.get("Notes",""),
                     datetime.now().isoformat(), pr_at_submit.strip() or None))
                conn.execute("UPDATE inventory SET ordered_qty = ordered_qty + ? "
                             "WHERE material_code = ?", (float(r["Order Qty"]), r["Code"]))
            st.session_state["_ord_draft"] = []
            pr_msg = f" (PR# {pr_at_submit.strip()})" if pr_at_submit.strip() else ""
            st.session_state["_ord_flash"] = f"✅ Order {new_id} submitted{pr_msg}."
            _commit_and_refresh(conn)
        st.markdown("</div>", unsafe_allow_html=True)
    conn.close()


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 6 — MASTER DATA
# ═══════════════════════════════════════════════════════════════════════════════
def tab_master_data():
    if not db_available():
        st.error("Database not found. Run `python setup_db.py` first.")
        return
    table_display = submode_radio("Select Table",
        ["Equipment", "LINING SYSTEM MATERIAL CONSM", "Materials_DetailsAvailable_Qty"],
        key="md_table_radio")
    sec_header("🗄️ Master Data — View, Add & Delete Records")
    db_table = {"Equipment":"equipment",
                "LINING SYSTEM MATERIAL CONSM":"recipe",
                "Materials_DetailsAvailable_Qty":"inventory"}[table_display]
    st.markdown("<hr>", unsafe_allow_html=True)

    conn = get_db()
    schema = conn.execute(f"PRAGMA table_info({db_table})").fetchall()
    cols = [c["name"] for c in schema]
    pk_col = next((c["name"] for c in schema if c["pk"]), "id")

    # ── ADD NEW ROW ────────────
    st.markdown('<div class="glass">', unsafe_allow_html=True)
    sec_header("➕ Add New Row")
    if db_table == "equipment":
        code_pairs = (recipe[["Lining_System_Code","Lining_System_Short_Name"]]
                      .drop_duplicates().sort_values("Lining_System_Code"))
        code_opts = [f'Code {r["Lining_System_Code"]} – {r["Lining_System_Short_Name"]}'
                     for _, r in code_pairs.iterrows()]
        sel_codes = st.multiselect("🔧 Select Lining System Code(s) *",
                                     code_opts, key="seq_codes_pre")
        if sel_codes:
            with st.form("smart_eq_form"):
                sec_header("Equipment Identity")
                cc1, cc2 = st.columns(2)
                tag_v = cc1.text_input("🏷️ Equipment Tag No. *", placeholder="e.g. V-1001",
                                         key="seq_tag")
                loc_v = cc2.selectbox("📍 Location *", LOCATION_ORDER, key="seq_loc")
                sec_header("Equipment Details (shared across all codes)")
                shared = {}
                shared["name"]      = st.text_input("Equipment Name", key="seq_sh_name")
                cc1, cc2, cc3 = st.columns(3)
                shared["type"]       = cc1.text_input("Type", key="seq_sh_type")
                shared["substrate"]  = cc2.text_input("Substrate", key="seq_sh_substrate")
                shared["material_spec"] = cc3.text_input("Material Spec.", key="seq_sh_spec")
                cc1, cc2 = st.columns(2)
                shared["design"]         = cc1.text_input("Design", key="seq_sh_design")
                shared["lining_systems"] = cc2.text_input("Lining System+", key="seq_sh_ls")
                sec_header("Per Lining System Code")
                sqm_per = {}
                for cl in sel_codes:
                    code = cl.split(" ")[1]
                    sqm_per[code] = st.number_input(f"Surface Area SQM * (Code {code})",
                                                     min_value=0.0, step=0.5, format="%.2f",
                                                     key=f"seq_sqm_{code}")
                if st.form_submit_button("💾 Save Equipment", type="primary"):
                    if not (tag_v and all(v > 0 for v in sqm_per.values())):
                        st.error("Tag and per-code SQM required.")
                    else:
                        for cl in sel_codes:
                            code = cl.split(" ")[1]
                            sname = code_pairs[code_pairs["Lining_System_Code"]==code]["Lining_System_Short_Name"].iloc[0]
                            conn.execute(
                                "INSERT INTO equipment (location, type, lining_system_code, "
                                "lining_system_short_name, lining_type, equipment_tag, name, "
                                "substrate, material_spec, design, surface_area_sqm, lining_systems) "
                                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                                (loc_v, shared["type"], code, sname, "",
                                 tag_v, shared["name"], shared["substrate"],
                                 shared["material_spec"], shared["design"],
                                 sqm_per[code], shared["lining_systems"]))
                            conn.execute(
                                "INSERT OR REPLACE INTO sqm_progress "
                                "(equipment_tag, lining_system_code, original_sqm, done_sqm) "
                                "VALUES (?,?,?,COALESCE((SELECT done_sqm FROM sqm_progress "
                                "WHERE equipment_tag=? AND lining_system_code=?),0))",
                                (tag_v, code, sqm_per[code], tag_v, code))
                        _commit_and_refresh(conn)
    else:
        # Skip cols with spaces/special chars (legacy artifacts of Excel ingestion)
        clean_cols = [c for c in cols
                      if c != pk_col
                      and " " not in c and "#" not in c and "/" not in c
                      and not c.startswith("Unnamed")]
        with st.form(f"dyn_{db_table}_form"):
            inputs = {}
            grid = st.columns(3)
            for i, c in enumerate(clean_cols):
                with grid[i % 3]:
                    if any(k in c.lower() for k in ("qty","sqm","amount","value","price","for_1")):
                        inputs[c] = st.number_input(c, min_value=0.0, value=0.0, step=0.001,
                                                     format="%.4f", key=f"dyn_{db_table}_{c}")
                    else:
                        inputs[c] = st.text_input(c, key=f"dyn_{db_table}_{c}")
            if st.form_submit_button(f"➕ Add Row to {table_display}", type="primary"):
                ccols = ",".join(f'"{k}"' for k in inputs.keys())
                qm = ",".join("?" * len(inputs))
                conn.execute(f"INSERT INTO {db_table} ({ccols}) VALUES ({qm})",
                              tuple(inputs.values()))
                _commit_and_refresh(conn)
    st.markdown("</div>", unsafe_allow_html=True)

    # ── VIEW / EDIT / DELETE ────────────
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="glass">', unsafe_allow_html=True)
    sec_header(f"📋 View, Edit & Delete — {table_display}")
    sc1, sc2 = st.columns([2, 1])
    with sc1:
        q = st.text_input("🔍 Search…", key=f"md_search_{db_table}")
    with sc2:
        col_scope = st.selectbox("in column", ["All columns"] + cols, key=f"md_col_{db_table}")

    df_all = pd.read_sql(f"SELECT * FROM {db_table}", conn)
    if q:
        if col_scope == "All columns":
            mask = df_all.astype(str).apply(lambda row: row.str.contains(q, case=False, na=False).any(), axis=1)
        else:
            mask = df_all[col_scope].astype(str).str.contains(q, case=False, na=False)
        df_all = df_all[mask]

    df_show = df_all.copy()
    df_show.insert(0, "☐ Select", False)
    df_show.insert(0, "Sl. No.", range(1, len(df_show)+1))
    edited = st.data_editor(
        df_show, num_rows="fixed", hide_index=True,
        use_container_width=True, height=min(600, 50 + len(df_show)*35),
        key=f"md_editor_{db_table}",
        column_config={
            "☐ Select": st.column_config.CheckboxColumn(),
            "Sl. No.": st.column_config.NumberColumn(disabled=True),
            pk_col: st.column_config.NumberColumn(disabled=True),
        },
    )
    st.caption(f"Total entries: {len(df_all)}")

    bc1, bc2, bc3 = st.columns([2, 2, 3])
    with bc1:
        if st.button("💾 Save Cell Edits", type="primary", key=f"save_edits_{db_table}"):
            # Compare edited vs df_show — find rows where columns changed
            for i, row in edited.iterrows():
                src = df_show.iloc[i]
                changes = {c: row[c] for c in cols
                           if c != pk_col and not pd.isna(row[c]) and row[c] != src[c]}
                if changes:
                    setter = ", ".join(f'"{k}" = ?' for k in changes)
                    conn.execute(f"UPDATE {db_table} SET {setter} WHERE \"{pk_col}\" = ?",
                                  (*changes.values(), row[pk_col]))
            _commit_and_refresh(conn)
    with bc2:
        if st.button("🗑️ Delete Selected", key=f"del_checked_{db_table}"):
            checked = edited[edited["☐ Select"] == True]
            if checked.empty:
                st.warning("Check at least one row.")
            else:
                for pkv in checked[pk_col]:
                    if db_table == "equipment":
                        tg = conn.execute("SELECT equipment_tag FROM equipment WHERE id = ?", (pkv,)).fetchone()
                        if tg: conn.execute("DELETE FROM sqm_progress WHERE equipment_tag = ?", (tg["equipment_tag"],))
                    conn.execute(f"DELETE FROM {db_table} WHERE \"{pk_col}\" = ?", (int(pkv),))
                _commit_and_refresh(conn)
    with bc3:
        st.download_button(
            f"⬇ Download {table_display[:14]}",
            excel_bytes(df_all, table_display, "overview"),
            file_name=f"{db_table}_export_{date.today():%Y%m%d}.xlsx",
            key=f"md_dl_{db_table}", use_container_width=True,
        )
    st.caption("⚠️ Deletion is permanent. Equipment rows also remove the matching sqm_progress record.")
    st.markdown("</div>", unsafe_allow_html=True)
    conn.close()


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
render_sidebar()
render_header()

tabs = st.tabs([
    "📊  Dashboard",
    "🔍  Entry",
    "📋  Reports",
    "⚙️  Plan",
    "📦  Inventory",
    "🗄️  Master Data",
])
with tabs[0]: tab_dashboard()
with tabs[1]: tab_entry()
with tabs[2]: tab_reports()
with tabs[3]: tab_plan()
with tabs[4]: tab_inventory()
with tabs[5]: tab_master_data()
