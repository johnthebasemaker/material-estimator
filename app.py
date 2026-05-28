"""
Smart Material Estimator · app.py (v3)
Run: streamlit run app.py
"""
import io, os, sys
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from validate_data     import clean_inventory, clean_recipe, clean_equipment
from allocation_engine import build_demand_matrix

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Smart Material Estimator",
                   page_icon="🏗️", layout="wide",
                   initial_sidebar_state="expanded")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PATH_A   = os.path.join(BASE_DIR, "Materials_DetailsAvailable_Qty.xlsx")
PATH_B   = os.path.join(BASE_DIR, "For_1_SQM.xlsx")
PATH_C   = os.path.join(BASE_DIR, "Equipment.xlsx")
SHEET_A, SHEET_B, SHEET_C = "Materials", "LINING SYSTEM MATERIAL CONSM", "Data Input"
LOCATION_ORDER = ["Brown Field", "TRAIN J", "TRAIN K"]

# ─────────────────────────────────────────────────────────────────────────────
# STYLES
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;600;700&display=swap');
:root{
  /* Map to Streamlit's native theme to seamlessly adapt to Light/Dark mode */
  --bg0: var(--background-color);
  --bg1: var(--secondary-background-color);
  --bg2: var(--secondary-background-color);
  --bg3: var(--background-color);
  --bg4: var(--secondary-background-color);
  --border: rgba(128, 128, 128, 0.2);
  --border2: rgba(128, 128, 128, 0.3);
  
  /* Preserving accent colors */
  --amber:#F59E0B;--amber2:#FCD34D;--amber-bg:rgba(245,158,11,.15);
  --green:#10B981;--green-bg:rgba(16,185,129,.15);
  --red:#EF4444;--red-bg:rgba(239,68,68,.15);
  --orange:#F97316;--orange-bg:rgba(249,115,22,.15);
  --yellow:#EAB308;--yellow-bg:rgba(234,179,8,.15);
  --blue:#3B82F6;--blue-bg:rgba(59,130,246,.15);
  
  /* Adaptive Text Colors using CSS color-mix */
  --t0: var(--text-color);
  --t1: var(--text-color);
  --t2: color-mix(in srgb, var(--text-color) 80%, transparent);
  --t3: color-mix(in srgb, var(--text-color) 60%, transparent);
  --t4: color-mix(in srgb, var(--text-color) 45%, transparent);
  --t5: color-mix(in srgb, var(--text-color) 30%, transparent);
}
html,body,[class*="css"]{font-family:'Sora',sans-serif!important;background:var(--bg0)!important;color:var(--t1);}
.main .block-container{padding-top:0!important;padding-bottom:3rem;max-width:1480px;}

/* sidebar */
[data-testid="stSidebar"]{background:var(--bg1)!important;border-right:1px solid var(--border)!important;}
[data-testid="stSidebar"] *{font-family:'Sora',sans-serif!important;}

/* ── TRUE STICKY HEADER & TABS ── */
header[data-testid="stHeader"]{display:none!important;}
[data-testid="stAppViewContainer"]{padding-top:0!important;}

/* 1. Freeze the main title block */
.sticky-header-wrap{
    position:sticky;top:0;z-index:9999;
    background:var(--bg0);
    padding:.8rem 1.5rem .5rem;
    margin-bottom:0;
}

/* 2. Freeze the Tabs right below the header */
[data-testid="stTabs"] > div:first-of-type {
    position: sticky;
    top: 50px; /* Sits right under sticky-header-wrap */
    z-index: 9998;
    background: var(--bg0);
    padding: 0.5rem 0;
    border-bottom: 2px solid var(--border);
}

/* tabs styling */
[data-testid="stTabs"] [data-baseweb="tab-list"]{background:transparent;border-bottom:none;gap:0;padding:0;}
[data-testid="stTabs"] [data-baseweb="tab"]{font-family:'JetBrains Mono',monospace!important;font-size:.7rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--t4)!important;padding:.7rem 1.3rem;border-bottom:3px solid transparent;transition:all .2s;}
[data-testid="stTabs"] [aria-selected="true"]{color:var(--amber)!important;border-bottom:3px solid var(--amber)!important;}

/* buttons */
.stButton>button{font-family:'JetBrains Mono',monospace!important;font-size:.68rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;background:var(--amber)!important;color:#000!important;border:none!important;border-radius:4px!important;padding:.48rem 1.1rem!important;transition:all .15s!important;}
.stButton>button:hover{background:#FBBF24!important;transform:translateY(-1px);box-shadow:0 4px 14px rgba(245,158,11,.3)!important;}

/* metrics */
[data-testid="stMetric"]{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:.85rem 1rem!important;transition:all .2s;cursor:default;}
[data-testid="stMetric"]:hover{background:var(--bg3)!important;border-color:var(--amber)!important;box-shadow:0 0 0 1px var(--amber);transform:translateY(-1px);}
[data-testid="stMetricLabel"]{font-family:'JetBrains Mono',monospace!important;font-size:.6rem!important;letter-spacing:.12em;text-transform:uppercase;color:var(--t3)!important;}
[data-testid="stMetricValue"]{font-family:'JetBrains Mono',monospace!important;font-size:1.8rem!important;color:var(--t0)!important;}
[data-testid="stMetricDelta"]{font-size:.72rem!important;}

/* info boxes */
[data-testid="stInfo"]   {background:var(--blue-bg)!important;border-left:3px solid var(--blue)!important; color:var(--t1)!important;}
[data-testid="stSuccess"]{background:var(--green-bg)!important;border-left:3px solid var(--green)!important; color:var(--t1)!important;}
[data-testid="stWarning"]{background:var(--amber-bg)!important;border-left:3px solid var(--amber)!important; color:var(--t1)!important;}
[data-testid="stError"]  {background:var(--red-bg)!important;border-left:3px solid var(--red)!important; color:var(--t1)!important;}

/* expander */
[data-testid="stExpander"]{background:var(--bg2)!important;border:1px solid var(--border)!important;border-radius:6px!important;}
[data-testid="stExpander"] summary{color:var(--t1)!important;}

/* select */
[data-baseweb="select"]>div,[data-baseweb="input"]>div{background:var(--bg2)!important;border-color:var(--border2)!important;color:var(--t0)!important;}

/* dataframe */
[data-testid="stDataFrame"]{border:1px solid var(--border)!important;border-radius:6px;overflow:hidden;}
hr{border-color:var(--border)!important;margin:.8rem 0!important;}

/* ── Custom components ── */
.sec-hdr{font-family:'JetBrains Mono',monospace;font-size:.6rem;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:var(--t5);border-bottom:1px solid var(--border);padding-bottom:.3rem;margin-bottom:.8rem;}
.card{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:1rem 1.2rem;margin-bottom:.6rem;}
.card-amber{border-left:4px solid var(--amber);}
.card-green{border-left:4px solid var(--green);}
.card-blue {border-left:4px solid var(--blue);}
.loc-badge{display:inline-block;font-family:'JetBrains Mono',monospace;font-size:.62rem;font-weight:700;letter-spacing:.07em;text-transform:uppercase;padding:.18rem .6rem;border-radius:3px;}
.loc-bf{background:rgba(59,130,246,.12);color:var(--blue);}
.loc-tj{background:rgba(245,158,11,.12);color:var(--amber);}
.loc-tk{background:rgba(16,185,129,.12);color:var(--green);}
.pill{display:inline-block;font-family:'JetBrains Mono',monospace;font-size:.68rem;font-weight:600;padding:.15rem .5rem;border-radius:20px;}
.pill-g{background:var(--green-bg);color:var(--green);}
.pill-y{background:var(--yellow-bg);color:var(--yellow);}
.pill-o{background:var(--orange-bg);color:var(--orange);}
.pill-r{background:var(--red-bg);color:var(--red);}
.tag-chip{display:inline-block;font-family:'JetBrains Mono',monospace;font-size:.7rem;background:var(--bg3);color:var(--amber);border:1px solid var(--border2);border-radius:4px;padding:.15rem .5rem;margin:.12rem;}
.syscode-block{background:var(--bg3);border:1px solid var(--border2);border-radius:6px;padding:.7rem .9rem;margin:.35rem 0;}
.syscode-hdr{display:flex;align-items:center;gap:.7rem;margin-bottom:.5rem;}
.code-badge{font-family:'JetBrains Mono',monospace;font-size:.68rem;font-weight:700;background:var(--bg4);color:var(--amber);border:1px solid var(--amber-bg);border-radius:4px;padding:.2rem .55rem;}
.session-equip{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:.9rem 1rem;margin-bottom:.5rem;}
.drag-handle{font-size:1rem;color:var(--t5);cursor:grab;user-select:none;padding:.2rem .4rem;}
.grand-box{background:linear-gradient(135deg, var(--amber-bg) 0%, var(--bg2) 70%);border:1px solid var(--border2);border-left:4px solid var(--amber);border-radius:8px;padding:1.1rem 1.4rem;}
.status-dot-g::before{content:"●";color:var(--green);margin-right:.4rem;}
.status-dot-o::before{content:"●";color:var(--orange);margin-right:.4rem;}
.status-dot-y::before{content:"●";color:var(--yellow);margin-right:.4rem;}
.status-dot-r::before{content:"●";color:var(--red);margin-right:.4rem;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADER
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Loading project data…")
def load_all():
    df_a_raw = pd.read_excel(PATH_A, sheet_name=SHEET_A)
    df_b_raw = pd.read_excel(PATH_B, sheet_name=SHEET_B)
    df_c_raw = pd.read_excel(PATH_C, sheet_name=SHEET_C)

    inv    = clean_inventory(df_a_raw)
    recipe = clean_recipe(df_b_raw)
    equip  = clean_equipment(df_c_raw)

    # ── (Tag, SysCode) SQM totals — sums multi-area rows correctly ────────
    equip_sc = equip.groupby(
        ["Equipment_Tag_No.", "Lining_System_Code", "Lining_System_Short_Name"],
        as_index=False
    )["Surface_Area_SQM"].sum().rename(columns={"Surface_Area_SQM": "Total_SQM"})

    # ── Detailed demand matrix: (tag, sys_code, material) ─────────────────
    dm = equip_sc.merge(recipe, on="Lining_System_Code", suffixes=("_e", "_r"))
    dm["Demand_Qty"] = dm["For_1_SQM"] * dm["Total_SQM"]
    if "Lining_System_Short_Name_e" in dm.columns:
        dm = dm.rename(columns={"Lining_System_Short_Name_e": "Lining_System_Short_Name"})
        dm.drop(columns=["Lining_System_Short_Name_r"], inplace=True, errors="ignore")
    dm = dm[["Equipment_Tag_No.", "Lining_System_Code", "Lining_System_Short_Name",
             "Total_SQM", "Material_Code", "Material_Name", "UOM", "Demand_Qty"]]

    # ── Equipment master (one row per tag) ─────────────────────────────────
    raw = df_c_raw.copy()
    raw.columns = raw.columns.str.strip()
    raw = raw.dropna(subset=["Equipment_Tag_No.", "Lining_System_Code"])
    raw["Equipment_Tag_No."] = raw["Equipment_Tag_No."].astype(str).str.strip()
    raw["Location"] = raw["Location"].astype(str).str.strip()
    raw["Type"]     = raw["Type"].astype(str).str.strip()
    raw["Surface_Area_SQM"] = pd.to_numeric(raw["Surface_Area_SQM"], errors="coerce")
    eq_master = raw.groupby("Equipment_Tag_No.", as_index=False).agg(
        Name          =("Name",          "first"),
        Description   =("Description",   "first"),
        Location      =("Location",      "first"),
        Type          =("Type",          "first"),
        Lining_Systems=("Lining_System+","first"),
        Material_Spec =("Material Spec.","first"),
        Design        =("Design",        "first"),
        Total_SQM     =("Surface_Area_SQM","sum"),
    )
    eq_master["Location"] = eq_master["Location"].str.strip()

    return inv, recipe, equip_sc, dm, eq_master


inv, recipe, equip_sc, dm, eq_master = load_all()
ALL_TAGS      = sorted(eq_master["Equipment_Tag_No."].tolist())
INV_POOL_INIT = inv.set_index("Material_Code")["Available_Qty"].to_dict()

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────
if "session_tags" not in st.session_state:
    st.session_state.session_tags = []

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def loc_badge(loc):
    cls = {"Brown Field":"loc-bf","TRAIN J":"loc-tj","TRAIN K":"loc-tk"}.get(loc,"loc-bf")
    return f'<span class="loc-badge {cls}">{loc}</span>'

def status_dot(pct):
    if pct >= 100: return "status-dot-g"
    if pct >= 90:  return "status-dot-o"
    if pct >= 80:  return "status-dot-y"
    return "status-dot-r"

def fulfil_pill(pct):
    cls = "pill-g" if pct>=100 else "pill-o" if pct>=90 else "pill-y" if pct>=80 else "pill-r"
    return f'<span class="pill {cls}">{pct:.1f}%</span>'

def cascade_allocate(tag_order: list[str]) -> pd.DataFrame:
    """
    Cascade inventory pool through equipment in order.
    Pool is GLOBAL per material (not per system code).
    Returns df with columns:
      Equipment_Tag_No. | Lining_System_Code | Lining_System_Short_Name |
      Total_SQM | Material_Code | Material_Name | UOM |
      Demand_Qty | Allocated_Qty | Shortfall_Qty | Fulfillment_Pct |
      Pool_Before | Pool_After
    """
    pool = dict(INV_POOL_INIT)  # mutable copy
    rows = []
    for tag in tag_order:
        tag_dm = dm[dm["Equipment_Tag_No."] == tag].copy()
        # Process system codes in numeric order for consistency
        for code in sorted(tag_dm["Lining_System_Code"].unique(), key=lambda x: int(x)):
            code_rows = tag_dm[tag_dm["Lining_System_Code"] == code]
            for _, r in code_rows.iterrows():
                mat     = r["Material_Code"]
                demand  = r["Demand_Qty"]
                before  = pool.get(mat, 0.0)
                alloc   = min(demand, before)
                short   = demand - alloc
                after   = max(0.0, before - alloc)
                pool[mat] = after
                rows.append({
                    "Equipment_Tag_No.":       tag,
                    "Lining_System_Code":      code,
                    "Lining_System_Short_Name": r["Lining_System_Short_Name"],
                    "Total_SQM":               r["Total_SQM"],
                    "Material_Code":           mat,
                    "Material_Name":           r["Material_Name"],
                    "UOM":                     r["UOM"],
                    "Demand_Qty":              round(demand, 4),
                    "Allocated_Qty":           round(alloc, 4),
                    "Shortfall_Qty":           round(short, 4),
                    "Pool_Before":             round(before, 4),
                    "Pool_After":              round(after, 4),
                })
    result = pd.DataFrame(rows)
    if not result.empty:
        result["Fulfillment_Pct"] = (
            result["Allocated_Qty"] / result["Demand_Qty"].replace(0, np.nan) * 100
        ).fillna(100).clip(0, 100).round(2)
    return result

def tag_fulfillment(alloc_df: pd.DataFrame, tag: str) -> float:
    t = alloc_df[alloc_df["Equipment_Tag_No."] == tag]
    if t.empty: return 0.0
    d = t["Demand_Qty"].sum()
    a = t["Allocated_Qty"].sum()
    return min(100.0, a / d * 100) if d > 0 else 100.0

def syscode_fulfillment(alloc_df: pd.DataFrame, tag: str, code: str) -> float:
    t = alloc_df[(alloc_df["Equipment_Tag_No."]==tag)&(alloc_df["Lining_System_Code"]==code)]
    if t.empty: return 0.0
    d = t["Demand_Qty"].sum()
    a = t["Allocated_Qty"].sum()
    return min(100.0, a / d * 100) if d > 0 else 100.0

def excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

def plotly_mat_table(df: pd.DataFrame, key_suffix: str, height: int = 380) -> None:
    """Renders a colour-coded native Streamlit dataframe that perfectly adapts to Light/Dark mode."""
    cols_show = ["Material_Code", "Material_Name", "UOM",
                 "Demand_Qty", "Allocated_Qty", "Shortfall_Qty", "Fulfillment_Pct"]
    
    # Copy data and rename columns FIRST to avoid Streamlit column_config conflicts
    df2 = df[cols_show].copy()
    clean_names = {
        "Material_Code": "Code",
        "Material_Name": "Material Name",
        "UOM": "UOM",
        "Demand_Qty": "Demand",
        "Allocated_Qty": "Allocated",
        "Shortfall_Qty": "Shortfall",
        "Fulfillment_Pct": "Fulfil %"
    }
    df2 = df2.rename(columns=clean_names)

    # 1. Row-by-Row Styling Logic
    def style_row(row):
        pct = row["Fulfil %"]
        if pd.isna(pct): 
            pct = 100.0
            
        if pct >= 100:
            bg = "rgba(16, 185, 129, 0.12)"
            tc = "#10B981" # Emerald Green
        elif pct >= 90:
            bg = "rgba(249, 115, 22, 0.12)"
            tc = "#F97316" # Orange
        elif pct >= 80:
            bg = "rgba(234, 179, 8, 0.12)"
            tc = "#EAB308" # Yellow
        else:
            bg = "rgba(239, 68, 68, 0.12)"
            tc = "#EF4444" # Red
            
        # Apply background to all columns, but add explicit text color/weight ONLY to the last column
        styles = [f"background-color: {bg}"] * len(df2.columns)
        styles[-1] = f"background-color: {bg}; color: {tc}; font-weight: 700;"
        return styles

    # 2. Apply styles and format numbers using the NEW column names
    styled_df = df2.style.apply(style_row, axis=1)
    
    styled_df = styled_df.format({
        "Demand": "{:,.3f}",
        "Allocated": "{:,.3f}",
        "Shortfall": "{:,.3f}",
        "Fulfil %": "{:.1f}%"
    })

    # 3. Render natively (No column_config allowed when passing a Styler object)
    st.dataframe(
        styled_df,
        hide_index=True,
        use_container_width=True,
        height=height,
        key=f"tbl_{key_suffix}"
    )



# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding:.5rem 0 1.2rem">
      <div style="font-family:'JetBrains Mono',monospace;font-size:1rem;
                  font-weight:700;color:#F59E0B;">🏗 SME</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:.56rem;
                  letter-spacing:.18em;text-transform:uppercase;color:var(--t5);margin-top:2px;">
        Smart Material Estimator v3</div>
    </div>""", unsafe_allow_html=True)

    st.markdown('<div class="sec-hdr">📍 Project Overview</div>', unsafe_allow_html=True)
    loc_counts = eq_master.groupby("Location")["Equipment_Tag_No."].count()
    for loc in LOCATION_ORDER:
        cnt   = loc_counts.get(loc, 0)
        badge = loc_badge(loc)
        st.markdown(
            f'<div style="display:flex;justify-content:space-between;'
            f'align-items:center;margin-bottom:.35rem;">'
            f'{badge}'
            f'<span style="font-family:\'JetBrains Mono\',monospace;'
            f'font-size:.75rem;color:var(--t3);">{cnt} equip.</span></div>',
            unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="sec-hdr">📦 Inventory</div>', unsafe_allow_html=True)
    st.caption(f"📦 {len(inv)} materials  ·  "
               f"⚠️ {(inv['Available_Qty']==0).sum()} at zero stock")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="sec-hdr">📋 Session</div>', unsafe_allow_html=True)
    n_sess = len(st.session_state.session_tags)
    if n_sess:
        for t in st.session_state.session_tags:
            st.caption(f"  · {t}")
        if st.button("🗑 Clear Session", key="clear_sidebar"):
            st.session_state.session_tags = []
            st.rerun()
    else:
        st.caption("No equipment added yet.")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""<div style="font-family:'JetBrains Mono',monospace;font-size:.6rem;
        letter-spacing:.08em;color:var(--t5);">
        🟢 100%  🟠 90–99%  🟡 80–89%  🔴 &lt;80%</div>""",
        unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# STICKY HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="sticky-header-wrap">
  <div style="display:flex;align-items:baseline;gap:.9rem;margin-bottom:.35rem;">
    <span style="font-family:'JetBrains Mono',monospace;font-size:1.2rem;
                 font-weight:700;color:var(--t0);letter-spacing:-.01em;">
      🏗 Smart Material Estimator</span>
    <span style="font-size:.72rem;color:var(--t4);letter-spacing:.03em;">
      System-code level · Cascading allocation · Priority-based</span>
  </div>
</div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "🔍  Equipment Entry",
    "📦  Session Order Report",
    "📍  Location Report",
    "⚙️  Execution Plan",
])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 · EQUIPMENT ENTRY
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    left, right = st.columns([1, 1.65], gap="large")

    # ── LEFT: search + session list ──────────────────────────────────────────
    with left:
        st.markdown('<div class="sec-hdr">🔍 Find Equipment</div>',
                    unsafe_allow_html=True)
        selected_tag = st.selectbox(
            "tag_search", [""] + ALL_TAGS,
            format_func=lambda t: (
                "" if t == "" else
                f"{t}  —  "
                f"{eq_master.set_index('Equipment_Tag_No.')['Name'].get(t,'')}"
            ),
            key="tag_select", label_visibility="collapsed",
        )
        ca, cb = st.columns([2,1])
        with ca:
            already = selected_tag in st.session_state.session_tags
            add_btn = st.button("＋ Add to Session", key="add_btn",
                                disabled=(selected_tag=="" or already))
        with cb:
            if already and selected_tag:
                st.markdown(
                    '<div style="padding:.45rem 0;font-family:\'JetBrains Mono\','
                    'monospace;font-size:.7rem;color:#10B981;">✓ In session</div>',
                    unsafe_allow_html=True)
        if add_btn and selected_tag:
            st.session_state.session_tags.append(selected_tag)
            st.rerun()

        # ── Priority order controls ──────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="sec-hdr">📋 Session Priority List</div>',
                    unsafe_allow_html=True)

        session_tags = st.session_state.session_tags
        if not session_tags:
            st.info("Add equipment tags above to build your session.")
        else:
            from streamlit_sortables import sort_items

            alloc_df = cascade_allocate(session_tags)
            tag_name = eq_master.set_index("Equipment_Tag_No.")["Name"].to_dict()
            tag_loc  = eq_master.set_index("Equipment_Tag_No.")["Location"].to_dict()

            # Static labels — no % to keep component state stable across reruns
            def _item_label(i, t):
                name = tag_name.get(t, t)[:28]
                loc  = tag_loc.get(t, "—")
                return f"#{i+1}  ||  {t}  ||  {name}  ||  {loc}"

            display_items = [_item_label(i, t) for i, t in enumerate(session_tags)]

            # Key tied to exact list — forces full re-init when items added/removed
            sort_key = "sess_sort_" + "_".join(session_tags)
            st.caption("⠿ Drag to reorder — order is applied instantly.")
            sorted_display = sort_items(display_items, direction="vertical", key=sort_key)

            # Parse tag back using explicit || delimiter
            def _parse_tag(label):
                parts = label.split("  ||  ")
                return parts[1].strip() if len(parts) > 1 else label.strip()

            new_order = [_parse_tag(l) for l in sorted_display if _parse_tag(l) in session_tags]
            # Safety fallback
            if len(new_order) != len(session_tags):
                new_order = session_tags[:]

            # Auto-apply if order changed
            if new_order != st.session_state.session_tags:
                st.session_state.session_tags = new_order
                st.rerun()

            # Show fulfillment pills below the sortable list
            alloc_df2 = cascade_allocate(new_order)
            for t in new_order:
                pct  = tag_fulfillment(alloc_df2, t)
                name = tag_name.get(t, t)
                loc  = tag_loc.get(t, "—")
                dot  = status_dot(pct)
                st.markdown(
                    f'<div class="session-equip" style="margin-bottom:.25rem;">'
                    f'<span class="{dot}" style="font-family:\'JetBrains Mono\','
                    f'monospace;font-size:.75rem;font-weight:600;color:var(--t1);">'
                    f'{t}</span>'
                    f'<span style="font-size:.78rem;color:var(--t3);margin-left:.5rem;">'
                    f'{name[:28]}</span>'
                    f'<span style="font-family:\'JetBrains Mono\',monospace;'
                    f'font-size:.65rem;color:var(--t4);margin-left:.5rem;">{loc}</span>'
                    f'<span style="float:right;">{fulfil_pill(pct)}</span>'
                    f'</div>', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑 Clear All", key="clear_all"):
                st.session_state.session_tags = []
                st.rerun()

    # ── RIGHT: equipment info card + system-code material tables ─────────────
    with right:
        if not selected_tag:
            st.markdown("""
            <div style="text-align:center;padding:4rem 1rem;">
              <div style="font-size:2.5rem;opacity:.12;margin-bottom:.8rem;">🔩</div>
              <div style="font-family:'JetBrains Mono',monospace;font-size:.75rem;
                          color:var(--t5);letter-spacing:.1em;">
                SELECT AN EQUIPMENT TAG TO VIEW DETAILS</div></div>""",
                unsafe_allow_html=True)
        else:
            row = eq_master[eq_master["Equipment_Tag_No."]==selected_tag].iloc[0]
            tag_codes = dm[dm["Equipment_Tag_No."]==selected_tag][
                ["Lining_System_Code","Lining_System_Short_Name","Total_SQM"]
            ].drop_duplicates().sort_values("Lining_System_Code", key=lambda x: x.astype(int))

            # Info card
            st.markdown(
                f'<div class="card card-amber">'
                f'<div style="display:flex;justify-content:space-between;'
                f'align-items:flex-start;margin-bottom:.7rem;">'
                f'<div><div style="font-family:\'JetBrains Mono\',monospace;'
                f'font-size:.62rem;color:var(--t4);letter-spacing:.1em;'
                f'text-transform:uppercase;">Equipment Tag</div>'
                f'<div style="font-family:\'JetBrains Mono\',monospace;'
                f'font-size:1.1rem;font-weight:700;color:#F59E0B;">'
                f'{row["Equipment_Tag_No."]}</div></div>'
                f'{loc_badge(str(row["Location"]))}</div>'
                f'<div style="font-size:.95rem;font-weight:600;color:var(--t0);'
                f'margin-bottom:.9rem;">{row["Name"]}</div>'
                f'<div style="display:grid;grid-template-columns:1fr 1fr 1fr;'
                f'gap:.4rem .7rem;font-size:.78rem;">'
                f'<div><span style="color:var(--t4);">Type</span><br>'
                f'<span style="color:var(--t1);">{row["Type"]}</span></div>'
                f'<div><span style="color:var(--t4);">Description</span><br>'
                f'<span style="color:var(--t1);">{row["Description"] or "—"}</span></div>'
                f'<div><span style="color:var(--t4);">Material Spec.</span><br>'
                f'<span style="color:var(--t1);">{row["Material_Spec"] or "—"}</span></div>'
                f'<div style="grid-column:1/-1;">'
                f'<span style="color:var(--t4);">Lining Systems</span><br>'
                f'<span style="color:var(--t2);font-size:.75rem;line-height:1.5;">'
                f'{str(row["Lining_Systems"]).replace(chr(10),"<br>")}'
                f'</span></div></div></div>',
                unsafe_allow_html=True)

            # System code sections
            st.markdown('<div class="sec-hdr" style="margin-top:1rem;">'
                        '⚗️ System Code Material Requirements</div>',
                        unsafe_allow_html=True)

            for _, sc_row in tag_codes.iterrows():
                code  = sc_row["Lining_System_Code"]
                sname = sc_row["Lining_System_Short_Name"]
                sqm   = sc_row["Total_SQM"]

                mat_rows = dm[
                    (dm["Equipment_Tag_No."]==selected_tag) &
                    (dm["Lining_System_Code"]==code)
                ].copy()
                mat_rows = mat_rows.merge(
                    inv[["Material_Code","Available_Qty"]], on="Material_Code", how="left")
                mat_rows["Available_Qty"]   = mat_rows["Available_Qty"].fillna(0)
                mat_rows["Allocated_Qty"]   = mat_rows["Available_Qty"].clip(
                    upper=mat_rows["Demand_Qty"])
                mat_rows["Shortfall_Qty"]   = (
                    mat_rows["Demand_Qty"] - mat_rows["Allocated_Qty"]).clip(lower=0)
                mat_rows["Fulfillment_Pct"] = (
                    mat_rows["Allocated_Qty"] /
                    mat_rows["Demand_Qty"].replace(0,np.nan) * 100
                ).fillna(100).clip(0,100).round(2)

                d_sum = mat_rows["Demand_Qty"].sum()
                a_sum = mat_rows["Allocated_Qty"].sum()
                pct   = min(100, a_sum/d_sum*100) if d_sum > 0 else 100
                short = mat_rows["Shortfall_Qty"].sum()

                with st.expander(
                    f"System Code {code}  ·  {sname}  ·  {sqm:,.2f} SQM  "
                    f"·  Coverage: {pct:.1f}%",
                    expanded=True,
                ):
                    mi1,mi2,mi3,mi4 = st.columns(4)
                    mi1.metric("System Code",   code)
                    mi2.metric("Short Name",    sname)
                    mi3.metric("Surface Area",  f"{sqm:,.2f} SQM")
                    mi4.metric("Coverage",      f"{pct:.1f}%")
                    plotly_mat_table(
                        mat_rows, f"entry_{selected_tag}_{code}",
                        height=65 + len(mat_rows)*30
                    )

            # Grand total for this equipment
            all_mat = dm[dm["Equipment_Tag_No."]==selected_tag].merge(
                inv[["Material_Code","Available_Qty"]], on="Material_Code", how="left")
            all_mat["Available_Qty"] = all_mat["Available_Qty"].fillna(0)
            all_mat["Shortfall"]     = (
                all_mat["Demand_Qty"] -
                all_mat["Available_Qty"].clip(upper=all_mat["Demand_Qty"])
            ).clip(lower=0)
            gt_demand = all_mat["Demand_Qty"].sum()
            gt_alloc  = all_mat["Available_Qty"].clip(upper=all_mat["Demand_Qty"]).sum()
            gt_short  = all_mat["Shortfall"].sum()
            gt_pct    = min(100, gt_alloc/gt_demand*100) if gt_demand > 0 else 100

            st.markdown(f"""
            <div class="grand-box" style="margin-top:.8rem;">
              <div style="font-family:'JetBrains Mono',monospace;font-size:.6rem;
                          letter-spacing:.14em;text-transform:uppercase;
                          color:#F59E0B;margin-bottom:.7rem;">
                Equipment Grand Total — {selected_tag}</div>
              <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:.7rem;">
                <div><div style="font-family:'JetBrains Mono',monospace;
                     font-size:1.3rem;font-weight:700;color:var(--t0);">
                     {len(tag_codes)}</div>
                     <div style="font-size:.62rem;text-transform:uppercase;
                     letter-spacing:.08em;color:var(--t4);">System Codes</div></div>
                <div><div style="font-family:'JetBrains Mono',monospace;
                     font-size:1.3rem;font-weight:700;color:var(--t0);">
                     {gt_demand:,.0f}</div>
                     <div style="font-size:.62rem;text-transform:uppercase;
                     letter-spacing:.08em;color:var(--t4);">Total Demand</div></div>
                <div><div style="font-family:'JetBrains Mono',monospace;
                     font-size:1.3rem;font-weight:700;
                     color:{'#EF4444' if gt_short>0 else '#10B981'};">
                     {gt_short:,.1f}</div>
                     <div style="font-size:.62rem;text-transform:uppercase;
                     letter-spacing:.08em;color:var(--t4);">Total Shortfall</div></div>
                <div><div style="font-family:'JetBrains Mono',monospace;
                     font-size:1.3rem;font-weight:700;
                     color:{'#10B981' if gt_pct>=100 else '#F97316' if gt_pct>=90 else '#EAB308' if gt_pct>=80 else '#EF4444'};">
                     {gt_pct:.1f}%</div>
                     <div style="font-size:.62rem;text-transform:uppercase;
                     letter-spacing:.08em;color:var(--t4);">Coverage</div></div>
              </div>
            </div>""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 · SESSION ORDER REPORT
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    session_tags = st.session_state.session_tags

    if not session_tags:
        st.info("Add equipment tags in the Entry tab to generate a report.")
    else:
        alloc_df = cascade_allocate(session_tags)
        tag_name = eq_master.set_index("Equipment_Tag_No.")["Name"].to_dict()
        tag_loc  = eq_master.set_index("Equipment_Tag_No.")["Location"].to_dict()

        # Session KPIs
        tot_demand = alloc_df["Demand_Qty"].sum()
        tot_alloc  = alloc_df["Allocated_Qty"].sum()
        tot_short  = alloc_df["Shortfall_Qty"].sum()
        ov_pct     = min(100, tot_alloc/tot_demand*100) if tot_demand > 0 else 100
        n_mats     = alloc_df["Material_Code"].nunique()
        n_short_m  = alloc_df[alloc_df["Shortfall_Qty"]>0]["Material_Code"].nunique()

        k1,k2,k3,k4,k5 = st.columns(5)
        k1.metric("Equipment",       len(session_tags))
        k2.metric("Materials",       n_mats)
        k3.metric("Need to Order",   n_short_m)
        k4.metric("Total Shortfall", f"{tot_short:,.1f}")
        k5.metric("Overall Coverage",f"{ov_pct:.1f}%")
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Priority reorder (updates global session_tags) ───────────────────
        st.markdown('<div class="sec-hdr">⠿ Drag to Reorder Priority — changes reflect everywhere</div>',
                    unsafe_allow_html=True)

        from streamlit_sortables import sort_items as _sort2

        def _t2_sl(i, t):
            return f"#{i+1}  ||  {t}  ||  {tag_name.get(t,t)[:28]}  ||  {tag_loc.get(t,'—')}"

        t2_display = [_t2_sl(i, t) for i, t in enumerate(session_tags)]
        t2_key     = "t2_sort_" + "_".join(session_tags)
        t2_sorted  = _sort2(t2_display, direction="vertical", key=t2_key)

        def _t2_parse(label):
            parts = label.split("  ||  ")
            return parts[1].strip() if len(parts) > 1 else label.strip()

        t2_new_order = [_t2_parse(l) for l in t2_sorted if _t2_parse(l) in session_tags]
        
        if len(t2_new_order) == len(session_tags) and t2_new_order != st.session_state.session_tags:
            st.session_state.session_tags = t2_new_order
            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        # ── Per-equipment expanders ──────────────────────────────────────────
        st.markdown('<div class="sec-hdr">Per-Equipment System Code Breakdown</div>',
                    unsafe_allow_html=True)

        for i, tag in enumerate(session_tags):
            tag_alloc = alloc_df[alloc_df["Equipment_Tag_No."]==tag]
            t_pct     = tag_fulfillment(alloc_df, tag)
            dot       = status_dot(t_pct)
            t_short   = tag_alloc["Shortfall_Qty"].sum()
            name      = tag_name.get(tag, tag)
            loc       = tag_loc.get(tag, "—")
            eq_row    = eq_master[eq_master["Equipment_Tag_No."]==tag].iloc[0]

            with st.expander(
                f"#{i+1}  {tag}  ·  {name}  ·  {loc}  "
                f"·  Coverage: {t_pct:.1f}%",
                expanded=False,
            ):
                # Equipment meta strip
                m1,m2,m3,m4 = st.columns(4)
                m1.markdown(f'**Type:** {eq_row["Type"]}')
                m2.markdown(f'**Description:** {eq_row["Description"] or "—"}')
                m3.markdown(f'**Material Spec.:** {eq_row["Material_Spec"] or "—"}')
                m4.markdown(f'**Total SQM:** `{eq_row["Total_SQM"]:,.2f}`')
                st.caption(f'**Lining Systems:** '
                           f'{str(eq_row["Lining_Systems"]).replace(chr(10)," | ")}')
                st.markdown("---")

                # Per system code tables
                for code in sorted(tag_alloc["Lining_System_Code"].unique(),
                                   key=lambda x: int(x)):
                    code_alloc = tag_alloc[tag_alloc["Lining_System_Code"]==code].copy()
                    sname = code_alloc["Lining_System_Short_Name"].iloc[0]
                    sqm   = code_alloc["Total_SQM"].iloc[0]
                    c_pct = syscode_fulfillment(alloc_df, tag, code)
                    c_short = code_alloc["Shortfall_Qty"].sum()
                    c_demand = code_alloc["Demand_Qty"].sum()
                    c_alloc  = code_alloc["Allocated_Qty"].sum()

                    st.markdown(
                        f'<div class="syscode-block">'
                        f'<div class="syscode-hdr">'
                        f'<span class="code-badge">Code {code}</span>'
                        f'<span style="font-size:.8rem;color:var(--t1);font-weight:500;">'
                        f'{sname}</span>'
                        f'<span style="font-family:\'JetBrains Mono\',monospace;'
                        f'font-size:.72rem;color:var(--t3);">{sqm:,.2f} SQM</span>'
                        f'<span style="margin-left:auto;">{fulfil_pill(c_pct)}</span>'
                        f'</div></div>',
                        unsafe_allow_html=True)

                    sc1,sc2,sc3 = st.columns(3)
                    sc1.metric("Demand",     f"{c_demand:,.2f}")
                    sc2.metric("Allocated",  f"{c_alloc:,.2f}")
                    sc3.metric("Shortfall",  f"{c_short:,.2f}")
                    plotly_mat_table(
                        code_alloc,
                        f"rep_{tag}_{code}",
                        height=65 + len(code_alloc)*30
                    )

                # Equipment grand total row
                st.markdown(
                    f'<div style="background:rgba(245,158,11,.05);'
                    f'border:1px solid rgba(245,158,11,.2);border-radius:6px;'
                    f'padding:.7rem 1rem;margin-top:.5rem;'
                    f'font-family:\'JetBrains Mono\',monospace;font-size:.8rem;">'
                    f'<span style="color:#F59E0B;font-weight:700;">GRAND TOTAL — {tag}</span>'
                    f'<span style="color:var(--t3);margin-left:1.5rem;">'
                    f'Demand: <strong style="color:var(--t1);">'
                    f'{tag_alloc["Demand_Qty"].sum():,.2f}</strong></span>'
                    f'<span style="color:var(--t3);margin-left:1rem;">'
                    f'Allocated: <strong style="color:var(--t1);">'
                    f'{tag_alloc["Allocated_Qty"].sum():,.2f}</strong></span>'
                    f'<span style="color:var(--t3);margin-left:1rem;">'
                    f'Shortfall: <strong style="color:#EF4444;">'
                    f'{tag_alloc["Shortfall_Qty"].sum():,.2f}</strong></span>'
                    f'<span style="margin-left:1rem;">{fulfil_pill(t_pct)}</span>'
                    f'</div>',
                    unsafe_allow_html=True)

        # ── Combined procurement list ─────────────────────────────────────────
        st.markdown('<div class="sec-hdr" style="margin-top:1.5rem;">'
                    '🛒 Combined Procurement List</div>',
                    unsafe_allow_html=True)

        combined = alloc_df.groupby(
            ["Material_Code","Material_Name","UOM"], as_index=False
        ).agg(
            Demand_Qty    =("Demand_Qty",    "sum"),
            Allocated_Qty =("Allocated_Qty", "sum"),
            Shortfall_Qty =("Shortfall_Qty", "sum"),
        )
        combined["Fulfillment_Pct"] = (
            combined["Allocated_Qty"] /
            combined["Demand_Qty"].replace(0, np.nan) * 100
        ).fillna(100).clip(0,100).round(2)
        combined = combined.sort_values("Fulfillment_Pct")

        # Stacked bar (shortage only)
        shortage_only = combined[combined["Shortfall_Qty"]>0].copy()
        if not shortage_only.empty:
            shortage_only["Label"] = (
                shortage_only["Material_Code"] + "  " +
                shortage_only["Material_Name"].fillna("").str[:22]
            )
            fig_bar = go.Figure()
            fig_bar.add_trace(go.Bar(
                name="Available", y=shortage_only["Label"],
                x=shortage_only["Allocated_Qty"], orientation="h",
                marker_color="#10B981", marker_opacity=.75,
                text=shortage_only["Allocated_Qty"].apply(lambda v:f"{v:,.0f}"),
                textposition="inside",
                textfont=dict(family="JetBrains Mono",size=9,color="#fff"),
            ))
            fig_bar.add_trace(go.Bar(
                name="To Order", y=shortage_only["Label"],
                x=shortage_only["Shortfall_Qty"], orientation="h",
                marker_color="#EF4444", marker_opacity=.75,
                text=shortage_only["Shortfall_Qty"].apply(lambda v:f"{v:,.1f}"),
                textposition="inside",
                textfont=dict(family="JetBrains Mono",size=9,color="#fff"),
            ))
            fig_bar.update_layout(
                barmode="stack",
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(family="JetBrains Mono", size=10), # Fixed Font
                legend=dict(orientation="h", yanchor="bottom", y=1.02,
                            bgcolor="rgba(0,0,0,0)", font=dict(size=10)),
                margin=dict(l=0, r=60, t=28, b=0),
                height=max(300, len(shortage_only)*42),
                xaxis=dict(gridcolor="rgba(128, 128, 128, 0.2)", zerolinecolor="rgba(128, 128, 128, 0.2)"), # Fixed Grids
                yaxis=dict(gridcolor="rgba(128, 128, 128, 0.2)"), # Fixed Grids
            )
            st.plotly_chart(fig_bar, use_container_width=True, key="session_bar")

        plotly_mat_table(combined, "session_combined",
                         height=90+len(combined)*30)

        # Grand total box
        st.markdown(f"""
        <div class="grand-box" style="margin-top:1rem;">
          <div style="font-family:'JetBrains Mono',monospace;font-size:.6rem;
                      letter-spacing:.14em;text-transform:uppercase;
                      color:#F59E0B;margin-bottom:.7rem;">
            ⭐ Grand Total — {len(session_tags)} Equipment Session</div>
          <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:.7rem;">
            <div><div style="font-family:'JetBrains Mono',monospace;font-size:1.25rem;
                 font-weight:700;color:var(--t0);">{len(session_tags)}</div>
                 <div style="font-size:.6rem;text-transform:uppercase;
                 letter-spacing:.08em;color:var(--t4);">Equipment</div></div>
            <div><div style="font-family:'JetBrains Mono',monospace;font-size:1.25rem;
                 font-weight:700;color:var(--t0);">{n_mats}</div>
                 <div style="font-size:.6rem;text-transform:uppercase;
                 letter-spacing:.08em;color:var(--t4);">Materials</div></div>
            <div><div style="font-family:'JetBrains Mono',monospace;font-size:1.25rem;
                 font-weight:700;color:var(--t0);">{tot_demand:,.0f}</div>
                 <div style="font-size:.6rem;text-transform:uppercase;
                 letter-spacing:.08em;color:var(--t4);">Total Demand</div></div>
            <div><div style="font-family:'JetBrains Mono',monospace;font-size:1.25rem;
                 font-weight:700;color:#EF4444;">{n_short_m}</div>
                 <div style="font-size:.6rem;text-transform:uppercase;
                 letter-spacing:.08em;color:var(--t4);">To Procure</div></div>
            <div><div style="font-family:'JetBrains Mono',monospace;font-size:1.25rem;
                 font-weight:700;color:#EF4444;">{tot_short:,.1f}</div>
                 <div style="font-size:.6rem;text-transform:uppercase;
                 letter-spacing:.08em;color:var(--t4);">Shortfall Units</div></div>
          </div>
        </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        d1, d2 = st.columns(2)
        with d1:
            st.download_button("⬇ Full Session Report",
                               data=excel_bytes(alloc_df),
                               file_name="session_full_report.xlsx",
                               mime="application/vnd.ms-excel",
                               use_container_width=True)
        with d2:
            if not shortage_only.empty:
                st.download_button("⬇ Order List Only",
                                   data=excel_bytes(shortage_only),
                                   file_name="order_list.xlsx",
                                   mime="application/vnd.ms-excel",
                                   use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 · LOCATION REPORT
# ═══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="sec-hdr">📍 All Equipment by Location — Cascading Balance</div>',
                unsafe_allow_html=True)
    st.caption("Drag to reorder equipment within each location. Order is independent from the session list.")

    # ── Per-location order state (independent from global session_tags) ──────
    if "loc_order" not in st.session_state:
        st.session_state.loc_order = {}

    from streamlit_sortables import sort_items as _sort3

    loc_color = {
        "Brown Field": ("loc-bf","#3B82F6"),
        "TRAIN J":     ("loc-tj","#F59E0B"),
        "TRAIN K":     ("loc-tk","#10B981"),
    }

    for loc in LOCATION_ORDER:
        # Initialise from file order if not set
        default_loc_order = eq_master[eq_master["Location"]==loc]["Equipment_Tag_No."].tolist()
        if loc not in st.session_state.loc_order:
            st.session_state.loc_order[loc] = default_loc_order

        loc_tags_all = st.session_state.loc_order[loc]
        if not loc_tags_all:
            continue

        badge_cls, accent = loc_color.get(loc, ("loc-bf","#3B82F6"))

        # Compute quick fulfillment with CURRENT order for labels only
        loc_alloc_preview = cascade_allocate(loc_tags_all)

        # Static labels — no % to keep component state stable across reruns
        def _l3_label(i, t):
            name = eq_master.set_index("Equipment_Tag_No.")["Name"].get(t, t)[:26]
            sqm  = eq_master.set_index("Equipment_Tag_No.")["Total_SQM"].get(t, 0)
            return f"#{i+1}  ||  {t}  ||  {name}  ||  {sqm:,.1f} SQM"

        loc_display = [_l3_label(i, t) for i, t in enumerate(loc_tags_all)]
        loc_key     = f"loc_sort_{loc}_" + "_".join(loc_tags_all)
        loc_sorted  = _sort3(loc_display, direction="vertical", key=loc_key)

        def _l3_parse(label):
            parts = label.split("  ||  ")
            return parts[1].strip() if len(parts) > 1 else label.strip()

        new_loc_order = [_l3_parse(l) for l in loc_sorted if _l3_parse(l) in loc_tags_all]
        
        if len(new_loc_order) == len(loc_tags_all) and new_loc_order != st.session_state.loc_order[loc]:
            st.session_state.loc_order[loc] = new_loc_order
            st.rerun()

        # Re-cascade with (possibly new) order
        loc_alloc = cascade_allocate(loc_tags_all)

        loc_demand = loc_alloc["Demand_Qty"].sum()
        loc_short  = loc_alloc["Shortfall_Qty"].sum()
        loc_pct    = min(100,
            loc_alloc["Allocated_Qty"].sum()/loc_demand*100
        ) if loc_demand > 0 else 100

        # Location header
        st.markdown(
            f'<div style="display:flex;align-items:center;gap:1rem;'
            f'margin:1.2rem 0 .7rem;">'
            f'<span class="loc-badge {badge_cls}" style="font-size:.76rem;'
            f'padding:.28rem .9rem;">{loc}</span>'
            f'<span style="font-family:\'JetBrains Mono\',monospace;font-size:.75rem;'
            f'color:var(--t4);">{len(loc_tags_all)} equipment</span>'
            f'<span style="font-family:\'JetBrains Mono\',monospace;font-size:.75rem;'
            f'color:{"#EF4444" if loc_short>0 else "#10B981"};">'
            f'Shortfall: {loc_short:,.1f}</span>'
            f'<span style="font-family:\'JetBrains Mono\',monospace;font-size:.75rem;'
            f'color:var(--t3);">Coverage: {loc_pct:.1f}%</span>'
            f'</div>',
            unsafe_allow_html=True)

        # Per-equipment expanders
        for tag in loc_tags_all:
            tag_alloc = loc_alloc[loc_alloc["Equipment_Tag_No."]==tag]
            t_pct     = tag_fulfillment(loc_alloc, tag)
            t_short   = tag_alloc["Shortfall_Qty"].sum()
            eq_row    = eq_master[eq_master["Equipment_Tag_No."]==tag].iloc[0]
            dot       = status_dot(t_pct)

            _dot_char = "🟢" if t_pct>=100 else "🟠" if t_pct>=90 else "🟡" if t_pct>=80 else "🔴"
            with st.expander(
                f"{_dot_char}  {tag}  ·  {eq_row['Name']}  ·  "
                f"{eq_row['Total_SQM']:,.2f} SQM  ·  Coverage: {t_pct:.1f}%",
                expanded=False,
            ):
                c1,c2,c3 = st.columns(3)
                c1.markdown(f'**Type:** {eq_row["Type"]}')
                c2.markdown(f'**Description:** {eq_row["Description"] or "—"}')
                c3.markdown(f'**Material Spec.:** {eq_row["Material_Spec"] or "—"}')
                st.caption(
                    f'**Lining:** {str(eq_row["Lining_Systems"]).replace(chr(10)," | ")}')
                st.markdown("---")

                # Per system code
                for code in sorted(tag_alloc["Lining_System_Code"].unique(),
                                   key=lambda x: int(x)):
                    code_alloc = tag_alloc[tag_alloc["Lining_System_Code"]==code].copy()
                    sname = code_alloc["Lining_System_Short_Name"].iloc[0]
                    sqm   = code_alloc["Total_SQM"].iloc[0]
                    c_pct = syscode_fulfillment(loc_alloc, tag, code)

                    st.markdown(
                        f'<div class="syscode-block">'
                        f'<div class="syscode-hdr">'
                        f'<span class="code-badge">Code {code}</span>'
                        f'<span style="font-size:.8rem;color:var(--t1);">{sname}</span>'
                        f'<span style="font-family:\'JetBrains Mono\',monospace;'
                        f'font-size:.72rem;color:var(--t3);">{sqm:,.2f} SQM</span>'
                        f'<span style="margin-left:auto;">{fulfil_pill(c_pct)}</span>'
                        f'</div></div>',
                        unsafe_allow_html=True)
                    plotly_mat_table(
                        code_alloc,
                        f"loc_{loc}_{tag}_{code}",
                        height=65 + len(code_alloc)*30
                    )

                # Equipment grand total
                st.markdown(
                    f'<div style="background:rgba(245,158,11,.05);'
                    f'border:1px solid rgba(245,158,11,.18);border-radius:6px;'
                    f'padding:.65rem .9rem;margin-top:.4rem;'
                    f'font-family:\'JetBrains Mono\',monospace;font-size:.77rem;">'
                    f'<span style="color:#F59E0B;font-weight:700;">TOTAL — {tag}</span>'
                    f'<span style="color:var(--t3);margin-left:1.2rem;">'
                    f'Demand: <b style="color:var(--t1);">'
                    f'{tag_alloc["Demand_Qty"].sum():,.2f}</b></span>'
                    f'<span style="color:var(--t3);margin-left:.8rem;">'
                    f'Shortfall: <b style="color:#EF4444;">'
                    f'{tag_alloc["Shortfall_Qty"].sum():,.2f}</b></span>'
                    f'<span style="margin-left:.8rem;">{fulfil_pill(t_pct)}</span>'
                    f'</div>',
                    unsafe_allow_html=True)

                # Add to session button
                if tag in st.session_state.session_tags:
                    st.markdown(
                        '<span style="font-family:\'JetBrains Mono\',monospace;'
                        'font-size:.7rem;color:#10B981;">✓ In session</span>',
                        unsafe_allow_html=True)
                else:
                    if st.button(f"＋ Add {tag} to Session",
                                 key=f"locadd_{loc}_{tag}"):
                        st.session_state.session_tags.append(tag)
                        st.rerun()

        # ── Bar chart (collapsible) per location ──────────────────────────────
        with st.expander(f"📊 Show Shortfall Chart — {loc}", expanded=False):
            chart_data = []
            for tag in loc_tags_all:
                tag_alloc = loc_alloc[loc_alloc["Equipment_Tag_No."]==tag]
                for code in sorted(tag_alloc["Lining_System_Code"].unique(),
                                   key=lambda x: int(x)):
                    ca = tag_alloc[tag_alloc["Lining_System_Code"]==code]
                    sname = ca["Lining_System_Short_Name"].iloc[0]
                    chart_data.append({
                        "Label": f"{tag}\nCode {code} ({sname})",
                        "Demand":    ca["Demand_Qty"].sum(),
                        "Allocated": ca["Allocated_Qty"].sum(),
                        "Shortfall": ca["Shortfall_Qty"].sum(),
                    })
            cdf = pd.DataFrame(chart_data)
            if not cdf.empty:
                fig_loc = go.Figure()
                fig_loc.add_trace(go.Bar(
                    name="Allocated", y=cdf["Label"], x=cdf["Allocated"],
                    orientation="h", marker_color=accent, marker_opacity=.65,
                    text=cdf["Allocated"].apply(lambda v:f"{v:,.0f}"),
                    textposition="inside",
                    textfont=dict(family="JetBrains Mono",size=8,color="#fff"),
                ))
                fig_loc.add_trace(go.Bar(
                    name="Shortfall", y=cdf["Label"], x=cdf["Shortfall"],
                    orientation="h", marker_color="#EF4444", marker_opacity=.75,
                    text=cdf["Shortfall"].apply(
                        lambda v:f"{v:,.0f}" if v > 0 else ""),
                    textposition="inside",
                    textfont=dict(family="JetBrains Mono",size=8,color="#fff"),
                ))
                fig_loc.update_layout(
                    barmode="stack",
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(family="JetBrains Mono",size=9,color="var(--t3)"),
                    legend=dict(orientation="h",yanchor="bottom",y=1.01,
                                bgcolor="rgba(0,0,0,0)"),
                    margin=dict(l=0,r=60,t=28,b=0),
                    height=max(350, len(cdf)*36),
                    xaxis=dict(gridcolor="#1E2E46",zerolinecolor="#1E2E46"),
                    yaxis=dict(gridcolor="#1E2E46"),
                    title=dict(text=f"{loc} — Demand vs Shortfall by System Code",
                               font=dict(family="JetBrains Mono",size=11,color="var(--t2)"))
                )
                st.plotly_chart(fig_loc, use_container_width=True,
                                key=f"loc_chart_{loc}")

        st.markdown(
            f'<div style="border-bottom:1px solid #1E2E46;margin:1rem 0;"></div>',
            unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 4 · EXECUTION PLAN
# ═══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="sec-hdr">⚙️ Execution Plan — Critical System Code Analysis</div>',
                unsafe_allow_html=True)

    session_tags = st.session_state.session_tags

    if not session_tags:
        st.info("Add equipment tags in the Entry tab first.")
    else:
        alloc_df   = cascade_allocate(session_tags)
        tag_name   = eq_master.set_index("Equipment_Tag_No.")["Name"].to_dict()

        sel_tag = st.selectbox(
            "Select Equipment",
            options=session_tags,
            format_func=lambda t: f"{t}  —  {tag_name.get(t,t)}",
            key="exec_tag",
        )

        tag_alloc = alloc_df[alloc_df["Equipment_Tag_No."]==sel_tag]
        avail_codes = sorted(tag_alloc["Lining_System_Code"].unique(),
                             key=lambda x: int(x))

        if not avail_codes:
            st.warning("No system code data for this equipment.")
        else:
            sel_code = st.selectbox(
                "Select Critical System Code",
                options=avail_codes,
                format_func=lambda c: (
                    f"Code {c}  —  "
                    f"{tag_alloc[tag_alloc['Lining_System_Code']==c]['Lining_System_Short_Name'].iloc[0]}"
                ),
                key="exec_code",
            )

            st.markdown("<br>", unsafe_allow_html=True)

            # Critical system code data
            crit = tag_alloc[tag_alloc["Lining_System_Code"]==sel_code].copy()
            crit_sname  = crit["Lining_System_Short_Name"].iloc[0]
            crit_sqm    = crit["Total_SQM"].iloc[0]
            crit_demand = crit["Demand_Qty"].sum()
            crit_alloc  = crit["Allocated_Qty"].sum()
            crit_short  = crit["Shortfall_Qty"].sum()
            crit_pct    = min(100, crit_alloc/crit_demand*100) if crit_demand > 0 else 100

            # Other system codes
            other_codes = [c for c in avail_codes if c != sel_code]
            other_alloc = tag_alloc[tag_alloc["Lining_System_Code"].isin(other_codes)]
            other_short = other_alloc["Shortfall_Qty"].sum()

            # ── Critical system code card ─────────────────────────────────────
            st.markdown(
                f'<div class="card card-amber">'
                f'<div style="font-family:\'JetBrains Mono\',monospace;'
                f'font-size:.6rem;letter-spacing:.14em;text-transform:uppercase;'
                f'color:#F59E0B;margin-bottom:.6rem;">Critical System Code</div>'
                f'<div style="display:flex;align-items:center;gap:1rem;'
                f'margin-bottom:.8rem;">'
                f'<span class="code-badge" style="font-size:.85rem;'
                f'padding:.3rem .8rem;">Code {sel_code}</span>'
                f'<span style="font-size:.95rem;font-weight:600;color:var(--t0);">'
                f'{crit_sname}</span>'
                f'<span style="font-family:\'JetBrains Mono\',monospace;'
                f'font-size:.8rem;color:var(--t3);">{crit_sqm:,.2f} SQM</span>'
                f'<span style="margin-left:auto;font-family:\'JetBrains Mono\','
                f'monospace;font-size:1.4rem;font-weight:700;'
                f'color:{"#10B981" if crit_pct>=100 else "#F97316" if crit_pct>=90 else "#EAB308" if crit_pct>=80 else "#EF4444"};">'
                f'{crit_pct:.1f}%</span>'
                f'</div>'
                f'<div style="font-size:.82rem;color:var(--t2);line-height:1.7;">'
                f'With current inventory allocation, <strong style="color:#F59E0B;">'
                f'{crit_pct:.1f}%</strong> of System Code {sel_code} ({crit_sname}) '
                f'can be completed for <strong style="color:var(--t0);">{sel_tag}</strong>. '
                + (
                    "✅ All materials for this system code are fully covered."
                    if crit_short == 0 else
                    f'⚠️ <strong style="color:#EF4444;">{crit_short:,.2f} units</strong>'
                    f' short across {(crit["Shortfall_Qty"]>0).sum()} material(s) — order these first to proceed.'
                ) +
                f'</div></div>',
                unsafe_allow_html=True)

            # Full materials table — all system codes
            st.markdown('<div class="sec-hdr" style="margin-top:1rem;">'
                        'All Materials — Status Overview</div>',
                        unsafe_allow_html=True)
            st.caption(
                f"Showing all {len(tag_alloc)} material rows across "
                f"{len(avail_codes)} system code(s) for {sel_tag}. "
                f"Cascade balance applied (priority position #{session_tags.index(sel_tag)+1})."
            )
            plotly_mat_table(
                tag_alloc.copy(),
                f"exec_all_{sel_tag}",
                height=80 + len(tag_alloc)*30
            )

            # ── ORDER PRIORITY SECTION ─────────────────────────────────────────
            st.markdown('<div class="sec-hdr" style="margin-top:1.2rem;">'
                        '📋 Procurement Order Priority</div>',
                        unsafe_allow_html=True)

            # 1️⃣ Critical code shortages
            crit_short_df = crit[crit["Shortfall_Qty"]>0][
                ["Material_Code","Material_Name","UOM","Demand_Qty",
                 "Allocated_Qty","Shortfall_Qty","Fulfillment_Pct"]
            ].copy()

            st.markdown(
                f'<div style="background:var(--red-bg);border:1px solid '
                f'rgba(239,68,68,.25);border-left:4px solid #EF4444;'
                f'border-radius:6px;padding:.8rem 1rem;margin-bottom:.6rem;">'
                f'<div style="font-family:\'JetBrains Mono\',monospace;font-size:.65rem;'
                f'letter-spacing:.12em;text-transform:uppercase;color:#EF4444;'
                f'margin-bottom:.4rem;">1️⃣ Order First — System Code {sel_code} '
                f'({crit_sname}) · Critical Path</div>'
                f'<div style="font-size:.8rem;color:var(--t2);">'
                f'{"No shortages on critical system code — fully covered ✅" if crit_short_df.empty else f"{len(crit_short_df)} material(s) need to be procured for this system code before work can begin."}'
                f'</div></div>',
                unsafe_allow_html=True)

            if not crit_short_df.empty:
                plotly_mat_table(crit_short_df,
                                 f"exec_crit_{sel_tag}_{sel_code}",
                                 height=65+len(crit_short_df)*30)

            # 2️⃣ Other system codes shortages
            for code in other_codes:
                code_alloc_df = tag_alloc[tag_alloc["Lining_System_Code"]==code]
                code_short    = code_alloc_df[code_alloc_df["Shortfall_Qty"]>0]
                sname_o       = code_alloc_df["Lining_System_Short_Name"].iloc[0]
                code_pct      = syscode_fulfillment(alloc_df, sel_tag, code)

                order_num = other_codes.index(code) + 2
                st.markdown(
                    f'<div style="background:var(--amber-bg);border:1px solid '
                    f'rgba(245,158,11,.2);border-left:4px solid #F59E0B;'
                    f'border-radius:6px;padding:.8rem 1rem;margin-bottom:.5rem;">'
                    f'<div style="font-family:\'JetBrains Mono\',monospace;'
                    f'font-size:.65rem;letter-spacing:.12em;text-transform:uppercase;'
                    f'color:#F59E0B;margin-bottom:.4rem;">'
                    f'{order_num}️⃣ Order Next — System Code {code} ({sname_o}) '
                    f'· Coverage: {code_pct:.1f}%</div>'
                    f'<div style="font-size:.8rem;color:var(--t2);">'
                    f'{"All materials covered ✅" if code_short.empty else f"{len(code_short)} material(s) short. Order after critical system code is secured."}'
                    f'</div></div>',
                    unsafe_allow_html=True)

                if not code_short.empty:
                    code_short_display = code_short[
                        ["Material_Code","Material_Name","UOM","Demand_Qty",
                         "Allocated_Qty","Shortfall_Qty","Fulfillment_Pct"]
                    ].copy()
                    plotly_mat_table(code_short_display,
                                     f"exec_other_{sel_tag}_{code}",
                                     height=65+len(code_short_display)*30)

            # Summary box
            all_short_df = tag_alloc[tag_alloc["Shortfall_Qty"]>0]
            total_to_order = all_short_df["Shortfall_Qty"].sum()

            st.markdown(f"""
            <div class="grand-box" style="margin-top:1rem;">
              <div style="font-family:'JetBrains Mono',monospace;font-size:.6rem;
                          letter-spacing:.14em;text-transform:uppercase;
                          color:#F59E0B;margin-bottom:.6rem;">
                Execution Summary — {sel_tag}</div>
              <div style="font-size:.82rem;color:var(--t2);line-height:1.8;">
                Critical system code <strong style="color:#F59E0B;">
                Code {sel_code} ({crit_sname})</strong> is at
                <strong style="color:{'#10B981' if crit_pct>=100 else '#EF4444'};">
                {crit_pct:.1f}%</strong> coverage.
                {"All critical materials are secured — proceed to other system codes." if crit_pct>=100
                  else f"Order {len(crit_short_df)} critical material(s) totalling "
                       f"<strong style='color:#EF4444;'>{crit_short:,.2f} units</strong> first."}
                {"" if other_short==0 else
                  f" Additionally, other system codes require "
                  f"<strong style='color:#F59E0B;'>{other_short:,.2f} units</strong>"
                  f" across {len(all_short_df[~all_short_df['Lining_System_Code'].isin([sel_code])])} material(s)."}
                <br>
                <strong style="color:var(--t0);">
                Total to order for full completion: {total_to_order:,.2f} units
                across {len(all_short_df)} material(s).</strong>
              </div>
            </div>""", unsafe_allow_html=True)

            if not all_short_df.empty:
                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button(
                    f"⬇ Download Execution Order List — {sel_tag}",
                    data=excel_bytes(all_short_df[
                        ["Lining_System_Code","Lining_System_Short_Name",
                         "Material_Code","Material_Name","UOM",
                         "Demand_Qty","Allocated_Qty","Shortfall_Qty","Fulfillment_Pct"]
                    ].sort_values(["Lining_System_Code","Shortfall_Qty"],
                                  ascending=[True,False])),
                    file_name=f"execution_plan_{sel_tag.replace('/','-')}.xlsx",
                    mime="application/vnd.ms-excel",
                    use_container_width=True,
                )