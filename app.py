"""
Smart Material Estimator  ·  app.py  (v2)
==========================================
Loads data directly from project files — no upload needed.
Run:  streamlit run app.py
Deps: streamlit, plotly, pandas, openpyxl
"""

import io, os, sys
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

# ── Sibling modules ───────────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from validate_data    import clean_inventory, clean_recipe, clean_equipment
from allocation_engine import build_demand_matrix

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Smart Material Estimator",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Hard-coded data paths (no upload needed) ──────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PATH_A   = os.path.join(BASE_DIR, "Materials_DetailsAvailable_Qty.xlsx")
PATH_B   = os.path.join(BASE_DIR, "For_1_SQM.xlsx")
PATH_C   = os.path.join(BASE_DIR, "Equipment.xlsx")

SHEET_A  = "Materials"
SHEET_B  = "LINING SYSTEM MATERIAL CONSM"
SHEET_C  = "Data Input"

LOCATION_ORDER = ["Brown Field", "TRAIN J", "TRAIN K"]

# ─────────────────────────────────────────────────────────────────────────────
# STYLES
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;600;700&display=swap');
:root{
  --bg0:#070C13;--bg1:#0D1520;--bg2:#131E2E;--bg3:#192638;
  --border:#1E2E46;--border2:#253550;
  --amber:#F59E0B;--amber2:#FCD34D;--amber-bg:rgba(245,158,11,.08);
  --green:#10B981;--green-bg:rgba(16,185,129,.08);
  --red:#EF4444;--red-bg:rgba(239,68,68,.08);
  --blue:#3B82F6;
  --t0:#F1F5F9;--t1:#CBD5E1;--t2:#94A3B8;--t3:#475569;--t4:#2E3D58;
}
html,body,[class*="css"]{font-family:'Sora',sans-serif!important;background:var(--bg0)!important;color:var(--t1);}
.main .block-container{padding-top:1.6rem;padding-bottom:3rem;max-width:1440px;}
/* sidebar */
[data-testid="stSidebar"]{background:var(--bg1)!important;border-right:1px solid var(--border)!important;}
[data-testid="stSidebar"] *{font-family:'Sora',sans-serif!important;}
/* tabs */
[data-testid="stTabs"] [data-baseweb="tab-list"]{background:var(--bg2);border-bottom:2px solid var(--border);gap:0;padding:0 .5rem;}
[data-testid="stTabs"] [data-baseweb="tab"]{font-family:'JetBrains Mono',monospace!important;font-size:.72rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--t3)!important;padding:.7rem 1.4rem;border-bottom:3px solid transparent;transition:all .2s;}
[data-testid="stTabs"] [aria-selected="true"]{color:var(--amber)!important;border-bottom:3px solid var(--amber)!important;background:transparent!important;}
/* buttons */
.stButton>button{font-family:'JetBrains Mono',monospace!important;font-size:.7rem;font-weight:700;letter-spacing:.1em;text-transform:uppercase;background:var(--amber)!important;color:#000!important;border:none!important;border-radius:4px!important;padding:.52rem 1.3rem!important;transition:all .15s!important;}
.stButton>button:hover{background:#FBBF24!important;transform:translateY(-1px);box-shadow:0 4px 16px rgba(245,158,11,.3)!important;}
.stButton>button[kind="secondary"]{background:var(--bg3)!important;color:var(--t1)!important;border:1px solid var(--border2)!important;}
.stButton>button[kind="secondary"]:hover{background:var(--bg2)!important;border-color:var(--amber)!important;}
/* metrics */
[data-testid="stMetric"]{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:.9rem 1.1rem!important;}
[data-testid="stMetricLabel"]{font-family:'JetBrains Mono',monospace!important;font-size:.62rem!important;letter-spacing:.12em;text-transform:uppercase;color:var(--t3)!important;}
[data-testid="stMetricValue"]{font-family:'JetBrains Mono',monospace!important;font-size:1.9rem!important;color:var(--t0)!important;}
/* info boxes */
[data-testid="stInfo"]   {background:rgba(59,130,246,.07)!important;border-left:3px solid var(--blue)!important;}
[data-testid="stSuccess"]{background:var(--green-bg)!important;border-left:3px solid var(--green)!important;}
[data-testid="stWarning"]{background:var(--amber-bg)!important;border-left:3px solid var(--amber)!important;}
[data-testid="stError"]  {background:var(--red-bg)!important;  border-left:3px solid var(--red)!important;}
/* select / text input */
[data-baseweb="select"]>div,[data-baseweb="input"]>div{background:var(--bg2)!important;border-color:var(--border2)!important;color:var(--t0)!important;}
/* dataframe */
[data-testid="stDataFrame"]{border:1px solid var(--border)!important;border-radius:6px;overflow:hidden;}
/* hr */
hr{border-color:var(--border)!important;margin:1rem 0!important;}
/* custom */
.sec-hdr{font-family:'JetBrains Mono',monospace;font-size:.62rem;font-weight:700;letter-spacing:.18em;text-transform:uppercase;color:var(--t4);border-bottom:1px solid var(--border);padding-bottom:.35rem;margin-bottom:.9rem;}
.card{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:1.1rem 1.3rem;margin-bottom:.7rem;}
.card-amber{border-left:4px solid var(--amber);}
.card-green{border-left:4px solid var(--green);}
.card-red  {border-left:4px solid var(--red);}
.tag-chip{display:inline-block;font-family:'JetBrains Mono',monospace;font-size:.72rem;background:var(--bg3);color:var(--amber);border:1px solid var(--border2);border-radius:4px;padding:.18rem .55rem;margin:.15rem;}
.loc-badge{display:inline-block;font-family:'JetBrains Mono',monospace;font-size:.65rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;padding:.2rem .65rem;border-radius:3px;}
.loc-bf{background:rgba(59,130,246,.12);color:#93C5FD;}
.loc-tj{background:rgba(245,158,11,.12);color:#FCD34D;}
.loc-tk{background:rgba(16,185,129,.12);color:#6EE7B7;}
.pill{display:inline-block;font-family:'JetBrains Mono',monospace;font-size:.7rem;font-weight:600;padding:.18rem .55rem;border-radius:20px;}
.pill-g{background:var(--green-bg);color:var(--green);}
.pill-y{background:var(--amber-bg);color:var(--amber);}
.pill-r{background:var(--red-bg);color:var(--red);}
.session-item{display:flex;align-items:center;gap:.7rem;background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:.55rem .9rem;margin-bottom:.3rem;}
.si-rank{font-family:'JetBrains Mono',monospace;font-size:.65rem;font-weight:700;color:var(--t3);min-width:1.5rem;}
.si-tag{font-family:'JetBrains Mono',monospace;font-size:.75rem;color:var(--amber);min-width:8rem;}
.si-name{font-size:.82rem;color:var(--t0);flex:1;}
.si-loc{font-size:.72rem;color:var(--t2);}
.grand-box{background:linear-gradient(135deg,rgba(245,158,11,.06) 0%,var(--bg2) 70%);border:1px solid rgba(245,158,11,.25);border-left:4px solid var(--amber);border-radius:8px;padding:1.2rem 1.5rem;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADER  (cached — loads once per session)
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Loading project data…")
def load_all():
    df_a_raw = pd.read_excel(PATH_A, sheet_name=SHEET_A)
    df_b_raw = pd.read_excel(PATH_B, sheet_name=SHEET_B)
    df_c_raw = pd.read_excel(PATH_C, sheet_name=SHEET_C)

    inv    = clean_inventory(df_a_raw)
    recipe = clean_recipe(df_b_raw)
    equip  = clean_equipment(df_c_raw)

    demand_df, inv_clean = build_demand_matrix(equip, recipe, inv)

    # Equipment master (one row per tag — full details for the info card)
    raw_for_master = df_c_raw.copy()
    raw_for_master.columns = raw_for_master.columns.str.strip()
    raw_for_master = raw_for_master.dropna(subset=["Equipment_Tag_No.", "Lining_System_Code"])
    raw_for_master["Equipment_Tag_No."] = raw_for_master["Equipment_Tag_No."].astype(str).str.strip()
    raw_for_master["Location"] = raw_for_master["Location"].astype(str).str.strip()
    raw_for_master["Type"]     = raw_for_master["Type"].astype(str).str.strip()
    raw_for_master["Surface_Area_SQM"] = pd.to_numeric(raw_for_master["Surface_Area_SQM"], errors="coerce")

    eq_master = raw_for_master.groupby("Equipment_Tag_No.", as_index=False).agg(
        Name            =("Name",            "first"),
        Description     =("Description",     "first"),
        Location        =("Location",        "first"),
        Type            =("Type",            "first"),
        Lining_Systems  =("Lining_System+",  "first"),
        Material_Spec   =("Material Spec.",  "first"),
        Design          =("Design",          "first"),
        Total_SQM       =("Surface_Area_SQM","sum"),
    )
    eq_master["Location"] = eq_master["Location"].replace({
        "Brown Field ": "Brown Field", "TRAIN J ": "TRAIN J"
    })

    return inv, recipe, equip, demand_df, inv_clean, eq_master


inv, recipe, equip, demand_df, inv_clean, eq_master = load_all()

ALL_TAGS = sorted(eq_master["Equipment_Tag_No."].tolist())

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────
if "session_tags" not in st.session_state:
    st.session_state.session_tags = []   # ordered list of tag strings

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def loc_badge(loc: str) -> str:
    cls = {"Brown Field": "loc-bf", "TRAIN J": "loc-tj", "TRAIN K": "loc-tk"}.get(loc, "loc-bf")
    return f'<span class="loc-badge {cls}">{loc}</span>'

def fulfil_pill(pct: float) -> str:
    cls = "pill-g" if pct >= 100 else "pill-y" if pct >= 50 else "pill-r"
    return f'<span class="pill {cls}">{pct:.1f}%</span>'

def get_tag_demand(tag: str) -> pd.DataFrame:
    """Per-material demand, available, shortfall for a single tag."""
    df = demand_df[demand_df["Equipment_Tag_No."] == tag].copy()
    df = df.merge(inv_clean[["Material_Code", "Available_Qty"]], on="Material_Code", how="left")
    df["Available_Qty"] = df["Available_Qty"].fillna(0)
    df["Shortfall"]     = (df["Demand_Qty"] - df["Available_Qty"]).clip(lower=0).round(3)
    df["Net_Balance"]   = (df["Available_Qty"] - df["Demand_Qty"]).round(3)
    df["Fulfillment_%"] = (
        df["Available_Qty"].clip(upper=df["Demand_Qty"]) /
        df["Demand_Qty"].replace(0, float("nan")) * 100
    ).fillna(100).clip(0, 100).round(1)
    return df[["Material_Code", "Material_Name", "UOM",
               "Demand_Qty", "Available_Qty", "Shortfall",
               "Net_Balance", "Fulfillment_%"]].sort_values("Fulfillment_%")

def get_session_combined(tags: list[str]) -> pd.DataFrame:
    """Aggregated demand across all session tags vs available stock."""
    if not tags:
        return pd.DataFrame()
    df = demand_df[demand_df["Equipment_Tag_No."].isin(tags)].copy()
    agg = df.groupby(["Material_Code", "Material_Name", "UOM"], as_index=False)["Demand_Qty"].sum()
    agg = agg.merge(inv_clean[["Material_Code", "Available_Qty"]], on="Material_Code", how="left")
    agg["Available_Qty"] = agg["Available_Qty"].fillna(0)
    agg["Shortfall"]     = (agg["Demand_Qty"] - agg["Available_Qty"]).clip(lower=0).round(3)
    agg["Net_Balance"]   = (agg["Available_Qty"] - agg["Demand_Qty"]).round(3)
    agg["Fulfillment_%"] = (
        agg["Available_Qty"].clip(upper=agg["Demand_Qty"]) /
        agg["Demand_Qty"].replace(0, float("nan")) * 100
    ).fillna(100).clip(0, 100).round(1)
    return agg.sort_values("Fulfillment_%")

def excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

def plotly_material_table(df: pd.DataFrame, height: int = 400) -> go.Figure:
    """Coloured Plotly table for material breakdown."""
    row_colors, text_colors_pct = [], []
    for pct in df.get("Fulfillment_%", []):
        if pct >= 100:
            row_colors.append("rgba(16,185,129,0.07)")
            text_colors_pct.append("#10B981")
        elif pct >= 50:
            row_colors.append("rgba(245,158,11,0.07)")
            text_colors_pct.append("#F59E0B")
        else:
            row_colors.append("rgba(239,68,68,0.07)")
            text_colors_pct.append("#EF4444")

    show_cols = ["Material_Code", "Material_Name", "UOM",
                 "Demand_Qty", "Available_Qty", "Shortfall",
                 "Net_Balance", "Fulfillment_%"]
    labels = ["Code", "Material Name", "UOM",
              "Demand", "Available", "Shortfall", "Net Balance", "Fulfil %"]

    cell_vals = []
    for c in show_cols:
        if c in ["Demand_Qty", "Available_Qty", "Shortfall"]:
            cell_vals.append([f"{v:,.3f}" for v in df[c]])
        elif c == "Net_Balance":
            cell_vals.append([f"{v:+,.3f}" for v in df[c]])
        elif c == "Fulfillment_%":
            cell_vals.append([f"{v:.1f}%" for v in df[c]])
        else:
            cell_vals.append(df[c].fillna("—").tolist())

    # Last column text colour = fulfillment colour
    n_cols = len(show_cols)
    cell_colors = [row_colors] * (n_cols - 1) + [[row_colors]]
    cell_font_colors = [["#CBD5E1"] * len(df)] * (n_cols - 1) + [text_colors_pct]

    fig = go.Figure(go.Table(
        columnwidth=[80, 200, 50, 80, 80, 80, 90, 70],
        header=dict(
            values=[f"<b>{l}</b>" for l in labels],
            fill_color="#172032", line_color="#253047",
            font=dict(family="JetBrains Mono", size=10, color="#94A3B8"),
            align="left", height=34,
        ),
        cells=dict(
            values=cell_vals,
            fill_color=cell_colors,
            line_color="#1E2E46",
            font=dict(family="JetBrains Mono", size=10,
                      color=cell_font_colors),
            align="left", height=30,
        ),
    ))
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=0, r=0, t=0, b=0), height=height,
    )
    return fig

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding:.5rem 0 1.4rem">
      <div style="font-family:'JetBrains Mono',monospace;font-size:1.05rem;
                  font-weight:700;color:#F59E0B;letter-spacing:.04em;">🏗 SME</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:.58rem;
                  letter-spacing:.18em;text-transform:uppercase;color:#2E3D58;margin-top:2px;">
        Smart Material Estimator</div>
    </div>""", unsafe_allow_html=True)

    st.markdown('<div class="sec-hdr">📊 Project Overview</div>', unsafe_allow_html=True)

    # Location breakdown
    loc_counts = eq_master.groupby("Location")["Equipment_Tag_No."].count()
    for loc in LOCATION_ORDER:
        cnt = loc_counts.get(loc, 0)
        badge = loc_badge(loc)
        st.markdown(
            f'<div style="display:flex;justify-content:space-between;'
            f'align-items:center;margin-bottom:.4rem;">'
            f'{badge}'
            f'<span style="font-family:\'JetBrains Mono\',monospace;font-size:.78rem;'
            f'color:#94A3B8;">{cnt} equip.</span></div>',
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="sec-hdr">🗂 Inventory Snapshot</div>', unsafe_allow_html=True)
    st.caption(f"📦  {len(inv_clean)} materials tracked")
    zero_stock = (inv_clean["Available_Qty"] == 0).sum()
    st.caption(f"⚠️  {zero_stock} material(s) at zero stock")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="sec-hdr">📋 Session</div>', unsafe_allow_html=True)
    n_sess = len(st.session_state.session_tags)
    if n_sess:
        st.markdown(
            f'<div style="font-family:\'JetBrains Mono\',monospace;font-size:.75rem;'
            f'color:#F59E0B;">{n_sess} equipment in session</div>',
            unsafe_allow_html=True,
        )
        for t in st.session_state.session_tags:
            st.caption(f"  · {t}")
        if st.button("🗑 Clear Session", key="clear_all_sidebar"):
            st.session_state.session_tags = []
            st.rerun()
    else:
        st.caption("No equipment added yet.")

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="display:flex;align-items:baseline;gap:1rem;margin-bottom:.3rem;">
  <span style="font-family:'JetBrains Mono',monospace;font-size:1.5rem;
               font-weight:700;color:#F1F5F9;letter-spacing:-.01em;">
    Smart Material Estimator</span>
  <span style="font-size:.8rem;color:#2E3D58;letter-spacing:.03em;">
    Equipment-level material planning · Session-based order reports</span>
</div><hr>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tab_entry, tab_report, tab_location = st.tabs([
    "🔍  Equipment Entry",
    "📦  Session Order Report",
    "📍  Location Report",
])


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1  ·  EQUIPMENT ENTRY PAGE
# ═══════════════════════════════════════════════════════════════════════════════
with tab_entry:

    left_col, right_col = st.columns([1, 1.6], gap="large")

    # ── LEFT: Search & Session List ────────────────────────────────────────
    with left_col:
        st.markdown('<div class="sec-hdr">🔍 Find Equipment</div>',
                    unsafe_allow_html=True)

        selected_tag = st.selectbox(
            "Enter or select Equipment Tag No.",
            options=[""] + ALL_TAGS,
            format_func=lambda t: (
                t if t == "" else
                f"{t}  —  {eq_master.set_index('Equipment_Tag_No.')['Name'].get(t, '')}"
            ),
            key="tag_select",
            help="Type to search by tag number or equipment name",
            label_visibility="collapsed",
        )

        # Add to session button
        c1, c2 = st.columns([2, 1])
        with c1:
            add_btn = st.button(
                "＋ Add to Session", key="add_btn",
                disabled=(selected_tag == "" or
                          selected_tag in st.session_state.session_tags),
            )
        with c2:
            if selected_tag and selected_tag in st.session_state.session_tags:
                st.markdown(
                    '<div style="padding:.45rem 0;font-family:\'JetBrains Mono\','
                    'monospace;font-size:.7rem;color:#10B981;">✓ In session</div>',
                    unsafe_allow_html=True,
                )

        if add_btn and selected_tag:
            st.session_state.session_tags.append(selected_tag)
            st.rerun()

        # ── Session list ──────────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="sec-hdr">📋 Session List</div>',
                    unsafe_allow_html=True)

        if not st.session_state.session_tags:
            st.info("Add equipment tags above to build your order session.")
        else:
            tag_name_map = eq_master.set_index("Equipment_Tag_No.")["Name"].to_dict()
            tag_loc_map  = eq_master.set_index("Equipment_Tag_No.")["Location"].to_dict()

            for i, t in enumerate(st.session_state.session_tags):
                name = tag_name_map.get(t, "—")
                loc  = tag_loc_map.get(t, "—")
                badge_html = loc_badge(loc)
                c_item, c_remove = st.columns([5, 1])
                with c_item:
                    st.markdown(
                        f'<div class="session-item">'
                        f'<span class="si-rank">#{i+1}</span>'
                        f'<span class="si-tag">{t}</span>'
                        f'<span class="si-name">{name}</span>'
                        f'{badge_html}'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                with c_remove:
                    if st.button("✕", key=f"rm_{t}_{i}",
                                 help=f"Remove {t}"):
                        st.session_state.session_tags.remove(t)
                        st.rerun()

            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑 Clear All", key="clear_session"):
                st.session_state.session_tags = []
                st.rerun()

    # ── RIGHT: Equipment Info Card ──────────────────────────────────────────
    with right_col:
        if not selected_tag:
            st.markdown("""
            <div style="text-align:center;padding:4rem 2rem;">
              <div style="font-family:'JetBrains Mono',monospace;font-size:2.5rem;
                          opacity:.15;margin-bottom:.8rem;">🔩</div>
              <div style="font-family:'JetBrains Mono',monospace;font-size:.78rem;
                          color:#2E3D58;letter-spacing:.1em;">
                SELECT AN EQUIPMENT TAG TO VIEW DETAILS</div>
            </div>""", unsafe_allow_html=True)
        else:
            row = eq_master[eq_master["Equipment_Tag_No."] == selected_tag].iloc[0]

            # ── Info card ─────────────────────────────────────────────────
            st.markdown(
                f'<div class="card card-amber">'
                f'<div style="display:flex;justify-content:space-between;'
                f'align-items:flex-start;margin-bottom:.8rem;">'
                f'<div>'
                f'<div style="font-family:\'JetBrains Mono\',monospace;'
                f'font-size:.65rem;color:#475569;letter-spacing:.12em;'
                f'text-transform:uppercase;margin-bottom:.25rem;">Equipment Tag</div>'
                f'<div style="font-family:\'JetBrains Mono\',monospace;'
                f'font-size:1.15rem;font-weight:700;color:#F59E0B;">'
                f'{row["Equipment_Tag_No."]}</div>'
                f'</div>'
                f'{loc_badge(str(row["Location"]))}'
                f'</div>'
                f'<div style="font-size:1rem;font-weight:600;color:#F1F5F9;'
                f'margin-bottom:1rem;">{row["Name"]}</div>'
                f'<div style="display:grid;grid-template-columns:1fr 1fr;'
                f'gap:.5rem .8rem;font-size:.8rem;">'
                f'<div><span style="color:#475569;">Description</span><br>'
                f'<span style="color:#CBD5E1;">{row["Description"] or "—"}</span></div>'
                f'<div><span style="color:#475569;">Type</span><br>'
                f'<span style="color:#CBD5E1;">{row["Type"]}</span></div>'
                f'<div><span style="color:#475569;">Total Lining Area</span><br>'
                f'<span style="font-family:\'JetBrains Mono\',monospace;'
                f'color:#F59E0B;font-weight:600;">{row["Total_SQM"]:,.3f} SQM</span></div>'
                f'<div><span style="color:#475569;">Material Spec.</span><br>'
                f'<span style="color:#CBD5E1;">{row["Material_Spec"] or "—"}</span></div>'
                f'<div style="grid-column:1/-1;">'
                f'<span style="color:#475569;">Lining System(s)</span><br>'
                f'<span style="color:#CBD5E1;font-size:.78rem;line-height:1.5;">'
                f'{str(row["Lining_Systems"]).replace(chr(10),"<br>") or "—"}</span></div>'
                f'</div></div>',
                unsafe_allow_html=True,
            )

            # ── Material requirement table ─────────────────────────────────
            st.markdown('<div class="sec-hdr" style="margin-top:1.2rem;">'
                        '⚗️ Material Requirements</div>', unsafe_allow_html=True)

            mat_df = get_tag_demand(selected_tag)

            if mat_df.empty:
                st.warning("No material recipe found for this equipment tag.")
            else:
                # Summary strip
                n_ok      = (mat_df["Fulfillment_%"] >= 100).sum()
                n_short   = (mat_df["Fulfillment_%"] < 100).sum()
                tot_demand = mat_df["Demand_Qty"].sum()
                tot_short  = mat_df["Shortfall"].sum()

                s1, s2, s3, s4 = st.columns(4)
                s1.metric("Total Materials",     len(mat_df))
                s2.metric("✅ Fully Stocked",    n_ok)
                s3.metric("⚠️ Need to Order",    n_short)
                s4.metric("Total Shortfall",
                          f"{tot_short:,.1f}" if tot_short > 0 else "0")

                # Table
                fig = plotly_material_table(mat_df,
                                            height=80 + len(mat_df) * 30)
                st.plotly_chart(fig, use_container_width=True)

                # Shortage-only download
                shortage_df = mat_df[mat_df["Shortfall"] > 0][
                    ["Material_Code", "Material_Name", "UOM",
                     "Demand_Qty", "Available_Qty", "Shortfall"]
                ].copy()
                if not shortage_df.empty:
                    st.download_button(
                        f"⬇ Download Shortage List for {selected_tag}",
                        data=excel_bytes(shortage_df),
                        file_name=f"shortage_{selected_tag.replace('/','-')}.xlsx",
                        mime="application/vnd.ms-excel",
                        use_container_width=True,
                    )


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2  ·  SESSION ORDER REPORT
# ═══════════════════════════════════════════════════════════════════════════════
with tab_report:
    session_tags = st.session_state.session_tags

    if not session_tags:
        st.markdown("""
        <div style="text-align:center;padding:4rem 2rem;">
          <div style="font-family:'JetBrains Mono',monospace;font-size:2rem;
                      opacity:.15;margin-bottom:.8rem;">📦</div>
          <div style="font-family:'JetBrains Mono',monospace;font-size:.78rem;
                      color:#2E3D58;letter-spacing:.1em;">
            ADD EQUIPMENT TAGS IN THE ENTRY TAB TO GENERATE A REPORT</div>
        </div>""", unsafe_allow_html=True)
    else:
        tag_name_map = eq_master.set_index("Equipment_Tag_No.")["Name"].to_dict()
        tag_loc_map  = eq_master.set_index("Equipment_Tag_No.")["Location"].to_dict()
        tag_sqm_map  = eq_master.set_index("Equipment_Tag_No.")["Total_SQM"].to_dict()

        # ── Session header strip ─────────────────────────────────────────
        st.markdown('<div class="sec-hdr">Session Equipment</div>',
                    unsafe_allow_html=True)
        chips = "".join(
            f'<span class="tag-chip">{t}</span>' for t in session_tags
        )
        st.markdown(
            f'<div style="margin-bottom:.8rem;">{chips}</div>',
            unsafe_allow_html=True,
        )

        # ── Per-equipment breakdown ──────────────────────────────────────
        st.markdown('<div class="sec-hdr" style="margin-top:1rem;">'
                    'Per-Equipment Material Breakdown</div>',
                    unsafe_allow_html=True)

        all_eq_rows = []  # collect for grand total

        for tag in session_tags:
            mat_df = get_tag_demand(tag)
            all_eq_rows.append(mat_df.assign(Equipment_Tag=tag))

            name    = tag_name_map.get(tag, tag)
            loc     = tag_loc_map.get(tag, "—")
            sqm     = tag_sqm_map.get(tag, 0)
            n_short = (mat_df["Shortfall"] > 0).sum()
            tot_sh  = mat_df["Shortfall"].sum()
            overall_pct = (
                mat_df["Available_Qty"].clip(upper=mat_df["Demand_Qty"]).sum() /
                mat_df["Demand_Qty"].sum() * 100
                if mat_df["Demand_Qty"].sum() > 0 else 100
            )

            with st.expander(
                f"{'🔴' if n_short > 0 else '✅'}  {tag}  ·  {name}  "
                f"·  {sqm:,.2f} SQM  ·  "
                f"{'Shortfall on ' + str(n_short) + ' material(s)' if n_short else 'All materials covered'}",
                expanded=False,
            ):
                # Equipment meta
                eq_row = eq_master[eq_master["Equipment_Tag_No."] == tag].iloc[0]
                ci1, ci2, ci3, ci4 = st.columns(4)
                ci1.markdown(f'**Location:** {loc}')
                ci2.markdown(f'**Type:** {eq_row["Type"]}')
                ci3.markdown(f'**Material Spec.:** {eq_row["Material_Spec"] or "—"}')
                ci4.markdown(f'**Total SQM:** `{sqm:,.3f}`')
                st.caption(
                    f'**Lining Systems:** '
                    f'{str(eq_row["Lining_Systems"]).replace(chr(10), " | ")}'
                )
                st.markdown("---")

                # KPIs
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("Total Materials",   len(mat_df))
                k2.metric("Fully Stocked",     (mat_df["Fulfillment_%"] >= 100).sum())
                k3.metric("Need to Order",     n_short)
                k4.metric("Overall Fulfillment", f"{overall_pct:.1f}%")

                # Material table
                fig = plotly_material_table(mat_df, height=70 + len(mat_df) * 30)
                st.plotly_chart(fig, use_container_width=True, key=f"report_mat_{tag}")

        # ── Combined order report ────────────────────────────────────────
        st.markdown('<div class="sec-hdr" style="margin-top:1.8rem;">'
                    '🛒 Combined Procurement List — All Session Equipment</div>',
                    unsafe_allow_html=True)

        combined_df = get_session_combined(session_tags)

        if not combined_df.empty:
            # KPIs for combined
            tot_demand  = combined_df["Demand_Qty"].sum()
            tot_avail   = combined_df["Available_Qty"].sum()
            tot_short   = combined_df["Shortfall"].sum()
            pct_overall = (
                combined_df["Available_Qty"].clip(upper=combined_df["Demand_Qty"]).sum()
                / tot_demand * 100 if tot_demand > 0 else 100
            )
            n_short_mats = (combined_df["Shortfall"] > 0).sum()

            g1, g2, g3, g4, g5 = st.columns(5)
            g1.metric("Equipment in Session",  len(session_tags))
            g2.metric("Total Demand (units)",  f"{tot_demand:,.0f}")
            g3.metric("Available (units)",     f"{tot_avail:,.0f}")
            g4.metric("Total Shortfall",       f"{tot_short:,.1f}")
            g5.metric("Overall Coverage",      f"{pct_overall:.1f}%")

            st.markdown("<br>", unsafe_allow_html=True)

            # Full balance table
            st.markdown('<div class="sec-hdr">Full Material Balance</div>',
                        unsafe_allow_html=True)
            fig_combined = plotly_material_table(
                combined_df, height=90 + len(combined_df) * 30
            )
            st.plotly_chart(fig_combined, use_container_width=True)

            # Shortage-only table
            shortage_only = combined_df[combined_df["Shortfall"] > 0].copy()
            if not shortage_only.empty:
                st.markdown('<div class="sec-hdr" style="margin-top:1.2rem;">'
                            '⚠️ Materials to Order (Shortage Only)</div>',
                            unsafe_allow_html=True)

                # Stacked bar chart
                so_chart = shortage_only.copy()
                so_chart["Label"] = (
                    so_chart["Material_Code"] + "  " +
                    so_chart["Material_Name"].fillna("").str[:20]
                )
                fig_bar = go.Figure()
                fig_bar.add_trace(go.Bar(
                    name="Available",
                    y=so_chart["Label"], x=so_chart["Available_Qty"],
                    orientation="h", marker_color="#10B981", marker_opacity=.75,
                    text=so_chart["Available_Qty"].apply(lambda v: f"{v:,.0f}"),
                    textposition="inside",
                    textfont=dict(family="JetBrains Mono", size=9, color="#fff"),
                ))
                fig_bar.add_trace(go.Bar(
                    name="To Order",
                    y=so_chart["Label"], x=so_chart["Shortfall"],
                    orientation="h", marker_color="#EF4444", marker_opacity=.75,
                    text=so_chart["Shortfall"].apply(lambda v: f"{v:,.1f}"),
                    textposition="inside",
                    textfont=dict(family="JetBrains Mono", size=9, color="#fff"),
                ))
                fig_bar.update_layout(
                    barmode="stack",
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    font=dict(family="JetBrains Mono", size=10, color="#94A3B8"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02,
                                bgcolor="rgba(0,0,0,0)", font=dict(size=10)),
                    margin=dict(l=0, r=60, t=28, b=0),
                    height=max(300, len(so_chart) * 40),
                    xaxis=dict(gridcolor="#1E2E46", zerolinecolor="#1E2E46"),
                    yaxis=dict(gridcolor="#1E2E46"),
                )
                st.plotly_chart(fig_bar, use_container_width=True, key="session_bar_chart")

                # Shopping list table
                shop = shortage_only[["Material_Code", "Material_Name", "UOM",
                                      "Demand_Qty", "Available_Qty",
                                      "Shortfall", "Net_Balance"]].copy()
                shop.columns = ["Code", "Material Name", "UOM",
                                "Total Demand", "Available", "TO ORDER", "Net Balance"]
                st.dataframe(shop, use_container_width=True, hide_index=True)

            # ── Grand total box ──────────────────────────────────────────
            st.markdown("<br>", unsafe_allow_html=True)
            grand_tags = "  ·  ".join(session_tags)
            st.markdown(f"""
            <div class="grand-box">
              <div style="font-family:'JetBrains Mono',monospace;font-size:.62rem;
                          letter-spacing:.15em;text-transform:uppercase;
                          color:#F59E0B;margin-bottom:.8rem;">
                ⭐ Grand Total Summary — {len(session_tags)} Equipment</div>
              <div style="font-size:.78rem;color:#94A3B8;margin-bottom:.8rem;">
                {grand_tags}</div>
              <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:.8rem;">
                <div><div style="font-family:'JetBrains Mono',monospace;
                     font-size:1.4rem;font-weight:700;color:#F1F5F9;">
                     {len(session_tags)}</div>
                     <div style="font-size:.68rem;letter-spacing:.08em;
                     text-transform:uppercase;color:#475569;">Equipment</div></div>
                <div><div style="font-family:'JetBrains Mono',monospace;
                     font-size:1.4rem;font-weight:700;color:#F1F5F9;">
                     {len(combined_df)}</div>
                     <div style="font-size:.68rem;letter-spacing:.08em;
                     text-transform:uppercase;color:#475569;">Materials</div></div>
                <div><div style="font-family:'JetBrains Mono',monospace;
                     font-size:1.4rem;font-weight:700;color:#EF4444;">
                     {n_short_mats}</div>
                     <div style="font-size:.68rem;letter-spacing:.08em;
                     text-transform:uppercase;color:#475569;">To Procure</div></div>
                <div><div style="font-family:'JetBrains Mono',monospace;
                     font-size:1.4rem;font-weight:700;color:#EF4444;">
                     {tot_short:,.1f}</div>
                     <div style="font-size:.68rem;letter-spacing:.08em;
                     text-transform:uppercase;color:#475569;">Total Shortfall Units</div></div>
              </div>
            </div>""", unsafe_allow_html=True)

            # Download buttons
            st.markdown("<br>", unsafe_allow_html=True)
            d1, d2 = st.columns(2)
            with d1:
                st.download_button(
                    "⬇ Download Full Session Report",
                    data=excel_bytes(combined_df),
                    file_name="session_full_report.xlsx",
                    mime="application/vnd.ms-excel",
                    use_container_width=True,
                )
            with d2:
                if not shortage_only.empty:
                    st.download_button(
                        "⬇ Download Order List Only",
                        data=excel_bytes(shortage_only),
                        file_name="session_order_list.xlsx",
                        mime="application/vnd.ms-excel",
                        use_container_width=True,
                    )


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3  ·  LOCATION REPORT  (full project, all equipment, grouped by location)
# ═══════════════════════════════════════════════════════════════════════════════
with tab_location:

    st.markdown('<div class="sec-hdr">📍 All Equipment — Grouped by Location</div>',
                unsafe_allow_html=True)
    st.markdown(
        '<div style="font-size:.82rem;color:#475569;margin-bottom:1.2rem;">'
        'Read-only view of every equipment in the project, '
        'segregated by site location. Use the Entry tab to build a targeted session.</div>',
        unsafe_allow_html=True,
    )

    # Project-wide KPIs
    all_demand  = demand_df.merge(
        inv_clean[["Material_Code", "Available_Qty"]], on="Material_Code", how="left"
    )
    all_demand["Available_Qty"] = all_demand["Available_Qty"].fillna(0)
    all_demand["Shortfall"]     = (all_demand["Demand_Qty"] - all_demand["Available_Qty"]).clip(lower=0)

    pk1, pk2, pk3, pk4 = st.columns(4)
    pk1.metric("Total Equipment",     len(eq_master))
    pk2.metric("Total Materials",     demand_df["Material_Code"].nunique())
    pk3.metric("Total Demand (units)",f"{all_demand['Demand_Qty'].sum():,.0f}")
    pk4.metric("Total Shortfall",     f"{all_demand['Shortfall'].sum():,.0f}")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Iterate locations ─────────────────────────────────────────────────
    loc_color_map = {
        "Brown Field": ("#3B82F6", "loc-bf"),
        "TRAIN J":     ("#F59E0B", "loc-tj"),
        "TRAIN K":     ("#10B981", "loc-tk"),
    }

    for loc in LOCATION_ORDER:
        loc_equip = eq_master[eq_master["Location"] == loc]["Equipment_Tag_No."].tolist()
        if not loc_equip:
            continue

        color, badge_cls = loc_color_map.get(loc, ("#94A3B8", "loc-bf"))

        # Location header
        loc_demand = all_demand[all_demand["Equipment_Tag_No."].isin(loc_equip)]
        loc_short  = loc_demand["Shortfall"].sum()
        loc_total  = loc_demand["Demand_Qty"].sum()
        loc_pct    = (
            loc_demand["Available_Qty"].clip(upper=loc_demand["Demand_Qty"]).sum()
            / loc_total * 100 if loc_total > 0 else 100
        )

        st.markdown(
            f'<div style="display:flex;align-items:center;gap:1rem;'
            f'margin:1.2rem 0 .6rem;">'
            f'<span class="loc-badge {badge_cls}" style="font-size:.78rem;'
            f'padding:.3rem .9rem;">{loc}</span>'
            f'<span style="font-family:\'JetBrains Mono\',monospace;font-size:.75rem;'
            f'color:#475569;">{len(loc_equip)} equipment</span>'
            f'<span style="font-family:\'JetBrains Mono\',monospace;font-size:.75rem;'
            f'color:{"#EF4444" if loc_short > 0 else "#10B981"};">'
            f'Shortfall: {loc_short:,.1f} units</span>'
            f'<span style="font-family:\'JetBrains Mono\',monospace;font-size:.75rem;'
            f'color:#94A3B8;">Coverage: {loc_pct:.1f}%</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

        # Equipment rows within this location
        for tag in loc_equip:
            eq_row  = eq_master[eq_master["Equipment_Tag_No."] == tag].iloc[0]
            mat_df  = get_tag_demand(tag)
            n_short = (mat_df["Shortfall"] > 0).sum()
            sqm     = eq_row["Total_SQM"]
            status_icon = "🔴" if n_short > 0 else "✅"

            with st.expander(
                f"{status_icon}  {tag}  ·  {eq_row['Name']}  ·  "
                f"{sqm:,.2f} SQM  ·  "
                f"{'Shortfall on ' + str(n_short) + ' material(s)' if n_short else 'All materials covered'}",
                expanded=False,
            ):
                ci1, ci2, ci3 = st.columns(3)
                ci1.markdown(f'**Type:** {eq_row["Type"]}')
                ci2.markdown(f'**Description:** {eq_row["Description"] or "—"}')
                ci3.markdown(f'**Material Spec.:** {eq_row["Material_Spec"] or "—"}')
                st.caption(
                    f'**Lining Systems:** '
                    f'{str(eq_row["Lining_Systems"]).replace(chr(10), " | ")}'
                )
                st.markdown("---")

                if mat_df.empty:
                    st.warning("No recipe data found.")
                else:
                    m1, m2, m3, m4 = st.columns(4)
                    tot_d = mat_df["Demand_Qty"].sum()
                    tot_s = mat_df["Shortfall"].sum()
                    ov_pct = (
                        mat_df["Available_Qty"].clip(upper=mat_df["Demand_Qty"]).sum()
                        / tot_d * 100 if tot_d > 0 else 100
                    )
                    m1.metric("Total Materials",   len(mat_df))
                    m2.metric("Fully Stocked",     (mat_df["Fulfillment_%"] >= 100).sum())
                    m3.metric("Need to Order",     n_short)
                    m4.metric("Coverage",          f"{ov_pct:.1f}%")

                    fig = plotly_material_table(mat_df, height=70 + len(mat_df) * 30)
                    st.plotly_chart(fig, use_container_width=True, key=f"loc_mat_{loc}_{tag}")

                # Quick-add to session button inside location report
                already_in = tag in st.session_state.session_tags
                if already_in:
                    st.markdown(
                        '<span style="font-family:\'JetBrains Mono\',monospace;'
                        'font-size:.7rem;color:#10B981;">✓ Already in session</span>',
                        unsafe_allow_html=True,
                    )
                else:
                    if st.button(f"＋ Add {tag} to Session",
                                 key=f"loc_add_{tag}"):
                        st.session_state.session_tags.append(tag)
                        st.rerun()

        st.markdown(
            f'<div style="border-bottom:1px solid #1E2E46;margin:1rem 0;"></div>',
            unsafe_allow_html=True,
        )