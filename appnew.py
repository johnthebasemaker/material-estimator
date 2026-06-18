"""
Smart Material Estimator · app.py (v3)
Run: streamlit run app.py
"""
import io, os, sys, sqlite3, base64
from PIL import Image as _PILImage
from datetime import date, datetime
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from validate_data     import clean_inventory, clean_recipe, clean_equipment
from allocation_engine import build_demand_matrix

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Smart Material Estimator & Planner",
                   page_icon="", layout="wide",
                   initial_sidebar_state="expanded")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PATH_A   = os.path.join(BASE_DIR, "Materials_DetailsAvailable_Qty.xlsx")
PATH_B   = os.path.join(BASE_DIR, "For_1_SQM.xlsx")
PATH_C   = os.path.join(BASE_DIR, "Equipment.xlsx")
SHEET_A, SHEET_B, SHEET_C = "Materials", "LINING SYSTEM MATERIAL CONSM", "Data Input"
LOCATION_ORDER = ["Brown Field", "TRAIN J", "TRAIN K"]
DB_PATH = os.path.join(BASE_DIR, "sme_database.db")
LOGO_PATH = os.path.join(BASE_DIR, "logo.png")

@st.cache_data(show_spinner=False)
def _logo_b64() -> str:
    if os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""

# ─────────────────────────────────────────────────────────────────────────────
# STYLES
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:ital,opsz,wght@0,14..32,300;0,14..32,400;0,14..32,500;0,14..32,600;0,14..32,700&family=JetBrains+Mono:wght@400;500;600;700&display=swap');

/* ── DESIGN TOKENS ── */
:root {
  --bg0: var(--background-color);
  --bg1: var(--secondary-background-color);
  --bg2: var(--secondary-background-color);
  --bg3: var(--background-color);
  --bg4: var(--secondary-background-color);
  --border:  rgba(128,128,128,.18);
  --border2: rgba(128,128,128,.28);

  --amber:      #F59E0B;
  --amber2:     #FCD34D;
  --amber3:     #D97706;
  --amber-bg:   rgba(245,158,11,.11);
  --amber-glow: rgba(245,158,11,.22);
  --green:      #10B981;
  --green-bg:   rgba(16,185,129,.11);
  --green-glow: rgba(16,185,129,.18);
  --red:        #EF4444;
  --red-bg:     rgba(239,68,68,.11);
  --orange:     #F97316;
  --orange-bg:  rgba(249,115,22,.11);
  --yellow:     #EAB308;
  --yellow-bg:  rgba(234,179,8,.11);
  --blue:       #3B82F6;
  --blue-bg:    rgba(59,130,246,.11);
  --blue-glow:  rgba(59,130,246,.18);

  --t0: var(--text-color);
  --t1: var(--text-color);
  --t2: color-mix(in srgb, var(--text-color) 82%, transparent);
  --t3: color-mix(in srgb, var(--text-color) 62%, transparent);
  --t4: color-mix(in srgb, var(--text-color) 45%, transparent);
  --t5: color-mix(in srgb, var(--text-color) 30%, transparent);

  --r-sm: 6px;
  --r-md: 10px;
  --r-lg: 14px;
}

/* ── CUSTOM SCROLLBAR ── */
::-webkit-scrollbar { width:5px; height:5px; }
::-webkit-scrollbar-track { background:transparent; }
::-webkit-scrollbar-thumb { background:var(--border2); border-radius:99px; }
::-webkit-scrollbar-thumb:hover { background:var(--amber); }

/* ── BASE ── */
html,body,[class*="css"] { font-family:'Inter',sans-serif!important; background:var(--bg0)!important; color:var(--t1); }
.main .block-container { padding-bottom:1.5rem; margin-top:0!important; max-width:1500px; }

/* ── SIDEBAR ── */
[data-testid="stSidebar"] { background:var(--bg1)!important; border-right:none!important; box-shadow:4px 0 24px rgba(0,0,0,.18)!important; border-radius:0 var(--r-lg) var(--r-lg) 0!important; }
[data-testid="stSidebar"] * { font-family:'Inter',sans-serif!important; }
[data-testid="stSidebar"]::before {
  content:''; display:block; height:3px;
  background:linear-gradient(90deg,var(--amber3) 0%,var(--amber) 50%,transparent 100%);
  position:sticky; top:0; z-index:1;
}

/* ── HEADER CHROME ── */
header[data-testid="stHeader"] { height:0!important; min-height:0!important; padding:0!important; background:transparent!important; overflow:visible!important; }
[data-testid="collapsedControl"] { display:flex!important; visibility:visible!important; opacity:1!important; position:fixed!important; top:.45rem!important; left:.5rem!important; z-index:10001!important; }
[data-testid="stAppViewContainer"] { padding-top:0!important; }

/* ── MAIN CONTAINER TOP PADDING (clears fixed sticky header) ── */
[data-testid="stAppViewBlockContainer"],
[data-testid="stMainBlockContainer"],
section.main > div.block-container { padding-top: 78px !important; }

/* ── HAMBURGER ICON (shows when sidebar is COLLAPSED — sits in sticky header, top-left) ── */
/* Covers BOTH old (`collapsedControl`) and new (`stSidebarCollapsedControl`) Streamlit testids */
[data-testid="collapsedControl"],
[data-testid="stSidebarCollapsedControl"],
[data-testid="stSidebarCollapseButton"]:not([data-testid="stSidebar"] *) {
  position: fixed !important;
  top: .9rem !important; left: 1rem !important;
  z-index: 1000001 !important;
  display: flex !important; visibility: visible !important; opacity: 1 !important;
  width: 38px !important; height: 38px !important;
}
[data-testid="collapsedControl"] button,
[data-testid="stSidebarCollapsedControl"] button,
[data-testid="stSidebarCollapsedControl"] > button {
  background: var(--bg1) !important;
  border: 1px solid var(--amber) !important;
  border-radius: var(--r-sm) !important;
  width: 38px !important; height: 38px !important;
  display: flex !important; align-items: center !important; justify-content: center !important;
  cursor: pointer !important; transition: all .12s !important;
  box-shadow: 0 0 0 1px var(--amber), 0 2px 10px var(--amber-glow) !important;
  padding: 0 !important;
  position: relative !important;
  font-size: 0 !important;
  color: transparent !important;
  overflow: hidden !important;
}
/* Nuke every native icon/text node inside (svg, span, "keyboard_double_arrow_right" text, etc.) */
[data-testid="collapsedControl"] button *,
[data-testid="stSidebarCollapsedControl"] button *,
[data-testid="stSidebarCollapsedControl"] > button * {
  display: none !important;
  visibility: hidden !important;
  font-size: 0 !important;
  width: 0 !important; height: 0 !important;
  opacity: 0 !important;
}
[data-testid="collapsedControl"] button::after,
[data-testid="stSidebarCollapsedControl"] button::after,
[data-testid="stSidebarCollapsedControl"] > button::after {
  content: '\2630' !important;             /* ☰ — Unicode trigram */
  font-family: 'Inter', 'Helvetica', sans-serif !important;
  font-size: 22px !important;
  font-weight: 700 !important;
  color: var(--amber) !important;
  line-height: 1 !important;
  display: block !important;
  visibility: visible !important;
  opacity: 1 !important;
  position: absolute !important;
  top: 50% !important; left: 50% !important;
  transform: translate(-50%, -50%) !important;
  width: auto !important; height: auto !important;
}
[data-testid="collapsedControl"] button:hover,
[data-testid="stSidebarCollapsedControl"] button:hover {
  background: var(--amber-bg) !important;
  box-shadow: 0 0 0 1px var(--amber), 0 4px 14px var(--amber-glow) !important;
}
[data-testid="collapsedControl"] button:hover::after,
[data-testid="stSidebarCollapsedControl"] button:hover::after {
  color: var(--amber2) !important;
}

/* ── SIDEBAR COLLAPSE BUTTON (shows when sidebar is OPEN — kills "keyboard_double_arrow_left" text) ── */
[data-testid="stSidebarCollapseButton"] button,
[data-testid="stSidebarCollapsedControl"] button,
[data-testid="stSidebar"] button[kind="header"],
[data-testid="stSidebarHeader"] button {
  font-size: 0 !important;
  color: transparent !important;
  position: relative !important;
  background: transparent !important;
  border: none !important;
  width: 32px !important; height: 32px !important;
  display: flex !important; align-items: center !important; justify-content: center !important;
}
[data-testid="stSidebarCollapseButton"] button > *,
[data-testid="stSidebarCollapseButton"] button svg,
[data-testid="stSidebarCollapseButton"] button span,
[data-testid="stSidebarCollapsedControl"] button > *,
[data-testid="stSidebarCollapsedControl"] button svg,
[data-testid="stSidebarCollapsedControl"] button span,
[data-testid="stSidebar"] button[kind="header"] > *,
[data-testid="stSidebar"] button[kind="header"] svg,
[data-testid="stSidebar"] button[kind="header"] span,
[data-testid="stSidebarHeader"] button > *,
[data-testid="stSidebarHeader"] button svg,
[data-testid="stSidebarHeader"] button span {
  display: none !important;
  visibility: hidden !important;
  font-size: 0 !important;
  width: 0 !important; height: 0 !important;
}
[data-testid="stSidebarCollapseButton"] button::after,
[data-testid="stSidebarCollapsedControl"] button::after,
[data-testid="stSidebar"] button[kind="header"]::after,
[data-testid="stSidebarHeader"] button::after {
  content: '✕' !important;
  font-family: 'Inter', sans-serif !important;
  font-size: 18px !important;
  font-weight: 700 !important;
  color: var(--amber) !important;
  line-height: 1 !important;
  display: block !important;
  position: absolute !important;
  top: 50% !important; left: 50% !important;
  transform: translate(-50%, -50%) !important;
}
[data-testid="stSidebarCollapseButton"] button:hover::after,
[data-testid="stSidebarCollapsedControl"] button:hover::after,
[data-testid="stSidebar"] button[kind="header"]:hover::after,
[data-testid="stSidebarHeader"] button:hover::after {
  color: var(--amber2) !important;
}

/* ── STICKY HEADER (fixed → persists across every tab while scrolling) ── */
.sticky-header-wrap {
  position: fixed !important;
  top: 0 !important;
  left: 0 !important;
  right: 0 !important;
  z-index: 999990 !important;
  background-color: var(--background-color) !important;
  width: 100%;
  padding: .55rem 1.5rem .55rem 4rem;
  padding-bottom: 10px;
  margin-bottom: 0;
  border-bottom: 1px solid rgba(128,128,128,0.2);
  box-shadow: 0 2px 24px rgba(0,0,0,.07);
  backdrop-filter: blur(8px);
  -webkit-backdrop-filter: blur(8px);
}
/* When sidebar is OPEN, shrink the fixed header so it doesn't overlap the sidebar */
[data-testid="stSidebar"][aria-expanded="true"] ~ section .sticky-header-wrap,
[data-testid="stAppViewContainer"]:has([data-testid="stSidebar"][aria-expanded="true"]) .sticky-header-wrap {
  left: var(--sidebar-width, 244px) !important;
}
.sticky-header-wrap::before {
  content:''; position:absolute; top:0; left:0; right:0; height:2px;
  background:linear-gradient(90deg,var(--amber3) 0%,var(--amber) 40%,var(--amber2) 70%,transparent 100%);
}

/* ── TABS (stick just under the fixed header) ── */
[data-testid="stTabs"] > div:first-of-type {
  position: sticky !important; top: 78px !important; z-index: 999985 !important;
  background: var(--bg0) !important;
  padding: .35rem .3rem;
  border-bottom: 1px solid var(--border);
}
[data-testid="stTabs"] [data-baseweb="tab-list"] { background:transparent; border-bottom:none; gap:.1rem; padding:0; }
[data-testid="stTabs"] [data-baseweb="tab"] {
  font-family:'JetBrains Mono',monospace!important;
  font-size:.67rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase;
  color:var(--t4)!important;
  padding:.5rem 1.1rem;
  border-radius:var(--r-sm);
  border-bottom:2px solid transparent;
  transition:all .18s;
}
[data-testid="stTabs"] [data-baseweb="tab"]:hover { color:var(--t2)!important; background:var(--amber-bg); }
[data-testid="stTabs"] [aria-selected="true"] { color:var(--amber)!important; border-bottom:2px solid var(--amber)!important; background:var(--amber-bg)!important; }

/* ── BUTTONS ── */
.stButton>button {
  font-family:'JetBrains Mono',monospace!important;
  font-size:.68rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase;
  background:linear-gradient(135deg,var(--amber3) 0%,var(--amber) 100%)!important;
  color:#000!important; border:none!important;
  border-radius:var(--r-sm)!important;
  padding:.5rem 1.2rem!important;
  transition:all .18s!important;
  box-shadow:0 2px 10px var(--amber-glow)!important;
}
.stButton>button:hover { background:linear-gradient(135deg,var(--amber) 0%,var(--amber2) 100%)!important; transform:translateY(-2px)!important; box-shadow:0 6px 20px var(--amber-glow)!important; }
.stButton>button:active { transform:translateY(0)!important; }

/* ── NATIVE METRICS ── */
[data-testid="stMetric"] {
  background:var(--bg2); border:1px solid var(--border);
  border-radius:var(--r-md); padding:.9rem 1rem!important;
  transition:all .12s; cursor:default;
  position:relative; overflow:hidden;
}
[data-testid="stMetric"]::before {
  content:''; position:absolute; top:0; left:0;
  width:3px; height:100%; background:var(--amber); border-radius:99px 0 0 99px;
}
[data-testid="stMetric"]:hover { border-color:var(--amber)!important; box-shadow:0 0 0 1px var(--amber),0 4px 22px var(--amber-glow)!important; transform:translateY(-2px)!important; }
[data-testid="stMetricLabel"] { font-family:'JetBrains Mono',monospace!important; font-size:.58rem!important; letter-spacing:.13em; text-transform:uppercase; color:var(--t4)!important; padding-left:.5rem; }
[data-testid="stMetricValue"] { font-family:'JetBrains Mono',monospace!important; font-size:1.9rem!important; font-weight:700!important; color:var(--t0)!important; padding-left:.5rem; }
[data-testid="stMetricDelta"] { font-size:.72rem!important; }

/* ── ALERT BOXES ── */
[data-testid="stInfo"]    { background:var(--blue-bg)!important;   border-left:3px solid var(--blue)!important;   border-radius:var(--r-sm)!important; color:var(--t1)!important; }
[data-testid="stSuccess"] { background:var(--green-bg)!important;  border-left:3px solid var(--green)!important;  border-radius:var(--r-sm)!important; color:var(--t1)!important; }
[data-testid="stWarning"] { background:var(--amber-bg)!important;  border-left:3px solid var(--amber)!important;  border-radius:var(--r-sm)!important; color:var(--t1)!important; }
[data-testid="stError"]   { background:var(--red-bg)!important;    border-left:3px solid var(--red)!important;    border-radius:var(--r-sm)!important; color:var(--t1)!important; }

/* ── EXPANDERS ── */
[data-testid="stExpander"] { background:var(--bg2)!important; border:1px solid var(--border)!important; border-radius:var(--r-md)!important; overflow:hidden; }
[data-testid="stExpander"] summary { color:var(--t1)!important; font-weight:600; padding:.7rem 1rem!important; }
[data-testid="stExpander"] summary:hover { background:var(--amber-bg)!important; }

/* ── SELECT / INPUT ── */
[data-baseweb="select"]>div,[data-baseweb="input"]>div { background:var(--bg2)!important; border-color:var(--border2)!important; border-radius:var(--r-sm)!important; color:var(--t0)!important; }

/* ── DATAFRAME ── */
[data-testid="stDataFrame"] { border:1px solid var(--border)!important; border-radius:var(--r-md); overflow:hidden; }
[data-testid="stDataFrame"] [role="columnheader"],[data-testid="stDataFrame"] th,
[data-testid="stDataFrame"] .dvn-column-header,[data-testid="stDataFrame"] .dvn-header-row [role="cell"] {
  font-weight:700!important; color:var(--text-color)!important; background:var(--bg2)!important;
}

/* ── DIVIDERS ── */
hr { border:none!important; height:1px!important; background:linear-gradient(90deg,transparent,var(--border2),transparent)!important; margin:.8rem 0!important; }

/* ── RADIO BUTTONS ── */
[data-testid="stRadio"] label { font-family:'JetBrains Mono',monospace!important; font-size:.72rem!important; font-weight:600!important; }

/* ── CUSTOM COMPONENTS ── */

.sec-hdr {
  font-family:'JetBrains Mono',monospace; font-size:.6rem; font-weight:700;
  letter-spacing:.18em; text-transform:uppercase; color:var(--t5);
  border-bottom:1px solid var(--border); padding-bottom:.35rem; margin-bottom:.9rem;
  position:relative;
}
.sec-hdr::after { content:''; position:absolute; bottom:-1px; left:0; width:36px; height:1px; background:var(--amber); }

.card { background:var(--bg2); border:1px solid var(--border); border-radius:var(--r-md); padding:1rem 1.2rem; margin-bottom:.6rem; transition:box-shadow .2s; }
.card:hover { box-shadow:0 2px 18px rgba(0,0,0,.1); }
.card-amber { border-left:3px solid var(--amber); }
.card-green { border-left:3px solid var(--green); }
.card-blue  { border-left:3px solid var(--blue); }

.loc-badge { display:inline-flex; align-items:center; flex-shrink:0; white-space:nowrap!important; font-family:'JetBrains Mono',monospace; font-size:.62rem; font-weight:700; letter-spacing:.07em; text-transform:uppercase; padding:.2rem .7rem; border-radius:99px; }
.loc-bf { background:rgba(59,130,246,.12); color:var(--blue);  border:1px solid rgba(59,130,246,.22); }
.loc-tj { background:rgba(245,158,11,.12); color:var(--amber); border:1px solid rgba(245,158,11,.22); }
.loc-tk { background:rgba(16,185,129,.12); color:var(--green); border:1px solid rgba(16,185,129,.22); }

.pill { display:inline-flex; align-items:center; flex-shrink:0; white-space:nowrap!important; font-family:'JetBrains Mono',monospace; font-size:.68rem; font-weight:600; padding:.18rem .6rem; border-radius:99px; }
.pill-g { background:var(--green-bg); color:var(--green); border:1px solid rgba(16,185,129,.18); }
.pill-y { background:var(--yellow-bg); color:var(--yellow); border:1px solid rgba(234,179,8,.18); }
.pill-o { background:var(--orange-bg); color:var(--orange); border:1px solid rgba(249,115,22,.18); }
.pill-r { background:var(--red-bg);    color:var(--red);    border:1px solid rgba(239,68,68,.18); }

.tag-chip { display:inline-flex; align-items:center; flex-shrink:0; white-space:nowrap!important; font-family:'JetBrains Mono',monospace; font-size:.7rem; background:var(--amber-bg); color:var(--amber); border:1px solid rgba(245,158,11,.2); border-radius:var(--r-sm); padding:.15rem .55rem; margin:.12rem; }
.syscode-block { background:var(--bg3); border:1px solid var(--border2); border-radius:var(--r-md); padding:.7rem .9rem; margin:.35rem 0; }
.syscode-hdr { display:flex; align-items:center; gap:.7rem; margin-bottom:.5rem; flex-wrap:nowrap; overflow-x:auto; }
.code-badge { display:inline-flex; align-items:center; flex-shrink:0; white-space:nowrap!important; font-family:'JetBrains Mono',monospace; font-size:.68rem; font-weight:700; background:var(--amber-bg); color:var(--amber); border:1px solid rgba(245,158,11,.22); border-radius:var(--r-sm); padding:.22rem .6rem; }
.session-equip { background:var(--bg2); border:1px solid var(--border); border-radius:var(--r-md); padding:.9rem 1rem; margin-bottom:.5rem; overflow:hidden; }
.drag-handle { font-size:1rem; color:var(--t5); cursor:grab; user-select:none; padding:.2rem .4rem; }

.grand-box {
  background:linear-gradient(135deg,var(--amber-bg) 0%,var(--bg2) 60%);
  border:1px solid var(--border2); border-left:3px solid var(--amber);
  border-radius:var(--r-lg); padding:1.2rem 1.5rem;
  position:relative; overflow:hidden;
}
.grand-box::after {
  content:''; position:absolute; top:-50px; right:-50px;
  width:140px; height:140px;
  background:radial-gradient(circle,var(--amber-glow) 0%,transparent 70%);
  pointer-events:none;
}

.status-dot-g::before { content:"●"; color:var(--green);  margin-right:.4rem; }
.status-dot-o::before { content:"●"; color:var(--orange); margin-right:.4rem; }
.status-dot-y::before { content:"●"; color:var(--yellow); margin-right:.4rem; }
.status-dot-r::before { content:"●"; color:var(--red);    margin-right:.4rem; }

/* ── POPOVER PERFORMANCE: kill open/close animations everywhere ── */
[data-baseweb="popover"],
[data-baseweb="popover"] *,
[data-testid="stPopover"],
[data-testid="stPopover"] *,
[data-testid="stPopoverBody"],
[data-baseweb="layer"] > div {
  transition: none !important;
  animation: none !important;
  animation-duration: 0s !important;
  transition-duration: 0s !important;
}

/* ── TOOLTIPS: keep them strictly below the fixed header so they never overlap it ── */
[data-baseweb="tooltip"],
[role="tooltip"],
[data-testid="stTooltipContent"],
[data-testid="stTooltipHoverTarget"] + div {
  z-index: 999985 !important;
}

/* ── KPI POPOVER BUTTONS ── */
[data-testid="stPopover"] button {
  background:var(--bg2)!important; border:1px solid var(--border)!important;
  border-radius:var(--r-md)!important; padding:.9rem 1rem!important;
  height:auto!important; min-height:80px!important;
  text-align:left!important; white-space:pre-wrap!important;
  transition:all .12s!important; color:var(--t0)!important;
  position:relative!important; overflow:hidden!important;
}
[data-testid="stPopover"] button::before {
  content:''; position:absolute; top:0; left:0;
  width:2px; height:100%; background:var(--amber); border-radius:99px 0 0 99px;
}
[data-testid="stPopover"] button:hover {
  background:var(--bg3)!important; border-color:var(--amber)!important;
  box-shadow:0 0 0 1px var(--amber),0 4px 22px var(--amber-glow)!important;
  transform:translateY(-2px)!important;
}
[data-testid="stPopover"] button p { font-family:'JetBrains Mono',monospace!important; font-size:.88rem!important; color:var(--t0)!important; line-height:1.5!important; }

/* ── MOBILE ── */
@media (max-width:768px) {
  .sticky-header-wrap { padding: .5rem .6rem .3rem .6rem !important; }
  [data-testid="collapsedControl"] { top: .5rem !important; left: .5rem !important; }
  [data-testid="stAppViewBlockContainer"],
  [data-testid="stMainBlockContainer"],
  section.main > div.block-container { padding-top: 70px !important; padding-left: .5rem !important; padding-right: .5rem !important; }
  [data-testid="stTabs"] > div:first-of-type { top: 64px !important; }
  [data-testid="stMetricValue"] { font-size:1.3rem!important; }
  [data-testid="stMetricLabel"] { font-size:.55rem!important; }
  .syscode-hdr { flex-wrap:wrap!important; overflow-x:visible!important; gap:.4rem!important; }
  .card { padding:.6rem .8rem!important; }
  .code-badge { font-size:.6rem!important; }
  .loc-badge { font-size:.56rem!important; }
}

/* ════════════════════════════════════════════════════════════════════════
   DESIGN INTEGRATION — appended additions (Claude design port)
   ════════════════════════════════════════════════════════════════════════ */

/* ── FROZEN TABLE HEADERS (sticky thead inside scrollable containers) ── */
[data-testid="stDataFrame"] [role="grid"] [role="row"]:first-child,
[data-testid="stDataFrame"] thead,
[data-testid="stDataFrame"] thead tr,
[data-testid="stDataFrame"] thead th {
  position: sticky !important;
  top: 0 !important;
  z-index: 5 !important;
  background: var(--bg2) !important;
}
.sme-scroll-table {
  max-height: 520px;
  overflow: auto;
  border: 1px solid var(--border);
  border-radius: var(--r-md);
}
.sme-scroll-table table { width:100%; border-collapse:collapse; font-size:.78rem; }
.sme-scroll-table thead th {
  position: sticky; top: 0; z-index: 4;
  background: var(--bg2);
  font-family:'JetBrains Mono',monospace;
  font-size:.6rem; font-weight:700; letter-spacing:.06em; text-transform:uppercase;
  color: var(--t4);
  padding:.6rem .75rem; text-align:left;
  border-bottom: 1px solid var(--border2);
  white-space:nowrap;
}
.sme-scroll-table tbody td {
  padding:.5rem .75rem;
  border-bottom: 1px solid var(--border);
  color: var(--t1);
}
.sme-scroll-table tbody tr:nth-child(even) { background: var(--bg3); }

/* ── LOGIN FORM POLISH (matches design's gradient panel) ── */
.sme-login-shell {
  min-height: 70vh;
  background: radial-gradient(ellipse at 40% 20%,
              color-mix(in srgb, var(--amber) 9%, var(--bg0)) 0%,
              var(--bg0) 70%);
  display:flex; align-items:center; justify-content:center;
}
.sme-login-card {
  width: min(420px, 92vw);
  background: var(--bg1);
  border: 1px solid var(--border);
  border-radius: 16px;
  padding: 40px 36px;
  box-shadow: 0 30px 60px rgba(0,0,0,.45);
  position: relative; overflow: hidden;
}
.sme-login-card::before {
  content:''; position:absolute; top:0; left:0; right:0; height:3px;
  background: linear-gradient(90deg, var(--amber3), var(--amber), var(--amber2));
}
.sme-login-title {
  font-family:'Inter',sans-serif;
  font-size: 21px; font-weight: 800;
  color: var(--t0); letter-spacing:-.5px;
  text-align: center; margin-top: 8px;
}
.sme-login-sub {
  text-align:center; font-size:10px; color: var(--t5);
  text-transform:uppercase; letter-spacing:1.8px; margin-top:5px;
}
.sme-login-shell .stTextInput input {
  background: var(--bg0) !important;
  border: 1px solid var(--border2) !important;
  border-radius: 9px !important;
  padding: 12px 14px !important;
  color: var(--t0) !important;
  font-family:'Inter',sans-serif !important;
}
.sme-login-shell .stButton>button {
  background: linear-gradient(135deg, var(--amber3), var(--amber)) !important;
  color: #fff !important;
  border-radius: 9px !important;
  padding: 12px !important;
  font-size: 15px !important;
  font-weight: 700 !important;
  letter-spacing: .2px !important;
}

/* ── KPI POPOVER REFINEMENT (match design's amber-card on hover) ── */
[data-testid="stPopover"] button {
  background: linear-gradient(180deg, var(--bg2), var(--bg3)) !important;
}
[data-testid="stPopover"] button p:first-child {
  font-family:'JetBrains Mono',monospace !important;
  font-size:.56rem !important;
  font-weight:700 !important;
  letter-spacing:.13em !important;
  text-transform:uppercase !important;
  color: var(--t4) !important;
  margin-bottom:.35rem !important;
  line-height:1.2 !important;
}
[data-testid="stPopover"] button p:nth-child(2) {
  font-family:'JetBrains Mono',monospace !important;
  font-size: 1.45rem !important;
  font-weight: 800 !important;
  color: var(--amber) !important;
  line-height: 1.05 !important;
  margin-bottom: .25rem !important;
}
[data-testid="stPopover"] button p:nth-child(3) {
  font-family:'Inter',sans-serif !important;
  font-size:.62rem !important;
  color: var(--t5) !important;
}

/* ── SIDEBAR REFINEMENTS (location/session rows from design) ── */
[data-testid="stSidebar"] hr { margin:.6rem 0 !important; }
[data-testid="stSidebar"] .stButton>button {
  background: transparent !important;
  color: var(--t4) !important;
  border: 1px solid var(--border2) !important;
  box-shadow: none !important;
  font-size:.6rem !important;
}
[data-testid="stSidebar"] .stButton>button:hover {
  background: var(--amber-bg) !important;
  border-color: var(--amber) !important;
  color: var(--amber) !important;
  transform: none !important;
}

/* ── DESIGN SVG GAUGE + HBAR (containers) ── */
.sme-viz-card {
  background: var(--bg2);
  border: 1px solid var(--border);
  border-radius: var(--r-md);
  padding: 18px;
  margin-bottom: 12px;
}
.sme-viz-card svg { display:block; width:100%; }
.sme-viz-legend {
  display:flex; gap:18px; margin-top:6px;
  font-size:10px; color: var(--t4);
  font-family:'JetBrains Mono',monospace;
}

/* ── THEME TOGGLE (explicit light mode overrides) ── */
html[data-sme-theme="light"] {
  --bg0: #F8FAFC !important;
  --bg1: #FFFFFF !important;
  --bg2: #FFFFFF !important;
  --bg3: #F1F5F9 !important;
  --bg4: #FFFFFF !important;
  --border:  rgba(15,23,42,.08) !important;
  --border2: rgba(15,23,42,.18) !important;
  --t0: #0F172A !important;
  --t1: #1E293B !important;
  --t2: rgba(15,23,42,.78) !important;
  --t3: rgba(15,23,42,.62) !important;
  --t4: rgba(15,23,42,.48) !important;
  --t5: rgba(15,23,42,.34) !important;
}
html[data-sme-theme="light"] body,
html[data-sme-theme="light"] [data-testid="stAppViewContainer"],
html[data-sme-theme="light"] .main .block-container {
  background: #F8FAFC !important;
  color: #1E293B !important;
}
html[data-sme-theme="light"] [data-testid="stSidebar"] {
  background: #FFFFFF !important;
  border-right: 1px solid rgba(15,23,42,.08) !important;
}
html[data-sme-theme="light"] .sticky-header-wrap {
  background-color: rgba(255,255,255,.92) !important;
  border-bottom: 1px solid rgba(15,23,42,.08) !important;
}
html[data-sme-theme="light"] [data-testid="stDataFrame"] thead th,
html[data-sme-theme="light"] .sme-scroll-table thead th {
  background: #F1F5F9 !important;
  color: #475569 !important;
}
html[data-sme-theme="light"] [data-testid="stPopover"] button {
  background: linear-gradient(180deg, #FFFFFF, #F8FAFC) !important;
  color: #1E293B !important;
}
html[data-sme-theme="light"] .sme-viz-card { background: #FFFFFF !important; }

/* ── THEME TOGGLE BUTTON STYLE (top of sidebar) ── */
.sme-theme-toggle {
  display:flex; align-items:center; justify-content:space-between;
  background: var(--bg3);
  border: 1px solid var(--border);
  border-radius: 99px;
  padding: 4px 6px;
  margin: .2rem 0 .8rem;
}
.sme-theme-toggle button {
  flex:1; background: transparent !important;
  color: var(--t4) !important; border: none !important;
  font-family:'JetBrains Mono',monospace !important;
  font-size: .58rem !important; padding: .3rem .5rem !important;
  letter-spacing:.1em; text-transform:uppercase;
  border-radius: 99px !important;
  box-shadow:none !important;
}
.sme-theme-toggle .active {
  background: var(--amber-bg) !important;
  color: var(--amber) !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# ADMIN LOGIN GATE  —  change credentials here
# ─────────────────────────────────────────────────────────────────────────────
_ADMIN_USER = "admin"
_ADMIN_PASS = "admin2026"

def _show_login():
    # Design integration: gradient shell + amber-accent card around the form.
    st.markdown('<div class="sme-login-shell"><div class="sme-login-card">',
                unsafe_allow_html=True)
    _, _logo_col, _ = st.columns([1, 1, 1])
    with _logo_col:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=200)
    st.markdown("""
    <div style="text-align:center;margin:.6rem 0 1.2rem;">
      <div style="font-size:38px;line-height:1;">🏗</div>
      <div class="sme-login-title">Smart Material Estimator</div>
      <div class="sme-login-sub">Enterprise Platform · v3</div>
    </div>""", unsafe_allow_html=True)
    user = st.text_input("Username", key="_login_user", placeholder="Enter username")
    pwd  = st.text_input("Password", type="password", key="_login_pass",
                         placeholder="Enter password")
    if st.button("🔐  Login", use_container_width=True, key="_login_btn"):
        if user == _ADMIN_USER and pwd == _ADMIN_PASS:
            st.session_state["_authenticated"] = True
            st.rerun()
        else:
            st.error("❌ Invalid credentials. Please try again.")
    st.markdown(
        '<div style="text-align:center;margin-top:14px;font-size:11px;'
        'color:var(--t5);">Demo: admin / admin2026</div>',
        unsafe_allow_html=True)
    st.markdown('</div></div>', unsafe_allow_html=True)

if "_authenticated" not in st.session_state:
    st.session_state["_authenticated"] = False
if not st.session_state["_authenticated"]:
    _show_login()
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# DATABASE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def db_available():
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


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADER  (SQLite when available, fallback to Excel)
# ─────────────────────────────────────────────────────────────────────────────
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
            .groupby(["Equipment_Tag_No.", "Lining_System_Code",
                      "Lining_System_Short_Name"], as_index=False)
            ["Surface_Area_SQM"].sum()
            .rename(columns={"Surface_Area_SQM": "Total_SQM_Original"}))
        equip_sc = equip_sc.merge(
            sqm_prog[["Equipment_Tag_No.", "Lining_System_Code",
                      "remaining_sqm", "done_sqm"]],
            on=["Equipment_Tag_No.", "Lining_System_Code"], how="left")
        equip_sc["remaining_sqm"] = equip_sc["remaining_sqm"].fillna(equip_sc["Total_SQM_Original"])
        equip_sc["done_sqm"]      = equip_sc["done_sqm"].fillna(0)
        equip_sc["Total_SQM"]     = equip_sc["remaining_sqm"]

    else:
        # Fallback: read from Excel using existing clean helpers
        from validate_data import clean_inventory, clean_recipe, clean_equipment
        df_a_raw = pd.read_excel(PATH_A, sheet_name=SHEET_A)
        df_b_raw = pd.read_excel(PATH_B, sheet_name=SHEET_B)
        df_c_raw = pd.read_excel(PATH_C, sheet_name=SHEET_C)
        inv    = clean_inventory(df_a_raw)
        recipe = clean_recipe(df_b_raw)
        equip_raw = clean_equipment(df_c_raw)

        inv_full = df_a_raw.copy()
        inv_full.columns = inv_full.columns.str.strip()
        inv_full["Material_Code"] = inv_full["Material_Code"].astype(str).str.strip()
        ordered_col = next((c for c in inv_full.columns
                            if c.strip() in ("Ordered_Qty","Balance To Be Received")), None)
        if ordered_col:
            inv_full[ordered_col] = pd.to_numeric(inv_full[ordered_col],errors="coerce").fillna(0)
            inv_ordered = inv_full.groupby("Material_Code",as_index=False).agg(
                Ordered_Qty=(ordered_col,"sum"))
            inv = inv.merge(inv_ordered,on="Material_Code",how="left")
            inv["Ordered_Qty"] = inv["Ordered_Qty"].fillna(0)
        else:
            inv["Ordered_Qty"] = 0.0

        raw = df_c_raw.copy()
        raw.columns = raw.columns.str.strip()
        raw = raw.dropna(subset=["Equipment_Tag_No.","Lining_System_Code"])
        raw["Equipment_Tag_No."] = raw["Equipment_Tag_No."].astype(str).str.strip()
        raw["Location"]           = raw["Location"].astype(str).str.strip()
        raw["Type"]               = raw["Type"].astype(str).str.strip()
        raw["Surface_Area_SQM"]   = pd.to_numeric(raw["Surface_Area_SQM"],errors="coerce")
        raw["Lining_System_Code"] = raw["Lining_System_Code"].astype(float).astype(int).astype(str)
        for col in ["Name","Substrate","Lining_System_Short_Name"]:
            if col in raw.columns:
                raw[col] = raw[col].astype(str).str.strip()
        equip_raw = raw

        equip_sc = (equip_raw
            .groupby(["Equipment_Tag_No.","Lining_System_Code","Lining_System_Short_Name"],
                     as_index=False)["Surface_Area_SQM"].sum()
            .rename(columns={"Surface_Area_SQM":"Total_SQM_Original"}))
        equip_sc["done_sqm"]      = 0.0
        equip_sc["remaining_sqm"] = equip_sc["Total_SQM_Original"]
        equip_sc["Total_SQM"]     = equip_sc["Total_SQM_Original"]

    # ── Demand matrix (uses remaining SQM) ───────────────────────────────
    dm = equip_sc.merge(recipe, on="Lining_System_Code", suffixes=("_e","_r"))
    dm["Demand_Qty"] = dm["For_1_SQM"] * dm["Total_SQM"]
    if "Lining_System_Short_Name_e" in dm.columns:
        dm = dm.rename(columns={"Lining_System_Short_Name_e":"Lining_System_Short_Name"})
        dm.drop(columns=["Lining_System_Short_Name_r"],inplace=True,errors="ignore")
    dm = dm[["Equipment_Tag_No.","Lining_System_Code","Lining_System_Short_Name",
             "Total_SQM","Material_Code","Material_Name","UOM","Demand_Qty"]]

    # ── Equipment master ─────────────────────────────────────────────────
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
    eq_master["Location"] = eq_master["Location"].str.strip()
    eq_master["Type"]     = eq_master["Type"].str.strip()

    # ── SQM reference (includes done_sqm for progress tracking) ─────────
    sqm_ref = equip_sc[["Equipment_Tag_No.","Lining_System_Code",
                         "Total_SQM","Total_SQM_Original","done_sqm"]].drop_duplicates()

    return inv, recipe, equip_sc, dm, eq_master, sqm_ref


inv, recipe, equip_sc, dm, eq_master, sqm_ref = load_all()
ALL_TAGS      = sorted(eq_master["Equipment_Tag_No."].tolist())
INV_POOL_INIT    = inv.set_index("Material_Code")["Available_Qty"].to_dict()
INV_ORDERED_INIT = inv.set_index("Material_Code")["Ordered_Qty"].to_dict()

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────
if "session_tags" not in st.session_state:
    st.session_state.session_tags = []

if "_session_key" not in st.session_state:
    import uuid as _uuid
    st.session_state["_session_key"] = str(_uuid.uuid4())

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

@st.cache_data(show_spinner=False)
def _cached_cascade_allocate(tag_order_tuple: tuple) -> pd.DataFrame:
    """Cached worker for cascade_allocate"""
    pool = dict(INV_POOL_INIT)  # mutable copy
    rows = []
    for tag in tag_order_tuple:
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

def cascade_allocate(tag_order: list[str]) -> pd.DataFrame:
    """
    Cascade inventory pool through equipment in order.
    Pool is GLOBAL per material (not per system code).
    """
    # Convert list to tuple so Streamlit's cache engine can hash it safely
    return _cached_cascade_allocate(tuple(tag_order))

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

def sqm_can_do(alloc_df: pd.DataFrame, tag: str, code: str) -> tuple[float, float, float]:
    """
    Returns (total_sqm, sqm_can_do, sqm_shortfall) for a (tag, system_code) pair.
    Method: weighted avg fulfillment % × Total_SQM
    """
    rows = alloc_df[(alloc_df["Equipment_Tag_No."]==tag) &
                    (alloc_df["Lining_System_Code"]==code)]
    if rows.empty:
        return 0.0, 0.0, 0.0
    total_sqm = sqm_ref[
        (sqm_ref["Equipment_Tag_No."]==tag) &
        (sqm_ref["Lining_System_Code"]==code)
    ]["Total_SQM"].sum()
    d = rows["Demand_Qty"].sum()
    a = rows["Allocated_Qty"].sum()
    pct = min(1.0, a / d) if d > 0 else 1.0
    can   = round(total_sqm * pct, 2)
    short = round(total_sqm - can, 2)
    return round(total_sqm, 2), can, short


def dbl_click_metric(
    label: str,
    value: str,
    state_key: str,
    drilldown_title: str,
    drilldown_df,
    help_text: str = "",
    delta: str = "",
    height: int = 95,
) -> None:
    """Renders a metric card popover — click opens the drill-down table."""
    btn_label = f"**{label}**\n\n{value}"
    if delta:
        btn_label += f"\n\n{delta}"
    with st.popover(btn_label, use_container_width=True):
        st.subheader(drilldown_title)
        if help_text:
            st.caption(help_text)
        _df = drilldown_df if drilldown_df is not None else pd.DataFrame()
        if len(_df):
            _MAX_ROWS = 200
            _total = len(_df)
            if _total > _MAX_ROWS:
                st.caption(f"Showing top {_MAX_ROWS:,} of {_total:,} rows")
                _df = _df.head(_MAX_ROWS)
            st.dataframe(_df, use_container_width=True, hide_index=True,
                         height=min(35 * (len(_df) + 1) + 3, 420))
        else:
            st.info("No detail data available for this metric.")


COLOR_SCHEMES = {
    "dashboard":   {"title_bg": "#1A2A3A", "header_bg": "#2D4A6A", "total_bg": "#F0C040", "total_fg": "#000000"},
    "brown_field": {"title_bg": "#0F2D52", "header_bg": "#1E5799", "total_bg": "#BDD7F0", "total_fg": "#000000"},
    "train_j":     {"title_bg": "#4A2E00", "header_bg": "#A0620A", "total_bg": "#FDE8A0", "total_fg": "#000000"},
    "train_k":     {"title_bg": "#0A2E1A", "header_bg": "#1A6B48", "total_bg": "#B3F0D8", "total_fg": "#000000"},
    "session":     {"title_bg": "#2D1A52", "header_bg": "#5B2D8E", "total_bg": "#E5D0F0", "total_fg": "#000000"},
    "execution":   {"title_bg": "#3A0A0A", "header_bg": "#8E1A1A", "total_bg": "#F5C6C6", "total_fg": "#000000"},
    "overview":    {"title_bg": "#0A2A2A", "header_bg": "#0E7490", "total_bg": "#A5F3FC", "total_fg": "#000000"},
}
_LOC_COLOR_MAP = {"Brown Field": "brown_field", "TRAIN J": "train_j", "TRAIN K": "train_k"}
_TABLE_COLOR_MAP = {"equipment": "brown_field", "recipe": "train_j", "inventory": "train_k"}


def generate_excel_report(df: pd.DataFrame,
                          report_title: str = "",
                          add_grand_total: bool = True,
                          color_scheme: str = "dashboard") -> bytes:
    """
    Professional Excel export using xlsxwriter:
    - Rows 0-3: space reserved for logo image (inserted at A1)
    - Row 4:    report title bar (merged, dark navy, white bold)
    - Row 5:    column headers  (navy bg, white bold, border)
    - Row 6+:   data rows       (border)
    - Last row: GRAND TOTAL     (gold bg, bold) when add_grand_total=True
    """
    buf = io.BytesIO()
    out_df = df.copy()

    # Drop helper columns that should not appear in exports
    for _drop in ("☐ Select", "Sl. No."):
        if _drop in out_df.columns:
            out_df = out_df.drop(columns=[_drop])

    # ── Grand total row ───────────────────────────────────────────────────
    if add_grand_total and len(out_df) > 0:
        num_cols = out_df.select_dtypes(include="number").columns.tolist()
        total_row = {c: "" for c in out_df.columns}
        total_row[out_df.columns[0]] = "GRAND TOTAL"
        for c in num_cols:
            try:
                total_row[c] = out_df[c].sum()
            except Exception:
                pass
        out_df = pd.concat([out_df, pd.DataFrame([total_row])], ignore_index=True)

    TITLE_ROW  = 4   # 0-indexed row for the title bar
    HEADER_ROW = 5   # 0-indexed row for column headers
    DATA_START = 6   # 0-indexed first data row
    n_cols = len(out_df.columns)
    cs = COLOR_SCHEMES.get(color_scheme, COLOR_SCHEMES["dashboard"])

    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        # Write data at HEADER_ROW so xlsxwriter positions cells correctly;
        # we will overwrite header row and data rows with formatted versions.
        out_df.to_excel(writer, index=False, sheet_name="Report",
                        startrow=HEADER_ROW)

        wb = writer.book
        ws = writer.sheets["Report"]

        # ── Formats ───────────────────────────────────────────────────────
        title_fmt = wb.add_format({
            "bold": True, "font_size": 13, "font_color": "#FFFFFF",
            "bg_color": cs["title_bg"], "align": "center", "valign": "vcenter",
            "border": 0,
        })
        header_fmt = wb.add_format({
            "bold": True, "font_size": 10, "font_color": "#FFFFFF",
            "bg_color": cs["header_bg"], "align": "center", "valign": "vcenter",
            "border": 1,
        })
        data_fmt = wb.add_format({
            "font_size": 9, "border": 1, "valign": "vcenter",
        })
        total_fmt = wb.add_format({
            "bold": True, "font_size": 10, "bg_color": cs["total_bg"],
            "font_color": cs["total_fg"], "border": 1, "valign": "vcenter",
        })

        # ── Logo — pre-resized to exactly 121×83 px @ 96 DPI = 1.26"×0.86" ─
        if os.path.exists(LOGO_PATH):
            _logo_buf = io.BytesIO()
            with _PILImage.open(LOGO_PATH) as _img:
                _img = _img.resize((121, 83), _PILImage.Resampling.LANCZOS)
                _img.save(_logo_buf, format="PNG", dpi=(96, 96))
            _logo_buf.seek(0)
            ws.insert_image(0, 0, "logo.png", {
                "image_data":      _logo_buf,
                "x_offset":        4,
                "y_offset":        4,
                "object_position": 1,
            })
        # 4 rows × 16 pts = 64 pts ≈ 0.889" — just enough to contain the 83 px logo
        for _r in range(4):
            ws.set_row(_r, 16)

        # ── Report metadata — right side of the header area ───────────────
        if n_cols >= 2:
            meta_label_fmt = wb.add_format({
                "font_size": 8, "bold": True, "align": "right",
                "valign": "vcenter", "font_color": "#555555",
            })
            meta_value_fmt = wb.add_format({
                "font_size": 8, "align": "left",
                "valign": "vcenter", "font_color": "#333333",
            })
            _gen_time = datetime.now().strftime("%Y-%m-%d  %H:%M")
            ws.write(1, n_cols - 2, "Report Generated:", meta_label_fmt)
            ws.write(1, n_cols - 1, _gen_time,           meta_value_fmt)
            ws.write(2, n_cols - 2, "Generated By:",     meta_label_fmt)
            ws.write(2, n_cols - 1, "Smart Material Estimator", meta_value_fmt)

        # ── Title row ─────────────────────────────────────────────────────
        if report_title and n_cols > 1:
            ws.merge_range(TITLE_ROW, 0, TITLE_ROW, n_cols - 1,
                           report_title, title_fmt)
        elif report_title:
            ws.write(TITLE_ROW, 0, report_title, title_fmt)
        ws.set_row(TITLE_ROW, 22)

        # ── Re-write header row with formatting ───────────────────────────
        for col_i, col_name in enumerate(out_df.columns):
            ws.write(HEADER_ROW, col_i, col_name, header_fmt)
        ws.set_row(HEADER_ROW, 18)
        ws.autofilter(HEADER_ROW, 0, HEADER_ROW, n_cols - 1)

        # ── Re-write data rows with formatting ────────────────────────────
        is_grand_total = add_grand_total and len(out_df) > 0
        for row_i, row_vals in enumerate(out_df.itertuples(index=False, name=None)):
            fmt = total_fmt if (is_grand_total and row_i == len(out_df) - 1) else data_fmt
            for col_i, val in enumerate(row_vals):
                cell_val = "" if (val is None or (isinstance(val, float) and np.isnan(val))) else val
                ws.write(DATA_START + row_i, col_i, cell_val, fmt)

        # ── Auto-width columns ────────────────────────────────────────────
        for col_i, col_name in enumerate(out_df.columns):
            col_data = out_df.iloc[:, col_i].fillna("").astype(str)
            max_len  = max(len(str(col_name)),
                           col_data.str.len().max() if len(col_data) else 0)
            ws.set_column(col_i, col_i, min(int(max_len) + 3, 42))

    return buf.getvalue()


def generate_multi_sheet_excel(sheets: list) -> bytes:
    """
    Build one workbook with one sheet per entry in `sheets`.
    Each sheet gets its own color scheme and AutoFilter.
    sheets: list of dicts with keys:
        name (str)             — Excel tab name (max 31 chars)
        df   (DataFrame)       — data
        title (str)            — title bar text
        color_scheme (str)     — key into COLOR_SCHEMES
        add_grand_total (bool) — default True
    """
    buf = io.BytesIO()
    TITLE_ROW, HEADER_ROW, DATA_START = 4, 5, 6

    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        for spec in sheets:
            out_df = spec["df"].copy()
            title  = spec.get("title", spec["name"])
            cs     = COLOR_SCHEMES.get(spec.get("color_scheme", "dashboard"), COLOR_SCHEMES["dashboard"])
            add_gt = spec.get("add_grand_total", True)
            sname  = spec["name"][:31]

            for _drop in ("☐ Select", "Sl. No."):
                if _drop in out_df.columns:
                    out_df = out_df.drop(columns=[_drop])

            if add_gt and len(out_df) > 0:
                num_cols = out_df.select_dtypes(include="number").columns.tolist()
                total_row = {c: "" for c in out_df.columns}
                total_row[out_df.columns[0]] = "GRAND TOTAL"
                for c in num_cols:
                    try: total_row[c] = out_df[c].sum()
                    except: pass
                out_df = pd.concat([out_df, pd.DataFrame([total_row])], ignore_index=True)

            n_cols = len(out_df.columns)
            out_df.to_excel(writer, index=False, sheet_name=sname, startrow=HEADER_ROW)
            ws = writer.sheets[sname]

            title_fmt  = wb.add_format({"bold": True, "font_size": 13, "font_color": "#FFFFFF",
                "bg_color": cs["title_bg"], "align": "center", "valign": "vcenter", "border": 0})
            header_fmt = wb.add_format({"bold": True, "font_size": 10, "font_color": "#FFFFFF",
                "bg_color": cs["header_bg"], "align": "center", "valign": "vcenter", "border": 1})
            data_fmt   = wb.add_format({"font_size": 9, "border": 1, "valign": "vcenter"})
            total_fmt  = wb.add_format({"bold": True, "font_size": 10, "border": 1,
                "bg_color": cs["total_bg"], "font_color": cs["total_fg"], "valign": "vcenter"})
            meta_label_fmt = wb.add_format({"font_size": 8, "bold": True, "align": "right",
                "valign": "vcenter", "font_color": "#555555"})
            meta_value_fmt = wb.add_format({"font_size": 8, "align": "left",
                "valign": "vcenter", "font_color": "#333333"})

            if os.path.exists(LOGO_PATH):
                _logo_buf = io.BytesIO()
                with _PILImage.open(LOGO_PATH) as _img:
                    _img = _img.resize((121, 83), _PILImage.Resampling.LANCZOS)
                    _img.save(_logo_buf, format="PNG", dpi=(96, 96))
                _logo_buf.seek(0)
                ws.insert_image(0, 0, "logo.png", {"image_data": _logo_buf,
                    "x_offset": 4, "y_offset": 4, "object_position": 1})
            for _r in range(4):
                ws.set_row(_r, 16)

            if n_cols >= 2:
                _gen_time = datetime.now().strftime("%Y-%m-%d  %H:%M")
                ws.write(1, n_cols - 2, "Report Generated:", meta_label_fmt)
                ws.write(1, n_cols - 1, _gen_time, meta_value_fmt)
                ws.write(2, n_cols - 2, "Generated By:", meta_label_fmt)
                ws.write(2, n_cols - 1, "Smart Material Estimator", meta_value_fmt)

            if title and n_cols > 1:
                ws.merge_range(TITLE_ROW, 0, TITLE_ROW, n_cols - 1, title, title_fmt)
            elif title:
                ws.write(TITLE_ROW, 0, title, title_fmt)
            ws.set_row(TITLE_ROW, 22)

            for col_i, col_name in enumerate(out_df.columns):
                ws.write(HEADER_ROW, col_i, col_name, header_fmt)
            ws.set_row(HEADER_ROW, 18)
            ws.autofilter(HEADER_ROW, 0, HEADER_ROW, n_cols - 1)

            is_gt = add_gt and len(out_df) > 0
            for row_i, row_vals in enumerate(out_df.itertuples(index=False, name=None)):
                fmt = total_fmt if (is_gt and row_i == len(out_df) - 1) else data_fmt
                for col_i, val in enumerate(row_vals):
                    cell_val = "" if (val is None or (isinstance(val, float) and np.isnan(val))) else val
                    ws.write(DATA_START + row_i, col_i, cell_val, fmt)

            for col_i, col_name in enumerate(out_df.columns):
                col_data = out_df.iloc[:, col_i].fillna("").astype(str)
                max_len  = max(len(str(col_name)), col_data.str.len().max() if len(col_data) else 0)
                ws.set_column(col_i, col_i, min(int(max_len) + 3, 42))

    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# SUGGESTION ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def _run_suggestion_engine(tag_list: list[str]) -> dict:
    """
    For each tag, tries moving it to every earlier position.
    Returns the best single-move gain per equipment and per system code.
    """
    if len(tag_list) < 2:
        return {"eq_suggestions": [], "sc_suggestions": [], "baseline_pcts": {}}

    baseline = cascade_allocate(tag_list)

    def _eq_pct(alloc, tag):
        t = alloc[alloc["Equipment_Tag_No."] == tag]
        d = t["Demand_Qty"].sum(); a = t["Allocated_Qty"].sum()
        return round(min(100.0, a/d*100), 2) if d > 0 else 100.0

    def _sc_pct(alloc, tag, code):
        t = alloc[(alloc["Equipment_Tag_No."]==tag)&(alloc["Lining_System_Code"]==code)]
        d = t["Demand_Qty"].sum(); a = t["Allocated_Qty"].sum()
        return round(min(100.0, a/d*100), 2) if d > 0 else 100.0

    baseline_pcts = {t: _eq_pct(baseline, t) for t in tag_list}
    eq_suggestions, sc_suggestions = [], []

    for i, target in enumerate(tag_list):
        if i == 0: continue
        base_pct = baseline_pcts[target]
        best_gain = 0; best_pos = i; best_new_pct = base_pct
        for new_pos in range(0, i):
            new_order = [t for t in tag_list if t != target]
            new_order.insert(new_pos, target)
            alloc   = cascade_allocate(new_order)
            new_pct = _eq_pct(alloc, target)
            gain    = new_pct - base_pct
            if gain > best_gain:
                best_gain = gain; best_pos = new_pos; best_new_pct = new_pct
        if best_gain > 0.4:
            eq_suggestions.append({
                "tag": target, "current_pos": i+1, "suggest_pos": best_pos+1,
                "current_pct": base_pct, "new_pct": best_new_pct,
                "gain": round(best_gain, 1),
            })

    for i, tag in enumerate(tag_list):
        if i == 0: continue
        tag_dm = dm[dm["Equipment_Tag_No."] == tag]
        for code in sorted(tag_dm["Lining_System_Code"].unique(), key=lambda x: int(x)):
            sname    = tag_dm[tag_dm["Lining_System_Code"]==code]["Lining_System_Short_Name"].iloc[0]
            base_pct = _sc_pct(baseline, tag, code)
            best_gain = 0; best_pos = i; best_new_pct = base_pct
            for new_pos in range(0, i):
                new_order = [t for t in tag_list if t != tag]
                new_order.insert(new_pos, tag)
                alloc   = cascade_allocate(new_order)
                new_pct = _sc_pct(alloc, tag, code)
                gain    = new_pct - base_pct
                if gain > best_gain:
                    best_gain = gain; best_pos = new_pos; best_new_pct = new_pct
            if best_gain > 0.4:
                sc_suggestions.append({
                    "tag": tag, "code": code, "sname": sname,
                    "current_pos": i+1, "suggest_pos": best_pos+1,
                    "current_pct": base_pct, "new_pct": best_new_pct,
                    "gain": round(best_gain, 1), "is_full": best_new_pct >= 99.9,
                })

    eq_suggestions.sort(key=lambda x: x["gain"], reverse=True)
    sc_suggestions.sort(key=lambda x: (x["is_full"], x["gain"]), reverse=True)
    return {
        "eq_suggestions": eq_suggestions[:5],
        "sc_suggestions": sc_suggestions[:8],
        "baseline_pcts":  baseline_pcts,
    }


def render_suggestion_panel(tag_list: list[str], panel_key: str) -> None:
    """Renders the Smart Reordering Suggestion panel."""
    st.markdown(
        '<div class="sec-hdr" style="margin-top:1.4rem;">'
        '💡 Smart Reordering Suggestions</div>', unsafe_allow_html=True)
    st.caption("Each suggestion shows the single best position change for one equipment "
               "or system code. Moving it there frees up inventory for that item first.")

    with st.spinner("Analysing reorder scenarios…"):
        result = _run_suggestion_engine(tag_list)

    eq_sugg = result["eq_suggestions"]
    sc_sugg = result["sc_suggestions"]

    if not eq_sugg and not sc_sugg:
        return  # hide silently when no improvements found

    c_eq, c_sc = st.columns(2, gap="large")

    with c_eq:
        st.markdown(
            '<div style="font-family:\'JetBrains Mono\',monospace;font-size:.65rem;'
            'font-weight:700;letter-spacing:.1em;text-transform:uppercase;'
            'color:var(--amber);margin-bottom:.6rem;"> By Equipment</div>',
            unsafe_allow_html=True)
        if not eq_sugg:
            st.caption("No equipment-level gains found.")
        for s in eq_sugg:
            gc = "#10B981" if s["new_pct"] >= 99.9 else "#F59E0B"
            tn = eq_master.set_index("Equipment_Tag_No.")["Name"].get(s["tag"], s["tag"])
            tag_label = f"{s['tag']}  ·  {tn[:24]}"
            move_label = (
                f"Move #{s['current_pos']} → #{s['suggest_pos']}  ·  "
                f"{s['current_pct']:.1f}% → {s['new_pct']:.1f}%"
                + ("  ✅ Full completion!" if s['new_pct'] >= 99.9 else "")
            )
            st.markdown(
                f'<div class="card" style="margin-bottom:.45rem;padding:.75rem 1rem;">'
                f'<div style="display:flex;justify-content:space-between;align-items:center;">'
                f'<span style="font-size:.82rem;font-weight:700;color:var(--t0);">{tag_label}</span>'
                f'<span style="font-family:JetBrains Mono,monospace;font-size:.75rem;'
                f'color:{gc};font-weight:700;">+{s["gain"]:.1f}%</span></div>'
                f'<div style="font-size:.76rem;color:var(--t2);margin-top:.35rem;">'
                f'{move_label}</div></div>',
                unsafe_allow_html=True)

    with c_sc:
        st.markdown(
            '<div style="font-family:\'JetBrains Mono\',monospace;font-size:.65rem;'
            'font-weight:700;letter-spacing:.1em;text-transform:uppercase;'
            'color:var(--amber);margin-bottom:.6rem;"> By System Code</div>',
            unsafe_allow_html=True)
        if not sc_sugg:
            st.caption("No system-code-level gains found.")
        for s in sc_sugg:
            gc   = "#10B981" if s["is_full"] else "#F59E0B"
            full = ('  <span style="background:rgba(16,185,129,.15);color:#10B981;'
                    'font-size:.65rem;padding:.1rem .4rem;border-radius:3px;">100% COMPLETE</span>'
                    if s["is_full"] else "")
            tn = eq_master.set_index("Equipment_Tag_No.")["Name"].get(s["tag"], s["tag"])
            sc_label   = f"Code {s['code']}  {s['sname']}  ·  {s['tag']} {tn[:20]}"
            sc_move    = (
                f"Move #{s['current_pos']} → #{s['suggest_pos']}  ·  "
                f"{s['current_pct']:.1f}% → {s['new_pct']:.1f}%"
                + ("  ✅ 100% COMPLETE" if s["is_full"] else "")
            )
            st.markdown(
                f'<div class="card" style="margin-bottom:.45rem;padding:.75rem 1rem;">'
                f'<div style="display:flex;justify-content:space-between;align-items:center;">'
                f'<span style="font-size:.8rem;font-weight:600;color:var(--t0);">{sc_label}</span>'
                f'<span style="font-family:JetBrains Mono,monospace;font-size:.75rem;'
                f'color:{gc};font-weight:700;">+{s["gain"]:.1f}%</span></div>'
                f'<div style="font-size:.76rem;color:var(--t2);margin-top:.3rem;">'
                f'{sc_move}</div></div>',
                unsafe_allow_html=True)


# ── Type/Description label helper ─────────────────────────────────────────────
def _eq_label(tag: str) -> str:
    """Returns 'TAG  —  Name  |  Type  |  Desc' for display labels."""
    row = eq_master[eq_master["Equipment_Tag_No."] == tag]
    if row.empty: return tag
    r = row.iloc[0]
    type_s = str(r.get("Type","") or "").strip()
    desc_s = str(r.get("Substrate","") or "").strip()
    name_s = str(r.get("Name","") or "").strip()
    parts  = [name_s]
    if type_s and type_s not in ("nan","—"): parts.append(type_s)
    if desc_s and desc_s not in ("nan","—"): parts.append(desc_s[:28])
    return f"{tag}  —  " + "  |  ".join(parts)


def plotly_mat_table(df: pd.DataFrame, key_suffix: str, height: int = 380,
                     show_sqm: bool = False, tag: str = "", code: str = "",
                     allocated_label: str = "Allocated") -> None:
    """Colour-coded material table. If show_sqm=True, adds SQM columns after qty cols."""
    base_cols = ["Material_Code", "Material_Name", "UOM",
                 "Demand_Qty", "Allocated_Qty", "Shortfall_Qty", "Fulfillment_Pct"]
    avail_cols = [c for c in base_cols if c in df.columns]
    df2 = df[avail_cols].copy()

    # Add Ordered_Qty if available
    if "Ordered_Qty" in df.columns:
        df2.insert(df2.columns.get_loc("Allocated_Qty"), "Ordered_Qty", df["Ordered_Qty"])

    rename_map = {
        "Material_Code":  "Code",
        "Material_Name":  "Material Name",
        "UOM":            "UOM",
        "Demand_Qty":     "Demand",
        "Ordered_Qty":    "On Order",
        "Allocated_Qty":  allocated_label,
        "Shortfall_Qty":  "Shortfall",
        "Fulfillment_Pct":"Fulfil %",
    }
    df2 = df2.rename(columns=rename_map)

    # SQM columns — per-material, based on each material's own fulfillment
    if show_sqm and tag and code:
        total_sqm_sc = sqm_ref[
            (sqm_ref["Equipment_Tag_No."]==tag) &
            (sqm_ref["Lining_System_Code"]==code)
        ]["Total_SQM"].sum()
        # Each material gets SQM proportional to its own fulfillment rate
        mat_fulfill = df["Fulfillment_Pct"].values / 100.0 if "Fulfillment_Pct" in df.columns else                       (df["Allocated_Qty"] / df["Demand_Qty"].replace(0, np.nan)).fillna(1.0).clip(0,1).values
        df2["SQM Total"]   = round(total_sqm_sc, 2)
        df2["SQM Done"]    = (total_sqm_sc * mat_fulfill).round(2)
        df2["SQM Deficit"] = (total_sqm_sc * (1 - mat_fulfill)).round(2)
        df2["SQM Done %"]  = (mat_fulfill * 100).round(1)

    # Pre-computed per-row SQM columns (for combined/aggregated tables)
    elif {"SQM_Total","SQM_Done","SQM_Deficit"}.issubset(df.columns):
        df2["SQM Total"]   = df["SQM_Total"].values
        df2["SQM Done"]    = df["SQM_Done"].values
        df2["SQM Deficit"] = df["SQM_Deficit"].values
        df2["SQM Done %"]  = np.where(
            df["SQM_Total"].values > 0,
            (df["SQM_Done"].values / df["SQM_Total"].replace(0, np.nan).values * 100),
            100.0
        ).round(1)

    fmt = {
        "Demand":          "{:,.3f}",
        allocated_label:   "{:,.3f}",
        "Shortfall":       "{:,.3f}",
        "Fulfil %":        "{:.1f}%",
    }
    if "On Order" in df2.columns:
        fmt["On Order"] = "{:,.3f}"
    if "SQM Total" in df2.columns:
        fmt.update({"SQM Total":"{:,.2f}","SQM Done":"{:,.2f}",
                    "SQM Deficit":"{:,.2f}","SQM Done %":"{:.1f}%"})

    fulfil_col = "Fulfil %"

    def style_row(row):
        pct = row.get(fulfil_col, 100)
        if pd.isna(pct): pct = 100.0
        if pct >= 100:
            bg, tc = "rgba(16,185,129,0.12)", "#10B981"
        elif pct >= 90:
            bg, tc = "rgba(249,115,22,0.12)",  "#F97316"
        elif pct >= 80:
            bg, tc = "rgba(234,179,8,0.12)",   "#EAB308"
        else:
            bg, tc = "rgba(239,68,68,0.12)",   "#EF4444"
        styles = [f"background-color:{bg}"] * len(row)
        fidx = list(row.index).index(fulfil_col) if fulfil_col in row.index else -1
        if fidx >= 0:
            styles[fidx] = f"background-color:{bg};color:{tc};font-weight:700"
        return styles

    styled_df = df2.style.apply(style_row, axis=1).format(fmt)
    st.dataframe(styled_df, hide_index=True, use_container_width=True,
                 height=height, key=f"tbl_{key_suffix}")



# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
# DESIGN INTEGRATION — theme toggle + inline SVG helpers (gauge + hbar)
# ─────────────────────────────────────────────────────────────────────────────
if "sme_theme" not in st.session_state:
    st.session_state.sme_theme = "dark"

def _apply_theme_attr():
    """Inject the data-sme-theme attribute on <html> so the light/dark CSS
    overrides take effect. Calc/data flow is untouched."""
    mode = st.session_state.get("sme_theme", "dark")
    st.markdown(
        f"""<script>
        (function(){{
          try {{
            const r = (window.parent || window).document.documentElement;
            r.setAttribute('data-sme-theme', '{mode}');
          }} catch(e) {{}}
        }})();
        </script>""",
        unsafe_allow_html=True,
    )

def _fc(p: float) -> str:
    """Coverage colour — matches design.fc()."""
    if p >= 100: return "#10B981"
    if p >= 90:  return "#F97316"
    if p >= 80:  return "#EAB308"
    return "#EF4444"

def render_design_gauge(pct: float, can_sqm: float, total_sqm: float) -> str:
    """Half-gauge SVG mirroring the Claude design's renderGauge."""
    import math
    pct = max(0.0, min(100.0, float(pct or 0)))
    w, h, cx, cy, R = 300, 168, 150, 158, 115
    sA = -math.pi
    vA = sA + (pct / 100.0) * math.pi
    def arc(r, s, e):
        x1, y1 = cx + r * math.cos(s), cy + r * math.sin(s)
        x2, y2 = cx + r * math.cos(e), cy + r * math.sin(e)
        large = 1 if (e - s) > math.pi else 0
        return f"M {x1:.1f} {y1:.1f} A {r} {r} 0 {large} 1 {x2:.1f} {y2:.1f}"
    col = _fc(pct)
    val_arc = (
        f'<path d="{arc(R, sA, vA)}" fill="none" stroke="{col}" '
        f'stroke-width="20" stroke-linecap="round"/>'
        if pct > 0 else ""
    )
    return f"""
    <div class="sme-viz-card">
      <svg viewBox="0 0 {w} {h}" preserveAspectRatio="xMidYMid meet">
        <path d="{arc(R, sA, sA + .5*math.pi)}" fill="none"
              stroke="rgba(239,68,68,.18)" stroke-width="22"/>
        <path d="{arc(R, sA + .5*math.pi, sA + .7*math.pi)}" fill="none"
              stroke="rgba(234,179,8,.18)" stroke-width="22"/>
        <path d="{arc(R, sA + .7*math.pi, sA + .85*math.pi)}" fill="none"
              stroke="rgba(249,115,22,.18)" stroke-width="22"/>
        <path d="{arc(R, sA + .85*math.pi, 0)}" fill="none"
              stroke="rgba(16,185,129,.18)" stroke-width="22"/>
        <path d="{arc(R, sA, 0)}" fill="none" stroke="rgba(128,128,128,.18)"
              stroke-width="20"/>
        {val_arc}
        <text x="{cx-R+2}" y="{cy+20}" fill="#94A3B8" font-size="10"
              font-family="JetBrains Mono, monospace">0%</text>
        <text x="{cx+R-22}" y="{cy+20}" fill="#94A3B8" font-size="10"
              font-family="JetBrains Mono, monospace">100%</text>
        <text x="{cx}" y="{cy-20}" text-anchor="middle" fill="{col}"
              font-size="32" font-weight="800"
              font-family="JetBrains Mono, monospace">{pct:.1f}%</text>
        <text x="{cx}" y="{cy-2}" text-anchor="middle" fill="#94A3B8"
              font-size="11" font-family="Inter, sans-serif">Overall Coverage</text>
        <text x="{cx}" y="{cy+14}" text-anchor="middle" fill="#94A3B8"
              font-size="10" font-family="JetBrains Mono, monospace">
          {can_sqm:,.1f} / {total_sqm:,.1f} SQM
        </text>
      </svg>
      <div class="sme-viz-legend">
        <span>■ Available: {can_sqm:,.1f} SQM</span>
        <span>■ Shortfall: {max(0.0, total_sqm-can_sqm):,.1f} SQM</span>
      </div>
    </div>
    """

def render_design_hbar(data: list[dict], title: str = "") -> str:
    """Horizontal bar chart SVG mirroring the Claude design's renderHBar.
    data: list of {"label": str, "val": float (0–100)} dicts."""
    if not data:
        return f'<div class="sme-viz-card">{title}<div style="color:var(--t4);font-size:.75rem;">No data.</div></div>'
    w, bH, gap, padL, padR = 460, 24, 7, 140, 60
    iW = w - padL - padR
    maxV = max([float(d.get("val") or 0) for d in data] + [1.0])
    total_h = len(data) * (bH + gap) + 16
    rows = []
    for i, d in enumerate(data):
        v = float(d.get("val") or 0)
        lbl = str(d.get("label", ""))[:20]
        y = i * (bH + gap) + 6
        bW = max(2.0, (v / maxV) * iW)
        col = _fc(v)
        rows.append(
            f'<text x="{padL-7}" y="{y+bH/2+4}" text-anchor="end" '
            f'fill="#94A3B8" font-size="11" font-family="Inter">{lbl}</text>'
            f'<rect x="{padL}" y="{y}" width="{iW}" height="{bH}" rx="4" '
            f'fill="rgba(128,128,128,.12)"/>'
            f'<rect x="{padL}" y="{y}" width="{bW:.1f}" height="{bH}" rx="4" '
            f'fill="{col}" opacity=".85"/>'
            f'<text x="{padL+bW+5:.1f}" y="{y+bH/2+4}" fill="#94A3B8" '
            f'font-size="11" font-family="JetBrains Mono, monospace">{v:.1f}%</text>'
        )
    title_html = (
        f'<div style="font-family:JetBrains Mono,monospace;font-size:.6rem;'
        f'font-weight:700;letter-spacing:.13em;text-transform:uppercase;'
        f'color:var(--t4);margin-bottom:.5rem;">{title}</div>'
        if title else ""
    )
    return f"""
    <div class="sme-viz-card">
      {title_html}
      <svg viewBox="0 0 {w} {total_h}" preserveAspectRatio="xMidYMid meet">
        {"".join(rows)}
      </svg>
    </div>
    """

# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    # Theme toggle — design integration (light/dark mode switch)
    _ttl, _ttr = st.columns(2, gap="small")
    with _ttl:
        if st.button("🌙 Dark",
                     key="_sme_theme_dark",
                     use_container_width=True,
                     type=("primary" if st.session_state.sme_theme == "dark" else "secondary")):
            st.session_state.sme_theme = "dark"
            st.rerun()
    with _ttr:
        if st.button("☀ Light",
                     key="_sme_theme_light",
                     use_container_width=True,
                     type=("primary" if st.session_state.sme_theme == "light" else "secondary")):
            st.session_state.sme_theme = "light"
            st.rerun()
    _apply_theme_attr()

    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=140)
    st.markdown("""
    <div style="padding:.3rem 0 1.2rem">
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
_hdr_logo = f'<img src="data:image/png;base64,{_logo_b64()}" style="height:34px;border-radius:6px;flex-shrink:0;">' if _logo_b64() else ""
st.markdown(f"""
<div class="sticky-header-wrap">
  <div style="display:flex;align-items:center;gap:1rem;">
    {_hdr_logo}
    <div style="display:flex;flex-direction:column;gap:.1rem;">
      <span style="font-family:'JetBrains Mono',monospace;font-size:1.05rem;font-weight:700;color:var(--t0);letter-spacing:-.01em;line-height:1.2;">
        Smart Material Estimator &amp; Planner</span>
      <span style="font-family:'JetBrains Mono',monospace;font-size:.58rem;color:var(--t5);letter-spacing:.08em;text-transform:uppercase;">
        System-code level · Cascading allocation · Priority-based</span>
    </div>
    <div style="margin-left:auto;display:flex;align-items:center;gap:.6rem;flex-shrink:0;">
      <span style="font-family:'JetBrains Mono',monospace;font-size:.58rem;color:var(--t5);letter-spacing:.1em;text-transform:uppercase;background:var(--amber-bg);border:1px solid rgba(245,158,11,.2);padding:.15rem .5rem;border-radius:99px;">v3</span>
      <span title="System online" style="width:7px;height:7px;border-radius:50%;background:var(--green);display:inline-block;box-shadow:0 0 7px var(--green-glow);"></span>
    </div>
  </div>
</div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM HAMBURGER (JS-injected — overrides Streamlit's invisible/missing toggle)
# ─────────────────────────────────────────────────────────────────────────────
import streamlit.components.v1 as _components
_components.html("""
<script>
(function(){
  const PARENT = window.parent.document;
  const BTN_ID = 'ge-custom-hamburger';

  function findToggle() {
    return PARENT.querySelector('[data-testid="stSidebarCollapsedControl"] button')
        || PARENT.querySelector('[data-testid="collapsedControl"] button')
        || PARENT.querySelector('[data-testid="stSidebarCollapseButton"] button')
        || PARENT.querySelector('[data-testid="stSidebarCollapseButton"]')
        || PARENT.querySelector('[data-testid="stSidebar"] button[kind="header"]')
        || PARENT.querySelector('[data-testid="stSidebar"] [data-testid="stSidebarHeader"] button');
  }

  function makeBtn() {
    if (PARENT.getElementById(BTN_ID)) return;
    const btn = PARENT.createElement('button');
    btn.id = BTN_ID;
    btn.type = 'button';
    btn.setAttribute('aria-label', 'Toggle sidebar');
    btn.title = 'Toggle sidebar';
    btn.textContent = '\\u2630';
    btn.style.cssText = [
      'position:fixed','top:14px','left:16px','z-index:1000001',
      'width:40px','height:40px',
      'background:rgba(245,158,11,0.12)',
      'border:1.5px solid #F59E0B',
      'border-radius:6px',
      'color:#F59E0B',
      'font-size:22px','font-weight:700','line-height:1',
      'font-family:Inter,Helvetica,sans-serif',
      'cursor:pointer',
      'display:flex','align-items:center','justify-content:center',
      'box-shadow:0 0 0 1px #F59E0B,0 2px 12px rgba(245,158,11,0.45)',
      'padding:0','margin:0',
      'transition:all .15s ease'
    ].join(';') + ';';
    btn.addEventListener('mouseenter', function(){
      btn.style.background = 'rgba(245,158,11,0.22)';
      btn.style.transform  = 'translateY(-1px)';
    });
    btn.addEventListener('mouseleave', function(){
      btn.style.background = 'rgba(245,158,11,0.12)';
      btn.style.transform  = 'translateY(0)';
    });
    btn.addEventListener('click', function(e){
      e.preventDefault();
      e.stopPropagation();
      const t = findToggle();
      if (t) { t.click(); }
    });
    PARENT.body.appendChild(btn);
  }

  // Try immediately + retry until the DOM is ready
  makeBtn();
  let tries = 0;
  const iv = setInterval(function(){
    makeBtn();
    if (++tries > 40) clearInterval(iv);
  }, 250);

  // Keep button alive across Streamlit re-renders
  const obs = new MutationObserver(function(){
    if (!PARENT.getElementById(BTN_ID)) makeBtn();
  });
  obs.observe(PARENT.body, { childList: true, subtree: false });
})();
</script>
""", height=0, width=0)

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tab0, tab1, tab2, tab3, tab4, tab_consume, tab5, tab_master = st.tabs([
    "📊  Dashboard",
    "🔍  Selective Equipment Entry",
    "📦  Session Order Report",
    "📍  Location Report",
    "⚙️  Execution Plan",
    "📦  Inventory",
    "📈  Total Overview",
    "🗄️  Master Data",
])



# ═══════════════════════════════════════════════════════════════════════════════
# TAB 0 · DASHBOARD (Project Overview + Material Requirement & Procurement)
# ═══════════════════════════════════════════════════════════════════════════════
with tab0:

    # ── Dashboard toggle ──────────────────────────────────────────────────────
    dash_view = st.radio(
        "View", ["📈 Project Overview", "🛒 Material Requirement & Procurement"],
        horizontal=True, key="dash_view", label_visibility="collapsed",
    )
    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Shared filter controls (used by both views) ───────────────────────────
    st.markdown('<div class="sec-hdr">🎛 Filter</div>', unsafe_allow_html=True)
    df1, df2_col, df3, df4 = st.columns(4)
    with df1:
        sel_locations = st.multiselect(" Location", options=LOCATION_ORDER,
                                        default=LOCATION_ORDER, key="dash_loc")
    with df2_col:
        all_types_d = sorted(eq_master["Type"].str.strip().unique().tolist())
        sel_types = st.multiselect(" Type", options=all_types_d,
                                    default=all_types_d, key="dash_type")
    with df3:
        all_codes_d = (
            dm[["Lining_System_Code","Lining_System_Short_Name"]].drop_duplicates()
            .sort_values("Lining_System_Code", key=lambda x: x.astype(int))
        )
        code_opts_d = [f"Code {r.Lining_System_Code} – {r.Lining_System_Short_Name}"
                       for _, r in all_codes_d.iterrows()]
        sel_codes_raw = st.multiselect(" System Code", options=code_opts_d,
                                        default=code_opts_d, key="dash_code")
        sel_codes = [c.split(" – ")[0].replace("Code ","").strip() for c in sel_codes_raw]
    with df4:
        all_desc_d = sorted(eq_master["Substrate"].dropna().unique().tolist())
        sel_substrate = st.multiselect(" Substrate", options=all_desc_d,
                                        default=all_desc_d, key="dash_substrate")

    # ── Apply filters ─────────────────────────────────────────────────────────
    filtered_eq = eq_master[
        eq_master["Location"].isin(sel_locations) &
        eq_master["Type"].str.strip().isin(sel_types) &
        eq_master["Substrate"].isin(sel_substrate)
    ]
    if sel_codes:
        tags_w_code = dm[dm["Lining_System_Code"].isin(sel_codes)]["Equipment_Tag_No."].unique()
        filtered_eq = filtered_eq[filtered_eq["Equipment_Tag_No."].isin(tags_w_code)]
    filtered_tags = filtered_eq["Equipment_Tag_No."].tolist()

    filtered_dm = dm[
        (dm["Equipment_Tag_No."].isin(filtered_tags)) &
        (dm["Lining_System_Code"].isin(sel_codes))
    ].copy()

    # ── Correct SQM using sqm_ref (never sum from dm — inflated x n_materials) ─
    proj_sqm = sqm_ref[
        sqm_ref["Equipment_Tag_No."].isin(filtered_tags) &
        sqm_ref["Lining_System_Code"].isin(sel_codes)
    ]["Total_SQM"].sum()

    # ── Material demand aggregation ───────────────────────────────────────────
    f_demand = (
        filtered_dm.groupby(["Material_Code","Material_Name","UOM"], as_index=False)
        ["Demand_Qty"].sum()
    )
    f_demand = f_demand.merge(
        inv[["Material_Code","Available_Qty","Ordered_Qty"]], on="Material_Code", how="left"
    )
    f_demand["Available_Qty"] = f_demand["Available_Qty"].fillna(0)
    f_demand["Ordered_Qty"]   = f_demand["Ordered_Qty"].fillna(0)
    f_demand["Shortfall"]     = (f_demand["Demand_Qty"] - f_demand["Available_Qty"]).clip(lower=0).round(3)
    f_demand["Net_Shortfall"] = (
        f_demand["Demand_Qty"] - f_demand["Available_Qty"] - f_demand["Ordered_Qty"]
    ).clip(lower=0).round(3)
    f_demand["Coverage_Pct"]  = (
        f_demand["Available_Qty"].clip(upper=f_demand["Demand_Qty"]) /
        f_demand["Demand_Qty"].replace(0, np.nan) * 100
    ).fillna(100).clip(0,100).round(1)

    f_total_demand = f_demand["Demand_Qty"].sum()
    f_total_avail  = f_demand["Available_Qty"].clip(upper=f_demand["Demand_Qty"]).sum()
    f_total_short  = f_demand["Shortfall"].sum()
    f_total_net    = f_demand["Net_Shortfall"].sum()
    f_cov          = (f_total_avail / f_total_demand * 100) if f_total_demand > 0 else 100
    can_sqm        = round(proj_sqm * min(1.0, f_cov/100), 2)
    short_sqm      = round(proj_sqm - can_sqm, 2)

    # ── SQM-based drill-down (per Equipment × System Code) ───────────────────
    _pair_d = filtered_dm.merge(
        inv[["Material_Code", "Available_Qty"]], on="Material_Code", how="left"
    )
    _pair_d["Available_Qty"] = _pair_d["Available_Qty"].fillna(0)
    _pair_d["Avail_Cap"]     = _pair_d[["Demand_Qty", "Available_Qty"]].min(axis=1)
    _pair_agg = _pair_d.groupby(
        ["Equipment_Tag_No.", "Lining_System_Code"], as_index=False
    ).agg(_Demand=("Demand_Qty", "sum"), _Avail=("Avail_Cap", "sum"))
    _pair_agg["Coverage %"] = (
        _pair_agg["_Avail"] / _pair_agg["_Demand"].replace(0, np.nan) * 100
    ).fillna(100).clip(0, 100).round(1)

    _dd_sqm_pair = sqm_ref[
        sqm_ref["Equipment_Tag_No."].isin(filtered_tags) &
        sqm_ref["Lining_System_Code"].isin(sel_codes)
    ][["Equipment_Tag_No.", "Lining_System_Code", "Total_SQM"]].merge(
        _pair_agg[["Equipment_Tag_No.", "Lining_System_Code", "Coverage %"]],
        on=["Equipment_Tag_No.", "Lining_System_Code"], how="left"
    )
    _dd_sqm_pair["Coverage %"]    = _dd_sqm_pair["Coverage %"].fillna(100)
    _dd_sqm_pair["Coverable SQM"] = (_dd_sqm_pair["Total_SQM"] * _dd_sqm_pair["Coverage %"] / 100).round(2)
    _dd_sqm_pair["SQM Deficit"]   = (_dd_sqm_pair["Total_SQM"] - _dd_sqm_pair["Coverable SQM"]).round(2)
    _dd_sqm_pair = _dd_sqm_pair.rename(columns={
        "Equipment_Tag_No.":  "Equipment Tag",
        "Lining_System_Code": "System Code",
        "Total_SQM":          "Total SQM",
    })

    _dd_cov_sqm_df = (_dd_sqm_pair[
        ["Equipment Tag", "System Code", "Total SQM", "Coverage %", "Coverable SQM"]
    ].sort_values("Coverage %").reset_index(drop=True))

    _dd_def_sqm_df = (_dd_sqm_pair[_dd_sqm_pair["SQM Deficit"] > 0][
        ["Equipment Tag", "System Code", "Total SQM", "Coverable SQM", "SQM Deficit"]
    ].sort_values("SQM Deficit", ascending=False).reset_index(drop=True))

    # ─────────────────────────────────────────────────────────────────────────
    if dash_view == "📈 Project Overview":
    # ─────────────────────────────────────────────────────────────────────────

        # KPI strip
        k1,k2,k3,k4,k5,k6,k7 = st.columns(7)
        _dd_equip_df = filtered_eq[["Equipment_Tag_No.","Name","Location","Type","Substrate"]].reset_index(drop=True)
        _dd_sqm_df   = (sqm_ref[sqm_ref["Equipment_Tag_No."].isin(filtered_tags) &
                                sqm_ref["Lining_System_Code"].isin(sel_codes)]
                        [["Equipment_Tag_No.","Lining_System_Code","Total_SQM"]]
                        .sort_values("Total_SQM", ascending=False).reset_index(drop=True))
        _dd_cov_df   = f_demand[["Material_Code","Material_Name","Demand_Qty","Available_Qty","Coverage_Pct"]].sort_values("Coverage_Pct").reset_index(drop=True)
        _dd_def_df   = f_demand[f_demand["Shortfall"]>0][["Material_Code","Material_Name","Demand_Qty","Available_Qty","Shortfall"]].sort_values("Shortfall", ascending=False).reset_index(drop=True)
        _dd_crit_df  = f_demand[f_demand["Coverage_Pct"]<50][["Material_Code","Material_Name","Demand_Qty","Available_Qty","Coverage_Pct"]].sort_values("Coverage_Pct").reset_index(drop=True)
        with k1:
            dbl_click_metric("Equipment", str(len(filtered_tags)), "t0_equip",
                "Equipment List", _dd_equip_df,
                help_text="Equipment tags matching current filter selection.")
        with k2:
            dbl_click_metric("Total SQM", f"{proj_sqm:,.1f}", "t0_sqm",
                "SQM by Equipment & System Code", _dd_sqm_df,
                help_text="Remaining surface area (m²) after deducting daily consumption entries.")
        with k3:
            dbl_click_metric("Available Coverage SQM", f"{can_sqm:,.2f}", "t0_cov_sqm",
                "Coverable SQM by Equipment & System Code", _dd_cov_sqm_df,
                help_text="Area (m²) coverable with currently available stock = Total SQM × Coverage %. Drill-down shows per-equipment SQM coverage.")
        with k4:
            dbl_click_metric("SQM Deficit", f"{short_sqm:,.2f}", "t0_deficit",
                "SQM Deficit by Equipment & System Code", _dd_def_sqm_df,
                help_text="Area (m²) that cannot be completed = Total SQM − Coverable SQM. Drill-down shows per-equipment SQM deficit.")
        with k5:
            dbl_click_metric("Overall Coverage", f"{f_cov:.1f}%", "t0_ov_cov",
                "Coverable SQM by Equipment & System Code", _dd_cov_sqm_df,
                delta=f"{f_cov-100:.1f}%",
                help_text="Allocated Qty ÷ Demand Qty × 100 across all filtered materials. Drill-down shows per-equipment SQM coverage.")
        with k6:
            dbl_click_metric("Shortfall SQM", f"{short_sqm:,.2f}", "t0_short_sqm",
                "SQM Deficit by Equipment & System Code", _dd_def_sqm_df,
                help_text="Area (m²) shortfall = Total SQM − Available Coverage SQM. Drill-down shows per-equipment SQM deficit.")
        with k7:
            dbl_click_metric("Critical (<50%)", str(int((f_demand["Coverage_Pct"]<50).sum())), "t0_critical",
                "Critical Materials (Coverage < 50%)", _dd_crit_df,
                help_text="Materials where Available Qty covers less than 50% of total demand.")
        st.markdown("<br>", unsafe_allow_html=True)

        row1a, row1b = st.columns([1,1.6], gap="large")

        with row1a:
            st.markdown('<div class="sec-hdr">🎯 Overall Coverage</div>', unsafe_allow_html=True)
            # ── Design integration: SVG gauge (mirrors Claude design) ──
            st.markdown(
                render_design_gauge(f_cov, can_sqm, proj_sqm),
                unsafe_allow_html=True,
            )
            fig_g = go.Figure(go.Indicator(
                mode="gauge+number+delta",
                value=round(f_cov,1),
                delta={"reference":100,"valueformat":".1f",
                       "decreasing":{"color":"#EF4444"},"increasing":{"color":"#10B981"}},
                number={"suffix":"%","font":{"family":"JetBrains Mono","size":36,"color":"var(--t0)"}},
                gauge={
                    "axis":{"range":[0,100],"tickwidth":1,
                            "tickfont":{"family":"JetBrains Mono","size":9}},
                    "bar":{"color":(
                        "#10B981" if f_cov>=100 else "#F97316" if f_cov>=90
                        else "#EAB308" if f_cov>=80 else "#EF4444"), "thickness":0.28},
                    "bgcolor":"rgba(0,0,0,0)","borderwidth":0,
                    "steps":[
                        {"range":[0,50],"color":"rgba(239,68,68,.08)"},
                        {"range":[50,80],"color":"rgba(234,179,8,.08)"},
                        {"range":[80,90],"color":"rgba(249,115,22,.08)"},
                        {"range":[90,100],"color":"rgba(16,185,129,.08)"},
                    ],
                },
                title={"text":f"Coverage  ·  {can_sqm:,.0f} / {proj_sqm:,.0f} SQM Available Material Coverage",
                       "font":{"family":"JetBrains Mono","size":9,"color":"rgba(148,163,184,.7)"}},
            ))
            fig_g.update_layout(paper_bgcolor="rgba(0,0,0,0)",
                                margin=dict(l=20,r=20,t=30,b=10),height=240)
            st.plotly_chart(fig_g, use_container_width=True, key="dash_gauge")

            # Demand vs Available mini stacked bar
            fig_dm = go.Figure()
            fig_dm.add_trace(go.Bar(name="Available",x=["Inventory"],y=[f_total_avail],
                marker_color="#10B981",marker_opacity=.8,
                text=[f"{f_total_avail:,.0f}"],textposition="auto",
                textfont=dict(family="JetBrains Mono",size=10)))
            fig_dm.add_trace(go.Bar(name="Shortfall",x=["Inventory"],y=[f_total_short],
                marker_color="#EF4444",marker_opacity=.8,
                text=[f"{f_total_short:,.0f}"],textposition="auto",
                textfont=dict(family="JetBrains Mono",size=10)))
            fig_dm.update_layout(barmode="stack",paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",margin=dict(l=0,r=0,t=0,b=0),height=120,
                showlegend=True,
                legend=dict(orientation="h",y=1.15,x=0,
                            font=dict(family="JetBrains Mono",size=9),bgcolor="rgba(0,0,0,0)"),
                xaxis=dict(showgrid=False,showticklabels=False),
                yaxis=dict(gridcolor="rgba(128,128,128,.1)",showticklabels=False))
            st.plotly_chart(fig_dm, use_container_width=True, key="dash_dmini")

        with row1b:
            st.markdown('<div class="sec-hdr">📍 Coverage by Location (SQM)</div>',
                        unsafe_allow_html=True)
            loc_rows = []
            for loc in sel_locations:
                loc_tags = filtered_eq[filtered_eq["Location"]==loc]["Equipment_Tag_No."].tolist()
                loc_dm_f = filtered_dm[filtered_dm["Equipment_Tag_No."].isin(loc_tags)]
                if loc_dm_f.empty: continue
                loc_agg = (loc_dm_f.groupby("Material_Code",as_index=False)["Demand_Qty"].sum()
                    .merge(inv[["Material_Code","Available_Qty"]],on="Material_Code",how="left"))
                loc_agg["Available_Qty"] = loc_agg["Available_Qty"].fillna(0)
                loc_d = loc_agg["Demand_Qty"].sum()
                loc_a = loc_agg["Available_Qty"].clip(upper=loc_agg["Demand_Qty"]).sum()
                loc_s = (loc_agg["Demand_Qty"]-loc_agg["Available_Qty"]).clip(lower=0).sum()
                loc_c = (loc_a/loc_d*100) if loc_d>0 else 100
                # ✅ Correct SQM using sqm_ref
                loc_sqm = sqm_ref[
                    sqm_ref["Equipment_Tag_No."].isin(loc_tags) &
                    sqm_ref["Lining_System_Code"].isin(sel_codes)
                ]["Total_SQM"].sum()
                loc_can = round(loc_sqm * min(1.0, loc_c/100), 2)
                loc_rows.append({"Location":loc,"Equipment":len(loc_tags),
                    "SQM":loc_sqm,"SQM_Can":loc_can,"SQM_Short":round(loc_sqm-loc_can,2),
                    "Demand":loc_d,"Available":loc_a,"Shortfall":loc_s,"Coverage_%":round(loc_c,1)})

            if loc_rows:
                loc_df = pd.DataFrame(loc_rows)
                loc_colors_map = {"Brown Field":"#3B82F6","TRAIN J":"#F59E0B","TRAIN K":"#10B981"}
                fig_loc = go.Figure()
                for _, lr in loc_df.iterrows():
                    c = loc_colors_map.get(lr["Location"],"#94A3B8")
                    fig_loc.add_trace(go.Bar(
                        x=[lr["Location"]],y=[lr["SQM_Can"]],name=f'{lr["Location"]} Can Do',
                        marker_color=c,marker_opacity=.8,
                        text=[f'{lr["Coverage_%"]:.0f}%\n{lr["SQM_Can"]:,.0f} SQM'],
                        textposition="inside",textfont=dict(family="JetBrains Mono",size=10,color="#fff"),
                        showlegend=False))
                    fig_loc.add_trace(go.Bar(
                        x=[lr["Location"]],y=[lr["SQM_Short"]],name=f'{lr["Location"]} Deficit',
                        marker_color="#EF4444",marker_opacity=.6,
                        text=[f'{lr["SQM_Short"]:,.0f} SQM deficit'],
                        textposition="inside",textfont=dict(family="JetBrains Mono",size=9,color="#fff"),
                        showlegend=False))
                fig_loc.update_layout(barmode="stack",paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",margin=dict(l=0,r=0,t=10,b=0),height=220,
                    xaxis=dict(tickfont=dict(family="JetBrains Mono",size=11)),
                    yaxis=dict(gridcolor="rgba(128,128,128,.08)",
                               title=dict(text="SQM",font=dict(family="JetBrains Mono",size=9))),
                    font=dict(family="JetBrains Mono",size=10,color="rgba(148,163,184,.8)"))
                st.plotly_chart(fig_loc, use_container_width=True, key="dash_loc_bar")

                # Location stat cards
                cols_loc = st.columns(len(loc_rows))
                for col, lr in zip(cols_loc, loc_rows):
                    dot = ("🟢" if lr["Coverage_%"]>=100 else "🟠" if lr["Coverage_%"]>=90
                           else "🟡" if lr["Coverage_%"]>=80 else "🔴")
                    loc_html = (
                        '<div class="card" style="text-align:center;padding:.7rem;">'
                        f'<div style="font-size:1.1rem;">{dot}</div>'
                        '<div style="font-family:\'JetBrains Mono\',monospace;font-size:.72rem;'
                        f'font-weight:700;color:var(--amber);margin:.2rem 0;">{lr["Location"]}</div>'
                        '<div style="font-family:\'JetBrains Mono\',monospace;font-size:1.1rem;'
                        f'font-weight:700;color:var(--t0);">{lr["Coverage_%"]:.1f}%</div>'
                        f'<div style="font-size:.68rem;color:var(--t3);">'
                        f'{lr["SQM_Can"]:,.0f} / {lr["SQM"]:,.0f} SQM</div>'
                        f'<div style="font-size:.65rem;color:var(--t3);">{lr["Equipment"]} equipment</div>'
                        '</div>'
                    )
                    col.markdown(loc_html, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        row2a, row2b = st.columns(2, gap="large")

        with row2a:
            st.markdown('<div class="sec-hdr"> Coverage by System Code (SQM)</div>',
                        unsafe_allow_html=True)
            sc_rows = []
            for code in sel_codes:
                sc_dm = filtered_dm[filtered_dm["Lining_System_Code"]==code]
                if sc_dm.empty: continue
                sname = sc_dm["Lining_System_Short_Name"].iloc[0]
                sc_agg = (sc_dm.groupby("Material_Code",as_index=False)["Demand_Qty"].sum()
                    .merge(inv[["Material_Code","Available_Qty"]],on="Material_Code",how="left"))
                sc_agg["Available_Qty"] = sc_agg["Available_Qty"].fillna(0)
                sc_d = sc_agg["Demand_Qty"].sum()
                sc_a = sc_agg["Available_Qty"].clip(upper=sc_agg["Demand_Qty"]).sum()
                sc_s = (sc_agg["Demand_Qty"]-sc_agg["Available_Qty"]).clip(lower=0).sum()
                sc_c = (sc_a/sc_d*100) if sc_d>0 else 100
                # ✅ Correct SQM
                sc_sqm = sqm_ref[
                    sqm_ref["Equipment_Tag_No."].isin(filtered_tags) &
                    (sqm_ref["Lining_System_Code"]==code)
                ]["Total_SQM"].sum()
                sc_can = round(sc_sqm * min(1.0, sc_c/100), 2)
                sc_rows.append({"Code":f"Code {code}","Short_Name":sname,
                    "SQM":sc_sqm,"SQM_Can":sc_can,"SQM_Short":round(sc_sqm-sc_can,2),
                    "Coverage_%":round(sc_c,1)})

            if sc_rows:
                sc_df = pd.DataFrame(sc_rows).sort_values("Coverage_%")
                # ── Design integration: SVG horizontal-bar ──
                st.markdown(
                    render_design_hbar(
                        [{"label": f"{r['Code']} – {str(r['Short_Name'])[:14]}",
                          "val": float(r["Coverage_%"])}
                         for _, r in sc_df.iterrows()],
                    ),
                    unsafe_allow_html=True,
                )
                bar_c = ["#10B981" if c>=100 else "#F97316" if c>=90
                         else "#EAB308" if c>=80 else "#EF4444" for c in sc_df["Coverage_%"]]
                fig_sc = go.Figure(go.Bar(
                    y=sc_df["Code"]+"  "+sc_df["Short_Name"],x=sc_df["Coverage_%"],
                    orientation="h",marker_color=bar_c,marker_opacity=.8,
                    text=[f"{v:.0f}%  ({r['SQM_Can']:,.0f}/{r['SQM']:,.0f} SQM)"
                          for v,(_,r) in zip(sc_df["Coverage_%"],sc_df.iterrows())],
                    textposition="inside",textfont=dict(family="JetBrains Mono",size=9,color="#fff"),
                ))
                fig_sc.add_vline(x=100,line_color="rgba(128,128,128,.2)",
                                 line_dash="dot",line_width=1)
                fig_sc.update_layout(paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    margin=dict(l=0,r=30,t=5,b=0),
                    height=max(220,len(sc_df)*42),
                    xaxis=dict(range=[0,115],gridcolor="rgba(128,128,128,.08)",
                               tickfont=dict(family="JetBrains Mono",size=9)),
                    yaxis=dict(gridcolor="rgba(128,128,128,.08)",
                               tickfont=dict(family="JetBrains Mono",size=9)),
                    font=dict(family="JetBrains Mono",size=10,color="rgba(148,163,184,.8)"))
                st.plotly_chart(fig_sc, use_container_width=True, key="dash_sc_bar")

                sc_show = sc_df.copy()
                sc_show.columns = ["Code","Short Name","SQM Total","Available Material Coverage (SQM)",
                                   "SQM Deficit","Coverage %"]
                sc_show[["SQM Total","Available Material Coverage (SQM)","SQM Deficit"]] = (
                    sc_show[["SQM Total","Available Material Coverage (SQM)","SQM Deficit"]].round(1))
                st.dataframe(sc_show,use_container_width=True,hide_index=True,
                             key="dash_sc_tbl")

        with row2b:
            st.markdown('<div class="sec-hdr">🧪 Coverage by Material</div>',
                        unsafe_allow_html=True)
            mat_rows_d = f_demand.copy().sort_values("Coverage_Pct")
            if not mat_rows_d.empty:
                # ── Design integration: SVG horizontal-bar ──
                st.markdown(
                    render_design_hbar(
                        [{"label": str(r["Material_Name"] or r["Material_Code"])[:18],
                          "val": float(r["Coverage_Pct"])}
                         for _, r in mat_rows_d.iterrows()],
                    ),
                    unsafe_allow_html=True,
                )
                bar_cm = ["#10B981" if c>=100 else "#F97316" if c>=90
                          else "#EAB308" if c>=80 else "#EF4444"
                          for c in mat_rows_d["Coverage_Pct"]]
                fig_mat = go.Figure(go.Bar(
                    y=mat_rows_d["Material_Code"]+"  "+mat_rows_d["Material_Name"].fillna("").str[:18],
                    x=mat_rows_d["Coverage_Pct"],orientation="h",
                    marker_color=bar_cm,marker_opacity=.8,
                    text=[f"{v:.0f}%" for v in mat_rows_d["Coverage_Pct"]],
                    textposition="inside",textfont=dict(family="JetBrains Mono",size=9,color="#fff"),
                    customdata=mat_rows_d[["Available_Qty","Demand_Qty","Shortfall","UOM"]].values,
                    hovertemplate=(
                        "<b>%{y}</b><br>Coverage: %{x:.1f}%<br>"
                        "Available: %{customdata[0]:,.1f} %{customdata[3]}<br>"
                        "Demand: %{customdata[1]:,.1f} %{customdata[3]}<br>"
                        "Shortfall: %{customdata[2]:,.1f} %{customdata[3]}<extra></extra>"
                    ),
                ))
                fig_mat.add_vline(x=100,line_color="rgba(128,128,128,.2)",
                                  line_dash="dot",line_width=1)
                fig_mat.update_layout(paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",margin=dict(l=0,r=30,t=5,b=0),
                    height=max(340,len(mat_rows_d)*34),
                    xaxis=dict(range=[0,115],gridcolor="rgba(128,128,128,.08)",
                               tickfont=dict(family="JetBrains Mono",size=9)),
                    yaxis=dict(gridcolor="rgba(128,128,128,.08)",
                               tickfont=dict(family="JetBrains Mono",size=9)),
                    font=dict(family="JetBrains Mono",size=10,color="rgba(148,163,184,.8)"))
                st.plotly_chart(fig_mat, use_container_width=True, key="dash_mat_bar")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="sec-hdr">📋 Full Material Balance</div>',
                    unsafe_allow_html=True)
        tbl_d = f_demand.sort_values("Coverage_Pct").copy()
        tbl_show = tbl_d[["Material_Code","Material_Name","UOM",
                           "Available_Qty","Ordered_Qty","Demand_Qty","Shortfall","Net_Shortfall","Coverage_Pct"]].copy()
        tbl_show.columns = ["Code","Material Name","UOM",
                             "Available","On Order","Total Demand","Shortfall","Net Shortfall","Coverage %"]

        def _style_cov(row):
            pct = row["Coverage %"]
            if pd.isna(pct): pct = 100.0
            if pct>=100:   bg,tc = "rgba(16,185,129,.1)","#10B981"
            elif pct>=90:  bg,tc = "rgba(249,115,22,.1)","#F97316"
            elif pct>=80:  bg,tc = "rgba(234,179,8,.1)", "#EAB308"
            else:          bg,tc = "rgba(239,68,68,.1)", "#EF4444"
            styles = [f"background-color:{bg}"]*len(row)
            styles[-1] = f"background-color:{bg};color:{tc};font-weight:700"
            return styles

        styled_tbl = tbl_show.style.apply(_style_cov,axis=1).format({
            "Available":"{:,.3f}","On Order":"{:,.3f}",
            "Total Demand":"{:,.3f}","Shortfall":"{:,.3f}",
            "Net Shortfall":"{:,.3f}","Coverage %":"{:.1f}%"})
        st.dataframe(styled_tbl,use_container_width=True,hide_index=True,
                     height=50+len(tbl_show)*35,key="dash_mat_tbl")

        # ── Stock-only materials (in inventory but not used in any recipe/demand) ──
        recipe_codes = set(dm["Material_Code"].unique())
        stock_only = inv[~inv["Material_Code"].isin(recipe_codes)].copy()
        if not stock_only.empty:
            st.markdown(
                '<div class="sec-hdr" style="margin-top:.8rem;">'
                '📦 Stock-Only Materials (No Demand in Any System Code)</div>',
                unsafe_allow_html=True)
            st.caption(
                "These materials are in your inventory but are not used "
                "in any lining system recipe. No demand is generated for them.")
            so_show = stock_only[["Material_Code","Material_Name","UOM",
                                   "Available_Qty","Ordered_Qty"]].copy()
            so_show["Ordered_Qty"] = so_show["Ordered_Qty"].fillna(0)
            so_show.columns = ["Code","Material Name","UOM","Available","On Order"]
            st.dataframe(so_show, use_container_width=True, hide_index=True,
                         key="dash_stock_only")

        da, db = st.columns(2)
        with da:
            st.download_button("⬇ Download Material Balance",
                data=generate_excel_report(tbl_show.reset_index(drop=True), "Material Balance", color_scheme="dashboard"),
                file_name="dashboard_material_balance.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

    # ─────────────────────────────────────────────────────────────────────────
    else:  # 🛒 Material Requirement & Procurement
    # ─────────────────────────────────────────────────────────────────────────

        st.markdown('<div class="sec-hdr">🛒 Material Requirement & Procurement — Location / System Code View</div>',
                    unsafe_allow_html=True)

        # KPI strip
        p1,p2,p3,p4,p5,p6 = st.columns(6)
        _dd_net_df = f_demand[f_demand["Net_Shortfall"]>0][["Material_Code","Material_Name","Demand_Qty","Available_Qty","Ordered_Qty","Net_Shortfall"]].sort_values("Net_Shortfall", ascending=False).reset_index(drop=True)
        with p1:
            dbl_click_metric("Equipment", str(len(filtered_tags)), "t0p_equip",
                "Equipment List", _dd_equip_df)
        with p2:
            dbl_click_metric("Total SQM", f"{proj_sqm:,.1f}", "t0p_sqm",
                "SQM by Equipment & System Code", _dd_sqm_df)
        with p3:
            dbl_click_metric("Available Coverage SQM", f"{can_sqm:,.2f}", "t0p_cov_sqm",
                "Coverable SQM by Equipment & System Code", _dd_cov_sqm_df)
        with p4:
            dbl_click_metric("SQM Deficit", f"{short_sqm:,.2f}", "t0p_deficit",
                "SQM Deficit by Equipment & System Code", _dd_def_sqm_df)
        with p5:
            dbl_click_metric("Shortfall Units", f"{f_total_short:,.0f}", "t0p_short",
                "Materials with Shortfall", _dd_def_df)
        with p6:
            dbl_click_metric("After Orders (Net)", f"{f_total_net:,.0f}", "t0p_net",
                "Materials with Net Shortfall (After Orders)", _dd_net_df)
        st.markdown("<br>", unsafe_allow_html=True)

        # Per-location, per-system-code breakdown
        for loc in sel_locations:
            loc_tags = filtered_eq[filtered_eq["Location"]==loc]["Equipment_Tag_No."].tolist()
            if not loc_tags: continue

            loc_dm = filtered_dm[filtered_dm["Equipment_Tag_No."].isin(loc_tags)]
            if loc_dm.empty: continue

            loc_sqm = sqm_ref[
                sqm_ref["Equipment_Tag_No."].isin(loc_tags) &
                sqm_ref["Lining_System_Code"].isin(sel_codes)
            ]["Total_SQM"].sum()

            loc_agg = (loc_dm.groupby("Material_Code",as_index=False)["Demand_Qty"].sum()
                .merge(inv[["Material_Code","Available_Qty","Ordered_Qty"]],
                       on="Material_Code",how="left"))
            loc_agg["Available_Qty"] = loc_agg["Available_Qty"].fillna(0)
            loc_agg["Ordered_Qty"]   = loc_agg["Ordered_Qty"].fillna(0)
            loc_agg["Shortfall"]     = (loc_agg["Demand_Qty"]-loc_agg["Available_Qty"]).clip(lower=0)
            loc_agg["Net_Shortfall"] = (loc_agg["Demand_Qty"]-loc_agg["Available_Qty"]-loc_agg["Ordered_Qty"]).clip(lower=0)
            loc_d = loc_agg["Demand_Qty"].sum()
            loc_a = loc_agg["Available_Qty"].clip(upper=loc_agg["Demand_Qty"]).sum()
            loc_c = (loc_a/loc_d*100) if loc_d>0 else 100
            loc_can_sqm = round(loc_sqm * min(1.0, loc_c/100), 2)
            loc_sh_sqm  = round(loc_sqm - loc_can_sqm, 2)
            loc_dot = "🟢" if loc_c>=100 else "🟠" if loc_c>=90 else "🟡" if loc_c>=80 else "🔴"

            badge_cls = {"Brown Field":"loc-bf","TRAIN J":"loc-tj","TRAIN K":"loc-tk"}.get(loc,"loc-bf")
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:.8rem;margin:1.2rem 0 .5rem;">'
                f'<span style="font-size:.95rem;">{loc_dot}</span>'
                f'<span class="loc-badge {badge_cls}">{loc}</span>'
                f'<span style="font-family:\'JetBrains Mono\',monospace;font-size:.75rem;'
                f'color:var(--t3);">{len(loc_tags)} equip  ·  {loc_can_sqm:,.1f}/{loc_sqm:,.1f} SQM  ·  {loc_c:.1f}%</span>'
                f'</div>', unsafe_allow_html=True)

            # Per system code within this location
            for code in sorted(sel_codes, key=lambda x: int(x)):
                code_dm = loc_dm[loc_dm["Lining_System_Code"]==code]
                if code_dm.empty: continue
                sname = code_dm["Lining_System_Short_Name"].iloc[0]

                # ✅ Correct SQM
                code_sqm = sqm_ref[
                    sqm_ref["Equipment_Tag_No."].isin(loc_tags) &
                    (sqm_ref["Lining_System_Code"]==code)
                ]["Total_SQM"].sum()

                code_agg = (code_dm.groupby(["Material_Code","Material_Name","UOM"],
                                            as_index=False)["Demand_Qty"].sum()
                    .merge(inv[["Material_Code","Available_Qty","Ordered_Qty"]],
                           on="Material_Code",how="left"))
                code_agg["Available_Qty"] = code_agg["Available_Qty"].fillna(0)
                code_agg["Ordered_Qty"]   = code_agg["Ordered_Qty"].fillna(0)
                code_agg["Shortfall"]     = (code_agg["Demand_Qty"]-code_agg["Available_Qty"]).clip(lower=0).round(3)
                code_agg["Net_Shortfall"] = (code_agg["Demand_Qty"]-code_agg["Available_Qty"]-code_agg["Ordered_Qty"]).clip(lower=0).round(3)
                code_agg["Coverage_Pct"]  = (
                    code_agg["Available_Qty"].clip(upper=code_agg["Demand_Qty"]) /
                    code_agg["Demand_Qty"].replace(0,np.nan)*100
                ).fillna(100).clip(0,100).round(1)
                code_agg["Fulfillment_Pct"] = code_agg["Coverage_Pct"]

                cd = code_agg["Demand_Qty"].sum()
                ca = code_agg["Available_Qty"].clip(upper=code_agg["Demand_Qty"]).sum()
                cc = (ca/cd*100) if cd>0 else 100
                c_can_sqm = round(code_sqm * min(1.0, cc/100), 2)
                c_sh_sqm  = round(code_sqm - c_can_sqm, 2)
                c_dot = "🟢" if cc>=100 else "🟠" if cc>=90 else "🟡" if cc>=80 else "🔴"

                with st.expander(
                    f"{c_dot}  Code {code} – {sname}  ·  "
                    f"{c_can_sqm:,.1f}/{code_sqm:,.1f} SQM  ·  {cc:.1f}%",
                    expanded=False,
                ):
                    pc1,pc2,pc3,pc4,pc5 = st.columns(5)
                    pc1.metric("System Code", f"Code {code}")
                    pc2.metric("Short Name",  sname)
                    pc3.metric("SQM Total",   f"{code_sqm:,.2f}")
                    pc4.metric("Available Material Coverage (SQM)", f"{c_can_sqm:,.2f}")
                    pc5.metric("SQM Deficit",    f"{c_sh_sqm:,.2f}")

                    # Table with Available, On Order, Demand, Shortfall, Net Shortfall
                    tbl_proc = code_agg[["Material_Code","Material_Name","UOM",
                                         "Available_Qty","Ordered_Qty","Demand_Qty",
                                         "Shortfall","Net_Shortfall","Fulfillment_Pct"]].copy()
                    tbl_proc.columns = ["Code","Material Name","UOM",
                                        "Available","On Order","Demand",
                                        "Shortfall","Net Shortfall (After Orders)","Fulfil %"]

                    def _style_proc(row):
                        pct = row["Fulfil %"]
                        if pd.isna(pct): pct = 100.0
                        if pct>=100:  bg,tc = "rgba(16,185,129,.1)","#10B981"
                        elif pct>=90: bg,tc = "rgba(249,115,22,.1)","#F97316"
                        elif pct>=80: bg,tc = "rgba(234,179,8,.1)", "#EAB308"
                        else:         bg,tc = "rgba(239,68,68,.1)", "#EF4444"
                        styles = [f"background-color:{bg}"]*len(row)
                        styles[-1] = f"background-color:{bg};color:{tc};font-weight:700"
                        return styles

                    styled_proc = tbl_proc.style.apply(_style_proc,axis=1).format({
                        "Available":"{:,.3f}","On Order":"{:,.3f}","Demand":"{:,.3f}",
                        "Shortfall":"{:,.3f}","Net Shortfall (After Orders)":"{:,.3f}",
                        "Fulfil %":"{:.1f}%"})
                    st.dataframe(styled_proc,use_container_width=True,hide_index=True,
                                 height=65+len(tbl_proc)*35,
                                 key=f"proc_{loc}_{code}")

            st.markdown('<div style="border-bottom:1px solid var(--border);margin:.8rem 0;"></div>',
                        unsafe_allow_html=True)

        # Grand total procurement table
        st.markdown('<div class="sec-hdr" style="margin-top:1rem;">📦 Grand Total — All Selected Equipment</div>',
                    unsafe_allow_html=True)

        grand = f_demand.sort_values("Coverage_Pct").copy()
        grand_show = grand[["Material_Code","Material_Name","UOM",
                             "Available_Qty","Ordered_Qty","Demand_Qty",
                             "Shortfall","Net_Shortfall","Coverage_Pct"]].copy()
        grand_show.columns = ["Code","Material Name","UOM",
                               "Available","On Order","Demand",
                               "Shortfall","Net Shortfall","Coverage %"]

        def _style_grand(row):
            pct = row["Coverage %"]
            if pd.isna(pct): pct=100.0
            if pct>=100:  bg,tc = "rgba(16,185,129,.1)","#10B981"
            elif pct>=90: bg,tc = "rgba(249,115,22,.1)","#F97316"
            elif pct>=80: bg,tc = "rgba(234,179,8,.1)", "#EAB308"
            else:         bg,tc = "rgba(239,68,68,.1)", "#EF4444"
            styles = [f"background-color:{bg}"]*len(row)
            styles[-1] = f"background-color:{bg};color:{tc};font-weight:700"
            return styles

        styled_grand = grand_show.style.apply(_style_grand,axis=1).format({
            "Available":"{:,.3f}","On Order":"{:,.3f}","Demand":"{:,.3f}",
            "Shortfall":"{:,.3f}","Net Shortfall":"{:,.3f}","Coverage %":"{:.1f}%"})
        st.dataframe(styled_grand,use_container_width=True,hide_index=True,
                     height=50+len(grand_show)*35,key="proc_grand_tbl")

        gc1, gc2 = st.columns(2)
        with gc1:
            st.download_button("⬇ Download Grand Procurement Table",
                data=generate_excel_report(grand_show.reset_index(drop=True), "Grand Procurement Table", color_scheme="dashboard"),
                file_name="procurement_grand_total.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with gc2:
            shortage_net = grand[grand["Net_Shortfall"]>0][
                ["Material_Code","Material_Name","UOM","Available_Qty","Ordered_Qty",
                 "Demand_Qty","Shortfall","Net_Shortfall"]].copy()
            shortage_net.columns = ["Code","Material Name","UOM","Available","On Order",
                                    "Demand","Shortfall","NET TO ORDER"]
            if not shortage_net.empty:
                st.download_button("⬇ Net Order List Only",
                    data=generate_excel_report(shortage_net.reset_index(drop=True), "Net Order List", color_scheme="dashboard"),
                    file_name="net_order_list.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 · EQUIPMENT ENTRY
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    left, right = st.columns([1, 1.65], gap="large")

    # ── LEFT: filter + search + session list ────────────────────────────────
    with left:
        st.markdown('<div class="sec-hdr">🎛 Filter Equipment</div>',
                    unsafe_allow_html=True)

        # 3 filter selectors
        f_loc  = st.multiselect(" Location", options=LOCATION_ORDER,
                                 default=[], key="t1_loc",
                                 placeholder="All locations")
        all_types = sorted(eq_master["Type"].str.strip().unique().tolist())
        f_type = st.multiselect(" Type", options=all_types,
                                 default=[], key="t1_type",
                                 placeholder="All types")
        all_codes_t1 = sorted(
            dm["Lining_System_Code"].unique().tolist(), key=lambda x: int(x))
        f_code = st.multiselect(
            " System Code", options=all_codes_t1,
            format_func=lambda c: f"Code {c} – "
                f"{dm[dm['Lining_System_Code']==c]['Lining_System_Short_Name'].iloc[0]}",
            default=[], key="t1_code", placeholder="All system codes")

        # Build filtered tag list
        filtered_eq = eq_master.copy()
        if f_loc:
            filtered_eq = filtered_eq[filtered_eq["Location"].isin(f_loc)]
        if f_type:
            filtered_eq = filtered_eq[filtered_eq["Type"].str.strip().isin(f_type)]
        if f_code:
            tags_with_code = dm[dm["Lining_System_Code"].isin(f_code)][
                "Equipment_Tag_No."].unique().tolist()
            filtered_eq = filtered_eq[filtered_eq["Equipment_Tag_No."].isin(tags_with_code)]
        filtered_tags_t1 = sorted(filtered_eq["Equipment_Tag_No."].tolist())

        st.markdown('<div class="sec-hdr" style="margin-top:.8rem;">🔍 Find Equipment</div>',
                    unsafe_allow_html=True)
        selected_tag = st.selectbox(
            "tag_search", [""] + filtered_tags_t1,
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

            # Show fulfillment rows with ✕ remove button
            alloc_df2 = cascade_allocate(new_order)
            for idx_t, t in enumerate(new_order):
                pct  = tag_fulfillment(alloc_df2, t)
                name = tag_name.get(t, t)
                loc  = tag_loc.get(t, "—")
                dot  = status_dot(pct)
                tag_total_sqm = sqm_ref[sqm_ref["Equipment_Tag_No."]==t]["Total_SQM"].sum()
                can_sqm  = round(tag_total_sqm * min(1.0, pct/100), 2)
                row_c, row_x = st.columns([9, 1])
                with row_c:
                    # Append system codes to session list display
                    _t_codes = sorted(
                        dm[dm["Equipment_Tag_No."]==t]["Lining_System_Code"].unique(),
                        key=lambda x: int(x))
                    _codes_badges = "  ".join(
                        f'<span style="font-family:\'JetBrains Mono\',monospace;'
                        f'font-size:.58rem;background:rgba(245,158,11,.15);'
                        f'color:var(--amber);border-radius:3px;padding:.1rem .3rem;">C{c}</span>'
                        for c in _t_codes)
                    _sess_parts = [
                        f'<div class="session-equip" style="margin-bottom:.22rem;">',
                        f'<span class="{dot}" style="font-family:JetBrains Mono,monospace;',
                        f'font-size:.75rem;font-weight:600;color:var(--t1);">{t}</span>',
                        f'<span style="font-size:.75rem;color:var(--t3);margin-left:.5rem;">{name[:22]}</span>',
                        f'<span style="margin-left:.4rem;">{_codes_badges}</span>',
                        f'<span style="font-family:JetBrains Mono,monospace;',
                        f'font-size:.62rem;color:var(--t4);margin-left:.4rem;">{loc}</span>',
                        f'<span style="float:right;font-family:JetBrains Mono,monospace;',
                        f'font-size:.64rem;color:var(--t3);">{can_sqm:,.1f}/{tag_total_sqm:,.1f} SQM&nbsp;&nbsp;</span>',
                        f'<span style="float:right;">{fulfil_pill(pct)}</span>',
                        '</div>',
                    ]
                    st.markdown("".join(_sess_parts), unsafe_allow_html=True)
                with row_x:
                    if st.button("✕", key=f"rm_{t}_{idx_t}", help=f"Remove {t}"):
                        st.session_state.session_tags.remove(t)
                        st.rerun()

            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑 Clear All", key="clear_all"):
                st.session_state.session_tags = []
                st.rerun()

    # ── RIGHT: equipment info card + system-code material tables ─────────────
    with right:
        if not selected_tag:
            st.markdown("""
            <div style="text-align:center;padding:4rem 1rem;">
              <div style="font-size:2.5rem;opacity:.12;margin-bottom:.8rem;"></div>
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
                f'<div><span style="color:var(--t4);">Substrate</span><br>'
                f'<span style="color:var(--t1);">{row["Substrate"] or "—"}</span></div>'
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
                    expanded=False,
                ):
                    mi1,mi2,mi3,mi4 = st.columns(4)
                    _t1_mat_dd = mat_rows[["Material_Code","Material_Name","Demand_Qty","Allocated_Qty","Shortfall_Qty","Fulfillment_Pct"]].reset_index(drop=True)
                    _t1_sk = f"t1_sc_{selected_tag}_{code}"
                    with mi1:
                        dbl_click_metric("System Code", str(code), f"{_t1_sk}_c",
                            f"Code {code} — Material Breakdown", _t1_mat_dd)
                    with mi2:
                        dbl_click_metric("Short Name", str(sname), f"{_t1_sk}_n",
                            f"{sname} — Material Breakdown", _t1_mat_dd)
                    with mi3:
                        dbl_click_metric("Surface Area", f"{sqm:,.2f} SQM", f"{_t1_sk}_s",
                            f"Code {code} — Material Breakdown", _t1_mat_dd)
                    with mi4:
                        dbl_click_metric("Coverage", f"{pct:.1f}%", f"{_t1_sk}_p",
                            f"Code {code} — Coverage Detail", _t1_mat_dd)
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
        _t2_equip_dd = pd.DataFrame({
            "Equipment Tag": session_tags,
            "Name":          [tag_name.get(t, t) for t in session_tags],
            "Location":      [tag_loc.get(t, "—") for t in session_tags],
        })
        _t2_mats_dd  = (alloc_df.groupby(["Material_Code","Material_Name"], as_index=False)
                        [["Demand_Qty","Allocated_Qty","Shortfall_Qty"]].sum()
                        .sort_values("Material_Code").reset_index(drop=True))
        _t2_mats_dd["Coverage_%"] = (_t2_mats_dd["Allocated_Qty"] /
            _t2_mats_dd["Demand_Qty"].replace(0, np.nan) * 100).fillna(100).clip(0,100).round(1)
        _t2_order_dd = _t2_mats_dd[_t2_mats_dd["Shortfall_Qty"]>0].reset_index(drop=True)
        _t2_cov_dd   = _t2_mats_dd.sort_values("Coverage_%").reset_index(drop=True)
        with k1:
            dbl_click_metric("Equipment", str(len(session_tags)), "t2_equip",
                "Session Equipment List", _t2_equip_dd)
        with k2:
            dbl_click_metric("Materials", str(n_mats), "t2_mats",
                "Material Demand Summary", _t2_mats_dd)
        with k3:
            dbl_click_metric("Need to Order", str(n_short_m), "t2_order",
                "Materials to Procure", _t2_order_dd)
        with k4:
            dbl_click_metric("Total Shortfall", f"{tot_short:,.1f}", "t2_short",
                "Shortfall Detail", _t2_order_dd)
        with k5:
            dbl_click_metric("Overall Coverage", f"{ov_pct:.1f}%", "t2_cov",
                "Coverage by Material", _t2_cov_dd)
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

            t_sqm      = sqm_ref[sqm_ref["Equipment_Tag_No."]==tag]["Total_SQM"].sum()
            t_can_sqm  = round(t_sqm * min(1.0, t_pct/100), 2)
            t_dot_char = "🟢" if t_pct>=100 else "🟠" if t_pct>=90 else "🟡" if t_pct>=80 else "🔴"
            _t2_type = str(eq_row.get("Type","") or "").strip()
            _t2_desc = str(eq_row.get("Substrate","") or "").strip()[:20]
            _t2_meta = "  |  ".join(p for p in [_t2_type,_t2_desc] if p and p not in ("nan","—"))
            with st.expander(
                f"{t_dot_char}  #{i+1}  {tag}  ·  {name}  ·  {_t2_meta}  ·  {loc}  "
                f"·  {t_can_sqm:,.1f}/{t_sqm:,.1f} SQM  ·  {t_pct:.1f}%",
                expanded=False,
            ):
                # Equipment meta strip
                m1,m2,m3,m4 = st.columns(4)
                m1.markdown(f'**Type:** {eq_row["Type"]}')
                m2.markdown(f'**Substrate:** {eq_row["Substrate"] or "—"}')
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

                    _, c_can_sqm, c_short_sqm = sqm_can_do(alloc_df, tag, code)
                    c_dot = "🟢" if c_pct>=100 else "🟠" if c_pct>=90 else "🟡" if c_pct>=80 else "🔴"
                    st.markdown(
                        f'<div class="syscode-block">'
                        f'<div class="syscode-hdr">'
                        f'<span style="font-size:.85rem;">{c_dot}</span>'
                        f'<span class="code-badge">Code {code}</span>'
                        f'<span style="font-size:.8rem;color:var(--t1);font-weight:500;">'
                        f'{sname}</span>'
                        f'<span style="font-family:\'JetBrains Mono\',monospace;'
                        f'font-size:.72rem;color:var(--t3);">'
                        f'{c_can_sqm:,.1f}/{sqm:,.1f} SQM</span>'
                        f'<span style="margin-left:auto;">{fulfil_pill(c_pct)}</span>'
                        f'</div></div>',
                        unsafe_allow_html=True)

                    sc1,sc2,sc3,sc4 = st.columns(4)
                    _t2e_dd = code_alloc[["Material_Code","Material_Name","Demand_Qty","Allocated_Qty","Shortfall_Qty"]].reset_index(drop=True)
                    _t2e_sk = f"t2e_{tag}_{code}"
                    with sc1:
                        dbl_click_metric("Demand", f"{c_demand:,.3f}", f"{_t2e_sk}_d",
                            f"Code {code} — {sname}: Material Detail", _t2e_dd)
                    with sc2:
                        dbl_click_metric("Allocated", f"{c_alloc:,.3f}", f"{_t2e_sk}_a",
                            f"Code {code} — {sname}: Material Detail", _t2e_dd)
                    if c_short > 0.001:
                        with sc3:
                            dbl_click_metric("Shortfall", f"{c_short:,.3f}", f"{_t2e_sk}_sh",
                                f"Code {code} — {sname}: Shortfall Detail",
                                _t2e_dd[_t2e_dd["Shortfall_Qty"]>0].reset_index(drop=True))
                        with sc4:
                            dbl_click_metric("SQM Deficit", f"{c_short_sqm:,.2f}", f"{_t2e_sk}_sq",
                                f"Code {code} — {sname}: Shortfall Detail",
                                _t2e_dd[_t2e_dd["Shortfall_Qty"]>0].reset_index(drop=True))
                    plotly_mat_table(
                        code_alloc,
                        f"rep_{tag}_{code}",
                        height=65 + len(code_alloc)*30,
                        show_sqm=True, tag=tag, code=code
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
                    f'{tag_alloc["Demand_Qty"].sum():,.3f}</strong></span>'
                    f'<span style="color:var(--t3);margin-left:1rem;">'
                    f'Allocated: <strong style="color:var(--t1);">'
                    f'{tag_alloc["Allocated_Qty"].sum():,.3f}</strong></span>'
                    + (f'<span style="color:var(--t3);margin-left:1rem;">'
                    f'Shortfall: <strong style="color:#EF4444;">'
                    f'{tag_alloc["Shortfall_Qty"].sum():,.3f}</strong></span>'
                    if tag_alloc["Shortfall_Qty"].sum() > 0.001 else "") +
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

        # ── SQM per material: sum of SQM across every (tag,code) cell it touches,
        #    weighted by that cell's fulfillment ──────────────────────────────
        _sqm_per_mat = alloc_df.copy()
        _sqm_per_mat["SQM_Done_Cell"] = (
            _sqm_per_mat["Total_SQM"] * _sqm_per_mat["Fulfillment_Pct"] / 100
        )
        _sqm_agg = _sqm_per_mat.groupby("Material_Code", as_index=False).agg(
            SQM_Total =("Total_SQM",     "sum"),
            SQM_Done  =("SQM_Done_Cell", "sum"),
        )
        _sqm_agg["SQM_Deficit"] = (_sqm_agg["SQM_Total"] - _sqm_agg["SQM_Done"]).round(2)
        _sqm_agg["SQM_Total"]   = _sqm_agg["SQM_Total"].round(2)
        _sqm_agg["SQM_Done"]    = _sqm_agg["SQM_Done"].round(2)
        combined = combined.merge(_sqm_agg, on="Material_Code", how="left")
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
                               data=generate_excel_report(alloc_df, "Session Full Report", color_scheme="session"),
                               file_name="session_full_report.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        with d2:
            if not shortage_only.empty:
                st.download_button("⬇ Order List Only",
                                   data=generate_excel_report(shortage_only, "Order List", color_scheme="session"),
                                   file_name="order_list.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)


        # ── Smart Reordering Suggestions ──────────────────────────────────────
        render_suggestion_panel(session_tags, "tab2")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 · LOCATION REPORT
# ═══════════════════════════════════════════════════════════════════════════════
with tab3:
    loc_report_mode = st.radio(
        "View Mode",
        ["📍 Location Based", "🌐 All Equipment"],
        horizontal=True, key="loc_report_mode",
        label_visibility="collapsed",
    )
    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Per-location order state (independent from global session_tags) ──────
    if "loc_order" not in st.session_state:
        st.session_state.loc_order = {}

    # ══════════════════════════════════════════════════════════════════════════
    # ALL EQUIPMENT MODE
    # ══════════════════════════════════════════════════════════════════════════
    if loc_report_mode == "🌐 All Equipment":
        st.markdown('<div class="sec-hdr">🌐 All Equipment — Global Cascading Balance</div>',
                    unsafe_allow_html=True)
        st.caption("All equipment in file order. Inventory pool is shared globally across all locations.")

        from streamlit_sortables import sort_items as _sort3_all

        if "all_eq_order" not in st.session_state:
            st.session_state.all_eq_order = eq_master["Equipment_Tag_No."].tolist()

        all_eq_tags = st.session_state.all_eq_order

        # ── Sortable list (static labels only) ───────────────────────────────
        def _ae_label(i, t):
            name = eq_master.set_index("Equipment_Tag_No.")["Name"].get(t, t)[:26]
            loc  = eq_master.set_index("Equipment_Tag_No.")["Location"].get(t, "")
            sqm  = eq_master.set_index("Equipment_Tag_No.")["Total_SQM"].get(t, 0)
            return f"#{i+1}  ||  {t}  ||  {name}  ||  {loc}  ||  {sqm:,.1f} SQM"

        _ae_display = [_ae_label(i, t) for i, t in enumerate(all_eq_tags)]
        _ae_key     = "ae_sort_" + str(len(all_eq_tags))
        st.caption("⇅ Drag to reorder — order determines cascade priority.")
        _ae_sorted  = _sort3_all(_ae_display, direction="vertical", key=_ae_key)

        def _ae_parse(label):
            parts = label.split("  ||  ")
            return parts[1].strip() if len(parts) > 1 else label.strip()

        _ae_new_order = [_ae_parse(l) for l in _ae_sorted if _ae_parse(l) in all_eq_tags]
        if len(_ae_new_order) == len(all_eq_tags) and _ae_new_order != st.session_state.all_eq_order:
            st.session_state.all_eq_order = _ae_new_order
            st.rerun()

        # ── Cascade allocation across all equipment ───────────────────────────
        ae_alloc = cascade_allocate(all_eq_tags)

        ae_demand = ae_alloc["Demand_Qty"].sum()
        ae_alloc_qty = ae_alloc["Allocated_Qty"].sum()
        ae_short  = ae_alloc["Shortfall_Qty"].sum()
        ae_pct    = min(100, ae_alloc_qty / ae_demand * 100) if ae_demand > 0 else 100
        ae_sqm    = sqm_ref["Total_SQM"].sum()
        ae_can_sqm = round(ae_sqm * min(1.0, ae_pct / 100), 2)

        # KPI strip
        ae_k1, ae_k2, ae_k3, ae_k4, ae_k5 = st.columns(5)
        ae_k1.metric("Equipment", str(len(all_eq_tags)))
        ae_k2.metric("Total SQM", f"{ae_sqm:,.1f}")
        ae_k3.metric("Available Coverage SQM", f"{ae_can_sqm:,.2f}")
        ae_k4.metric("SQM Deficit", f"{ae_sqm - ae_can_sqm:,.2f}")
        ae_k5.metric("Overall Coverage", f"{ae_pct:.1f}%")
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Per-equipment expanders ────────────────────────────────────────────
        st.markdown('<div class="sec-hdr">Per-Equipment Detail</div>', unsafe_allow_html=True)
        for i, tag in enumerate(all_eq_tags):
            tag_alloc_ae = ae_alloc[ae_alloc["Equipment_Tag_No."] == tag]
            t_pct_ae = tag_fulfillment(ae_alloc, tag)
            t_short_ae = tag_alloc_ae["Shortfall_Qty"].sum()
            eq_row_ae = eq_master[eq_master["Equipment_Tag_No."] == tag].iloc[0]
            _t3a_sqm = sqm_ref[sqm_ref["Equipment_Tag_No."] == tag]["Total_SQM"].sum()
            _t3a_cansqm = round(_t3a_sqm * min(1.0, t_pct_ae / 100), 2)
            _t3a_dot = "✅" if t_pct_ae >= 100 else "🟠" if t_pct_ae >= 90 else "🟡" if t_pct_ae >= 80 else "🔴"
            _t3a_type = str(eq_row_ae.get("Type", "") or "").strip()
            _t3a_desc = str(eq_row_ae.get("Substrate", "") or "").strip()[:20]
            _t3a_loc  = str(eq_row_ae.get("Location", "") or "").strip()
            _t3a_meta = "  |  ".join(p for p in [_t3a_type, _t3a_desc] if p and p not in ("nan", "—"))
            with st.expander(
                f"{_t3a_dot}  #{i+1}  {tag}  ·  {eq_row_ae['Name']}  ·  {_t3a_meta}  ·  {_t3a_loc}  "
                f"·  {_t3a_cansqm:,.1f}/{_t3a_sqm:,.1f} SQM  ·  {t_pct_ae:.1f}%",
                expanded=False,
            ):
                for code in sorted(tag_alloc_ae["Lining_System_Code"].unique(), key=lambda x: int(x)):
                    code_alloc_ae = tag_alloc_ae[tag_alloc_ae["Lining_System_Code"] == code].copy()
                    sname_ae = code_alloc_ae["Lining_System_Short_Name"].iloc[0]
                    sqm_ae   = code_alloc_ae["Total_SQM"].iloc[0]
                    c_pct_ae = syscode_fulfillment(ae_alloc, tag, code)
                    _, c_can_ae, _ = sqm_can_do(ae_alloc, tag, code)
                    c_dot_ae = "🟢" if c_pct_ae >= 100 else "🟠" if c_pct_ae >= 90 else "🟡" if c_pct_ae >= 80 else "🔴"
                    st.markdown(
                        f'<div class="syscode-block"><div class="syscode-hdr">'
                        f'<span style="font-size:.85rem;">{c_dot_ae}</span>'
                        f'<span class="code-badge">Code {code}</span>'
                        f'<span style="font-size:.8rem;color:var(--t1);">{sname_ae}</span>'
                        f'<span style="font-family:\'JetBrains Mono\',monospace;font-size:.72rem;color:var(--t3);">'
                        f'{c_can_ae:,.1f}/{sqm_ae:,.1f} SQM</span>'
                        f'<span style="margin-left:auto;">{fulfil_pill(c_pct_ae)}</span>'
                        f'</div></div>', unsafe_allow_html=True)
                    plotly_mat_table(
                        code_alloc_ae, f"ae_{tag}_{code}",
                        height=65 + len(code_alloc_ae) * 30,
                        show_sqm=True, tag=tag, code=code, allocated_label="Available"
                    )
                # Add to session
                if tag in st.session_state.session_tags:
                    st.markdown('<span style="font-family:\'JetBrains Mono\',monospace;font-size:.7rem;color:#10B981;">✓ In session</span>', unsafe_allow_html=True)
                else:
                    if st.button(f"＋ Add {tag} to Session", key=f"aeadd_{tag}"):
                        st.session_state.session_tags.append(tag)
                        st.rerun()

        # Smart Reordering Suggestions
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("💡 Smart Reordering Suggestions — All Equipment", expanded=False):
            render_suggestion_panel(all_eq_tags, "tab3_all")

        # Download
        st.markdown("<br>", unsafe_allow_html=True)
        _ae_export = ae_alloc.copy()
        st.download_button(
            "⬇ Download All Equipment Report",
            data=generate_excel_report(_ae_export, "All Equipment — Global Report", color_scheme="overview"),
            file_name=f"all_equipment_report_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_ae_all",
        )

    from streamlit_sortables import sort_items as _sort3

    loc_color = {
        "Brown Field": ("loc-bf","#3B82F6"),
        "TRAIN J":     ("loc-tj","#F59E0B"),
        "TRAIN K":     ("loc-tk","#10B981"),
    }

    if loc_report_mode == "📍 Location Based":
        st.markdown('<div class="sec-hdr">📍 All Equipment by Location — Cascading Balance</div>',
                    unsafe_allow_html=True)
        st.caption("Drag to reorder equipment within each location. Order is independent from the session list.")

    for loc in LOCATION_ORDER:
        # Skip location rendering in All Equipment mode
        if loc_report_mode != "📍 Location Based":
            # Still seed loc_order state so it's ready when user switches modes
            default_loc_order = eq_master[eq_master["Location"]==loc]["Equipment_Tag_No."].tolist()
            if loc not in st.session_state.loc_order:
                st.session_state.loc_order[loc] = default_loc_order
            continue
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

        # ── Location title above sortable list ──────────────────────────────
        st.markdown(
            f'<div style="font-family:\'JetBrains Mono\',monospace;font-size:.68rem;'
            f'font-weight:700;letter-spacing:.1em;color:var(--amber);margin-bottom:.3rem;">'
            f'📍 {loc} — Drag to set build priority</div>',
            unsafe_allow_html=True)

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

        # Location SQM (correct: sum unique (tag,code) SQM, not from dm)
        loc_sqm_total = sqm_ref[sqm_ref["Equipment_Tag_No."].isin(loc_tags_all)]["Total_SQM"].sum()
        loc_can_sqm   = round(loc_sqm_total * min(1.0, loc_pct/100), 2)
        loc_dot       = "🟢" if loc_pct>=100 else "🟠" if loc_pct>=90 else "🟡" if loc_pct>=80 else "🔴"

        # Location header
        st.markdown(
            f'<div style="display:flex;align-items:center;gap:1rem;'
            f'margin:1.2rem 0 .7rem;">'
            f'<span style="font-size:.95rem;">{loc_dot}</span>'
            f'<span class="loc-badge {badge_cls}" style="font-size:.76rem;'
            f'padding:.28rem .9rem;">{loc}</span>'
            f'<span style="font-family:\'JetBrains Mono\',monospace;font-size:.75rem;'
            f'color:var(--t4);">{len(loc_tags_all)} equip</span>'
            f'<span style="font-family:\'JetBrains Mono\',monospace;font-size:.75rem;'
            f'color:var(--t3);">{loc_can_sqm:,.1f}/{loc_sqm_total:,.1f} SQM</span>'
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

            _dot_char  = "✅" if t_pct>=100 else "🟠" if t_pct>=90 else "🟡" if t_pct>=80 else "🔴"
            _t3_sqm    = sqm_ref[sqm_ref["Equipment_Tag_No."]==tag]["Total_SQM"].sum()
            _t3_cansqm = round(_t3_sqm * min(1.0, t_pct/100), 2)
            _t3_type   = str(eq_row.get("Type","") or "").strip()
            _t3_desc   = str(eq_row.get("Substrate","") or "").strip()[:20]
            _t3_meta   = "  |  ".join(p for p in [_t3_type,_t3_desc] if p and p not in ("nan","—"))
            with st.expander(
                f"{_dot_char}  {tag}  ·  {eq_row['Name']}  ·  {_t3_meta}  ·  "
                f"{_t3_cansqm:,.1f}/{_t3_sqm:,.1f} SQM  ·  {t_pct:.1f}%",
                expanded=False,
            ):
                c1,c2,c3 = st.columns(3)
                c1.markdown(f'**Type:** {eq_row["Type"]}')
                c2.markdown(f'**Substrate:** {eq_row["Substrate"] or "—"}')
                c3.markdown(f'**Material Spec.:** {eq_row["Material_Spec"] or "—"}')
                st.caption(
                    f'**Lining:** {str(eq_row["Lining_Systems"]).replace(chr(10)," | ")}')
                if t_pct >= 100:
                    st.markdown(
                        '<div style="background:var(--green-bg);border:1px solid var(--green);'
                        'border-radius:6px;padding:.5rem .9rem;margin-bottom:.5rem;'
                        'font-family:\'JetBrains Mono\',monospace;font-size:.78rem;color:var(--green);">'
                        '✅ All materials fully covered — ready to proceed</div>',
                        unsafe_allow_html=True)
                st.markdown("---")

                # Per system code
                for code in sorted(tag_alloc["Lining_System_Code"].unique(),
                                   key=lambda x: int(x)):
                    code_alloc = tag_alloc[tag_alloc["Lining_System_Code"]==code].copy()
                    sname = code_alloc["Lining_System_Short_Name"].iloc[0]
                    sqm   = code_alloc["Total_SQM"].iloc[0]
                    c_pct = syscode_fulfillment(loc_alloc, tag, code)

                    _, c3_can, c3_short_sqm = sqm_can_do(loc_alloc, tag, code)
                    c3_dot = "🟢" if c_pct>=100 else "🟠" if c_pct>=90 else "🟡" if c_pct>=80 else "🔴"
                    st.markdown(
                        f'<div class="syscode-block">'
                        f'<div class="syscode-hdr">'
                        f'<span style="font-size:.85rem;">{c3_dot}</span>'
                        f'<span class="code-badge">Code {code}</span>'
                        f'<span style="font-size:.8rem;color:var(--t1);">{sname}</span>'
                        f'<span style="font-family:\'JetBrains Mono\',monospace;'
                        f'font-size:.72rem;color:var(--t3);">'
                        f'{c3_can:,.1f}/{sqm:,.1f} SQM</span>'
                        f'<span style="margin-left:auto;">{fulfil_pill(c_pct)}</span>'
                        f'</div></div>',
                        unsafe_allow_html=True)
                    plotly_mat_table(
                        code_alloc,
                        f"loc_{loc}_{tag}_{code}",
                        height=65 + len(code_alloc)*30,
                        show_sqm=True, tag=tag, code=code,
                        allocated_label="Available"
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
                    f'{tag_alloc["Demand_Qty"].sum():,.3f}</b></span>'
                    + (f'<span style="color:var(--t3);margin-left:.8rem;">'
                    f'Shortfall: <b style="color:#EF4444;">'
                    f'{tag_alloc["Shortfall_Qty"].sum():,.3f}</b></span>'
                    if tag_alloc["Shortfall_Qty"].sum() > 0.001 else "") +
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


    # ── Per-location Excel downloads ────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    if loc_report_mode == "📍 Location Based":
        st.markdown('<div class="sec-hdr">📥 Download Report per Location</div>',
                    unsafe_allow_html=True)
    _all_loc_sheets = []
    dl_loc_cols = st.columns(len(LOCATION_ORDER))
    for _dl_col, _loc_dl in zip(dl_loc_cols, LOCATION_ORDER):
        if loc_report_mode != "📍 Location Based":
            continue
        _loc_tags_dl = st.session_state.loc_order.get(
            _loc_dl,
            eq_master[eq_master["Location"] == _loc_dl]["Equipment_Tag_No."].tolist()
        )
        if not _loc_tags_dl:
            continue
        _loc_alloc_dl = cascade_allocate(_loc_tags_dl)
        if _loc_alloc_dl.empty:
            continue
        # Enrich with inventory data for export
        _inv_lu = inv[["Material_Code","Available_Qty","Ordered_Qty"]].groupby(
            "Material_Code", as_index=False).first()
        _loc_report = _loc_alloc_dl.merge(_inv_lu, on="Material_Code", how="left")
        _loc_report["Available_Qty"] = _loc_report["Available_Qty"].fillna(0)
        _loc_report["Ordered_Qty"]   = _loc_report["Ordered_Qty"].fillna(0)
        _export_cols = [
            "Equipment_Tag_No.", "Lining_System_Code", "Lining_System_Short_Name",
            "Total_SQM", "Material_Code", "Material_Name", "UOM",
            "Demand_Qty", "Available_Qty", "Ordered_Qty",
            "Allocated_Qty", "Shortfall_Qty", "Fulfillment_Pct",
        ]
        _export_df = _loc_report[[c for c in _export_cols if c in _loc_report.columns]]
        _loc_cs = _LOC_COLOR_MAP.get(_loc_dl, "dashboard")
        _all_loc_sheets.append({
            "name":            _loc_dl[:31],
            "df":              _export_df,
            "title":           f"Location Report — {_loc_dl}",
            "color_scheme":    _loc_cs,
            "add_grand_total": True,
        })
        _dl_col.download_button(
            label=f"⬇ {_loc_dl}",
            data=generate_excel_report(_export_df, f"Location Report – {_loc_dl}",
                                       color_scheme=_loc_cs),
            file_name=f"location_report_{_loc_dl.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"dl_loc_{_loc_dl}",
        )

    if _all_loc_sheets:
        st.download_button(
            "⬇ All Locations — Combined (Multi-Sheet)",
            data=generate_multi_sheet_excel(_all_loc_sheets),
            file_name=f"location_report_all_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_loc_all",
        )

    if loc_report_mode == "📍 Location Based":
        # ── Print Report button ───────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="sec-hdr">🖨 Print Location Report</div>',
                    unsafe_allow_html=True)
        st.markdown("""
        <button onclick="window.print()"
            style="font-family:'JetBrains Mono',monospace;font-size:.68rem;
                   font-weight:700;letter-spacing:.08em;text-transform:uppercase;
                   background:#F59E0B;color:#000;border:none;border-radius:4px;
                   padding:.52rem 1.3rem;cursor:pointer;transition:all .15s;">
            🖨 Print / Save as PDF
        </button>
        <style>
        @media print {
            [data-testid="stSidebar"], [data-testid="stHeader"],
            .sticky-header-wrap, [data-testid="stTabs"] > div:first-of-type,
            button[onclick="window.print()"] { display:none!important; }
            [data-testid="stExpander"] { break-inside:avoid; }
            body { background:#fff!important; color:#000!important; }
        }
        </style>""", unsafe_allow_html=True)

        # ── Smart Reordering Suggestions per location ─────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        for _loc_sugg in LOCATION_ORDER:
            _loc_tags_sugg = st.session_state.loc_order.get(
                _loc_sugg,
                eq_master[eq_master["Location"]==_loc_sugg]["Equipment_Tag_No."].tolist()
            )
            if len(_loc_tags_sugg) < 2:
                continue
            with st.expander(f"💡 Smart Reordering Suggestions — {_loc_sugg}", expanded=False):
                render_suggestion_panel(_loc_tags_sugg, f"tab3_{_loc_sugg}")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 4 · EXECUTION PLAN
# ═══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="sec-hdr">⚙️ Execution Plan — Critical System Code Analysis</div>',
                unsafe_allow_html=True)

    exec_subview = st.radio(
        "View", ["⚙️ Execution Plan", "📋 Progress List"],
        horizontal=True, key="exec_subview", label_visibility="collapsed"
    )
    st.markdown("<hr>", unsafe_allow_html=True)

    if exec_subview == "📋 Progress List":
        if not db_available():
            st.warning("Database required for the Progress List.")
        else:
            conn = get_db()
            prog_df = pd.read_sql("""
                WITH dynamic_done AS (
                    SELECT
                        equipment_tag,
                        lining_system_code,
                        COALESCE(SUM(sqm_completed), 0.0) AS done_sqm
                    FROM consumption_log
                    GROUP BY equipment_tag, lining_system_code
                )
                SELECT
                    sp.equipment_tag                                        AS "Equipment Tag",
                    sp.lining_system_code                                   AS "System Code",
                    e.lining_system_short_name                              AS "System Name",
                    e.location                                              AS "Location",
                    e.name                                                  AS "Equipment Name",
                    sp.original_sqm                                         AS "Total SQM",
                    COALESCE(dd.done_sqm, 0.0)                             AS "Done SQM",
                    (sp.original_sqm - COALESCE(dd.done_sqm, 0.0))        AS "Remaining SQM",
                    ROUND(COALESCE(dd.done_sqm, 0.0) * 100.0
                          / NULLIF(sp.original_sqm, 0), 1)                 AS "Completion %"
                FROM sqm_progress sp
                LEFT JOIN dynamic_done dd
                       ON sp.equipment_tag      = dd.equipment_tag
                      AND sp.lining_system_code = dd.lining_system_code
                LEFT JOIN equipment e
                       ON sp.equipment_tag      = e.equipment_tag
                      AND sp.lining_system_code = e.lining_system_code
                ORDER BY e.location, sp.equipment_tag,
                         CAST(sp.lining_system_code AS INTEGER)
            """, conn)
            conn.close()

            prog_df["Status"] = prog_df["Completion %"].apply(
                lambda p: "✅ Complete"    if (p or 0) >= 100
                          else "🔄 In Progress" if (p or 0) > 0
                          else "⏳ Not Started"
            )

            tot_orig = prog_df["Total SQM"].sum()
            tot_done = prog_df["Done SQM"].sum()
            tot_rem  = prog_df["Remaining SQM"].sum()
            tot_pct  = (tot_done / tot_orig * 100) if tot_orig > 0 else 0.0

            pk1, pk2, pk3, pk4 = st.columns(4)
            pk1.metric("Total SQM",     f"{tot_orig:,.2f}")
            pk2.metric("Done SQM",      f"{tot_done:,.2f}")
            pk3.metric("Remaining SQM", f"{tot_rem:,.2f}")
            pk4.metric("Completion",    f"{tot_pct:.1f}%")
            st.markdown("<br>", unsafe_allow_html=True)

            pf1, pf2 = st.columns(2)
            with pf1:
                prog_locs   = ["All"] + sorted(prog_df["Location"].dropna().unique().tolist())
                prog_loc_f  = st.selectbox("Filter Location", prog_locs, key="prog_loc_f")
            with pf2:
                prog_stat_opts = ["All", "✅ Complete", "🔄 In Progress", "⏳ Not Started"]
                prog_status_f  = st.selectbox("Filter Status", prog_stat_opts, key="prog_status_f")

            filt_prog = prog_df.copy()
            if prog_loc_f    != "All": filt_prog = filt_prog[filt_prog["Location"] == prog_loc_f]
            if prog_status_f != "All": filt_prog = filt_prog[filt_prog["Status"]   == prog_status_f]
            filt_prog = filt_prog.reset_index(drop=True)

            def _style_prog(row):
                p = row["Completion %"] or 0
                if p >= 100:  bg, tc = "rgba(16,185,129,.1)", "#10B981"
                elif p > 0:   bg, tc = "rgba(245,158,11,.1)",  "#F59E0B"
                else:         bg, tc = "rgba(239,68,68,.1)",   "#EF4444"
                styles = [f"background-color:{bg}"] * len(row)
                ci = list(row.index).index("Completion %")
                styles[ci] = f"background-color:{bg};color:{tc};font-weight:700"
                return styles

            st.dataframe(
                filt_prog.style.apply(_style_prog, axis=1).format({
                    "Total SQM":     "{:,.2f}",
                    "Done SQM":      "{:,.2f}",
                    "Remaining SQM": "{:,.2f}",
                    "Completion %":  "{:.1f}%",
                }),
                use_container_width=True, hide_index=True,
                height=min(700, 60 + len(filt_prog) * 35),
                key="prog_list_tbl"
            )

            st.download_button(
                "⬇ Download Progress List",
                data=generate_excel_report(
                    filt_prog.drop(columns=["Status"], errors="ignore").reset_index(drop=True),
                    f"Progress List — {date.today()}", color_scheme="overview"),
                file_name=f"progress_list_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_prog_list"
            )

    session_tags = st.session_state.session_tags

    if exec_subview == "📋 Progress List":
        pass  # progress list already rendered above
    elif not session_tags:
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
                    f'⚠️ <strong style="color:#EF4444;">{crit_short:,.3f} units</strong>'
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
                       f"<strong style='color:#EF4444;'>{crit_short:,.3f} units</strong> first."}
                {"" if other_short==0 else
                  f" Additionally, other system codes require "
                  f"<strong style='color:#F59E0B;'>{other_short:,.3f} units</strong>"
                  f" across {len(all_short_df[~all_short_df['Lining_System_Code'].isin([sel_code])])} material(s)."}
                <br>
                <strong style="color:var(--t0);">
                Total to order for full completion: {total_to_order:,.3f} units
                across {len(all_short_df)} material(s).</strong>
              </div>
            </div>""", unsafe_allow_html=True)

            if not all_short_df.empty:
                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button(
                    f"⬇ Download Execution Order List — {sel_tag}",
                    data=generate_excel_report(
                        all_short_df[
                            ["Lining_System_Code","Lining_System_Short_Name",
                             "Material_Code","Material_Name","UOM",
                             "Demand_Qty","Allocated_Qty","Shortfall_Qty","Fulfillment_Pct"]
                        ].sort_values(["Lining_System_Code","Shortfall_Qty"],
                                      ascending=[True,False]),
                        f"Execution Plan – {sel_tag}", color_scheme="execution"),
                    file_name=f"execution_plan_{sel_tag.replace('/','-')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

# ═══════════════════════════════════════════════════════════════════════════════
# TAB: DAILY CONSUMPTION ENTRY
# ═══════════════════════════════════════════════════════════════════════════════
with tab_consume:
    st.markdown('<div class="sec-hdr">📦 Inventory — Consumption & Receipts</div>',
                unsafe_allow_html=True)

    if not db_available():
        st.error("Database not found. Run `python setup_db.py` first to initialise the database.")
        st.stop()

    inv_mode = st.radio(
        "Mode", ["📊 Main Inventory", "📅 Consumption", "📦 Receipts", "📋 Ordered"],
        horizontal=True, key="inv_mode", label_visibility="collapsed")
    st.markdown("<hr>", unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════════
    # CONSUMPTION MODE
    # ═══════════════════════════════════════════════════════════════════════════
    if inv_mode == "📅 Consumption":

        # ── Cascading dropdowns: Location → Type → Equipment → System Code ───
        st.markdown('<div class="sec-hdr">Step 1 — Select Work Location & Equipment</div>',
                    unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            ce_loc = st.selectbox(" Location", options=[""] + LOCATION_ORDER,
                                  key="ce_loc", label_visibility="visible")
        with col2:
            if ce_loc:
                type_opts = sorted(
                    eq_master[eq_master["Location"]==ce_loc]["Type"].dropna().unique().tolist())
            else:
                type_opts = sorted(eq_master["Type"].dropna().unique().tolist())
            ce_type = st.selectbox(" Type", options=[""] + type_opts,
                                   key="ce_type", label_visibility="visible")
        with col3:
            eq_filter = eq_master.copy()
            if ce_loc:  eq_filter = eq_filter[eq_filter["Location"]==ce_loc]
            if ce_type: eq_filter = eq_filter[eq_filter["Type"]==ce_type]
            tag_opts = sorted(eq_filter["Equipment_Tag_No."].tolist())
            ce_tag = st.selectbox(" Equipment Tag", options=[""] + tag_opts,
                                  format_func=lambda t: "" if t=="" else _eq_label(t),
                                  key="ce_tag", label_visibility="visible")
        with col4:
            if ce_tag:
                code_opts = sorted(
                    equip_sc[equip_sc["Equipment_Tag_No."]==ce_tag]["Lining_System_Code"].unique(),
                    key=lambda x: int(x))
                code_labels = [
                    f"Code {c}  –  {equip_sc[(equip_sc['Equipment_Tag_No.']==ce_tag)&(equip_sc['Lining_System_Code']==c)]['Lining_System_Short_Name'].iloc[0]}"
                    for c in code_opts]
            else:
                code_opts, code_labels = [], []
            ce_code_raw = st.selectbox(" System Code", options=[""] + code_labels,
                                       key="ce_code", label_visibility="visible")
            ce_code = ce_code_raw.split("  –  ")[0].replace("Code ","").strip() if ce_code_raw else ""

        # ── SQM Entry ─────────────────────────────────────────────────────────
        if ce_tag and ce_code:
            st.markdown("<hr>", unsafe_allow_html=True)
            st.markdown('<div class="sec-hdr">Step 2 — Enter SQM Completed Today</div>',
                        unsafe_allow_html=True)

            sqm_row = sqm_ref[
                (sqm_ref["Equipment_Tag_No."]==ce_tag) &
                (sqm_ref["Lining_System_Code"]==ce_code)]
            total_sqm_orig  = float(sqm_row["Total_SQM_Original"].iloc[0]) if not sqm_row.empty else 0
            done_sqm_prev   = float(sqm_row["done_sqm"].iloc[0]) if not sqm_row.empty else 0
            remaining_sqm   = max(0.0, total_sqm_orig - done_sqm_prev)
            sname           = (equip_sc[(equip_sc["Equipment_Tag_No."]==ce_tag) &
                                        (equip_sc["Lining_System_Code"]==ce_code)]
                               ["Lining_System_Short_Name"].iloc[0]
                               if not sqm_row.empty else ce_code)

            sc1,sc2,sc3,sc4 = st.columns(4)
            _consume_rec = recipe[recipe["Lining_System_Code"]==ce_code][["Material_Code","Material_Name","For_1_SQM","UOM"]].copy()
            _consume_rec["Demand for Remaining SQM"] = (_consume_rec["For_1_SQM"] * remaining_sqm).round(4)
            _consume_rec = _consume_rec.reset_index(drop=True)
            with sc1:
                dbl_click_metric("System Code", str(sname), f"tce_sc_{ce_tag}_{ce_code}",
                    f"{sname} — Recipe Detail", _consume_rec)
            with sc2:
                dbl_click_metric("Original SQM", f"{total_sqm_orig:,.2f}", f"tce_orig_{ce_tag}_{ce_code}",
                    f"{sname} — Recipe Detail", _consume_rec)
            with sc3:
                dbl_click_metric("Already Done SQM", f"{done_sqm_prev:,.2f}", f"tce_done_{ce_tag}_{ce_code}",
                    f"{sname} — Completed so far ({done_sqm_prev:,.2f} SQM)", _consume_rec)
            with sc4:
                dbl_click_metric("Remaining SQM", f"{remaining_sqm:,.2f}", f"tce_rem_{ce_tag}_{ce_code}",
                    f"{sname} — Material Needed for Remaining {remaining_sqm:,.2f} SQM", _consume_rec,
                    help_text="Remaining = Original − already completed")

            sc_recipe = recipe[recipe["Lining_System_Code"]==ce_code].copy()
            sc_recipe = sc_recipe.merge(
                inv[["Material_Code","Available_Qty","Ordered_Qty"]],
                on="Material_Code", how="left")
            sc_recipe["Available_Qty"] = sc_recipe["Available_Qty"].fillna(0)
            sc_recipe["Ordered_Qty"]   = sc_recipe["Ordered_Qty"].fillna(0)

            if sc_recipe.empty:
                st.warning(f"No recipe found for System Code {ce_code}.")
            else:
                st.markdown("<hr>", unsafe_allow_html=True)
                with st.form(key="ce_form", clear_on_submit=False):
                    st.markdown('<div class="sec-hdr">Step 2 — Enter SQM Completed Today</div>',
                                unsafe_allow_html=True)
                    col_date, col_sqm = st.columns(2)
                    with col_date:
                        ce_date_form = st.date_input("📅 Work Date", value=date.today(),
                                                     key="form_ce_date")
                    with col_sqm:
                        ce_sqm_form = st.number_input(
                            "SQM Completed Today",
                            min_value=0.0, max_value=float(remaining_sqm),
                            value=0.0, step=0.5, format="%.2f",
                            key="form_ce_sqm",
                            help=f"Maximum {remaining_sqm:,.2f} m² remaining")

                    st.markdown('<div class="sec-hdr" style="margin-top:.8rem;">'
                                'Step 3 — Material Quantities Consumed</div>',
                                unsafe_allow_html=True)
                    st.caption(
                        "Actual Consumed defaults to For_1_SQM × SQM if left at 0. "
                        "Override with actual site usage if different.")

                    h1,h2,h3,h4,h5,h6,h7 = st.columns([2,3,1,1.5,1.5,1.5,1.5])
                    for hdr, col in zip(
                        ["Code","Material Name","UOM","Available",
                         "For 1 SQM","Actual Consumed","On Order"],
                        [h1,h2,h3,h4,h5,h6,h7]
                    ):
                        col.markdown(f"**{hdr}**")
                    st.markdown("---")

                    mat_inputs = {}
                    for _, mrow in sc_recipe.iterrows():
                        mc    = str(mrow["Material_Code"])
                        for_1 = float(mrow.get("For_1_SQM", 0) or 0)
                        avail = float(mrow.get("Available_Qty", 0) or 0)
                        onord = float(mrow.get("Ordered_Qty",  0) or 0)
                        c1,c2,c3,c4,c5,c6,c7 = st.columns([2,3,1,1.5,1.5,1.5,1.5])
                        c1.markdown(f"<code>{mc}</code>", unsafe_allow_html=True)
                        c2.write(str(mrow.get("Material_Name", "")))
                        c3.write(str(mrow.get("UOM", "")))
                        c4.write(f"{avail:,.3f}")
                        c5.write(f"{for_1:,.3f}")
                        actual = c6.number_input(
                            "qty", min_value=0.0, value=0.0,
                            step=0.001, format="%.3f",
                            key=f"form_mat_{mc}",
                            label_visibility="collapsed")
                        c7.write(f"{onord:,.3f}")
                        mat_inputs[mc] = {
                            "material_name": str(mrow.get("Material_Name", "")),
                            "uom":           str(mrow.get("UOM", "")),
                            "for_1_sqm":     for_1,
                            "actual_input":  float(actual),
                        }

                    ce_notes_form = st.text_area(
                        "📝 Notes (optional)",
                        placeholder="Weather conditions, issues, remarks…",
                        key="form_notes", height=70)

                    # ── Live variance preview (outside form, inside else block) ──
                    _sqm_preview = st.session_state.get("form_ce_sqm", 0.0)
                    if _sqm_preview > 0:
                        for _, _mrow in sc_recipe.iterrows():
                            _mc_p   = str(_mrow["Material_Code"])
                            _for1_p = float(_mrow.get("For_1_SQM", 0) or 0)
                            _exp_p  = round(_for1_p * _sqm_preview, 4)
                            _act_p  = float(st.session_state.get(f"form_mat_{_mc_p}", 0.0))
                            _eff_p  = _act_p if _act_p > 0 else _exp_p
                            if _exp_p > 0:
                                _var_p = (_eff_p - _exp_p) / _exp_p * 100
                                if _var_p > 1.0:
                                    st.warning(f"⚠️ {_mc_p}: Entered {_eff_p:.3f} vs expected {_exp_p:.3f} — **Over Consumption** (+{_var_p:.1f}%)")
                                elif _var_p < -1.0:
                                    st.info(f"ℹ️ {_mc_p}: Entered {_eff_p:.3f} vs expected {_exp_p:.3f} — **Less Consumption** ({_var_p:.1f}%)")

                    add_to_grid_btn = st.form_submit_button(
                        "➕ Add to Grid",
                        use_container_width=False)

                # Clear Form button (outside form)
                if st.button("🧹 Clear Form", key="ce_clr_btn"):
                    for _k in list(st.session_state.keys()):
                        if _k.startswith(("form_ce_", "form_mat_", "form_notes")):
                            del st.session_state[_k]
                    st.rerun()

                # ── Add to Draft Grid ─────────────────────────────────────────
                if add_to_grid_btn:
                    sqm_val   = st.session_state.get("form_ce_sqm",  0.0)
                    date_val  = st.session_state.get("form_ce_date", date.today())
                    notes_val = st.session_state.get("form_notes",   "")
                    if sqm_val <= 0:
                        st.error("❌ Enter SQM Completed > 0 before adding to grid.")
                    elif sqm_val > remaining_sqm + 0.001:
                        st.error(f"❌ SQM entered ({sqm_val:.2f} m²) exceeds remaining ({remaining_sqm:.2f} m²). Entry blocked.")
                    else:
                        try:
                            conn = get_db(); cur = conn.cursor()
                            _sk = st.session_state["_session_key"]
                            for mc, vals in mat_inputs.items():
                                expected_qty = round(vals["for_1_sqm"] * sqm_val, 4)
                                actual_qty   = float(vals["actual_input"])
                                effective_qty = actual_qty if actual_qty > 0 else expected_qty
                                if expected_qty > 0:
                                    var_pct = round((effective_qty - expected_qty) / expected_qty * 100, 2)
                                else:
                                    var_pct = 0.0
                                if abs(var_pct) < 1.0:
                                    var_status = "OK"
                                elif var_pct > 0:
                                    var_status = "Over Consumption"
                                else:
                                    var_status = "Less Consumption"
                                cur.execute("""
                                    INSERT INTO draft_consumption
                                      (session_key, entry_date, equipment_tag, lining_system_code,
                                       lining_system_name, sqm_completed, material_code,
                                       material_name, uom, expected_qty, actual_qty,
                                       effective_qty, variance_pct, variance_status, notes)
                                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                                """, (_sk, str(date_val), ce_tag, ce_code, sname,
                                      sqm_val, mc, vals["material_name"], vals["uom"],
                                      expected_qty, actual_qty, effective_qty,
                                      var_pct, var_status, notes_val))
                            conn.commit(); conn.close()
                            st.success(f"✅ {len(mat_inputs)} material(s) added to draft grid for {ce_tag} · Code {ce_code}.")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Database error: {e}")

        # ── Draft Consumption Grid ────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="sec-hdr">📋 Draft Consumption Grid — Pending Submission</div>',
                    unsafe_allow_html=True)
        _sk = st.session_state["_session_key"]
        _conn_draft = get_db()
        _draft_df = pd.read_sql(
            "SELECT * FROM draft_consumption WHERE session_key = ? ORDER BY added_at",
            _conn_draft, params=[_sk])
        _conn_draft.close()

        if _draft_df.empty:
            st.info("No pending entries. Use 'Add to Grid' above to stage consumption data before submitting.")
        else:
            # Colour-code by variance status
            _draft_show = _draft_df[[
                "id","entry_date","equipment_tag","lining_system_code","material_code",
                "material_name","uom","expected_qty","effective_qty","variance_pct","variance_status","notes"
            ]].copy()
            _draft_show.insert(0, "☐ Del", False)

            def _style_draft(row):
                vs = row.get("variance_status","OK")
                if vs == "Over Consumption":   bg = "rgba(245,158,11,.15)"
                elif vs == "Less Consumption": bg = "rgba(59,130,246,.12)"
                else:                          bg = "rgba(16,185,129,.08)"
                return [f"background-color:{bg}"] * len(row)

            _de_state_key = "draft_ce_editor"
            st.data_editor(
                _draft_show.style.apply(_style_draft, axis=1),
                key=_de_state_key,
                num_rows="fixed",
                hide_index=True,
                use_container_width=True,
                height=min(500, 55 + len(_draft_show) * 35),
                column_config={
                    "id":               st.column_config.NumberColumn("ID", disabled=True),
                    "☐ Del":            st.column_config.CheckboxColumn("🗑", default=False),
                    "variance_status":  st.column_config.SelectboxColumn(
                        "Variance Status",
                        options=["OK","Over Consumption","Less Consumption"]),
                },
            )

            _dg1, _dg2, _dg3 = st.columns([2, 2, 4])
            with _dg1:
                if st.button("🗑️ Delete Selected", key="del_draft_rows"):
                    _de_estate = st.session_state.get(_de_state_key, {})
                    _de_edits  = _de_estate.get("edited_rows", {})
                    _del_ids   = [int(_draft_df.iloc[int(i)]["id"])
                                  for i, ch in _de_edits.items()
                                  if ch.get("☐ Del", False)]
                    if _del_ids:
                        _c = get_db(); _cc = _c.cursor()
                        for _did in _del_ids:
                            _cc.execute("DELETE FROM draft_consumption WHERE id=?", (_did,))
                        _c.commit(); _c.close()
                        st.cache_data.clear(); st.rerun()
                    else:
                        st.warning("Check the 🗑 column on rows you want to delete.")
            with _dg2:
                if st.button("🧹 Clear All Draft", key="clear_draft_all"):
                    _c = get_db()
                    _c.execute("DELETE FROM draft_consumption WHERE session_key=?", (_sk,))
                    _c.commit(); _c.close()
                    st.cache_data.clear(); st.rerun()

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown('<div class="sec-hdr">✅ Submit All Draft Entries</div>',
                        unsafe_allow_html=True)
            if st.button("✅ Submit Consumption", key="submit_draft_btn", type="primary"):
                try:
                    _c = get_db(); _cc = _c.cursor()
                    # Process each draft row
                    for _, _dr in _draft_df.iterrows():
                        _cc.execute("""
                            INSERT INTO consumption_log
                              (entry_date, equipment_tag, lining_system_code,
                               lining_system_name, sqm_completed, material_code,
                               material_name, uom, expected_qty, consumed_qty,
                               variance_status, variance_pct, notes)
                            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                        """, (str(_dr["entry_date"]), str(_dr["equipment_tag"]),
                              str(_dr["lining_system_code"]), str(_dr.get("lining_system_name","")),
                              float(_dr["sqm_completed"]), str(_dr["material_code"]),
                              str(_dr.get("material_name","")), str(_dr.get("uom","")),
                              float(_dr.get("expected_qty") or 0),
                              float(_dr.get("effective_qty") or 0),
                              str(_dr.get("variance_status","OK")),
                              float(_dr.get("variance_pct") or 0),
                              str(_dr.get("notes",""))))
                        eff = float(_dr.get("effective_qty") or 0)
                        if eff > 0:
                            _cc.execute(
                                "UPDATE inventory SET available_qty = MAX(0, available_qty - ?) WHERE material_code = ?",
                                (eff, str(_dr["material_code"])))
                    # Update sqm_progress for each unique (tag, code) pair
                    _sqm_updates = _draft_df.groupby(
                        ["equipment_tag","lining_system_code"], as_index=False
                    )["sqm_completed"].first()
                    for _, _sr in _sqm_updates.iterrows():
                        _cc.execute(
                            "UPDATE sqm_progress SET done_sqm = done_sqm + ? WHERE equipment_tag=? AND lining_system_code=?",
                            (float(_sr["sqm_completed"]), str(_sr["equipment_tag"]), str(_sr["lining_system_code"])))
                    # Delete draft rows
                    _cc.execute("DELETE FROM draft_consumption WHERE session_key=?", (_sk,))
                    _c.commit(); _c.close()
                    st.cache_data.clear()

                    # Generate multi-sheet Excel
                    _sub_sheets = []
                    _full_cols = ["entry_date","equipment_tag","lining_system_code",
                                  "lining_system_name","sqm_completed","material_code",
                                  "material_name","uom","expected_qty","effective_qty",
                                  "variance_pct","variance_status","notes"]
                    _full_export = _draft_df[[c for c in _full_cols if c in _draft_df.columns]].copy()
                    _full_export.columns = [c.replace("_"," ").title() for c in _full_export.columns]
                    _full_export = _full_export.rename(columns={"Variance Status": "Variance Status"})
                    _sub_sheets.append({
                        "name": "Full Day Summary",
                        "df":   _full_export,
                        "title": f"Consumption Summary — {date.today()}",
                        "color_scheme": "overview",
                        "add_grand_total": True,
                    })
                    for _scode in sorted(_draft_df["lining_system_code"].unique(), key=int):
                        _sc_rows = _draft_df[_draft_df["lining_system_code"] == _scode]
                        _sc_sname = str(_sc_rows["lining_system_name"].iloc[0]) if "lining_system_name" in _sc_rows else _scode
                        _sc_export = _sc_rows[[c for c in _full_cols if c in _sc_rows.columns]].copy()
                        _sc_export.columns = [c.replace("_"," ").title() for c in _sc_export.columns]
                        _tab_name = f"Code {_scode}"[:31]
                        _sub_sheets.append({
                            "name": _tab_name,
                            "df":   _sc_export,
                            "title": f"Code {_scode} — {_sc_sname}",
                            "color_scheme": "train_j",
                            "add_grand_total": True,
                        })
                    _excel_bytes = generate_multi_sheet_excel(_sub_sheets)
                    st.success(f"✅ {len(_draft_df)} entries submitted to consumption log.")
                    st.download_button(
                        "⬇ Download Consumption Report",
                        data=_excel_bytes,
                        file_name=f"consumption_report_{date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_cons_report",
                    )
                    st.rerun()
                except Exception as _e:
                    st.error(f"❌ Database error during submission: {_e}")

        # ── Consumption History ───────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("📜 View & Edit Consumption History", expanded=False):
            conn = get_db()
            _log_raw = pd.read_sql("""
                SELECT id,
                       entry_date AS Date,
                       equipment_tag AS Equipment,
                       lining_system_code AS Code,
                       lining_system_name AS System,
                       sqm_completed AS "SQM Done",
                       material_code AS Material,
                       material_name AS "Material Name",
                       uom AS UOM,
                       expected_qty AS "Expected Qty",
                       consumed_qty AS "Consumed Qty",
                       notes AS Notes,
                       submitted_at AS "Submitted At"
                FROM consumption_log
                ORDER BY submitted_at DESC
                LIMIT 200
            """, conn)
            conn.close()

            if _log_raw.empty:
                st.info("No consumption entries yet.")
            else:
                fh1, fh2 = st.columns(2)
                with fh1:
                    log_tags = ["All"] + sorted(_log_raw["Equipment"].unique().tolist())
                    fh_tag   = st.selectbox("Filter by Equipment", log_tags, key="log_tag")
                with fh2:
                    log_dates = ["All"] + sorted(_log_raw["Date"].unique().tolist(), reverse=True)
                    fh_date   = st.selectbox("Filter by Date", log_dates, key="log_date")

                df_show = _log_raw.copy()
                if fh_tag  != "All": df_show = df_show[df_show["Equipment"]==fh_tag]
                if fh_date != "All": df_show = df_show[df_show["Date"]==fh_date]
                df_show = df_show.reset_index(drop=True)

                # Build editable grid
                _clog_display = df_show.copy()
                _clog_display.insert(0, "Sl. No.", range(1, len(_clog_display)+1))
                _clog_display.insert(1, "☐ Select", False)

                st.data_editor(
                    _clog_display,
                    key="cons_log_editor",
                    num_rows="fixed",
                    hide_index=True,
                    use_container_width=True,
                    height=min(600, 50 + len(_clog_display) * 35),
                    column_config={
                        "id":           st.column_config.NumberColumn("ID", disabled=True),
                        "Sl. No.":      st.column_config.NumberColumn("Sl. No.", disabled=True),
                        "☐ Select":     st.column_config.CheckboxColumn("☐", default=False),
                        "Date":         st.column_config.TextColumn("Date", disabled=True),
                        "Equipment":    st.column_config.TextColumn("Equipment", disabled=True),
                        "Code":         st.column_config.TextColumn("Code", disabled=True),
                        "Submitted At": st.column_config.TextColumn("Submitted At", disabled=True),
                    },
                )

                _cb1, _cb2, _cb3 = st.columns([2, 2, 4])
                with _cb1:
                    if st.button("💾 Save Cell Edits", key="cons_log_save"):
                        _estate = st.session_state.get("cons_log_editor", {})
                        _edits  = _estate.get("edited_rows", {})
                        _skip   = {"Sl. No.", "☐ Select", "id", "Date", "Equipment",
                                   "Code", "Submitted At"}
                        _saved  = 0
                        try:
                            conn = get_db(); cur = conn.cursor()
                            for _ridx, _changes in _edits.items():
                                _safe = {k: v for k, v in _changes.items() if k not in _skip}
                                if not _safe:
                                    continue
                                _pk = int(df_show.iloc[int(_ridx)]["id"])
                                _set_sql = ", ".join([f'"{k}" = ?' for k in _safe])
                                # Map display column names back to DB column names
                                _col_map = {
                                    "System": "lining_system_name",
                                    "SQM Done": "sqm_completed",
                                    "Material": "material_code",
                                    "Material Name": "material_name",
                                    "UOM": "uom",
                                    "Expected Qty": "expected_qty",
                                    "Consumed Qty": "consumed_qty",
                                    "Notes": "notes",
                                }
                                _db_safe = {_col_map.get(k, k): v for k, v in _safe.items()}
                                _set_sql = ", ".join([f'"{k}" = ?' for k in _db_safe])
                                cur.execute(
                                    f"UPDATE consumption_log SET {_set_sql} WHERE id = ?",
                                    list(_db_safe.values()) + [_pk])
                                _saved += 1
                            conn.commit(); conn.close()
                            st.cache_data.clear()
                            st.success(f"✅ {_saved} row(s) updated.")
                            st.rerun()
                        except Exception as _e:
                            st.error(f"❌ Database error: {_e}")

                with _cb2:
                    if st.button("🗑️ Delete Selected", key="cons_log_del"):
                        _estate   = st.session_state.get("cons_log_editor", {})
                        _edits    = _estate.get("edited_rows", {})
                        _del_idxs = [int(i) for i, ch in _edits.items()
                                     if ch.get("☐ Select", False)]
                        if not _del_idxs:
                            st.warning("Check the ☐ column on rows you want to delete first.")
                        else:
                            try:
                                conn = get_db(); cur = conn.cursor()
                                for _di in _del_idxs:
                                    _pk = int(df_show.iloc[_di]["id"])
                                    cur.execute("DELETE FROM consumption_log WHERE id = ?", (_pk,))
                                conn.commit(); conn.close()
                                st.cache_data.clear()
                                st.success(f"✅ {len(_del_idxs)} row(s) deleted.")
                                st.rerun()
                            except Exception as _e:
                                st.error(f"❌ Database error: {_e}")

                st.download_button(
                    "⬇ Download Consumption Log",
                    data=generate_excel_report(
                        df_show.drop(columns=["id"], errors="ignore"),
                        "Daily Consumption Log", color_scheme="overview"),
                    file_name=f"consumption_log_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=False)

    # ═══════════════════════════════════════════════════════════════════════════
    # RECEIPTS MODE
    # ═══════════════════════════════════════════════════════════════════════════
    elif inv_mode == "📦 Receipts":
        st.markdown('<div class="sec-hdr">📦 Record Material Receipt — New Stock Received</div>',
                    unsafe_allow_html=True)
        st.caption("Log when new materials arrive on site. "
                   "Available Qty increases by the received amount; Ordered Qty decreases accordingly.")

        # ── Step 1: Order ID selector (appears first; material list derived from it) ─
        _rc_open_conn = get_db()
        _rc_all_orders = pd.read_sql(
            "SELECT DISTINCT order_id FROM orders_log "
            "WHERE status != 'Fulfilled' AND (ordered_qty - fulfilled_qty) > 0 "
            "ORDER BY id DESC",
            _rc_open_conn)
        _rc_open_conn.close()
        _rc_order_opts = ["— None —"] + _rc_all_orders["order_id"].tolist()
        _rc_sel_order  = st.selectbox(
            "🔗 Link to Order ID / PR# (Optional — select first to filter materials)",
            options=_rc_order_opts, key="rc_order_id",
            help="Select an order to restrict the Material dropdown to only that order's items. "
                 "Choose '— None —' to see all available materials.")

        # ── Step 2: Material selector — filtered by selected order (or all) ──────
        if _rc_sel_order != "— None —":
            _rc_ord_conn = get_db()
            _rc_ord_mats = pd.read_sql(
                "SELECT DISTINCT material_code FROM orders_log WHERE order_id = ?",
                _rc_ord_conn, params=[_rc_sel_order])["material_code"].tolist()
            _rc_ord_conn.close()
            mat_opts_r = [m for m in inv["Material_Code"].tolist() if m in _rc_ord_mats]
        else:
            mat_opts_r = inv["Material_Code"].tolist()

        rc_mat = st.selectbox(
            "🧪 Select Material",
            options=mat_opts_r,
            format_func=lambda m: f"{m}  —  "
                f"{inv.set_index('Material_Code')['Material_Name'].get(m, m)[:35]}",
            key="rc_mat_sel")

        # Auto-fill disabled read-only boxes
        if rc_mat:
            _rc_row   = inv[inv["Material_Code"] == rc_mat].iloc[0]
            _rc_avail = float(_rc_row.get("Available_Qty", 0) or 0)
            _rc_ord   = float(_rc_row.get("Ordered_Qty",  0) or 0)
            _ri1, _ri2 = st.columns(2)
            _ri1.text_input("Current Available Qty", value=f"{_rc_avail:,.3f}",
                            disabled=True, key="rc_avail_disp")
            _ri2.text_input("Current Ordered Qty",  value=f"{_rc_ord:,.3f}",
                            disabled=True, key="rc_ord_disp")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Receipt form ──────────────────────────────────────────────────────
        with st.form(key="receipt_form", clear_on_submit=False):
            rc1, rc2, rc3 = st.columns(3)
            with rc1:
                rc_date = st.date_input("📅 Receipt Date", value=date.today(), key="rc_date")
            with rc2:
                rc_qty = st.number_input("Qty Received", min_value=0.0,
                                         value=0.0, step=1.0, format="%.3f", key="rc_qty")
            with rc3:
                rc_notes = st.text_input("Notes / PO Ref.", key="rc_notes",
                                         placeholder="PO number, supplier…")

            # Dynamic extra columns from receipt_log schema
            _RECEIPT_FIXED = {"id","entry_date","material_code","material_name","uom",
                               "received_qty","notes","submitted_at","order_id"}
            _rc_conn = get_db()
            _rc_extra_cols = [(r[1], r[2]) for r in
                              _rc_conn.execute("PRAGMA table_info(receipt_log)").fetchall()
                              if r[1].lower() not in _RECEIPT_FIXED]
            _rc_conn.close()
            rc_extra_inputs = {}
            if _rc_extra_cols:
                st.markdown('<div class="sec-hdr" style="margin-top:.5rem;">'
                            'Additional Fields</div>', unsafe_allow_html=True)
                for _ri in range(0, len(_rc_extra_cols), 3):
                    _rcols = st.columns(3)
                    for _rj, (_rcn, _rct) in enumerate(_rc_extra_cols[_ri:_ri+3]):
                        with _rcols[_rj]:
                            if any(kw in _rcn.lower() for kw in ("qty","sqm","amount","value")):
                                rc_extra_inputs[_rcn] = st.number_input(
                                    _rcn.replace("_"," ").title(),
                                    value=0.0, step=0.001, format="%.3f",
                                    key=f"rc_ext_{_rcn}")
                            else:
                                rc_extra_inputs[_rcn] = st.text_input(
                                    _rcn.replace("_"," ").title(),
                                    key=f"rc_ext_{_rcn}")

            rc_submit = st.form_submit_button("✅  Record Receipt", use_container_width=False)

        # Clear Form button (outside form)
        if st.button("🧹 Clear Form", key="rc_clr_btn"):
            for _k in list(st.session_state.keys()):
                if _k.startswith(("rc_date", "rc_qty", "rc_notes", "rc_ext_",
                                   "rc_avail_disp", "rc_ord_disp", "rc_mat_sel")):
                    del st.session_state[_k]
            st.rerun()

        # ── DB write ──────────────────────────────────────────────────────────
        if rc_submit:
            rc_qty_val    = st.session_state.get("rc_qty",      0.0)
            rc_date_val   = st.session_state.get("rc_date",     date.today())
            rc_notes_val  = st.session_state.get("rc_notes",    "")
            rc_order_link = st.session_state.get("rc_order_id", "— None —")
            if not rc_mat:
                st.error("Please select a material.")
            elif rc_qty_val <= 0:
                st.error("❌ Please enter a quantity > 0.")
            else:
                try:
                    conn = get_db(); cur = conn.cursor()
                    _linked_order = rc_order_link if rc_order_link != "— None —" else None
                    cur.execute("""
                        INSERT INTO receipt_log
                          (entry_date, material_code, material_name,
                           uom, received_qty, notes, order_id)
                        SELECT ?, ?, material_name, uom, ?, ?, ?
                        FROM inventory WHERE material_code = ?
                    """, (str(rc_date_val), rc_mat, rc_qty_val, rc_notes_val,
                          _linked_order, rc_mat))
                    _new_receipt_id = cur.lastrowid
                    # Available_Qty += received; Ordered_Qty -= received (floor 0)
                    cur.execute("""
                        UPDATE inventory
                        SET available_qty = available_qty + ?,
                            ordered_qty   = MAX(0, ordered_qty - ?)
                        WHERE material_code = ?
                    """, (rc_qty_val, rc_qty_val, rc_mat))
                    # Link to order if provided
                    if _linked_order:
                        cur.execute("""
                            UPDATE orders_log
                            SET fulfilled_qty = fulfilled_qty + ?,
                                receipt_ids = CASE WHEN receipt_ids IS NULL
                                    THEN CAST(? AS TEXT)
                                    ELSE receipt_ids || ',' || CAST(? AS TEXT) END,
                                status = CASE
                                    WHEN fulfilled_qty + ? >= ordered_qty THEN 'Fulfilled'
                                    WHEN fulfilled_qty + ? > 0            THEN 'Partial'
                                    ELSE 'Pending' END
                            WHERE order_id = ? AND material_code = ?
                        """, (rc_qty_val, _new_receipt_id, _new_receipt_id,
                              rc_qty_val, rc_qty_val, _linked_order, rc_mat))
                    conn.commit(); conn.close()
                    st.cache_data.clear()
                    _mat_nm = inv.set_index("Material_Code")["Material_Name"].get(rc_mat, rc_mat)
                    _order_msg = f" Linked to Order **{_linked_order}**." if _linked_order else ""
                    st.success(f"✅ Received {rc_qty_val:,.3f} units of {_mat_nm} "
                               f"added to inventory. Ordered Qty adjusted.{_order_msg}")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error: {e}")

        # ── Recent Receipts ───────────────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("📜 View & Edit Receipt History", expanded=False):
            conn = get_db()
            _rlog_raw = pd.read_sql(
                "SELECT id, entry_date AS Date, material_code AS Code, "
                "material_name AS Material, uom AS UOM, "
                "received_qty AS \"Received Qty\", notes AS Notes, "
                "submitted_at AS \"Received At\" "
                "FROM receipt_log ORDER BY submitted_at DESC LIMIT 100", conn)
            conn.close()
            if _rlog_raw.empty:
                st.info("No receipt entries yet.")
            else:
                _rf1, _rf2 = st.columns(2)
                with _rf1:
                    _r_mats  = ["All"] + sorted(_rlog_raw["Code"].unique().tolist())
                    _r_fmat  = st.selectbox("Filter by Material", _r_mats, key="rlog_mat")
                with _rf2:
                    _r_dates = ["All"] + sorted(_rlog_raw["Date"].unique().tolist(), reverse=True)
                    _r_fdate = st.selectbox("Filter by Date", _r_dates, key="rlog_date")

                rlog = _rlog_raw.copy()
                if _r_fmat  != "All": rlog = rlog[rlog["Code"]==_r_fmat]
                if _r_fdate != "All": rlog = rlog[rlog["Date"]==_r_fdate]
                rlog = rlog.reset_index(drop=True)

                _rlog_display = rlog.copy()
                _rlog_display.insert(0, "Sl. No.", range(1, len(_rlog_display)+1))
                _rlog_display.insert(1, "☐ Select", False)

                st.data_editor(
                    _rlog_display,
                    key="receipt_log_editor",
                    num_rows="fixed",
                    hide_index=True,
                    use_container_width=True,
                    height=min(500, 50 + len(_rlog_display) * 35),
                    column_config={
                        "id":          st.column_config.NumberColumn("ID", disabled=True),
                        "Sl. No.":     st.column_config.NumberColumn("Sl. No.", disabled=True),
                        "☐ Select":    st.column_config.CheckboxColumn("☐", default=False),
                        "Code":        st.column_config.TextColumn("Code", disabled=True),
                        "Received At": st.column_config.TextColumn("Received At", disabled=True),
                    },
                )

                _rb1, _rb2, _rb3 = st.columns([2, 2, 4])
                with _rb1:
                    if st.button("💾 Save Cell Edits", key="rlog_save"):
                        _restate = st.session_state.get("receipt_log_editor", {})
                        _redits  = _restate.get("edited_rows", {})
                        _rskip   = {"Sl. No.", "☐ Select", "id", "Code", "Received At"}
                        _rsaved  = 0
                        try:
                            conn = get_db(); cur = conn.cursor()
                            for _ridx, _changes in _redits.items():
                                _safe = {k: v for k, v in _changes.items() if k not in _rskip}
                                if not _safe:
                                    continue
                                _pk = int(rlog.iloc[int(_ridx)]["id"])
                                _col_map = {
                                    "Date":         "entry_date",
                                    "Material":     "material_name",
                                    "UOM":          "uom",
                                    "Received Qty": "received_qty",
                                    "Notes":        "notes",
                                }
                                _db_safe = {_col_map.get(k, k): v for k, v in _safe.items()}
                                _set_sql = ", ".join([f'"{k}" = ?' for k in _db_safe])
                                cur.execute(
                                    f"UPDATE receipt_log SET {_set_sql} WHERE id = ?",
                                    list(_db_safe.values()) + [_pk])
                                _rsaved += 1
                            conn.commit(); conn.close()
                            st.cache_data.clear()
                            st.success(f"✅ {_rsaved} row(s) updated.")
                            st.rerun()
                        except Exception as _e:
                            st.error(f"❌ Database error: {_e}")

                with _rb2:
                    if st.button("🗑️ Delete Selected", key="rlog_del"):
                        _restate  = st.session_state.get("receipt_log_editor", {})
                        _redits   = _restate.get("edited_rows", {})
                        _del_idxs = [int(i) for i, ch in _redits.items()
                                     if ch.get("☐ Select", False)]
                        if not _del_idxs:
                            st.warning("Check the ☐ column on rows you want to delete first.")
                        else:
                            try:
                                conn = get_db(); cur = conn.cursor()
                                for _di in _del_idxs:
                                    _pk = int(rlog.iloc[_di]["id"])
                                    cur.execute("DELETE FROM receipt_log WHERE id = ?", (_pk,))
                                conn.commit(); conn.close()
                                st.cache_data.clear()
                                st.success(f"✅ {len(_del_idxs)} row(s) deleted.")
                                st.rerun()
                            except Exception as _e:
                                st.error(f"❌ Database error: {_e}")

                st.download_button(
                    "⬇ Download Receipt Log",
                    data=generate_excel_report(
                        rlog.drop(columns=["id"], errors="ignore"),
                        "Material Receipt Log", color_scheme="overview"),
                    file_name=f"receipt_log_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_receipt_log")

    # ═══════════════════════════════════════════════════════════════════════════
    # MAIN INVENTORY DASHBOARD
    # ═══════════════════════════════════════════════════════════════════════════
    elif inv_mode == "📊 Main Inventory":
        from datetime import timedelta
        st.markdown('<div class="sec-hdr">📊 Main Inventory Dashboard — Movements by Date Range</div>',
                    unsafe_allow_html=True)

        _mi_c1, _mi_c2, _mi_c3 = st.columns(3)
        with _mi_c1:
            _mi_from = st.date_input("From Date", value=date.today() - timedelta(days=30), key="mi_from")
        with _mi_c2:
            _mi_to   = st.date_input("To Date",   value=date.today(), key="mi_to")
        with _mi_c3:
            _mi_view = st.radio("Group By", ["By Material", "By System Code"],
                                horizontal=True, key="mi_view")

        st.markdown("<hr>", unsafe_allow_html=True)

        _mi_conn = get_db()
        # Receipts in range
        _mi_recv = pd.read_sql(
            "SELECT material_code, SUM(received_qty) AS total_received "
            "FROM receipt_log WHERE entry_date BETWEEN ? AND ? GROUP BY material_code",
            _mi_conn, params=[str(_mi_from), str(_mi_to)])
        # Consumed in range
        _mi_cons = pd.read_sql(
            "SELECT material_code, SUM(consumed_qty) AS total_consumed "
            "FROM consumption_log WHERE entry_date BETWEEN ? AND ? GROUP BY material_code",
            _mi_conn, params=[str(_mi_from), str(_mi_to)])
        _mi_conn.close()

        # Build main inventory table
        _mi_base = inv[["Material_Code","Material_Name","UOM","Available_Qty","Ordered_Qty"]].copy()
        _mi_base = _mi_base.rename(columns={
            "Material_Code": "material_code", "Material_Name": "material_name",
            "UOM": "uom", "Available_Qty": "current_stock", "Ordered_Qty": "ordered_qty"})
        _mi_base = _mi_base.merge(_mi_recv, on="material_code", how="left")
        _mi_base = _mi_base.merge(_mi_cons, on="material_code", how="left")
        _mi_base["total_received"] = _mi_base["total_received"].fillna(0)
        _mi_base["total_consumed"] = _mi_base["total_consumed"].fillna(0)

        if _mi_view == "By Material":
            _mi_show = _mi_base.rename(columns={
                "material_code": "Code", "material_name": "Material Name", "uom": "UOM",
                "total_received": "Total Receipts", "total_consumed": "Total Consumed",
                "current_stock": "Current Stock", "ordered_qty": "Ordered Qty"
            })[["Code","Material Name","UOM","Total Receipts","Total Consumed","Current Stock","Ordered Qty"]]

            def _style_mi(row):
                stk = row["Current Stock"]
                if stk <= 0:    bg,tc = "rgba(239,68,68,.12)","#EF4444"
                elif stk < 50:  bg,tc = "rgba(245,158,11,.12)","#F59E0B"
                else:           bg,tc = "rgba(16,185,129,.08)","#10B981"
                styles = [f"background-color:{bg}"] * len(row)
                styles[list(row.index).index("Current Stock")] = f"background-color:{bg};color:{tc};font-weight:700"
                return styles

            st.dataframe(
                _mi_show.style.apply(_style_mi, axis=1).format({
                    "Total Receipts":"{:,.3f}","Total Consumed":"{:,.3f}",
                    "Current Stock":"{:,.3f}","Ordered Qty":"{:,.3f}"}),
                use_container_width=True, hide_index=True,
                height=60 + len(_mi_show)*35, key="mi_mat_tbl")

            st.download_button("⬇ Download Inventory Dashboard",
                data=generate_excel_report(_mi_show.reset_index(drop=True),
                    f"Inventory Dashboard {_mi_from} to {_mi_to}", color_scheme="overview"),
                file_name=f"inventory_dashboard_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_mi_mat")

        else:  # By System Code
            _mi_recipe_codes = recipe[["Lining_System_Code","Lining_System_Short_Name","Material_Code"]].copy()
            _mi_recipe_codes = _mi_recipe_codes.rename(columns={"Material_Code":"material_code"})
            _mi_merged = _mi_recipe_codes.merge(_mi_base, on="material_code", how="left")
            _mi_merged["total_received"] = _mi_merged["total_received"].fillna(0)
            _mi_merged["total_consumed"] = _mi_merged["total_consumed"].fillna(0)
            _mi_merged["current_stock"]  = _mi_merged["current_stock"].fillna(0)

            for _micode in sorted(_mi_merged["Lining_System_Code"].unique(), key=int):
                _mi_sc = _mi_merged[_mi_merged["Lining_System_Code"]==_micode]
                _mi_sn = _mi_sc["Lining_System_Short_Name"].iloc[0]
                _mi_cov = (_mi_sc["current_stock"].sum() / max(_mi_sc["current_stock"].sum() + _mi_sc["total_consumed"].sum(), 1)) * 100
                with st.expander(f"Code {_micode} — {_mi_sn}  ·  {len(_mi_sc)} materials  ·  {_mi_cov:.0f}% stock coverage",
                                 expanded=False):
                    _mi_sc_show = _mi_sc[["material_code","material_name","uom",
                                          "total_received","total_consumed","current_stock","ordered_qty"]].copy()
                    _mi_sc_show.columns = ["Code","Material Name","UOM",
                                           "Total Receipts","Total Consumed","Current Stock","Ordered Qty"]
                    st.dataframe(_mi_sc_show.style.format({
                        "Total Receipts":"{:,.3f}","Total Consumed":"{:,.3f}",
                        "Current Stock":"{:,.3f}","Ordered Qty":"{:,.3f}"}),
                        use_container_width=True, hide_index=True,
                        height=60 + len(_mi_sc_show)*35, key=f"mi_sc_{_micode}")

    # ═══════════════════════════════════════════════════════════════════════════
    # ORDERED — ORDER MANAGEMENT SYSTEM
    # ═══════════════════════════════════════════════════════════════════════════
    elif inv_mode == "📋 Ordered":
        st.markdown('<div class="sec-hdr">📋 Order Management System</div>',
                    unsafe_allow_html=True)

        _ord_mode = st.radio(
            "Generate From",
            ["📍 Location Needs", "⚙️ System Code Needs", "📜 View Past Orders"],
            horizontal=True, key="ord_mode",
        )
        st.markdown("<hr>", unsafe_allow_html=True)

        # ── Shared: Generate Shortfall DataFrame ─────────────────────────────
        def _compute_order_shortfalls(tags: list, label: str) -> pd.DataFrame:
            if not tags:
                return pd.DataFrame()
            _alloc = cascade_allocate(tags)
            _short = (_alloc[_alloc["Shortfall_Qty"] > 0]
                      .groupby(["Material_Code","Material_Name","UOM"], as_index=False)
                      ["Shortfall_Qty"].sum())
            _short = _short.merge(inv[["Material_Code","Available_Qty"]],
                                  on="Material_Code", how="left")
            _short["Available_Qty"]   = _short["Available_Qty"].fillna(0)
            _short["Order Qty"]       = _short["Shortfall_Qty"].round(3)
            _short["Notes"]           = ""
            _short["☑ Remove"]        = False
            _short["Source"]          = label
            return _short.rename(columns={
                "Material_Code":"Material Code","Material_Name":"Material Name",
                "UOM":"UOM","Available_Qty":"Current Stock","Shortfall_Qty":"Shortfall Qty"})

        # ── Location Needs ────────────────────────────────────────────────────
        if _ord_mode == "📍 Location Needs":
            _ord_loc_sel = st.selectbox("Select Location",
                options=["All Locations"] + LOCATION_ORDER, key="ord_loc_sel")
            if st.button("🔍 Calculate Shortfalls", key="calc_ord_short"):
                if _ord_loc_sel == "All Locations":
                    _ord_tags = eq_master["Equipment_Tag_No."].tolist()
                else:
                    _ord_tags = eq_master[eq_master["Location"]==_ord_loc_sel]["Equipment_Tag_No."].tolist()
                _ord_sf = _compute_order_shortfalls(_ord_tags, _ord_loc_sel)
                st.session_state["_ord_draft"] = _ord_sf.to_dict("records")
                st.session_state["_ord_source_detail"] = _ord_loc_sel

        # ── System Code Needs ─────────────────────────────────────────────────
        elif _ord_mode == "⚙️ System Code Needs":
            _ord_code_opts = sorted(dm["Lining_System_Code"].unique().tolist(), key=int)
            _ord_code_sel  = st.selectbox("Select System Code", _ord_code_opts,
                format_func=lambda c: f"Code {c} — {dm[dm['Lining_System_Code']==c]['Lining_System_Short_Name'].iloc[0]}",
                key="ord_code_sel")
            if st.button("🔍 Calculate Shortfalls", key="calc_ord_sc_short"):
                _ord_tags_sc = dm[dm["Lining_System_Code"]==_ord_code_sel]["Equipment_Tag_No."].unique().tolist()
                _ord_sf_sc = _compute_order_shortfalls(_ord_tags_sc, f"Code {_ord_code_sel}")
                st.session_state["_ord_draft"] = _ord_sf_sc.to_dict("records")
                st.session_state["_ord_source_detail"] = f"Code {_ord_code_sel}"

        # ── View Past Orders ──────────────────────────────────────────────────
        elif _ord_mode == "📜 View Past Orders":
            _past_conn = get_db()
            _past_ids  = pd.read_sql(
                "SELECT DISTINCT order_id, order_date, generation_source, source_detail "
                "FROM orders_log ORDER BY id DESC", _past_conn)
            _past_conn.close()
            if _past_ids.empty:
                st.info("No orders submitted yet.")
            else:
                _sel_ord = st.selectbox("Select Order ID",
                    _past_ids["order_id"].tolist(), key="view_past_ord")
                _past_ord_conn = get_db()
                _ord_detail = pd.read_sql(
                    "SELECT order_id AS 'Order ID', pr_number AS 'PR#', "
                    "material_code AS 'Material Code', material_name AS 'Material Name', "
                    "uom AS UOM, ordered_qty AS 'Ordered Qty', fulfilled_qty AS 'Fulfilled Qty', "
                    "MAX(0, ordered_qty - fulfilled_qty) AS 'Pending Qty', status AS Status, notes AS Notes "
                    "FROM orders_log WHERE order_id = ?", _past_ord_conn, params=[_sel_ord])
                _past_ord_conn.close()
                if not _ord_detail.empty:
                    # PR# entry (current value from first row, all rows share same pr_number)
                    _cur_pr_val = str(_ord_detail["PR#"].iloc[0] or "") if "PR#" in _ord_detail.columns else ""
                    _pr_c1, _pr_c2 = st.columns([3, 1])
                    with _pr_c1:
                        _pr_input = st.text_input(
                            "🔖 PR# (Purchase Request Number)",
                            value=_cur_pr_val,
                            placeholder="Enter official PR number…",
                            key="pr_input")
                    with _pr_c2:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("💾 Update PR#", key="update_pr_btn"):
                            if _pr_input.strip():
                                _upr_conn = get_db()
                                _upr_conn.execute(
                                    "UPDATE orders_log SET pr_number = ? WHERE order_id = ?",
                                    (_pr_input.strip(), _sel_ord))
                                _upr_conn.commit(); _upr_conn.close()
                                st.cache_data.clear()
                                st.success(f"✅ PR# '{_pr_input.strip()}' saved for {_sel_ord}.")
                                st.rerun()
                            else:
                                st.error("Please enter a PR# value.")
                    def _style_ord_status(row):
                        s = row["Status"]
                        if s == "Fulfilled":  bg,tc = "rgba(16,185,129,.1)","#10B981"
                        elif s == "Partial":  bg,tc = "rgba(245,158,11,.1)","#F59E0B"
                        else:                 bg,tc = "rgba(239,68,68,.1)","#EF4444"
                        styles = [f"background-color:{bg}"] * len(row)
                        styles[-3] = f"background-color:{bg};color:{tc};font-weight:700"
                        return styles
                    st.dataframe(_ord_detail.style.apply(_style_ord_status, axis=1).format({
                        "Ordered Qty":"{:,.3f}","Fulfilled Qty":"{:,.3f}","Pending Qty":"{:,.3f}"}),
                        use_container_width=True, hide_index=True, height=60+len(_ord_detail)*35,
                        key="ord_status_tbl")
                    st.download_button("⬇ Download Order Status",
                        data=generate_excel_report(_ord_detail, f"Order Status — {_sel_ord}", color_scheme="execution"),
                        file_name=f"order_status_{_sel_ord}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_ord_status")

        # ── Editable Order Draft Grid (shared by Location & System Code modes) ─
        if _ord_mode in ("📍 Location Needs", "⚙️ System Code Needs"):
            _ord_draft_records = st.session_state.get("_ord_draft", [])
            if _ord_draft_records:
                st.markdown('<div class="sec-hdr" style="margin-top:1rem;">✏️ Review & Edit Order List</div>',
                            unsafe_allow_html=True)
                st.caption("Edit Order Qty, uncheck ☑ Remove to exclude rows, or add new rows manually.")
                _ord_draft_df = pd.DataFrame(_ord_draft_records)

                _ord_editor_key = "ord_draft_editor"
                _ord_edit_df = st.data_editor(
                    _ord_draft_df,
                    key=_ord_editor_key,
                    num_rows="dynamic",
                    hide_index=True,
                    use_container_width=True,
                    height=min(500, 60 + len(_ord_draft_df) * 35),
                    column_config={
                        "Material Code":  st.column_config.TextColumn("Material Code"),
                        "Material Name":  st.column_config.TextColumn("Material Name"),
                        "UOM":            st.column_config.TextColumn("UOM"),
                        "Current Stock":  st.column_config.NumberColumn("Current Stock", disabled=True, format="%.3f"),
                        "Shortfall Qty":  st.column_config.NumberColumn("Shortfall Qty", disabled=True, format="%.3f"),
                        "Order Qty":      st.column_config.NumberColumn("Order Qty", format="%.3f", step=0.001),
                        "Notes":          st.column_config.TextColumn("Notes"),
                        "☑ Remove":       st.column_config.CheckboxColumn("☑ Remove", default=False),
                        "Source":         st.column_config.TextColumn("Source", disabled=True),
                    },
                )

                if st.button("📤 Submit Order", key="submit_ord_btn", type="primary"):
                    # Read final editor state
                    _oe_state   = st.session_state.get(_ord_editor_key, {})
                    _oe_edited  = _oe_state.get("edited_rows", {})
                    _oe_added   = _oe_state.get("added_rows", [])
                    _oe_deleted = set(_oe_state.get("deleted_rows", []))

                    # Build final list from original + edits, excluding removed rows
                    _final_rows = []
                    for _ri, _row in enumerate(list(_ord_draft_df.itertuples(index=False, name=None))):
                        if _ri in _oe_deleted:
                            continue
                        _row_dict = dict(zip(_ord_draft_df.columns, _row))
                        if _ri in _oe_edited:
                            _row_dict.update(_oe_edited[_ri])
                        if _row_dict.get("☑ Remove", False):
                            continue
                        if float(_row_dict.get("Order Qty", 0) or 0) > 0:
                            _final_rows.append(_row_dict)
                    for _ar in _oe_added:
                        if float(_ar.get("Order Qty", 0) or 0) > 0:
                            _final_rows.append(_ar)

                    if not _final_rows:
                        st.error("❌ No valid rows with Order Qty > 0. Nothing to submit.")
                    else:
                        try:
                            _oc = get_db()
                            _oid = _next_order_id(_oc)
                            _cur = _oc.cursor()
                            _src_detail = st.session_state.get("_ord_source_detail", "")
                            _gen_src = "Location" if _ord_mode == "📍 Location Needs" else "System Code"
                            for _fr in _final_rows:
                                _cur.execute("""
                                    INSERT INTO orders_log
                                      (order_id, order_date, generated_by,
                                       generation_source, source_detail,
                                       material_code, material_name, uom,
                                       ordered_qty, fulfilled_qty, status, notes)
                                    VALUES (?,?,?,?,?,?,?,?,?,0,'Pending',?)
                                """, (_oid, str(date.today()),
                                      "Smart Material Estimator & Planner",
                                      _gen_src, _src_detail,
                                      str(_fr.get("Material Code","")),
                                      str(_fr.get("Material Name","")),
                                      str(_fr.get("UOM","")),
                                      float(_fr.get("Order Qty", 0) or 0),
                                      str(_fr.get("Notes",""))))
                            _oc.commit(); _oc.close()
                            st.cache_data.clear()
                            st.success(f"✅ Order **{_oid}** submitted — {len(_final_rows)} material(s).")

                            # Generate Excel order sheet
                            _ord_export_cols = ["Material Code","Material Name","UOM","Order Qty","Notes"]
                            _ord_export = pd.DataFrame([{c: _r.get(c,"") for c in _ord_export_cols}
                                                        for _r in _final_rows])
                            _ord_export.insert(0, "Order ID", _oid)
                            _ord_export.insert(1, "Order Date", str(date.today()))
                            _ord_export.insert(2, "Generated By", "Smart Material Estimator & Planner")
                            st.download_button(
                                f"⬇ Download Order {_oid}",
                                data=generate_excel_report(_ord_export,
                                    f"Procurement Order — {_oid}", color_scheme="execution"),
                                file_name=f"order_{_oid}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_new_order",
                            )
                            del st.session_state["_ord_draft"]
                            st.rerun()
                        except Exception as _oe:
                            st.error(f"❌ Database error: {_oe}")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 5 · TOTAL OVERVIEW  (Master Filterable Datatable)
# ═══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="sec-hdr">📈 Total Overview — Master Equipment & Material Table</div>',
                unsafe_allow_html=True)

    # ── Build master table: one row per (Equipment, System Code) ─────────────
    # Join eq_master + sqm_ref + demand aggregation
    sqm_ref_ov = equip_sc[["Equipment_Tag_No.","Lining_System_Code",
                            "Lining_System_Short_Name","Total_SQM_Original",
                            "done_sqm","Total_SQM"]].drop_duplicates()

    # Total demand per (tag, code) from recipe × SQM
    dm_agg = dm.groupby(["Equipment_Tag_No.","Lining_System_Code"],
                         as_index=False).agg(
        Total_Demand_Qty=("Demand_Qty","sum"))
    dm_agg = dm_agg.merge(
        inv[["Material_Code","Available_Qty"]].groupby("Material_Code",as_index=False).first(),
        how="cross")  # we need per-code shortfall

    # Simpler: shortfall per (tag,code) = demand − min(demand, available)
    # Use cascade_allocate with all tags in file order for true shortfall
    _all_tags_ov = eq_master["Equipment_Tag_No."].tolist()
    _alloc_ov    = cascade_allocate(_all_tags_ov)

    sc_shortfall = _alloc_ov.groupby(
        ["Equipment_Tag_No.","Lining_System_Code"], as_index=False
    ).agg(
        Shortfall_Qty=("Shortfall_Qty","sum"),
        Demand_Qty   =("Demand_Qty","sum"),
        Allocated_Qty=("Allocated_Qty","sum"),
    )

    # Master table
    master = sqm_ref_ov.rename(columns={
        "Total_SQM_Original":"Total_SQM",
        "done_sqm":          "Done_SQM",
        "Total_SQM":         "Remaining_SQM",
    })
    # Merge equipment master (with all columns from Data Input sheet)
    master = master.merge(
        eq_master[["Equipment_Tag_No.","Name","Substrate","Location","Type",
                   "Lining_Systems","Lining_Type","Material_Spec","Design"]],
        on="Equipment_Tag_No.", how="left")

    # Merge equipment_sc for Lining_Area (surface area per row, not summed)
    lining_area_ref = equip_sc[
        ["Equipment_Tag_No.","Lining_System_Code","Total_SQM_Original"]
    ].drop_duplicates().rename(columns={"Total_SQM_Original":"Lining_Area_SQM"})
    lining_area_ref["Lining_Area_SQM"] = lining_area_ref["Lining_Area_SQM"].round(3)
    master = master.merge(lining_area_ref,
                          on=["Equipment_Tag_No.","Lining_System_Code"], how="left")

    master = master.merge(sc_shortfall[["Equipment_Tag_No.","Lining_System_Code",
                                         "Shortfall_Qty","Demand_Qty","Allocated_Qty"]],
                          on=["Equipment_Tag_No.","Lining_System_Code"], how="left")
    master["Shortfall_Qty"] = master["Shortfall_Qty"].fillna(0)
    master["Fulfillment_%"] = (master["Allocated_Qty"] /
        master["Demand_Qty"].replace(0,np.nan) * 100).fillna(100).clip(0,100).round(1)

    # Add serial number
    master = master.reset_index(drop=True)
    master.insert(0, "S.No", master.index + 1)

    display_master = master[[
        "S.No","Equipment_Tag_No.","Name","Substrate","Type","Location",
        "Lining_Systems","Lining_System_Code","Lining_System_Short_Name",
        "Lining_Type","Material_Spec","Design",
        "Total_SQM","Lining_Area_SQM","Done_SQM","Remaining_SQM",
        "Demand_Qty","Allocated_Qty","Shortfall_Qty","Fulfillment_%"
    ]].copy()
    display_master.columns = [
        "S.No","Equipment No","Name","Substrate","Type","Location",
        "Lining System+","System Code","System Name",
        "Lining Type","Material Spec.","Design",
        "Total SQM","Lining Area SQM","Already Done SQM","Remaining SQM",
        "Total Demand","Allocated","Shortfall Qty","Fulfil %"
    ]

    # ── Filter controls ───────────────────────────────────────────────────────
    st.markdown('<div class="sec-hdr">🎛 Filters</div>', unsafe_allow_html=True)
    ff1,ff2,ff3,ff4 = st.columns(4)
    with ff1:
        f_loc_ov = st.multiselect(" Location",
            options=LOCATION_ORDER, default=LOCATION_ORDER, key="ov_loc")
    with ff2:
        type_opts_ov = sorted(display_master["Type"].dropna().unique().tolist())
        f_type_ov = st.multiselect(" Type",
            options=type_opts_ov, default=type_opts_ov, key="ov_type")
    with ff3:
        codes_ov = sorted(display_master["System Code"].unique().tolist(), key=int)
        f_code_ov = st.multiselect(" System Code",
            options=codes_ov, default=codes_ov, key="ov_code")
    with ff4:
        status_opts = ["All","Fully Ready (100%)","Partial (50-99%)","Blocked (<50%)"]
        f_status_ov = st.selectbox(" Status", options=status_opts, key="ov_status")

    # Apply filters
    filtered_master = display_master[
        display_master["Location"].isin(f_loc_ov) &
        display_master["Type"].isin(f_type_ov) &
        display_master["System Code"].isin(f_code_ov)
    ].copy()
    if f_status_ov == "Fully Ready (100%)":
        filtered_master = filtered_master[filtered_master["Fulfil %"] >= 100]
    elif f_status_ov == "Partial (50-99%)":
        filtered_master = filtered_master[(filtered_master["Fulfil %"] >= 50) &
                                          (filtered_master["Fulfil %"] < 100)]
    elif f_status_ov == "Blocked (<50%)":
        filtered_master = filtered_master[filtered_master["Fulfil %"] < 50]

    # Renumber after filter
    filtered_master = filtered_master.reset_index(drop=True)
    filtered_master["S.No"] = filtered_master.index + 1

    # ── Dynamic summary KPIs ─────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    ov1,ov2,ov3,ov4,ov5,ov6 = st.columns(6)
    _ov_sqm_deficit = (
        filtered_master["Remaining SQM"] *
        (1 - filtered_master["Fulfil %"].clip(0, 100) / 100)
    ).sum()
    _t5_base_cols = ["Equipment","Sys_Code","Total SQM","Already Done SQM","Remaining SQM","Fulfil %"]
    _t5_base = filtered_master[[c for c in _t5_base_cols if c in filtered_master.columns]].copy()
    _t5_short_dd = _t5_base[filtered_master["Fulfil %"]<100].copy()
    _t5_short_dd["Shortfall SQM"] = (_t5_short_dd["Remaining SQM"] * (1 - _t5_short_dd["Fulfil %"].clip(0,100)/100)).round(2)
    _t5_short_dd = _t5_short_dd.sort_values("Shortfall SQM", ascending=False).reset_index(drop=True)
    with ov1:
        dbl_click_metric("No. of Items (filtered)", str(len(filtered_master)), "t5_rows",
            "All Filtered by No. of Items", _t5_base.reset_index(drop=True),
            help_text="Number of (Equipment, System Code) pairs in current filter.")
    with ov2:
        dbl_click_metric("Total SQM", f'{filtered_master["Total SQM"].sum():,.1f}', "t5_sqm",
            "Total SQM by No. of Items (sorted desc)",
            _t5_base.sort_values("Total SQM", ascending=False).reset_index(drop=True),
            help_text="Sum of original SQM for filtered rows.")
    with ov3:
        dbl_click_metric("Already Done SQM", f'{filtered_master["Already Done SQM"].sum():,.1f}', "t5_done",
            "Completed SQM by No. of Items (sorted desc)",
            _t5_base.sort_values("Already Done SQM", ascending=False).reset_index(drop=True),
            help_text="SQM already completed (from daily consumption entries).")
    with ov4:
        dbl_click_metric("Remaining SQM", f'{filtered_master["Remaining SQM"].sum():,.1f}', "t5_rem",
            "Remaining SQM by No. of Items (sorted desc)",
            _t5_base.sort_values("Remaining SQM", ascending=False).reset_index(drop=True),
            help_text="SQM still to be completed = Total − Done.")
    with ov5:
        dbl_click_metric("Shortfall SQM", f"{_ov_sqm_deficit:,.1f}", "t5_short",
            "by No. of Items with SQM Shortfall (sorted desc)", _t5_short_dd,
            help_text="SQM that cannot be completed across filtered rows, weighted by material fulfillment %.")
    with ov6:
        dbl_click_metric("Avg Coverage",
            f'{filtered_master["Fulfil %"].mean():.1f}%' if len(filtered_master) else "0%",
            "t5_avg_cov",
            "Coverage by by No. of Items (sorted asc)",
            _t5_base.sort_values("Fulfil %").reset_index(drop=True),
            help_text="Average fulfillment % across filtered (Equipment, System Code) pairs.")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Colour-coded master table ─────────────────────────────────────────────
    def _style_master(row):
        pct = row["Fulfil %"]
        if pct >= 100:  bg,tc = "rgba(16,185,129,.1)","#10B981"
        elif pct >= 90: bg,tc = "rgba(249,115,22,.1)","#F97316"
        elif pct >= 80: bg,tc = "rgba(234,179,8,.1)", "#EAB308"
        else:           bg,tc = "rgba(239,68,68,.1)", "#EF4444"
        styles = [f"background-color:{bg}"] * len(row)
        ci = list(row.index).index("Fulfil %")
        styles[ci] = f"background-color:{bg};color:{tc};font-weight:700"
        return styles

    styled_master = (filtered_master.style
        .apply(_style_master, axis=1)
        .format({
            "Total SQM":        "{:,.2f}",
            "Lining Area SQM":  "{:,.3f}",
            "Already Done SQM": "{:,.2f}",
            "Remaining SQM":    "{:,.2f}",
            "Total Demand":     "{:,.3f}",
            "Allocated":        "{:,.3f}",
            "Shortfall Qty":    "{:,.3f}",
            "Fulfil %":         "{:.1f}%",
        }))
    st.dataframe(styled_master, use_container_width=True, hide_index=True,
                 height=min(700, 50 + len(filtered_master)*35),
                 key="ov_master_tbl")

    # ── Per-System-Code material detail (expandable) ─────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)

    # Sub-metrics for the whole filtered selection
    _ov_total_sqm   = filtered_master["Total SQM"].sum()
    _ov_done_sqm    = filtered_master["Already Done SQM"].sum()
    _ov_pending_sqm = filtered_master["Remaining SQM"].sum()
    _ov_comp_pct    = (_ov_done_sqm / _ov_total_sqm * 100) if _ov_total_sqm > 0 else 0.0
    _sm1, _sm2, _sm3, _sm4 = st.columns(4)
    _sm1.metric("Total SQM",        f"{_ov_total_sqm:,.2f}")
    _sm2.metric("Already Done SQM", f"{_ov_done_sqm:,.2f}")
    _sm3.metric("Pending SQM",      f"{_ov_pending_sqm:,.2f}")
    _sm4.metric("Completion",       f"{_ov_comp_pct:.1f}%")

    st.markdown('<div class="sec-hdr">🔬 Material Detail by System Code</div>',
                unsafe_allow_html=True)

    for code in sorted(filtered_master["System Code"].unique().tolist(), key=int):
        sc_dm = dm[dm["Lining_System_Code"]==code]
        if sc_dm.empty: continue
        sname   = filtered_master[filtered_master["System Code"]==code]["System Name"].iloc[0]
        sc_sqm  = filtered_master[filtered_master["System Code"]==code]["Total SQM"].sum()
        done_sq = filtered_master[filtered_master["System Code"]==code]["Already Done SQM"].sum()
        sc_mat  = sc_dm.groupby(["Material_Code","Material_Name","UOM"],
                                 as_index=False)["Demand_Qty"].sum()
        sc_mat  = sc_mat.merge(inv[["Material_Code","Available_Qty"]],
                               on="Material_Code", how="left")
        sc_mat["Available_Qty"] = sc_mat["Available_Qty"].fillna(0)
        sc_mat["Shortfall"]     = (sc_mat["Demand_Qty"]-sc_mat["Available_Qty"]).clip(lower=0).round(3)
        sc_mat["Coverage_%"]    = (
            sc_mat["Available_Qty"].clip(upper=sc_mat["Demand_Qty"])
            / sc_mat["Demand_Qty"].replace(0,np.nan)*100
        ).fillna(100).clip(0,100).round(1)
        # SQM achievable based on MINIMUM material coverage (bottleneck)
        sc_cov_min = sc_mat["Coverage_%"].min() if len(sc_mat) else 100
        sc_cov_avg = (sc_mat["Available_Qty"].clip(upper=sc_mat["Demand_Qty"]).sum() /
                      sc_mat["Demand_Qty"].sum()*100) if sc_mat["Demand_Qty"].sum()>0 else 100
        sc_can      = sc_sqm * min(1.0, sc_cov_avg/100)
        dot = "🟢" if sc_cov_avg>=100 else "🟠" if sc_cov_avg>=90 else "🟡" if sc_cov_avg>=80 else "🔴"

        with st.expander(
            f"{dot}  Code {code}  ·  {sname}  ·  "
            f"{sc_can:,.1f}/{sc_sqm:,.1f} SQM  ·  Done: {done_sq:,.1f}  ·  {sc_cov_avg:.1f}%",
            expanded=False,
        ):
            m1,m2,m3,m4,m5 = st.columns(5)
            _t5sc_dd = sc_mat[["Material_Code","Material_Name","UOM","Available_Qty","Demand_Qty","Shortfall","Coverage_%"]].rename(columns={"Coverage_%":"Coverage %"}).reset_index(drop=True)
            _t5sc_sk = f"t5sc_{code}"
            with m1:
                dbl_click_metric("System Code", f"Code {code}", f"{_t5sc_sk}_c",
                    f"Code {code} — Material Breakdown", _t5sc_dd)
            with m2:
                dbl_click_metric("Short Name", str(sname), f"{_t5sc_sk}_n",
                    f"{sname} — Material Breakdown", _t5sc_dd)
            with m3:
                dbl_click_metric("Total SQM", f"{sc_sqm:,.2f}", f"{_t5sc_sk}_s",
                    f"Code {code} — Material Breakdown", _t5sc_dd)
            with m4:
                dbl_click_metric("Already Done SQM", f"{done_sq:,.2f}", f"{_t5sc_sk}_d",
                    f"Code {code} — Material Breakdown", _t5sc_dd,
                    help_text="SQM completed via Daily Consumption entries.")
            with m5:
                dbl_click_metric("Coverage SQM", f"{sc_can:,.2f}  ({sc_cov_avg:.1f}%)", f"{_t5sc_sk}_p",
                    f"Code {code} — Material Breakdown", _t5sc_dd,
                    help_text="SQM coverable with current available material balance.")

            mat_show = sc_mat[["Material_Code","Material_Name","UOM",
                                "Available_Qty","Demand_Qty","Shortfall",
                                "Coverage_%"]].copy()
            mat_show.columns = ["Code","Material Name","UOM",
                                 "Available","Total Demand","Shortfall","Coverage %"]
            def _style_ov_det(row):
                pct=row["Coverage %"]
                if pct>=100:  bg,tc="rgba(16,185,129,.1)","#10B981"
                elif pct>=90: bg,tc="rgba(249,115,22,.1)","#F97316"
                elif pct>=80: bg,tc="rgba(234,179,8,.1)", "#EAB308"
                else:         bg,tc="rgba(239,68,68,.1)", "#EF4444"
                styles=[f"background-color:{bg}"]*len(row)
                ci=list(row.index).index("Coverage %")
                styles[ci]=f"background-color:{bg};color:{tc};font-weight:700"
                return styles
            st.dataframe(
                mat_show.style.apply(_style_ov_det,axis=1).format({
                    "Available":"{:,.3f}","Total Demand":"{:,.3f}",
                    "Shortfall":"{:,.3f}","Coverage %":"{:.1f}%"}),
                use_container_width=True, hide_index=True,
                height=65+len(mat_show)*35, key=f"ov_det_{code}")

    # ── Downloads ─────────────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            "⬇ Download Filtered Master Table",
            data=generate_excel_report(filtered_master, f"Total Overview — {date.today()}", color_scheme="overview"),
            file_name=f"total_overview_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with dl2:
        if db_available():
            conn = get_db()
            full_log = pd.read_sql("SELECT * FROM consumption_log ORDER BY submitted_at DESC", conn)
            conn.close()
            st.download_button(
                "⬇ Download Full Consumption Log",
                data=generate_excel_report(full_log, "Full Consumption Log", color_scheme="overview"),
                file_name=f"consumption_log_full_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# TAB: MASTER DATA — ADD EQUIPMENT
# ─────────────────────────────────────────────────────────────────────────────


# ─────────────────────────────────────────────────────────────────────────────
# TAB: MASTER DATA — FULL TABLE MANAGEMENT
# ─────────────────────────────────────────────────────────────────────────────
with tab_master:
    st.markdown('<div class="sec-hdr">🗄️ Master Data — View, Add & Delete Records</div>',
                unsafe_allow_html=True)

    if not db_available():
        st.error("Database not found. Run `python setup_db.py` first.")
        st.stop()

    # ── Radio selector ─────────────────────────────────────────────────────────
    md_table_sel = st.radio(
        "Select Table to Manage",
        options=[
            "Equipment",
            "LINING SYSTEM MATERIAL CONSM",
            "Materials_DetailsAvailable_Qty",
        ],
        key="md_table_radio",
        horizontal=True,
    )
    TABLE_MAP = {
        "Equipment":                      "equipment",
        "LINING SYSTEM MATERIAL CONSM":   "recipe",
        "Materials_DetailsAvailable_Qty":  "inventory",
    }
    db_table = TABLE_MAP[md_table_sel]
    st.markdown("<hr>", unsafe_allow_html=True)

    # ── Auto-fill helper for Equipment smart entry ─────────────────────────────
    def _get_autofill(code: str) -> dict:
        conn = get_db()
        rec = conn.execute(
            "SELECT lining_system_short_name, lining_type FROM recipe "
            "WHERE lining_system_code = ? LIMIT 1", (code,)
        ).fetchone()
        eq_row = conn.execute(
            'SELECT "Lining_System", "Material Spec.", "Lining_Area/location" FROM equipment '
            'WHERE lining_system_code = ? AND "Lining_System" IS NOT NULL LIMIT 1', (code,)
        ).fetchone()
        conn.close()
        return {
            "lining_system_short_name": rec["lining_system_short_name"] if rec else "",
            "lining_type":              rec["lining_type"]              if rec else "",
            "Lining_System":            eq_row["Lining_System"]         if eq_row else "",
            "Material Spec.":           eq_row["Material Spec."]        if eq_row else "",
            "Lining_Area/location":     eq_row["Lining_Area/location"]  if eq_row else None,
        }

    # ══════════════════════════════════════════════════════════════════════════
    # ADD NEW ROW SECTION
    # ══════════════════════════════════════════════════════════════════════════
    st.markdown('<div class="sec-hdr">➕ Add New Row</div>', unsafe_allow_html=True)

    # ── EQUIPMENT: Smart Entry ─────────────────────────────────────────────────
    if md_table_sel == "Equipment":
        conn = get_db()
        recipe_codes_df = pd.read_sql(
            "SELECT DISTINCT lining_system_code, lining_system_short_name, lining_type "
            "FROM recipe ORDER BY CAST(lining_system_code AS INTEGER)", conn)
        conn.close()

        code_opts = [
            f"Code {r.lining_system_code} — {r.lining_system_short_name}"
            for _, r in recipe_codes_df.iterrows()
        ]
        sel_codes_display = st.multiselect(
            "🔧 Select Lining System Code(s) *",
            options=code_opts, key="seq_codes_pre",
            help="One row will be inserted into the equipment table per selected code.",
        )

        if sel_codes_display:
            conn = get_db()
            eq_col_info = conn.execute("PRAGMA table_info(equipment)").fetchall()
            conn.close()

            SKIP_FOR_FORM = {
                "id", "lining_system_code", "lining_system_short_name",
                "lining_type",
                "lining_system",         # actual DB col: "Lining_System"
                "material spec.",        # actual DB col: "Material Spec."
                "lining_area/location",  # actual DB col: "Lining_Area/location"
                "equipment_tag", "surface_area_sqm", "location",
                "sl. #", "sl.#", "sl. no.", "sl. no", "sl.no.",
            }
            _skip_ff_lower = {s.lower() for s in SKIP_FOR_FORM}
            shared_cols = [(n, t) for (_, n, t, *__) in eq_col_info
                           if n.lower() not in _skip_ff_lower]

            with st.form(key="smart_eq_form"):
                st.markdown('<div class="sec-hdr">Equipment Identity</div>',
                            unsafe_allow_html=True)
                _seq_c1, _seq_c2 = st.columns(2)
                with _seq_c1:
                    eq_tag_inp = st.text_input("🏷️ Equipment Tag No. *",
                                               placeholder="e.g. V-1001", key="seq_tag")
                with _seq_c2:
                    loc_inp = st.selectbox("📍 Location *", options=LOCATION_ORDER,
                                           key="seq_loc")

                if shared_cols:
                    st.markdown(
                        '<div class="sec-hdr" style="margin-top:.8rem;">'
                        'Equipment Details (shared across all selected codes)</div>',
                        unsafe_allow_html=True)
                    shared_inputs = {}
                    for _si in range(0, len(shared_cols), 3):
                        _row_cols = st.columns(3)
                        for _sj, (_sn, _st) in enumerate(shared_cols[_si:_si+3]):
                            with _row_cols[_sj]:
                                if any(kw in _sn.lower() for kw in ("sqm", "qty", "for_1")):
                                    shared_inputs[_sn] = st.number_input(
                                        _sn.replace("_", " ").title(),
                                        value=0.0, step=0.1, key=f"seq_sh_{_sn}")
                                else:
                                    shared_inputs[_sn] = st.text_input(
                                        _sn.replace("_", " ").title(),
                                        key=f"seq_sh_{_sn}")
                else:
                    shared_inputs = {}

                st.markdown('<div class="sec-hdr" style="margin-top:.8rem;">Per Lining System Code</div>',
                            unsafe_allow_html=True)
                per_code_sqm = {}
                for _cd in sel_codes_display:
                    _code = _cd.split(" — ")[0].replace("Code ", "").strip()
                    _af   = _get_autofill(_code)
                    st.markdown(
                        f'<div style="margin:.5rem 0 .3rem;">'
                        f'<span class="code-badge">Code {_code}</span>'
                        f'<span style="font-size:.8rem;color:var(--t2);margin-left:.6rem;">'
                        f'{_af["lining_system_short_name"]}</span></div>',
                        unsafe_allow_html=True)
                    _ca1, _ca2, _ca3 = st.columns(3)
                    with _ca1:
                        st.text_input("Lining System Short Name",
                                      value=_af["lining_system_short_name"],
                                      disabled=True, key=f"seq_sn_{_code}")
                    with _ca2:
                        st.text_input("Lining Type",
                                      value=_af["lining_type"],
                                      disabled=True, key=f"seq_lt_{_code}")
                    with _ca3:
                        st.text_input("Material Spec.",
                                      value=_af["Material Spec."] or "",
                                      disabled=True, key=f"seq_ms_{_code}")
                    _cb1, _cb2 = st.columns(2)
                    with _cb1:
                        st.text_input("Lining System+",
                                      value=_af["Lining_System"] or "",
                                      disabled=True, key=f"seq_ls_{_code}")
                    with _cb2:
                        per_code_sqm[_code] = st.number_input(
                            f"Surface Area SQM * (Code {_code})",
                            min_value=0.0, value=0.0, step=0.1, format="%.2f",
                            key=f"seq_sqm_{_code}",
                            help="Required — used for material demand calculation.")

                st.markdown("<br>", unsafe_allow_html=True)
                seq_submit = st.form_submit_button("💾 Save Equipment",
                                                   use_container_width=False)

            if seq_submit:
                _eq_tag_val = st.session_state.get("seq_tag", "").strip()
                if not _eq_tag_val:
                    st.error("Equipment Tag No. is required.")
                elif not sel_codes_display:
                    st.error("Select at least one Lining System Code.")
                else:
                    _missing_shared = [
                        k.replace("_", " ").title()
                        for k, v in shared_inputs.items()
                        if isinstance(v, str) and v.strip() == ""
                    ]
                    _bad = [c for c in per_code_sqm if per_code_sqm[c] <= 0]
                    if _missing_shared:
                        st.error(f"Please fill in all mandatory fields: {', '.join(_missing_shared)}")
                    elif _bad:
                        st.error(f"Surface Area SQM must be > 0 for codes: {', '.join(_bad)}")
                    else:
                        try:
                            conn = get_db(); cur = conn.cursor()
                            for _cd in sel_codes_display:
                                _code = _cd.split(" — ")[0].replace("Code ", "").strip()
                                _af   = _get_autofill(_code)
                                _sqm  = per_code_sqm[_code]
                                # Build INSERT dynamically from shared_inputs + fixed fields
                                _fixed = {
                                    "location":                st.session_state.get("seq_loc", LOCATION_ORDER[0]),
                                    "lining_system_code":      _code,
                                    "lining_system_short_name": _af["lining_system_short_name"],
                                    "lining_type":             _af["lining_type"],
                                    "equipment_tag":           _eq_tag_val,
                                    "Material Spec.":          _af["Material Spec."] or None,
                                    "surface_area_sqm":        _sqm,
                                    "Lining_System":           _af["Lining_System"] or None,
                                    "Lining_Area/location":    _af.get("Lining_Area/location"),
                                }
                                _all_vals = {**_fixed}
                                for _k, _v in shared_inputs.items():
                                    _all_vals[_k] = _v if _v != "" else None
                                _cols_str = ", ".join([f'"{c}"' for c in _all_vals.keys()])
                                _ph       = ", ".join(["?"] * len(_all_vals))
                                cur.execute(
                                    f"INSERT INTO equipment ({_cols_str}) VALUES ({_ph})",
                                    list(_all_vals.values()))
                                cur.execute("""
                                    INSERT INTO sqm_progress
                                        (equipment_tag, lining_system_code,
                                         original_sqm, done_sqm)
                                    VALUES (?, ?, ?, 0)
                                    ON CONFLICT(equipment_tag, lining_system_code)
                                    DO UPDATE SET original_sqm = excluded.original_sqm
                                """, (_eq_tag_val, _code, _sqm))
                            conn.commit(); conn.close()
                            st.cache_data.clear()
                            st.success(
                                f"✅ Equipment **{_eq_tag_val}** saved for "
                                f"{len(sel_codes_display)} system code(s).")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Database error: {e}")

            # Clear Form button (outside form, still inside if sel_codes_display)
            if st.button("🧹 Clear Equipment Form", key="seq_clr_btn"):
                _clr_keys = (
                    ["seq_tag", "seq_loc", "seq_codes_pre"]
                    + [f"seq_sh_{n}" for n, _ in shared_cols]
                    + [k for k in list(st.session_state.keys())
                       if k.startswith(("seq_sn_", "seq_lt_", "seq_ms_", "seq_ls_", "seq_sqm_"))]
                )
                for _k in _clr_keys:
                    st.session_state.pop(_k, None)
                st.rerun()

        else:
            st.info("Select one or more Lining System Codes above to build the entry form.")

    # ── RECIPE / INVENTORY: Dynamic form from PRAGMA ───────────────────────────
    else:
        conn = get_db()
        col_info = conn.execute(f"PRAGMA table_info({db_table})").fetchall()
        conn.close()

        SKIP_DYN = {"id", "sl. #", "sl.#", "sl. no.", "sl. no", "sl.no."}
        _skip_dyn_lower = {s.lower() for s in SKIP_DYN}
        editable_cols = [(n, t) for (_, n, t, *__) in col_info
                         if n.lower() not in _skip_dyn_lower]

        with st.form(key=f"dyn_add_{db_table}"):
            dyn_inputs = {}
            for _di in range(0, len(editable_cols), 3):
                _drow = st.columns(3)
                for _dj, (_dn, _dt) in enumerate(editable_cols[_di:_di+3]):
                    with _drow[_dj]:
                        if any(kw in _dn.lower() for kw in ("sqm", "qty", "for_1")):
                            dyn_inputs[_dn] = st.number_input(
                                _dn.replace("_", " ").title(),
                                value=0.0, step=0.001, format="%.4f",
                                key=f"dyn_{db_table}_{_dn}")
                        else:
                            dyn_inputs[_dn] = st.text_input(
                                _dn.replace("_", " ").title(),
                                key=f"dyn_{db_table}_{_dn}")
            st.markdown("<br>", unsafe_allow_html=True)
            dyn_submit = st.form_submit_button(f"➕ Add Row to {md_table_sel}")

        if dyn_submit:
            _missing_fields = []
            for _fn, _fv in dyn_inputs.items():
                if isinstance(_fv, str) and _fv.strip() == "":
                    _missing_fields.append(_fn.replace("_", " ").title())
                elif isinstance(_fv, float) and _fv == 0.0:
                    if any(kw in _fn.lower() for kw in ("sqm", "qty")):
                        _missing_fields.append(_fn.replace("_", " ").title())
            if _missing_fields:
                st.error(f"Please fill in all mandatory fields: {', '.join(_missing_fields)}")
            else:
                try:
                    conn = get_db(); cur = conn.cursor()
                    _dcols = list(dyn_inputs.keys())
                    _dvals = [dyn_inputs[c] if dyn_inputs[c] != "" else None for c in _dcols]
                    _cols_str = ", ".join([f'"{c}"' for c in _dcols])
                    _ph       = ", ".join(["?"] * len(_dcols))
                    cur.execute(f"INSERT INTO {db_table} ({_cols_str}) VALUES ({_ph})", _dvals)
                    conn.commit(); conn.close()
                    st.cache_data.clear()
                    st.success("✅ Row added successfully.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Database error: {e}")

        # Clear Form button (outside form)
        if st.button("🧹 Clear Form", key=f"dyn_clr_{db_table}"):
            for _k in list(st.session_state.keys()):
                if _k.startswith(f"dyn_{db_table}_"):
                    del st.session_state[_k]
            st.rerun()

    # ══════════════════════════════════════════════════════════════════════════
    # VIEW & DELETE SECTION
    # ══════════════════════════════════════════════════════════════════════════
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f'<div class="sec-hdr">📋 View, Edit & Delete — {md_table_sel}</div>',
                unsafe_allow_html=True)

    conn = get_db()
    try:
        view_df = pd.read_sql(f"SELECT * FROM {db_table} ORDER BY rowid", conn)
    except Exception:
        view_df = pd.DataFrame()
    conn.close()

    PK_MAP = {
        "equipment":  ("id",            int),
        "recipe":     ("id",            int),
        "inventory":  ("material_code", str),
    }
    pk_col, pk_cast = PK_MAP.get(db_table, ("id", int))

    if view_df.empty:
        st.info("No records found in this table.")
    else:
        view_df_display = view_df.copy()
        _sl_db_cols = [c for c in view_df_display.columns
                       if c.lower().strip() in {"sl. #", "sl.#", "sl. no.", "sl.no.", "sl. no"}]
        view_df_display = view_df_display.drop(columns=_sl_db_cols, errors="ignore")
        view_df_display.insert(0, "Sl. No.", range(1, len(view_df_display) + 1))

        # ── Search filter ─────────────────────────────────────────────────────
        _search_cols = [c for c in view_df_display.columns if c not in ("Sl. No.", "☐ Select")]
        _srch1, _srch2 = st.columns([2, 1])
        with _srch1:
            _md_search = st.text_input(
                "🔍 Search table...", key=f"md_search_{db_table}",
                placeholder="Type to filter…")
        with _srch2:
            _md_col = st.selectbox(
                "in column", options=["All columns"] + _search_cols,
                key=f"md_col_{db_table}", label_visibility="visible")
        if _md_search.strip():
            if _md_col == "All columns":
                _mask = view_df_display[_search_cols].apply(
                    lambda col: col.astype(str).str.contains(
                        _md_search.strip(), case=False, na=False)
                ).any(axis=1)
            else:
                _mask = view_df_display[_md_col].astype(str).str.contains(
                    _md_search.strip(), case=False, na=False)
            view_df_display = view_df_display[_mask]
        view_df_display = view_df_display.reset_index(drop=True)
        view_df_display["Sl. No."] = range(1, len(view_df_display) + 1)

        # Add checkbox column for bulk delete
        view_df_display.insert(0, "☐ Select", False)

        # Build column_config
        _col_cfg = {
            "Sl. No.": st.column_config.NumberColumn("Sl. No.", disabled=True),
            "☐ Select": st.column_config.CheckboxColumn(
                "☐", help="Check rows to delete, then click 'Delete Selected'",
                default=False),
        }
        if pk_col in view_df_display.columns:
            _cfg_type = (st.column_config.NumberColumn if pk_cast == int
                         else st.column_config.TextColumn)
            _col_cfg[pk_col] = _cfg_type(pk_col, disabled=True)

        st.data_editor(
            view_df_display,
            key=f"md_editor_{db_table}",
            num_rows="fixed",
            hide_index=True,
            use_container_width=True,
            height=min(600, 50 + len(view_df_display) * 35),
            column_config=_col_cfg,
        )
        st.caption(f"Total entries: {len(view_df_display)}")

        _btn1, _btn2, _btn3 = st.columns([2, 2, 3])

        with _btn1:
            if st.button("💾 Save Cell Edits", type="primary", key=f"save_edits_{db_table}"):
                _editor_state = st.session_state.get(f"md_editor_{db_table}", {})
                _edited_rows  = _editor_state.get("edited_rows", {})
                if not _edited_rows:
                    st.info("No changes detected in the grid.")
                else:
                    try:
                        conn = get_db(); cur = conn.cursor()
                        n_saved = 0
                        for _row_idx, _changes in _edited_rows.items():
                            # Guard: skip structural + checkbox columns
                            _safe = {k: v for k, v in _changes.items()
                                     if k not in ("Sl. No.", pk_col, "☐ Select")}
                            if not _safe:
                                continue
                            _pk_val    = view_df.iloc[int(_row_idx)][pk_col]
                            _set_parts = [f'"{k}" = ?' for k in _safe.keys()]
                            _set_sql   = ", ".join(_set_parts)
                            _vals      = list(_safe.values()) + [pk_cast(_pk_val)]
                            cur.execute(
                                f'UPDATE {db_table} SET {_set_sql} WHERE "{pk_col}" = ?',
                                _vals)
                            n_saved += 1
                        conn.commit(); conn.close()
                        st.cache_data.clear()
                        st.success(f"✅ {n_saved} row(s) updated successfully.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Database error: {e}")

        with _btn2:
            if st.button("🗑️ Delete Selected Rows", type="secondary",
                         key=f"del_checked_{db_table}"):
                _editor_state = st.session_state.get(f"md_editor_{db_table}", {})
                _edited_rows  = _editor_state.get("edited_rows", {})
                _del_indices = [int(idx) for idx, changes in _edited_rows.items()
                                if changes.get("☐ Select", False)]
                if not _del_indices:
                    st.warning("No rows checked for deletion. Check the ☐ column first.")
                else:
                    try:
                        conn = get_db(); cur = conn.cursor()
                        n_deleted = 0
                        for _di in _del_indices:
                            _del_id = view_df.iloc[_di][pk_col]
                            if db_table == "equipment":
                                _match = view_df[view_df[pk_col] == _del_id]
                                if not _match.empty:
                                    _er = _match.iloc[0]
                                    cur.execute(
                                        "DELETE FROM sqm_progress "
                                        "WHERE equipment_tag = ? AND lining_system_code = ?",
                                        (_er["equipment_tag"], _er["lining_system_code"]))
                            cur.execute(
                                f'DELETE FROM {db_table} WHERE "{pk_col}" = ?',
                                (pk_cast(_del_id),))
                            n_deleted += 1
                        conn.commit(); conn.close()
                        st.cache_data.clear()
                        st.success(f"✅ {n_deleted} row(s) deleted from `{db_table}`.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Database error: {e}")

        st.caption("⚠️ Deletion is permanent. Equipment rows also remove the matching sqm_progress record.")

        st.download_button(
            f"⬇ Download {md_table_sel} Table",
            data=generate_excel_report(
                view_df_display.drop(columns=["Sl. No.", "☐ Select"], errors="ignore"),
                f"{md_table_sel} — Smart Material Estimator",
                color_scheme=_TABLE_COLOR_MAP.get(db_table, "dashboard")),
            file_name=f"{db_table}_export_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{db_table}",
        )
