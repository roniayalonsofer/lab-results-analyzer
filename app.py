# app.py  --  Streamlit UI for the Lab Results Analyzer
# Run: py -3 -m streamlit run app.py
import sys, os, io, collections, socket, base64
from datetime import date

# ── add soil_lab_tool to path ─────────────────────────────────────
ROOT     = os.path.dirname(os.path.abspath(__file__))
TOOL_DIR = os.path.join(ROOT, 'soil_lab_tool')
if TOOL_DIR not in sys.path:
    sys.path.insert(0, TOOL_DIR)

THRESH_DIR = os.path.join(TOOL_DIR, 'thresholds')

# ── company logo (base64 embed) ────────────────────────────────────
def _logo_b64() -> str:
    for name in ('תמונה1.png', 'logo.png'):
        logo_path = os.path.join(ROOT, name)
        if os.path.exists(logo_path):
            with open(logo_path, 'rb') as f:
                return base64.b64encode(f.read()).decode()
    return ""

LOGO_B64 = _logo_b64()
LOGO_TAG = (f'<img src="data:image/png;base64,{LOGO_B64}" '
            f'style="width:100%;max-width:200px;display:block;">'
            if LOGO_B64 else '🧪')
LAB_DIR    = os.path.join(ROOT, 'Laboratory_results')

import streamlit as st
import pandas as pd

# ── page config ───────────────────────────────────────────────────
st.set_page_config(
    page_title="מערכת ניתוח תוצאות מעבדה",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── detect LAN IP for sharing ─────────────────────────────────────
def _local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "localhost"

LAN_IP  = _local_ip()
APP_URL = f"http://{LAN_IP}:8501"

# ══════════════════════════════════════════════════════════════════
# CSS — full design system
# ══════════════════════════════════════════════════════════════════
st.markdown(f"""
<style>
/* ── base layout ── */
/* Apply RTL only to content, not to root — prevents RTL from fighting
   Streamlit's LTR CSS grid and causing the sidebar/main overlap.       */
html, body {{ direction: ltr; }}
[data-testid="stMain"] {{ direction: rtl; }}
[data-testid="stMain"] .block-container {{
    direction: rtl;
    padding-top: 1rem;
    max-width: 1200px;
    padding-left: 2rem;
    padding-right: 2rem;
    /* do not use margin: 0 auto here — let Streamlit control horizontal
       positioning so the block never drifts under the sidebar           */
}}

/* ── sidebar — RTL position (right side) + always expanded ── */
[data-testid="stSidebar"] {{
    direction: rtl;
    background: #1a2d38;
    min-width: 15rem !important;
    max-width: 16rem !important;
    width: 15.5rem !important;
    flex-shrink: 0 !important;
    /* move sidebar to right side for Hebrew/RTL layout */
    right: 0;
    left: auto;
    /* override any JS-driven collapse transform */
    transform: none !important;
    transition: none !important;
    visibility: visible !important;
    display: block !important;
    margin-left: 0 !important;
}}
[data-testid="stSidebarContent"] {{
    direction: rtl;
}}
/* also override the collapsed-state that Streamlit applies via aria */
[data-testid="stSidebar"][aria-expanded="false"] {{
    transform: none !important;
    width: 15.5rem !important;
    min-width: 15rem !important;
    overflow: visible !important;
}}
/* hide ALL collapse/expand controls — cover every known selector variant
   so neither the in-sidebar button nor the floating re-open button shows */
[data-testid="stSidebarCollapseButton"],
[data-testid="stSidebarNavCollapseButton"],
[data-testid="collapsedControl"],
button[aria-label="Close sidebar"],
button[aria-label="Open sidebar"],
button[aria-label="פתח סרגל צד"],
button[aria-label="סגור סרגל צד"] {{
    display: none !important;
}}

[data-testid="stSidebar"] * {{ color: #e2e8f0 !important; }}
/* input / select text — black so it shows on white background */
[data-testid="stSidebar"] input {{ color: #111827 !important; }}
[data-testid="stSidebar"] input::placeholder {{ color: #6b7280 !important; }}
[data-testid="stSidebar"] [data-baseweb="select"] span,
[data-testid="stSidebar"] [data-baseweb="select"] div {{ color: #111827 !important; }}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stTextInput label,
[data-testid="stSidebar"] .stCheckbox label {{ color: #cbd5e1 !important; }}
[data-testid="stSidebar"] .stMarkdown h2,
[data-testid="stSidebar"] .stMarkdown h3 {{ color: #f1f5f9 !important; }}
[data-testid="stSidebar"] hr {{ border-color: #2a4050; }}

/* hide footer / menu / header toolbar
   use display:none (not visibility:hidden) so these elements are fully
   removed from layout and cannot intercept sidebar-related clicks       */
#MainMenu {{ display: none !important; }}
footer {{ display: none !important; }}
header {{ display: none !important; }}

/* ── typography — explicit RTL for all text elements ── */
.stMarkdown p, .stMarkdown li,
.stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4 {{
    direction: rtl; text-align: right;
}}
.stTextInput label, .stSelectbox label,
.stMultiSelect label, .stFileUploader label,
.stCheckbox label, .stRadio label,
.stMetric label, .stMetric div,
[data-testid="stMetricLabel"], [data-testid="stMetricValue"],
[data-testid="stCaptionContainer"] {{
    direction: rtl; text-align: right;
}}

/* ── hero header ── */
.hero {{
    background: linear-gradient(135deg, #2d4a5a 0%, #4a7a8a 60%, #6a9aaa 100%);
    color: white;
    padding: 1.5rem 2rem;
    border-radius: 2px;
    margin-bottom: 1.5rem;
    box-shadow: 0 2px 12px rgba(45,74,90,0.2);
    display: flex;
    align-items: center;
    direction: rtl;
}}
.hero-title {{
    font-size: 1.6rem;
    font-weight: 700;
    margin: 0 0 0.25rem;
    letter-spacing: -0.5px;
}}
.hero-sub {{
    font-size: 0.9rem;
    opacity: 0.85;
    margin: 0;
}}

/* ── section cards ── */
.section-card {{
    background: white;
    border-radius: 2px;
    border: 1px solid #e2e8f0;
    padding: 1.25rem 1.5rem;
    margin-bottom: 1.25rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}}
.section-title {{
    font-size: 0.85rem;
    font-weight: 700;
    color: #2d4a5a;
    letter-spacing: 0.4px;
    text-transform: uppercase;
    margin: 0 0 0.75rem;
    padding-bottom: 0.5rem;
    border-bottom: 2px solid #4a7a8a;
    direction: rtl;
    display: flex;
    align-items: center;
    gap: 0.4rem;
}}

/* ── step indicators ── */
.step-row {{
    display: flex;
    gap: 0.75rem;
    margin-bottom: 1.25rem;
    direction: rtl;
}}
.step-pill {{
    display: flex;
    align-items: center;
    gap: 0.4rem;
    background: #f1f5f9;
    border: 1.5px solid #cbd5e1;
    border-radius: 2px;
    padding: 0.3rem 0.9rem;
    font-size: 0.8rem;
    color: #64748b;
    font-weight: 500;
    flex: 1;
    justify-content: center;
}}
.step-pill.active {{
    background: #eef3f5;
    border-color: #4a7a8a;
    color: #2d4a5a;
    font-weight: 700;
}}
.step-pill.done {{
    background: #f0fdf4;
    border-color: #22c55e;
    color: #15803d;
    border-radius: 2px;
}}

/* ── stat cards ── */
.stats-row {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 0.75rem;
    margin: 0.75rem 0;
}}
.stat-card {{
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 2px;
    padding: 0.9rem 1rem;
    text-align: center;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}}
.stat-num {{
    font-size: 1.75rem;
    font-weight: 800;
    color: #2d4a5a;
    line-height: 1;
    margin-bottom: 0.2rem;
}}
.stat-label {{
    font-size: 0.75rem;
    color: #64748b;
    font-weight: 500;
}}

/* ── upload zone ── */
[data-testid="stFileUploader"] {{
    border: 2.5px dashed #a8c8d2;
    border-radius: 12px;
    padding: 0.5rem;
    background: #eef3f5;
    border-radius: 2px;
    transition: border-color 0.2s;
}}
[data-testid="stFileUploader"]:hover {{
    border-color: #4a7a8a;
    background: #d5e8ec;
}}
[data-testid="stFileUploader"] label {{
    font-size: 0.95rem !important;
    font-weight: 600 !important;
    color: #2d4a5a !important;
}}

/* ── download button ── */
.stDownloadButton button {{
    width: 100%;
    background: linear-gradient(135deg, #16a34a, #15803d);
    color: white;
    font-weight: 700;
    font-size: 1rem;
    border-radius: 2px;
    border: none;
    padding: 0.75rem;
    letter-spacing: 0.5px;
    box-shadow: 0 2px 6px rgba(22,163,74,0.25);
}}
.stDownloadButton button:hover {{
    background: linear-gradient(135deg, #15803d, #166534);
    box-shadow: 0 4px 12px rgba(22,163,74,0.4);
}}

/* ── type badge ── */
.type-badge {{
    display: inline-block;
    padding: 3px 10px;
    border-radius: 2px;
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 0.3px;
    margin: 2px 3px;
    color: white;
}}

/* ── info banner ── */
.info-banner {{
    background: #eef3f5;
    border: 1px solid #b8d4da;
    border-radius: 2px;
    padding: 0.75rem 1rem;
    direction: rtl;
    font-size: 0.875rem;
    color: #2d4a5a;
    margin-bottom: 0.75rem;
    border-right: 3px solid #4a7a8a;
}}
.success-banner {{
    background: #f0fdf4;
    border: 1px solid #86efac;
    border-radius: 2px;
    padding: 0.75rem 1rem;
    direction: rtl;
    font-size: 0.875rem;
    color: #15803d;
    margin-bottom: 0.75rem;
    border-right: 3px solid #16a34a;
}}

/* ── sidebar label ── */
.sidebar-label {{
    font-size: 0.7rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: #94a3b8;
    margin: 1rem 0 0.25rem;
    padding-bottom: 0.2rem;
    border-bottom: 1px solid #2a4050;
}}

/* ── metric overrides ── */
[data-testid="metric-container"] {{
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 2px;
    padding: 0.75rem 1rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}}
[data-testid="stMetricValue"] {{
    font-size: 1.6rem !important;
    color: #2d4a5a !important;
    font-weight: 800 !important;
}}
[data-testid="stMetricLabel"] {{
    font-size: 0.8rem !important;
    color: #64748b !important;
}}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# LOAD THRESHOLD MANAGER
# ══════════════════════════════════════════════════════════════════
@st.cache_resource(show_spinner="טוען ערכי סף...")
def load_threshold_manager(_mtime):
    from core.threshold_manager import ThresholdManager
    MAIN_THRESH  = os.path.join(THRESH_DIR, 'soil_vsl_tier1_v7_2024.xlsx')
    VSL_FULL     = os.path.join(THRESH_DIR, 'soil_vsl_v7_full.xlsx')
    PFAS_THRESH     = os.path.join(LAB_DIR,   'נספח לטבלת ערכי סף - PFAS.xlsx')
    PFAS_THRESH_ALT = os.path.join(THRESH_DIR, 'pfas_thresholds.xlsx')
    vsl_full_path = VSL_FULL if os.path.exists(VSL_FULL) else None
    pfas_path = (PFAS_THRESH     if os.path.exists(PFAS_THRESH)
                 else PFAS_THRESH_ALT if os.path.exists(PFAS_THRESH_ALT)
                 else None)
    return ThresholdManager(MAIN_THRESH, pfas_path=pfas_path, vsl_full_path=vsl_full_path)

# ══════════════════════════════════════════════════════════════════
# SIDEBAR — defined before module imports so it always renders
# ══════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(
        f'<div style="text-align:center;padding:0.5rem 0 0.75rem;">'
        f'<div style="background:white;border-radius:2px;padding:0.6rem 0.8rem;'
        f'margin-bottom:0.5rem;display:inline-block;width:90%;">'
        f'{LOGO_TAG}</div>'
        f'<div style="font-size:0.7rem;color:#94a3b8;margin-top:4px;">Lab Results Analyzer</div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown('<hr style="margin:0.5rem 0 1rem;">', unsafe_allow_html=True)

    st.markdown('<div class="sidebar-label">📋 פרטי פרויקט</div>', unsafe_allow_html=True)
    client_name  = st.text_input("שם לקוח",  value="", label_visibility="collapsed",
                                  placeholder="שם לקוח (לדוג׳: סונול)")
    project_name = st.text_input("שם האתר",  value="", label_visibility="collapsed",
                                  placeholder="שם האתר (לדוג׳: צומת שמשון)")

    st.markdown('<div class="sidebar-label">🏭 מעבדה וקטגוריה</div>', unsafe_allow_html=True)
    lab = st.selectbox("מעבדה", ["🔍 זיהוי אוטומטי", "KTE", "מכון הנפט", "בקטוכם", "Alchem", "ALS", "Aminolab", "RJ Lee", "אלכם (XRF)"],
                       label_visibility="collapsed")
    category_display = {
        "🔍 זיהוי אוטומטי":           "auto",
        "🪨 קרקע (soil)":             "soil",
        "💧 מי תהום (groundwater)":   "groundwater",
        "🧬 PFAS":                    "pfas",
        "📊 PR format (KTE מתכות)":   "pr",
        "💨 גז קרקע (soil_gas)":      "soil_gas",
        "🪨 גרנולומטריה (grain_size)": "grain_size",
    }
    cat_label    = st.selectbox("קטגוריה", list(category_display.keys()),
                                label_visibility="collapsed")
    category_raw = category_display[cat_label]

    st.markdown('<hr style="margin:1rem 0 0.5rem;">', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# HERO HEADER — defined before module imports so it always renders
# ══════════════════════════════════════════════════════════════════
_hero_logo = (
    f'<div style="background:white;border-radius:2px;padding:0.4rem 0.7rem;margin-left:1rem;">'
    f'<img src="data:image/png;base64,{LOGO_B64}" style="height:56px;display:block;"></div>'
    if LOGO_B64 else ''
)
st.html(f"""
<div class="hero">
  <div style="display:flex;align-items:center;gap:1rem;">
    <div>
      <div class="hero-title">מערכת לניתוח ועיבוד תוצאות מעבדה — אדמה</div>
      <div class="hero-sub">העלה קובץ דוח מעבדה · בחר ערכי סף · הורד Excel מסודר</div>
    </div>
  </div>
  {_hero_logo}
</div>
""")

# ══════════════════════════════════════════════════════════════════
# MODULE IMPORTS
# ══════════════════════════════════════════════════════════════════
try:
    from core.excel_output import LabReportExcel
    from core.word_output  import (LabReportWord, parse_als_file, build_word_report,
                                   load_threshold_file, get_tier1_col, tier1_label)
    from core.xrf_output   import build_xrf_simple_excel
    from parsers import get_parser, auto_detect_category, auto_detect_lab
    _tm_py    = os.path.join(TOOL_DIR, 'core', 'threshold_manager.py')
    _tm_mtime = os.path.getmtime(_tm_py)
    tm = load_threshold_manager(_tm_mtime)
except Exception as e:
    st.error(f"שגיאת טעינת מודולים: {e}")
    st.stop()

# ══════════════════════════════════════════════════════════════════
# STEP INDICATOR (used by Excel tab)
# ══════════════════════════════════════════════════════════════════
def _build_pivot_table(recs):
    """Build a compound × sample_id pivot DataFrame for preview."""
    rows = []
    for r in recs:
        val  = r.get('value')
        flag = r.get('flag', '')
        loq  = r.get('loq')
        if flag in ('ND', '<LOD', '<LOQ') and loq is not None:
            display_val = f"<{loq}"
        elif val is not None:
            display_val = f"{val:.4g}" if isinstance(val, (int, float)) else str(val)
        else:
            display_val = flag or ''
        rows.append({
            'תרכובת':    r.get('compound', ''),
            'sample_id': r.get('sample_id', ''),
            'value':     display_val,
        })
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    try:
        pivot = df.pivot_table(index='תרכובת', columns='sample_id',
                               values='value', aggfunc='first')
        pivot.columns.name = None
        return pivot
    except Exception:
        return df[['תרכובת', 'sample_id', 'value']]


_PREVIEW_LEGEND = (
    '<div style="font-size:0.75rem;direction:rtl;margin-top:6px;display:flex;'
    'gap:10px;align-items:center;">'
    '<span style="background:#FFFF00;padding:2px 10px;border-radius:2px;'
    'border:1px solid #ccc;">מעל VSL</span>'
    '<span style="background:#FFC000;padding:2px 10px;border-radius:2px;'
    'border:1px solid #ccc;">מעל Tier 1</span>'
    '</div>'
)


def _build_styled_pivot(recs: list[dict], tm, selected_thresholds: list[str]):
    """Return (Styler, has_any_color) for a compound × sample_id pivot.

    Yellow (#FFFF00) = value exceeds the selected VSL threshold.
    Orange (#FFC000) = value exceeds a selected Tier-1 threshold (takes
                       priority over yellow when both are exceeded).
    """
    if not recs:
        return pd.DataFrame().style, False

    rows_d, rows_n, cas_map = [], [], {}
    for r in recs:
        val      = r.get('value')
        flag     = r.get('flag', '')
        loq      = r.get('loq')
        compound = r.get('compound', '')
        sample   = r.get('sample_id', '')
        cas      = r.get('cas', '')

        if flag in ('ND', '<LOD', '<LOQ') and loq is not None:
            disp = f"<{loq}"
        elif val is not None:
            disp = f"{val:.4g}" if isinstance(val, (int, float)) else str(val)
        else:
            disp = flag or ''

        rows_d.append({'תרכובת': compound, '_s': sample, '_v': disp})
        rows_n.append({'תרכובת': compound, '_s': sample,
                       '_v': val if isinstance(val, (int, float)) else None})
        cas_map.setdefault(compound, cas)

    try:
        piv_d = (pd.DataFrame(rows_d)
                 .pivot_table(index='תרכובת', columns='_s', values='_v', aggfunc='first'))
        piv_n = (pd.DataFrame(rows_n)
                 .pivot_table(index='תרכובת', columns='_s', values='_v', aggfunc='first'))
        piv_d.columns.name = None
        piv_n.columns.name = None
    except Exception:
        return pd.DataFrame(rows_d).style, False

    # VSL keys → yellow; everything else (Tier1, GW, GAS_*) → orange
    vsl_keys   = [k for k in selected_thresholds if 'VSL' in k]
    tier1_keys = [k for k in selected_thresholds if k not in vsl_keys]
    has_colors = bool((vsl_keys or tier1_keys) and tm is not None)

    colors = pd.DataFrame('', index=piv_d.index, columns=piv_d.columns)

    if has_colors:
        for compound in piv_d.index:
            cas = cas_map.get(compound, '')
            for sample in piv_d.columns:
                try:
                    raw = piv_n.loc[compound, sample]
                    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
                        continue
                    num_val = float(raw)
                except (TypeError, ValueError, KeyError):
                    continue

                # Orange: any Tier-1 key exceeded (checked first — higher priority)
                for tk in tier1_keys:
                    thresh = tm.get_threshold_with_name(cas, tk, compound)
                    if thresh is not None and num_val > thresh:
                        colors.loc[compound, sample] = 'background-color: #FFC000'
                        break

                if colors.loc[compound, sample]:
                    continue

                # Yellow: any VSL key exceeded
                for vk in vsl_keys:
                    thresh = tm.get_threshold_with_name(cas, vk, compound)
                    if thresh is not None and num_val > thresh:
                        colors.loc[compound, sample] = 'background-color: #FFFF00'
                        break

    styled = piv_d.style.apply(lambda _: colors, axis=None)
    return styled, has_colors


def _steps(step: int):
    labels = ["① העלאת קובץ", "② בחירת ערכי סף", "③ הורדת דוח"]
    pills = ""
    for i, lbl in enumerate(labels, 1):
        cls = "active" if i == step else ("done" if i < step else "step-pill")
        icon = "✅ " if i < step else ""
        pills += f'<div class="step-pill {cls}">{icon}{lbl}</div>'
    st.markdown(f'<div class="step-row">{pills}</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# TABS
# tab_word must be rendered BEFORE tab_excel so that st.stop() calls
# inside tab_excel don't prevent the Word tab from appearing.
# ══════════════════════════════════════════════════════════════════
tab_excel, tab_word = st.tabs(["📊 יצוא Excel", "📄 יצוא Word"])

# ══════════════════════════════════════════════════════════════════
# WORD TAB
# ══════════════════════════════════════════════════════════════════
with tab_word:
    st.markdown("#### 📄 יצוא דוח Word מקובץ ALS")
    st.caption("העלה קבצי ALS מהמעבדה, בחר ערכי סף והורד דוח Word מעוצב")
    st.markdown("---")

    # ── Threshold settings ────────────────────────────────────────
    st.markdown("##### ⚙️ הגדרות ערכי סף")
    wc1, wc2, wc3, wc4 = st.columns([2, 2, 2, 2])
    with wc1:
        w_thresh_file = st.file_uploader(
            "📂 קובץ ערכי סף (Excel)",
            type=["xlsx", "xls"],
            key="w_thresh",
            help="קובץ Excel עם ערכי VSL ו-TIER 1",
        )
    with wc2:
        w_land = st.selectbox("Land Use", ["Industrial", "Residential"], key="w_land")
    with wc3:
        w_aquifer = st.selectbox(
            "Aquifer Sensitivity", ["A-1, A, B", "B-1 or C"], key="w_aquifer"
        )
    with wc4:
        _w_depth_opts = ["Not Applicable"] if "b-1" in w_aquifer.lower() else ["0 - 6 m", ">6 m"]
        w_depth = st.selectbox("Depth to GW", _w_depth_opts, key="w_depth")

    w_t1col = get_tier1_col(w_land, w_aquifer, w_depth)
    w_t1lbl = tier1_label(w_land, w_aquifer, w_depth)
    st.caption(f"📌 TIER 1: **{w_land}** | {w_aquifer} | {w_depth}")

    st.markdown("---")

    # ── ALS file upload ───────────────────────────────────────────
    st.markdown("##### 📤 קבצי ALS")
    w_files = st.file_uploader(
        "העלה קבצי ALS (Excel)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="w_als",
    )

    st.markdown("---")

    # ── Table options (4 columns) ─────────────────────────────────
    st.markdown("##### 📋 הגדרות טבלאות")
    tcol1, tcol2, tcol3, tcol4 = st.columns(4)

    with tcol1:
        with st.expander("🛢️ TPH", expanded=True):
            w_tph_inc   = st.checkbox("כלול בדוח", value=True,  key="w_tph_inc")
            w_tph_title = st.text_input("כותרת", value="טבלה 1 – TPH",     key="w_tph_title")
            w_tph_page  = st.selectbox("גודל דף", ["A4", "Tabloid"],        key="w_tph_page")
            w_tph_land  = st.selectbox("כיוון", ["לרוחב", "לאורך"],        key="w_tph_orient") == "לרוחב"

    with tcol2:
        with st.expander("⚗️ Metals", expanded=True):
            w_met_inc   = st.checkbox("כלול בדוח", value=True,  key="w_met_inc")
            w_met_title = st.text_input("כותרת", value="טבלה 2 – מתכות",   key="w_met_title")
            w_met_page  = st.selectbox("גודל דף", ["A4", "Tabloid"],        key="w_met_page")
            w_met_land  = st.selectbox("כיוון", ["לרוחב", "לאורך"],        key="w_met_orient") == "לרוחב"

    with tcol3:
        with st.expander("🧪 VOC+SVOC", expanded=True):
            w_voc_inc   = st.checkbox("כלול בדוח", value=True,  key="w_voc_inc")
            w_voc_title = st.text_input("כותרת", value="טבלה 3 – VOC+SVOC", key="w_voc_title")
            w_voc_page  = st.selectbox("גודל דף", ["A4", "Tabloid"],         key="w_voc_page")
            w_voc_land  = st.selectbox("כיוון", ["לרוחב", "לאורך"],         key="w_voc_orient") == "לרוחב"

    with tcol4:
        with st.expander("🔬 PFAS", expanded=True):
            w_pfas_inc   = st.checkbox("כלול בדוח", value=True,  key="w_pfas_inc")
            w_pfas_title = st.text_input("כותרת", value="טבלה 4 – PFAS",   key="w_pfas_title")
            w_pfas_page  = st.selectbox("גודל דף", ["A4", "Tabloid"],       key="w_pfas_page")
            w_pfas_land  = st.selectbox("כיוון", ["לרוחב", "לאורך"],       key="w_pfas_orient") == "לרוחב"

    st.markdown("---")

    # ── Generate button ───────────────────────────────────────────
    _w_ready = bool(w_files and w_thresh_file)
    if not w_files:
        st.info("👆 העלה קבצי ALS כדי להתחיל")
    elif not w_thresh_file:
        st.info("👆 העלה קובץ ערכי סף כדי לצור את הדוח")

    if _w_ready:
        if st.button("📄 צור דוח Word", type="primary",
                     use_container_width=True, key="w_gen"):
            with st.spinner("⏳ בונה דוח..."):
                try:
                    _thresh_dict = load_threshold_file(w_thresh_file.read())

                    _all_dfs: dict = {"TPH": [], "Metals": [], "VOC+SVOC": [], "PFAS": []}
                    _w_errors = []
                    for _wf in w_files:
                        _df, _err = parse_als_file(_wf.read(), _wf.name)
                        if _err:
                            _w_errors.append(f"{_wf.name}: {_err}")
                        elif _df is not None and not _df.empty:
                            _grp = _df["group"].str.upper().str.strip()
                            if _grp.str.contains("PFAS", na=False).any():
                                _all_dfs["PFAS"].append(
                                    _df[_grp.str.contains("PFAS", na=False)]
                                )
                            if _grp.str.contains("VOC|SVOC|BTEX", na=False).any():
                                _all_dfs["VOC+SVOC"].append(
                                    _df[_grp.str.contains("VOC|SVOC|BTEX", na=False)]
                                )
                            if _grp.str.contains(
                                "METAL|INORGANIC|ICP|ELEMENT", na=False
                            ).any():
                                _all_dfs["Metals"].append(
                                    _df[_grp.str.contains(
                                        "METAL|INORGANIC|ICP|ELEMENT", na=False
                                    )]
                                )
                            if _grp.str.contains(
                                "TPH|PETROLEUM|HYDROCARBON|DRO|ORO", na=False
                            ).any():
                                _all_dfs["TPH"].append(
                                    _df[_grp.str.contains(
                                        "TPH|PETROLEUM|HYDROCARBON|DRO|ORO", na=False
                                    )]
                                )

                    for _e in _w_errors:
                        st.warning(f"⚠️ {_e}")

                    _merged = {
                        k: pd.concat(v, ignore_index=True) if v else None
                        for k, v in _all_dfs.items()
                    }

                    _table_cfgs = []
                    if w_tph_inc and _merged.get("TPH") is not None and not _merged["TPH"].empty:
                        _table_cfgs.append({"type": "TPH",      "df": _merged["TPH"],
                                            "title": w_tph_title,  "page_size": w_tph_page,
                                            "landscape": w_tph_land})
                    if w_met_inc and _merged.get("Metals") is not None and not _merged["Metals"].empty:
                        _table_cfgs.append({"type": "Metals",   "df": _merged["Metals"],
                                            "title": w_met_title,  "page_size": w_met_page,
                                            "landscape": w_met_land})
                    if w_voc_inc and _merged.get("VOC+SVOC") is not None and not _merged["VOC+SVOC"].empty:
                        _table_cfgs.append({"type": "VOC+SVOC", "df": _merged["VOC+SVOC"],
                                            "title": w_voc_title,  "page_size": w_voc_page,
                                            "landscape": w_voc_land})
                    if w_pfas_inc and _merged.get("PFAS") is not None and not _merged["PFAS"].empty:
                        _table_cfgs.append({"type": "PFAS",     "df": _merged["PFAS"],
                                            "title": w_pfas_title, "page_size": w_pfas_page,
                                            "landscape": w_pfas_land})

                    if not _table_cfgs:
                        st.warning("⚠️ לא נמצאו נתונים מסווגים בקבצים שהועלו")
                    else:
                        _docx_bytes = build_word_report(
                            _table_cfgs, _thresh_dict, w_t1col, w_t1lbl
                        )
                        st.session_state["w_docx_bytes"] = _docx_bytes
                        st.success(f"✅ הדוח נוצר — {len(_table_cfgs)} טבלאות")

                except Exception as _ex:
                    st.error(f"❌ שגיאה: {_ex}")
                    import traceback
                    st.code(traceback.format_exc())

    if st.session_state.get("w_docx_bytes"):
        st.download_button(
            "⬇️ הורד דוח Word",
            data=st.session_state["w_docx_bytes"],
            file_name="word_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="w_dl",
        )

# ══════════════════════════════════════════════════════════════════
# EXCEL TAB  (existing flow — unchanged)
# ══════════════════════════════════════════════════════════════════
with tab_excel:
    _steps(1)

    # ══════════════════════════════════════════════════════════════
    # UPLOAD
    # ══════════════════════════════════════════════════════════════
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📤 שלב 1 — העלאת קובץ דוח מעבדה</div>',
                unsafe_allow_html=True)

    col_up, col_meta = st.columns([3, 1])

    with col_up:
        uploaded_files = st.file_uploader(
            "גרור קבצים לכאן או לחץ לבחירה",
            type=["xlsx", "xls", "csv", "pdf"],
            accept_multiple_files=True,
            help="ניתן להעלות מספר קבצים מאותה מעבדה | XLSX / XLS / CSV / PDF",
            label_visibility="visible",
        )

    with col_meta:
        # Detect lab early if auto-detect mode and files are already available
        _display_lab = lab
        if uploaded_files and lab == "🔍 זיהוי אוטומטי":
            try:
                _peek = uploaded_files[0].getvalue()
                _det  = auto_detect_lab(uploaded_files[0].name, _peek)
                _display_lab = _det or "KTE"
            except Exception:
                _display_lab = "?"
        elif lab == "🔍 זיהוי אוטומטי":
            _display_lab = "—"

        cat_clean = cat_label.split(" ", 1)[-1] if " " in cat_label else cat_label
        st.markdown(f"""
        <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:2px;
                    padding:0.9rem 1rem;margin-top:1.75rem;text-align:center;">
          <div style="font-size:0.65rem;color:#94a3b8;font-weight:700;letter-spacing:0.8px;
                      text-transform:uppercase;margin-bottom:4px;">מעבדה</div>
          <div style="font-size:1.2rem;font-weight:800;color:#2d4a5a;">{_display_lab}</div>
          <div style="font-size:0.7rem;color:#64748b;margin-top:4px;">{cat_clean}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    if not uploaded_files:
        st.markdown("""
        <div class="info-banner">
          ℹ️ העלה קובץ דוח מעבדה כדי להתחיל — המערכת תזהה אוטומטית את סוג הניתוח
        </div>
        """, unsafe_allow_html=True)
        st.stop()

    # ══════════════════════════════════════════════════════════════
    # PARSE
    # ══════════════════════════════════════════════════════════════
    all_raw: list[tuple[str, bytes]] = [(uf.name, uf.read()) for uf in uploaded_files]
    fname     = " | ".join(f for f, _ in all_raw)
    raw_bytes = all_raw[0][1]

    if lab == "🔍 זיהוי אוטומטי":
        detected_lab = None
        for _fn, _fb in all_raw:
            _det = auto_detect_lab(_fn, _fb)
            if _det and _det != "KTE":
                detected_lab = _det
                break
        if not detected_lab:
            detected_lab = auto_detect_lab(all_raw[0][0], all_raw[0][1]) or "KTE"
        lab = detected_lab

    if category_raw == 'auto':
        category = None
        for _fn, _fb in all_raw:
            _cat = auto_detect_category(_fn, _fb)
            if _cat and _cat != "soil":
                category = _cat
                break
        if not category:
            category = auto_detect_category(all_raw[0][0], all_raw[0][1]) or "soil"
        cat_info = f"זוהה אוטומטית: **{category}**"
    else:
        category = category_raw
        cat_info  = f"קטגוריה: **{category}**"

    try:
        try:
            parser = get_parser(lab, category)
        except KeyError:
            fallback = "soil"
            st.warning(f"⚠️ אין parser עבור {lab} / {category}. מנסה: **{fallback}**")
            category = fallback
            cat_info  = f"ברירת מחדל: **{fallback}**"
            parser   = get_parser(lab, fallback)
    except Exception as e:
        st.error(f"שגיאת טעינת parser: {e}")
        st.exception(e)
        st.stop()

    all_records:    list[dict] = []
    file_summaries: list[dict] = []
    n_files = len(all_raw)

    # Separate companion PDFs from Excel/CSV data files.
    # Collect bytes from ALL uploaded PDFs so every Lab ID → borehole
    # mapping is available regardless of how many PDFs were uploaded.
    _pdf_raws    = [b for n, b in all_raw if n.lower().endswith(".pdf")]
    _data_files  = [(n, b) for n, b in all_raw if not n.lower().endswith(".pdf")]
    _parse_files = _data_files if _data_files else all_raw

    with st.spinner(f"מנתח {'קבצים' if n_files > 1 else 'קובץ'}..."):
        for fname_i, raw_i in _parse_files:
            try:
                try:
                    file_records = parser.parse(io.BytesIO(raw_i), pdf_bytes=_pdf_raws)
                except TypeError:
                    file_records = parser.parse(io.BytesIO(raw_i))
                all_records.extend(file_records)
                file_summaries.append({"name": fname_i, "records": len(file_records), "ok": True})
            except Exception as e:
                st.error(f"שגיאת פרסינג: {fname_i} — {e}")
                file_summaries.append({"name": fname_i, "records": 0, "ok": False})

    records = all_records

    if not records:
        st.markdown('<div class="info-banner">⚠️ לא נמצאו רשומות — בדוק פורמט הקובץ ובחירת מעבדה / קטגוריה</div>',
                    unsafe_allow_html=True)
        st.stop()

    # ── stats ─────────────────────────────────────────────────────
    by_type  = collections.Counter(r.get('analysis_type', '?') for r in records)
    samples  = sorted(set(r['sample_id'] for r in records))
    detected = [r for r in records if r.get('flag') not in ('ND', '<LOD') and r.get('value') is not None]

    # success line
    st.markdown(f"""
    <div class="success-banner">
      ✅ {cat_info} &nbsp;|&nbsp; Parser: <code>{type(parser).__name__}</code>
      {"&nbsp;|&nbsp; " + " ".join(f'<b>{s["name"]}</b>: {s["records"]} רשומות' for s in file_summaries) if n_files > 1 else ""}
    </div>
    """, unsafe_allow_html=True)

    # metric cards
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("סה\"כ רשומות",  f"{len(records):,}")
    with c2: st.metric("ערכים מזוהים",  f"{len(detected):,}")
    with c3: st.metric("דגימות",        f"{len(samples):,}")
    with c4: st.metric("סוגי ניתוח",   f"{len(by_type):,}")

    # analysis-type badges
    BADGE_COLORS = {
        "SOIL_GAS_VOC": "#7c3aed", "SOIL_VOC":    "#0d9488",
        "SOIL_TPH":     "#0891b2", "SOIL_MBTEX":  "#0f766e",
        "SOIL_METALS":  "#4f46e5", "SOIL_PFAS":   "#db2777",
        "GW_VOC":       "#4a7a8a", "GW_PFAS":     "#9333ea",
        "LOWFLOW":      "#6b7280",
    }
    badges = " ".join(
        f'<span class="type-badge" style="background:{BADGE_COLORS.get(t,"#94a3b8")};">'
        f'{t}: {cnt}</span>'
        for t, cnt in by_type.most_common()
    )
    st.markdown(f'<div style="margin:0.5rem 0;">{badges}</div>', unsafe_allow_html=True)

    _steps(2)

    # ══════════════════════════════════════════════════════════════
    # THRESHOLD SELECTION
    # ══════════════════════════════════════════════════════════════
    found_atypes  = list(by_type.keys())
    has_soil      = any(t in found_atypes for t in ("SOIL_VOC","SOIL_TPH","SOIL_METALS","SOIL_MBTEX","XRF"))
    has_soil_pfas = "SOIL_PFAS" in found_atypes
    has_soil_gas  = "SOIL_GAS_VOC" in found_atypes
    has_gw        = any(t.startswith("GW_") for t in found_atypes)

    selected_thresholds: list[str] = []

    _SENS_MAP  = {"רגיש מאוד": "vh", "רגיש/בינוני": "hm", "לא רגיש": "low", "—": None}
    _DEPTH_MAP = {"0-6מ'": "0_6", ">6מ'": "6"}

    def _soil_tier1_key(land_use: str, sens_code, depth_label) -> str | None:
        if not sens_code:
            return None
        pfx = "RES" if land_use == "res" else "IND"
        if sens_code == "vh":  return f"TIER1_{pfx}_SOIL_VH"
        if sens_code == "hm":
            d = _DEPTH_MAP.get(depth_label, "0_6")
            return f"TIER1_{pfx}_SOIL_HM_{d}"
        if sens_code == "low": return f"TIER1_{pfx}_SOIL_LOW"
        return None

    _PFAS_SENS_MAP = {"רגישות גבוהה מאוד": "vh", "רגישות גבוהה/בינונית": "hm", "רגישות נמוכה": "low", "—": None}

    def _pfas_tier1_key(land_use: str, sens_code, depth_label) -> str | None:
        if not sens_code:
            return None
        pfx = "RES" if land_use == "res" else "IND"
        if sens_code == "vh":  return f"PFAS_TIER1_{pfx}_VERY_HIGH"
        if sens_code == "hm":
            return f"PFAS_TIER1_{pfx}_0_6" if depth_label == "0-6מ'" else f"PFAS_TIER1_{pfx}_6PLUS"
        if sens_code == "low": return f"PFAS_TIER1_{pfx}_NO_GW"
        return None

    any_shown = False

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📋 שלב 2 — בחירת ערכי סף להשוואה</div>',
                unsafe_allow_html=True)

    # ── Soil ─────────────────────────────────────────────────────
    if has_soil:
        any_shown = True
        st.markdown("##### 🪨 קרקע")
        col_vsl, col_t1r, col_t1i = st.columns(3)

        with col_vsl:
            st.markdown('<div style="font-size:0.85rem;font-weight:700;color:#374151;margin-bottom:6px;">VSL — ישיר</div>', unsafe_allow_html=True)
            use_vsl = st.checkbox("VSL (Direct Contact)", value=True, key="vsl_cb")

        with col_t1r:
            st.markdown('<div style="font-size:0.85rem;font-weight:700;color:#374151;margin-bottom:6px;">Tier 1 מגורים (Residential)</div>', unsafe_allow_html=True)
            sens_res = st.selectbox("רגישות אקוויפר", ["—","רגיש מאוד","רגיש/בינוני","לא רגיש"], key="sens_res", label_visibility="collapsed")
            depth_res = None
            if sens_res == "רגיש/בינוני":
                depth_res = st.radio('עומק מי"ת', ["0-6מ'",">6מ'"], horizontal=True, key="depth_res", label_visibility="collapsed")

        with col_t1i:
            st.markdown('<div style="font-size:0.85rem;font-weight:700;color:#374151;margin-bottom:6px;">Tier 1 תעשייה (Industrial)</div>', unsafe_allow_html=True)
            sens_ind = st.selectbox("רגישות אקוויפר", ["—","רגיש מאוד","רגיש/בינוני","לא רגיש"], key="sens_ind", label_visibility="collapsed")
            depth_ind = None
            if sens_ind == "רגיש/בינוני":
                depth_ind = st.radio('עומק מי"ת', ["0-6מ'",">6מ'"], horizontal=True, key="depth_ind", label_visibility="collapsed")

        if use_vsl: selected_thresholds.append("VSL_SOIL")
        k = _soil_tier1_key("res", _SENS_MAP.get(sens_res), depth_res)
        if k: selected_thresholds.append(k)
        k = _soil_tier1_key("ind", _SENS_MAP.get(sens_ind), depth_ind)
        if k: selected_thresholds.append(k)

    # ── Soil PFAS ─────────────────────────────────────────────────
    if has_soil_pfas:
        any_shown = True
        st.markdown("##### 🧬 קרקע PFAS")
        cp_vsl, cp_res, cp_ind = st.columns(3)

        with cp_vsl:
            st.markdown('<div style="font-size:0.85rem;font-weight:700;color:#374151;margin-bottom:6px;">VSL — ישיר</div>', unsafe_allow_html=True)
            use_pfas_vsl = st.checkbox("PFAS VSL", value=True, key="pfas_vsl")

        with cp_res:
            st.markdown('<div style="font-size:0.85rem;font-weight:700;color:#374151;margin-bottom:6px;">Tier 1 מגורים (Residential)</div>', unsafe_allow_html=True)
            pfas_sens_res = st.selectbox("רגישות", ["—", "רגישות גבוהה מאוד", "רגישות גבוהה/בינונית", "רגישות נמוכה"],
                                         key="pfas_sens_res", label_visibility="collapsed")
            pfas_depth_res = None
            if pfas_sens_res == "רגישות גבוהה/בינונית":
                pfas_depth_res = st.radio("עומק", ["0-6מ'", ">6מ'"], horizontal=True,
                                          key="pfas_depth_res", label_visibility="collapsed")

        with cp_ind:
            st.markdown('<div style="font-size:0.85rem;font-weight:700;color:#374151;margin-bottom:6px;">Tier 1 תעשייה (Industrial)</div>', unsafe_allow_html=True)
            pfas_sens_ind = st.selectbox("רגישות", ["—", "רגישות גבוהה מאוד", "רגישות גבוהה/בינונית", "רגישות נמוכה"],
                                         key="pfas_sens_ind", label_visibility="collapsed")
            pfas_depth_ind = None
            if pfas_sens_ind == "רגישות גבוהה/בינונית":
                pfas_depth_ind = st.radio("עומק", ["0-6מ'", ">6מ'"], horizontal=True,
                                          key="pfas_depth_ind", label_visibility="collapsed")

        if use_pfas_vsl:
            selected_thresholds.append("PFAS_VSL")
        k = _pfas_tier1_key("res", _PFAS_SENS_MAP.get(pfas_sens_res), pfas_depth_res)
        if k: selected_thresholds.append(k)
        k = _pfas_tier1_key("ind", _PFAS_SENS_MAP.get(pfas_sens_ind), pfas_depth_ind)
        if k: selected_thresholds.append(k)

    # ── Soil gas ──────────────────────────────────────────────────
    if has_soil_gas:
        any_shown = True
        st.markdown("##### 💨 גז קרקע VOC")
        sg_col_r, sg_col_i = st.columns(2)
        with sg_col_r:
            st.markdown('<div style="font-size:0.8rem;font-weight:600;color:#374151;">Tier 1 מגורים</div>', unsafe_allow_html=True)
            sg_res_in  = st.checkbox("Indoor — פנים",  value=True,  key="sg_res_in")
            sg_res_out = st.checkbox("Outdoor — חוץ",  value=False, key="sg_res_out")
        with sg_col_i:
            st.markdown('<div style="font-size:0.8rem;font-weight:600;color:#374151;">Tier 1 תעשייה</div>', unsafe_allow_html=True)
            sg_ind_in  = st.checkbox("Indoor — פנים",  value=False, key="sg_ind_in")
            sg_ind_out = st.checkbox("Outdoor — חוץ",  value=False, key="sg_ind_out")
        if sg_res_in:  selected_thresholds.append("GAS_INDOOR_RES")
        if sg_res_out: selected_thresholds.append("GAS_OUTDOOR_RES")
        if sg_ind_in:  selected_thresholds.append("GAS_INDOOR_IND")
        if sg_ind_out: selected_thresholds.append("GAS_OUTDOOR_IND")

    # ── Groundwater ───────────────────────────────────────────────
    if has_gw:
        any_shown = True
        st.markdown("##### 💧 מי תהום")
        use_gw = st.checkbox('ערך סף מי"ת (GW Standard)', value=True, key="gw_cb")
        if use_gw: selected_thresholds.append("GW")

    if not any_shown:
        st.info("ℹ️ LOWFLOW — ממצאי שדה בלבד, ללא ערכי סף")
    elif not selected_thresholds:
        st.warning("⚠️ לא נבחרו ערכי סף — הדוח ייצא ללא עמודות השוואה")

    # ── Combine options ───────────────────────────────────────────
    has_tph_and_voc   = "SOIL_TPH" in found_atypes and "SOIL_VOC"   in found_atypes
    has_tph_and_mbtex = "SOIL_TPH" in found_atypes and "SOIL_MBTEX" in found_atypes
    combine_tph_voc   = False
    combine_tph_mbtex = False

    if has_tph_and_voc or has_tph_and_mbtex:
        st.markdown('<div style="margin-top:0.5rem;"></div>', unsafe_allow_html=True)
        cc1, cc2 = st.columns(2)
        with cc1:
            if has_tph_and_voc:
                combine_tph_voc = st.checkbox("שלב TPH + BTEX בגיליון אחד", value=False, key="combine_tph_voc")
        with cc2:
            if has_tph_and_mbtex:
                combine_tph_mbtex = st.checkbox("שלב TPH + MBTEX בגיליון אחד", value=False, key="combine_tph_mbtex")

    st.markdown('</div>', unsafe_allow_html=True)  # end section-card

    # ══════════════════════════════════════════════════════════════
    # PREVIEW TABLE
    # ══════════════════════════════════════════════════════════════
    with st.expander("📊 תצוגה מקדימה של הנתונים", expanded=False):
        def build_preview(recs):
            rows = []
            for r in recs:
                val = r.get('value')
                rows.append({
                    'דגימה':   r.get('sample_id', ''),
                    'תרכובת':  r.get('compound', ''),
                    'CAS':      r.get('cas', ''),
                    'ערך':      f"{val:.4g}" if isinstance(val, float) else (str(val) if val is not None else ''),
                    'יחידות':  r.get('unit', ''),
                    'flag':     r.get('flag', ''),
                })
            return pd.DataFrame(rows)

        analysis_types = list(by_type.keys())
        if len(analysis_types) > 1:
            tabs = st.tabs([f"{t} ({by_type[t]})" for t in analysis_types])
            for tab, atype in zip(tabs, analysis_types):
                with tab:
                    subset = [r for r in records if r.get('analysis_type') == atype]
                    st.dataframe(build_preview(subset), use_container_width=True, height=280)
        else:
            st.dataframe(build_preview(records), use_container_width=True, height=320)

    # ══════════════════════════════════════════════════════════════
    # BUILD EXCEL + DOWNLOAD
    # ══════════════════════════════════════════════════════════════
    _steps(3)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📥 שלב 3 — הורדת דוח Excel</div>', unsafe_allow_html=True)

    _sniff   = raw_bytes.lstrip()[:200]
    _is_kte_gw = (
        lab == "KTE" and category == "groundwater" and
        (b"<?xml" in _sniff or b"<Workbook" in _sniff)
    )

    excel_buf = io.BytesIO()
    excel_ok  = False
    word_buf  = io.BytesIO()
    word_ok   = False

    if _is_kte_gw:
        try:
            from core.excel_output import build_kte_gw_btex_simple_from_xml
            build_kte_gw_btex_simple_from_xml(raw_bytes, excel_buf)
            excel_ok = True
        except Exception as e:
            st.error(f"שגיאת בניית Excel: {e}")
            st.exception(e)
    elif lab.lower() in ("xrf",) or lab in ("אלכם", "אלכם (XRF)"):
        try:
            build_xrf_simple_excel(
                records,
                excel_buf,
                threshold_manager   = tm,
                selected_thresholds = selected_thresholds,
                project_name        = project_name,
                client              = client_name,
                report_date         = date.today().strftime('%d.%m.%Y'),
            )
            excel_buf.seek(0)
            excel_ok = True
        except Exception as e:
            st.error(f"שגיאת בניית Excel: {e}")
            st.exception(e)
    else:
        thresh_display = ", ".join(tm.threshold_label(k) for k in selected_thresholds) or "ללא ערכי סף"
        st.caption(f"📌 ערכי סף: **{thresh_display}**")
        try:
            builder = LabReportExcel(
                records             = records,
                threshold_manager   = tm,
                output_path         = excel_buf,
                project_name        = project_name,
                client              = client_name,
                report_date         = date.today().strftime('%d.%m.%Y'),
                selected_thresholds = selected_thresholds,
                combine_tph_voc     = combine_tph_voc,
                combine_tph_mbtex   = combine_tph_mbtex,
            )
            builder.build()
            excel_buf.seek(0)
            excel_ok = True
            try:
                LabReportWord(
                    records             = records,
                    threshold_manager   = tm,
                    output_path         = word_buf,
                    project_name        = project_name,
                    client              = client_name,
                    report_date         = date.today().strftime('%d.%m.%Y'),
                    selected_thresholds = selected_thresholds,
                    combine_tph_voc     = combine_tph_voc,
                    combine_tph_mbtex   = combine_tph_mbtex,
                ).build()
                word_buf.seek(0)
                word_ok = True
            except Exception as e:
                st.warning(f"⚠️ שגיאת בניית Word: {e}")
        except Exception as e:
            st.error(f"שגיאת בניית Excel: {e}")
            st.exception(e)

    if excel_ok:
        def _safe(s: str) -> str:
            import re
            return re.sub(r'[\\/*?:"<>|\s]+', '_', s.strip()).strip('_') or 'x'
        _parts = ["lab_report"]
        if client_name.strip():  _parts.append(_safe(client_name))
        if project_name.strip(): _parts.append(_safe(project_name))
        out_filename = f"{'_'.join(_parts)}.xlsx"
        size_kb = len(excel_buf.getvalue()) / 1024

        dl_col, wd_col, prev_col, info_col = st.columns([2, 2, 1.5, 1])
        with dl_col:
            st.download_button(
                label     = "⬇️ הורד דוח Excel",
                data      = excel_buf.getvalue(),
                file_name = out_filename,
                mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with wd_col:
            if word_ok:
                st.download_button(
                    label     = "⬇️ הורד דוח Word",
                    data      = word_buf.getvalue(),
                    file_name = out_filename.replace(".xlsx", ".docx"),
                    mime      = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        with prev_col:
            if st.button("👁️ תצוגה מקדימה", use_container_width=True, key="xl_preview_btn"):
                st.session_state["xl_show_preview"] = not st.session_state.get("xl_show_preview", False)
        with info_col:
            st.markdown(f"""
            <div style="padding:0.5rem 0;font-size:0.82rem;color:#64748b;direction:rtl;">
              <div>📄 <b>{out_filename}</b></div>
              <div>📦 גודל: {size_kb:.1f} KB</div>
              <div>📅 {date.today().strftime('%d.%m.%Y')}</div>
            </div>
            """, unsafe_allow_html=True)

        if st.session_state.get("xl_show_preview", False):
            with st.expander("📊 תצוגה מקדימה — נתונים מורחבת", expanded=True):
                _preview_atypes = list(by_type.keys())
                if len(_preview_atypes) > 1:
                    _preview_tabs = st.tabs([f"{t} ({by_type[t]})" for t in _preview_atypes])
                    for _ptab, _patype in zip(_preview_tabs, _preview_atypes):
                        with _ptab:
                            _subset = [r for r in records if r.get('analysis_type') == _patype]
                            _styled, _has_clr = _build_styled_pivot(
                                _subset, tm, selected_thresholds)
                            st.dataframe(_styled, use_container_width=True)
                            if _has_clr:
                                st.markdown(_PREVIEW_LEGEND, unsafe_allow_html=True)
                else:
                    _styled, _has_clr = _build_styled_pivot(
                        records, tm, selected_thresholds)
                    st.dataframe(_styled, use_container_width=True)
                    if _has_clr:
                        st.markdown(_PREVIEW_LEGEND, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)  # end section-card

    # ── footer ────────────────────────────────────────────────────
    st.markdown(
        f'<div style="text-align:center;color:#94a3b8;font-size:0.75rem;margin-top:1rem;">'
        f'🔬 {lab} / {category} &nbsp;·&nbsp; '
        f'📁 {fname[:80]}{"…" if len(fname)>80 else ""} &nbsp;·&nbsp; '
        f'📅 {date.today().strftime("%d.%m.%Y")}'
        f'</div>',
        unsafe_allow_html=True,
    )
