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
    logo_path = os.path.join(ROOT, 'logo.png')
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
html, body {{ direction: rtl; }}
.main .block-container {{
    direction: rtl;
    padding-top: 1rem;
    max-width: 1300px;
}}
[data-testid="stSidebar"] {{
    direction: rtl;
    background: #1e293b;
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
[data-testid="stSidebar"] hr {{ border-color: #334155; }}

/* hide footer / menu */
#MainMenu {{ visibility: hidden; }}
footer {{ visibility: hidden; }}
header {{ visibility: hidden; }}

/* ── typography ── */
.stMarkdown p, .stMarkdown li {{ direction: rtl; text-align: right; }}
.stTextInput label, .stSelectbox label,
.stMultiSelect label, .stFileUploader label,
.stCheckbox label, .stRadio label {{ direction: rtl; text-align: right; }}

/* ── hero header ── */
.hero {{
    background: linear-gradient(135deg, #1e3a8a 0%, #1d4ed8 60%, #0ea5e9 100%);
    color: white;
    padding: 1.5rem 2rem;
    border-radius: 16px;
    margin-bottom: 1.5rem;
    box-shadow: 0 4px 24px rgba(30,64,175,0.25);
    display: flex;
    align-items: center;
    justify-content: space-between;
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
.hero-badge {{
    background: rgba(255,255,255,0.15);
    border: 1px solid rgba(255,255,255,0.3);
    border-radius: 10px;
    padding: 0.5rem 1rem;
    font-size: 0.8rem;
    text-align: center;
    direction: ltr;
    min-width: 160px;
}}
.hero-badge a {{
    color: #bfdbfe;
    font-weight: 600;
    font-size: 0.85rem;
    word-break: break-all;
}}

/* ── section cards ── */
.section-card {{
    background: white;
    border-radius: 12px;
    border: 1px solid #e2e8f0;
    padding: 1.25rem 1.5rem;
    margin-bottom: 1.25rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}}
.section-title {{
    font-size: 1rem;
    font-weight: 700;
    color: #1e293b;
    margin: 0 0 0.75rem;
    padding-bottom: 0.5rem;
    border-bottom: 2px solid #e2e8f0;
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
    border-radius: 20px;
    padding: 0.3rem 0.9rem;
    font-size: 0.8rem;
    color: #64748b;
    font-weight: 500;
    flex: 1;
    justify-content: center;
}}
.step-pill.active {{
    background: #eff6ff;
    border-color: #3b82f6;
    color: #1d4ed8;
    font-weight: 700;
}}
.step-pill.done {{
    background: #f0fdf4;
    border-color: #22c55e;
    color: #15803d;
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
    border-radius: 10px;
    padding: 0.9rem 1rem;
    text-align: center;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}}
.stat-num {{
    font-size: 1.75rem;
    font-weight: 800;
    color: #1e40af;
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
    border: 2.5px dashed #93c5fd;
    border-radius: 12px;
    padding: 0.5rem;
    background: #eff6ff;
    transition: border-color 0.2s;
}}
[data-testid="stFileUploader"]:hover {{
    border-color: #3b82f6;
    background: #dbeafe;
}}
[data-testid="stFileUploader"] label {{
    font-size: 0.95rem !important;
    font-weight: 600 !important;
    color: #1d4ed8 !important;
}}

/* ── download button ── */
.stDownloadButton button {{
    width: 100%;
    background: linear-gradient(135deg, #16a34a, #15803d);
    color: white;
    font-weight: 700;
    font-size: 1rem;
    border-radius: 10px;
    border: none;
    padding: 0.75rem;
    letter-spacing: 0.3px;
    box-shadow: 0 2px 8px rgba(22,163,74,0.3);
}}
.stDownloadButton button:hover {{
    background: linear-gradient(135deg, #15803d, #166534);
    box-shadow: 0 4px 12px rgba(22,163,74,0.4);
}}

/* ── type badge ── */
.type-badge {{
    display: inline-block;
    padding: 3px 12px;
    border-radius: 20px;
    font-size: 0.78rem;
    font-weight: 600;
    margin: 2px 3px;
    color: white;
}}

/* ── info banner ── */
.info-banner {{
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    border-radius: 10px;
    padding: 0.75rem 1rem;
    direction: rtl;
    font-size: 0.9rem;
    color: #1e40af;
    margin-bottom: 0.75rem;
}}
.success-banner {{
    background: #f0fdf4;
    border: 1px solid #86efac;
    border-radius: 10px;
    padding: 0.75rem 1rem;
    direction: rtl;
    font-size: 0.9rem;
    color: #15803d;
    margin-bottom: 0.75rem;
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
    border-bottom: 1px solid #334155;
}}

/* ── metric overrides ── */
[data-testid="metric-container"] {{
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 0.75rem 1rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}}
[data-testid="stMetricValue"] {{
    font-size: 1.6rem !important;
    color: #1e40af !important;
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
    PFAS_THRESH  = os.path.join(LAB_DIR, 'נספח לטבלת ערכי סף - PFAS.xlsx')
    vsl_full_path = VSL_FULL    if os.path.exists(VSL_FULL)    else None
    pfas_path     = PFAS_THRESH if os.path.exists(PFAS_THRESH) else None
    return ThresholdManager(MAIN_THRESH, pfas_path=pfas_path, vsl_full_path=vsl_full_path)

try:
    from core.excel_output import LabReportExcel
    from parsers import get_parser, auto_detect_category
    _tm_py    = os.path.join(TOOL_DIR, 'core', 'threshold_manager.py')
    _tm_mtime = os.path.getmtime(_tm_py)
    tm = load_threshold_manager(_tm_mtime)
except Exception as e:
    st.error(f"שגיאת טעינת מודולים: {e}")
    st.stop()

# ══════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(
        f'<div style="text-align:center;padding:0.5rem 0 0.75rem;">'
        f'<div style="background:white;border-radius:10px;padding:0.6rem 0.8rem;'
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
    lab = st.selectbox("מעבדה", ["KTE", "מכון הנפט", "בקטוכם", "Alchem"],
                       label_visibility="collapsed")
    category_display = {
        "🔍 זיהוי אוטומטי":           "auto",
        "🪨 קרקע (soil)":             "soil",
        "💧 מי תהום (groundwater)":   "groundwater",
        "🧬 PFAS":                    "pfas",
        "📊 PR format (KTE מתכות)":   "pr",
        "💨 גז קרקע (soil_gas)":      "soil_gas",
    }
    cat_label    = st.selectbox("קטגוריה", list(category_display.keys()),
                                label_visibility="collapsed")
    category_raw = category_display[cat_label]


    st.markdown('<hr style="margin:1rem 0 0.5rem;">', unsafe_allow_html=True)

    # ── Network share box ─────────────────────────────────────────
    st.markdown(
        f'<div style="background:#0f172a;border:1px solid #1e40af;border-radius:8px;'
        f'padding:0.6rem 0.75rem;margin-top:0.5rem;">'
        f'<div style="font-size:0.7rem;color:#93c5fd;font-weight:700;margin-bottom:4px;">'
        f'🔗 קישור לשיתוף עם הצוות</div>'
        f'<div style="font-size:0.75rem;color:#bfdbfe;word-break:break-all;'
        f'font-family:monospace;">{APP_URL}</div>'
        f'<div style="font-size:0.65rem;color:#64748b;margin-top:4px;">'
        f'זמין ברשת המקומית בלבד</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

# ══════════════════════════════════════════════════════════════════
# HERO HEADER
# ══════════════════════════════════════════════════════════════════
_hero_logo = (
    f'<div style="background:white;border-radius:8px;padding:0.4rem 0.7rem;">'
    f'<img src="data:image/png;base64,{LOGO_B64}" style="height:48px;display:block;"></div>'
    if LOGO_B64 else ''
)
st.markdown(f"""
<div class="hero">
  <div style="display:flex;align-items:center;gap:1rem;">
    {_hero_logo}
    <div>
      <div class="hero-title">מערכת ניתוח תוצאות מעבדה</div>
      <div class="hero-sub">העלה קובץ דוח מעבדה · בחר ערכי סף · הורד Excel מסודר</div>
    </div>
  </div>
  <div class="hero-badge">
    <div style="font-size:0.65rem;margin-bottom:4px;opacity:0.7;">קישור לצוות</div>
    <a href="{APP_URL}" target="_blank">{APP_URL}</a>
    <div style="font-size:0.65rem;margin-top:4px;opacity:0.6;">רשת מקומית · פורט 8501</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# STEP INDICATOR
# ══════════════════════════════════════════════════════════════════
def _steps(step: int):
    labels = ["① העלאת קובץ", "② בחירת ערכי סף", "③ הורדת דוח"]
    pills = ""
    for i, lbl in enumerate(labels, 1):
        cls = "active" if i == step else ("done" if i < step else "step-pill")
        icon = "✅ " if i < step else ""
        pills += f'<div class="step-pill {cls}">{icon}{lbl}</div>'
    st.markdown(f'<div class="step-row">{pills}</div>', unsafe_allow_html=True)

_steps(1)

# ══════════════════════════════════════════════════════════════════
# UPLOAD
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">📤 שלב 1 — העלאת קובץ דוח מעבדה</div>',
            unsafe_allow_html=True)

col_up, col_meta = st.columns([3, 1])

with col_up:
    uploaded_files = st.file_uploader(
        "גרור קבצים לכאן או לחץ לבחירה",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="ניתן להעלות מספר קבצים מאותה מעבדה | XLSX / XLS / CSV",
        label_visibility="visible",
    )

with col_meta:
    # selected lab/category display
    cat_clean = cat_label.split(" ", 1)[-1] if " " in cat_label else cat_label
    st.markdown(f"""
    <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;
                padding:0.9rem 1rem;margin-top:1.75rem;text-align:center;">
      <div style="font-size:0.7rem;color:#94a3b8;font-weight:600;margin-bottom:4px;">מעבדה</div>
      <div style="font-size:1.3rem;font-weight:800;color:#1e40af;">{lab}</div>
      <div style="font-size:0.75rem;color:#64748b;margin-top:4px;">{cat_clean}</div>
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

# ══════════════════════════════════════════════════════════════════
# PARSE
# ══════════════════════════════════════════════════════════════════
all_raw: list[tuple[str, bytes]] = [(uf.name, uf.read()) for uf in uploaded_files]
fname     = " | ".join(f for f, _ in all_raw)
raw_bytes = all_raw[0][1]

if category_raw == 'auto':
    category = auto_detect_category(all_raw[0][0], raw_bytes)
    cat_info  = f"זוהה אוטומטית: **{category}**"
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

with st.spinner(f"מנתח {'קבצים' if n_files > 1 else 'קובץ'}..."):
    for fname_i, raw_i in all_raw:
        try:
            file_records = parser.parse(io.BytesIO(raw_i))
            all_records.extend(file_records)
            file_summaries.append({"name": fname_i, "records": len(file_records), "ok": True})
        except Exception as e:
            st.error(f"שגיאת פרסינג: {fname_i} — {e}")
            file_summaries.append({"name": fname_i, "records": 0, "ok": False})

records = all_records

if not records:
    st.warning("⚠️ לא נמצאו רשומות — בדוק פורמט הקובץ ובחירת מעבדה / קטגוריה")
    st.stop()

# ── stats ─────────────────────────────────────────────────────────
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
    "GW_VOC":       "#2563eb", "GW_PFAS":     "#9333ea",
    "LOWFLOW":      "#6b7280",
}
badges = " ".join(
    f'<span class="type-badge" style="background:{BADGE_COLORS.get(t,"#94a3b8")};">'
    f'{t}: {cnt}</span>'
    for t, cnt in by_type.most_common()
)
st.markdown(f'<div style="margin:0.5rem 0;">{badges}</div>', unsafe_allow_html=True)

_steps(2)

# ══════════════════════════════════════════════════════════════════
# THRESHOLD SELECTION
# ══════════════════════════════════════════════════════════════════
found_atypes  = list(by_type.keys())
has_soil      = any(t in found_atypes for t in ("SOIL_VOC","SOIL_TPH","SOIL_METALS","SOIL_MBTEX"))
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

any_shown = False

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">📋 שלב 2 — בחירת ערכי סף להשוואה</div>',
            unsafe_allow_html=True)

# ── Soil ─────────────────────────────────────────────────────────
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

# ── Soil PFAS ─────────────────────────────────────────────────────
if has_soil_pfas:
    any_shown = True
    st.markdown("##### 🧬 קרקע PFAS")
    cp1, cp2, cp3 = st.columns(3)
    with cp1: use_pfas_vsl    = st.checkbox("PFAS VSL",           value=True,  key="pfas_vsl")
    with cp2: use_pfas_t1_res = st.checkbox("PFAS Tier 1 מגורים", value=False, key="pfas_t1r")
    with cp3: use_pfas_t1_ind = st.checkbox("PFAS Tier 1 תעשייה", value=False, key="pfas_t1i")
    if use_pfas_vsl:    selected_thresholds.append("PFAS_VSL")
    if use_pfas_t1_res: selected_thresholds.append("PFAS_TIER1_RES")
    if use_pfas_t1_ind: selected_thresholds.append("PFAS_TIER1_IND")

# ── Soil gas ──────────────────────────────────────────────────────
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

# ── Groundwater ───────────────────────────────────────────────────
if has_gw:
    any_shown = True
    st.markdown("##### 💧 מי תהום")
    use_gw = st.checkbox('ערך סף מי"ת (GW Standard)', value=True, key="gw_cb")
    if use_gw: selected_thresholds.append("GW")

if not any_shown:
    st.info("ℹ️ LOWFLOW — ממצאי שדה בלבד, ללא ערכי סף")
elif not selected_thresholds:
    st.warning("⚠️ לא נבחרו ערכי סף — הדוח ייצא ללא עמודות השוואה")

# ── Combine options ───────────────────────────────────────────────
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

# ══════════════════════════════════════════════════════════════════
# PREVIEW TABLE
# ══════════════════════════════════════════════════════════════════
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

# ══════════════════════════════════════════════════════════════════
# BUILD EXCEL + DOWNLOAD
# ══════════════════════════════════════════════════════════════════
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

if _is_kte_gw:
    try:
        from core.excel_output import build_kte_gw_btex_simple_from_xml
        build_kte_gw_btex_simple_from_xml(raw_bytes, excel_buf)
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
            selected_thresholds = selected_thresholds if selected_thresholds else None,
            combine_tph_voc     = combine_tph_voc,
            combine_tph_mbtex   = combine_tph_mbtex,
        )
        builder.build()
        excel_buf.seek(0)
        excel_ok = True
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

    dl_col, info_col = st.columns([2, 1])
    with dl_col:
        st.download_button(
            label     = f"⬇️ הורד דוח Excel",
            data      = excel_buf.getvalue(),
            file_name = out_filename,
            mime      = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with info_col:
        st.markdown(f"""
        <div style="padding:0.5rem 0;font-size:0.82rem;color:#64748b;direction:rtl;">
          <div>📄 <b>{out_filename}</b></div>
          <div>📦 גודל: {size_kb:.1f} KB</div>
          <div>📅 {date.today().strftime('%d.%m.%Y')}</div>
        </div>
        """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)  # end section-card

# ── footer ────────────────────────────────────────────────────────
st.markdown(
    f'<div style="text-align:center;color:#94a3b8;font-size:0.75rem;margin-top:1rem;">'
    f'🔬 {lab} / {category} &nbsp;·&nbsp; '
    f'📁 {fname[:80]}{"…" if len(fname)>80 else ""} &nbsp;·&nbsp; '
    f'📅 {date.today().strftime("%d.%m.%Y")}'
    f'</div>',
    unsafe_allow_html=True,
)
