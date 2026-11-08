import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date, datetime, time, timedelta
import io
import base64
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# إعدادات الصفحة الأساسية
# ==========================================
st.set_page_config(page_title="SNB Promises Dashboard", layout="wide", initial_sidebar_state="collapsed")

# ==========================================
# ثيم Streamlit الفاتح: بيخلّي كل الحقول والمربعات (القوائم، الاختيارات، الرفع) بيضاء
# ==========================================
_LIGHT_THEME = {
    "base": "light",
    "primaryColor": "#2CA549",
    "backgroundColor": "#FFFFFF",
    "secondaryBackgroundColor": "#FFFFFF",
    "textColor": "#1F4373",
    "borderColor": "#99AABF",
}
for _k, _v in _LIGHT_THEME.items():
    try:
        st._config.set_option(f"theme.{_k}", _v)
    except Exception:
        pass
if not st.session_state.get('_light_theme_applied'):
    st.session_state['_light_theme_applied'] = True
    st.rerun()

if 'current_page' not in st.session_state:
    st.session_state.current_page = 'Overview'
if 'theme_choice' not in st.session_state:
    st.session_state.theme_choice = 'ثيم إجادة (Navy)'
if 'act_uploader_ver' not in st.session_state:
    st.session_state.act_uploader_ver = 0

# ==========================================
# دالات استجابة سريعة للزراير (Callbacks) والأيقونات
# ==========================================
PHONE_ICON_B64 = "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"
PHONE_ICON_COLOR = "#1F4373"

# الأيقونات الجديدة باللون الأخضر للمكالمات الناجحة والمكالمات المغطاة
GREEN_CHECK_SVG = """<svg viewBox="0 0 100 100" width="32" height="32" xmlns="http://www.w3.org/2000/svg"><path fill="#2CA549" d="M50 4 C54 4 57 8 61 9 C65 10 70 8 73 10 C77 13 77 18 80 21 C83 24 88 24 90 28 C92 31 90 36 91 40 C92 44 96 47 96 51 C96 55 92 58 91 62 C90 66 92 71 90 74 C88 78 83 78 80 81 C77 84 77 89 73 92 C70 94 65 92 61 93 C57 94 54 98 50 98 C46 98 43 94 39 93 C35 92 30 94 27 92 C23 89 23 84 20 81 C17 78 12 78 10 74 C8 71 10 66 9 62 C8 58 4 55 4 51 C4 47 8 44 9 40 C10 36 8 31 10 28 C12 24 17 24 20 21 C23 18 23 13 27 10 C30 8 35 10 39 9 C43 8 46 4 50 4 Z"/><circle cx="50" cy="51" r="28" fill="#82D197"/><path d="M38 51 L46 59 L63 42" fill="none" stroke="#FFFFFF" stroke-width="8" stroke-linecap="round" stroke-linejoin="round"/></svg>"""

GREEN_PHONE_X_SVG = """<svg viewBox="0 0 24 24" width="30" height="30" fill="none" stroke="#2CA549" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" xmlns="http://www.w3.org/2000/svg"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/><line x1="15" y1="3" x2="21" y2="9" stroke="#2CA549" stroke-width="2.6"/><line x1="21" y1="3" x2="15" y2="9" stroke="#2CA549" stroke-width="2.6"/></svg>"""

def toggle_theme():
    if st.session_state.theme_choice == 'ثيم إجادة (Navy)':
        st.session_state.theme_choice = 'التصميم الفاتح (Light)'
    else:
        st.session_state.theme_choice = 'ثيم إجادة (Navy)'

def set_page(page_name):
    st.session_state.current_page = page_name

def clear_activity_files():
    st.session_state.page_files['Activity'] = None
    st.session_state.act_uploader_ver += 1

# ==========================================
# دالة عامة للبحث عن الأعمدة
# ==========================================
def find_col_in(target_df, candidates):
    if target_df is None or target_df.empty:
        return None
    for c in candidates:
        for col in target_df.columns:
            if col.lower().strip() == c.lower().strip():
                return col
    return None

# ==========================================
# إعدادات الثيم والألوان
# ==========================================
ejada_navy_dark = "#1F4373"
ejada_navy_light = "#304D73"
ejada_grey_blue = "#586F8C"
ejada_border_light = "#99AABF"
white_bg = "#FFFFFF"  
snb_green = "#2CA549"  

if st.session_state.theme_choice == "ثيم إجادة (Navy)":
    bg_color = ejada_navy_dark 
    card_color = white_bg 
    border_color = ejada_border_light
    primary_color = ejada_navy_light 
    primary_text = "#FFFFFF"
    text_color = ejada_navy_dark
    label_color = "#FFFFFF"
    sub_text_color = "#E2E8F0"
    gauge_bg = "#E2E8F0"
    nav_btn_bg = "#FFFFFF"
    nav_btn_text = ejada_navy_dark
    nav_btn_border = ejada_border_light
    kpi_title_color = "#586F8C"
    header_logo_color = "#FFFFFF"
    input_bg = "#FFFFFF"
    input_text = "#1F4373"
else:
    bg_color = "#F4F6F9"
    card_color = "#FFFFFF"
    border_color = "#CBD5E1"
    primary_color = ejada_navy_dark
    primary_text = "#FFFFFF"
    text_color = "#1F4373" 
    label_color = "#1F4373"
    sub_text_color = "#475569"
    gauge_bg = "#E2E8F0"
    nav_btn_bg = "#FFFFFF"
    nav_btn_text = "#1F4373"
    nav_btn_border = "#CBD5E1"
    kpi_title_color = "#475569"
    header_logo_color = "#1F4373"
    input_bg = "#FFFFFF"
    input_text = "#1F4373"

# ==========================================
# تصميم الـ UI الديناميكي (CSS) وتبييض الشرائط السوداء بالكامل
# ==========================================
st.markdown(f"""
<style>
    [data-testid="collapsedControl"] {{display: none !important;}}
    [data-testid="stSidebar"] {{display: none !important;}}
    
    .stApp {{ background-color: {bg_color} !important; }}

    .block-container {{ padding-top: 2.5rem !important; padding-left: 1.5rem !important; padding-right: 1.5rem !important; max-width: 100% !important; direction: rtl; text-align: right; }}
    
    div[data-testid="stButton"] {{
        position: relative !important;
        z-index: 9999 !important;
    }}

    div[data-testid="stButton"] > button[kind="secondary"] {{
        background-color: {nav_btn_bg} !important;
        color: {nav_btn_text} !important;
        border: 1px solid {nav_btn_border} !important;
        border-radius: 8px !important;
        padding: 8px 16px !important;
        font-weight: 800 !important;
        font-size: 15px !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
    }}
    div[data-testid="stButton"] > button[kind="secondary"] p {{ color: {nav_btn_text} !important; font-weight: 800 !important; }}
    
    div[data-testid="stButton"] > button[kind="primary"] {{
        background-color: {snb_green} !important; 
        color: #FFFFFF !important; 
        border: none !important;
        border-radius: 8px !important;
        padding: 8px 16px !important;
        font-weight: 800 !important;
        font-size: 15px !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.3) !important;
    }}
    div[data-testid="stButton"] > button[kind="primary"] p {{ color: #FFFFFF !important; font-weight: 800 !important; }}

    div.st-key-theme_toggle div[data-testid="stButton"] > button {{
        background: linear-gradient(90deg, {ejada_navy_dark} 50%, {ejada_grey_blue} 50%) !important;
        border: 2px solid #FFFFFF !important;
        border-radius: 50% !important; 
        width: 40px !important;
        height: 40px !important;
        min-width: 40px !important;
        min-height: 40px !important;
        padding: 0 !important;
        color: transparent !important; 
        box-shadow: 0 2px 6px rgba(0,0,0,0.4) !important;
        cursor: pointer !important;
    }}

    /* تصغير حدود المربعات (Containers) لتكون على قد خانة رفع الملفات بالضبط */
    div[data-testid="stVerticalBlockBorderWrapper"] {{
        padding: 10px 14px !important;
        border-radius: 10px !important;
        border: 1px solid {border_color} !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"] > div {{
        gap: 0.4rem !important;
    }}

    /* زرار X الصغير لمسح الملف المرفوع */
    div.st-key-clear_act_files_btn div[data-testid="stButton"] > button {{
        background-color: #EF4444 !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 6px !important;
        padding: 2px 10px !important;
        min-height: 26px !important;
        height: 26px !important;
        font-size: 13px !important;
        font-weight: 900 !important;
        line-height: 1 !important;
    }}
    div.st-key-clear_act_files_btn div[data-testid="stButton"] > button p {{
        color: #FFFFFF !important;
        font-size: 13px !important;
        font-weight: 900 !important;
    }}

    /* ===== تبييض كل المربعات السوداء (Selectbox / Multiselect / Checkbox / Expander) ===== */
    html body .stApp div[data-baseweb="select"] > div,
    html body .stApp [data-testid="stSelectbox"] div[data-baseweb="select"] > div,
    html body .stApp [data-testid="stMultiSelect"] div[data-baseweb="select"] > div,
    html body .stApp div[data-baseweb="base-input"],
    html body .stApp div[data-baseweb="input"] {{
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
        color: {input_text} !important;
        border: 1px solid {border_color} !important;
        border-radius: 8px !important;
        direction: rtl !important;
        text-align: right !important;
    }}
    html body .stApp div[data-baseweb="select"] div,
    html body .stApp div[data-baseweb="select"] span,
    html body .stApp div[data-baseweb="select"] input {{
        color: {input_text} !important;
        -webkit-text-fill-color: {input_text} !important;
    }}
    html body .stApp div[data-baseweb="select"] input {{
        background-color: transparent !important;
        background: transparent !important;
    }}
    html body .stApp div[data-baseweb="select"] svg {{
        fill: {input_text} !important;
        color: {input_text} !important;
    }}

    /* القائمة المنسدلة لما تفتح */
    html body div[data-baseweb="popover"],
    html body div[data-baseweb="popover"] > div,
    html body div[data-baseweb="popover"] ul,
    html body div[data-baseweb="menu"],
    html body ul[role="listbox"],
    html body li[role="option"] {{
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
        color: {input_text} !important;
        direction: rtl !important;
        text-align: right !important;
    }}
    html body li[role="option"] *,
    html body ul[role="listbox"] * {{
        color: {input_text} !important;
        -webkit-text-fill-color: {input_text} !important;
    }}
    html body li[role="option"]:hover,
    html body li[role="option"][aria-selected="true"] {{
        background-color: #E2E8F0 !important;
    }}

    /* الاختيارات المحددة داخل Multiselect */
    html body .stApp span[data-baseweb="tag"] {{
        background-color: {ejada_navy_dark} !important;
        background: {ejada_navy_dark} !important;
        color: #FFFFFF !important;
        border-radius: 6px !important;
    }}
    html body .stApp span[data-baseweb="tag"] span,
    html body .stApp span[data-baseweb="tag"] svg {{
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        fill: #FFFFFF !important;
    }}

    /* مربع الـ Checkbox */
    html body .stApp label[data-baseweb="checkbox"] > span:first-child,
    html body .stApp label[data-baseweb="checkbox"] > div:first-child,
    html body .stApp div[data-baseweb="checkbox"] > span:first-child,
    html body .stApp div[data-baseweb="checkbox"] > div:first-child {{
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
        border: 2px solid {ejada_grey_blue} !important;
        border-radius: 4px !important;
    }}
    html body .stApp label[data-baseweb="checkbox"]:has(input:checked) > span:first-child,
    html body .stApp label[data-baseweb="checkbox"]:has(input:checked) > div:first-child {{
        background-color: {snb_green} !important;
        background: {snb_green} !important;
        border-color: {snb_green} !important;
    }}
    html body .stApp label[data-baseweb="checkbox"] svg {{
        fill: #FFFFFF !important;
        color: #FFFFFF !important;
    }}

    /* الـ Expander (تفصيل أداء المحصلين) */
    html body .stApp [data-testid="stExpander"] details,
    html body .stApp [data-testid="stExpander"] summary {{
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
        color: {input_text} !important;
        border-color: {border_color} !important;
        border-radius: 10px !important;
    }}
    html body .stApp [data-testid="stExpander"] summary *,
    html body .stApp [data-testid="stExpander"] summary p {{
        color: {input_text} !important;
        -webkit-text-fill-color: {input_text} !important;
    }}
    html body .stApp [data-testid="stExpander"] summary svg {{
        fill: {input_text} !important;
    }}

    /* التابات وخيارات الراديو */
    html body .stApp [role="tab"],
    html body .stApp [role="tab"] *,
    html body .stApp [role="radiogroup"] label,
    html body .stApp [role="radiogroup"] label * {{
        color: {label_color} !important;
    }}

    /* تبييض منطقة رفع الملفات وتصغير ارتفاعها وإظهار علامة X بوضوح */
    [data-testid="stFileUploader"] > section {{
        background-color: #FFFFFF !important;
        border: 2px dashed {border_color} !important;
        border-radius: 8px !important;
        padding: 8px 12px !important;
        min-height: unset !important;
        direction: rtl !important;
    }}
    [data-testid="stFileUploader"] section * {{
        color: {input_text} !important;
    }}
    [data-testid="stFileUploader"] button {{
        background-color: #FFFFFF !important;
        color: {input_text} !important;
        border: 2px solid {ejada_navy_dark} !important;
        border-radius: 8px !important;
        font-weight: 800 !important;
    }}
    [data-testid="stFileUploader"] button:hover {{
        background-color: {ejada_navy_dark} !important;
        color: #FFFFFF !important;
    }}
    [data-testid="stFileUploaderFile"] {{
        background-color: #FFFFFF !important;
        border: 1px solid {border_color} !important;
        border-radius: 6px !important;
        padding: 4px 8px !important;
        margin-top: 4px !important;
    }}
    [data-testid="stFileUploaderFile"] * {{
        color: {label_color} !important;
    }}
    [data-testid="stFileUploaderDeleteBtn"] button {{
        background-color: #EF4444 !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 50% !important;
        width: 22px !important;
        height: 22px !important;
        min-width: 22px !important;
        min-height: 22px !important;
        padding: 0 !important;
    }}
    [data-testid="stFileUploaderDeleteBtn"] svg {{
        fill: #FFFFFF !important;
        color: #FFFFFF !important;
    }}

    /* عناوين الفلاتر والـ Checkboxes */
    label[data-testid="stWidgetLabel"] p,
    div[data-testid="stCheckbox"] p {{
        color: {label_color} !important;
        font-weight: 800 !important;
        font-size: 14px !important;
    }}

    /* تصغير بطاقات الـ KPI وتوسيط الأسماء والأرقام في المنتصف */
    .kpi-card {{
        background-color: {card_color};
        border-radius: 10px;
        padding: 12px 14px;
        box-shadow: 0 3px 8px rgba(0,0,0,0.12);
        border: 1px solid {border_color};
        border-right-width: 5px;
        border-right-style: solid;
        text-align: center;
        direction: rtl;
        position: relative;
    }}
    .kpi-title {{
        color: {kpi_title_color};
        font-size: 13px;
        font-weight: 800;
        margin-bottom: 4px;
        text-align: center;
    }}
    .kpi-value {{
        color: {text_color};
        font-size: 24px;
        font-weight: 900;
        margin-bottom: 2px;
        text-align: center;
    }}
    .kpi-sub {{
        color: {kpi_title_color};
        font-size: 11px;
        font-weight: 700;
        text-align: center;
    }}
    .kpi-icon {{
        position: absolute;
        top: 12px;
        left: 12px;
        font-size: 22px;
        opacity: 0.95;
        display: flex;
        align-items: center;
        justify-content: center;
    }}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 0. الصف العلوي: زر القمر + Overview
# ==========================================
top_cols = st.columns([0.6, 1.6, 7.8], gap="small")

with top_cols[0]:
    st.button("🌓", key="theme_toggle", on_click=toggle_theme, help="تغيير الثيم (كحلي / أبيض)")

with top_cols[1]:
    st.button("📊 Overview", key="top_overview_btn", type="primary" if st.session_state.current_page == 'Overview' else "secondary", use_container_width=True, on_click=set_page, args=('Overview',))

# ==========================================
# 1. الهيدر الرئيسي
# ==========================================
st.markdown(f"""
<div style="position: relative; width: 100%; min-height: 110px; margin-top: -45px; margin-bottom: 20px; pointer-events: none; z-index: 1;">
    <div style="text-align: center; line-height: 1.1;">
        <div style="color: {header_logo_color}; font-size: 45px; font-weight: 900; font-family: Tahoma, sans-serif;">إجـــادة</div>
        <div style="color: {header_logo_color}; font-size: 20px; letter-spacing: 10px; font-weight: 900; margin-top: -8px; margin-right:-10px;">EJADA</div>
        <div style="color: {header_logo_color}; font-size: 18px; font-weight: bold; margin-top: 5px;">لتحصيل الديون</div>
    </div>
    <div style="position: absolute; top: 5px; left: 10px; text-align: left; line-height: 1.2; direction: ltr;">
        <div style="white-space: nowrap;">
            <span style="color: {snb_green}; font-weight: 900; font-size: 24px; letter-spacing: 1px; vertical-align: middle;">SNB</span>
            <span style="color: {snb_green}; font-size: 19px; font-weight: bold; font-family: Tahoma, sans-serif; margin-left: 6px; vertical-align: middle;">البنك الأهلي السعودي</span>
        </div>
        <div style="color: {snb_green}; font-size: 15px; font-weight: 900; text-align: center; letter-spacing: 1.5px; margin-top: 2px;">
            (30-60-NPL)
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 2. شريط التنقل
# ==========================================
nav_cols = st.columns([1.1, 1.6, 1.3, 1.3, 4.7], gap="small") 

with nav_cols[0]:
    st.button("🏠 الرئيسية", key="nav_home_btn", type="primary" if st.session_state.current_page == 'Overview' else "secondary", use_container_width=True, on_click=set_page, args=('Overview',))

with nav_cols[1]:
    st.button("🤝 الوعود القائمة والمكسورة", key="nav_prom_btn", type="primary" if st.session_state.current_page == 'Promises' else "secondary", use_container_width=True, on_click=set_page, args=('Promises',))

with nav_cols[2]:
    st.button("🚨 تقرير الإهمال", key="nav_neg_btn", type="primary" if st.session_state.current_page == 'Neglect' else "secondary", use_container_width=True, on_click=set_page, args=('Neglect',))

with nav_cols[3]:
    st.button("📈 النشاط", key="nav_act_btn", type="primary" if st.session_state.current_page == 'Activity' else "secondary", use_container_width=True, on_click=set_page, args=('Activity',))

with nav_cols[4]:
    st.markdown("")

st.markdown(f"<hr style='margin: 10px 0 20px 0; border-color: {border_color}; clear: both;'>", unsafe_allow_html=True)

# ==========================================
# دوال تقرير الوعود والإكسيل
# ==========================================
PIVOT_TEAM = 'Sales Team'
PIVOT_COLL = 'Salesperson'
PIVOT_NOTE = 'ملاحظات-التمويل'
PIVOT_ACC = 'Account Number'
PIVOT_AMT = 'Net Amount'
FIXED_NOTES = ['NPL', 'Dpd60', 'Dpd30']
SUM_FMT = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'


def build_pivot(df, order='desc'):
    need = [PIVOT_TEAM, PIVOT_COLL, PIVOT_NOTE, PIVOT_ACC, PIVOT_AMT]
    if df is None or df.empty or any(c not in df.columns for c in need):
        return {'notes': list(FIXED_NOTES), 'teams': [], 'grand': {'by_note': {n: (0, 0.0) for n in FIXED_NOTES}, 'cnt': 0, 'amt': 0.0}}

    d = df[need].copy()
    d['_c'] = d[PIVOT_ACC].notna().astype(int)
    d[PIVOT_AMT] = pd.to_numeric(d[PIVOT_AMT], errors='coerce').fillna(0.0)
    d[PIVOT_TEAM] = d[PIVOT_TEAM].fillna('(فارغ)').astype(str)
    d[PIVOT_COLL] = d[PIVOT_COLL].fillna('(فارغ)').astype(str)
    d[PIVOT_NOTE] = d[PIVOT_NOTE].fillna('(فارغ)').astype(str).str.strip()

    def _ordered(values):
        vals = list(pd.unique(values))
        return sorted(vals, reverse=True) if order == 'desc' else vals

    notes = list(FIXED_NOTES) + sorted([x for x in d[PIVOT_NOTE].unique() if x not in FIXED_NOTES], reverse=True)
    g = d.groupby([PIVOT_TEAM, PIVOT_COLL, PIVOT_NOTE]).agg(cnt=('_c', 'sum'), amt=(PIVOT_AMT, 'sum'))

    teams = []
    grand_by_note = {n: [0, 0.0] for n in notes}
    for team in _ordered(d[PIVOT_TEAM]):
        rows = []
        for coll in _ordered(d.loc[d[PIVOT_TEAM] == team, PIVOT_COLL]):
            by_note, t_cnt, t_amt = {}, 0, 0.0
            for n in notes:
                if (team, coll, n) in g.index:
                    c, a = int(g.loc[(team, coll, n), 'cnt']), float(g.loc[(team, coll, n), 'amt'])
                else:
                    c, a = 0, 0.0
                by_note[n] = (c, a)
                t_cnt += c
                t_amt += a
                grand_by_note[n][0] += c
                grand_by_note[n][1] += a
            rows.append({'coll': coll, 'by_note': by_note, 'cnt': t_cnt, 'amt': t_amt})
        teams.append({'team': team, 'rows': rows})

    grand = {
        'by_note': {n: (v[0], v[1]) for n, v in grand_by_note.items()},
        'cnt': sum(v[0] for v in grand_by_note.values()),
        'amt': sum(v[1] for v in grand_by_note.values()),
    }
    return {'notes': notes, 'teams': teams, 'grand': grand}


def pivot_to_frame(p):
    cols = ['اسم المشرف', 'اسم المحصل']
    for n in p['notes']:
        cols += [f'{n} | عدد الحسابات', f'{n} | متبقي المديونية']
    cols += ['Total عدد الحسابات', 'Total متبقي المديونية']
    out = []
    for t in p['teams']:
        for i, r in enumerate(t['rows']):
            row = [t['team'] if i == 0 else '', r['coll']]
            for n in p['notes']:
                row += [r['by_note'][n][0], round(r['by_note'][n][1], 2)]
            row += [r['cnt'], round(r['amt'], 2)]
            out.append(row)
    g = p['grand']
    row = ['الاجمالي', '']
    for n in p['notes']:
        row += [g['by_note'][n][0], round(g['by_note'][n][1], 2)]
    row += [g['cnt'], round(g['amt'], 2)]
    out.append(row)
    return pd.DataFrame(out, columns=cols)


def _prepare_for_output(df):
    d = df.copy()
    for c in ['Amount', 'Net Amount', 'Payment']:
        if c in d.columns:
            d[c] = pd.to_numeric(d[c], errors='coerce')
    for c in ['Debit Date', 'Assign Date', 'Follow up Due Date', 'Follow up Last Date', 'Expected Closing']:
        if c in d.columns:
            d[c] = pd.to_datetime(d[c], errors='coerce')
    return d


def _write_df_sheet(ws, df):
    if df is None or len(df.columns) == 0:
        return
    df = _prepare_for_output(df)
    bold = Font(bold=True)
    for j, c in enumerate(df.columns, 1):
        ws.cell(row=1, column=j, value=c).font = bold
        ws.column_dimensions[get_column_letter(j)].width = min(40, max(12, len(str(c)) + 4))
    for i, row in enumerate(df.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            if v is None or (not isinstance(v, (list, tuple)) and pd.isna(v)):
                continue
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            elif hasattr(v, 'item') and not isinstance(v, (str, bytes)):
                v = v.item()
            cell = ws.cell(row=i, column=j, value=v)
            if isinstance(v, datetime):
                cell.number_format = 'yyyy-mm-dd'
            elif isinstance(v, float):
                cell.number_format = '#,##0.00'


def _write_pivot_sheet(ws, p, title):
    notes = p['notes']
    n = len(notes)
    last_col = 2 + 2 * n + 2
    ws.sheet_view.rightToLeft = True

    dark = PatternFill('solid', fgColor='244062')
    light = PatternFill('solid', fgColor='C5D9F1')
    white_b = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    thin, thick = Side(style='thin'), Side(style='thick')
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    def hdr(r, c, v=None, size=11):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name='Calibri', bold=True, color='FFFFFF', size=size)
        cell.fill = dark
        cell.alignment = center
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        return cell

    for r in (1, 2):
        for c in range(1, last_col + 1):
            hdr(r, c, None)
    ws.cell(row=1, column=1, value=title).font = Font(name='Calibri', bold=True, color='FFFFFF', size=18)
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=last_col)

    for c in range(1, last_col + 1):
        hdr(3, c)
        hdr(4, c)
        hdr(5, c)
    ws.cell(row=3, column=3, value=PIVOT_NOTE)
    ws.cell(row=3, column=4, value='Values')
    for i, note in enumerate(notes):
        c0 = 3 + 2 * i
        ws.cell(row=4, column=c0, value=note).font = Font(name='Calibri', bold=True, color='FFFFFF', size=14)
        ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + 1)
        ws.cell(row=5, column=c0, value='عدد الحسابات')
        ws.cell(row=5, column=c0 + 1, value='متبقي المديونية')
    tc = 3 + 2 * n
    ws.cell(row=4, column=tc, value='Total عدد الحسابات')
    ws.merge_cells(start_row=4, start_column=tc, end_row=5, end_column=tc)
    ws.cell(row=4, column=tc + 1, value='Total متبقي المديونية')
    ws.merge_cells(start_row=4, start_column=tc + 1, end_row=5, end_column=tc + 1)
    ws.cell(row=5, column=1, value='اسم المشرف ')
    ws.cell(row=5, column=2, value='اسم المحصل ')

    def body(r, c, v, fmt=None, bold=False):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name='Calibri', bold=bold, size=11)
        cell.alignment = center
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        if fmt:
            cell.number_format = fmt
        return cell

    r = 6
    for t in p['teams']:
        start = r
        for row in t['rows']:
            body(r, 2, row['coll'])
            for i, note in enumerate(notes):
                body(r, 3 + 2 * i, row['by_note'][note][0])
                body(r, 4 + 2 * i, row['by_note'][note][1], SUM_FMT)
            body(r, tc, row['cnt'])
            body(r, tc + 1, row['amt'], SUM_FMT)
            r += 1
        a = body(start, 1, t['team'])
        a.fill = light
        a.border = Border(left=thick, right=thin, top=thin, bottom=thin)
        for rr in range(start + 1, r):
            cell = ws.cell(row=rr, column=1)
            cell.fill = light
            cell.border = Border(left=thick, right=thin, top=thin, bottom=thin)
        if r - 1 > start:
            ws.merge_cells(start_row=start, start_column=1, end_row=r - 1, end_column=1)

    g = p['grand']
    for c in range(1, last_col + 1):
        cell = ws.cell(row=r, column=c)
        cell.font = white_b
        cell.fill = dark
        cell.alignment = center
        cell.border = Border(left=thick if c == 1 else thin, right=thin, top=thin, bottom=thick)
    ws.cell(row=r, column=1, value='الاجمالي')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    for i, note in enumerate(notes):
        ws.cell(row=r, column=3 + 2 * i, value=g['by_note'][note][0])
        x = ws.cell(row=r, column=4 + 2 * i, value=g['by_note'][note][1])
        x.number_format = SUM_FMT
    ws.cell(row=r, column=tc, value=g['cnt'])
    x = ws.cell(row=r, column=tc + 1, value=g['amt'])
    x.number_format = SUM_FMT

    ws.column_dimensions['A'].width = 35.7
    ws.column_dimensions['B'].width = 30.7
    for c in range(3, 3 + 2 * n):
        ws.column_dimensions[get_column_letter(c)].width = 18.9
    ws.column_dimensions[get_column_letter(tc)].width = 16.7
    ws.column_dimensions[get_column_letter(tc + 1)].width = 17.6


import re
import zipfile
from xml.sax.saxutils import escape as _xml_escape

_ILLEGAL_XML = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _xa(v):
    return _xml_escape(_ILLEGAL_XML.sub('', str(v)), {'"': '&quot;'})


def _cell_kind(v):
    if isinstance(v, datetime):
        return 'd'
    if isinstance(v, bool):
        return 's'
    if isinstance(v, (int, float)):
        return 'n'
    return 's'


def _normalize_value(v):
    if v is None or (not isinstance(v, (list, tuple)) and pd.isna(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, 'item') and not isinstance(v, (str, bytes)):
        return v.item()
    return v


def _pivot_parts(data_df, data_sheet, p, order, idx):
    df = _prepare_for_output(data_df)
    cols = list(df.columns)
    n_rows = len(df)
    team_i, coll_i, note_i = cols.index(PIVOT_TEAM), cols.index(PIVOT_COLL), cols.index(PIVOT_NOTE)
    acc_i, amt_i = cols.index(PIVOT_ACC), cols.index(PIVOT_AMT)

    values = [[_normalize_value(v) for v in df[c].tolist()] for c in cols]

    kinds = []
    for col_vals in values:
        ks = {_cell_kind(v) for v in col_vals if v is not None}
        if not ks:
            kinds.append('blank')
        elif ks == {'n'}:
            kinds.append('n')
        elif ks == {'d'}:
            kinds.append('d')
        else:
            kinds.append('s')

    shared_cols = {team_i, coll_i, note_i}
    shared_maps = {}
    for ci in shared_cols:
        vals = ['' if v is None else str(v).strip() if ci == note_i else str(v) for v in values[ci]]
        values[ci] = vals
        uniq = list(dict.fromkeys(vals))
        if ci == note_i:
            uniq += [x for x in p['notes'] if x not in uniq]
        shared_maps[ci] = {u: k for k, u in enumerate(uniq)}

    fields_xml = []
    for ci, name in enumerate(cols):
        if ci in shared_cols:
            items = ''.join(f'<s v="{_xa(u)}"/>' for u in shared_maps[ci])
            fields_xml.append(f'<cacheField name="{_xa(name)}" numFmtId="0"><sharedItems count="{len(shared_maps[ci])}">{items}</sharedItems></cacheField>')
            continue
        col_vals = values[ci]
        has_blank = any(v is None for v in col_vals)
        k = kinds[ci]
        if k == 'blank':
            si = '<sharedItems containsNonDate="0" containsString="0" containsBlank="1"/>'
        elif k == 'n':
            nums = [float(v) for v in col_vals if v is not None]
            is_int = all(float(v).is_integer() for v in nums)
            attrs = 'containsSemiMixedTypes="0" containsString="0" containsNumber="1"' if not has_blank else 'containsString="0" containsBlank="1" containsNumber="1"'
            if is_int:
                attrs += ' containsInteger="1"'
            si = f'<sharedItems {attrs} minValue="{repr(min(nums))}" maxValue="{repr(max(nums))}"/>'
        elif k == 'd':
            ds = [v for v in col_vals if v is not None]
            attrs = 'containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0"' if not has_blank else 'containsNonDate="0" containsDate="1" containsString="0" containsBlank="1"'
            si = f'<sharedItems {attrs} minDate="{min(ds).strftime("%Y-%m-%dT%H:%M:%S")}" maxDate="{max(ds).strftime("%Y-%m-%dT%H:%M:%S")}"/>'
        else:
            si = '<sharedItems containsBlank="1"/>' if has_blank else '<sharedItems/>'
        fields_xml.append(f'<cacheField name="{_xa(name)}" numFmtId="0">{si}</cacheField>')

    last_col_letter = get_column_letter(len(cols))
    ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    cache_def = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition {ns} r:id="rId1" refreshedBy="Ejada" '
        f'refreshedDate="46330.5" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="{n_rows}">'
        f'<cacheSource type="worksheet"><worksheetSource ref="A1:{last_col_letter}{n_rows + 1}" sheet="{_xa(data_sheet)}"/></cacheSource>'
        f'<cacheFields count="{len(cols)}">{"".join(fields_xml)}</cacheFields></pivotCacheDefinition>'
    )

    recs = []
    for ri in range(n_rows):
        cells = []
        for ci in range(len(cols)):
            v = values[ci][ri]
            if ci in shared_cols:
                cells.append(f'<x v="{shared_maps[ci][v]}"/>')
            elif v is None:
                cells.append('<m/>')
            elif kinds[ci] == 'n':
                cells.append(f'<n v="{repr(float(v)) if not float(v).is_integer() else int(float(v))}"/>')
            elif kinds[ci] == 'd':
                cells.append(f'<d v="{v.strftime("%Y-%m-%dT%H:%M:%S")}"/>')
            else:
                cells.append(f'<s v="{_xa(v if not isinstance(v, datetime) else v.strftime("%Y-%m-%d %H:%M:%S"))}"/>')
        recs.append('<r>' + ''.join(cells) + '</r>')
    cache_records = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords {ns} count="{n_rows}">{"".join(recs)}</pivotCacheRecords>'

    notes = p['notes']
    n_notes = len(notes)
    team_order = [t['team'] for t in p['teams']]
    coll_order = []
    for t in p['teams']:
        for r in t['rows']:
            if r['coll'] not in coll_order:
                coll_order.append(r['coll'])
    team_pos = {t: k for k, t in enumerate(team_order)}
    coll_pos = {c: k for k, c in enumerate(coll_order)}

    def items_xml(order_list, ci):
        return f'<items count="{len(order_list)}">' + ''.join(f'<item x="{shared_maps[ci][o]}"/>' for o in order_list) + '</items>'

    pf = []
    for ci in range(len(cols)):
        base = 'compact="0" outline="0" showAll="0" defaultSubtotal="0"'
        if ci == note_i:
            pf.append(f'<pivotField axis="axisCol" compact="0" outline="0" showAll="1" defaultSubtotal="0" sortType="descending">{items_xml(notes, ci)}</pivotField>')
        elif ci == team_i:
            pf.append(f'<pivotField name="اسم المشرف " axis="axisRow" {base}>{items_xml(team_order, ci)}</pivotField>')
        elif ci == coll_i:
            st_attr = ' sortType="descending"' if order == 'desc' else ''
            pf.append(f'<pivotField name="اسم المحصل " axis="axisRow" {base}{st_attr}>{items_xml(coll_order, ci)}</pivotField>')
        elif ci == acc_i:
            pf.append(f'<pivotField dataField="1" compact="0" outline="0" showAll="0" defaultSubtotal="0"/>')
        elif ci == amt_i:
            pf.append(f'<pivotField dataField="1" compact="0" outline="0" showAll="0" defaultSubtotal="0"/>')
        else:
            pf.append(f'<pivotField compact="0" outline="0" showAll="0" defaultSubtotal="0"/>')

    row_items = []
    for t in p['teams']:
        for k, r in enumerate(t['rows']):
            if k == 0:
                tp = team_pos[t['team']]
                tx = '<x/>' if tp == 0 else f'<x v="{tp}"/>'
                cp = coll_pos[r['coll']]
                cx = '<x/>' if cp == 0 else f'<x v="{cp}"/>'
                row_items.append(f'<i>{tx}{cx}</i>')
            else:
                row_items.append(f'<i r="1"><x v="{coll_pos[r["coll"]]}"/></i>')
    row_items.append('<i t="grand"><x/></i>')

    col_items = []
    for k in range(n_notes):
        first = '<x/>' if k == 0 else f'<x v="{k}"/>'
        col_items.append(f'<i>{first}<x/></i>')
        col_items.append('<i r="1" i="1"><x v="1"/></i>')
    col_items.append('<i t="grand"><x/></i>')
    col_items.append('<i t="grand" i="1"><x/></i>')

    n_body = len(row_items) - 1
    last_row = 5 + n_body + 1
    last_col = 2 + 2 * n_notes + 2
    ref = f'A3:{get_column_letter(last_col)}{last_row}'

    pivot_def = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'name="PivotTable{idx}" cacheId="{idx}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" '
        f'applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" grandTotalCaption="الاجمالي" missingCaption="0" '
        f'updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" mergeItem="1" createdVersion="8" indent="0" '
        f'compact="0" compactData="0" multipleFieldFilters="0" rowHeaderCaption="المحصل" colHeaderCaption="الفرع">'
        f'<location ref="{ref}" firstHeaderRow="1" firstDataRow="3" firstDataCol="2"/>'
        f'<pivotFields count="{len(cols)}">{"".join(pf)}</pivotFields>'
        f'<rowFields count="2"><field x="{team_i}"/><field x="{coll_i}"/></rowFields>'
        f'<rowItems count="{len(row_items)}">{"".join(row_items)}</rowItems>'
        f'<colFields count="2"><field x="{note_i}"/><field x="-2"/></colFields>'
        f'<colItems count="{len(col_items)}">{"".join(col_items)}</colItems>'
        f'<dataFields count="2"><dataField name="عدد الحسابات" fld="{acc_i}" subtotal="count" baseField="{coll_i}" baseItem="0"/>'
        f'<dataField name="متبقي المديونية" fld="{amt_i}" baseField="0" baseItem="0" numFmtId="3"/></dataFields>'
        f'<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
        f'</pivotTableDefinition>'
    )
    return cache_def, cache_records, pivot_def


def inject_pivot_tables(xlsx_bytes, specs):
    zin = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    files = {n: zin.read(n) for n in zin.namelist()}
    zin.close()

    wb_xml = files['xl/workbook.xml'].decode('utf-8')
    wb_rels = files['xl/_rels/workbook.xml.rels'].decode('utf-8')
    ct = files['[Content_Types].xml'].decode('utf-8')

    sheet_rid = {}
    for m in re.finditer(r'<sheet\b[^>]*?>', wb_xml):
        tag = m.group(0)
        nm = re.search(r'name="([^"]*)"', tag).group(1)
        rid = re.search(r'r:id="([^"]*)"', tag).group(1)
        sheet_rid[nm.replace('&amp;', '&')] = rid
    rid_target = {}
    for m in re.finditer(r'<Relationship\b[^>]*?>', wb_rels):
        tag = m.group(0)
        rid_target[re.search(r'Id="([^"]*)"', tag).group(1)] = re.search(r'Target="([^"]*)"', tag).group(1)

    pivot_caches_xml = ''
    new_rels = ''
    new_ct = ''
    for idx, sp in enumerate(specs, 1):
        cache_def, cache_records, pivot_def = _pivot_parts(sp['data_df'], sp['data_sheet'], sp['pivot'], sp['order'], idx)
        files[f'xl/pivotCache/pivotCacheDefinition{idx}.xml'] = cache_def.encode('utf-8')
        files[f'xl/pivotCache/pivotCacheRecords{idx}.xml'] = cache_records.encode('utf-8')
        files[f'xl/pivotCache/_rels/pivotCacheDefinition{idx}.xml.rels'] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords{idx}.xml"/></Relationships>'
        ).encode('utf-8')
        files[f'xl/pivotTables/pivotTable{idx}.xml'] = pivot_def.encode('utf-8')
        files[f'xl/pivotTables/_rels/pivotTable{idx}.xml.rels'] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition{idx}.xml"/></Relationships>'
        ).encode('utf-8')

        target = rid_target[sheet_rid[sp['summary_sheet']]]
        sheet_file = target.lstrip('/') if target.startswith('/') else 'xl/' + target
        rels_path = sheet_file.replace('worksheets/', 'worksheets/_rels/') + '.rels'
        rel = (f'<Relationship Id="rIdPv{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" '
               f'Target="../pivotTables/pivotTable{idx}.xml"/>')
        if rels_path in files:
            files[rels_path] = files[rels_path].decode('utf-8').replace('</Relationships>', rel + '</Relationships>').encode('utf-8')
        else:
            files[rels_path] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                + rel + '</Relationships>').encode('utf-8')

        new_rels += (f'<Relationship Id="rIdPc{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" '
                     f'Target="pivotCache/pivotCacheDefinition{idx}.xml"/>')
        pivot_caches_xml += f'<pivotCache cacheId="{idx}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdPc{idx}"/>'
        new_ct += (f'<Override PartName="/xl/pivotCache/pivotCacheDefinition{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>'
                   f'<Override PartName="/xl/pivotCache/pivotCacheRecords{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>'
                   f'<Override PartName="/xl/pivotTables/pivotTable{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>')

    if specs:
        wb_rels = wb_rels.replace('</Relationships>', new_rels + '</Relationships>')
        ct = ct.replace('</Types>', new_ct + '</Types>')
        if re.search(r'<calcPr\b[^>]*/>', wb_xml):
            wb_xml = re.sub(r'(<calcPr\b[^>]*/>)', r'\1<pivotCaches>' + pivot_caches_xml.replace('\\', '\\\\') + '</pivotCaches>', wb_xml, count=1)
        else:
            wb_xml = wb_xml.replace('</workbook>', '<pivotCaches>' + pivot_caches_xml + '</pivotCaches></workbook>')
        files['xl/workbook.xml'] = wb_xml.encode('utf-8')
        files['xl/_rels/workbook.xml.rels'] = wb_rels.encode('utf-8')
        files['[Content_Types].xml'] = ct.encode('utf-8')

    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        zout.writestr('[Content_Types].xml', files.pop('[Content_Types].xml'))
        for n, b in files.items():
            zout.writestr(n, b)
    out.seek(0)
    return out


def _can_pivot(df):
    need = [PIVOT_TEAM, PIVOT_COLL, PIVOT_NOTE, PIVOT_ACC, PIVOT_AMT]
    return (df is not None and len(df) > 0 and all(c in df.columns for c in need)
            and len(set(df.columns)) == len(df.columns))


def build_report_xlsx(active_df, broken_df, neglect_df=None):
    wb = Workbook()
    ws = wb.active
    ws.title = 'تقرير الوعود القائمة'
    _write_df_sheet(ws, active_df)

    ws = wb.create_sheet('ملخص تقرير الوعود القائمة')
    _write_pivot_sheet(ws, build_pivot(active_df, order='first'), 'تقرير الوعود القائمة البنك الاهلي السعودي SNB (NPL&60Dpd)')

    ws = wb.create_sheet('تقرير الوعود المكسورة')
    _write_df_sheet(ws, broken_df)

    ws = wb.create_sheet('ملخص تقرير الوعود المكسورة')
    _write_pivot_sheet(ws, build_pivot(broken_df, order='desc'), 'تقرير الوعود المكسورة البنك الاهلي SNB (NPL&60Dpd)')

    if neglect_df is not None and len(neglect_df.columns) > 0:
        ws = wb.create_sheet('تقرير الإهمال')
        _write_df_sheet(ws, neglect_df)

    buf = io.BytesIO()
    wb.save(buf)

    specs = []
    if _can_pivot(active_df):
        specs.append({'summary_sheet': 'ملخص تقرير الوعود القائمة', 'data_sheet': 'تقرير الوعود القائمة',
                      'data_df': active_df, 'pivot': build_pivot(active_df, order='first'), 'order': 'first'})
    if _can_pivot(broken_df):
        specs.append({'summary_sheet': 'ملخص تقرير الوعود المكسورة', 'data_sheet': 'تقرير الوعود المكسورة',
                      'data_df': broken_df, 'pivot': build_pivot(broken_df, order='desc'), 'order': 'desc'})
    if specs:
        try:
            return inject_pivot_tables(buf.getvalue(), specs)
        except Exception:
            pass
    buf.seek(0)
    return buf

# ==========================================
# إدارة الملفات لكل صفحة تقرير بشكل منفصل تماماً
# ==========================================
if 'page_files' not in st.session_state:
    st.session_state.page_files = {
        'Overview': None,
        'Promises': None,
        'Neglect': None,
        'Activity': None
    }

curr_page = st.session_state.current_page

# دالة معالجة بيانات أي صفحة
def process_page_data(uploaded_files):
    if not uploaded_files:
        return None, None, None, None, None, 0, 0, 0, 0, 0, pd.DataFrame()
    
    today = pd.Timestamp(date.today()).normalize()
    promises_dfs = []
    activity_dfs = []
    all_dfs = []

    for f in uploaded_files:
        try:
            try:
                f.seek(0)
            except Exception:
                pass
            temp_df = pd.read_csv(f, dtype=str) if f.name.endswith('.csv') else pd.read_excel(f, dtype=str)
            temp_df.columns = temp_df.columns.astype(str).str.strip()
            all_dfs.append(temp_df)
            if 'Follow up Due Date' in temp_df.columns and 'Follow up Last Date' in temp_df.columns:
                promises_dfs.append(temp_df)
            has_created = any(c.lower().strip() in ['created on', 'created_on', 'createdon', 'تاريخ الإنشاء', 'وقت المعالجة', 'date', 'التاريخ'] for c in temp_df.columns)
            if has_created or not ('Follow up Due Date' in temp_df.columns):
                activity_dfs.append(temp_df)
        except Exception as e:
            st.error(f"خطأ في قراءة الملف {f.name}: {e}")

    df_all = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
    df_promises = pd.concat(promises_dfs, ignore_index=True) if promises_dfs else pd.DataFrame()
    df_act_raw = pd.concat(activity_dfs, ignore_index=True) if activity_dfs else df_all.copy()

    created_col = find_col_in(df_act_raw, ['created on', 'created_on', 'createdon', 'تاريخ الإنشاء', 'وقت المعالجة', 'date', 'التاريخ'])
    if created_col:
        df_act_raw['Created_DateTime'] = pd.to_datetime(df_act_raw[created_col], errors='coerce')
        df_act_raw['Date'] = df_act_raw['Created_DateTime'].dt.date
        df_act_raw['Time'] = df_act_raw['Created_DateTime'].dt.time

    collector_col = find_col_in(df_act_raw, ['collecter', 'collector', 'المحصل', 'اسم المحصل', 'الموظف', 'user', 'agent'])
    if not collector_col:
        collector_col = find_col_in(df_all, ['collecter', 'collector', 'المحصل', 'اسم المحصل', 'الموظف', 'user', 'agent'])

    id_col = find_col_in(df_act_raw, ['id numer', 'id number', 'id_number', 'id', 'رقم الهوية', 'رقم العميل', 'account id'])
    notes_col = find_col_in(df_act_raw, ['notes', 'note', 'الملاحظات', 'ملاحظات', 'feedback'])

    dedup_cols = [c for c in [collector_col, id_col, notes_col] if c is not None and c in df_act_raw.columns]
    activity_base_df = df_act_raw.drop_duplicates(subset=dedup_cols).copy() if len(dedup_cols) == 3 else df_act_raw.copy()

    active_promises_df = pd.DataFrame()
    broken_promises_df = pd.DataFrame()
    neglect_df = pd.DataFrame()
    total_unhandled = 0

    if not df_promises.empty and 'Follow up Due Date' in df_promises.columns and 'Follow up Last Date' in df_promises.columns:
        due_date = pd.to_datetime(df_promises['Follow up Due Date'], errors='coerce').dt.normalize()
        last_date = pd.to_datetime(df_promises['Follow up Last Date'], errors='coerce').dt.normalize()
        base_condition = (
            (df_promises['حالة المعالجة - التمويل'].astype(str).str.strip() == 'غير معالج') & 
            (df_promises['تحديث فئة العميل - التمويل'].astype(str).str.strip().str.lower() != 'writeoff') & 
            (df_promises['Classification'].astype(str).str.strip() == 'وعود السداد قيد التنفيذ')
        )
        active_promises_df = df_promises[base_condition & (due_date == today) & (last_date != today)].copy()
        broken_promises_df = df_promises[base_condition & (due_date < today)].copy()
        _diff = (last_date[broken_promises_df.index] - due_date[broken_promises_df.index]).dt.days
        broken_promises_df = broken_promises_df[_diff < 0]
        if not broken_promises_df.empty:
            broken_promises_df = broken_promises_df.drop(columns=['الفرق '], errors='ignore')
            broken_promises_df.insert(broken_promises_df.columns.get_loc('Follow up Last Date') + 1, 'الفرق ', _diff[_diff < 0].astype(int))

        neglect_base = df_promises[
            (df_promises['حالة المعالجة - التمويل'].astype(str).str.strip() == 'غير معالج') & 
            (df_promises['تحديث فئة العميل - التمويل'].astype(str).str.strip().str.lower() != 'writeoff')
        ].copy()
        n_due_date = pd.to_datetime(neglect_base['Follow up Due Date'], errors='coerce').dt.normalize()
        neglect_base = neglect_base[n_due_date != today]
        n_last_date = pd.to_datetime(neglect_base['Follow up Last Date'], errors='coerce').dt.normalize()
        neglect_base['Neglect_Days'] = (today - n_last_date).dt.days
        neglect_df = neglect_base[neglect_base['Neglect_Days'] > 3].copy()

        total_unhandled = len(df_promises[base_condition])

    output_buffer = build_report_xlsx(active_promises_df, broken_promises_df, neglect_df)
    report_bytes = output_buffer.getvalue()

    total_rows = len(df_promises) if not df_promises.empty else len(df_all)
    active_count = len(active_promises_df)
    broken_count = len(broken_promises_df)
    neglect_count = len(neglect_df)
    total_activity_count = len(activity_base_df)

    return active_promises_df, broken_promises_df, neglect_df, activity_base_df, report_bytes, total_rows, total_unhandled, active_count, broken_count, neglect_count, total_activity_count


# ==========================================
# صفحة النشاط: العنوان + الخانات التلاتة بحجم محكوم على قد خانة الرفع وزر X للمسح
# ==========================================
def render_activity_header():
    st.markdown(f"""
    <div style="background-color: {card_color}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; margin-bottom: 20px; display:flex; align-items:center; justify-content:center; gap:14px;">
        <div style="width:42px; height:42px; flex-shrink:0; background-color:{PHONE_ICON_COLOR}; -webkit-mask: url(data:image/png;base64,{PHONE_ICON_B64}) no-repeat center / contain; mask: url(data:image/png;base64,{PHONE_ICON_B64}) no-repeat center / contain;"></div>
        <div style="margin:0; color:{text_color}; font-weight:900; font-size:28px; line-height:1.3;">تقرير نشاط المحصلين</div>
    </div>
    """, unsafe_allow_html=True)

def render_activity_boxes(activity_base_df=None):
    """الترتيب (من اليمين): ملف الموديل - ارفعي الملف هنا يا رنا (مع زر X للمسح) - اسم المحصل"""
    col_model, col_file, col_coll = st.columns([1, 1, 1], gap="medium")

    with col_model:
        with st.container(border=True):
            st.markdown(f"<p style='font-size: 13px; color: {sub_text_color}; direction: rtl; text-align: right; margin: 0 0 2px 0; font-weight: bold;'>التقييم (1) يعني 'ناجحة'، والتقييم (0) يعني 'مغطاة'.</p>", unsafe_allow_html=True)
            mapping_file = st.file_uploader("شيت موديل النشاط", type=["xlsx"], key="mapping_uploader")

    with col_file:
        with st.container(border=True):
            cur_files = st.session_state.page_files.get('Activity') or []
            if cur_files:
                cur_names = "، ".join(f.name for f in cur_files)
                hdr_col1, hdr_col2 = st.columns([8.5, 1.5])
                with hdr_col1:
                    st.markdown(f"<p style='font-size: 13px; color: {sub_text_color}; direction: rtl; text-align: right; margin: 0; font-weight: bold;'>الملف الحالي: {cur_names}</p>", unsafe_allow_html=True)
                with hdr_col2:
                    st.button("✖", key="clear_act_files_btn", on_click=clear_activity_files, help="إزالة الملف المرفوع")
            else:
                st.markdown(f"<p style='font-size: 13px; color: {sub_text_color}; direction: rtl; text-align: right; margin: 0 0 2px 0; font-weight: bold;'>الملف الحالي: لسه ما اترفعش ملف</p>", unsafe_allow_html=True)

            uploader_key = f"uploader_inpage_Activity_{st.session_state.act_uploader_ver}"
            new_files = st.file_uploader("ارفعي الملف هنا يا رنا", type=["xlsx", "csv"], accept_multiple_files=True, key=uploader_key)
            if new_files:
                new_sig = [(f.name, f.size) for f in new_files]
                cur_sig = [(f.name, f.size) for f in cur_files]
                if new_sig != cur_sig:
                    st.session_state.page_files['Activity'] = new_files
                    st.rerun()

    collector_col_name = None
    selected_collector = "الكل (جميع المحصلين)"
    with col_coll:
        with st.container(border=True):
            st.markdown(f"<p style='font-size: 13px; color: {sub_text_color}; direction: rtl; text-align: right; margin: 0 0 2px 0; font-weight: bold;'>فلترة حسب المحصل</p>", unsafe_allow_html=True)
            collector_options = ["الكل (جميع المحصلين)"]
            if activity_base_df is not None:
                collector_col_name = find_col_in(activity_base_df, ['collecter', 'collector', 'المحصل', 'اسم المحصل'])
                if collector_col_name and collector_col_name in activity_base_df.columns:
                    unique_colls = activity_base_df[collector_col_name].dropna().astype(str).unique().tolist()
                    collector_options.extend(sorted(unique_colls))
            selected_collector = st.selectbox("👤 اسم المحصل (Collector):", options=collector_options, disabled=(activity_base_df is None))

    return selected_collector, mapping_file, collector_col_name

# التحقق من الملفات الخاصة بالصفحة الحالية فقط
page_uploaded = st.session_state.page_files[curr_page]

if page_uploaded is None:
    if curr_page == 'Overview':
        st.markdown(f"""
        <div style="background-color: {card_color}; border-radius: 15px; text-align: center; padding: 60px 20px; direction: rtl; border: 1px solid {border_color}; margin-top: 20px; margin-bottom: 30px; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
            <h2 style="color: {text_color}; margin-bottom: 15px; font-weight:900;">أهلاً بيكي يا رنا في منصة تقارير شركة إجادة لتحصيل الديون</h2>
            <p style="color: {kpi_title_color}; font-size: 18px; font-weight: bold;">من شريط التنقل بالأعلى، اختاري أي تقرير لعرضه ورفع ملفاته الخاصة.</p>
        </div>
        """, unsafe_allow_html=True)
    elif curr_page == 'Activity':
        render_activity_header()
        render_activity_boxes(None)
        st.markdown(f"<p style='color: {label_color}; font-weight: 800; font-size: 16px; text-align: right; direction: rtl; margin-top: 10px;'>ارفعي ملف النشاط من الخانة اللي فوق عشان يظهر التقرير.</p>", unsafe_allow_html=True)
    else:
        st.markdown(f"<p style='color: {label_color}; font-weight: 800; font-size: 18px; text-align: right; direction: rtl; margin-top: 20px;'>ارفعي الملفات هنا ي رنا</p>", unsafe_allow_html=True)
        col_up1, col_up2 = st.columns([7, 3])
        with col_up1:
            files_in = st.file_uploader("ارفعي الملفات هنا ي رنا", type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed", key=f"uploader_{curr_page}")
            if files_in:
                st.session_state.page_files[curr_page] = files_in
                st.rerun()
else:
    active_promises_df, broken_promises_df, neglect_df, activity_base_df, report_bytes, total_rows, total_unhandled, active_count, broken_count, neglect_count, total_activity_count = process_page_data(st.session_state.page_files[curr_page])

    # ==========================================
    # صفحة الـ Overview الشاملة
    # ==========================================
    if curr_page == 'Overview':
        st.markdown(f"<h2 style='text-align:center; margin:0; color:{label_color}; font-weight:900;'>📊 لوحة القيادة الشاملة (Executive Overview)</h2>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align:center; color:{sub_text_color}; font-weight: bold; margin-bottom: 30px;'>ملخص أداء محفظة الديون ونشاط المحصلين</p>", unsafe_allow_html=True)
        
        st.markdown(f"""
        <div style="display: grid; grid-template-columns: repeat(6, 1fr); gap: 12px; margin-bottom: 25px;">
            <div class="kpi-card" style="border-right-color: {border_color};">
                <div class="kpi-icon">📋</div>
                <div class="kpi-title">إجمالي الحسابات</div>
                <div class="kpi-value">{total_rows:,}</div>
            </div>
            <div class="kpi-card" style="border-right-color: #586F8C;">
                <div class="kpi-icon">⏳</div>
                <div class="kpi-title">وعود قيد التنفيذ</div>
                <div class="kpi-value">{total_unhandled:,}</div>
            </div>
            <div class="kpi-card" style="border-right-color: #2CA549;">
                <div class="kpi-icon">{GREEN_CHECK_SVG}</div>
                <div class="kpi-title">وعود قائمة</div>
                <div class="kpi-value">{active_count:,}</div>
            </div>
            <div class="kpi-card" style="border-right-color: #586F8C;">
                <div class="kpi-icon">⚠️</div>
                <div class="kpi-title">وعود مكسورة</div>
                <div class="kpi-value">{broken_count:,}</div>
            </div>
            <div class="kpi-card" style="border-right-color: #99AABF;">
                <div class="kpi-icon">🚨</div>
                <div class="kpi-title">حالات إهمال</div>
                <div class="kpi-value">{neglect_count:,}</div>
            </div>
            <div class="kpi-card" style="border-right-color: {border_color};">
                <div class="kpi-icon">📞</div>
                <div class="kpi-title">إجمالي النشاط</div>
                <div class="kpi-value">{total_activity_count:,}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        chart_col1, chart_col2, chart_col3 = st.columns(3)

        with chart_col1:
            with st.container(border=True):
                fig_donut = go.Figure(go.Pie(
                    labels=['الوعود القائمة', 'الوعود المكسورة', 'حالات الإهمال'],
                    values=[active_count, broken_count, neglect_count],
                    hole=0.6,
                    marker_colors=["#1F4373", "#304D73", "#586F8C"],
                    textinfo='percent+label',
                    textfont_color="#FFFFFF"
                ))
                fig_donut.update_layout(
                    title=dict(text="توزيع الوعود وحالات الإهمال", x=0.5, xanchor='center', font=dict(color=ejada_navy_dark, size=15, weight='bold')),
                    paper_bgcolor='white', plot_bgcolor='white', font_color=ejada_navy_dark, 
                    margin=dict(t=50, b=10, l=10, r=10), showlegend=False
                )
                st.plotly_chart(fig_donut, use_container_width=True, config={'displayModeBar': False})

        with chart_col2:
            with st.container(border=True):
                total_extracted = active_count + broken_count
                gauge_val = int((active_count / total_extracted * 100)) if total_extracted > 0 else 0
                fig_gauge = go.Figure(go.Indicator(
                    mode="gauge+number",
                    value=gauge_val,
                    number={'suffix': "%", 'font': {'color': ejada_navy_dark, 'size': 34, 'weight': 'bold'}},
                    title_text="نسبة الوعود القائمة من الإجمالي",
                    gauge={
                        'axis': {'range': [0, 100], 'visible': False},
                        'bar': {'color': "#304D73"}, 'bgcolor': "#E2E8F0", 'borderwidth': 0,
                        'steps': [{'range': [0, 50], 'color': '#99AABF'}, {'range': [50, 100], 'color': '#586F8C'}]
                    }
                ))
                fig_gauge.update_layout(
                    title_font=dict(color=ejada_navy_dark, size=15, weight='bold'),
                    title_x=0.5,
                    paper_bgcolor='white', 
                    font_color=ejada_navy_dark, 
                    margin=dict(l=20, r=20, t=50, b=20)
                )
                st.plotly_chart(fig_gauge, use_container_width=True, config={'displayModeBar': False})

        with chart_col3:
            with st.container(border=True):
                collector_col = find_col_in(activity_base_df, ['collecter', 'collector', 'المحصل', 'اسم المحصل'])
                if collector_col and collector_col in activity_base_df.columns and total_activity_count > 0:
                    top_colls = activity_base_df[collector_col].value_counts().head(5)
                    fig_bar = go.Figure(go.Bar(
                        x=top_colls.index, y=top_colls.values,
                        marker_color=["#1F4373", "#304D73", "#586F8C", "#99AABF", "#1F4373"],
                        text=top_colls.values,
                        textposition='auto',
                        textfont=dict(color="#FFFFFF", weight='bold')
                    ))
                    fig_bar.update_layout(
                        title=dict(text="أكثر 5 محصلين نشاطاً", x=0.5, xanchor='center', font=dict(color=ejada_navy_dark, size=15, weight='bold')),
                        paper_bgcolor='white', plot_bgcolor='white', font_color=ejada_navy_dark, margin=dict(t=50, b=10, l=10, r=10)
                    )
                    fig_bar.update_yaxes(showgrid=False, visible=False)
                    st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar': False})
                else:
                    st.info("لا توجد بيانات نشاط كافية لعرض رسم بياني.")
                
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(label=f"📥 تحميل التقرير الكامل.xlsx", data=report_bytes, file_name=f"Daily_Promises_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    # ==========================================
    # صفحة الوعود القائمة والمكسورة (مدمجة)
    # ==========================================
    elif curr_page == 'Promises':
        st.markdown(f"""
        <div style="background-color: {card_color}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; text-align: right; margin-bottom: 20px;">
            <h3 style="margin:0; color:#1F4373; font-weight:900;">🤝 إدارة ومتابعة الوعود (القائمة والمكسورة)</h3>
        </div>
        """, unsafe_allow_html=True)
        
        tab_active, tab_broken = st.tabs([f"✅ الوعود القائمة ({active_count:,})", f"⚠️ الوعود المكسورة ({broken_count:,})"])
        
        with tab_active:
            st.markdown(f"<h4 style='text-align:right; direction:rtl; color:{label_color}; font-weight:900;'>📊 ملخص تقرير الوعود القائمة (Pivot)</h4>", unsafe_allow_html=True)
            st.dataframe(pivot_to_frame(build_pivot(active_promises_df, order='first')), use_container_width=True, hide_index=True)
            st.markdown(f"<h4 style='text-align:right; direction:rtl; color:{label_color}; font-weight:900; margin-top:20px;'>📋 تفاصيل شيت الوعود القائمة</h4>", unsafe_allow_html=True)
            st.dataframe(active_promises_df, use_container_width=True, height=400)
            
        with tab_broken:
            st.markdown(f"<h4 style='text-align:right; direction:rtl; color:{label_color}; font-weight:900;'>📊 ملخص تقرير الوعود المكسورة (Pivot)</h4>", unsafe_allow_html=True)
            st.dataframe(pivot_to_frame(build_pivot(broken_promises_df, order='desc')), use_container_width=True, hide_index=True)
            st.markdown(f"<h4 style='text-align:right; direction:rtl; color:{label_color}; font-weight:900; margin-top:20px;'>📋 تفاصيل شيت الوعود المكسورة</h4>", unsafe_allow_html=True)
            st.dataframe(broken_promises_df, use_container_width=True, height=400)
            
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button("📥 تحميل التقرير الشامل (الإكسيل بالملخصات).xlsx", data=report_bytes, file_name=f"Daily_Promises_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_promises_all")
        
    elif curr_page == 'Neglect':
        st.markdown(f"""
        <div style="background-color: {card_color}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; text-align: right; margin-bottom: 20px;">
            <h3 style="margin:0; color:#586F8C; font-weight:900;">🚨 تقرير الإهمال (تأخير أكثر من 3 أيام)</h3>
        </div>
        """, unsafe_allow_html=True)
        default_sub_states = ['وعد بسداد مبلغ المعالجة', 'واعد بسداد مبلغ المعالجه', 'رافض السداد', 'الجوال مغلق', 'مماطل', 'تم سداد جزء وليس مبلغ المعالجه', 'لا يجيب', 'متقاعد الراتب علي بنك آخر', 'يرغب بجدوله المديونيه']
        selected_sub_states = st.multiselect("فلاتر الـ Sub State المشمولة:", options=default_sub_states, default=default_sub_states)
        sub_col = 'Sub State' if 'Sub State' in neglect_df.columns else 'sub state' if 'sub state' in neglect_df.columns else None
        final_neglect_df = neglect_df[neglect_df[sub_col].isin(selected_sub_states)] if sub_col else neglect_df
        st.markdown(f"<p style='text-align:right; direction:rtl; color:{label_color}; font-weight: bold;'>العدد الإجمالي: <b style='color:#586F8C; font-size:18px;'>{len(final_neglect_df):,}</b> حساب مهمل</p>", unsafe_allow_html=True)
        st.dataframe(final_neglect_df, use_container_width=True, height=450)

    # ==========================================
    # صفحة تقرير النشاط
    # ==========================================
    elif curr_page == 'Activity':
        render_activity_header()
        selected_collector, mapping_file, collector_col_name = render_activity_boxes(activity_base_df)

        # فلاتر المدى الزمني والشفت والبريك
        filter_act_col1, filter_act_col2, filter_act_col3, filter_act_col4 = st.columns([1.5, 2, 1.5, 1.5])
        with filter_act_col1:
            time_range_mode = st.selectbox("📅 مدى التقرير:", ["يومي", "أسبوعي", "شهري"])
        with filter_act_col2:
            shift_option = st.selectbox("⏰ فترة النشاط (الشفت):", ["كل اليوم بدون تحديد", "من 2:00 م إلى 12:00 منتصف الليل (14:00 - 00:00)", "من 12:00 ظهراً إلى 4:00 عصراً (12:00 - 16:00)", "من 4:00 عصراً إلى آخر اليوم (16:00 - 23:59)"])
        with filter_act_col3:
            exclude_last_hour = st.checkbox("⏳ استبعاد آخر ساعة", value=False)
        with filter_act_col4:
            exclude_break = st.checkbox("☕ استبعاد البريك (12:00 - 12:30)", value=False)

        act_df = activity_base_df.copy()

        if collector_col_name and collector_col_name in act_df.columns and selected_collector != "الكل (جميع المحصلين)":
            act_df = act_df[act_df[collector_col_name].astype(str).str.strip() == selected_collector]

        created_col = find_col_in(act_df, ['created on', 'created_on', 'createdon', 'تاريخ الإنشاء', 'وقت المعالجة', 'date', 'التاريخ'])
        if created_col and 'Created_DateTime' in act_df.columns and act_df['Created_DateTime'].notna().any():
            act_df = act_df.dropna(subset=['Created_DateTime']).copy()
            max_dt = act_df['Created_DateTime'].max()
            if time_range_mode == "يومي":
                act_df = act_df[act_df['Created_DateTime'].dt.date == max_dt.date()]
            elif time_range_mode == "أسبوعي":
                start_week = max_dt - timedelta(days=7)
                act_df = act_df[act_df['Created_DateTime'] >= start_week]
            elif time_range_mode == "شهري":
                start_month = max_dt - timedelta(days=30)
                act_df = act_df[act_df['Created_DateTime'] >= start_month]

            act_times = act_df['Created_DateTime'].dt.time
            if shift_option == "من 2:00 م إلى 12:00 منتصف الليل (14:00 - 00:00)":
                start_t, end_t = time(14, 0), time(23, 0) if exclude_last_hour else time(23, 59, 59)
                act_df = act_df[(act_times >= start_t) & (act_times <= end_t)]
            elif shift_option == "من 12:00 ظهراً إلى 4:00 عصراً (12:00 - 16:00)":
                start_t, end_t = time(12, 0), time(15, 0) if exclude_last_hour else time(16, 0)
                act_df = act_df[(act_times >= start_t) & (act_times <= end_t)]
            elif shift_option == "من 4:00 عصراً إلى آخر اليوم (16:00 - 23:59)":
                start_t, end_t = time(16, 0), time(23, 0) if exclude_last_hour else time(23, 59, 59)
                act_df = act_df[(act_times >= start_t) & (act_times <= end_t)]
            if exclude_break:
                curr_times = act_df['Created_DateTime'].dt.time
                act_df = act_df[~((curr_times >= time(12, 0)) & (curr_times <= time(12, 30)))]

        # ربط التقييم بعمود Notes والتعامل مع الملاحظات الجديدة
        notes_col_name = find_col_in(act_df, ['notes', 'note', 'الملاحظات', 'ملاحظات', 'feedback'])
        
        if mapping_file is not None and notes_col_name:
            map_df = pd.read_excel(mapping_file)
            map_df.columns = map_df.columns.str.strip()
            
            map_notes_col = find_col_in(map_df, ['notes', 'note', 'الملاحظات', 'ملاحظات', 'feedback'])
            map_eval_col = [c for c in map_df.columns if c != map_notes_col][0] if len(map_df.columns) > 1 else map_df.columns[0]
            
            if map_notes_col:
                mapping_dict = dict(zip(map_df[map_notes_col].astype(str).str.strip(), map_df[map_eval_col]))
                
                all_sheet_notes = act_df[notes_col_name].dropna().astype(str).str.strip().unique()
                unmapped_notes = [n for n in all_sheet_notes if n not in mapping_dict]
                
                if unmapped_notes:
                    st.warning(f"⚠️ وجدنا ملاحظات جديدة ({len(unmapped_notes)} ملاحظة) غير موجودة في موديل التقييم المرفوع:")
                    
                    if 'custom_note_evals' not in st.session_state:
                        st.session_state.custom_note_evals = {}
                        
                    for un_note in unmapped_notes:
                        eval_choice = st.radio(f"كيف تحسبين الملاحظة الآتية: **({un_note})**؟", options=["ناجحة (1)", "مغطاة (0)"], key=f"eval_radio_{un_note}", horizontal=True)
                        st.session_state.custom_note_evals[un_note] = 1 if "ناجحة" in eval_choice else 0
                        
                def classify_with_model(val):
                    v_str = str(val).strip()
                    if v_str in mapping_dict:
                        val_eval = str(mapping_dict[v_str]).strip()
                        return 'ناجحة' if val_eval in ['1', '1.0', '1,0', 'true'] else 'مغطاة'
                    elif v_str in st.session_state.get('custom_note_evals', {}):
                        return 'ناجحة' if st.session_state.custom_note_evals[v_str] == 1 else 'مغطاة'
                    else:
                        return 'غير مصنف'
                        
                act_df['Call_Status'] = act_df[notes_col_name].apply(classify_with_model)
            else:
                act_df['Call_Status'] = 'غير مصنف'
        else:
            target_col = notes_col_name if notes_col_name else find_col_in(act_df, ['التقييم', 'status', 'الحالة', 'sub state', 'result'])
            if target_col:
                available_vals = act_df[target_col].dropna().astype(str).unique().tolist()
                selected_success_vals = st.multiselect(f"حددي القيم التي تعتبر (ناجحة = 1) من عمود [{target_col}]:", options=available_vals, default=[v for v in available_vals if str(v).strip() in ['1', '1.0', 'ناجح', 'وعد', 'سداد', 'تم']])
                act_df['Call_Status'] = act_df[target_col].astype(str).apply(lambda x: 'ناجحة' if x in selected_success_vals else 'مغطاة')
            else:
                act_df['Call_Status'] = 'غير مصنف'

        total_activity = len(act_df)
        success_count = len(act_df[act_df['Call_Status'] == 'ناجحة'])
        covered_count = len(act_df[act_df['Call_Status'] == 'مغطاة'])
        success_rate = (success_count / total_activity * 100) if total_activity > 0 else 0

        collector_col = collector_col_name
        if collector_col and collector_col in act_df.columns and total_activity > 0:
            top_collector_series = act_df[collector_col].value_counts()
            if not top_collector_series.empty:
                top_name = top_collector_series.index[0]
                top_calls = top_collector_series.values[0]
                top_pct = (top_calls / total_activity * 100) if total_activity > 0 else 0
                st.markdown(f"""
                <div style="background-color: {card_color}; border-radius: 12px; padding: 20px 30px; border: 2px solid #586F8C; direction:rtl; text-align: right; margin-bottom: 20px; display: flex; align-items: center; justify-content: space-between;">
                    <div>
                        <h4 style="margin:0; color: #586F8C; font-weight:900;">🏆 بطل الشفت (الأكثر نشاطاً)</h4>
                        <p style="margin: 5px 0 0 0; color: {text_color}; font-size: 18px;">المحصل: <b style="color:#1F4373; font-size: 22px;">{top_name}</b></p>
                    </div>
                    <div style="text-align: left; background-color: {ejada_navy_dark}; padding: 10px 20px; border-radius: 8px; border: 1px solid #586F8C;">
                        <h3 style="margin:0; color: #FFFFFF; font-weight:900;">{top_calls:,} مكالمة</h3>
                        <p style="margin: 0; color: #E2E8F0; font-size: 14px; font-weight: bold;">استحوذ على <span style="color:#2CA549;">{top_pct:.1f}%</span> من إجمالي شغل الفترة</p>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown(f"""
        <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 18px; direction: rtl;">
            <div class="kpi-card" style="border-right-color: {border_color};">
                <div class="kpi-icon">📞</div>
                <div class="kpi-title">إجمالي المحاولات</div>
                <div class="kpi-value">{total_activity:,}</div>
                <div class="kpi-sub">بعد إزالة التكرار</div>
            </div>
            <div class="kpi-card" style="border-right-color: #2CA549;">
                <div class="kpi-icon">{GREEN_CHECK_SVG}</div>
                <div class="kpi-title">المكالمات الناجحة (1)</div>
                <div class="kpi-value">{success_count:,}</div>
                <div class="kpi-sub" style="color: #2CA549;">تقييم ناجح</div>
            </div>
            <div class="kpi-card" style="border-right-color: #2CA549;">
                <div class="kpi-icon">{GREEN_PHONE_X_SVG}</div>
                <div class="kpi-title">المكالمات المغط
