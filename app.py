import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date, datetime, time, timedelta
import io
import base64
import os
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# إعدادات الصفحة الأساسية
# ==========================================
st.set_page_config(page_title="SNB Promises Dashboard", layout="wide", initial_sidebar_state="collapsed")

# ==========================================
# ثيم Streamlit الفاتح: بيخلّي كل الحقول والمربعات بيضاء
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
# الأيقونات والثيمات
# ==========================================
PHONE_ICON_COLOR = "#1F4373"

PHONE_HEADER_SVG = f"""<svg viewBox="0 0 24 24" width="38" height="38" fill="{PHONE_ICON_COLOR}" style="display:block; flex-shrink:0;">
  <path d="M6.62 10.79c1.44 2.83 3.76 5.14 6.59 6.59l2.2-2.2c.27-.27.67-.36 1.02-.24 1.12.37 2.33.57 3.57.57.55 0 1 .45 1 1V20c0 .55-.45 1-1 1-9.39 0-17-7.61-17-17 0-.55.45-1 1-1h3.5c.55 0 1 .45 1 1 0 1.25.2 2.45.57 3.57.11.35.03.74-.25 1.02l-2.2 2.2z"/>
</svg>"""

GREEN_CHECK_SVG = """<svg viewBox="0 0 100 100" width="26" height="26" style="display:block;">
  <path fill="#2CA549" d="M50 4 C54 4 57 9 61 10 C65 11 70 8 73 10 C77 13 76 19 79 22 C82 25 88 24 90 28 C92 31 89 36 90 40 C91 44 96 46 96 50 C96 54 91 56 90 60 C89 64 92 69 90 72 C88 76 82 75 79 78 C76 81 77 87 73 90 C70 92 65 89 61 90 C57 91 54 96 50 96 C46 96 43 91 39 90 C35 89 30 92 27 90 C23 87 24 81 21 78 C18 75 12 76 10 72 C8 69 11 64 10 60 C9 56 4 54 4 50 C4 46 9 44 10 40 C11 36 8 31 10 28 C12 24 18 25 21 22 C24 19 23 13 27 10 C30 8 35 11 39 10 C43 9 46 4 50 4 Z"/>
  <circle cx="50" cy="50" r="28" fill="#82D197"/>
  <path d="M38 51 L46 59 L63 42" fill="none" stroke="#FFFFFF" stroke-width="7" stroke-linecap="round" stroke-linejoin="round"/>
</svg>"""

GREEN_PHONE_X_SVG = """<svg viewBox="0 0 100 100" width="26" height="26" style="display:block;">
  <path d="M26 20 C17 20 12 28 12 37 C12 63 37 88 63 88 C72 88 80 83 80 74 C80 69 76 65 71 63 L61 60 C57 59 52 60 49 64 L45 68 C36 63 29 56 24 47 L28 43 C32 40 33 35 32 31 L29 21 C28 18 27 20 26 20 Z" fill="none" stroke="#2CA549" stroke-width="8" stroke-linecap="round" stroke-linejoin="round"/>
  <path d="M58 18 L84 44 M84 18 L58 44" fill="none" stroke="#2CA549" stroke-width="10" stroke-linecap="round"/>
</svg>"""

def format_12h(dt_series):
    return (
        dt_series.dt.strftime('%I:%M %p')
        .str.replace('AM', 'ص', regex=False)
        .str.replace('PM', 'م', regex=False)
    )

def format_dt_12h(dt_series):
    return (
        dt_series.dt.strftime('%Y-%m-%d %I:%M %p')
        .str.replace('AM', 'ص', regex=False)
        .str.replace('PM', 'م', regex=False)
    )

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

def find_col_in(target_df, candidates):
    if target_df is None or target_df.empty:
        return None
    for c in candidates:
        for col in target_df.columns:
            if col.lower().strip() == c.lower().strip():
                return col
    return None

ejada_navy_dark = "#1F4373"
ejada_navy_light = "#304D73"
ejada_grey_blue = "#586F8C"
ejada_border_light = "#99AABF"
ejada_soft_grey = "#F2F2F2"
white_bg = "#FFFFFF"  
snb_green = "#2CA549"  

CHART_PALETTE = ["#1F4373", "#304D73", "#586F8C", "#99AABF", "#F2F2F2"]

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
    kpi_title_color = "#1F4373"
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
    kpi_title_color = "#1F4373"
    header_logo_color = "#1F4373"
    input_bg = "#FFFFFF"
    input_text = "#1F4373"

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

    div[data-testid="stVerticalBlockBorderWrapper"] {{
        padding: 10px 14px !important;
        border-radius: 10px !important;
        border: 1px solid {border_color} !important;
        background-color: #FFFFFF !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"] > div {{
        gap: 0.4rem !important;
    }}

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
        font-weight: 800 !important;
    }}
    html body .stApp [data-testid="stExpander"] summary svg {{
        fill: {input_text} !important;
    }}
    html body .stApp [data-testid="stExpander"] details > div,
    html body .stApp [data-testid="stExpander"] details > div p,
    html body .stApp [data-testid="stExpander"] details > div label,
    html body .stApp [data-testid="stExpander"] details > div span {{
        color: {input_text} !important;
    }}

    html body .stApp [role="tab"],
    html body .stApp [role="tab"] * {{
        color: {label_color} !important;
    }}

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
        color: {input_text} !important;
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

    label[data-testid="stWidgetLabel"] p,
    div[data-testid="stCheckbox"] p {{
        color: {label_color} !important;
        font-weight: 800 !important;
        font-size: 14px !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"] label[data-testid="stWidgetLabel"] p,
    [data-testid="stExpander"] label[data-testid="stWidgetLabel"] p {{
        color: {input_text} !important;
    }}

    .kpi-grid-4 {{
        display: grid;
        grid-template-columns: repeat(4, minmax(180px, 270px));
        justify-content: center;
        gap: 14px;
        margin-bottom: 16px;
        direction: rtl;
    }}
    .kpi-grid-6 {{
        display: grid;
        grid-template-columns: repeat(6, 1fr);
        gap: 10px;
        margin-bottom: 20px;
        direction: rtl;
    }}
    .kpi-card {{
        background-color: {card_color};
        border-radius: 10px;
        padding: 8px 12px !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.12);
        border: 1px solid {border_color};
        border-right-width: 5px;
        border-right-style: solid;
        text-align: center !important;
        direction: rtl;
        position: relative;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        min-height: 68px !important;
    }}
    .kpi-title {{
        color: {kpi_title_color};
        font-size: 18px !important;
        font-weight: 900 !important;
        margin-bottom: 2px !important;
        line-height: 1.2 !important;
        text-align: center !important;
        width: 100%;
    }}
    .kpi-value {{
        color: {text_color};
        font-size: 22px !important;
        font-weight: 900 !important;
        margin-bottom: 1px !important;
        line-height: 1.1 !important;
        text-align: center !important;
        width: 100%;
    }}
    .kpi-sub {{
        color: #586F8C;
        font-size: 11px !important;
        font-weight: 700 !important;
        margin-top: 0 !important;
        text-align: center !important;
        width: 100%;
    }}
    .kpi-icon {{
        position: absolute;
        top: 8px;
        left: 10px;
        font-size: 20px;
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
# دوال تقرير الوعود والإكسيل (سليمة 100% وتفتح على جميع إصدارات Excel بدون أي أخطاء)
# ==========================================
PIVOT_TEAM = 'Sales Team'
PIVOT_COLL = 'Salesperson'
PIVOT_NOTE = 'ملاحظات-التمويل'
PIVOT_ACC = 'Account Number'
PIVOT_AMT = 'Net Amount'
FIXED_NOTES = ['NPL', 'DPD60', 'DPD30']
SUM_FMT = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'
_ILLEGAL_EXCEL_CHARS = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _clean_excel_str(val):
    if isinstance(val, str):
        return _ILLEGAL_EXCEL_CHARS.sub('', val)
    return val


def _normalize_note_bucket(val):
    """توحيد مسميات المحفظة (Dpd30 / DPD30 / 30Dpd -> DPD30) لمنع ظهور عمود مكرر بأصفار"""
    if val is None or (not isinstance(val, (list, tuple)) and pd.isna(val)):
        return '(فارغ)'
    s = _clean_excel_str(str(val)).strip()
    if not s or s.lower() == 'nan':
        return '(فارغ)'
    key = re.sub(r'[\s_\-]+', '', s).upper()
    if key == 'NPL':
        return 'NPL'
    if key in ('DPD60', '60DPD'):
        return 'DPD60'
    if key in ('DPD30', '30DPD'):
        return 'DPD30'
    return s


def build_pivot(df, order='desc'):
    need = [PIVOT_TEAM, PIVOT_COLL, PIVOT_NOTE, PIVOT_ACC, PIVOT_AMT]
    if df is None or df.empty or any(c not in df.columns for c in need):
        return {'notes': list(FIXED_NOTES), 'teams': [], 'grand': {'by_note': {n: (0, 0.0) for n in FIXED_NOTES}, 'cnt': 0, 'amt': 0.0}}

    d = df[need].copy()
    d['_c'] = d[PIVOT_ACC].notna().astype(int)
    d[PIVOT_AMT] = pd.to_numeric(d[PIVOT_AMT], errors='coerce').fillna(0.0)
    d[PIVOT_TEAM] = d[PIVOT_TEAM].fillna('(فارغ)').astype(str).apply(_clean_excel_str)
    d[PIVOT_COLL] = d[PIVOT_COLL].fillna('(فارغ)').astype(str).apply(_clean_excel_str)
    d[PIVOT_NOTE] = d[PIVOT_NOTE].apply(_normalize_note_bucket)

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
    if PIVOT_NOTE in d.columns:
        d[PIVOT_NOTE] = d[PIVOT_NOTE].apply(_normalize_note_bucket)
    for c in ['Amount', 'Net Amount', 'Payment']:
        if c in d.columns:
            d[c] = pd.to_numeric(d[c], errors='coerce')
    for c in ['Debit Date', 'Assign Date', 'Follow up Due Date', 'Follow up Last Date', 'Expected Closing']:
        if c in d.columns:
            d[c] = pd.to_datetime(d[c], errors='coerce')
    return d


def _write_df_sheet(ws, df):
    ws.sheet_view.rightToLeft = True
    if df is None or len(df.columns) == 0:
        return
    df = _prepare_for_output(df)
    bold = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    hdr_fill = PatternFill('solid', fgColor='1F4373')
    center = Alignment(horizontal='center', vertical='center')

    for j, c in enumerate(df.columns, 1):
        clean_col = _clean_excel_str(str(c))
        cell = ws.cell(row=1, column=j, value=clean_col)
        cell.font = bold
        cell.fill = hdr_fill
        cell.alignment = center
        ws.column_dimensions[get_column_letter(j)].width = min(40, max(14, len(clean_col) + 5))

    for i, row in enumerate(df.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            if v is None or (not isinstance(v, (list, tuple)) and pd.isna(v)):
                continue
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            elif hasattr(v, 'item') and not isinstance(v, (str, bytes)):
                v = v.item()
            elif isinstance(v, str):
                v = _clean_excel_str(v)
            cell = ws.cell(row=i, column=j, value=v)
            if isinstance(v, datetime):
                cell.number_format = 'yyyy-mm-dd'
            elif isinstance(v, float):
                cell.number_format = '#,##0.00'

    if len(df) > 0:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
        ws.freeze_panes = "A2"


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
    ws.cell(row=4, column=tc + 1, value='Total متبقي المديونية')
    ws.merge_cells(start_row=4, start_column=tc, end_row=5, end_column=tc)
    ws.merge_cells(start_row=4, start_column=tc + 1, end_row=5, end_column=tc + 1)
    ws.cell(row=5, column=1, value='اسم المشرف')
    ws.cell(row=5, column=2, value='اسم المحصل')

    def body(r, c, v, fmt=None, bold=False):
        cell = ws.cell(row=r, column=c, value=_clean_excel_str(v))
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
        a = body(start, 1, t['team'], bold=True)
        a.fill = light
        a.border = Border(left=thick, right=thin, top=thin, bottom=thin)
        for rr in range(start + 1, r):
            cell = ws.cell(row=rr, column=1, value=_clean_excel_str(t['team']))
            cell.font = Font(name='Calibri', bold=True, size=11)
            cell.alignment = center
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

    ws.freeze_panes = "C6"
    ws.column_dimensions['A'].width = 35.7
    ws.column_dimensions['B'].width = 30.7
    for c in range(3, 3 + 2 * n):
        ws.column_dimensions[get_column_letter(c)].width = 18.9
    ws.column_dimensions[get_column_letter(tc)].width = 16.7
    ws.column_dimensions[get_column_letter(tc + 1)].width = 17.6


def build_report_xlsx(active_df, broken_df, neglect_df=None):
    """بناء ملف الإكسيل بشكل نقي 100% عبر openpyxl ليفتح فوراً على جميع أجهزة وأنظمة الإكسيل بدون أي تلف"""
    wb = Workbook()
    ws = wb.active
    ws.title = 'تقرير الوعود القائمة'
    _write_df_sheet(ws, active_df)

    p_active = build_pivot(active_df, order='first')
    ws = wb.create_sheet('ملخص تقرير الوعود القائمة')
    _write_pivot_sheet(ws, p_active, 'تقرير الوعود القائمة البنك الاهلي السعودي SNB (NPL&60Dpd&30Dpd)')

    ws = wb.create_sheet('تقرير الوعود المكسورة')
    _write_df_sheet(ws, broken_df)

    p_broken = build_pivot(broken_df, order='desc')
    ws = wb.create_sheet('ملخص تقرير الوعود المكسورة')
    _write_pivot_sheet(ws, p_broken, 'تقرير الوعود المكسورة البنك الاهلي SNB (NPL&60Dpd&30Dpd)')

    if neglect_df is not None and len(neglect_df.columns) > 0:
        ws = wb.create_sheet('تقرير الإهمال')
        _write_df_sheet(ws, neglect_df)

    buf = io.BytesIO()
    wb.save(buf)
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
            if PIVOT_NOTE in temp_df.columns:
                temp_df[PIVOT_NOTE] = temp_df[PIVOT_NOTE].apply(_normalize_note_bucket)
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
# صفحة النشاط: العنوان مع لوجو التليفون الكامل + الخانات التلاتة
# ==========================================
def render_activity_header():
    st.markdown(f"""
    <div style="background-color: {card_color}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; margin-bottom: 20px; display:flex; align-items:center; justify-content:center; gap:14px;">
        {PHONE_HEADER_SVG}
        <div style="margin:0; color:{text_color}; font-weight:900; font-size:28px; line-height:1.3;">تقرير نشاط المحصلين</div>
    </div>
    """, unsafe_allow_html=True)

def render_activity_boxes(activity_base_df=None):
    col_model, col_file, col_coll = st.columns([1, 1, 1], gap="medium")

    with col_model:
        with st.container(border=True):
            st.markdown(f"<p style='font-size: 13px; color: {input_text}; direction: rtl; text-align: right; margin: 0 0 2px 0; font-weight: bold;'>التقييم (1) يعني 'ناجحة'، والتقييم (0) يعني 'مغطاة'.</p>", unsafe_allow_html=True)
            mapping_file = st.file_uploader("شيت موديل النشاط", type=["xlsx"], key="mapping_uploader")

    with col_file:
        with st.container(border=True):
            cur_files = st.session_state.page_files.get('Activity') or []
            if cur_files:
                cur_names = "، ".join(f.name for f in cur_files)
                hdr_col1, hdr_col2 = st.columns([8.5, 1.5])
                with hdr_col1:
                    st.markdown(f"<p style='font-size: 13px; color: {input_text}; direction: rtl; text-align: right; margin: 0; font-weight: bold;'>الملف الحالي: {cur_names}</p>", unsafe_allow_html=True)
                with hdr_col2:
                    st.button("✖", key="clear_act_files_btn", on_click=clear_activity_files, help="إزالة الملف المرفوع")
            else:
                st.markdown(f"<p style='font-size: 13px; color: {input_text}; direction: rtl; text-align: right; margin: 0 0 2px 0; font-weight: bold;'>الملف الحالي: لسه ما اترفعش ملف</p>", unsafe_allow_html=True)

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
            st.markdown(f"<p style='font-size: 13px; color: {input_text}; direction: rtl; text-align: right; margin: 0 0 2px 0; font-weight: bold;'>فلترة حسب المحصل</p>", unsafe_allow_html=True)
            collector_options = ["الكل (جميع المحصلين)"]
            if activity_base_df is not None:
                collector_col_name = find_col_in(activity_base_df, ['collecter', 'collector', 'المحصل', 'اسم المحصل'])
                if collector_col_name and collector_col_name in activity_base_df.columns:
                    unique_colls = activity_base_df[collector_col_name].dropna().astype(str).unique().tolist()
                    collector_options.extend(sorted(unique_colls))
            selected_collector = st.selectbox("👤 اسم المحصل (Collector):", options=collector_options, disabled=(activity_base_df is None))

    return selected_collector, mapping_file, collector_col_name

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

    if curr_page == 'Overview':
        st.markdown(f"<h2 style='text-align:center; margin:0; color:{label_color}; font-weight:900;'>📊 لوحة القيادة الشاملة (Executive Overview)</h2>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align:center; color:{sub_text_color}; font-weight: bold; margin-bottom: 30px;'>ملخص أداء محفظة الديون ونشاط المحصلين</p>", unsafe_allow_html=True)
        
        st.markdown(f"""
        <div class="kpi-grid-6">
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
            <div class="kpi-card" style="border-right-color: {snb_green};">
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
        st.download_button(label="📥 تحميل التقرير الكامل.xlsx", data=report_bytes, file_name="Daily_Promises_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

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
        st.download_button("📥 تحميل التقرير الشامل (الإكسيل بالملخصات).xlsx", data=report_bytes, file_name="Daily_Promises_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_promises_all")
        
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
                    if 'custom_note_evals' not in st.session_state:
                        st.session_state.custom_note_evals = {}
                    
                    with st.expander(f"⚠️ ملاحظات جديدة غير موجودة في الموديل ({len(unmapped_notes)} ملاحظة) - اضغطي هنا لفتحها أو قفلها وتحديد تقييمها", expanded=False):
                        for un_note in unmapped_notes:
                            eval_choice = st.radio(
                                f"كيف تحسبين الملاحظة الآتية من عمود [Notes]: **({un_note})**؟",
                                options=["ناجحة (1)", "مغطاة (0)"],
                                key=f"eval_radio_{un_note}",
                                horizontal=True
                            )
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
                with st.expander(f"📝 تحديد الملاحظات التي تعتبر (ناجحة = 1) من عمود [Notes] - اضغطي هنا للفتح أو القفل", expanded=False):
                    selected_success_vals = st.multiselect(
                        f"حددي القيم التي تعتبر (ناجحة = 1) من عمود [Notes]:",
                        options=available_vals,
                        default=[v for v in available_vals if str(v).strip() in ['1', '1.0', 'ناجح', 'وعد', 'سداد', 'تم']]
                    )
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
                <div style="background-color: {card_color}; border-radius: 12px; padding: 14px 22px; border: 2px solid #586F8C; direction:rtl; text-align: right; margin-bottom: 16px; display: flex; align-items: center; justify-content: space-between;">
                    <div>
                        <h4 style="margin:0; color: #586F8C; font-weight:900;">🏆 بطل الشفت (الأكثر نشاطاً)</h4>
                        <p style="margin: 5px 0 0 0; color: {text_color}; font-size: 18px;">المحصل: <b style="color:#1F4373; font-size: 22px;">{top_name}</b></p>
                    </div>
                    <div style="text-align: left; background-color: {ejada_navy_dark}; padding: 8px 18px; border-radius: 8px; border: 1px solid #586F8C;">
                        <h3 style="margin:0; color: #FFFFFF; font-weight:900;">{top_calls:,} مكالمة</h3>
                        <p style="margin: 0; color: #E2E8F0; font-size: 14px; font-weight: bold;">استحوذ على <span style="color:#2CA549;">{top_pct:.1f}%</span> من إجمالي شغل الفترة</p>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown(f"""
        <div class="kpi-grid-4">
            <div class="kpi-card" style="border-right-color: {border_color};">
                <div class="kpi-icon">📞</div>
                <div class="kpi-title">إجمالي المحاولات</div>
                <div class="kpi-value">{total_activity:,}</div>
                <div class="kpi-sub">بعد إزالة التكرار</div>
            </div>
            <div class="kpi-card" style="border-right-color: {snb_green};">
                <div class="kpi-icon">{GREEN_CHECK_SVG}</div>
                <div class="kpi-title">المكالمات الناجحة (1)</div>
                <div class="kpi-value">{success_count:,}</div>
                <div class="kpi-sub" style="color: {snb_green};">تقييم ناجح</div>
            </div>
            <div class="kpi-card" style="border-right-color: {snb_green};">
                <div class="kpi-icon">{GREEN_PHONE_X_SVG}</div>
                <div class="kpi-title">المكالمات المغطاة (0)</div>
                <div class="kpi-value">{covered_count:,}</div>
                <div class="kpi-sub">تقييم مغطاة</div>
            </div>
            <div class="kpi-card" style="border-right-color: #586F8C;">
                <div class="kpi-icon">🎯</div>
                <div class="kpi-title">نسبة النجاح الإجمالية</div>
                <div class="kpi-value">{success_rate:.1f}%</div>
                <div class="kpi-sub" style="color: #586F8C;">معدل الوصول للعملاء</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # ==========================================
        # الرسوم البيانية الثلاثة في صفحة النشاط
        # ==========================================
        if collector_col and collector_col in act_df.columns and total_activity > 0:
            ch_col1, ch_col2, ch_col3 = st.columns(3, gap="medium")

            # 1) الشارت الأولى: أكبر محصلين مكالمات
            with ch_col1:
                with st.container(border=True):
                    top_c = act_df[collector_col].value_counts().head(10).iloc[::-1]
                    bar_colors = [CHART_PALETTE[i % 4] for i in range(len(top_c))][::-1]
                    fig_top = go.Figure(go.Bar(
                        x=top_c.values,
                        y=top_c.index,
                        orientation='h',
                        marker=dict(color=bar_colors, line=dict(color="#1F4373", width=1)),
                        text=top_c.values,
                        textposition='auto',
                        textfont=dict(color="#FFFFFF", weight='bold', size=12)
                    ))
                    fig_top.update_layout(
                        title=dict(text="<b>📞 أكبر المحصلين مكالمات</b>", x=0.5, xanchor='center', font=dict(color="#1F4373", size=19, weight='bold')),
                        paper_bgcolor="#FFFFFF",
                        plot_bgcolor="#F2F2F2",
                        font=dict(color="#1F4373", family="Tahoma", size=12),
                        height=max(520, len(top_c) * 58 + 170),
                        margin=dict(l=20, r=20, t=65, b=45),
                        xaxis=dict(title="عدد المكالمات", gridcolor="#FFFFFF"),
                        yaxis=dict(title="", automargin=True)
                    )
                    st.plotly_chart(fig_top, use_container_width=True, config={'displayModeBar': False})

            # 2) الشارت الثانية: ميعاد أول وآخر إفادة (بالوقت فقط ونظام 12 ساعة) وبخط عريض Bold
            with ch_col2:
                with st.container(border=True):
                    if 'Created_DateTime' in act_df.columns and act_df['Created_DateTime'].notna().any():
                        dt_valid = act_df.dropna(subset=['Created_DateTime', collector_col]).sort_values('Created_DateTime')
                        first_rows = dt_valid.groupby(collector_col, as_index=False).first()
                        last_rows = dt_valid.groupby(collector_col, as_index=False).last()
                        fl_df = pd.merge(
                            first_rows[[collector_col, 'Created_DateTime'] + ([notes_col_name] if notes_col_name else [])],
                            last_rows[[collector_col, 'Created_DateTime'] + ([notes_col_name] if notes_col_name else [])],
                            on=collector_col,
                            suffixes=('_first', '_last')
                        )
                        top_names_list = act_df[collector_col].value_counts().head(10).index.tolist()
                        fl_df = fl_df[fl_df[collector_col].isin(top_names_list)].copy()
                        fl_df['first_hour'] = fl_df['Created_DateTime_first'].dt.hour + fl_df['Created_DateTime_first'].dt.minute / 60.0
                        fl_df['last_hour'] = fl_df['Created_DateTime_last'].dt.hour + fl_df['Created_DateTime_last'].dt.minute / 60.0
                        fl_df['first_time_12h'] = format_12h(fl_df['Created_DateTime_first'])
                        fl_df['last_time_12h'] = format_12h(fl_df['Created_DateTime_last'])
                        first_note_txt = fl_df[f"{notes_col_name}_first"].astype(str) if notes_col_name else pd.Series([""] * len(fl_df))
                        last_note_txt = fl_df[f"{notes_col_name}_last"].astype(str) if notes_col_name else pd.Series([""] * len(fl_df))

                        fig_fl = go.Figure()
                        fig_fl.add_trace(go.Bar(
                            name="أول إفادة",
                            y=fl_df[collector_col],
                            x=fl_df['first_hour'],
                            orientation='h',
                            marker_color="#99AABF",
                            text=fl_df['first_time_12h'],
                            textposition='outside',
                            textfont=dict(color="#1F4373", weight='bold', size=12),
                            cliponaxis=False,
                            customdata=list(zip(fl_df['first_time_12h'], first_note_txt)),
                            hovertemplate="<b>%{y}</b><br>أول إفادة: %{customdata[0]}<br>الإفادة: %{customdata[1]}<extra></extra>"
                        ))
                        fig_fl.add_trace(go.Bar(
                            name="آخر إفادة",
                            y=fl_df[collector_col],
                            x=fl_df['last_hour'],
                            orientation='h',
                            marker_color="#1F4373",
                            text=fl_df['last_time_12h'],
                            textposition='outside',
                            textfont=dict(color="#1F4373", weight='bold', size=12),
                            cliponaxis=False,
                            customdata=list(zip(fl_df['last_time_12h'], last_note_txt)),
                            hovertemplate="<b>%{y}</b><br>آخر إفادة: %{customdata[0]}<br>الإفادة: %{customdata[1]}<extra></extra>"
                        ))
                        fig_fl.update_layout(
                            barmode='group',
                            title=dict(text="<b>🕒 ميعاد أول وآخر إفادة</b>", x=0.5, xanchor='center', font=dict(color="#1F4373", size=19, weight='bold')),
                            paper_bgcolor="#FFFFFF",
                            plot_bgcolor="#F2F2F2",
                            font=dict(color="#1F4373", family="Tahoma", size=12),
                            height=max(520, len(fl_df) * 58 + 170),
                            margin=dict(l=20, r=65, t=65, b=45),
                            xaxis=dict(
                                title="الوقت (ساعات ودقائق - نظام 12 ساعة)",
                                range=[0, 27],
                                tickvals=[0, 4, 8, 12, 16, 20, 24],
                                ticktext=["12:00 ص", "4:00 ص", "8:00 ص", "12:00 م", "4:00 م", "8:00 م", "12:00 ص"],
                                gridcolor="#FFFFFF"
                            ),
                            yaxis=dict(title="", automargin=True),
                            legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="center", x=0.5)
                        )
                        st.plotly_chart(fig_fl, use_container_width=True, config={'displayModeBar': False})
                    else:
                        st.info("لا يوجد عمود تاريخ/وقت لحساب أول وآخر إفادة.")

            # 3) الشارت الثالثة: ميعاد أول إفادة بعد البريك (بنظام 12 ساعة)
            with ch_col3:
                with st.container(border=True):
                    if 'Created_DateTime' in act_df.columns and act_df['Created_DateTime'].notna().any():
                        after_break_df = act_df[act_df['Created_DateTime'].dt.time > time(12, 30)].dropna(subset=['Created_DateTime', collector_col]).sort_values('Created_DateTime')
                        if not after_break_df.empty:
                            first_ab = after_break_df.groupby(collector_col, as_index=False).first()
                            top_names_list = act_df[collector_col].value_counts().head(10).index.tolist()
                            first_ab = first_ab[first_ab[collector_col].isin(top_names_list)].copy()
                            first_ab['ab_hour'] = first_ab['Created_DateTime'].dt.hour + first_ab['Created_DateTime'].dt.minute / 60.0
                            first_ab['ab_time_12h'] = format_12h(first_ab['Created_DateTime'])
                            ab_note_txt = first_ab[notes_col_name].astype(str) if notes_col_name and notes_col_name in first_ab.columns else pd.Series(["-"] * len(first_ab))

                            ab_colors = [CHART_PALETTE[i % 4] for i in range(len(first_ab))]
                            fig_ab = go.Figure(go.Bar(
                                y=first_ab[collector_col],
                                x=first_ab['ab_hour'],
                                orientation='h',
                                marker=dict(color=ab_colors, line=dict(color="#1F4373", width=1)),
                                text=first_ab['ab_time_12h'],
                                textposition='outside',
                                textfont=dict(color="#1F4373", weight='bold', size=12),
                                cliponaxis=False,
                                customdata=list(zip(first_ab['ab_time_12h'], ab_note_txt)),
                                hovertemplate="<b>%{y}</b><br>ميعاد أول إفادة بعد البريك: %{customdata[0]}<br>الإفادة (Notes): %{customdata[1]}<extra></extra>"
                            ))
                            fig_ab.update_layout(
                                title=dict(text="<b>☕ ميعاد أول إفادة بعد البريك (12:30 م)</b>", x=0.5, xanchor='center', font=dict(color="#1F4373", size=19, weight='bold')),
                                paper_bgcolor="#FFFFFF",
                                plot_bgcolor="#F2F2F2",
                                font=dict(color="#1F4373", family="Tahoma", size=12),
                                height=max(520, len(first_ab) * 58 + 170),
                                margin=dict(l=20, r=65, t=65, b=45),
                                xaxis=dict(
                                    title="ميعاد أول إفادة (نظام 12 ساعة)",
                                    range=[12, 26],
                                    tickvals=[12, 14, 16, 18, 20, 22, 24],
                                    ticktext=["12:00 م", "2:00 م", "4:00 م", "6:00 م", "8:00 م", "10:00 م", "12:00 ص"],
                                    gridcolor="#FFFFFF"
                                ),
                                yaxis=dict(title="", automargin=True)
                            )
                            st.plotly_chart(fig_ab, use_container_width=True, config={'displayModeBar': False})
                        else:
                            st.info("لا توجد إفادات مسجلة بعد وقت البريك (12:30 م).")
                    else:
                        st.info("لا يوجد عمود تاريخ/وقت لعرض أول إفادة بعد البريك.")

        if collector_col and collector_col in act_df.columns:
            pivot_summary = act_df.groupby([collector_col, 'Call_Status']).size().unstack(fill_value=0)
            if 'ناجحة' not in pivot_summary.columns: pivot_summary['ناجحة'] = 0
            if 'مغطاة' not in pivot_summary.columns: pivot_summary['مغطاة'] = 0
            pivot_summary['إجمالي المحاولات'] = pivot_summary['ناجحة'] + pivot_summary['مغطاة']
            pivot_summary['نسبة النجاح %'] = (pivot_summary['ناجحة'] / pivot_summary['إجمالي المحاولات'] * 100).round(1).astype(str) + '%'
            pivot_summary = pivot_summary.sort_values(by='إجمالي المحاولات', ascending=False).reset_index()

            with st.expander("📊 تفصيل أداء المحصلين (ناجحة مقابل مغطاة)", expanded=True):
                st.dataframe(pivot_summary, use_container_width=True)

        st.markdown("<div style='text-align: right; direction: rtl;'>", unsafe_allow_html=True)
        extra_time_cols = [c for c in ['Date', 'Time', 'Call_Status'] if c in act_df.columns]
        display_cols = extra_time_cols + [c for c in act_df.columns if c not in ['Created_DateTime', 'Date', 'Time', 'Call_Status']]
        st.dataframe(act_df[display_cols], use_container_width=True, height=450)
        st.markdown("</div>", unsafe_allow_html=True)

        wb_act = Workbook()
        ws_act = wb_act.active
        ws_act.title = 'تقرير النشاط والتقييم'
        _write_df_sheet(ws_act, act_df[display_cols])
        act_buf = io.BytesIO()
        wb_act.save(act_buf)
        act_buf.seek(0)
        st.download_button(
            "📥 تحميل تقرير النشاط مفلتر.xlsx",
            data=act_buf.getvalue(),
            file_name="Activity_Success_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
