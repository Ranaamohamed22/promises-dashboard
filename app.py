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

if 'current_page' not in st.session_state:
    st.session_state.current_page = 'Overview'
if 'theme_choice' not in st.session_state:
    st.session_state.theme_choice = 'ثيم إجادة (Navy)'

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
# إعدادات الثيم والألوان بالترتيب المطلوب
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
    sub_text_color = ejada_grey_blue
    gauge_bg = "#E2E8F0"
    nav_btn_bg = "#FFFFFF"
    nav_btn_text = ejada_navy_dark
    nav_btn_border = ejada_border_light
    kpi_title_color = "#586F8C"
    header_logo_color = "#FFFFFF"
    input_bg = "#FFFFFF"
    input_text = "#1F4373"
else:
    bg_color = "#FFFFFF"
    card_color = "#F3F4F6"
    border_color = "#9CA3AF"
    primary_color = ejada_navy_light
    primary_text = "#FFFFFF"
    text_color = "#1F2937" 
    sub_text_color = "#4B5563"
    gauge_bg = "#E2E8F0"
    nav_btn_bg = "#FFFFFF"
    nav_btn_text = "#1F2937"
    nav_btn_border = "#9CA3AF"
    kpi_title_color = "#4B5563"
    header_logo_color = "#1F4373"
    input_bg = "#FFFFFF"
    input_text = "#1F2937"

# ==========================================
# تصميم الـ UI الديناميكي (CSS)
# ==========================================
st.markdown(f"""
<style>
    [data-testid="collapsedControl"] {{display: none !important;}}
    [data-testid="stSidebar"] {{display: none !important;}}
    
    .stApp {{ background-color: {bg_color} !important; }}

    .block-container {{ padding-top: 3.8rem !important; padding-left: 1rem !important; padding-right: 1rem !important; max-width: 100% !important; direction: rtl; text-align: right; }}
    
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
        width: 38px !important;
        height: 38px !important;
        min-width: 38px !important;
        min-height: 38px !important;
        padding: 0 !important;
        margin-top: 1px !important; 
        color: transparent !important; 
        box-shadow: 0 2px 6px rgba(0,0,0,0.4) !important;
    }}

    .stTextInput input, div[data-baseweb="select"] > div, div[data-baseweb="base-input"], div[data-baseweb="tag"], div[data-baseweb="menu"] {{
        background-color: {input_bg} !important;
        color: {input_text} !important;
        border: 1px solid {border_color} !important;
        border-radius: 8px !important;
        direction: rtl;
        text-align: right;
    }}
    
    div[data-baseweb="select"] span, div[data-baseweb="select"] div, span[data-baseweb="tag"] span, div[data-baseweb="menu"] div {{
        color: {input_text} !important;
    }}

    .stTextInput input::placeholder {{ color: {sub_text_color} !important; }}
    
    [data-testid="stFileUploader"] > section {{ background-color: {input_bg} !important; border: 1px dashed {border_color} !important; border-radius: 8px !important; }}
    [data-testid="stFileUploader"] section * {{ color: {input_text} !important; }}

    .kpi-card {{ background-color: {card_color}; border-radius: 12px; padding: 20px; box-shadow: 0 4px 10px rgba(0,0,0,0.2); border: 1px solid {border_color}; border-right-width: 6px; border-right-style: solid; text-align: right; direction: rtl; position: relative; }}
    .kpi-title {{ color: {kpi_title_color}; font-size: 13px; text-transform: uppercase; font-weight: 800; margin-bottom: 10px; }}
    .kpi-value {{ color: {text_color}; font-size: 30px; font-weight: 900; margin-bottom: 5px; }}
    .kpi-sub {{ color: {sub_text_color}; font-size: 12px; font-weight: 600; }}
    .kpi-icon {{ position: absolute; top: 20px; left: 20px; font-size: 24px; opacity: 0.9; }}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 0. الصف العلوي: زر القمر (تبديل الثيم) + Overview
# ==========================================
top_cols = st.columns([0.5, 1.5, 6], gap="small")

with top_cols[0]:
    if st.button(" ", key="theme_toggle"): 
        if st.session_state.theme_choice == 'التصميم الأبيض (Light)':
            st.session_state.theme_choice = 'ثيم إجادة (Navy)'
        else:
            st.session_state.theme_choice = 'التصميم الأبيض (Light)'
        st.rerun()

with top_cols[1]:
    if st.button("📊 Overview", type="primary" if st.session_state.current_page == 'Overview' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Overview'
        st.rerun()

# ==========================================
# 1. الهيدر الرئيسي (إجادة في المنتصف + البنك الأهلي SNB ناحية اليسار تماماً)
# ==========================================
st.markdown(f"""
<div style="position: relative; width: 100%; min-height: 110px; margin-top: -50px; margin-bottom: 25px;">
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
nav_cols = st.columns([1, 1.5, 1.3, 1.3, 4.9], gap="small") 

with nav_cols[0]:
    if st.button("🏠 الرئيسية", type="primary" if st.session_state.current_page == 'Overview' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Overview'
        st.rerun()

with nav_cols[1]:
    if st.button("🤝 الوعود القائمة والمكسورة", type="primary" if st.session_state.current_page == 'Promises' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Promises'
        st.rerun()

with nav_cols[2]:
    if st.button("🚨 تقرير الإهمال", type="primary" if st.session_state.current_page == 'Neglect' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Neglect'
        st.rerun()

with nav_cols[3]:
    if st.button("📈 النشاط", type="primary" if st.session_state.current_page == 'Activity' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Activity'
        st.rerun()

with nav_cols[4]:
    st.markdown("")

st.markdown(f"<hr style='margin: 10px 0 15px 0; border-color: {border_color}; clear: both;'>", unsafe_allow_html=True)

# ==========================================
# دوال تقرير الوعود والإكسيل
# ==========================================
PIVOT_TEAM = 'Sales Team'
PIVOT_COLL = 'Salesperson'
PIVOT_NOTE = 'ملاحظات-التمويل'
PIVOT_ACC = 'Account Number'
PIVOT_AMT = 'Net Amount'
SUM_FMT = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'

def build_pivot(df, order='desc'):
    need = [PIVOT_TEAM, PIVOT_COLL, PIVOT_NOTE, PIVOT_ACC, PIVOT_AMT]
    if df is None or df.empty or any(c not in df.columns for c in need):
        return {'notes': ['NPL'], 'teams': [], 'grand': {'by_note': {'NPL': (0, 0.0)}, 'cnt': 0, 'amt': 0.0}}

    d = df[need].copy()
    d['_c'] = d[PIVOT_ACC].notna().astype(int)
    d[PIVOT_AMT] = pd.to_numeric(d[PIVOT_AMT], errors='coerce').fillna(0.0)
    d[PIVOT_TEAM] = d[PIVOT_TEAM].fillna('(فارغ)').astype(str)
    d[PIVOT_COLL] = d[PIVOT_COLL].fillna('(فارغ)').astype(str)
    d[PIVOT_NOTE] = d[PIVOT_NOTE].fillna('(فارغ)').astype(str).str.strip()

    def _ordered(values):
        vals = list(pd.unique(values))
        return sorted(vals, reverse=True) if order == 'desc' else vals

    notes = sorted(d[PIVOT_NOTE].unique(), reverse=True)
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
