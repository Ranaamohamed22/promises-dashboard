import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date, datetime, time, timedelta
import io
import base64
import os
import re
import zipfile
from xml.sax.saxutils import escape as _xml_escape
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# إعدادات الصفحة الأساسية
# ==========================================
st.set_page_config(page_title="EJADA | SNB & STC Dashboard", layout="wide", initial_sidebar_state="collapsed")

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

BANKS = {'SNB': 'البنك الأهلي السعودي', 'STC': 'بنك STC'}
if 'selected_bank' not in st.session_state:
    st.session_state.selected_bank = 'SNB'

def set_bank(bank):
    st.session_state.selected_bank = bank

if 'current_page' not in st.session_state:
    st.session_state.current_page = 'Overview'
if 'theme_choice' not in st.session_state:
    st.session_state.theme_choice = 'ثيم إجادة (Navy)'
if 'act_uploader_ver' not in st.session_state:
    st.session_state.act_uploader_ver = {'SNB': 0, 'STC': 0}

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
        
    )

def format_dt_12h(dt_series):
    return (
        dt_series.dt.strftime('%Y-%m-%d %I:%M %p')
        
    )


def _clock_ticks(start_hour, end_hour, step=2):
    """Return 12-hour clock tick positions and AM/PM labels for hour-based axes."""
    ticks = list(range(int(start_hour), int(end_hour) + 1, int(step)))
    if not ticks or ticks[-1] != int(end_hour):
        ticks.append(int(end_hour))

    labels = []
    for hour in ticks:
        hour_24 = int(hour) % 24
        label = datetime(2000, 1, 1, hour_24, 0).strftime('%I:%M %p').lstrip('0')
        labels.append(label)
    return ticks, labels

THEMES = ['ثيم إجادة (Navy)', 'الوضع الداكن (Black)', 'التصميم الفاتح (Light)']

def toggle_theme():
    current = st.session_state.theme_choice
    st.session_state.theme_choice = THEMES[(THEMES.index(current) + 1) % len(THEMES)] if current in THEMES else THEMES[0]

def set_page(page_name):
    st.session_state.current_page = page_name

def clear_activity_files():
    st.session_state.page_files[st.session_state.selected_bank]['Activity'] = None
    st.session_state.act_uploader_ver[st.session_state.selected_bank] += 1

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
elif st.session_state.theme_choice == "الوضع الداكن (Black)":
    bg_color = "#101216"
    card_color = "#FFFFFF"
    border_color = "#596474"
    primary_color = "#1F4373"
    primary_text = "#FFFFFF"
    text_color = "#1F4373"
    label_color = "#FFFFFF"
    sub_text_color = "#D1D5DB"
    gauge_bg = "#E2E8F0"
    nav_btn_bg = "#FFFFFF"
    nav_btn_text = "#1F4373"
    nav_btn_border = "#596474"
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
    /* Remove the large white Streamlit header strip, but keep its toolbar/menu accessible. */
    header[data-testid="stHeader"] {{ background: transparent !important; height: 2.5rem !important; }}
    [data-testid="stDecoration"] {{ display: none !important; }}

    .block-container {{ padding-top: 1rem !important; padding-left: 1.5rem !important; padding-right: 1.5rem !important; max-width: 100% !important; direction: rtl; text-align: right; }}
    
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
        padding: 4px 6px !important;
        border-radius: 7px !important;
        border: 1px solid {border_color} !important;
        background-color: #FFFFFF !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"] > div {{
        gap: 0.3rem !important;
    }}
    div[data-testid="stFileUploaderDropzone"] {{
        min-height: 40px !important;
        padding: 2px 5px !important;
    }}
    div[data-testid="stFileUploaderDropzoneInstructions"] {{
        gap: 3px !important;
    }}
    /* Equal outer dimensions for Activity + Model and Promises upload cards. */
    div[data-testid="stVerticalBlockBorderWrapper"]:has([class*="st-key-overview_activity_model_"]) {{
        height: 230px !important;
        min-height: 230px !important;
        max-height: 230px !important;
        box-sizing: border-box !important;
        overflow: hidden !important;
        padding: 4px 6px !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"]:has([class*="st-key-overview_promises_"]) {{
        height: 230px !important;
        min-height: 230px !important;
        max-height: 230px !important;
        box-sizing: border-box !important;
        overflow: hidden !important;
        padding: 4px 6px !important;
    }}
    /* Make each uploader compact: button and file information share one row. */
    div[data-testid="stFileUploaderDropzone"] {{
        height: 58px !important;
        min-height: 58px !important;
        box-sizing: border-box !important;
        padding: 3px 6px !important;
    }}
    div[data-testid="stFileUploaderDropzone"] > div {{
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        align-items: center !important;
        justify-content: space-between !important;
        gap: 6px !important;
        width: 100% !important;
        min-height: 48px !important;
    }}
    div[data-testid="stFileUploaderDropzoneInstructions"] {{
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        align-items: center !important;
        white-space: nowrap !important;
        gap: 4px !important;
        min-width: 0 !important;
    }}
    div[data-testid="stFileUploaderDropzoneInstructions"] small {{
        white-space: nowrap !important;
        font-size: 10px !important;
    }}
    /* Keep the two main Overview upload cards compact and equal in size. */
    div[data-testid="stVerticalBlockBorderWrapper"]:has([class*="st-key-overview_activity_"]):has([class*="st-key-overview_activity_model_"]) {{
        min-height: 176px !important;
        box-sizing: border-box !important;
        padding: 4px 6px !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"]:has([class*="st-key-overview_promises_"]) {{
        min-height: 176px !important;
        box-sizing: border-box !important;
        padding: 4px 6px !important;
    }}
    div[data-testid="stFileUploader"] {{
        margin-bottom: 0 !important;
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
    /* ===== صفحة النشاط: 3 مربعات في سطر واحد - أصغر - والعناوين في النص - والحدود على القد ===== */
    html body .stApp div[class*="st-key-act_boxes_row"] {{
        display: grid !important;
        grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
        gap: 12px !important;
        width: 100% !important;
        margin: 0 auto 14px auto !important;
        align-items: stretch !important;
        direction: rtl !important;
    }}
    html body .stApp div[class*="st-key-act_box_"] {{
        border: 1px solid {border_color} !important;
        border-radius: 8px !important;
        padding: 6px 10px 8px 10px !important;
        min-height: 118px !important;
        height: auto !important;
        box-sizing: border-box !important;
        gap: 8px !important;
        text-align: center !important;
        height: 100% !important;
    }}
    html body .stApp div[class*="st-key-act_box_"] p {{
        line-height: 1.6 !important;
        margin: 0 0 4px 0 !important;
        padding: 0 !important;
        position: static !important;
        transform: none !important;
        text-align: center !important;
    }}
    html body .stApp div[class*="st-key-act_box_"] [data-testid="stElementContainer"],
    html body .stApp div[class*="st-key-act_box_"] [data-testid="stMarkdownContainer"] {{
        width: 100% !important;
        height: auto !important;
        min-height: 0 !important;
        overflow: visible !important;
    }}
    html body .stApp div[class*="st-key-act_box_"] [data-testid="stWidgetLabel"] {{
        width: 100% !important;
        display: flex !important;
        justify-content: center !important;
        margin-bottom: 4px !important;
        height: auto !important;
        min-height: 24px !important;
        line-height: 1.6 !important;
    }}
    html body .stApp div[class*="st-key-act_box_"] [data-testid="stWidgetLabel"] p {{
        width: 100% !important;
        text-align: center !important;
        margin: 0 !important;
        color: {label_color} !important;
    }}
    html body .stApp div[class*="st-key-act_box_"] div[data-testid="stFileUploaderDropzone"] {{
        height: 44px !important;
        min-height: 44px !important;
        padding: 2px 6px !important;
    }}
    html body .stApp div[class*="st-key-act_box_"] div[data-testid="stFileUploaderDropzone"] > div {{
        min-height: 36px !important;
    }}
    html body .stApp div[class*="st-key-act_box_"] div[data-baseweb="select"] > div {{
        min-height: 38px !important;
    }}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 0. الصف العلوي: زر القمر + Overview
# ==========================================
top_cols = st.columns([0.6, 1.6, 7.8], gap="small")

with top_cols[0]:
    st.button("🌓", key="theme_toggle", on_click=toggle_theme, help="تغيير الثيم (كحلي / أسود / أبيض)")

with top_cols[1]:
    st.button("📊 Overview", key="top_overview_btn", type="primary" if st.session_state.current_page == 'Overview' else "secondary", use_container_width=True, on_click=set_page, args=('Overview',))

# ==========================================
# 1. الهيدر الرئيسي
# ==========================================
# وضع اسم البنك على يسار الشاشة مباشرة، بعيداً عن انعكاس أعمدة RTL.
# الشعار في منتصف الصفحة، ولا نستخدم وسوماً غير مغلقة.
_bank_title = 'SNB البنك الأهلي السعودي' if st.session_state.selected_bank == 'SNB' else 'STC'
_bank_subtitle = '(30-60-NPL)' if st.session_state.selected_bank == 'SNB' else ''
st.markdown(f"""
<div class="ejada-header" style="position:relative;width:100%;height:72px;margin-top:-54px;direction:ltr;pointer-events:none;z-index:1;">
  <div style="position:absolute;left:12px;right:auto;top:9px;text-align:left;direction:ltr;color:{snb_green};font-weight:900;white-space:nowrap;">
    <div style="font-size:22px;">{_bank_title}</div>
    <div style="font-size:14px;">{_bank_subtitle}</div>
  </div>
  <div style="position:absolute;left:50%;top:0;transform:translateX(-50%);text-align:center;line-height:1.03;white-space:nowrap;">
    <div style="color:{header_logo_color};font-size:43px;font-weight:900;font-family:Tahoma,sans-serif;">إجـــادة</div>
    <div style="color:{header_logo_color};font-size:18px;letter-spacing:9px;font-weight:900;">EJADA</div>
    <div style="color:{header_logo_color};font-size:16px;font-weight:900;">لتحصيل الديون</div>
  </div>
</div>
""", unsafe_allow_html=True)

if st.session_state.current_page != 'Overview':
    bank_col1, bank_col2, _bank_spacer = st.columns([1.3, 1.3, 7.4], gap='small')
    with bank_col1:
        st.button('🏦 SNB | البنك الأهلي السعودي', key='bank_SNB',
                  type='primary' if st.session_state.selected_bank == 'SNB' else 'secondary',
                  use_container_width=True, on_click=set_bank, args=('SNB',))
    with bank_col2:
        st.button('🏦 STC', key='bank_STC',
                  type='primary' if st.session_state.selected_bank == 'STC' else 'secondary',
                  use_container_width=True, on_click=set_bank, args=('STC',))


# ==========================================
# 2. شريط التنقل
# ==========================================
nav_cols = st.columns([0.9, 1.5, 1.05, 0.85, 5.7], gap="small") 

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
# دوال تقرير الوعود والإكسيل (مع Pivot Table حقيقي وتفاعلي 100% يفتح فوراً على Excel)
# ==========================================
PIVOT_TEAM = 'Sales Team'
PIVOT_COLL = 'Salesperson'
PIVOT_NOTE = 'ملاحظات-التمويل'
PIVOT_ACC = 'Account Number'
PIVOT_AMT = 'Net Amount'
FIXED_NOTES = ['NPL', 'DPD60', 'DPD30']
SUM_FMT = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'
_ILLEGAL_XML = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')


def _clean_excel_str(val):
    if isinstance(val, str):
        return _ILLEGAL_XML.sub('', val)
    return val


def _xa(v):
    return _xml_escape(_ILLEGAL_XML.sub('', str(v)), {'"': '&quot;'})


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

    # نعرض الأعمدة التي لها بيانات فعلية بترتيب NPL ثم DPD60 ثم DPD30 حتى لا يظهر عمود وهمي بأصفار
    present_notes = list(d[PIVOT_NOTE].unique())
    ordered_fixed = [n for n in FIXED_NOTES if n in present_notes]
    extra_notes = sorted([x for x in present_notes if x not in FIXED_NOTES], reverse=True)
    notes = ordered_fixed + extra_notes if (ordered_fixed or extra_notes) else list(FIXED_NOTES)

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
    bold = Font(name='Calibri', bold=True)
    for j, c in enumerate(df.columns, 1):
        clean_col = _clean_excel_str(str(c))
        ws.cell(row=1, column=j, value=clean_col).font = bold
        ws.column_dimensions[get_column_letter(j)].width = min(40, max(12, len(clean_col) + 4))
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


def _write_pivot_sheet(ws, p, title, is_real_pivot=False):
    """كتابة شيت الملخص؛ عندما يكون is_real_pivot=True نضع عنوان التقرير فقط في الصفين 1-2 ونترك منطقة الـ Pivot بدون دمج خلايا متعارض"""
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
        if not is_real_pivot:
            ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + 1)
        ws.cell(row=5, column=c0, value='عدد الحسابات')
        ws.cell(row=5, column=c0 + 1, value='متبقي المديونية')
    tc = 3 + 2 * n
    ws.cell(row=4, column=tc, value='Total عدد الحسابات')
    ws.cell(row=4, column=tc + 1, value='Total متبقي المديونية')
    if not is_real_pivot:
        ws.merge_cells(start_row=4, start_column=tc, end_row=5, end_column=tc)
        ws.merge_cells(start_row=4, start_column=tc + 1, end_row=5, end_column=tc + 1)
    else:
        ws.cell(row=5, column=tc, value='Total عدد الحسابات')
        ws.cell(row=5, column=tc + 1, value='Total متبقي المديونية')
    ws.cell(row=5, column=1, value='اسم المشرف ')
    ws.cell(row=5, column=2, value='اسم المحصل ')

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
        a = body(start, 1, t['team'])
        a.fill = light
        a.border = Border(left=thick, right=thin, top=thin, bottom=thin)
        for rr in range(start + 1, r):
            cell = ws.cell(row=rr, column=1)
            cell.fill = light
            cell.border = Border(left=thick, right=thin, top=thin, bottom=thin)
        if not is_real_pivot and r - 1 > start:
            ws.merge_cells(start_row=start, start_column=1, end_row=r - 1, end_column=1)

    g = p['grand']
    for c in range(1, last_col + 1):
        cell = ws.cell(row=r, column=c)
        cell.font = white_b
        cell.fill = dark
        cell.alignment = center
        cell.border = Border(left=thick if c == 1 else thin, right=thin, top=thin, bottom=thick)
    ws.cell(row=r, column=1, value='الاجمالي')
    if not is_real_pivot:
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
    if isinstance(v, str):
        return _clean_excel_str(v)
    return v


def _pivot_parts(data_df, data_sheet, p, order, idx):
    """بناء XML سليم 100% للـ Pivot Table متوافق مع جميع نسخ Microsoft Excel"""
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
        vals = ['(فارغ)' if v is None else _normalize_note_bucket(v) if ci == note_i else _clean_excel_str(str(v)) for v in values[ci]]
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
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition {ns} r:id="rId1" refreshOnLoad="1" refreshedBy="Ejada" '
        f'refreshedDate="46330.5" createdVersion="3" refreshedVersion="3" minRefreshableVersion="3" recordCount="{n_rows}">'
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
        return f'<items count="{len(order_list) + 1}">' + ''.join(f'<item x="{shared_maps[ci][o]}"/>' for o in order_list) + '<item t="default"/></items>'

    pf = []
    for ci in range(len(cols)):
        base = 'compact="0" outline="0" showAll="0" defaultSubtotal="0"'
        if ci == note_i:
            pf.append(f'<pivotField axis="axisCol" compact="0" outline="0" showAll="0" defaultSubtotal="0" sortType="descending">{items_xml(notes, ci)}</pivotField>')
        elif ci == team_i:
            pf.append(f'<pivotField axis="axisRow" {base}>{items_xml(team_order, ci)}</pivotField>')
        elif ci == coll_i:
            st_attr = ' sortType="descending"' if order == 'desc' else ''
            pf.append(f'<pivotField axis="axisRow" {base}{st_attr}>{items_xml(coll_order, ci)}</pivotField>')
        elif ci == acc_i or ci == amt_i:
            pf.append('<pivotField dataField="1" compact="0" outline="0" showAll="0" defaultSubtotal="0"/>')
        else:
            pf.append('<pivotField compact="0" outline="0" showAll="0" defaultSubtotal="0"/>')

    row_items = []
    for t in p['teams']:
        for k, r in enumerate(t['rows']):
            tp = team_pos[t['team']]
            cp = coll_pos[r['coll']]
            if k == 0:
                row_items.append(f'<i><x v="{tp}"/><x v="{cp}"/></i>')
            else:
                row_items.append(f'<i r="1"><x v="{cp}"/></i>')
    row_items.append('<i t="grand"><x/></i>')

    col_items = []
    for k in range(n_notes):
        col_items.append(f'<i><x v="{k}"/><x v="0"/></i>')
        col_items.append('<i r="1" i="1"><x v="1"/></i>')
    col_items.append('<i t="grand"><x v="0"/></i>')
    col_items.append('<i t="grand" i="1"><x v="1"/></i>')

    n_body = len(row_items) - 1
    last_row = 5 + n_body + 1
    last_col = 2 + 2 * n_notes + 2
    ref = f'A3:{get_column_letter(last_col)}{last_row}'

    pivot_def = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'name="PivotTable{idx}" cacheId="{idx}" dataOnRows="0" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" '
        f'applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" grandTotalCaption="الاجمالي" '
        f'updatedVersion="3" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="3" indent="0" '
        f'compact="0" compactData="0" gridDropZones="1">'
        f'<location ref="{ref}" firstHeaderRow="1" firstDataRow="3" firstDataCol="2"/>'
        f'<pivotFields count="{len(cols)}">{"".join(pf)}</pivotFields>'
        f'<rowFields count="2"><field x="{team_i}"/><field x="{coll_i}"/></rowFields>'
        f'<rowItems count="{len(row_items)}">{"".join(row_items)}</rowItems>'
        f'<colFields count="2"><field x="{note_i}"/><field x="-2"/></colFields>'
        f'<colItems count="{len(col_items)}">{"".join(col_items)}</colItems>'
        f'<dataFields count="2">'
        f'<dataField name="عدد الحسابات" fld="{acc_i}" subtotal="count"/>'
        f'<dataField name="متبقي المديونية" fld="{amt_i}" subtotal="sum" numFmtId="3"/>'
        f'</dataFields>'
        f'<pivotTableStyleInfo name="PivotStyleMedium9" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
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

        target = rid_target[sheet_rid[sp['summary_sheet']]].lstrip('/')
        sheet_file = target if target.startswith('xl/') else 'xl/' + target
        sheet_dir, sheet_name = sheet_file.rsplit('/', 1)
        rels_path = f'{sheet_dir}/_rels/{sheet_name}.rels'

        rel = (f'<Relationship Id="rIdPv{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" '
               f'Target="../pivotTables/pivotTable{idx}.xml"/>')
        if rels_path in files:
            files[rels_path] = files[rels_path].decode('utf-8').replace('</Relationships>', rel + '</Relationships>').encode('utf-8')
        else:
            files[rels_path] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                + rel + '</Relationships>').encode('utf-8')

        new_rels += (f'<Relationship Id="rIdPc{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" '
                     f'Target="pivotCache/pivotCacheDefinition{idx}.xml"/>')
        pivot_caches_xml += f'<pivotCache cacheId="{idx}" r:id="rIdPc{idx}"/>'
        new_ct += (f'<Override PartName="/xl/pivotCache/pivotCacheDefinition{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>'
                   f'<Override PartName="/xl/pivotCache/pivotCacheRecords{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>'
                   f'<Override PartName="/xl/pivotTables/pivotTable{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>')

    if specs:
        wb_rels = wb_rels.replace('</Relationships>', new_rels + '</Relationships>')
        ct = ct.replace('</Types>', new_ct + '</Types>')
        if 'xmlns:r=' not in wb_xml:
            wb_xml = wb_xml.replace(
                '<workbook ',
                '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
                1
            )
        wb_xml = wb_xml.replace('</workbook>', f'<pivotCaches>{pivot_caches_xml}</pivotCaches></workbook>')
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


def build_report_xlsx(active_df, broken_df, neglect_df=None, bank="SNB"):
    wb = Workbook()
    ws = wb.active
    ws.title = 'تقرير الوعود القائمة'
    _write_df_sheet(ws, active_df)

    can_pv_active = _can_pivot(active_df)
    p_active = build_pivot(active_df, order='first')
    ws = wb.create_sheet('ملخص تقرير الوعود القائمة')
    _write_pivot_sheet(ws, p_active, f'تقرير الوعود القائمة {BANKS[bank]} {bank}', is_real_pivot=can_pv_active)

    ws = wb.create_sheet('تقرير الوعود المكسورة')
    _write_df_sheet(ws, broken_df)

    can_pv_broken = _can_pivot(broken_df)
    p_broken = build_pivot(broken_df, order='desc')
    ws = wb.create_sheet('ملخص تقرير الوعود المكسورة')
    _write_pivot_sheet(ws, p_broken, f'تقرير الوعود المكسورة {BANKS[bank]} {bank}', is_real_pivot=can_pv_broken)

    if neglect_df is not None and len(neglect_df.columns) > 0:
        ws = wb.create_sheet('تقرير الإهمال')
        _write_df_sheet(ws, neglect_df)

    buf = io.BytesIO()
    wb.save(buf)

    specs = []
    if can_pv_active:
        specs.append({'summary_sheet': 'ملخص تقرير الوعود القائمة', 'data_sheet': 'تقرير الوعود القائمة',
                      'data_df': active_df, 'pivot': p_active, 'order': 'first'})
    if can_pv_broken:
        specs.append({'summary_sheet': 'ملخص تقرير الوعود المكسورة', 'data_sheet': 'تقرير الوعود المكسورة',
                      'data_df': broken_df, 'pivot': p_broken, 'order': 'desc'})
    if specs:
        return inject_pivot_tables(buf.getvalue(), specs)
    buf.seek(0)
    return buf

# ==========================================
# إدارة الملفات لكل صفحة تقرير بشكل منفصل تماماً
# ==========================================
if 'page_files' not in st.session_state or not all(
    isinstance(st.session_state.page_files.get(b), dict) for b in BANKS
):
    old_files = st.session_state.get('page_files', {})
    st.session_state.page_files = {
        b: {page: (old_files.get(page) if b == 'SNB' and not isinstance(old_files.get('SNB'), dict) else None)
            for page in ('Overview', 'Promises', 'Neglect', 'Activity')}
        for b in BANKS
    }

# ملفات لوحة Overview مستقلة لكل بنك ولكل نوع تقرير.
if 'overview_uploads' not in st.session_state:
    st.session_state.overview_uploads = {
        b: {'Activity': [], 'Promises': [], 'ActivityModel': None} for b in BANKS
    }
else:
    for b in BANKS:
        st.session_state.overview_uploads.setdefault(
            b, {'Activity': [], 'Promises': [], 'ActivityModel': None}
        )
        st.session_state.overview_uploads[b].setdefault('Activity', [])
        st.session_state.overview_uploads[b].setdefault('Promises', [])
        st.session_state.overview_uploads[b].setdefault('ActivityModel', None)

if 'overview_uploader_versions' not in st.session_state:
    st.session_state.overview_uploader_versions = {
        b: {'Activity': 0, 'Promises': 0, 'ActivityModel': 0} for b in BANKS
    }
else:
    for b in BANKS:
        st.session_state.overview_uploader_versions.setdefault(
            b, {'Activity': 0, 'Promises': 0, 'ActivityModel': 0}
        )
        for report_kind in ('Activity', 'Promises', 'ActivityModel'):
            st.session_state.overview_uploader_versions[b].setdefault(report_kind, 0)

def clear_overview_upload(bank, report_kind):
    if report_kind == 'ActivityModel':
        st.session_state.overview_uploads[bank][report_kind] = None
    else:
        st.session_state.overview_uploads[bank][report_kind] = []
    st.session_state.overview_uploader_versions[bank][report_kind] += 1

selected_bank = st.session_state.selected_bank
bank_files = st.session_state.page_files[selected_bank]

curr_page = st.session_state.current_page

# دالة معالجة بيانات أي صفحة
def process_page_data(uploaded_files=None, bank="SNB", activity_files=None, promises_files=None):
    explicit_sources = activity_files is not None or promises_files is not None
    if explicit_sources:
        source_files = (
            [(f, 'activity') for f in (activity_files or [])] +
            [(f, 'promises') for f in (promises_files or [])]
        )
    else:
        source_files = [(f, None) for f in (uploaded_files or [])]

    if not source_files:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), None, 0, 0, 0, 0, 0, 0

    today = pd.Timestamp(date.today()).normalize()
    promises_dfs = []
    activity_dfs = []
    all_dfs = []

    for f, source_type in source_files:
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
            if source_type == 'promises':
                promises_dfs.append(temp_df)
            elif source_type == 'activity':
                activity_dfs.append(temp_df)
            else:
                if 'Follow up Due Date' in temp_df.columns and 'Follow up Last Date' in temp_df.columns:
                    promises_dfs.append(temp_df)
                has_created = any(c.lower().strip() in ['created on', 'created_on', 'createdon', 'تاريخ الإنشاء', 'وقت المعالجة', 'date', 'التاريخ'] for c in temp_df.columns)
                if has_created or not ('Follow up Due Date' in temp_df.columns):
                    activity_dfs.append(temp_df)
        except Exception as e:
            st.error(f"خطأ في قراءة الملف {f.name}: {e}")

    df_all = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
    df_promises = pd.concat(promises_dfs, ignore_index=True) if promises_dfs else pd.DataFrame()
    if activity_dfs:
        df_act_raw = pd.concat(activity_dfs, ignore_index=True)
    else:
        df_act_raw = pd.DataFrame() if explicit_sources else df_all.copy()

    created_col = find_col_in(df_act_raw, ['created on', 'created_on', 'createdon', 'تاريخ الإنشاء', 'وقت المعالجة', 'date', 'التاريخ'])
    if created_col:
        df_act_raw['Created_DateTime'] = pd.to_datetime(df_act_raw[created_col], errors='coerce')
        df_act_raw['Date'] = df_act_raw['Created_DateTime'].dt.date
        df_act_raw['Time'] = df_act_raw['Created_DateTime'].dt.time

    collector_col = find_col_in(df_act_raw, ['collecter', 'collector', 'المحصل', 'اسم المحصل', 'الموظف', 'user', 'agent'])
    if not collector_col and not explicit_sources:
        collector_col = find_col_in(df_all, ['collecter', 'collector', 'المحصل', 'اسم المحصل', 'الموظف', 'user', 'agent'])

    id_col = find_col_in(df_act_raw, ['id numer', 'id number', 'id_number', 'id', 'رقم الهوية', 'رقم العميل', 'account id'])
    notes_col = find_col_in(df_act_raw, ['notes', 'note', 'الملاحظات', 'ملاحظات', 'feedback'])

    dedup_cols = [c for c in [collector_col, id_col, notes_col] if c is not None and c in df_act_raw.columns]
    activity_base_df = df_act_raw.drop_duplicates(subset=dedup_cols).copy() if len(dedup_cols) == 3 else df_act_raw.copy()

    active_promises_df = pd.DataFrame()
    broken_promises_df = pd.DataFrame()
    neglect_df = pd.DataFrame()
    total_unhandled = 0

    if not df_promises.empty and all(c in df_promises.columns for c in ['Follow up Due Date', 'Follow up Last Date', 'حالة المعالجة - التمويل', 'تحديث فئة العميل - التمويل', 'Classification']):
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

    output_buffer = build_report_xlsx(active_promises_df, broken_promises_df, neglect_df, bank=bank)
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
    """3 مربعات في سطر واحد (الترتيب من اليمين): ملف الموديل - ارفعي الملف هنا - اسم المحصل"""
    cap_style = f"font-size: 12px; color: {label_color}; direction: rtl; text-align: center; margin: 0; font-weight: bold;"
    collector_col_name = None
    selected_collector = "الكل (جميع المحصلين)"

    with st.container(key="act_boxes_row"):
        # ---------- 1) شيت موديل النشاط ----------
        with st.container(key="act_box_model"):
            st.markdown(f"<p style='{cap_style}'>التقييم (1) يعني 'ناجحة'، والتقييم (0) يعني 'مغطاة'.</p>", unsafe_allow_html=True)
            activity_model_version = st.session_state.overview_uploader_versions[selected_bank]['ActivityModel']
            uploaded_mapping_file = st.file_uploader(
                "شيت موديل النشاط",
                type=["xlsx"],
                key=f"overview_activity_model_{selected_bank}_{activity_model_version}"
            )
            if uploaded_mapping_file is not None:
                st.session_state.overview_uploads[selected_bank]['ActivityModel'] = uploaded_mapping_file
            mapping_file = (
                uploaded_mapping_file
                if uploaded_mapping_file is not None
                else st.session_state.overview_uploads[selected_bank].get('ActivityModel')
            )

        # ---------- 2) ارفعي الملف هنا يا رنا ----------
        with st.container(key="act_box_file"):
            cur_files = bank_files.get('Activity') or []
            if cur_files:
                cur_names = "، ".join(f.name for f in cur_files)
                hdr_col1, hdr_col2 = st.columns([8.5, 1.5])
                with hdr_col1:
                    st.markdown(f"<p style='{cap_style}'>الملف الحالي: {cur_names}</p>", unsafe_allow_html=True)
                with hdr_col2:
                    st.button("✖", key="clear_act_files_btn", on_click=clear_activity_files, help="إزالة الملف المرفوع")
            else:
                st.markdown(f"<p style='{cap_style}'>الملف الحالي: لسه ما اترفعش ملف</p>", unsafe_allow_html=True)

            uploader_key = f"uploader_inpage_{selected_bank}_Activity_{st.session_state.act_uploader_ver[selected_bank]}"
            new_files = st.file_uploader("ارفعي الملف هنا يا رنا", type=["xlsx", "csv"], accept_multiple_files=True, key=uploader_key)
            if new_files:
                new_sig = [(f.name, f.size) for f in new_files]
                cur_sig = [(f.name, f.size) for f in cur_files]
                if new_sig != cur_sig:
                    bank_files['Activity'] = new_files
                    st.rerun()

        # ---------- 3) اسم المحصل ----------
        with st.container(key="act_box_coll"):
            st.markdown(f"<p style='{cap_style}'>فلترة حسب المحصل</p>", unsafe_allow_html=True)
            collector_options = ["الكل (جميع المحصلين)"]
            if activity_base_df is not None:
                collector_col_name = find_col_in(activity_base_df, ['collecter', 'collector', 'المحصل', 'اسم المحصل'])
                if collector_col_name and collector_col_name in activity_base_df.columns:
                    unique_colls = activity_base_df[collector_col_name].dropna().astype(str).unique().tolist()
                    collector_options.extend(sorted(unique_colls))
            selected_collector = st.selectbox("👤 اسم المحصل (Collector):", options=collector_options, disabled=(activity_base_df is None), key=f"collector_{selected_bank}")

    return selected_collector, mapping_file, collector_col_name

st.markdown(f"<div style='text-align:right; color:{label_color}; font-weight:900; font-size:18px; margin-bottom:12px'>🏦 تقارير {'البنك الأهلي السعودي SNB' if selected_bank == 'SNB' else 'STC'}</div>", unsafe_allow_html=True)

if curr_page == 'Overview':
    st.markdown(f"<h3 style='text-align:center;color:{label_color};font-weight:900;margin:6px 0 12px;'>اختاري البنك</h3>", unsafe_allow_html=True)
    overview_bank_cols = st.columns(2, gap='large')
    with overview_bank_cols[0]:
        st.button(
            '🏦 SNB | البنك الأهلي السعودي',
            key='overview_choose_SNB',
            type='primary' if selected_bank == 'SNB' else 'secondary',
            use_container_width=True,
            on_click=set_bank,
            args=('SNB',)
        )
    with overview_bank_cols[1]:
        st.button(
            '🏦 STC',
            key='overview_choose_STC',
            type='primary' if selected_bank == 'STC' else 'secondary',
            use_container_width=True,
            on_click=set_bank,
            args=('STC',)
        )

    overview_versions = st.session_state.overview_uploader_versions[selected_bank]
    upload_cols = st.columns([1, 1], gap='small')

    # العمود الأول: صندوق واحد يحتوي خانتين متجاورتين وواضحتي الاسم.
    with upload_cols[0]:
        with st.container(border=True):
            activity_model_cols = st.columns(2, gap='small')

            # الخانة الأولى: تقرير النشاط.
            with activity_model_cols[0]:
                st.markdown(
                    "<div style='text-align:center;font-weight:900;color:#1F4373;"
                    "font-size:14px;margin:0 0 3px;'>📈 1) تقرير النشاط</div>"
                    "",
                    unsafe_allow_html=True
                )
                activity_widget_files = st.file_uploader(
                    'ارفعي ملف تقرير النشاط في هذه الخانة',
                    type=['xlsx', 'csv'],
                    accept_multiple_files=True,
                    key=f"overview_activity_{selected_bank}_{overview_versions['Activity']}"
                )
                if activity_widget_files:
                    st.session_state.overview_uploads[selected_bank]['Activity'] = activity_widget_files

                saved_activity_files = st.session_state.overview_uploads[selected_bank]['Activity']
                if saved_activity_files:
                    st.caption('المرفوع: ' + '، '.join(f.name for f in saved_activity_files))
                    st.button(
                        'مسح تقرير النشاط',
                        key=f'clear_overview_activity_{selected_bank}',
                        on_click=clear_overview_upload,
                        args=(selected_bank, 'Activity')
                    )
                else:
                    st.caption('لم يُرفع تقرير النشاط بعد.')

            # الخانة الثانية: موديل النشاط.
            with activity_model_cols[1]:
                st.markdown(
                    "<div style='text-align:center;font-weight:900;color:#1F4373;"
                    "font-size:14px;margin:0 0 3px;'>🧩 2) موديل النشاط</div>"
                    "",
                    unsafe_allow_html=True
                )
                activity_model_file = st.file_uploader(
                    'ارفعي ملف موديل النشاط في هذه الخانة',
                    type=['xlsx'],
                    accept_multiple_files=False,
                    key=f"overview_activity_model_{selected_bank}_{overview_versions['ActivityModel']}"
                )
                if activity_model_file is not None:
                    st.session_state.overview_uploads[selected_bank]['ActivityModel'] = activity_model_file

                saved_activity_model = st.session_state.overview_uploads[selected_bank].get('ActivityModel')
                if saved_activity_model is not None:
                    st.caption(f'الموديل المرفوع: {saved_activity_model.name}')
                    st.button(
                        'مسح موديل النشاط',
                        key=f'clear_overview_activity_model_{selected_bank}',
                        on_click=clear_overview_upload,
                        args=(selected_bank, 'ActivityModel')
                    )
                else:
                    st.caption('لم يُرفع موديل النشاط بعد.')

    # العمود الثاني: تقرير الوعود والإهمال.
    with upload_cols[1]:
        with st.container(border=True):
            st.markdown("<div style='text-align:center;font-weight:900;color:#1F4373;font-size:14px;margin-bottom:3px;'>🤝 تقرير الوعود والإهمال (Promises & Neglect)</div>", unsafe_allow_html=True)
            promises_widget_files = st.file_uploader(
                'ارفعي تقرير الوعود والإهمال هنا',
                type=['xlsx', 'csv'],
                accept_multiple_files=True,
                key=f"overview_promises_{selected_bank}_{overview_versions['Promises']}"
            )
            if promises_widget_files:
                st.session_state.overview_uploads[selected_bank]['Promises'] = promises_widget_files
            saved_promises_files = st.session_state.overview_uploads[selected_bank]['Promises']
            if saved_promises_files:
                st.caption('الملفات المحفوظة: ' + '، '.join(f.name for f in saved_promises_files))
                st.button(
                    'مسح تقرير الوعود والإهمال',
                    key=f'clear_overview_promises_{selected_bank}',
                    on_click=clear_overview_upload,
                    args=(selected_bank, 'Promises')
                )
            else:
                st.caption('لم يتم رفع تقرير الوعود والإهمال بعد.')

    overview_activity_files = st.session_state.overview_uploads[selected_bank]['Activity']
    overview_promises_files = st.session_state.overview_uploads[selected_bank]['Promises']
    overview_combined_files = (overview_activity_files or []) + (overview_promises_files or [])
    bank_files['Overview'] = overview_combined_files if overview_combined_files else None

page_uploaded = bank_files[curr_page]

if page_uploaded is None:
    if curr_page == 'Overview':
        st.markdown(f"""
        <div style="background-color:{card_color};border:1px solid {border_color};border-radius:15px;padding:22px 18px;margin:8px 0 18px;text-align:center;direction:rtl;">
            <h2 style="color:{text_color};margin:0 0 8px;font-weight:900;">📊 لوحة التحكم الرئيسية | EJADA Overview</h2>
            <p style="color:{kpi_title_color};margin:0;font-weight:700;">اختاري البنك، ثم ارفعي تقرير النشاط وتقرير الوعود والإهمال لبناء لوحة المؤشرات.</p>
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
            files_in = st.file_uploader("ارفعي الملفات هنا ي رنا", type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed", key=f"uploader_{selected_bank}_{curr_page}")
            if files_in:
                bank_files[curr_page] = files_in
                st.rerun()
else:
    if curr_page == 'Overview':
        active_promises_df, broken_promises_df, neglect_df, activity_base_df, report_bytes, total_rows, total_unhandled, active_count, broken_count, neglect_count, total_activity_count = process_page_data(
            bank=selected_bank,
            activity_files=st.session_state.overview_uploads[selected_bank]['Activity'],
            promises_files=st.session_state.overview_uploads[selected_bank]['Promises']
        )
    else:
        active_promises_df, broken_promises_df, neglect_df, activity_base_df, report_bytes, total_rows, total_unhandled, active_count, broken_count, neglect_count, total_activity_count = process_page_data(bank_files[curr_page], bank=selected_bank)

    if curr_page == 'Overview':
        st.markdown(f"<h2 style='text-align:center; margin:4px 0 0; color:{label_color}; font-weight:900;'>📊 لوحة القيادة الشاملة | {selected_bank}</h2>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align:center; color:{sub_text_color}; font-weight: bold; margin-bottom: 22px;'>ملخص المكالمات والوعود والإهمال للبنك المختار</p>", unsafe_allow_html=True)

        if not st.session_state.overview_uploads[selected_bank]['Activity']:
            st.warning('ارفعي تقرير النشاط لإظهار مؤشر إجمالي المكالمات ورسوم أداء المحصلين.')
        if not st.session_state.overview_uploads[selected_bank]['Promises']:
            st.warning('ارفعي تقرير الوعود والإهمال لإظهار مؤشرات الوعود القائمة والمكسورة والإهمال.')

        st.markdown(f"""
        <div class="kpi-grid-4">
            <div class="kpi-card" style="border-right-color:#304D73;">
                <div class="kpi-icon">📞</div>
                <div class="kpi-title">إجمالي المكالمات</div>
                <div class="kpi-value">{total_activity_count:,}</div>
            </div>
            <div class="kpi-card" style="border-right-color:#1F4373;">
                <div class="kpi-icon">🤝</div>
                <div class="kpi-title">الوعود القائمة</div>
                <div class="kpi-value">{active_count:,}</div>
            </div>
            <div class="kpi-card" style="border-right-color:#586F8C;">
                <div class="kpi-icon">⚠️</div>
                <div class="kpi-title">الوعود المكسورة</div>
                <div class="kpi-value">{broken_count:,}</div>
            </div>
            <div class="kpi-card" style="border-right-color:#99AABF;">
                <div class="kpi-icon">🚨</div>
                <div class="kpi-title">إجمالي الإهمال</div>
                <div class="kpi-value">{neglect_count:,}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # الرسوم الرئيسية: النشاط، حالة الوعود، المحصلون في الوعود المكسورة، وتوزيع الإهمال.
        def _overview_clean_name(value):
            cleaned = re.sub(r'(?i)(?<![A-Z0-9])(?:SNB|NPL|II)(?![A-Z0-9])', ' ', str(value))
            cleaned = re.sub(r'[|]+', ' ', cleaned)
            return re.sub(r'\\s+', ' ', cleaned).strip(' |،-')

        overview_row1 = st.columns(2, gap='large')

        with overview_row1[0]:
            with st.container(border=True):
                collector_col = find_col_in(activity_base_df, ['collecter', 'collector', 'المحصل', 'اسم المحصل', 'الموظف', 'user', 'agent'])
                if collector_col and not activity_base_df.empty:
                    top_calls = activity_base_df[collector_col].dropna().astype(str).value_counts().head(10)
                    call_labels = [_overview_clean_name(n) for n in top_calls.index]
                    fig_calls = go.Figure(go.Bar(
                        x=top_calls.values, y=call_labels, orientation='h',
                        marker=dict(color='#304D73', line=dict(width=0)),
                        text=[f'<b>{int(v):,}</b>' for v in top_calls.values],
                        textposition='outside', cliponaxis=False,
                        textfont=dict(color='#1F4373', size=13, family='Tahoma')
                    ))
                    fig_calls.update_layout(
                        title=dict(text='<b>أكثر 10 محصلين حسب عدد المكالمات</b>', x=0.5, xanchor='center',
                                   font=dict(color='#1F4373', size=18, family='Tahoma')),
                        paper_bgcolor='white', plot_bgcolor='white',
                        font=dict(color='#1F4373', family='Tahoma'),
                        height=max(350, len(top_calls) * 34 + 115),
                        margin=dict(t=65, b=25, l=20, r=55), showlegend=False
                    )
                    fig_calls.update_xaxes(showticklabels=False, showgrid=False, zeroline=False, title='')
                    fig_calls.update_yaxes(autorange='reversed', showgrid=False, title='', tickfont=dict(size=12))
                    st.plotly_chart(fig_calls, use_container_width=True, config={'displayModeBar': False})
                else:
                    st.info('ارفعي تقرير النشاط الذي يحتوي على عمود اسم المحصل لعرض الرسم.')

        with overview_row1[1]:
            with st.container(border=True):
                promise_total = active_count + broken_count
                if promise_total > 0:
                    fig_promise_status = go.Figure(go.Pie(
                        labels=['الوعود القائمة', 'الوعود المكسورة'],
                        values=[active_count, broken_count],
                        hole=0.62,
                        marker=dict(colors=['#1F4373', '#586F8C'], line=dict(color='#FFFFFF', width=2)),
                        textinfo='label+percent',
                        textfont=dict(color='#FFFFFF', size=13, family='Tahoma'),
                        sort=False
                    ))
                    fig_promise_status.update_layout(
                        title=dict(text='<b>توزيع الوعود القائمة والمكسورة</b>', x=0.5, xanchor='center',
                                   font=dict(color='#1F4373', size=18, family='Tahoma')),
                        paper_bgcolor='white', plot_bgcolor='white',
                        font=dict(color='#1F4373', family='Tahoma'),
                        height=390, margin=dict(t=65, b=25, l=25, r=25), showlegend=False
                    )
                    st.plotly_chart(fig_promise_status, use_container_width=True, config={'displayModeBar': False})
                else:
                    st.info('لا توجد بيانات وعود كافية لعرض توزيع الوعود.')

        overview_row2 = st.columns(2, gap='large')

        with overview_row2[0]:
            with st.container(border=True):
                broken_collector_col = find_col_in(broken_promises_df, ['collecter', 'collector', 'المحصل', 'اسم المحصل', 'الموظف', 'user', 'agent'])
                if broken_collector_col and not broken_promises_df.empty:
                    top_broken = broken_promises_df[broken_collector_col].dropna().astype(str).value_counts().head(10)
                    broken_labels = [_overview_clean_name(n) for n in top_broken.index]
                    fig_broken = go.Figure(go.Bar(
                        x=top_broken.values, y=broken_labels, orientation='h',
                        marker=dict(color='#586F8C', line=dict(width=0)),
                        text=[f'<b>{int(v):,}</b>' for v in top_broken.values],
                        textposition='outside', cliponaxis=False,
                        textfont=dict(color='#1F4373', size=13, family='Tahoma')
                    ))
                    fig_broken.update_layout(
                        title=dict(text='<b>أكثر المحصلين في الوعود المكسورة</b>', x=0.5, xanchor='center',
                                   font=dict(color='#1F4373', size=18, family='Tahoma')),
                        paper_bgcolor='white', plot_bgcolor='white',
                        font=dict(color='#1F4373', family='Tahoma'),
                        height=max(350, len(top_broken) * 34 + 115),
                        margin=dict(t=65, b=25, l=20, r=55), showlegend=False
                    )
                    fig_broken.update_xaxes(showticklabels=False, showgrid=False, zeroline=False, title='')
                    fig_broken.update_yaxes(autorange='reversed', showgrid=False, title='', tickfont=dict(size=12))
                    st.plotly_chart(fig_broken, use_container_width=True, config={'displayModeBar': False})
                else:
                    st.info('لا توجد بيانات محصلين داخل تقرير الوعود المكسورة لعرض الرسم.')

        with overview_row2[1]:
            with st.container(border=True):
                neglect_group_col = find_col_in(neglect_df, ['Sub State', 'sub state', 'substate', 'الحالة الفرعية', 'حالة العميل'])
                if not neglect_group_col:
                    neglect_group_col = find_col_in(neglect_df, ['collecter', 'collector', 'المحصل', 'اسم المحصل', 'الموظف'])
                if neglect_group_col and not neglect_df.empty:
                    top_neglect = neglect_df[neglect_group_col].dropna().astype(str).value_counts().head(10)
                    neglect_labels = [_overview_clean_name(n) for n in top_neglect.index]
                    fig_neglect = go.Figure(go.Bar(
                        x=top_neglect.values, y=neglect_labels, orientation='h',
                        marker=dict(color='#000000', line=dict(color='rgba(0,0,0,0)', width=0)),
                        text=[f'<b>{int(v):,}</b>' for v in top_neglect.values],
                        textposition='outside', cliponaxis=False,
                        textfont=dict(color='#1F4373', size=13, family='Tahoma')
                    ))
                    fig_neglect.update_layout(
                        title=dict(text=f'<b>توزيع الإهمال حسب {neglect_group_col}</b>', x=0.5, xanchor='center',
                                   font=dict(color='#1F4373', size=18, family='Tahoma')),
                        paper_bgcolor='white', plot_bgcolor='white',
                        font=dict(color='#1F4373', family='Tahoma'),
                        height=max(350, len(top_neglect) * 34 + 115),
                        margin=dict(t=65, b=25, l=20, r=55), showlegend=False
                    )
                    fig_neglect.update_xaxes(showticklabels=False, showgrid=False, zeroline=False, title='')
                    fig_neglect.update_yaxes(autorange='reversed', showgrid=False, title='', tickfont=dict(size=12))
                    st.plotly_chart(fig_neglect, use_container_width=True, config={'displayModeBar': False})
                else:
                    st.info('لا توجد بيانات كافية لعرض توزيع حالات الإهمال.')

        st.markdown("<br>", unsafe_allow_html=True)
        if report_bytes:
            st.download_button(label="📥 تحميل التقرير الكامل.xlsx", data=report_bytes, file_name=f"{selected_bank}_Daily_Promises_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

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
        st.download_button("📥 تحميل التقرير الشامل (الإكسيل بالملخصات).xlsx", data=report_bytes, file_name=f"{selected_bank}_Daily_Promises_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_promises_all_{selected_bank}")
        
    elif curr_page == 'Neglect':
        st.markdown(f"""
        <div style="background-color: {card_color}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; text-align: right; margin-bottom: 20px;">
            <h3 style="margin:0; color:#586F8C; font-weight:900;">🚨 تقرير الإهمال (تأخير أكثر من 3 أيام)</h3>
        </div>
        """, unsafe_allow_html=True)
        default_sub_states = ['وعد بسداد مبلغ المعالجة', 'واعد بسداد مبلغ المعالجه', 'رافض السداد', 'الجوال مغلق', 'مماطل', 'تم سداد جزء وليس مبلغ المعالجه', 'لا يجيب', 'متقاعد الراتب علي بنك آخر', 'يرغب بجدوله المديونيه']
        selected_sub_states = st.multiselect("فلاتر الـ Sub State المشمولة:", options=default_sub_states, default=default_sub_states, key=f"neglect_states_{selected_bank}")
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
            time_range_mode = st.selectbox("📅 مدى التقرير:", ["يومي", "أسبوعي", "شهري"], key=f"time_range_{selected_bank}")
        with filter_act_col2:
            shift_option = st.selectbox("⏰ فترة النشاط (الشفت):", ["كل اليوم بدون تحديد", "من 2:00 PM إلى 12:00 منتصف الليل (14:00 - 00:00)", "من 12:00 PM إلى 4:00 PM (12:00 - 16:00)", "من 4:00 PM إلى آخر اليوم (16:00 - 23:59)"], key=f"shift_{selected_bank}")
        with filter_act_col3:
            exclude_last_hour = st.checkbox("⏳ استبعاد آخر ساعة", value=False, key=f"last_hour_{selected_bank}")
        with filter_act_col4:
            exclude_break = st.checkbox("☕ استبعاد البريك (12:00 - 12:30)", value=False, key=f"break_{selected_bank}")

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
            if shift_option == "من 2:00 PM إلى 12:00 منتصف الليل (14:00 - 00:00)":
                start_t, end_t = time(14, 0), time(23, 0) if exclude_last_hour else time(23, 59, 59)
                act_df = act_df[(act_times >= start_t) & (act_times <= end_t)]
            elif shift_option == "من 12:00 PM إلى 4:00 PM (12:00 - 16:00)":
                start_t, end_t = time(12, 0), time(15, 0) if exclude_last_hour else time(16, 0)
                act_df = act_df[(act_times >= start_t) & (act_times <= end_t)]
            elif shift_option == "من 4:00 PM إلى آخر اليوم (16:00 - 23:59)":
                start_t, end_t = time(16, 0), time(23, 0) if exclude_last_hour else time(23, 59, 59)
                act_df = act_df[(act_times >= start_t) & (act_times <= end_t)]
            if exclude_break:
                curr_times = act_df['Created_DateTime'].dt.time
                act_df = act_df[~((curr_times >= time(12, 0)) & (curr_times <= time(12, 30)))]

        notes_col_name = find_col_in(act_df, ['notes', 'note', 'الملاحظات', 'ملاحظات', 'feedback'])
        
        if mapping_file is not None and notes_col_name:
            try:
                mapping_file.seek(0)
            except Exception:
                pass
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
                        st.session_state.custom_note_evals = {'SNB': {}, 'STC': {}}
                    
                    with st.expander(f"⚠️ ملاحظات جديدة غير موجودة في الموديل ({len(unmapped_notes)} ملاحظة) - اضغطي هنا لفتحها أو قفلها وتحديد تقييمها", expanded=False):
                        for un_note in unmapped_notes:
                            eval_choice = st.radio(
                                f"كيف تحسبين الملاحظة الآتية من عمود [Notes]: **({un_note})**؟",
                                options=["ناجحة (1)", "مغطاة (0)"],
                                key=f"eval_radio_{selected_bank}_{un_note}",
                                horizontal=True
                            )
                            st.session_state.custom_note_evals[selected_bank][un_note] = 1 if "ناجحة" in eval_choice else 0
                        
                def classify_with_model(val):
                    v_str = str(val).strip()
                    if v_str in mapping_dict:
                        val_eval = str(mapping_dict[v_str]).strip()
                        return 'ناجحة' if val_eval in ['1', '1.0', '1,0', 'true'] else 'مغطاة'
                    elif v_str in st.session_state.get('custom_note_evals', {}).get(selected_bank, {}):
                        return 'ناجحة' if st.session_state.custom_note_evals[selected_bank][v_str] == 1 else 'مغطاة'
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
                        default=[v for v in available_vals if str(v).strip() in ['1', '1.0', 'ناجح', 'وعد', 'سداد', 'تم']], key=f"success_vals_{selected_bank}"
                    )
                act_df['Call_Status'] = act_df[target_col].astype(str).apply(lambda x: 'ناجحة' if x in selected_success_vals else 'مغطاة')
            else:
                act_df['Call_Status'] = 'غير مصنف'

        total_activity = len(act_df)
        success_count = len(act_df[act_df['Call_Status'] == 'ناجحة'])
        covered_count = len(act_df[act_df['Call_Status'] == 'مغطاة'])
        success_rate = (success_count / total_activity * 100) if total_activity > 0 else 0

        collector_col = collector_col_name
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
                <div class="kpi-title">المكالمات الناجحة</div>
                <div class="kpi-value">{success_count:,}</div>
                <div class="kpi-sub" style="color: {snb_green};">تقييم ناجح</div>
            </div>
            <div class="kpi-card" style="border-right-color: {snb_green};">
                <div class="kpi-icon">{GREEN_PHONE_X_SVG}</div>
                <div class="kpi-title">المكالمات المغطاة</div>
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
        # الرسوم البيانية المطلوبة لتقرير النشاط
        # ==========================================
        if collector_col and collector_col in act_df.columns and total_activity > 0:

            def _clean_collector_name(full_name):
                value = str(full_name)
                value = re.sub(r"(?i)(?<![A-Z0-9])(?:SNB|NPL|II)(?![A-Z0-9])", " ", value)
                value = re.sub(r"[|]+", " ", value)
                value = re.sub(r"\s+", " ", value).strip(" |،-")
                return value or str(full_name)

            def _fmt_hour_12_ar(dt_value):
                hour_24 = int(dt_value.hour)
                minute = int(dt_value.minute)
                suffix = "ص" if hour_24 < 12 else "م"
                hour_12 = hour_24 % 12
                if hour_12 == 0:
                    hour_12 = 12
                return f"{hour_12}:{minute:02d} {suffix}"

            def _break_overlap_minutes(start_dt, end_dt):
                if pd.isna(start_dt) or pd.isna(end_dt) or end_dt <= start_dt:
                    return 0.0
                total_overlap = 0.0
                start_day = start_dt.normalize()
                end_day = end_dt.normalize()
                current_day = start_day
                while current_day <= end_day:
                    break_start = current_day + pd.Timedelta(hours=12)
                    break_end = current_day + pd.Timedelta(hours=12, minutes=30)
                    overlap_start = max(start_dt, break_start)
                    overlap_end = min(end_dt, break_end)
                    if overlap_end > overlap_start:
                        total_overlap += (overlap_end - overlap_start).total_seconds() / 60.0
                    current_day += pd.Timedelta(days=1)
                return total_overlap

            # ------------------------------------------
            # تجهيز بيانات النشاط
            # ------------------------------------------
            activity_chart_df = act_df.copy()
            activity_chart_df = activity_chart_df.dropna(subset=[collector_col]).copy()
            activity_chart_df['_collector_name'] = activity_chart_df[collector_col].astype(str).apply(_clean_collector_name)
            if 'Call_Status' in activity_chart_df.columns:
                activity_chart_df['_call_status'] = activity_chart_df['Call_Status'].fillna('غير مصنف').astype(str)
            else:
                activity_chart_df['_call_status'] = 'غير مصنف'

            # ------------------------------------------
            # الشارت الأولى: جودة المكالمات لكل محصل - Column Chart
            # ------------------------------------------
            quality_table = (
                activity_chart_df.groupby(['_collector_name', '_call_status'])
                .size()
                .unstack(fill_value=0)
            )
            if 'ناجحة' not in quality_table.columns:
                quality_table['ناجحة'] = 0
            if 'مغطاة' not in quality_table.columns:
                quality_table['مغطاة'] = 0

            # تضمين جميع المحصلين، وليس أعلى 10 فقط.
            quality_table['إجمالي'] = quality_table['ناجحة'] + quality_table['مغطاة']
            quality_table = quality_table.sort_values('إجمالي', ascending=False)

            if not quality_table.empty:
                quality_names = quality_table.index.tolist()
                quality_success = quality_table['ناجحة'].astype(int).tolist()
                quality_covered = quality_table['مغطاة'].astype(int).tolist()

                # حجز مساحة بجانب الشارت لمخطط توزيع العملاء الفعالين وغير الفعالين.
                quality_col, customer_pie_col = st.columns([2.15, 1], gap='small')

                with quality_col:
                    fig_quality = go.Figure()
                    fig_quality.add_trace(go.Bar(
                        x=quality_names,
                        y=quality_success,
                        name='المكالمات الناجحة',
                        marker=dict(color='#1F4373', line=dict(width=0)),
                        text=quality_success,
                        textposition='inside',
                        insidetextanchor='middle',
                        hovertemplate='<b>%{x}</b><br>ناجحة: %{y}<extra></extra>'
                    ))
                    fig_quality.add_trace(go.Bar(
                        x=quality_names,
                        y=quality_covered,
                        name='المكالمات المغطاة',
                        marker=dict(color='#99AABF', line=dict(width=0)),
                        text=quality_covered,
                        textposition='inside',
                        insidetextanchor='middle',
                        hovertemplate='<b>%{x}</b><br>مغطاة: %{y}<extra></extra>'
                    ))
                    fig_quality.update_layout(
                        title=dict(
                            text='<b>جودة المكالمات: ناجحة مقابل مغطاة لكل محصل</b>',
                            x=0.5, xanchor='center',
                            font=dict(size=17, color='#1F4373', family='Tahoma')
                        ),
                        barmode='stack',
                        height=390,
                        autosize=True,
                        paper_bgcolor='#FFFFFF',
                        plot_bgcolor='#FFFFFF',
                        font=dict(family='Tahoma', color='#1F4373', size=10),
                        margin=dict(l=12, r=12, t=65, b=105),
                        legend=dict(
                            orientation='h', yanchor='bottom', y=1.01,
                            xanchor='center', x=0.5,
                            font=dict(size=10)
                        ),
                        xaxis=dict(
                            title='',
                            showgrid=False,
                            showline=False,
                            linecolor='rgba(0,0,0,0)',
                            linewidth=0,
                            zeroline=False,
                            tickangle=-50,
                            tickmode='array',
                            tickvals=quality_names,
                            ticktext=[f'<b>{str(name)}</b>' for name in quality_names],
                            tickfont=dict(size=9, color='#1F4373', family='Tahoma'),
                            automargin=True,
                            categoryorder='array',
                            categoryarray=quality_names
                        ),
                        yaxis=dict(
                            title='عدد المكالمات',
                            showgrid=True,
                            gridcolor='#E8EDF3',
                            showline=False,
                            linecolor='rgba(0,0,0,0)',
                            linewidth=0,
                            zeroline=False,
                            tickfont=dict(size=9)
                        ),
                        bargap=0.24
                    )
                    with st.container(border=True):
                        st.plotly_chart(
                            fig_quality,
                            use_container_width=True,
                            config={'displayModeBar': False}
                        )

                with customer_pie_col:
                    with st.container(border=True):
                        st.markdown(
                            "<div style='text-align:center;color:#1F4373;"
                            "font-weight:900;font-size:16px;margin:8px 0 4px;'>"
                            "توزيع العملاء حسب الحالة الرئيسية</div>"
                            "<div style='text-align:center;color:#586F8C;"
                            "font-size:11px;margin-bottom:8px;'>Distinct Count of Debitor</div>",
                            unsafe_allow_html=True
                        )

                        debitor_col = find_col_in(
                            activity_chart_df,
                            ['Debitor', 'Debitor Number', 'Debitor No', 'Debtor', 'Debtor Number',
                             'Debtor No', 'رقم المدين', 'رقم العميل']
                        )
                        main_state_col = find_col_in(
                            activity_chart_df,
                            ['Main State', 'Main_State', 'main state', 'main_state',
                             'MainState', 'الحالة الرئيسية']
                        )

                        if debitor_col and main_state_col:
                            customer_state_df = activity_chart_df[
                                [debitor_col, main_state_col]
                            ].copy()
                            customer_state_df[debitor_col] = (
                                customer_state_df[debitor_col].astype(str).str.strip()
                            )
                            customer_state_df[main_state_col] = (
                                customer_state_df[main_state_col].astype(str).str.strip()
                            )
                            customer_state_df = customer_state_df[
                                customer_state_df[debitor_col].ne('')
                                & customer_state_df[debitor_col].ne('nan')
                                & customer_state_df[main_state_col].ne('')
                                & customer_state_df[main_state_col].ne('nan')
                            ].drop_duplicates(subset=[debitor_col, main_state_col])

                            customer_state_counts = (
                                customer_state_df.groupby(main_state_col)[debitor_col]
                                .nunique()
                                .sort_values(ascending=False)
                            )

                            if not customer_state_counts.empty:
                                customer_palette = [
                                    '#1F4373', '#304D73', '#586F8C',
                                    '#99AABF', '#F2F2F2'
                                ]
                                pie_colors = (
                                    customer_palette
                                    * ((len(customer_state_counts) // len(customer_palette)) + 1)
                                )[:len(customer_state_counts)]

                                fig_customers = go.Figure(data=[go.Pie(
                                    labels=customer_state_counts.index.astype(str).tolist(),
                                    values=customer_state_counts.astype(int).tolist(),
                                    hole=0.28,
                                    sort=False,
                                    marker=dict(
                                        colors=pie_colors,
                                        line=dict(color='#FFFFFF', width=1)
                                    ),
                                    textinfo='percent',
                                    textposition='inside',
                                    textfont=dict(
                                        family='Tahoma',
                                        size=11,
                                        color='#FFFFFF'
                                    ),
                                    hovertemplate=(
                                        '<b>%{label}</b><br>'
                                        'عدد العملاء الفريدين: %{value:,}<br>'
                                        'النسبة: %{percent}<extra></extra>'
                                    )
                                )])
                                fig_customers.update_layout(
                                    height=390,
                                    autosize=True,
                                    paper_bgcolor='#FFFFFF',
                                    plot_bgcolor='#FFFFFF',
                                    font=dict(
                                        family='Tahoma',
                                        color='#1F4373',
                                        size=10
                                    ),
                                    margin=dict(l=5, r=5, t=12, b=12),
                                    showlegend=True,
                                    legend=dict(
                                        orientation='h',
                                        yanchor='bottom',
                                        y=0,
                                        xanchor='center',
                                        x=0.5,
                                        font=dict(size=9, color='#1F4373')
                                    )
                                )
                                st.plotly_chart(
                                    fig_customers,
                                    use_container_width=True,
                                    config={'displayModeBar': False}
                                )
                                st.caption(
                                    f"إجمالي العملاء الفريدين: "
                                    f"{customer_state_df[debitor_col].nunique():,}"
                                )
                            else:
                                st.warning(
                                    "لا توجد قيم صالحة في Debitor وMain State لعرض توزيع العملاء."
                                )
                        else:
                            missing_cols = []
                            if not debitor_col:
                                missing_cols.append('Debitor')
                            if not main_state_col:
                                missing_cols.append('Main State')
                            st.warning(
                                "لم يتم العثور على العمود/الأعمدة التالية في ملف النشاط: "
                                + " و ".join(missing_cols)
                            )
            else:
                st.info('لا توجد بيانات كافية لعرض جودة المكالمات للمحصلين.')

            # ------------------------------------------
            # الشارت الثانية: منحنى النشاط خلال ساعات اليوم
            # ------------------------------------------
            if 'Created_DateTime' in activity_chart_df.columns:
                hourly_df = activity_chart_df.dropna(subset=['Created_DateTime']).copy()
                if not hourly_df.empty:
                    hourly_df['_hour_bucket'] = hourly_df['Created_DateTime'].dt.floor('h')
                    hourly_counts = (
                        hourly_df.groupby('_hour_bucket')
                        .size()
                        .sort_index()
                        .rename('عدد المكالمات')
                    )
                    hourly_chart = hourly_counts.reset_index()
                    hourly_chart['hour_label'] = hourly_chart['_hour_bucket'].apply(_fmt_hour_12_ar)

                    fig_area = go.Figure()
                    fig_area.add_trace(go.Scatter(
                        x=hourly_chart['hour_label'],
                        y=hourly_chart['عدد المكالمات'],
                        mode='lines+markers',
                        fill='tozeroy',
                        line=dict(color='#1F4373', width=3),
                        marker=dict(color='#99AABF', size=8),
                        hovertemplate='الساعة: %{x}<br>عدد المكالمات: %{y}<extra></extra>'
                    ))
                    fig_area.update_layout(
                        title=dict(
                            text='<b>منحنى النشاط خلال ساعات اليوم</b>',
                            x=0.5, xanchor='center',
                            font=dict(size=21, color='#1F4373', family='Tahoma')
                        ),
                        height=420,
                        paper_bgcolor='#FFFFFF',
                        plot_bgcolor='#FFFFFF',
                        font=dict(family='Tahoma', color='#1F4373', size=13),
                        margin=dict(l=20, r=25, t=70, b=55),
                        xaxis=dict(title='', showgrid=False, tickangle=-20),
                        yaxis=dict(title='عدد المكالمات', gridcolor='#E7EDF4', zeroline=False)
                    )
                    with st.container(border=True):
                        st.plotly_chart(fig_area, use_container_width=True, config={'displayModeBar': False})
                else:
                    st.info('لا توجد بيانات وقت صالحة لعرض منحنى النشاط خلال ساعات اليوم.')
            else:
                st.info('عمود Created_DateTime غير متوفر لعرض منحنى النشاط خلال ساعات اليوم.')

            # ------------------------------------------
            # الشارت الثالثة: تحليل الوقت الضائع
            # ------------------------------------------
            if 'Created_DateTime' in activity_chart_df.columns:
                gap_df = activity_chart_df.dropna(subset=['Created_DateTime']).copy()
                gap_df = gap_df.sort_values([collector_col, 'Created_DateTime']).copy()

                lost_time_rows = []
                for collector_name, grp in gap_df.groupby('_collector_name'):
                    grp = grp.sort_values('Created_DateTime').reset_index(drop=True)
                    total_lost = 0.0
                    longest_gap = 0.0

                    for idx in range(len(grp) - 1):
                        current_dt = grp.loc[idx, 'Created_DateTime']
                        next_dt = grp.loc[idx + 1, 'Created_DateTime']
                        current_status = str(grp.loc[idx, '_call_status']).strip()

                        raw_gap = (next_dt - current_dt).total_seconds() / 60.0
                        if raw_gap <= 0:
                            continue

                        adjusted_gap = max(0.0, raw_gap - _break_overlap_minutes(current_dt, next_dt))
                        if adjusted_gap <= 0:
                            continue

                        threshold = 20.0 if current_status == 'ناجحة' else 5.0 if current_status == 'مغطاة' else 5.0
                        lost_minutes = max(0.0, adjusted_gap - threshold)

                        total_lost += lost_minutes
                        longest_gap = max(longest_gap, adjusted_gap)

                    lost_time_rows.append({
                        'المحصل': collector_name,
                        'إجمالي الدقائق الضائعة': round(total_lost, 1),
                        'أطول فجوة توقف': round(longest_gap, 1),
                    })

                lost_summary = pd.DataFrame(lost_time_rows)
                lost_summary = lost_summary.sort_values(
                    by=['إجمالي الدقائق الضائعة', 'أطول فجوة توقف'],
                    ascending=False
                ).head(10)

                if not lost_summary.empty:
                    fig_lost = go.Figure()
                    fig_lost.add_trace(go.Bar(
                        y=lost_summary['المحصل'],
                        x=lost_summary['إجمالي الدقائق الضائعة'],
                        name='إجمالي الدقائق الضائعة',
                        orientation='h',
                        marker=dict(color='#1F4373', line=dict(width=0)),
                        text=[f"{v:.0f}" for v in lost_summary['إجمالي الدقائق الضائعة']],
                        textposition='outside',
                        hovertemplate='<b>%{y}</b><br>إجمالي الدقائق الضائعة: %{x:.1f}<extra></extra>'
                    ))
                    fig_lost.add_trace(go.Bar(
                        y=lost_summary['المحصل'],
                        x=lost_summary['أطول فجوة توقف'],
                        name='أطول فجوة توقف',
                        orientation='h',
                        marker=dict(color='#99AABF', line=dict(width=0)),
                        text=[f"{v:.0f}" for v in lost_summary['أطول فجوة توقف']],
                        textposition='outside',
                        hovertemplate='<b>%{y}</b><br>أطول فجوة توقف: %{x:.1f} دقيقة<extra></extra>'
                    ))
                    fig_lost.update_layout(
                        title=dict(
                            text='<b>تحليل الوقت الضائع للمحصلين بالدقائق</b>',
                            x=0.5, xanchor='center',
                            font=dict(size=21, color='#1F4373', family='Tahoma')
                        ),
                        barmode='group',
                        height=max(420, len(lost_summary) * 36 + 140),
                        paper_bgcolor='#FFFFFF',
                        plot_bgcolor='#FFFFFF',
                        font=dict(family='Tahoma', color='#1F4373', size=13),
                        margin=dict(l=20, r=45, t=70, b=35),
                        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
                        xaxis=dict(title='الدقائق', gridcolor='#E7EDF4', zeroline=False),
                        yaxis=dict(title='', autorange='reversed')
                    )
                    with st.container(border=True):
                        st.plotly_chart(fig_lost, use_container_width=True, config={'displayModeBar': False})
                else:
                    st.info('لا توجد بيانات كافية لحساب الوقت الضائع للمحصلين.')
            else:
                st.info('عمود Created_DateTime غير متوفر لحساب الوقت الضائع للمحصلين.')

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
            file_name=f"{selected_bank}_Activity_Success_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
