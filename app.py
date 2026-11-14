import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date, datetime, time, timedelta
import io
import base64
import os
import re
import zipfile
import hashlib
import gzip
import hmac
import json
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from cryptography.fernet import Fernet, InvalidToken
from difflib import SequenceMatcher, get_close_matches
from xml.sax.saxutils import escape as _xml_escape
from openpyxl import Workbook, load_workbook
from copy import copy
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

# ترحيل جلسات التشغيل القديمة: الرئيسية أصبحت صفحة مستقلة عن Overview.
if not st.session_state.get('_separate_home_overview_v1', False):
    st.session_state.current_page = 'Home'
    st.session_state['_separate_home_overview_v1'] = True
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

_overview_upload_panel_bg = (
    "#1F4373" if st.session_state.theme_choice == "ثيم إجادة (Navy)"
    else "#222B39" if st.session_state.theme_choice == "الوضع الداكن (Black)"
    else "#FFFFFF"
)
_overview_upload_caption_color = (
    "#E7EEF8" if st.session_state.theme_choice != "التصميم الفاتح (Light)"
    else "#64748B"
)

st.markdown(f"""
<style>
    [data-testid="collapsedControl"] {{display: none !important;}}
    [data-testid="stSidebar"] {{display: none !important;}}
    
    .stApp {{ background-color: {bg_color} !important; }}
    /* Hide any stale/raw markdown artifact from older bank-label markup after redeploy. */
    div[data-testid="stMarkdownContainer"] > p:has(code) {{
        display: none !important;
    }}
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
        border: 2px solid {header_logo_color} !important;
        border-radius: 50% !important;
        width: 44px !important;
        height: 44px !important;
        min-width: 44px !important;
        min-height: 44px !important;
        padding: 0 !important;
        color: transparent !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.28) !important;
        cursor: pointer !important;
    }}

    /* Compact Overview: border hugs the label instead of filling the column. */
    div.st-key-top_overview_btn div[data-testid="stButton"] {{
        width: fit-content !important;
    }}
    div.st-key-top_overview_btn div[data-testid="stButton"] > button {{
        width: auto !important;
        min-width: 0 !important;
        min-height: 42px !important;
        padding: 6px 14px !important;
        border-radius: 8px !important;
    }}
    div.st-key-top_overview_btn div[data-testid="stButton"] > button p {{
        font-size: 19px !important;
        font-weight: 900 !important;
        line-height: 1.1 !important;
        white-space: nowrap !important;
    }}

    /* One real flex row: each button is a separate item with ~0.5 cm (19px) gap. */
    div.st-key-report_nav_group {{
        width: 100% !important;
        max-width: 100% !important;
        direction: rtl !important;
        overflow-x: auto !important;
    }}
    div.st-key-report_nav_group[data-testid="stVerticalBlock"],
    div.st-key-report_nav_group [data-testid="stVerticalBlock"] {{
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        align-items: center !important;
        justify-content: flex-start !important;
        direction: rtl !important;
        gap: 19px !important;
        width: 100% !important;
        min-width: max-content !important;
    }}
    div.st-key-report_nav_group [data-testid="stElementContainer"] {{
        flex: 0 0 auto !important;
        width: fit-content !important;
        min-width: 0 !important;
        margin: 0 !important;
    }}

    /* Report navigation boxes should hug their text exactly. */
    div.st-key-nav_prom_btn div[data-testid="stButton"],
    div.st-key-nav_neg_btn div[data-testid="stButton"],
    div.st-key-nav_act_btn div[data-testid="stButton"] {{
        width: fit-content !important;
        margin: 0 !important;
    }}
    div.st-key-nav_prom_btn div[data-testid="stButton"] > button,
    div.st-key-nav_neg_btn div[data-testid="stButton"] > button,
    div.st-key-nav_act_btn div[data-testid="stButton"] > button {{
        width: auto !important;
        min-width: 0 !important;
        min-height: 36px !important;
        padding: 4px 10px !important;
        border-radius: 7px !important;
        margin: 0 !important;
    }}
    div.st-key-nav_prom_btn div[data-testid="stButton"] > button p,
    div.st-key-nav_neg_btn div[data-testid="stButton"] > button p,
    div.st-key-nav_act_btn div[data-testid="stButton"] > button p {{
        font-size: 17px !important;
        font-weight: 900 !important;
        line-height: 1.1 !important;
        white-space: nowrap !important;
    }}

    /* Fix the Streamlit header: never give the full-width toolbar a white background.
       Only the small actions cluster receives the white card. */
    header[data-testid="stHeader"] {{
        background: transparent !important;
        background-color: transparent !important;
        box-shadow: none !important;
        border: none !important;
        height: 2.5rem !important;
        min-height: 0 !important;
        pointer-events: none !important;
    }}
    header[data-testid="stHeader"] [data-testid="stToolbar"] {{
        position: relative !important;
        top: 0 !important;
        left: auto !important;
        right: auto !important;
        display: flex !important;
        width: 100% !important;
        height: 2.5rem !important;
        padding: 0 !important;
        margin: 0 !important;
        background: transparent !important;
        background-color: transparent !important;
        box-shadow: none !important;
        border: none !important;
        pointer-events: none !important;
    }}

    /* Keep Share, star, edit and GitHub together in a compact white box. */
    header[data-testid="stHeader"] [data-testid="stToolbarActions"] {{
        position: absolute !important;
        top: 7px !important;
        left: 8px !important;
        right: auto !important;
        display: inline-flex !important;
        flex-direction: row !important;
        align-items: center !important;
        justify-content: flex-start !important;
        width: max-content !important;
        max-width: calc(100vw - 36px) !important;
        height: auto !important;
        padding: 5px 10px !important;
        margin: 0 !important;
        background: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 10px !important;
        box-shadow: 0 2px 7px rgba(0,0,0,0.12) !important;
        direction: ltr !important;
        pointer-events: auto !important;
        z-index: 20 !important;
    }}
    header[data-testid="stHeader"] [data-testid="stToolbarActions"] :is(button, a, span, [role="button"]),
    header[data-testid="stHeader"] [data-testid="stToolbarActions"] :is(button, a, span, [role="button"]) * {{
        color: {snb_green} !important;
        -webkit-text-fill-color: {snb_green} !important;
    }}
    /* Filter the complete SVG so both solid and outline icons remain green;
       do not force fill on every path (which ruins outlined icons). */
    header[data-testid="stHeader"] [data-testid="stToolbarActions"] svg,
    header[data-testid="stHeader"] [data-testid="stToolbarActions"] img,
    header[data-testid="stHeader"] [data-testid="stToolbar"] svg,
    header[data-testid="stHeader"] [data-testid="stToolbar"] img {{
        color: {snb_green} !important;
        filter: brightness(0) saturate(100%) invert(49%) sepia(77%)
                saturate(575%) hue-rotate(82deg) brightness(92%) contrast(88%) !important;
        opacity: 1 !important;
    }}
    /* Allow the Streamlit menu at the opposite edge to remain clickable. */
    header[data-testid="stHeader"] [data-testid="stMainMenu"],
    header[data-testid="stHeader"] [data-testid="stMainMenu"] *,
    header[data-testid="stHeader"] button[aria-label*="menu" i] {{
        pointer-events: auto !important;
        color: {snb_green} !important;
    }}

    /* Bank selection: real two-column row aligned to screen left. */
    .stApp [data-testid="stHorizontalBlock"]:has(.st-key-bank_SNB),
    .stApp [data-testid="stHorizontalBlock"]:has(.st-key-overview_choose_SNB) {{
        direction: ltr !important;
        flex-wrap: nowrap !important;
        justify-content: flex-start !important;
        align-items: center !important;
        gap: 8px !important;
    }}
    .stApp [data-testid="stHorizontalBlock"]:has(.st-key-bank_SNB) > [data-testid="stColumn"],
    .stApp [data-testid="stHorizontalBlock"]:has(.st-key-overview_choose_SNB) > [data-testid="stColumn"] {{
        min-width: 0 !important;
    }}
    .stApp [data-testid="stHorizontalBlock"]:has(.st-key-bank_SNB) > [data-testid="stColumn"]:nth-child(-n+2),
    .stApp [data-testid="stHorizontalBlock"]:has(.st-key-overview_choose_SNB) > [data-testid="stColumn"]:nth-child(-n+2) {{
        flex: 0 0 118px !important;
        width: 118px !important;
        max-width: 118px !important;
    }}
    div.st-key-bank_SNB div[data-testid="stButton"],
    div.st-key-bank_STC div[data-testid="stButton"],
    div.st-key-overview_choose_SNB div[data-testid="stButton"],
    div.st-key-overview_choose_STC div[data-testid="stButton"] {{
        width: fit-content !important;
    }}
    div.st-key-bank_SNB button,
    div.st-key-bank_STC button,
    div.st-key-overview_choose_SNB button,
    div.st-key-overview_choose_STC button {{
        width: 116px !important;
        min-width: 0 !important;
        min-height: 44px !important;
        padding: 6px 12px !important;
        border-radius: 9px !important;
        white-space: nowrap !important;
    }}
    div.st-key-bank_SNB button p,
    div.st-key-bank_STC button p,
    div.st-key-overview_choose_SNB button p,
    div.st-key-overview_choose_STC button p {{
        font-size: 21px !important;
        font-weight: 900 !important;
        line-height: 1.15 !important;
        white-space: nowrap !important;
    }}

    /* Home navigation is intentionally icon-only and hugs the icon. */
    div.st-key-nav_home_btn div[data-testid="stButton"] {{
        width: fit-content !important;
        margin: 0 !important;
    }}
    div.st-key-nav_home_btn div[data-testid="stButton"] > button {{
        min-width: 38px !important;
        width: 38px !important;
        min-height: 36px !important;
        height: 36px !important;
        padding: 3px 6px !important;
        border-radius: 7px !important;
        margin: 0 !important;
    }}
    div.st-key-nav_home_btn div[data-testid="stButton"] > button p {{
        font-size: 19px !important;
        line-height: 1 !important;
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
    [class*="st-key-report_upload_bar_"] [data-testid="stHorizontalBlock"]:has(.ejada-upload-name) {{
        align-items: center !important;
        flex-wrap: nowrap !important;
    }}
    /* Dedicated report header: compact inline title, responsive upload box. */
    [class*="st-key-report_upload_bar_"] {{
        width: 100% !important;
    }}
    [class*="st-key-report_upload_bar_"] [data-testid="stFileUploader"] {{
        width: 100% !important;
        min-width: 0 !important;
        margin: 0 !important;
        overflow: visible !important;
    }}
    [class*="st-key-report_upload_bar_"] [data-testid="stFileUploaderDropzone"] {{
        box-sizing: border-box !important;
        height: auto !important;
        min-height: 64px !important;
        max-height: none !important;
        width: 100% !important;
        padding: 8px 10px !important;
        border: 1px dashed {border_color} !important;
        background-color: #FFFFFF !important;
        border-radius: 9px !important;
        overflow: hidden !important;
    }}
    [class*="st-key-report_upload_bar_"] [data-testid="stFileUploaderDropzone"] > div {{
        display: flex !important;
        min-height: 46px !important;
        height: auto !important;
        flex-direction: row !important;
        align-items: center !important;
        justify-content: center !important;
        flex-wrap: wrap !important;
        gap: 5px !important;
        overflow: hidden !important;
    }}
    [class*="st-key-report_upload_bar_"] [data-testid="stFileUploaderDropzoneInstructions"],
    [class*="st-key-report_upload_bar_"] [data-testid="stFileUploaderDropzone"] small {{
        display: none !important;
    }}
    [class*="st-key-report_upload_bar_"] [data-testid="stFileUploaderDropzone"] button {{
        height: 38px !important;
        min-height: 38px !important;
        font-size: 15px !important;
        padding: 5px 15px !important;
        margin: 0 auto !important;
    }}
    /* The uploaded name is shown above in a dedicated ellipsized line;
       hide Streamlit's duplicate list to prevent overlapping file chips. */
    [class*="st-key-report_upload_bar_"] [data-testid="stFileUploaderFile"],
    [class*="st-key-report_upload_bar_"] [data-testid="stFileUploaderFileList"] {{
        display: none !important;
    }}
    [class*="st-key-report_upload_bar_"] .ejada-upload-name {{
        width: 100%;
        max-width: 100%;
        white-space: nowrap;
        text-align: right;
        direction: rtl;
        overflow: hidden;
        text-overflow: ellipsis;
        font-size: 15px;
        font-weight: 900;
        min-height: 25px;
        line-height: 25px;
    }}
    div[class*="st-key-small_clear_report_"] button {{
        width: 34px !important;
        min-width: 34px !important;
        height: 36px !important;
        min-height: 36px !important;
        padding: 0 !important;
        border-radius: 8px !important;
        color: #FFFFFF !important;
        background: #C5393B !important;
        border: 0 !important;
        font-size: 19px !important;
        font-weight: 900 !important;
    }}
    div[class*="st-key-small_clear_report_"] button p {{
        color: #FFFFFF !important;
        font-size: 19px !important;
        font-weight: 900 !important;
    }}

    /* Compact promise diagnostic with smaller bold labels and centered heading. */
    div[class*="st-key-promise_diagnostic_"] {{
        width: min(1400px, 100%) !important;
        max-width: 1400px !important;
        min-width: 0 !important;
        margin: 8px auto 12px !important;
        box-sizing: border-box !important;
    }}
    div[class*="st-key-promise_diagnostic_"] [data-testid="stExpander"] {{
        width: 100% !important;
        max-width: 100% !important;
        margin: 0 auto !important;
        box-sizing: border-box !important;
    }}
    div[class*="st-key-promise_diagnostic_"] [data-testid="stExpander"] summary {{
        position: relative !important;
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
        direction: rtl !important;
        text-align: center !important;
        min-height: 44px !important;
        padding-right: 34px !important;
        padding-left: 34px !important;
    }}
    /* Streamlit places the collapse arrow and title inside separate elements.
       Center the title independently of the arrow on both open and closed states. */
    div[class*="st-key-promise_diagnostic_"] [data-testid="stExpander"] summary [data-testid="stExpanderToggleContent"] {{
        flex: 1 1 auto !important;
        width: 100% !important;
        text-align: center !important;
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
    }}
    div[class*="st-key-promise_diagnostic_"] [data-testid="stExpander"] summary [data-testid="stExpanderToggleIcon"] {{
        position: absolute !important;
        right: 10px !important;
        top: 50% !important;
        transform: translateY(-50%) !important;
    }}
    div[class*="st-key-promise_diagnostic_"] [data-testid="stExpander"] summary [data-testid="stMarkdownContainer"] {{
        position: static !important;
    }}
    div[class*="st-key-promise_diagnostic_"] [data-testid="stExpander"] summary p {{
        position: absolute !important;
        left: 50% !important;
        top: 50% !important;
        transform: translate(-50%, -50%) !important;
        width: max-content !important;
        max-width: calc(100% - 74px) !important;
        margin: 0 !important;
        padding: 0 !important;
        font-size: 16px !important;
        font-weight: 900 !important;
        text-align: center !important;
        direction: rtl !important;
        line-height: 1.45 !important;
        white-space: normal !important;
        color: #1F4373 !important;
    }}
    div[class*="st-key-promise_diagnostic_"] [data-testid="stExpander"] [data-testid="stExpanderDetails"] {{
        padding: 8px 10px !important;
    }}
    table.ejada-diagnostic-table {{
        width: max-content !important;
        max-width: 100% !important;
        margin-right: 0 !important;
        margin-left: auto !important;
        border: 1px solid #B9C8DB !important;
        border-collapse: collapse !important;
        table-layout: auto !important;
        direction: rtl;
        font-size: 13px;
        font-weight: 800;
        color: #1F4373;
        background: #FFFFFF;
    }}
    table.ejada-diagnostic-table th,
    table.ejada-diagnostic-table td {{
        padding: 7px 10px;
        border: 1px solid #B9C8DB !important;
        font-size: 13px;
        font-weight: 800;
        white-space: normal;
        overflow-wrap: anywhere;
    }}
    div[class*="st-key-promise_details_row_"] [class*="st-key-promise_diagnostic_"] {{
        width: 100% !important;
        max-width: 100% !important;
        margin: 0 !important;
    }}
    div[class*="st-key-promise_details_row_"] [data-testid="stHorizontalBlock"] {{
        align-items: flex-start !important;
    }}
    table.ejada-diagnostic-table thead th {{
        background: #2CA549 !important;
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-weight: 900 !important;
    }}
    table.ejada-diagnostic-table th:first-child,
    table.ejada-diagnostic-table td:first-child {{
        width: auto !important;
        min-width: 0 !important;
        text-align: right !important;
    }}
    table.ejada-diagnostic-table th:last-child,
    table.ejada-diagnostic-table td:last-child {{
        width: 100px !important;
        min-width: 80px !important;
        max-width: 100px !important;
        text-align: left !important;
        font-variant-numeric: tabular-nums;
        white-space: nowrap !important;
    }}

    /* Compact same-height file panel and title across SNB/STC reports.
       The X sits above the real native Streamlit uploader (+), outside file panel. */
    [class*="st-key-report_upload_bar_"] [data-testid="stHorizontalBlock"]:has(.ejada-report-title-holder) {{
        direction: rtl !important;
        flex-wrap: nowrap !important;
        align-items: flex-start !important;
    }}
    [class*="st-key-report_upload_bar_"] [data-testid="stHorizontalBlock"]:has(.ejada-report-title-holder) > [data-testid="stColumn"] {{
        min-width: 0 !important;
        align-self: flex-start !important;
    }}
    [class*="st-key-report_upload_bar_"] .ejada-report-title-holder {{
        display: flex !important;
        align-items: flex-start !important;
        justify-content: center !important;
        width: 100% !important;
        height: 90px !important;
        min-height: 90px !important;
        max-height: 90px !important;
        box-sizing: border-box !important;
        padding: 0 !important;
        margin: 0 !important;
        direction: rtl !important;
    }}
    [class*="st-key-report_upload_bar_"] .ejada-report-title {{
        display: inline-flex !important;
        align-items: center !important;
        justify-content: center !important;
        text-align: center !important;
        height: 90px !important;
        min-height: 90px !important;
        max-height: 90px !important;
        box-sizing: border-box !important;
        width: fit-content !important;
        max-width: 100% !important;
        padding: 8px 13px !important;
        border: 1px solid !important;
        border-radius: 9px !important;
        font-weight: 900 !important;
        line-height: 1.25 !important;
        white-space: normal !important;
        margin: 0 !important;
    }}
    [class*="st-key-report_upload_bar_"] .ejada-current-file-box {{
        box-sizing: border-box !important;
        width: 100% !important;
        height: 90px !important;
        min-height: 90px !important;
        max-height: 90px !important;
        background: #FFFFFF !important;
        border: 1px dashed {border_color} !important;
        border-radius: 9px !important;
        display: flex !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: center !important;
        padding: 8px 9px !important;
        margin: 0 !important;
        overflow: hidden !important;
        color: #1F4373 !important;
        direction: ltr !important;
        text-align: center !important;
    }}
    [class*="st-key-report_upload_bar_"] .ejada-file-name {{
        font-size: 14px !important;
        font-weight: 800 !important;
        line-height: 1.4 !important;
        width: 100% !important;
        max-width: 100% !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        color: #1F4373 !important;
    }}
    [class*="st-key-report_upload_bar_"] .ejada-file-size {{
        font-size: 12px !important;
        font-weight: 700 !important;
        white-space: nowrap !important;
        color: #586F8C !important;
        line-height: 1.6 !important;
    }}
    [class*="st-key-report_upload_bar_"] .ejada-no-file {{
        color: #1F4373 !important;
        font-size: 14px !important;
        font-weight: 900 !important;
    }}
    [class*="st-key-report_upload_bar_"] .ejada-upload-name {{
        max-width: 100% !important;
        min-height: 25px !important;
        line-height: 25px !important;
        margin: 0 !important;
        overflow: hidden !important;
        white-space: nowrap !important;
        text-overflow: ellipsis !important;
        text-align: right !important;
        direction: rtl !important;
    }}
    [class*="st-key-mini_actions_"] [data-testid="stVerticalBlock"] {{
        gap: 4px !important;
    }}
    [class*="st-key-mini_actions_"] .ejada-mini-x-placeholder {{
        width: 40px !important;
        height: 40px !important;
        min-height: 40px !important;
        margin: 0 !important;
    }}
    [class*="st-key-mini_actions_"] [data-testid="stFileUploader"] {{
        width: 40px !important;
        min-width: 40px !important;
        max-width: 40px !important;
        margin: 0 !important;
        padding: 0 !important;
    }}
    [class*="st-key-mini_actions_"] [data-testid="stFileUploaderDropzone"] {{
        box-sizing: border-box !important;
        width: 40px !important;
        min-width: 40px !important;
        max-width: 40px !important;
        height: 40px !important;
        min-height: 40px !important;
        max-height: 40px !important;
        padding: 1px !important;
        margin: 0 !important;
        border: 1px solid #1F4373 !important;
        border-radius: 8px !important;
        background: #FFFFFF !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        overflow: hidden !important;
    }}
    [class*="st-key-mini_actions_"] [data-testid="stFileUploaderDropzone"] > div {{
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        width: 100% !important;
        min-height: 0 !important;
        height: 100% !important;
        overflow: hidden !important;
    }}
    [class*="st-key-mini_actions_"] [data-testid="stFileUploaderDropzoneInstructions"],
    [class*="st-key-mini_actions_"] [data-testid="stFileUploaderDropzone"] small,
    [class*="st-key-mini_actions_"] [data-testid="stFileUploaderFile"],
    [class*="st-key-mini_actions_"] [data-testid="stFileUploaderFileList"] {{
        display: none !important;
    }}
    [class*="st-key-mini_actions_"] [data-testid="stFileUploaderDropzone"] button {{
        position: relative !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        width: 36px !important;
        min-width: 36px !important;
        height: 36px !important;
        min-height: 36px !important;
        padding: 0 !important;
        margin: 0 !important;
        font-size: 0 !important;
        color: transparent !important;
        background: #FFFFFF !important;
        border: 0 !important;
        box-shadow: none !important;
        cursor: pointer !important;
    }}
    [class*="st-key-mini_actions_"] [data-testid="stFileUploaderDropzone"] button * {{
        font-size: 0 !important;
        color: transparent !important;
        visibility: hidden !important;
    }}
    [class*="st-key-mini_actions_"] [data-testid="stFileUploaderDropzone"] button::after {{
        content: "+" !important;
        font-size: 25px !important;
        font-weight: 500 !important;
        line-height: 1 !important;
        color: #1F4373 !important;
        visibility: visible !important;
    }}
    div[class*="st-key-small_clear_report_"] button {{
        width: 40px !important;
        min-width: 40px !important;
        height: 40px !important;
        min-height: 40px !important;
        padding: 0 !important;
        border-radius: 8px !important;
        background: #C5393B !important;
        border: 0 !important;
        color: #FFFFFF !important;
        box-shadow: none !important;
        font-size: 20px !important;
        font-weight: 900 !important;
    }}
    div[class*="st-key-small_clear_report_"] button p {{
        color: #FFFFFF !important;
        font-size: 20px !important;
        font-weight: 900 !important;
    }}


    /* Compact promise KPIs replace the old report title, without extra row. */
    [class*="st-key-report_upload_bar_"] .ejada-promises-header-kpis {{
        display: flex !important;
        align-items: stretch !important;
        justify-content: center !important;
        gap: 9px !important;
        direction: rtl !important;
        height: 90px !important;
        width: 100% !important;
        margin: 0 !important;
    }}
    [class*="st-key-report_upload_bar_"] .ejada-mini-kpi {{
        display: flex !important;
        flex: 1 1 0 !important;
        min-width: 0 !important;
        max-width: 210px !important;
        box-sizing: border-box !important;
        flex-direction: column !important;
        align-items: center !important;
        justify-content: center !important;
        gap: 6px !important;
        text-align: center !important;
        padding: 7px 6px !important;
        background: #FFFFFF !important;
        border: 1px solid #99AABF !important;
        border-radius: 9px !important;
        color: #1F4373 !important;
        font-weight: 900 !important;
        overflow: hidden !important;
    }}
    [class*="st-key-report_upload_bar_"] .ejada-mini-kpi-active {{
        border-color: #2CA549 !important;
    }}
    [class*="st-key-report_upload_bar_"] .ejada-mini-kpi span {{
        font-size: clamp(11px, 0.88vw, 16px) !important;
        font-weight: 900 !important;
        white-space: normal !important;
        line-height: 1.25 !important;
    }}
    [class*="st-key-report_upload_bar_"] .ejada-mini-kpi strong {{
        font-size: clamp(20px, 1.6vw, 29px) !important;
        line-height: 1.1 !important;
        font-weight: 900 !important;
        color: #1F4373 !important;
    }}
    [class*="st-key-report_upload_bar_"] .ejada-mini-kpi-active strong {{
        color: #2CA549 !important;
    }}
    /* Center the details expander title independently of its right-side caret. */
    div[class*="st-key-active_sheet_detail_"] [data-testid="stExpander"] summary {{
        position: relative !important;
        min-height: 48px !important;
    }}
    div[class*="st-key-active_sheet_detail_"] [data-testid="stExpander"] summary p {{
        position: absolute !important;
        top: 50% !important;
        left: 50% !important;
        transform: translate(-50%, -50%) !important;
        text-align: center !important;
        margin: 0 !important;
        width: max-content !important;
        max-width: calc(100% - 70px) !important;
        font-size: 15px !important;
        font-weight: 900 !important;
    }}
    div[class*="st-key-promise_downloads_"] [data-testid="stDownloadButton"] {{
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
        width: 100% !important;
    }}
    div[class*="st-key-promise_downloads_"] [data-testid="stDownloadButton"] button {{
        width: fit-content !important;
        max-width: 100% !important;
        margin: 0 auto !important;
        min-height: 44px !important;
        padding: 8px 18px !important;
        text-align: center !important;
        font-weight: 900 !important;
    }}
    div[class*="st-key-promise_downloads_"] [data-testid="stDownloadButton"] button p {{
        font-size: 16px !important;
        text-align: center !important;
        font-weight: 900 !important;
        white-space: normal !important;
    }}
    /* One HTML grid for both diagnostic tables prevents staggered header rows. */
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual {{
        display: grid !important;
        grid-template-columns: minmax(0, 1.4fr) minmax(0, 1fr) !important;
        align-items: start !important;
        direction: rtl !important;
        column-gap: 12px !important;
        width: 100% !important;
        margin-top: 0 !important;
    }}
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual > div {{
        min-width: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
    }}
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual table.ejada-diagnostic-table {{
        width: 100% !important;
        margin: 0 !important;
        border-collapse: collapse !important;
    }}
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual thead {{
        vertical-align: top !important;
    }}
    /* Two bordered diagnostic tables side by side: steps on right, class counts on left. */
    div[class*="st-key-promise_diagnostic_"] [data-testid="stHorizontalBlock"]:has(.ejada-diagnostic-table) {{
        display: flex !important;
        flex-wrap: nowrap !important;
        align-items: flex-start !important;
        direction: rtl !important;
        gap: 12px !important;
    }}
    div[class*="st-key-promise_diagnostic_"] [data-testid="stHorizontalBlock"]:has(.ejada-diagnostic-table) [data-testid="stColumn"] {{
        min-width: 0 !important;
        align-self: flex-start !important;
    }}
    div[class*="st-key-promise_diagnostic_"] [data-testid="stHorizontalBlock"] table.ejada-diagnostic-table {{
        width: 100% !important;
        min-width: 0 !important;
        max-width: 100% !important;
        table-layout: auto !important;
    }}
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-small-caption {{
        text-align: center !important;
        color: #1F4373 !important;
        font-size: 12px !important;
        font-weight: 900 !important;
        margin-bottom: 6px !important;
    }}
    @media (max-width: 1100px) {{
        [class*="st-key-report_upload_bar_"] .ejada-promises-header-kpis {{
            gap: 5px !important;
        }}
        [class*="st-key-report_upload_bar_"] .ejada-mini-kpi {{
            padding: 5px 3px !important;
        }}
        [class*="st-key-report_upload_bar_"] .ejada-mini-kpi span {{
            font-size: 11px !important;
        }}
    }}

    /* Layout repair: diagnostic columns must never overflow their white expander. */
    div[class*="st-key-promise_diagnostic_"] [data-testid="stExpander"] details {{
        max-width: 100% !important;
        box-sizing: border-box !important;
        overflow: hidden !important;
    }}
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual {{
        display: grid !important;
        grid-template-columns: minmax(0, 1.55fr) minmax(0, 1fr) !important;
        align-items: start !important;
        gap: 12px !important;
        width: 100% !important;
        max-width: 100% !important;
        min-width: 0 !important;
        box-sizing: border-box !important;
        direction: rtl !important;
    }}
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual > div {{
        min-width: 0 !important;
        max-width: 100% !important;
        box-sizing: border-box !important;
    }}
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-caption-spacer,
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual .ejada-diagnostic-small-caption {{
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        text-align: center !important;
        height: 32px !important;
        min-height: 32px !important;
        max-height: 32px !important;
        box-sizing: border-box !important;
        padding: 0 3px !important;
        margin: 0 !important;
        font-size: 12px !important;
        line-height: 1.25 !important;
        white-space: normal !important;
        overflow-wrap: anywhere !important;
    }}
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual table.ejada-diagnostic-table {{
        width: 100% !important;
        max-width: 100% !important;
        min-width: 0 !important;
        table-layout: fixed !important;
        border-collapse: collapse !important;
        box-sizing: border-box !important;
        margin: 0 !important;
    }}
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual table.ejada-diagnostic-table th,
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual table.ejada-diagnostic-table td {{
        white-space: normal !important;
        overflow-wrap: anywhere !important;
        word-break: normal !important;
        box-sizing: border-box !important;
        min-width: 0 !important;
    }}
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual table.ejada-diagnostic-table th:last-child,
    div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual table.ejada-diagnostic-table td:last-child {{
        width: 86px !important;
        max-width: 86px !important;
    }}
    /* STC active and broken export buttons: far right, one under the other. */
    div.st-key-promise_downloads_STC [data-testid="stDownloadButton"] {{
        display: flex !important;
        direction: rtl !important;
        justify-content: flex-start !important;
        width: 100% !important;
        margin: 0 0 8px 0 !important;
    }}
    div.st-key-promise_downloads_STC [data-testid="stDownloadButton"] button {{
        width: fit-content !important;
        max-width: 100% !important;
        margin: 0 !important;
    }}
    @media (max-width: 780px) {{
        div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual {{
            gap: 7px !important;
        }}
        div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual table.ejada-diagnostic-table {{
            font-size: 11px !important;
        }}
        div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual table.ejada-diagnostic-table th,
        div[class*="st-key-promise_diagnostic_"] .ejada-diagnostic-dual table.ejada-diagnostic-table td {{
            padding: 5px 4px !important;
        }}
    }}

    /* Activity-page upload cards only: compact outer border and dropzone. */
    div[class*="st-key-activity_model_upload_card"],
    div[class*="st-key-activity_file_upload_card"] {{
        box-sizing: border-box !important;
        width: 100% !important;
        max-width: 340px !important;
        min-width: 0 !important;
        margin-left: auto !important;
        margin-right: auto !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-activity_model_upload_card),
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-activity_file_upload_card) {{
        box-sizing: border-box !important;
        width: 100% !important;
        max-width: 340px !important;
        min-width: 0 !important;
        height: auto !important;
        min-height: 0 !important;
        padding: 10px 10px 7px !important;
        margin-left: auto !important;
        margin-right: auto !important;
    }}
    div[class*="st-key-activity_model_upload_card"] [data-testid="stFileUploader"],
    div[class*="st-key-activity_file_upload_card"] [data-testid="stFileUploader"] {{
        box-sizing: border-box !important;
        width: 100% !important;
        max-width: 100% !important;
        min-width: 0 !important;
    }}
    div[class*="st-key-activity_model_upload_card"] [data-testid="stFileUploaderDropzone"],
    div[class*="st-key-activity_file_upload_card"] [data-testid="stFileUploaderDropzone"] {{
        box-sizing: border-box !important;
        width: 100% !important;
        max-width: 100% !important;
        min-width: 0 !important;
        padding: 7px 9px !important;
        min-height: 50px !important;
        border-radius: 8px !important;
    }}


    /* Activity uploaders: RESTORE the native Streamlit 2026 file chips and add button.
       The uploaded filename, native navy X, and native plus are INSIDE the dropzone.
       Never force the entire dropzone to a 36px square! */
    div[class*="st-key-activity_model_upload_card"],
    div[class*="st-key-activity_file_upload_card"] {{
        box-sizing: border-box !important;
        width: 100% !important;
        max-width: 315px !important;
        min-width: 0 !important;
        margin-left: auto !important;
        margin-right: auto !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-activity_model_upload_card),
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-activity_file_upload_card) {{
        box-sizing: border-box !important;
        width: 100% !important;
        max-width: 315px !important;
        min-width: 0 !important;
        min-height: 0 !important;
        height: auto !important;
        padding: 7px 9px !important;
        margin: 0 auto !important;
    }}
    div[class*="st-key-activity_model_upload_card"] [data-testid="stFileUploader"],
    div[class*="st-key-activity_file_upload_card"] [data-testid="stFileUploader"] {{
        box-sizing: border-box !important;
        display: block !important;
        direction: ltr !important;
        width: 100% !important;
        min-width: 0 !important;
        max-width: 100% !important;
    }}
    div[class*="st-key-activity_model_upload_card"] [data-testid="stFileUploaderDropzone"],
    div[class*="st-key-activity_file_upload_card"] [data-testid="stFileUploaderDropzone"] {{
        box-sizing: border-box !important;
        display: flex !important;
        flex-flow: row nowrap !important;
        justify-content: center !important;
        align-items: center !important;
        width: 100% !important;
        min-width: 0 !important;
        max-width: 100% !important;
        min-height: 54px !important;
        height: auto !important;
        max-height: none !important;
        padding: 7px 8px !important;
        overflow: visible !important;
    }}
    /* Do NOT hide the original upload button or its icon. */
    div[class*="st-key-activity_model_upload_card"] [data-testid="stFileUploaderDropzone"] button,
    div[class*="st-key-activity_file_upload_card"] [data-testid="stFileUploaderDropzone"] button {{
        font-size: inherit !important;
        color: #1F4373 !important;
        visibility: visible !important;
        text-indent: 0 !important;
    }}
    div[class*="st-key-activity_model_upload_card"] [data-testid="stFileUploaderDropzone"] button *,
    div[class*="st-key-activity_file_upload_card"] [data-testid="stFileUploaderDropzone"] button * {{
        font-size: inherit !important;
        visibility: visible !important;
    }}
    /* The redesigned uploader renders both native X and native plus inside stFileChips. */
    div[class*="st-key-activity_model_upload_card"] [data-testid="stFileChips"],
    div[class*="st-key-activity_file_upload_card"] [data-testid="stFileChips"] {{
        display: flex !important;
        flex-flow: row nowrap !important;
        align-items: center !important;
        justify-content: flex-start !important;
        gap: 5px !important;
        width: 100% !important;
        max-width: 100% !important;
        min-width: 0 !important;
        overflow: visible !important;
    }}
    div[class*="st-key-activity_model_upload_card"] [data-testid="stFileChips"] > ul,
    div[class*="st-key-activity_file_upload_card"] [data-testid="stFileChips"] > ul {{
        flex: 1 1 0 !important;
        min-width: 0 !important;
        width: auto !important;
        max-width: calc(100% - 35px) !important;
        margin: 0 !important;
        padding: 0 !important;
        display: flex !important;
        flex-flow: row nowrap !important;
    }}
    div[class*="st-key-activity_model_upload_card"] [data-testid="stFileChips"] > button,
    div[class*="st-key-activity_file_upload_card"] [data-testid="stFileChips"] > button {{
        flex: 0 0 28px !important;
        min-width: 28px !important;
        width: 28px !important;
        height: 28px !important;
        min-height: 28px !important;
        margin: 0 !important;
        padding: 2px !important;
    }}
    div[class*="st-key-activity_model_upload_card"] [data-testid="stFileChip"],
    div[class*="st-key-activity_file_upload_card"] [data-testid="stFileChip"] {{
        box-sizing: border-box !important;
        flex: 1 1 auto !important;
        min-width: 0 !important;
        max-width: 100% !important;
        overflow: hidden !important;
    }}
    div[class*="st-key-activity_model_upload_card"] [data-testid="stFileChip"] > div,
    div[class*="st-key-activity_file_upload_card"] [data-testid="stFileChip"] > div {{
        min-width: 0 !important;
    }}
    div[class*="st-key-activity_model_upload_card"] [data-testid="stFileChipName"],
    div[class*="st-key-activity_file_upload_card"] [data-testid="stFileChipName"] {{
        display: block !important;
        max-width: 120px !important;
        white-space: nowrap !important;
        text-overflow: ellipsis !important;
        overflow: hidden !important;
    }}
    /* Bank buttons and report navigation now share a single parent row. */
    .stApp [data-testid="stHorizontalBlock"]:has(.st-key-bank_SNB) > [data-testid="stColumn"]:last-child {{
        flex: 1 1 auto !important;
        min-width: 0 !important;
    }}
    .stApp [data-testid="stHorizontalBlock"]:has(.st-key-bank_SNB) div.st-key-report_nav_group {{
        margin-top: 4px !important;
        min-width: 0 !important;
        overflow-x: auto !important;
    }}

    /* Consistent 315x185 upload boxes at the far right on all report pages. */
    div[data-testid="stHorizontalBlock"]:has(.st-key-activity_model_upload_card),
    div[data-testid="stHorizontalBlock"]:has(.st-key-overview_activity_upload_card),
    div[class*="st-key-report_upload_bar_"] [data-testid="stHorizontalBlock"]:has([class*="st-key-ejada_shared_upload_card_"]) {{
        direction: rtl !important;
        justify-content: flex-start !important;
        align-items: flex-start !important;
        flex-wrap: nowrap !important;
        gap: 12px !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-activity_model_upload_card),
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-activity_file_upload_card),
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-overview_activity_upload_card),
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-overview_model_upload_card),
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-overview_promises_upload_card),
    div[data-testid="stVerticalBlockBorderWrapper"]:has([class*="st-key-ejada_shared_upload_card_"]) {{
        box-sizing: border-box !important;
        width: min(315px, 100%) !important;
        max-width: 315px !important;
        min-width: 0 !important;
        height: 185px !important;
        min-height: 185px !important;
        max-height: 185px !important;
        margin-left: auto !important;
        margin-right: 0 !important;
        padding: 9px 10px !important;
        overflow: visible !important;
        background: #FFFFFF !important;
        border: 1px solid {border_color} !important;
        border-radius: 10px !important;
    }}
    div[class*="st-key-activity_model_upload_card"],
    div[class*="st-key-activity_file_upload_card"],
    div.st-key-overview_activity_upload_card,
    div.st-key-overview_model_upload_card,
    div.st-key-overview_promises_upload_card,
    div[class*="st-key-ejada_shared_upload_card_"] {{
        box-sizing: border-box !important;
        max-width: 100% !important;
        min-width: 0 !important;
    }}
    div.st-key-overview_activity_upload_card [data-testid="stFileUploader"],
    div.st-key-overview_model_upload_card [data-testid="stFileUploader"],
    div.st-key-overview_promises_upload_card [data-testid="stFileUploader"] {{
        width: 100% !important;
        max-width: 100% !important;
        min-width: 0 !important;
    }}
    div.st-key-overview_activity_upload_card [data-testid="stFileUploaderDropzone"],
    div.st-key-overview_model_upload_card [data-testid="stFileUploaderDropzone"],
    div.st-key-overview_promises_upload_card [data-testid="stFileUploaderDropzone"] {{
        width: 100% !important;
        min-width: 0 !important;
        min-height: 54px !important;
        height: auto !important;
        padding: 6px !important;
    }}
    div.st-key-overview_activity_upload_card [data-testid="stCaptionContainer"],
    div.st-key-overview_model_upload_card [data-testid="stCaptionContainer"],
    div.st-key-overview_promises_upload_card [data-testid="stCaptionContainer"] {{
        max-width: 100% !important;
        overflow: hidden !important;
        white-space: nowrap !important;
        text-overflow: ellipsis !important;
    }}
    div[class*="st-key-report_upload_bar_"] [data-testid="stHorizontalBlock"]:has([class*="st-key-ejada_shared_upload_card_"]) > [data-testid="stColumn"]:first-child {{
        flex: 0 0 315px !important;
        min-width: 315px !important;
        max-width: 315px !important;
    }}
    @media (max-width: 1050px) {{
        div[data-testid="stHorizontalBlock"]:has(.st-key-overview_activity_upload_card),
        div[data-testid="stHorizontalBlock"]:has(.st-key-activity_model_upload_card) {{
            flex-wrap: wrap !important;
        }}
    }}

    /* Overview only: compact uploader row in place of report/home navigation. */
    div.st-key-overview_upload_header {{
        width: 100% !important;
        max-width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
    }}
    div.st-key-overview_upload_header [data-testid="stHorizontalBlock"]:has(.st-key-overview_activity_upload_card) {{
        direction: rtl !important;
        justify-content: flex-start !important;
        align-items: flex-start !important;
        flex-wrap: nowrap !important;
        gap: 12px !important;
        width: 100% !important;
        margin: 0 !important;
    }}
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-overview_activity_upload_card),
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-overview_model_upload_card),
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-overview_promises_upload_card) {{
        box-sizing: border-box !important;
        width: min(300px, 100%) !important;
        max-width: 300px !important;
        min-width: 0 !important;
        height: 178px !important;
        min-height: 178px !important;
        max-height: 178px !important;
        border: 1px solid {border_color} !important;
        border-radius: 9px !important;
        padding: 6px 8px !important;
        margin-left: auto !important;
        margin-right: 0 !important;
        background-color: {_overview_upload_panel_bg} !important;
        overflow: visible !important;
    }}
    div.st-key-overview_activity_upload_card [data-testid="stCaptionContainer"],
    div.st-key-overview_model_upload_card [data-testid="stCaptionContainer"],
    div.st-key-overview_promises_upload_card [data-testid="stCaptionContainer"],
    div.st-key-overview_activity_upload_card [data-testid="stCaptionContainer"] p,
    div.st-key-overview_model_upload_card [data-testid="stCaptionContainer"] p,
    div.st-key-overview_promises_upload_card [data-testid="stCaptionContainer"] p {{
        color: {_overview_upload_caption_color} !important;
    }}
    @media (max-width: 1040px) {{
        div.st-key-overview_upload_header [data-testid="stHorizontalBlock"]:has(.st-key-overview_activity_upload_card) {{
            flex-wrap: wrap !important;
        }}
    }}

    /* Activity: compact uploader and adjacent activity kind, no redundant title. */
    div.st-key-snb_activity_header_row {{
        box-sizing: border-box !important;
        width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
    }}
    div.st-key-snb_activity_header_row [data-testid="stHorizontalBlock"]:has(.st-key-activity_file_upload_card) {{
        direction: rtl !important;
        display: flex !important;
        flex-wrap: nowrap !important;
        align-items: center !important;
        justify-content: flex-start !important;
        gap: 12px !important;
        width: 100% !important;
    }}
    div.st-key-snb_activity_kind_compact {{
        box-sizing: border-box !important;
        width: 172px !important;
        max-width: 100% !important;
        margin-left: 0 !important;
        margin-right: auto !important;
    }}
    div.st-key-snb_activity_kind_compact [data-testid="stSelectbox"] {{
        width: 172px !important;
        max-width: 100% !important;
        min-width: 0 !important;
        margin: 0 !important;
    }}
    div.st-key-snb_activity_kind_compact [data-testid="stSelectbox"] > div {{
        width: 100% !important;
        max-width: 100% !important;
        min-width: 0 !important;
    }}
    div.st-key-snb_activity_header_row div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-activity_file_upload_card),
    div.st-key-snb_activity_header_row div[data-testid="stVerticalBlockBorderWrapper"]:has(.st-key-activity_model_upload_card) {{
        height: auto !important;
        min-height: 0 !important;
        max-height: none !important;
        width: min(315px, 100%) !important;
        max-width: 315px !important;
        padding: 7px 8px !important;
        margin: 0 !important;
    }}
    div.st-key-snb_activity_daily_options [data-testid="stHorizontalBlock"] {{
        direction: rtl !important;
        align-items: flex-end !important;
        justify-content: flex-start !important;
        gap: 14px !important;
    }}
    div.st-key-snb_activity_daily_options [data-testid="stCheckbox"] {{
        margin-bottom: 10px !important;
    }}
    @media (max-width: 1060px) {{
        div.st-key-snb_activity_header_row [data-testid="stHorizontalBlock"]:has(.st-key-activity_file_upload_card) {{
            flex-wrap: wrap !important;
        }}
    }}

    /* Clickable dashboard toggle in the Activity header. */
    div[class*="st-key-activity_dashboard_toggle_"] button {{
        border-radius: 9px !important;
        min-height: 46px !important;
        width: 100% !important;
        font-weight: 900 !important;
        white-space: nowrap !important;
    }}
    div[class*="st-key-ejada_activity_dashboard_"] {{
        margin-top: 10px !important;
        margin-bottom: 16px !important;
    }}

    /* Activity summary: visibly distinct KPI cards in both SNB and STC. */
    div.st-key-ejada_activity_kpi_panel {{
        width: 100% !important;
        max-width: 100% !important;
        margin-top: 5px !important;
        margin-bottom: 12px !important;
    }}
    div.st-key-ejada_activity_kpi_panel [data-testid="stMetric"] {{
        box-sizing: border-box !important;
        border: 1px solid {border_color} !important;
        background: {card_color} !important;
        border-radius: 12px !important;
        min-height: 108px !important;
        padding: 12px 14px !important;
        box-shadow: 0 2px 7px rgba(0,0,0,0.05) !important;
        text-align: center !important;
    }}
    div.st-key-ejada_activity_kpi_panel [data-testid="stMetric"] label,
    div.st-key-ejada_activity_kpi_panel [data-testid="stMetricLabel"],
    div.st-key-ejada_activity_kpi_panel [data-testid="stMetricLabel"] p {{
        color: {text_color} !important;
        font-size: 17px !important;
        font-weight: 800 !important;
        text-align: center !important;
        justify-content: center !important;
    }}
    div.st-key-ejada_activity_kpi_panel [data-testid="stMetricValue"],
    div.st-key-ejada_activity_kpi_panel [data-testid="stMetricValue"] > div {{
        color: {text_color} !important;
        font-weight: 900 !important;
        font-size: 31px !important;
        text-align: center !important;
    }}

</style>
""", unsafe_allow_html=True)

# ==========================================
# 0. الصف العلوي: Overview وبجواره زر القمر
# ==========================================
top_cols = st.columns([0.48, 1.55, 7.97], gap="small")

with top_cols[0]:
    st.button("🌓", key="theme_toggle", on_click=toggle_theme, help="تغيير الثيم (كحلي / أسود / أبيض)")

with top_cols[1]:
    st.button("📊 Overview", key="top_overview_btn", type="primary" if st.session_state.current_page == 'Overview' else "secondary", use_container_width=False, on_click=set_page, args=('Overview',))

# ==========================================
# 1. الهيدر الرئيسي
# ==========================================
# وضع اسم البنك على يسار الشاشة مباشرة، بعيداً عن انعكاس أعمدة RTL.
# الشعار في منتصف الصفحة، ولا نستخدم وسوماً غير مغلقة.
# Keep EJADA centered; no duplicate green bank title above the bank selectors.
st.markdown(f"""
<div class="ejada-header" style="position:relative;width:100%;height:72px;margin-top:-54px;direction:ltr;pointer-events:none;z-index:1;">
  <div style="position:absolute;left:50%;top:0;transform:translateX(-50%);text-align:center;line-height:1.03;white-space:nowrap;">
    <div style="color:{header_logo_color};font-size:43px;font-weight:900;font-family:Tahoma,sans-serif;">إجـــادة</div>
    <div style="color:{header_logo_color};font-size:18px;letter-spacing:9px;font-weight:900;">EJADA</div>
    <div style="color:{header_logo_color};font-size:16px;font-weight:900;">لتحصيل الديون</div>
  </div>
</div>
""", unsafe_allow_html=True)

# Keep report navigation in the SAME row as SNB/STC, aligned to the top right.
def _render_ejada_main_navigation():
    with st.container(key="report_nav_group"):
        st.button("🏠", key="nav_home_btn", type="primary" if st.session_state.current_page == 'Home' else "secondary", use_container_width=False, on_click=set_page, args=('Home',), help="الرئيسية")
        st.button("🤝 الوعود القائمة والمكسورة", key="nav_prom_btn", type="primary" if st.session_state.current_page == 'Promises' else "secondary", use_container_width=False, on_click=set_page, args=('Promises',))
        st.button("🚨 تقرير الإهمال", key="nav_neg_btn", type="primary" if st.session_state.current_page == 'Neglect' else "secondary", use_container_width=False, on_click=set_page, args=('Neglect',))
        st.button("📈 النشاط", key="nav_act_btn", type="primary" if st.session_state.current_page == 'Activity' else "secondary", use_container_width=False, on_click=set_page, args=('Activity',))

if st.session_state.current_page not in ('Overview', 'Home'):
    bank_snb_col, bank_stc_col, nav_col = st.columns([0.82, 0.82, 8.36], gap='small', vertical_alignment='top')
    with bank_snb_col:
        st.button('🏦 SNB', key='bank_SNB',
                  type='primary' if st.session_state.selected_bank == 'SNB' else 'secondary',
                  use_container_width=False, on_click=set_bank, args=('SNB',))
        if st.session_state.selected_bank == 'SNB':
            st.markdown(
                f'<div style="color:{snb_green};font-size:19px;font-weight:900;'
                'width:116px;text-align:center;direction:ltr;white-space:nowrap;'
                'margin-top:-12px;">(30-60-NPL)</div>',
                unsafe_allow_html=True
            )
    with bank_stc_col:
        st.button('🏦 STC', key='bank_STC',
                  type='primary' if st.session_state.selected_bank == 'STC' else 'secondary',
                  use_container_width=False, on_click=set_bank, args=('STC',))

    with nav_col:
        _render_ejada_main_navigation()
elif st.session_state.current_page == 'Overview':
    # Overview must always expose navigation back to the three report pages.
    # Place the original upload widgets directly below this navigation row.
    _render_ejada_main_navigation()
    _overview_upload_header = st.container(key="overview_upload_header")
else:
    _render_ejada_main_navigation()

st.markdown(
    f"<hr style='margin: 10px 0 20px 0; border-color: {border_color}; clear: both;'>",
    unsafe_allow_html=True
)

# Click on empty canvas to close the current report and return Home.
# All controls, file uploads, charts and actual report content are excluded.
_BLANK_CLICK_SCRIPT = r"""
<script>
(() => {
  try {
    const root = window.parent !== window ? window.parent : window;
    const doc = root.document;
    if (root.__ejadaBlankCloseHandler) {
      doc.removeEventListener('click', root.__ejadaBlankCloseHandler, true);
      root.__ejadaBlankCloseHandler = null;
    }
    if (!__ACTIVE__) return;
    const excluded = [
      'button','a','input','select','textarea','label','iframe',
      '[role="button"]','[role="tab"]','[data-testid="stHeader"]',
      '[data-testid="stFileUploader"]','[data-testid="stMarkdownContainer"]',
      '[data-testid="stPlotlyChart"]','[data-testid="stDataFrame"]',
      '[data-testid="stImage"]','[data-testid="stAlert"]',
      '[data-testid="stExpander"]','[data-testid="stTabs"]',
      '.kpi-card','.kpi-grid-4','.kpi-grid-6',
      '[data-testid="stVerticalBlockBorderWrapper"]'
    ].join(',');
    const empty = [
      '.stApp','main','[data-testid="stMain"]',
      '[data-testid="stMainBlockContainer"]','.block-container',
      '[data-testid="stVerticalBlock"]','[data-testid="stHorizontalBlock"]',
      '[data-testid="stColumn"]','[data-testid="stElementContainer"]'
    ].join(',');
    const listener = e => {
      if (e.button !== 0 || !e.target || e.target.nodeType !== 1) return;
      if (e.target.closest(excluded) || !e.target.matches(empty)) return;
      const home = doc.querySelector('div.st-key-nav_home_btn button');
      if (home) home.click();
    };
    root.__ejadaBlankCloseHandler = listener;
    doc.addEventListener('click', listener, true);
  } catch (_) {}
})();
</script>
"""
_BLANK_CLICK_SCRIPT = _BLANK_CLICK_SCRIPT.replace(
    '__ACTIVE__', 'true' if st.session_state.current_page != 'Home' else 'false'
)
try:
    st.html(_BLANK_CLICK_SCRIPT, unsafe_allow_javascript=True)
except TypeError:
    from streamlit.components.v1 import html as _legacy_html
    _legacy_html(_BLANK_CLICK_SCRIPT, height=0, width=0)


# ==========================================
# دوال تقرير الوعود والإكسيل (مع Pivot Table حقيقي وتفاعلي 100% يفتح فوراً على Excel)
# ==========================================
PIVOT_TEAM = 'Sales Team'
PIVOT_COLL = 'Salesperson'
PIVOT_NOTE = 'ملاحظات-التمويل'
PIVOT_ACC = 'Account Number'
PIVOT_AMT = 'Net Amount'
FIXED_NOTES = ['NPL', 'Dpd60', 'Dpd30']
SUM_FMT = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'
COUNT_FMT = '#,##0'
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
        return 'Dpd60'
    if key in ('DPD30', '30DPD'):
        return 'Dpd30'
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


def style_pivot_frame(frame):
    """Center and bold Pivot column headings; show numbers with commas and no decimals."""
    numeric_cols = frame.select_dtypes(include='number').columns.tolist()
    number_formats = {col: '{:,.0f}' for col in numeric_cols}
    return (
        frame.style
        .format(number_formats, na_rep='')
        .set_table_styles([
            {'selector': 'th', 'props': [
                ('text-align', 'center'),
                ('font-weight', 'bold'),
                ('vertical-align', 'middle'),
            ]},
            {'selector': 'th.col_heading', 'props': [
                ('text-align', 'center'),
                ('font-weight', 'bold'),
            ]},
            {'selector': 'td', 'props': [('vertical-align', 'middle')]},
        ])
        .set_properties(subset=numeric_cols, **{'text-align': 'center'})
    )


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
    ws.sheet_view.rightToLeft = False
    if df is None or len(df.columns) == 0:
        return
    df = _prepare_for_output(df)
    bold = Font(name='Calibri', bold=True)
    for j, c in enumerate(df.columns, 1):
        clean_col = _clean_excel_str(str(c))
        ws.cell(row=1, column=j, value=clean_col).font = bold
        sample_lengths = [
            len(str(v)) for v in df[c].head(500).tolist()
            if v is not None and not (isinstance(v, float) and pd.isna(v))
        ]
        content_width = max([len(clean_col)] + sample_lengths)
        ws.column_dimensions[get_column_letter(j)].width = min(44, max(9, content_width + 2))
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
            column_name = str(df.columns[j - 1]).strip()
            if isinstance(v, datetime):
                if column_name in ['Next_Call_DateTime', 'Created_DateTime', 'وقت المكالمة التالية لنفس المحصل']:
                    cell.number_format = 'yyyy-mm-dd hh:mm:ss AM/PM'
                else:
                    cell.number_format = 'yyyy-mm-dd'
            elif isinstance(v, time):
                cell.number_format = 'hh:mm:ss AM/PM'
            elif isinstance(v, date):
                cell.number_format = 'yyyy-mm-dd'
            elif isinstance(v, float):
                cell.number_format = '#,##0.00'



def _break_overlap_minutes_for_report(start_dt, end_dt):
    """Subtract only the overlap with the daily 12:00–12:30 break."""
    if pd.isna(start_dt) or pd.isna(end_dt) or end_dt <= start_dt:
        return 0.0
    total_overlap = 0.0
    current_day = start_dt.normalize()
    end_day = end_dt.normalize()
    while current_day <= end_day:
        break_start = current_day + pd.Timedelta(hours=12)
        break_end = current_day + pd.Timedelta(hours=12, minutes=30)
        overlap_start = max(start_dt, break_start)
        overlap_end = min(end_dt, break_end)
        if overlap_end > overlap_start:
            total_overlap += (overlap_end - overlap_start).total_seconds() / 60.0
        current_day += pd.Timedelta(days=1)
    return total_overlap


def _attach_wasted_time_metrics(activity_df, collector_col, datetime_col='Created_DateTime',
                                status_col='Call_Status'):
    """
    Calculate the gap and wasted time on the CURRENT call row, matching the
    reference Excel report:
      successful call -> 20 minutes allowed
      unsuccessful/covered call -> 5 minutes allowed
      subtract the 12:00–12:30 break overlap before calculating waste.
    """
    result = activity_df.copy()
    if collector_col not in result.columns or datetime_col not in result.columns:
        result['فرق التوقيت (دقائق)'] = 0.0
        result['فرق_التوقيت_المعدّل (دقائق)'] = 0.0
        result['الوقت المهدر (دقائق)'] = 0.0
        return result

    result[datetime_col] = pd.to_datetime(result[datetime_col], errors='coerce')
    result = result.dropna(subset=[datetime_col]).copy()
    result = result.sort_values([collector_col, datetime_col], kind='mergesort').copy()

    previous_dt = result.groupby(collector_col)[datetime_col].shift(1)
    raw_gap = (result[datetime_col] - previous_dt).dt.total_seconds() / 60.0
    raw_gap = raw_gap.fillna(0).clip(lower=0)

    break_overlap = pd.Series(
        [
            _break_overlap_minutes_for_report(prev_dt, curr_dt)
            for prev_dt, curr_dt in zip(previous_dt.tolist(), result[datetime_col].tolist())
        ],
        index=result.index,
        dtype='float64'
    )
    adjusted_gap = (raw_gap - break_overlap).clip(lower=0)

    if status_col in result.columns:
        thresholds = result[status_col].astype(str).str.strip().map(
            lambda status: 20.0 if status == 'ناجحة' else 5.0
        )
    else:
        thresholds = pd.Series(5.0, index=result.index)

    # Round the gap first, matching the supplied Excel sample calculations.
    result['فرق التوقيت (دقائق)'] = raw_gap.round(2)
    result['فرق_التوقيت_المعدّل (دقائق)'] = adjusted_gap.round(2)
    result['الوقت المهدر (دقائق)'] = (
        (adjusted_gap.round(2) - thresholds).clip(lower=0).round(2)
    )
    result['التوقيت'] = result[datetime_col].dt.strftime('%H:%M:%S')
    return result


def _build_reference_activity_frames(processed_df):
    """Build the report tabs that match the supplied SNB Excel workbook."""
    df = processed_df.copy()
    if df.empty:
        empty = pd.DataFrame()
        return {
            'Productivity_Report': empty,
            'Calls_Over_30_Min': empty,
            'Calls_Under_1_Min': empty,
            'First_Call_Per_Collector': empty,
            'Last_Call_Per_Collector': empty,
            'Wasted_Time_Per_Collector': empty,
            'Main State_Distribution': empty,
        }

    collector_col = 'Collector'
    gap_col = 'فرق التوقيت (دقائق)'
    waste_col = 'الوقت المهدر (دقائق)'
    adjusted_col = 'فرق_التوقيت_المعدّل (دقائق)'
    status_col = 'التصنيف'
    time_col = 'التوقيت'
    notes_col = 'Notes'
    created_col = 'Created on'

    df[created_col] = pd.to_datetime(df[created_col], errors='coerce')
    df[gap_col] = pd.to_numeric(df[gap_col], errors='coerce').fillna(0).round(2)
    df[waste_col] = pd.to_numeric(df[waste_col], errors='coerce').fillna(0).round(2)
    df[adjusted_col] = pd.to_numeric(df[adjusted_col], errors='coerce').fillna(0).round(2)
    df[status_col] = pd.to_numeric(df[status_col], errors='coerce').fillna(0).astype(int)
    df = df.sort_values([collector_col, created_col], kind='mergesort').copy()

    # Productivity report: KPI score/classification are left blank because
    # their calculation rule is not present in the supplied source workbook.
    productivity_rows = []
    for collector_name, group in df.groupby(collector_col, sort=True):
        total = len(group)
        successful = int((group[status_col] == 1).sum())
        unsuccessful = total - successful
        avg_gap = round(float(group[gap_col].mean()), 2) if total else 0.0
        success_rate = round((successful / total * 100), 2) if total else 0.0
        productivity_rows.append({
            'Collector': collector_name,
            'إجمالي_المكالمات': total,
            'عدد_الناجحة': successful,
            'عدد_غير_الناجحة': unsuccessful,
            'متوسط_فرق_التوقيت_بالدقائق': avg_gap,
            'نسبة_النجاح_%': success_rate,
            'مؤشر_الأداء_%': None,
            'تصنيف_KPI': None,
        })
    productivity = pd.DataFrame(productivity_rows)

    def _short_calls_frame(subset):
        if subset.empty:
            return pd.DataFrame(columns=['المحصل', 'التوقيت', 'فرق التوقيت (دقائق)', 'الإفادة'])
        return pd.DataFrame({
            'المحصل': subset[collector_col].astype(str).tolist(),
            'التوقيت': subset[time_col].astype(str).tolist(),
            'فرق التوقيت (دقائق)': subset[gap_col].astype(float).tolist(),
            'الإفادة': subset[notes_col].tolist() if notes_col in subset.columns else [None] * len(subset),
        })

    over_30 = _short_calls_frame(df[df[gap_col] > 30].copy())
    under_1 = _short_calls_frame(df[(df[gap_col] > 0) & (df[gap_col] < 1)].copy())

    first_rows, last_rows = [], []
    for collector_name, group in df.groupby(collector_col, sort=True):
        group = group.sort_values(created_col, kind='mergesort')
        first_call = group.iloc[0]
        last_call = group.iloc[-1]
        first_rows.append({
            'المحصل': collector_name,
            'توقيت أول إفادة': str(first_call[time_col]),
            'حالة النجاح': 'ناجحة' if int(first_call[status_col]) == 1 else 'غير ناجحة',
        })
        last_rows.append({
            'المحصل': collector_name,
            'توقيت آخر إفادة': str(last_call[time_col]),
            'حالة النجاح': 'ناجحة' if int(last_call[status_col]) == 1 else 'غير ناجحة',
        })

    wasted_rows = []
    for collector_name, group in df.groupby(collector_col, sort=True):
        positive = group.loc[group[waste_col] > 0, waste_col]
        wasted_rows.append({
            'المحصل': collector_name,
            'إجمالي_الوقت_المهدر (دق)': round(float(group[waste_col].sum()), 2),
            'عدد_المكالمات_المهدرة': int((group[waste_col] > 0).sum()),
            'متوسط_الوقت_المهدر (دق)': round(float(positive.mean()), 2) if len(positive) else 0.0,
            'أعلى_وقت_مهدر (دق)': round(float(positive.max()), 2) if len(positive) else 0.0,
        })
    wasted_summary = pd.DataFrame(wasted_rows).sort_values(
        'إجمالي_الوقت_المهدر (دق)', ascending=False
    ).reset_index(drop=True)

    if 'Main State' in df.columns:
        state_counts = df['Main State'].fillna('غير محدد').astype(str).value_counts()
        state_rows = [{'الحالة': state, 'عدد_الحسابات': int(count)}
                      for state, count in state_counts.items()]
        state_rows.append({'الحالة': 'الإجمالي', 'عدد_الحسابات': int(len(df))})
        state_distribution = pd.DataFrame(state_rows)
    else:
        state_distribution = pd.DataFrame(columns=['الحالة', 'عدد_الحسابات'])

    return {
        'Productivity_Report': productivity,
        'Calls_Over_30_Min': over_30,
        'Calls_Under_1_Min': under_1,
        'First_Call_Per_Collector': pd.DataFrame(first_rows),
        'Last_Call_Per_Collector': pd.DataFrame(last_rows),
        'Wasted_Time_Per_Collector': wasted_summary,
        'Main State_Distribution': state_distribution,
    }


def _write_df_to_template_sheet(ws, frame):
    """Replace worksheet data while retaining the template's formatting."""
    frame = frame.copy() if frame is not None else pd.DataFrame()
    headers = list(frame.columns)

    # Keep the original header styling, but update header text to the frame schema.
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.alignment = copy(cell.alignment)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Clear old values, preserving the sample workbook's row styles and widths.
    old_max_row = ws.max_row
    max_col = max(ws.max_column, len(headers))
    for row_idx in range(2, old_max_row + 1):
        for col_idx in range(1, max_col + 1):
            ws.cell(row=row_idx, column=col_idx).value = None

    # Extend the alternating row style if the new report is longer than the template.
    for row_idx, values in enumerate(frame.itertuples(index=False, name=None), start=2):
        if row_idx > old_max_row:
            style_source_row = 2 if row_idx % 2 == 0 else 3
            for col_idx in range(1, max_col + 1):
                source_cell = ws.cell(row=style_source_row, column=col_idx)
                target_cell = ws.cell(row=row_idx, column=col_idx)
                if source_cell.has_style:
                    target_cell._style = copy(source_cell._style)
                if source_cell.number_format:
                    target_cell.number_format = source_cell.number_format
        for col_idx, value in enumerate(values, 1):
            if value is None or (not isinstance(value, (list, tuple)) and pd.isna(value)):
                continue
            if isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()
            elif hasattr(value, 'item') and not isinstance(value, (str, bytes)):
                value = value.item()
            elif isinstance(value, str):
                value = _clean_excel_str(value)
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            header = str(headers[col_idx - 1]).strip()
            if isinstance(value, datetime):
                cell.number_format = 'yyyy-mm-dd hh:mm:ss'
            elif isinstance(value, time):
                cell.number_format = 'hh:mm:ss'
            elif isinstance(value, (float, int)) and not isinstance(value, bool):
                if any(token in header for token in ['دقائق', 'الوقت', 'النسبة', '%']):
                    cell.number_format = '#,##0.00'

    # Remove unused trailing rows so the sheet opens at the actual report size.
    final_row = max(1, len(frame) + 1)
    if ws.max_row > final_row:
        ws.delete_rows(final_row + 1, ws.max_row - final_row)
    if ws.title == 'Processed_Data':
        ws.auto_filter.ref = f"A1:{get_column_letter(max(1, len(headers)))}{final_row}"


def _build_processed_activity_export(activity_df, template_headers, collector_col,
                                     created_col, notes_col):
    base = activity_df.copy()
    if 'Created_DateTime' not in base.columns and created_col and created_col in base.columns:
        base['Created_DateTime'] = pd.to_datetime(base[created_col], errors='coerce')
    elif 'Created_DateTime' in base.columns:
        base['Created_DateTime'] = pd.to_datetime(base['Created_DateTime'], errors='coerce')

    if collector_col and collector_col in base.columns:
        sort_cols = [collector_col]
        if 'Created_DateTime' in base.columns:
            sort_cols.append('Created_DateTime')
        base = base.sort_values(sort_cols, kind='mergesort').copy()

    def _source_series(header):
        col = find_col_in(base, [header])
        return base[col] if col else pd.Series([None] * len(base), index=base.index)

    output = pd.DataFrame(index=base.index)
    for header in template_headers:
        if header == 'Collector':
            output[header] = base[collector_col] if collector_col in base.columns else None
        elif header == 'Created on':
            if created_col and created_col in base.columns:
                output[header] = base[created_col]
            elif 'Created_DateTime' in base.columns:
                output[header] = base['Created_DateTime']
            else:
                output[header] = None
        elif header == 'التوقيت':
            output[header] = (
                base['Created_DateTime'].dt.strftime('%H:%M:%S')
                if 'Created_DateTime' in base.columns else None
            )
        elif header == 'فرق التوقيت (دقائق)':
            metric = base['Gap_Minutes'] if 'Gap_Minutes' in base.columns else (
                base[header] if header in base.columns else pd.Series(0.0, index=base.index)
            )
            output[header] = pd.to_numeric(metric, errors='coerce').fillna(0).round(2)
        elif header == 'الوقت المهدر (دقائق)':
            metric = base['Wasted_Minutes'] if 'Wasted_Minutes' in base.columns else (
                base[header] if header in base.columns else pd.Series(0.0, index=base.index)
            )
            output[header] = pd.to_numeric(metric, errors='coerce').fillna(0).round(2)
        elif header == 'فرق_التوقيت_المعدّل (دقائق)':
            metric = base['Effective_Gap_Minutes'] if 'Effective_Gap_Minutes' in base.columns else (
                base[header] if header in base.columns else pd.Series(0.0, index=base.index)
            )
            output[header] = pd.to_numeric(metric, errors='coerce').fillna(0).round(2)
        elif header == 'التصنيف':
            if 'Call_Status' in base.columns:
                output[header] = base['Call_Status'].astype(str).apply(
                    lambda value: 1 if 'ناجحة' in value and 'غير' not in value else 0
                )
            else:
                output[header] = _source_series(header)
        elif header == 'Notes':
            output[header] = base[notes_col] if notes_col and notes_col in base.columns else _source_series(header)
        else:
            output[header] = _source_series(header)

    output = output.reset_index(drop=True)
    return output



# القالب المرجعي مدمج داخل الكود؛ لا يلزم رفع ملف Excel منفصل بجانب app.py.
_ACTIVITY_TEMPLATE_B64 = (
    "UEsDBBQABgAIAAAAIQCg7NQdnAEAAHoJAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADMlsluwjAQhu+V+g6Rr1VioC2lFYFDl2OLVPoAbjxJ"
    "LOJFttnevrZZRFEKQiA1lzixPf//eSJ5pj9c8CqagTZMihS1kxaKQGSSMlGk6Gv8FvdQZCwRlFRSQIqWYNBwcH3VHy8VmMhFC5Oi"
    "0lr1hLHJSuDEJFKBcCu51JxY96kLrEg2IQXgTqvVxZkUFoSNrddAg/4L5GRa2eh14aZXJEoUKHpe7fNWKWLcx/t5XBuhoTJ7IUSp"
    "imXEurPhmaB7XPGaKXGRYY8pmTI3DvwPB7/ym2nXYB334ZKpGYVoRLR9J9yR40WF51JPvqWcJIdFaihlnrMMqMym3OUsMUoDoaYE"
    "sLxKwphwwsSG+4B/2GxwGNoXBvHnC8IncnQawnHbEI67hnDcN4Sj2xCOh4Zw9BrC8fhPHNZVGMDhef4VFmSOXFjGLiswl762g+gx"
    "55JooJ9Wu1p8cYBd7SMcVJO5R8Drl/PzvhY65Ovq3UhLZVyvoOH07G9Ku4+OlRMCbRlsi3tdkdw6uqJ+9u8G38lQoDXeOHROgx8A"
    "AAD//wMAUEsDBBQABgAIAAAAIQC1VTAj9AAAAEwCAAALAAgCX3JlbHMvLnJlbHMgogQCKKAAAgAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArJJNT8MwDIbvSPyHyPfV3ZAQQkt3QUi7IVR+gEncD7WN"
    "oyQb3b8nHBBUGoMDR3+9fvzK2908jerIIfbiNKyLEhQ7I7Z3rYaX+nF1ByomcpZGcazhxBF21fXV9plHSnkodr2PKqu4qKFLyd8j"
    "RtPxRLEQzy5XGgkTpRyGFj2ZgVrGTVneYviuAdVCU+2thrC3N6Dqk8+bf9eWpukNP4g5TOzSmRXIc2Jn2a58yGwh9fkaVVNoOWmw"
    "Yp5yOiJ5X2RswPNEm78T/XwtTpzIUiI0Evgyz0fHJaD1f1q0NPHLnXnENwnDq8jwyYKLH6jeAQAA//8DAFBLAwQUAAYACAAAACEA"
    "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"
    "ZXRzL3NoZWV0MS54bWyck11vmzAUhu8n7T9Yvg+GtMlWFFJty9gibVK1drs35gBWsI2OnSbR1P++A1lIpdxERWBsH/s5H369uN+b"
    "lj0Deu1sxpMo5gyscqW2dcZ/P+WTj5z5IG0pW2ch4wfw/H75/t1i53DjG4DAiGB9xpsQulQIrxow0keuA0uWyqGRgYZYC98hyHLY"
    "ZFoxjeO5MFJbfiSkeA3DVZVWsHJqa8CGIwShlYHi943u/Ilm1DU4I3Gz7SbKmY4QhW51OAxQzoxK17V1KIuW8t4nt1KxPdI7pe/m"
    "5GaYv/BktELnXRUiIotjzJfp34k7IdVIusz/KkxyKxCedX+AZ9T0bSEls5E1PcNu3gibj7C+XJhudZnxv/H/Z0L/pG/ic3OyvfDl"
    "YtDJAy4XQRZfXOuQYV1kPM+TfJ5//czFciHGNaUmOfQlYAhVxj8lvXmw/tGw86/6zDdu9w11+UNbINmS4MnBI7SgAlB8CWe9tAvn"
    "Nv3WNU3FZ1+vafkg7QdkJVRy24ZfbvcddN0Egswo4V4xaXlYgVckVcJE09kY1koGSSF2soafEmttPWuhGlZ9oM145MRRPwiuG+Iq"
    "XAjODN2GLhKQYOKI7JVz4TQgZolyR9eXYdrXG9flUAwxXtjlPwAAAP//AwBQSwMEFAAGAAgAAAAhAAteESzDRQEAQ+8HABgAAAB4"
    "bC93b3Jrc2hlZXRzL3NoZWV0Mi54bWycfduSG0eW5Pua7T/Q+N5F5D0hk3qM3aouFAtFNkipAOGNTZVatCFFLcm+2dr++57MiEzE"
    "cXdkJoamaXEUnidOXBwnLh4R3/7Xvz9+ePLPx89f3n/67bun2dXq6ZPH3959+vn9b3//7umPP/zlD+3TJ1++vv3t57cfPv32+N3T"
    "/zx+efpff/zf/+vbf336/N9ffn18/PrELPz25bunv379+vs3z559effr48e3X64+/f74m6X88unzx7df7f/9/PdnX37//Pj25/6j"
    "jx+e5atV/ezj2/e/PQ0Wvvm8xManX355/+7x+0/v/vHx8bevwcjnxw9vv5r/X359//uXwdrHd0vMfXz7+b//8fsf3n36+LuZ+Nv7"
    "D++//qc3+vTJx3ff3P79t0+f3/7tg5X731n59t2Tf3+2f3L7v2LIpv/vlNPH9+8+f/ry6ZevV2b5WfCZi79+tn729t1oicu/yExW"
    "Pvv8+M/3XQOeTOX/M5eyarSVn4wV/0Nj9Wisq67P3/zj/c/fPf2/q/jnD/bvrPuf1R9W1hf6vw1p/+/pH7/9+b21cFeqJ58ff/nu"
    "6fPsm+fPq1X19Nkfv+270MP7x399Sf7+5Oun37ePv3z98+OHD989/VP29EnXRf/26dN/d8hby3plVr88fnh813WWJ2/tX/98DOg/"
    "W2m//J8+H/ur5fBszCL9+5DdX/pO/dfPT35+/OXtPz58ff3pX5vH93//9asxqLKidn3lm5//8/3jl3fWSS3jq7z3+92nD2bC/vfJ"
    "x/cd26yTvf33d0+tdv71/uevv373NDe2vfvHl6+fPu7Df8g6Z8YPyviB/Xv4wP468YE50+dQnz6wkk580MQP7N8xh8y8nPjAHO5z"
    "sH8PLk3nsI4f2L+HD6z8Ezlk9pPUZ9H9JX5S2F+nPhlqtnN+KMdMLuZ1yMX+Ej9ZTxc9M4vhk1MDdj5OOTa0YNdNhlyePvn1/c8/"
    "P4YOMfWxtWLI79ScmTXU1CdDe3a4obqt5qc+GVo0OzVpOd0ts6FNu78MuUzXXT40aveXoVGnWygfGrX7y7Lunw+N2v1l+MQqb6L4"
    "+dCo3V+GrjPzydCoecLLaRLYz0Foyu4vMZdymsr50PrdX5ZRLR9av/vLUBbZ+s/CL1P/o/f9269v//jt50//emIByar6y+9vu/Ce"
    "fdNZ0T9t9hPVgZ936P4b+837Yj+2//zj6ttn/7Rf0HcR0f0qdwhr9BGSNVnmUX/WqNqjvufcwM41I3Jv4y+MKDziRuSSteDKRoIa"
    "b+lWgloPeiFBaw+6Y1DpEVtGVB5xrzJaQ3O9lCCo5VcSBBX9VwZBHe4YARX4mhFQe28YAVX3g3AWCv2jKo+vvAcBgQLvBQS61kFA"
    "oB1/EhBoyKOAQNU+F8TMTpX7zLg+Er4bCC0nfIe2UUvKZmiSPwlIk0Fl9cMvsNNkUF3fs6Uig/q6ZkwOzfuXACn6n6eyzsviqlgX"
    "WV71Y09w7EZ4n63BsU0A1eoH73Yi7cVE2l1Ia3qb2apeVWWR1VCUraoRaPv7AdP9GgNzX6Zpq6t12azTP9DRXg3obkTa/dz/1flY"
    "W/2VTVtVRbuCCtqJWlyD9dcKA2V5I1oXzPygzMDvyI+q2qAjPaRlq+qybsu6afPcWsL/FOyVMXDqEDCnLgfl+knZAKePYKMAnj1/"
    "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"
    "wyCY4JKKCjUMibqFQ+a5YHwl2a7klWlIFC+ogfj1VtT6zFayW+hKql539YjJsgaKtSiEeXOkyCtZ8ryppKLrAnQfSUSZd1TwC+E/"
    "Prtx73ueSsGU1DI3PjDjJua+/DmeY8papr7+m2hIhBXfC1vAC1X4uZBI3HKFF7LxJ8kmLZlNl0p2IkvR3+B0jWAmdghGQWSHd9c/"
    "tFxkAipsVXmK5ym6J8kqQni5cP3zW/CDfnfv2XZcS/lqF57BTWChuId9cu34Q3kZz+muMC/y8I2LzdZA78cQpK1ykr2tuGbQXkDj"
    "h7FlYrIAdzB6pbD7BNqDHt18EJnZpiiMkMd22sjyT/OCnMwaA8imM4D5ZECA4wMDyJgzgPnsAWw/MAD/zgDmk0E8HTLAjRSXmRU1"
    "dLlQ8uBB70E8uqZ2J5MESIZTATmw2HsLdiaQIw3F2i/JZD5b4D0knZ1ADwMgMptcgx4HQdNr0GrQ3bwFYZDQ6oA03a7DgqF4yLvo"
    "mEYdGQOYKOrKaEATl4ox8eOugj7JJLhgrgRAyW8XYMEpgnq1AmZxJ/4+hESzoFOGBjRr4h/7YaeWqz5JeEnBVfgQy+3hW/B1/sNx"
    "J7UPfQyJ447IxwZ0yn/gj8NOBw05umhsBDTHRbMparrh36naiEp7Bc/dUTCFE0I1h0Xg2wcja7cJ1tLAxne3W/jCcdghgQ/ruZTm"
    "/GCPo/abufwPAAD//wMAUEsDBBQABgAIAAAAIQAar6Iy4gkAAKw+AAAYAAAAeGwvd29ya3NoZWV0cy9zaGVldDUueG1snJtRb9s4"
    "FoXfF9j/YOi9sURSkh3EGUxsDHYeBlgMZnafVUdJjNqWYTtNi8X+96EoihbvPazECdq4aY4OeXlJfSRFPfz07bCffa3Pl11zXCXZ"
    "XZrM6uO2ed4dX1fJn3/88mmRzC7X6vhc7ZtjvUq+15fkp8d//uPhozl/ubzV9XWmHY6XVfJ2vZ7u5/PL9q0+VJe75lQf9W9emvOh"
    "uuofz6/zy+lcV8/mosN+LtK0mB+q3THpHO7PUzyal5fdtt402/dDfbx2Jud6X111/S9vu9Oldztsp9gdqvOX99OnbXM4aYvPu/3u"
    "+t2YJrPD9v7X12Nzrj7vddzfMlVtZ9/O+o/Qf2VfjPl/VtJhtz03l+bleqed512defjL+XJebZ0Tj3+STabm5/rrrk3gzUr8vSpl"
    "ufMSNzP5N80KZ9Y21/n+ffe8Sv6X2q9P+jNrv6Wf0rz9Nvj6f/L48LzTGW6jmp3rl1Xyc3a/KYtk/vhgOtB/dvXHZfDvWdsfPzfN"
    "l/YXv+py0lY6Z9pfTH/893n2XL9U7/vr783Hv+rd69tVd/5c17JN8/3z90192er+pW3uRN46bZu9Lk5/nx127UDR/aP6Zj4/ds/X"
    "t1UiVDLbvl+uzeG/3X9k9rLuAt2c5gL9aS/ItMcPLtBNZi7Qn30J+tofXKDLNxfoT3tBoa/lF8y7UEzLbKpr9fhwbj5muvPp+lxO"
    "VTuUs3ttgptCt0Gr/bkVm0t0G110tr4+ZsVy8TD/qht9a0VPQJQtCl+0hqLSF21gcUsnmusQXBy6mabH0Yp18pKZi0NmioTBNZm+"
    "eZEwOlFhmiK9WwoSAPcQ8tZcXv111qbXvxWvEp2uH9SfazKRZaT+nWhh67+gCeAewfrrykyvfysea3+u0fUnDbzuRH37LyRpf+4R"
    "rL++DUyvfysea3+u0fXPSft3or79b13bDLcNt8jFbSB53aeIqX4rHmt+rsmEIO277kSu+Wn35x7B5i9j6t+K/ebPSNM9AYmQtPd3"
    "or7183LpfZFb1oZbqlt7eMloZ0+T76mt2E8GCwZIhKRDoRP1uShJT9sAj0Fv9Kq/jKl+Kx7JBZAIRXPRifpcDFq2GwrAo7xZeNXP"
    "9Gw2AmmteqT9jSPRiJyOZavqM1AQLm6Ay3BE+THEYbkD5ZAHrA+1Ew+aKJEzLHcqd0NiMQCXNJSHKCRngLd0UCONyAm01lbloMxi"
    "AFS+DSY/DVFUzgByWQgIywWdVlinPg0FHcygJBEKIQrMGaAuCwGRuaBosE59FsjsagMKCt1N2xl6xHgG3GURIDbT8bo25a6SPgkU"
    "fhv7++GoU8WtM/odKYrPGYAviwEBumQdySe0YmOBuwTTEIXobAKjkUaUpKesrcpRms7xgEtwlpRFkdmox9CA2FySZK2tUz8Y8sKf"
    "abDhzU2DwzsK1tkEWiONWFBcW5XLil7LD7/YcOcFh/qZiMK3UY/kCGnEguLbqhw2aFKASygpIoreRj0yhUIasaT3XKsKDhXgEsxC"
    "3Hp6ArwFWlAvWRb8FbViWeAu5eJ21/O3BKLoLSbQG2nEkk5ArMqtqulNF7mEJlHtBlPEvsYEfBtHOpld0juWVQUns8BlkYXyEAVw"
    "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"
    "cXxtcHSb7+/+6QJjEv/l6OP34z+8mR/k/X92t/bxrds3r598NFh/YKPc81jnT94PRu/8xyeD88Ed8u6rvD0eU43F62se/U83wg+/"
    "aW+wS4R/Pzo49+uDd/cHe3uDi9c/Or7BD92H4i36c37X+s/XTz50bzn++MOr3KDGnoE+d3zjv3ffMp5OhovReD4dd/9y9unp/bOv"
    "Tr8//db95bPTx/ztHu/424AfPj19eHbv9KezTwdnn/LL87PPzr46+7T7Kd702eDMffT+6cPTn3SRwekDu1b++OBf/oXLlVvQB5+X"
    "77rLTd0//d59wQvd7uDsk9PnulD0PAPe8Yh/HrjP3j/7/PTHweljvufbsy/tPvXjgKf59OzLYOxteubD83/58NUn6fDkxrGbpPF4"
    "OhyP5rP5MLipT7ndezY63M0LbuWOe9MjjcXZVwPG8TH3/tXpNwM+okF9zC+P06A94Ec9IqP9Khewidj2mefM5ldnd5j/weD0CfP0"
    "rS2Ir9zymo1X8+lsuBpFyyF44DtMny0GPc1PWkdn7qK2kuwWf9T3Mku/HmjDhCtu431nd/n1M601voMRfXj69ekTW01+WFji6Yts"
    "BQxOn519MUgr9fRxsB5ae/Hw+Nafj7tvGQ1nk/F8Pp66keibWU0XA/AJq1J3rcX44Ozz1tLUzvuR8WL2/bgP58vRZD6cxYNSBo7L"
    "PuAqL8qA2M45fXT6N/exr5ntHxmSZ4PTp2df6DuDN91/lTdFV+qbQS2rz1ncZ5+wpn/gBzYkN8HCcE88Wa6Gs+VoOHUbn6XPrX/P"
    "M6b9rWs8sBUrc/UsjXH+2/aZPXd86+Sa+2amdcEiX8yjDczV72H4+MLnmkHN2c8aP3eX9zA198pb6vzcYZzzxx4yEp/w63fBR++n"
    "qycLWj6sJdnsFdlojRxWutrSHlPda0K5R1uGDGfaeCyVH07/yhVZmK+wp9KKtSd6bgPPltJrD1hTst1pYtJrzzXfpz/4fTQcjsar"
    "4cqPNl5Brqe1Q57aBVlC2jpfnN2NrMl4PppOVsvZyG+V7+22NPxl4TSvyBZzfz/6lTgajqeL+WoyXbn1kDxUMc1uHj871Q3b7Wot"
    "yLJpe58+SiPylImL9vpoOV5Op9ORt7H3ZU3z+pFX+5yVJ5uShj0tTXcX1ez8yIq0uTL/LFfC+sIhsXx0k9Fq5/7uaLVrUr8pe7fH"
    "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"
    "6sSv7p+0V76EfbIPIfVl07eTP9pzqccJv4qIajyt9iHI+sLOcxUfiNe27lLHUT05pXPjgYF0y+/gD1r3p73xKOK+7H8ofPB6MXFc"
    "dVrTvVOfTepnOu0k7vN+NU7BMyWeAkqm8691Us3vJvUs8jLXFcF5XV/Mp61T8MHzhUldI7yiv6+QYdMByv8Fb8v/vd6U0nbzkyTF"
    "5odVJ9I6Dv7v+3PQi2ydRjfo/ln7c3S/kM5byDb6ieqW4LPWfZgKzZuSmtk64n3pumP8Fu9SD9U6TrZmWTwNRRd78EH3gWRrh62X"
    "lBrFjp+ojgeu6n4DcFL5UvQf40vgg+Yp0RstPobfqm4z0QPh517vAE8meojbdV8li5e4YPyZefC6huFn6g7EKcWBid6S/W9Wp4he"
    "oTzfeE54I35i/kaeLvvcRp6u+TJ+qP5JHNQ63YRnhldbXgBPVl4UHVjzCNu/VMFist9mso8l8ULz5WwpND6Z1Njs6AaaF4M/qk/y"
    "Hie6pe6Pyj4EzTs0zyIoW1zGD6lf+L7cSL6iTxJfVJdjXWicgnfpesw+QK3fkT/ac6MT2rqDhwvORISc6Oqat1YdTesCyePaeG5+"
    "4JhjzuxxMnHd3X64DlouGcrO95tfR8Ps7n/EqYCfz+f2tlGdBXpG2xzjy1mtmz+uJzvuc0HKx8fN/ecP42xBbaK/fU6u83t8/Vxr"
    "9snN97myYmzo/64uRP+mNztu2L9+4+3X2w/jvoKcc+3nSP55/MW9uvMo39zIEaEceT3/83+++vrfPz+nXd16xeGMSbv6wSfvpQ2j"
    "zjD4M8bfrle19DdzOMmQ0dZNRT9M2+a8R24eLOf4tD5v0r6rVndTuxOs1+rOveRkLN+O603nY2leWqc5bh32s+NpjhdmzjjOfXw0"
    "zn/kkNWdP9ZDGLeOWX/z/x/43bjX5nqcOLm+e9qkjy+TvE71U8Y+buiuw5ztQEyNL59W+Ou8g7fefvPc9M5r55ZXXjm3vPreu3aa"
    "h3XKdQE5D/nLzaMX15MxuSj5zvVIL73//gcv/wcAAP//AwBQSwMEFAAGAAgAAAAhAE2B0aQRAgAA8QMAABgAAAB4bC9kcmF3aW5n"
    "cy9kcmF3aW5nMS54bWycU9uO2yAQfa/Uf0C8O8beJHWtOKtmE1eRqjaq2g9gMY6RDFhALqvV/nsHsDfb9qWqH/ABhuHMOcPq/ip7"
    "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"
    "enf5tL8ZdzWI/iOTFl+b2sbXttUL/lqvnq/pypEeJw0U7SYlXvdlzd9tIzB/9s/KK/ulj3fa1pi/8sxdqilK2gxJvNLzeyffPDv5"
    "ZtGIos9OqIaBNix8MX28AKJO6HtAM/QSNHfm41K5+iwprt1x4Z737z649YfyM6drVJvUVx2vPPLDoX9fu3PRzeruz7qsuzuc4+fb"
    "1B75/NKTSvVZYtxas+bSA3tr/Xxx28G/UZVvztRvOG/3x88WHy6qtO3AWHv61OH/HPh4VNHu3arXTU4dzJ038MDP0EoERrB++uIW"
    "qfosLZw8/oerXrzomVXrCypOO0W8mtMnKg5sWr/wsbeGX/DlYtWQda0nD81r633Pg372gHp00X3bLr5o4V/WHbbtvOF0wY6/3P3G"
    "xa+p5odN7TJzsoeKdr9r+p7reGT8es47Y+/+YuVPxxzHWVtR9NOmv9z9ysUvqGebbTfjQh/HB/hG+EWzs/9/c6Tqs7RwsmrBDd/c"
    "+fyOHYWnz9ap5/WuPXvy9MGtez6e8+1Vl+79d6HjiZSJXX6R4X3PGX2u/6Xq8vFkxTN53778ZcmxSqV7n7HuZMmq57+7bdaZKsdG"
    "CXnX0gaIYBGBobHKeZbqs2TXhpr/e/TEnJVndpXWnlZ3Wa6vPV569j+flV/y4Nk56km3M+Jv7+et51jqw5erYmW1ccack/M2Vym7"
    "ra2u+s/fTl2w3Kw6z9Y/fVWqAKKJFnFVV1nlb0HNG+WwzDYmk0laIf3MzGgjAF9Ol313z9Hd//X+e90jefzH3fu0cY3rZe/sXjo3"
    "mCldDVPaXvNypxQ/t169762ba1zWdV4w6BJ/WxPLv7vUubjcuLh+q/v9jE4icFJ28rRer9fphOWnlbRe+UkERois+6jPPQUiCBld"
    "5v/z6p/7HjroyFeX/WrbHuGvC+be/vqFnm8BP/CPt5yLy43re9+1n1zT+F04gBvCL8Km7thns/63YIMpwJeJ7Cv7PHVnt3SXtRvW"
    "X/5QUFNaZ6Q98p57j2ZndYeem7Z3jd9zYrW7drVMj80AABAASURBVNBbt2YJIFBEYITFmQfvrf1IBEE3+3fn3Jzt49bluv/89fSt"
    "/p39jv9/bRaODXTq1ur//K3q1u3eSynJ+qdva3NdDmMeIGC+w29TMAY0Iqf041Jf1WfJgerVs0rdV2fe1LFfPxEow8/TLvW/+hwC"
    "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"
    "J77wL7U1cmGlvnE29Clj++HqMUJKZAmdnrpqqnDaeeKY4bft7bK2j2O4CXvWc9y/69dV4Z6Cnf/XaUxCYrt0EQzp2KbfeLGqGj7/"
    "q4IbpBya1v4l6ZpdhJi9+7On75ncDTgwCSHQFD4mA/Q5T6Ayr5HlP/9mPJCut9eurRMirY+nlyjvFdh0DT4Nujhd6QL2ZsdFe/fJ"
    "JxMBzKdkm0DD+2QLlk/hmP3JugAAgCc+JgP0OU+gMrPQUb+m7VV8+zM564k0+W5Z92mgmphtnavPwpIN7UfuNC+ixZ7P+jguVqXX"
    "Wq5gheW62rGZ02zDAIDmaGCSXJCtKC9wL1A2nsh8X1jJpVh7lhy0dGpTp8Kz9lKSCsf55U75SMpZ0holgdpWvjTVnvXSr/aYqb0e"
    "vFTmlv9on+qebZXE5zX3SeXgJRVCqYarrpq/3SyvTL8oNBfLPrjMS+wxdwdRDQBaAgrQQBP0bd1e+llXvcPDcyeOS/VikdLFQ148"
    "/7vhneyTG8n9i/26bpRq0Eo5WHqJ+hpVldgsKT/gHKbkP6eTFSXlK23O3/5MPpnwN1uru06rE610KmCZfVj+7/HSWsuxWk44pjsO"
    "HgAAD9KT5cRSfcZD16E9Z45JPxOSB7o/Nfoq1bS9Xi59PZBq0G923HikQunJ5ch9Tc62VuVLrNlQTuiej9xCac21HoByd3bdkceV"
    "V1nuSgYAtHB9pz6oTmS2GnQjF1brPrQ/29FSFLaWoaePGZPs4S3slVDLRZyc+6wlbw+dhJSBqsqXfPatPFJzhXM+kt/X2mXKkrzO"
    "HyrvUMp6lpUJnW4Y7cfB2+yokZ9KS2rrcgSqorZX93ym5HGnC+179slr0tJzAivBuzcPWygNA9YzAQ9X08p6b7nbpRpADRqxggI0"
    "EHaj28vJVWo3tp4TJI75P+erYs8VbYV8YeyU7B2JzZ5oXZrHqw77vplLubB3JqVt6YRArnRLV+yW0wvLTMSN7lbdddp6/uGeUOXT"
    "DsuhWj7skccZjgMAEGovZVrnTuio5M2ETg86t6RW1Jzw+uI9n/WxZLG0Dv2sKbXp2VZqSD4sN+hebblGfWmqfKW9L9/rMNbWS2Xl"
    "ABwl+G/fky/XLYOEAABaNGVoDkci89QlyFLPtdQ3hacLK7kobHtWqgW7jwIh5ztrOdjyFsp1nOU/b9dotmtVS43YZZSM7Xn/s9V5"
    "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"
    "Ci0vsa7Pfm3ivJXy2rEvrHpgtMs2rrsVi2fc/eJh6yaT5iydO05KpXfvmCktuB1n8O/SWVou+eCuy98uFuqXyB9c2F/u8ZBc9w8A"
    "AAAAAfJ4yVnx0rr8dfZNkjq/MqR9O6dXWTbwsD4GFDkuypSLO8tFa69blIs72zVjV9sVooXlctKxC/sFoPSqZ8U8+7JyeWjV+G4t"
    "14PyZebB33q6MPTnXdSX0q4H0HLRAxpAWNWsf/4vfUb+5elNNaJ1l2tm21ZLxeiR8vo+Eag+S1l8s1R9dl6Z/sjooYtHD32krVQS"
    "L/q8RsQOKZXOLrj33aU/rFr6w5ycHYXK2uxe+xctLnTdVkqW8marXr5XLHpN3dzaNe/9d6f3knKzXJUeNXfVw5PkWrZ7eva8W+lE"
    "wfLuY1fOfm6d9yNtyrvIpwX/nrBslfJG4r6JPt/Ifkjyf1SfAQAAADSVl0tOqRhtWb/Yucp8vGTnNHV5Orase+ruF3s+bLv2LC5R"
    "1nbJFm8vcftKpAKx7XJy/tIS1RNZ176y7JZsqX4tFXzVy6478LnbFyZ8f99T64V3fr6L9VL63eni7XnuV8QtEgVoAGH06FWv3rDY"
    "0xOtuzy24cG9Gx78YrYIu7R7Rg99JSdRtSbxkiHdB4sYltOtq+XPcQ/YWoxzfjsz58WF3hJt8Y7D2QO7iiDeyNdux42bJAoOhSAZ"
    "u7/L+tfeLp4009Zle9zV93b5fkV0368EAAAAoMXwfsl5In/aus3T1u106QXVLmvA4tG5MdwXpldutuXPztMet996O+G3txS8/kGR"
    "x+1LCgtEz+wsEQRfux09YazYX1wiQqTryCldim29vtbPmph3nvzfXXJJunDp9ROfm/WUZc0MayW95IO7zrNuY+tBteY5y0P1PbtR"
    "igI0gDCbdvWPGx58bFjSgeV/uXiO9HjzDRcofZ93HBCix6U3/UlEh/KDUuJ/5kTiDUMGXJIkYoecAr+/b6LbAFJSlXb/Gpd24H1v"
    "323JdvNWTpg+LZgCtOfdWq1Zs9JeCm8il3cpLD4k1BXzzt16ikOFRT52sHJ2njXTz1gasrMNAAAAALHL9ZJT7hpl6fvcuZOofW/3"
    "seMCCqnsK197ul2L5Vw7vdu/NzivlC9mpQu3y98W984YJYLiabdW6775Pti6tieFG1YcHjvF0utr3VPzhOW+22W3CFv3qe/FBGnN"
    "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"
    "f5ZwW5TgAAAAAAAAAAC4BAloAAAAAAAAAIBLkIAGAAAAAMAbNBauNG8pEwCAcR3NMsckmGNWFjWLYVADGgAAAACMK2N1zZLpcqXh"
    "xUUH8mRu3ocJ0ZP0J74rinrz+Z5+9k+J1FVH/OyIeJ3mgnVLc5u0xdjtJRsX1BctX5VfG7+5Ymu0aB/pX5u9861kKVy5fkeD9TXW"
    "Zy30lygLwSkHs+RZ2/J+c6CtT1lOTEZ5eKqyqiAt57u+ek3RtrhG+9UmZlpatKcc3uiQt2JZ7vMUAPfWT0Be8fozL8wLsOvTO4br"
    "oV5z+UTJ/KdOiofqiWymtH27kkL0yKlFbP3ZvbOUcFeXZU4vtb6kbwCs3JKQXawuKK+KO9SzrK1B13/wDzpqt1o9htuecnijt7yV"
    "vg8HNfB21fW8Vtp7Q0bADGgAAAAAMKikZcesKQnNzBmTlJzFKxHvfycy3fzOY717a09FvVJUJxI2Jy9JvE1ZztLcsO0lRRUlRQdT"
    "5Wy93moKP51fWN+3r5IgrlB7bk4s7f1siPmtfSnhSpZByXfYL/d5fbBJcg8clQFWmxlbnJFzdJC3OtStpBxU3+TONMl/tqBRABhH"
    "34Cc99TuCK0l4tWGy8rzdc12ZxBlReBEJe8cEVVSeUUmz3vEUwO4kqLNqEvbp8fJsGpr7A0PrnvNIcQpCeIKawDc0+vCjuhtSuhW"
    "89dKitZ+uY/+g79ltftSJDfb8VmntlKmR/vmslL10DB89Wdqg2eFicGQgAYAAAAAQ3rs8IZgOVGt5CNs/vN9bcbcF9cv9+18csUi"
    "+wnRXsr6oT0weWNSiN4UtnZN2I79lf33r286K2GhIXLr4tem9pMxsYiLS5S6s/UyQoIWPmqq/VKb2V2Wo16XrX6t01MnzQXrLC1Z"
    "1t+xNKf/DgmD5sQBjJxBAvLcvDXBk6Wlsqj3RSrNp3eps55PNn0nHs4acuM22q4OCV2TEtrfKT1NU3WDKdKZKD1o8A+JWhzcVD0C"
    "UVrZimjHgsaPvoxf+6itXS2I1Cvw2qL3yqLmspz0UqnNXa8FaltPPYxXbskol4b8pQlqPSUDBXAS0AAAAABgQC+9E3lvXfX8p761"
    "azvyfKb2lFnJX8hX1Q5FNjJW11RtMoe1VL66e0WheBsl7at+aF/Xd1Jb3LI06ww1m+IM7SP9qnxJXebcZc5hySmhH1T1X6CzrKzY"
    "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"
    "Xe1ICExOSWyw3fNKrbwR/mhUoLKg1gCN3qbeHVGn/6aVh0qHstZb6gxgREwKfsEari8PoTiSkpVWOtdsCJ58peFFLXPtvRoLs6zF"
    "Nywn4XooQTK0tLzW8UXDC+wDrnaE2dX00H8v9W2rjmapNfoXbN2ZFqy9n5AwsVXeKCuiBAcAAAAAwCV6psvZOfJ41JGbdXPs48EC"
    "42YdWmXeoS2Hp+7cFqJeqmyzYKVaXtP6yForQ0xp+3bZz1m2XgqdvXz2zrfiqvSCFekrZw00A673ah3ZNtRz2XWfrfxGsnu2mGxf"
    "MSN6W8nmLQnrY3J7vTzy9PqYBEvLAgkKS81fqv7WsYnxcjPqzL4hdwYwMgaLw3lP7bbml61VlXo1er2ghbOzrdUnlHC9UYnDdT3P"
    "KjHNVJxrezhgYD+aZU5XzrqVrt8SUvR0/boBQu5Aq3XUf2Af4lbUGtD6yvWaHrIsbe969XeMj03UmsNm1y1N0CosBaccVOspbTyY"
    "um6p5eim/GpR4vZ8rl5vlWHr7u7Wv9ubOuUOGZ5XIgywBz3MJktJOHb+GGDnjyHbzr94+ZpgdE2dbLlomJ0/+mw732kXL1316U1p"
    "1L/rXDRC4K/FA7wxnzwH4L5sQyPACQwP+uCQh9FEAMdw3DSAO40SHAAAAAAAAAAAlyABDQAADGZETsJjbE2YNlUAuCX+94RBue3w"
    "gP+nMGr4Y4PbIgENAAAMxteXAYzhTfv+HAHglu6KCBfAgNx2eMAhD6OGAA63xec3AABgMBP8xwkM7gcpy4SZ7IAb8vGJemKFAAbk"
    "tsMDDnkYJQRwuDES0AAAwGD8/HwDyEEbnCl63gPLkwSAm4l6ctWsR7ghNgzJbYcHHPIwOgjgcGd+AgAAYDQTAvz9/f1aWts7Ojq7"
    "ursFBjR/3RNzFv+ouuDdr6tOXG3+RgCMnUmBdwfFPDz/l7+YSa0AGJnbDg845MF1COAwBBLQAADAkHxvu23ihPECI5v6UOS9D0UK"
    "AAAjxG2HBxzyAHgzSnAAAAAAAAAAAFyCBDQAAAAAAAAAwCVIQAMAAAAAAAAAXIIENAAAAAAAAADAJUhAAwAAAAAAAABcggQ0AAAA"
    "AAAAAMAlSEDDyEw/WnL4w7R/39zfw8C5e//9mZqqTTV/eubfN4cJAAAAAAAAgNFGAhrGlbTs31+8795JfuPH9/NwUXr039/TWbz+"
    "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"
    "lpboxfFytt6tJ0GTgHbKJ/syDjUO+NDi+J6Nv97zib7c+N62X9v6nPjXXz+z7f2/Cpxit7f/emjHMxt//czGHe9906cTO99lvnk/"
    "Q93nv3b8m1f/Oaz71n5Z+Sfr798It86685/51+N9nmHnu5gaUp6xfPXdn2pUsf6LsPMBAGNsqu//mqF8kO22/5DcXHY99rlrGd6Z"
    "iVbS0Oqc3zK7jIM6ke21/Wf66RxiChWnOazWlvLY6tys6hFVlrM0N0zLs29O7LdDcMr21Lr0rEoBMIaU7PML8wKU7HPEz7Tpz0nL"
    "jlkKbszNe8dyr8KXIqeLtHz20UmBlobOtMvYqgYJ8rPCxVkh5rVz8veU2rWMZJBXp1FLabZ6mjCjXBrsq3B4DhLQt0zNL+SfHehh"
    "j0+qv5k65bNjtjzRFPmk+IS6cLzi2ynTBM7ovbeP/5+SO1fn/NPulDv/8F7vfBw731Ua39t74u6Uf9qd83/N++Td3pm1xspPZJqc"
    "rrQ2Ksv/RzsHcOLY36ZNFQzfifcsO/9H35ax80fXw08roUb9Wny/3Bl0t/1TWlSpqT5hfczOBwCMqduW/dI/RCvB8dP3u0OWTMi0"
    "LwbdcxPC7p3Z12ILOsRjNRZmWaeqqRPEZoXZPReYnBJaWl7r+KKysmIJC3W2AMWAqx0FDWfqRH//vdVXHWowReq/kb4THPtYhSVv"
    "Tjud7f5VRAEP0XMTwknBL1RtUktqZKxWs8+iVthQCz1XrX7J/gXT1W5KuzlMvnp/94pC8WqVW6wnzBzL9w8UjdVZxr0PB7dkwcqU"
    "s6Xl4gpqrLZei6OVWupVRcQWyZUEuiU33VSt14Duaak8VGoriu2m/AS36HuL03bfsy/j6/4f2iiph3lr4hp/o2QlHp6nNtw574d/"
    "q/hE5km1/HCOfCJwQq+9/c1fv4mIVPftPXdN+/avf1WetXZj57vSlKB7lO9BQTMvVSj/ELZM3Df/eWJm3FpT2WvHG3+yWI16d/9w"
    "zjfKOYAfSoXMmSenBcNmMmmnVX6o/KX3zoGy80fLXw+VfZPw03n2Tcrprh/+9CdNb6gR5odqAzsfADDG1DsNjrNvCIy7vTxOX+z7"
    "lOcKWjg7e2lCvrZsStu3UclN1PU8G/10qqk41/awPD2h3Npzl311i6NZ5vRSkdL1W0KKnq5ftzS3SSR7+eydA1xt3We1jmwbkvDU"
    "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"
    "AjQAAAAAAAkBGgAAAACAhAANAAAAAEBCgAYAAAAAICFAAwAAAACQEKABJgAAAAAKAjQAAAAAAAkBGgAAAACAhAANAAAAAEBCgAYA"
    "AAAAICFAAwAAAACQEKABAAAAAEgI0AAAAAAAJARoAAAAAAASAjQAAAAAAAkBGgAAAACAhAANAAAAAEBCgAYAAAAAICFAAwAAAACQ"
    "EKABAAAAAEgI0AAAAAAAJARoAAAAAAASAjQAAAAAAAkBGgAAAACAhAANAAAAAEBCgAYAAAAAICFAAwAAAACQEKABAAAAAEgI0AAA"
    "AAAAJARoAAAAAAASAjQAAAAAAAkBGgAAAACAhAANAAAAAEBCgAYAAAAAICFAAwAAAACQEKABAAAAAEgI0AAAAAAAJARoAAAAAAAS"
    "AjQAAAAAAAkBGgAAAACAhAANAAAAAEBCgAYAAAAAICFAAwAAAACQEKABAAAAAEgI0AAAAAAAJARoAAAAAAASAjQAAAAAAAkBGgAA"
    "AACAhAANAAAAAEBCgAYAAAAAICFAAwAAAACQEKABAAAAAEgI0AAAAAAAJARoAAAAAAASAjQAAAAAAAkBGgAAAACAhAANAAAAAEBC"
    "gAYAAAAAICFAAwAAAACQEKABAAAAAEgI0AAAAAAAJARoAAAAAAASAjQAAAAAAAkBGgAAAACAhAANAAAAAEBCgAYAAAAAICFAAwAA"
    "AACQEKABAAAAAEgI0AAAAAAAJARoAAAAAAASAjQAAAAAAAkBGgAAAACAhAANAAAAAEBCgAYAAAAAICFAAwAAAACQEKABAAAAAEgI"
    "0AAAAAAAJARoAAAAAAASAjQAAAAAAAkBGgAAAACAhAANAAAAAEBCgAYAAAAAICFAAwAAAACQEKABAAAAAEgI0AAAAAAAJARoAAAA"
    "AAASAjQAAAAAAAkBGgAAAACAhAANAAAAAEBCgAYAAAAAICFAAwAAAACQEKABAAAAAEgI0AAAAAAAJARoAAAAAAASAjQAAAAAAAkB"
    "GgAAAACAhAANAAAAAEBCgAYAAAAAICFAAwAAAACQEKABAAAAAEgI0AAAAAAAJARoAAAAAAASAjQAAAAAAAkBGgAAAACAhAANAAAA"
    "AEBCgAYAAAAAICFAAwAAAACQEKABAAAAAEgI0AAAAAAAJARoAAAAAAASAjQAAAAAAAkBGgAAAACAhAANAAAAAEBCgAYAAAAAICFA"
    "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"
)


def _load_activity_template_workbook(template_path=None):
    """تحميل قالب Excel من ملف محلي إن وُجد، وإلا من النسخة المدمجة داخل app.py."""
    if template_path and os.path.isfile(template_path):
        return load_workbook(template_path)
    try:
        embedded_bytes = base64.b64decode("".join(_ACTIVITY_TEMPLATE_B64.split()))
        return load_workbook(io.BytesIO(embedded_bytes))
    except Exception as exc:
        raise RuntimeError(
            "تعذر تحميل قالب تقرير النشاط المدمج داخل التطبيق. أعيدي تنزيل نسخة app.py الكاملة."
        ) from exc

def _build_full_activity_excel_bytes(activity_df, collector_col, created_col, notes_col,
                                     selected_bank, template_path=None):
    # استخدم القالب الموجود بجانب التطبيق إن وُجد، وإلا استخدم النسخة المدمجة داخل الكود.
    report_wb = _load_activity_template_workbook(template_path)

    processed_ws = report_wb['Processed_Data']
    template_headers = [cell.value for cell in processed_ws[1] if cell.value is not None]
    processed_df = _build_processed_activity_export(
        activity_df, template_headers, collector_col, created_col, notes_col
    )

    report_frames = _build_reference_activity_frames(processed_df)

    # Keep the reference KPI index only when the collector's key productivity
    # figures still match the template. Otherwise leave the unknown KPI formula blank.
    if 'Productivity_Report' in report_wb.sheetnames:
        old_kpi = {}
        old_ws = report_wb['Productivity_Report']
        for row in old_ws.iter_rows(min_row=2, values_only=True):
            if row and row[0] is not None:
                old_kpi[str(row[0])] = {
                    'calls': row[1],
                    'success': row[2],
                    'avg_gap': row[4],
                    'score': row[6],
                    'class': row[7],
                }
        prod = report_frames['Productivity_Report']
        for idx, row in prod.iterrows():
            previous = old_kpi.get(str(row['Collector']))
            if previous:
                same_calls = previous['calls'] == row['إجمالي_المكالمات']
                same_success = previous['success'] == row['عدد_الناجحة']
                try:
                    same_avg = abs(float(previous['avg_gap']) - float(row['متوسط_فرق_التوقيت_بالدقائق'])) < 0.01
                except (TypeError, ValueError):
                    same_avg = False
                if same_calls and same_success and same_avg:
                    prod.at[idx, 'مؤشر_الأداء_%'] = previous['score']
                    prod.at[idx, 'تصنيف_KPI'] = previous['class']

    _write_df_to_template_sheet(processed_ws, processed_df)
    for sheet_name, frame in report_frames.items():
        if sheet_name in report_wb.sheetnames:
            _write_df_to_template_sheet(report_wb[sheet_name], frame)

    # Keep the dashboard artwork from the supplied template; all detailed
    # data tabs above are rebuilt from the newly uploaded activity file.
    report_buffer = io.BytesIO()
    report_wb.save(report_buffer)
    report_buffer.seek(0)
    return report_buffer.getvalue(), processed_df, report_frames


def _write_pivot_sheet(ws, p, title, is_real_pivot=False):
    """كتابة شيت الملخص؛ عندما يكون is_real_pivot=True نضع عنوان التقرير فقط في الصفين 1-2 ونترك منطقة الـ Pivot بدون دمج خلايا متعارض"""
    notes = p['notes']
    n = len(notes)
    last_col = 2 + 2 * n + 2
    # فتح الملخص من أعلى الشيت وبمقياس مناسب حتى تظهر مجموعات الـ Pivot كاملة.
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.zoomScale = 80
    ws.sheet_view.zoomScaleNormal = 80
    ws.sheet_view.showGridLines = False
    ws.sheet_view.topLeftCell = 'A1'
    if ws.sheet_view.selection:
        ws.sheet_view.selection[0].activeCell = 'A1'
        ws.sheet_view.selection[0].sqref = 'A1'
    ws.row_dimensions[1].height = 27
    ws.row_dimensions[2].height = 10

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

    if not is_real_pivot:
        msg = ws.cell(
            row=4, column=1,
            value='لا توجد بيانات كافية أو أن أعمدة الـ Pivot المطلوبة غير متاحة.'
        )
        msg.font = Font(name='Calibri', bold=True, color='244062', size=12)
        msg.alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(start_row=4, start_column=1, end_row=5, end_column=last_col)
        ws.column_dimensions['A'].width = 35.7
        return

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
                body(r, 3 + 2 * i, row['by_note'][note][0], COUNT_FMT)
                body(r, 4 + 2 * i, row['by_note'][note][1], SUM_FMT)
            body(r, tc, row['cnt'], COUNT_FMT)
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
        ws.cell(row=r, column=3 + 2 * i, value=g['by_note'][note][0]).number_format = COUNT_FMT
        x = ws.cell(row=r, column=4 + 2 * i, value=g['by_note'][note][1])
        x.number_format = SUM_FMT
    ws.cell(row=r, column=tc, value=g['cnt']).number_format = COUNT_FMT
    x = ws.cell(row=r, column=tc + 1, value=g['amt'])
    x.number_format = SUM_FMT

    # أعمدة مضغوطة لتظهر كل مجموعات الملاحظات والإجماليات على شاشة واحدة.
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 32
    for c in range(3, 3 + 2 * n):
        ws.column_dimensions[get_column_letter(c)].width = 15.5
    ws.column_dimensions[get_column_letter(tc)].width = 15.5
    ws.column_dimensions[get_column_letter(tc + 1)].width = 16


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
        f'name="PivotTable{idx}" cacheId="{idx}" dataOnRows="0" applyNumberFormats="1" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" '
        f'applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" grandTotalCaption="الاجمالي" '
        f'missingCaption="0" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" '
        f'mergeItem="1" createdVersion="8" indent="0" compact="0" compactData="0" multipleFieldFilters="0" '
        f'rowHeaderCaption="المحصل" colHeaderCaption="الفرع">'
        f'<location ref="{ref}" firstHeaderRow="1" firstDataRow="3" firstDataCol="2"/>'
        f'<pivotFields count="{len(cols)}">{"".join(pf)}</pivotFields>'
        f'<rowFields count="2"><field x="{team_i}"/><field x="{coll_i}"/></rowFields>'
        f'<rowItems count="{len(row_items)}">{"".join(row_items)}</rowItems>'
        f'<colFields count="2"><field x="{note_i}"/><field x="-2"/></colFields>'
        f'<colItems count="{len(col_items)}">{"".join(col_items)}</colItems>'
        f'<dataFields count="2">'
        f'<dataField name="عدد الحسابات" fld="{acc_i}" subtotal="count" numFmtId="3"/>'
        f'<dataField name="متبقي المديونية" fld="{amt_i}" subtotal="sum" numFmtId="3"/>'
        f'</dataFields>'
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
        root_match = re.search(r'<workbook\b[^>]*>', wb_xml)
        if root_match and not re.search(r'\bxmlns:r=', root_match.group(0)):
            root_tag = root_match.group(0)
            updated_root = root_tag[:-1] + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            wb_xml = wb_xml.replace(root_tag, updated_root, 1)
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
    if bank == 'SNB':
        active_title = 'تقرير الوعود القائمة البنك الاهلي السعودي SNB (NPL&60Dpd)'
        broken_title = 'تقرير الوعود المكسورة البنك الاهلي SNB (NPL&60Dpd)'
    else:
        active_title = f'تقرير الوعود القائمة {BANKS[bank]} {bank}'
        broken_title = f'تقرير الوعود المكسورة {BANKS[bank]} {bank}'
    _write_pivot_sheet(ws, p_active, active_title, is_real_pivot=can_pv_active)

    ws = wb.create_sheet('تقرير الوعود المكسورة')
    _write_df_sheet(ws, broken_df)

    can_pv_broken = _can_pivot(broken_df)
    p_broken = build_pivot(broken_df, order='desc')
    ws = wb.create_sheet('ملخص تقرير الوعود المكسورة')
    _write_pivot_sheet(ws, p_broken, broken_title, is_real_pivot=can_pv_broken)

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


def build_single_promise_report_xlsx(df, report_kind="active", bank="STC"):
    """Build one standalone promises workbook (active OR broken) with its summary."""
    is_active = report_kind == "active"
    data_sheet = 'تقرير الوعود القائمة' if is_active else 'تقرير الوعود المكسورة'
    summary_sheet = 'ملخص تقرير الوعود القائمة' if is_active else 'ملخص تقرير الوعود المكسورة'
    order = 'first' if is_active else 'desc'

    if bank == 'SNB':
        title = (
            'تقرير الوعود القائمة البنك الاهلي السعودي SNB (NPL&60Dpd)'
            if is_active
            else 'تقرير الوعود المكسورة البنك الاهلي SNB (NPL&60Dpd)'
        )
    else:
        title = (
            f'تقرير الوعود القائمة {BANKS[bank]} {bank}'
            if is_active
            else f'تقرير الوعود المكسورة {BANKS[bank]} {bank}'
        )

    wb = Workbook()
    ws = wb.active
    ws.title = data_sheet
    _write_df_sheet(ws, df)

    can_pv = _can_pivot(df)
    pivot_df = build_pivot(df, order=order)
    ws = wb.create_sheet(summary_sheet)
    _write_pivot_sheet(ws, pivot_df, title, is_real_pivot=can_pv)

    buf = io.BytesIO()
    wb.save(buf)

    if can_pv:
        return inject_pivot_tables(
            buf.getvalue(),
            [{
                'summary_sheet': summary_sheet,
                'data_sheet': data_sheet,
                'data_df': df,
                'pivot': pivot_df,
                'order': order,
            }]
        )

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
            for page in ('Home', 'Overview', 'Promises', 'Neglect', 'Activity')}
        for b in BANKS
    }

# إضافة صفحة Home للجلسات القديمة دون التأثير على ملفات التقارير المرفوعة.
for _bank in BANKS:
    st.session_state.page_files[_bank].setdefault('Home', None)

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

# نسخة مستقلة لخانة رفع كل تقرير، عشان زر الإزالة يفضّي الـ widget فعلياً.
if 'report_uploader_versions' not in st.session_state:
    st.session_state.report_uploader_versions = {
        b: {page: 0 for page in ('Promises', 'Neglect', 'Activity')}
        for b in BANKS
    }
else:
    for b in BANKS:
        st.session_state.report_uploader_versions.setdefault(
            b, {page: 0 for page in ('Promises', 'Neglect', 'Activity')}
        )
        for page in ('Promises', 'Neglect', 'Activity'):
            st.session_state.report_uploader_versions[b].setdefault(page, 0)

def clear_report_page_file(bank, page):
    st.session_state.page_files[bank][page] = None
    st.session_state.report_uploader_versions[bank][page] += 1

# Independent Activity-page uploads: one model and one activity file per bank.
# Files are kept in Streamlit session state across bank/page navigation.
if 'activity_page_uploads' not in st.session_state:
    st.session_state.activity_page_uploads = {
        bank: {'model': None, 'activity': None} for bank in BANKS
    }
else:
    for bank in BANKS:
        st.session_state.activity_page_uploads.setdefault(
            bank, {'model': None, 'activity': None}
        )
        st.session_state.activity_page_uploads[bank].setdefault('model', None)
        st.session_state.activity_page_uploads[bank].setdefault('activity', None)


def save_activity_page_upload(bank, kind, widget_key):
    """Keep both uploads per bank, and open the new SNB report screen automatically."""
    st.session_state.activity_page_uploads[bank][kind] = st.session_state.get(
        widget_key
    )
    if bank == 'SNB' and kind == "activity":
        # A different upload should not silently keep its predecessor's shift.
        st.session_state.pop("snb_activity_daily_shift", None)
        st.session_state.pop("snb_activity_exclude_after_shift", None)
        st.session_state.pop("snb_activity_exclude_last_hour", None)
    if bank == 'SNB':
        uploads = st.session_state.activity_page_uploads[bank]
        st.session_state.snb_activity_screen = (
            'report'
            if uploads.get('model') is not None
            and uploads.get('activity') is not None
            else 'upload'
        )



# Upload keys are versioned so removing a file clears the native uploader too.
if "activity_upload_versions" not in st.session_state:
    st.session_state.activity_upload_versions = {
        bank: {"model": 0, "activity": 0} for bank in BANKS
    }
else:
    for bank in BANKS:
        versions = st.session_state.activity_upload_versions.setdefault(bank, {})
        versions.setdefault("model", 0)
        versions.setdefault("activity", 0)


def clear_activity_page_upload(bank, kind):
    st.session_state.activity_page_uploads[bank][kind] = None
    st.session_state.activity_upload_versions[bank][kind] += 1
    if bank == "SNB":
        st.session_state.snb_activity_screen = "upload"

def _snb_header_key(value):
    """Normalize labels for robust English/Arabic column lookup."""
    return re.sub(
        r"[\s_\-–—]+", "", str(value or "")
    ).casefold().strip()


def _snb_find_column(frame, *names):
    available = {_snb_header_key(col): col for col in frame.columns}
    for name in names:
        hit = available.get(_snb_header_key(name))
        if hit is not None:
            return hit
    return None


def _snb_normalize_note(value):
    """Normalize Notes (= إفادة) for exact matching to the uploaded model."""
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().casefold()



# ---------------------------------------------------------------------------
# Persistent user-approved SNB note evaluations.
# Only an ENCRYPTED reference file is editable from Streamlit. The app.py source,
# passwords, uploaded portfolios and individual debtor rows are never saved.
# ---------------------------------------------------------------------------
_SNB_MODEL_OVERRIDES_PATH = "data/snb_activity_model_overrides.enc"


def _snb_secret(name, default=""):
    try:
        return str(st.secrets.get(name, default) or "").strip()
    except (FileNotFoundError, KeyError, AttributeError):
        return default


def _snb_model_git_settings():
    return {
        "token": _snb_secret("GITHUB_TOKEN"),
        "repo": _snb_secret("GITHUB_REPO", "Ranaamohamed22/promises-dashboard"),
        "branch": _snb_secret("GITHUB_BRANCH", "main"),
        "admin_password": _snb_secret("MODEL_ADMIN_PASSWORD"),
    }


def _snb_github_model_request(settings, method="GET", payload=None):
    """Read/write exactly one reference file via GitHub Contents API."""
    from urllib.parse import quote
    repo_name = settings["repo"]
    repo_parts = repo_name.split("/")
    if len(repo_parts) != 2 or any(not part for part in repo_parts):
        raise ValueError("إعداد GITHUB_REPO يجب أن يكون owner/repository.")
    path = quote(_SNB_MODEL_OVERRIDES_PATH, safe="/")
    url = (
        "https://api.github.com/repos/"
        + quote(repo_parts[0], safe="")
        + "/" + quote(repo_parts[1], safe="")
        + "/contents/" + path
    )
    if method == "GET":
        url += "?ref=" + quote(settings["branch"], safe="")
    data = None if payload is None else json.dumps(
        payload, ensure_ascii=False
    ).encode("utf-8")
    request = Request(
        url, data=data, method=method,
        headers={
            "Authorization": "Bearer " + settings["token"],
            "Accept": "application/vnd.github+json",
            "Content-Type": "application/json",
            "User-Agent": "EJADA-Dashboard-Model-Manager",
        },
    )
    try:
        with urlopen(request, timeout=12) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        if exc.code == 404 and method == "GET":
            return None
        if exc.code in (401, 403):
            raise RuntimeError(
                "GitHub رفض صلاحية الحفظ أو انتهت صلاحية الرمز؛ راجعي Streamlit Secrets."
            ) from None
        if exc.code == 409:
            raise RuntimeError(
                "تم تعديل الموديل في جلسة أخرى؛ أعيدي تحميل الصفحة ثم احفظي مرة أخرى."
            ) from None
        raise RuntimeError("تعذّر الاتصال بملف الموديل على GitHub (HTTP %s)." % exc.code) from None
    except (URLError, TimeoutError) as exc:
        raise RuntimeError("تعذّر الاتصال بخدمة حفظ موديل النشاط على GitHub.") from exc



def _snb_get_fernet():
    """Require private encryption key; never place the key in a public repository."""
    secret = _snb_secret("MODEL_ENCRYPTION_KEY")
    if not secret:
        raise ValueError(
            "لازم إضافة MODEL_ENCRYPTION_KEY إلى Streamlit Secrets لتشغيل الموديل المشفر."
        )
    try:
        return Fernet(secret.encode("ascii"))
    except (ValueError, UnicodeError) as exc:
        raise ValueError(
            "مفتاح MODEL_ENCRYPTION_KEY غير صالح؛ استخدمي السطر الموجود في ملف الإعداد."
        ) from exc


def _snb_secret_model_evaluations():
    """Load the encrypted 777-row baseline model from Streamlit private Secrets.

    No customer Notes appear in public code or GitHub history.
    """
    ciphertext = _snb_secret("SNB_MODEL_ENCRYPTED")
    if not ciphertext:
        raise ValueError(
            "الموديل الثابت لم يتم تفعيله؛ أضيفي SNB_MODEL_ENCRYPTED إلى Streamlit Secrets."
        )
    try:
        compressed = _snb_get_fernet().decrypt(ciphertext.encode("ascii"))
        source = json.loads(gzip.decompress(compressed).decode("utf-8"))
    except (InvalidToken, OSError, ValueError, UnicodeError, json.JSONDecodeError) as exc:
        raise ValueError(
            "تعذر فك تشفير موديل النشاط الثابت؛ راجعي سطري SNB_MODEL_ENCRYPTED "
            "وMODEL_ENCRYPTION_KEY في Secrets."
        ) from exc
    entries = source.get("entries", [])
    if not isinstance(entries, list):
        raise ValueError("ملف موديل النشاط الثابت غير صالح.")
    mapping, conflicts, invalid = {}, set(), 0
    for entry in entries:
        if not isinstance(entry, dict):
            invalid += 1
            continue
        key = _snb_normalize_note(entry.get("notes"))
        rating = entry.get("evaluation")
        if not key or type(rating) is not int or rating not in (0, 1):
            invalid += 1
            continue
        if key in mapping and mapping[key] != rating:
            conflicts.add(key)
        else:
            mapping[key] = rating
    for key in conflicts:
        mapping.pop(key, None)
    return mapping, len(conflicts), invalid


def _snb_load_saved_model_notes(settings):
    """Return normalized mapping, original labels, and current file SHA."""
    if not settings["token"]:
        return {}, {}, None
    response = _snb_github_model_request(settings)
    if response is None:
        return {}, {}, None
    if not isinstance(response, dict) or response.get("type") != "file":
        raise ValueError("ملف موديل النشاط على GitHub غير صالح.")
    encoded = response.get("content", "")
    if not encoded:
        raise ValueError("ملف الموديل المحفوظ لا يحتوي على بيانات قابلة للقراءة.")
    try:
        encrypted = base64.b64decode(encoded).strip()
        decrypted = _snb_get_fernet().decrypt(encrypted)
        document = json.loads(decrypted.decode("utf-8"))
    except (InvalidToken, ValueError, UnicodeError, json.JSONDecodeError) as exc:
        raise ValueError(
            "تعذّر فتح موديل الإضافات المشفّر. راجعي MODEL_ENCRYPTION_KEY."
        ) from exc
    if not isinstance(document, dict) or not isinstance(document.get("entries", []), list):
        raise ValueError("صيغة ملف الموديل المحفوظ غير صحيحة.")
    ratings, labels = {}, {}
    for entry in document.get("entries", []):
        if not isinstance(entry, dict):
            continue
        note = str(entry.get("notes", "")).strip()
        key = _snb_normalize_note(note)
        value = entry.get("evaluation")
        if key and type(value) is int and value in (0, 1):
            if key in ratings and ratings[key] != value:
                raise ValueError("موديل النشاط المحفوظ يحتوي على تقييمين متعارضين.")
            ratings[key], labels[key] = value, note
    return ratings, labels, response.get("sha")


def _snb_save_model_notes(settings, additions):
    """Merge entries using a fresh SHA; do not overwrite concurrent updates."""
    if not settings["token"]:
        raise RuntimeError("لم يتم إعداد GITHUB_TOKEN داخل Streamlit Secrets.")
    current, labels, sha = _snb_load_saved_model_notes(settings)
    for note, rating in additions:
        key = _snb_normalize_note(note)
        if not key or rating not in (0, 1):
            raise ValueError("لا يمكن حفظ إفادة فارغة أو تقييم غير صالح.")
        if _snb_special_review_mask(pd.Series([note])).iloc[0]:
            raise ValueError(
                "الإفادات التي تحتوي على / أو // أو «معالج» أو «تم» "
                "تحتاج مراجعة لكل صف، ولا تُحفظ كتقييم عام."
            )
        current[key] = int(rating)
        labels[key] = str(note).strip()
    entries = [
        {"notes": labels[key], "evaluation": current[key]}
        for key in sorted(current)
    ]
    raw = json.dumps(
        {"schema_version": 1, "entries": entries},
        ensure_ascii=False, indent=2
    ).encode("utf-8")
    payload = {
        "message": "Update approved EJADA SNB activity note evaluations",
        "branch": settings["branch"],
        "content": base64.b64encode(
            _snb_get_fernet().encrypt(raw)
        ).decode("ascii"),
    }
    if sha:
        payload["sha"] = sha
    _snb_github_model_request(settings, method="PUT", payload=payload)
    return len(additions)


def _snb_render_model_management(notes, base_mapping, manual, row_decisions):
    """Visible INSIDE Activity only after all required note ratings are approved."""
    settings = _snb_model_git_settings()
    st.markdown("### إدارة موديل النشاط")
    st.caption(
        "تظهر هنا الإفادات الجديدة بعد اعتماد تقييماتها. "
        "الحفظ يضيف الإفادات العادية فقط إلى ملف مرجعي دائم على GitHub؛ "
        "الإفادات الخاصة (/ أو // أو معالج أو تم) تظل للمراجعة في كل تقرير."
    )
    previous_notice = st.session_state.pop("snb_model_manager_saved", None)
    if previous_notice:
        st.success(previous_notice)
    if not settings["token"]:
        st.info("فعّلي GITHUB_TOKEN من Streamlit Secrets لتشغيل الحفظ الدائم.")
        return
    if not settings["admin_password"]:
        st.warning("أضيفي MODEL_ADMIN_PASSWORD إلى Streamlit Secrets قبل السماح بتعديل الموديل.")
        return
    if not _snb_secret("MODEL_ENCRYPTION_KEY"):
        st.warning("أضيفي MODEL_ENCRYPTION_KEY إلى Streamlit Secrets لتشفير الإفادات قبل حفظها.")
        return

    if not st.session_state.get("snb_model_manager_is_admin", False):
        with st.form("snb_model_manager_admin_login"):
            password = st.text_input(
                "كلمة مرور إدارة موديل النشاط", type="password"
            )
            login = st.form_submit_button("فتح إدارة موديل النشاط")
        if login:
            if hmac.compare_digest(password, settings["admin_password"]):
                st.session_state.snb_model_manager_is_admin = True
                st.rerun()
            else:
                st.error("كلمة مرور الإدارة غير صحيحة.")
        return

    if st.button("قفل إدارة موديل النشاط", key="snb_model_manager_logout"):
        st.session_state.snb_model_manager_is_admin = False
        st.rerun()

    try:
        persisted, _, _ = _snb_load_saved_model_notes(settings)
    except (RuntimeError, ValueError, json.JSONDecodeError) as exc:
        st.error(str(exc))
        return

    display_labels = {}
    candidates = {}
    conflicted = set()
    normalized = notes.map(_snb_normalize_note)
    for index, (note, key) in enumerate(zip(notes, normalized)):
        if not key or _snb_special_review_mask(pd.Series([note])).iloc[0]:
            continue  # Special notes require per-row review, never a generic rating.
        # The original model's ratings are already known. Only approved manual
        # ratings are eligible to be added to the persistent additions file.
        if key not in manual:
            continue
        value = manual[key]
        if value not in (0, 1):
            continue
        if key in candidates and candidates[key] != value:
            conflicted.add(key)
        candidates[key] = value
        display_labels.setdefault(key, str(note).strip())
    for key in conflicted:
        candidates.pop(key, None)
    rows = [
        {
            "إضافة للموديل": False,
            "الإفادة المعتمدة": display_labels[key],
            "التقييم": "ناجحة" if rating == 1 else "غير ناجحة",
            "الحالة": (
                "موجودة بنفس التقييم" if persisted.get(key) == rating
                else "تحديث تقييم محفوظ" if key in persisted
                else "تحديث الموديل الأساسي" if key in base_mapping
                else "جديدة"
            ),
        }
        for key, rating in sorted(candidates.items())
        if persisted.get(key) != rating
    ]
    st.caption(f"الإفادات المعتمدة المتاحة للإضافة: {len(rows)}")
    if not rows:
        st.info("لا توجد إفادات جديدة معتمدة لإضافتها حاليًا.")
    with st.form("snb_model_manager_approved_form"):
        if rows:
            edited = st.data_editor(
                pd.DataFrame(rows),
                hide_index=True,
                use_container_width=True,
                disabled=["الإفادة المعتمدة", "التقييم", "الحالة"],
                column_config={
                    "إضافة للموديل": st.column_config.CheckboxColumn(
                        "إضافة للموديل", default=False
                    )
                },
                key="snb_model_manager_editor",
            )
        else:
            edited = pd.DataFrame()
        add_note = st.text_input("إضافة إفادة جديدة يدويًا (اختياري)")
        add_rating = st.selectbox(
            "تقييم الإفادة الجديدة",
            ["ناجحة", "غير ناجحة"],
            key="snb_model_manager_new_rating",
        )
        overwrite = st.checkbox(
            "أوافق على استبدال التقييم القديم لو الإفادة موجودة",
            value=False,
        )
        save = st.form_submit_button("💾 حفظ الإفادات المحددة في الموديل")
    if not save:
        return
    additions = []
    if not edited.empty:
        for _, row in edited.iterrows():
            if bool(row["إضافة للموديل"]):
                additions.append((
                    row["الإفادة المعتمدة"],
                    1 if row["التقييم"] == "ناجحة" else 0
                ))
    if add_note.strip():
        additions.append((add_note.strip(), 1 if add_rating == "ناجحة" else 0))
    if not additions:
        st.warning("حددي إفادة معتمدة أو اكتبي إفادة جديدة للحفظ.")
        return
    batch = {}
    batch_labels = {}
    for note, rating in additions:
        key = _snb_normalize_note(note)
        if not key:
            st.error("لا يمكن حفظ إفادة فارغة.")
            return
        if _snb_special_review_mask(pd.Series([note])).iloc[0]:
            st.error("هذه الإفادة من الحالات الخاصة التي تتطلب مراجعة كل صف؛ لن تُضاف كتقييم عام.")
            return
        if key in batch and batch[key] != rating:
            st.error("نفس الإفادة موجودة بتقييمين مختلفين.")
            return
        existing = persisted.get(key, base_mapping.get(key))
        if existing is not None and existing != rating and not overwrite:
            st.warning("فيه إفادة بتقييم مختلف عن الموديل؛ فعّلي الموافقة على الاستبدال أولًا.")
            return
        batch[key] = rating
        batch_labels[key] = note.strip()
    try:
        count = _snb_save_model_notes(
            settings,
            [(batch_labels[key], rating) for key, rating in batch.items()]
        )
    except (RuntimeError, ValueError) as exc:
        st.error(f"تعذّر الحفظ على GitHub: {exc}")
        return
    st.session_state.snb_model_manager_saved = (
        f"اتحفظت {count} إفادة في موديل النشاط الدائم على GitHub."
    )
    st.rerun()


def _snb_model_evaluations(uploaded_model):
    """Map model Notes -> 0/1; conflicts and unrecognized values are not guessed."""
    try:
        model = pd.read_excel(io.BytesIO(uploaded_model.getvalue()), dtype=str)
    except Exception as exc:
        raise ValueError(f"تعذر قراءة موديل النشاط Excel: {exc}") from exc
    model.columns = model.columns.astype(str).str.strip()
    notes_col = _snb_find_column(
        model, "Notes", "Note", "الإفادة", "افادة"
    )
    evaluation_col = _snb_find_column(
        model, "التقيم", "التقييم", "تقييم", "Evaluation", "Rating", "Score"
    )
    missing = [
        label for label, col in [
            ("Notes", notes_col), ("التقيم", evaluation_col)
        ] if col is None
    ]
    if missing:
        raise ValueError(
            "موديل النشاط لازم يحتوي على العمودين Notes والتقيم. "
            "الناقص: " + "، ".join(missing)
        )

    model["__key"] = model[notes_col].map(_snb_normalize_note)
    model["__value"] = pd.to_numeric(
        model[evaluation_col].astype(str).str.strip(),
        errors="coerce"
    )
    valid = model[
        model["__key"].ne("") & model["__value"].isin([0, 1])
    ].copy()

    # If identical Notes have both 0 and 1 in the model, leave them unclassified.
    conflicts = valid.groupby("__key")["__value"].nunique()
    conflict_keys = set(conflicts[conflicts > 1].index)
    valid = valid[
        ~valid["__key"].isin(conflict_keys)
    ].drop_duplicates(subset=["__key"], keep="first")
    mapping = dict(
        zip(valid["__key"], valid["__value"].astype(int))
    )
    rejected = int(
        (model["__key"].eq("") | ~model["__value"].isin([0, 1])).sum()
    )
    return mapping, len(conflict_keys), rejected


def _prepare_snb_activity_file(uploaded_file):
    """Deduplicate Debitor+Notes+Collector, create Time, and combine CC/PF."""
    file_name = uploaded_file.name
    file_bytes = uploaded_file.getvalue()
    if file_name.lower().endswith(".csv"):
        try:
            frame = pd.read_csv(
                io.BytesIO(file_bytes), dtype=str, encoding="utf-8-sig",
                low_memory=False
            )
        except UnicodeDecodeError:
            frame = pd.read_csv(
                io.BytesIO(file_bytes), dtype=str, encoding="cp1256",
                low_memory=False
            )
    else:
        frame = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
    frame.columns = frame.columns.astype(str).str.strip()

    debitor = _snb_find_column(frame, "Debitor", "Debtor")
    notes = _snb_find_column(
        frame, "Notes", "Note", "الإفادة", "افادة"
    )
    collector = _snb_find_column(
        frame, "Collector", "Collecter"
    )
    created = _snb_find_column(
        frame, "Created on", "Created_on", "CreatedOn"
    )
    # Merge CC and PF in the EXISTING product/finance column, not a new one.
    # The actual SNB activity column is "Claim/نوع المتنج-التمويل".
    # Exact matching has priority; keep the original header unchanged.
    product = _snb_find_column(
        frame,
        "Claim/نوع المتنج-التمويل",
        "Claim/نوع المنتج-التمويل",
        "نوع المتنج-التمويل/Claim",
        "نوع المنتج-التمويل/Claim",
        "نوع المنتج - التمويل/Claim",
        "نوع المنتج-التمويل / Claim",
        "Claim/نوع المنتج-التمويل",
        "نوع المنتج-التمويل",
        "نوع المنتج - التمويل",
        "نوع المتنج-التمويل",
        "نوع المنتج",
        "Product Type - Finance",
        "Product Type - Financing",
        "Product Type/Claim",
        "Product Type", "ProductType"
    )
    if product is None:
        product_candidates = []
        for col in frame.columns:
            normalized = _snb_header_key(col)
            is_finance_product = (
                ("منتج" in normalized or "متنج" in normalized)
                and "تمويل" in normalized
            )
            is_english_product = (
                "product" in normalized and "type" in normalized
            )
            if not (is_finance_product or is_english_product):
                continue
            # Explicit Claim/product/finance names outrank generic fields.
            score = (
                4 * ("claim" in normalized)
                + 2 * is_finance_product
                + is_english_product
            )
            product_candidates.append((score, col))
        if product_candidates:
            product_candidates.sort(key=lambda item: item[0], reverse=True)
            product = product_candidates[0][1]
    missing = [
        label for label, col in [
            ("Debitor", debitor), ("Notes (الإفادة)", notes),
            ("Collector", collector), ("Created on", created)
        ] if col is None
    ]
    if missing:
        raise ValueError(
            "الأعمدة المطلوبة في ملف النشاط غير موجودة: "
            + "، ".join(missing)
        )

    initial_rows = len(frame)
    key_columns = [
        "__ejada_key_debitor", "__ejada_key_notes",
        "__ejada_key_collector"
    ]
    for key, name in zip(key_columns, [debitor, notes, collector]):
        frame[key] = frame[name].fillna("").astype(str).str.strip()

    # Updated rule: any CC OR PF product becomes CC+PF in the SAME column,
    # even when only one of the two codes is present. Preserve other values.
    combined_products = 0
    if product is not None:
        product_text = frame[product].fillna("").astype(str)
        has_cc_or_pf = product_text.str.contains(
            r"(?<![A-Za-z0-9])(?:CC|PF)(?![A-Za-z0-9])",
            case=False, regex=True
        )
        needs_change = (
            has_cc_or_pf
            & product_text.str.strip().str.upper().ne("CC+PF")
        )
        combined_products = int(needs_change.sum())
        frame.loc[has_cc_or_pf, product] = "CC+PF"

    frame = (
        frame.loc[~frame.duplicated(subset=key_columns, keep="first")]
        .drop(columns=key_columns)
        .copy()
    )
    duplicates = initial_rows - len(frame)

    parsed = pd.to_datetime(frame[created], errors="coerce")
    # Pandas 2+ can parse heterogeneous export timestamp formats with "mixed".
    # Only apply that fallback when it actually recovers more timestamps.
    if parsed.isna().any():
        try:
            mixed = pd.to_datetime(
                frame[created], errors="coerce", format="mixed"
            )
            if mixed.notna().sum() > parsed.notna().sum():
                parsed = mixed
        except (ValueError, TypeError):
            pass

    if len(frame) and parsed.notna().sum() == 0:
        raise ValueError("تعذر قراءة وقت Created on من ملف النشاط.")

    for name in ("Time", "مده المكالمه", "مدة المكالمة"):
        if name in frame.columns:
            frame = frame.drop(columns=[name])
    frame.insert(
        frame.columns.get_loc(created) + 1,
        "Time",
        parsed.dt.strftime("%I:%M:%S %p").fillna("")
    )
    frame["__ejada_parsed_created"] = parsed
    return frame, {
        "original": initial_rows, "duplicates": duplicates,
        "created": created, "notes": notes,
        "collector": collector, "product": product,
        "combined_products": combined_products,
        "bad_created": int(parsed.isna().sum())
    }


def _infer_snb_daily_period(frame):
    """Infer shift suggestion: 02:00–12:00 / 12:30–16:00 / full day."""
    times = frame["__ejada_parsed_created"].dropna()
    if times.empty:
        return "full", True
    latest_day = times.max().normalize()
    # Infer from the most recent calendar date in this upload.
    current_day = times[times.dt.normalize() == latest_day]
    minute_of_day = current_day.dt.hour * 60 + current_day.dt.minute
    first = ((minute_of_day >= 120) & (minute_of_day < 720)).any()
    second = ((minute_of_day >= 750) & (minute_of_day < 960)).any()
    outside = ((minute_of_day < 120) | (minute_of_day >= 960)).any()

    # Full-day uploads must have activity in BOTH shift windows.
    # An isolated out-of-shift/overtime call does not prove full-day coverage.
    if first and second:
        return "full", False
    if first:
        return "first", False
    if second:
        return "second", False
    # Without evidence for either shift we cannot assume a full-day file.
    return "full", True


def _filter_snb_activity_period(frame, activity_kind, daily_period,
                                exclude_after_shift=True):
    """Select the daily shift; optionally discard overtime AFTER its end.

    The original end-of-shift hour is always included: 11:00-12:00 first,
    and 15:00-16:00 second. The optional extra hour starts at 12:00/16:00.
    """
    valid = frame.dropna(subset=["__ejada_parsed_created"]).copy()
    if valid.empty:
        return valid, 0
    timestamps = valid["__ejada_parsed_created"]
    day_start = timestamps.max().normalize()
    overtime_excluded = 0

    if activity_kind == "يومي":
        if daily_period == "first":
            lower = day_start + pd.Timedelta(hours=2)
            shift_end = day_start + pd.Timedelta(hours=12)
        elif daily_period == "second":
            lower = day_start + pd.Timedelta(hours=12, minutes=30)
            shift_end = day_start + pd.Timedelta(hours=16)
        else:
            lower, shift_end = day_start, day_start + pd.Timedelta(days=1)

        # The extra after-shift hour can be kept only when explicitly allowed.
        if daily_period in ("first", "second"):
            upper = shift_end if exclude_after_shift else (
                shift_end + pd.Timedelta(hours=1)
            )
            in_overtime = (
                (timestamps >= shift_end)
                & (timestamps < shift_end + pd.Timedelta(hours=1))
                & (timestamps.dt.normalize() == day_start)
            )
            if exclude_after_shift:
                overtime_excluded = int(in_overtime.sum())
        else:
            upper = shift_end
        valid = valid.loc[
            (timestamps >= lower) & (timestamps < upper)
        ].copy()
    elif activity_kind == "أسبوعي":
        valid = valid.loc[
            timestamps >= day_start - pd.Timedelta(days=6)
        ].copy()
    elif activity_kind == "شهري":
        valid = valid.loc[
            timestamps >= day_start - pd.Timedelta(days=29)
        ].copy()
    return valid, overtime_excluded

def _snb_after_break_notes(prepared, selected, notes_col, collector_col):
    """ONE row per collector per day with first post-break Note and duty span.

    After-break range is 12:30–16:00. The first/last daily Notes and duty
    span use the full uploaded activity on the dates represented in the report.
    Duty means last timestamp minus first timestamp, with no inferred shift
    start/end or payroll adjustments.
    """
    columns = [
        "التاريخ", "المحصل", "أول إفادة بعد البريك",
        "ميعاد أول إفادة بعد البريك", "ميعاد أول إفادة",
        "ميعاد آخر إفادة", "مدة الدوام",
        "دقائق من نهاية البريك", "الحالة"
    ]
    stamp = "__ejada_parsed_created"
    if prepared.empty or selected.empty:
        return pd.DataFrame(columns=columns)
    active_days = selected[stamp].dropna().dt.normalize().unique()
    source = prepared.dropna(subset=[stamp]).copy()
    source["__ejada_day"] = source[stamp].dt.normalize()
    source = source.loc[source["__ejada_day"].isin(active_days)].copy()
    if source.empty:
        return pd.DataFrame(columns=columns)

    source["__ejada_collector"] = (
        source[collector_col].fillna("").astype(str).str.strip()
        .replace("", "(بدون اسم)")
    )
    source = source.sort_values(
        ["__ejada_day", "__ejada_collector", stamp], kind="stable"
    )
    rows = []
    for (day, collector), group in source.groupby(
        ["__ejada_day", "__ejada_collector"], sort=True, dropna=False
    ):
        first_time = group.iloc[0][stamp]
        last_time = group.iloc[-1][stamp]
        duty_seconds = max(
            0, int(round((last_time - first_time).total_seconds()))
        )
        duty_text = (
            f"{duty_seconds // 3600:02d}:"
            f"{(duty_seconds % 3600) // 60:02d}:"
            f"{duty_seconds % 60:02d}"
        )
        minute_of_day = group[stamp].dt.hour * 60 + group[stamp].dt.minute
        post_break = group.loc[
            (minute_of_day >= 750) & (minute_of_day < 960)
        ]
        post_note, post_time, post_wait = "", "", ""
        status = "لم يسجل إفادة بعد البريك"
        if not post_break.empty:
            first_post = post_break.iloc[0]
            when = first_post[stamp]
            post_note = (
                "" if pd.isna(first_post[notes_col])
                else str(first_post[notes_col])
            )
            post_time = when.strftime("%I:%M:%S %p")
            break_end = pd.Timestamp(day) + pd.Timedelta(
                hours=12, minutes=30
            )
            post_wait = round(
                max(0, (when - break_end).total_seconds() / 60), 2
            )
            status = "سجل أول إفادة بعد البريك"
        rows.append({
            "التاريخ": pd.Timestamp(day).strftime("%Y-%m-%d"),
            "المحصل": collector,
            "أول إفادة بعد البريك": post_note,
            "ميعاد أول إفادة بعد البريك": post_time,
            "ميعاد أول إفادة": first_time.strftime("%I:%M:%S %p"),
            "ميعاد آخر إفادة": last_time.strftime("%I:%M:%S %p"),
            "مدة الدوام": duty_text,
            "دقائق من نهاية البريك": post_wait,
            "الحالة": status,
        })
    return pd.DataFrame(rows, columns=columns)


def _snb_model_similarity(note_key, model_keys):
    """Return similarity to nearest model Notes, NEVER infer the call evaluation."""
    if note_key in model_keys:
        return 100, note_key
    if not note_key or not model_keys:
        return 0, ""
    closest = get_close_matches(note_key, model_keys, n=1, cutoff=0)
    if not closest:
        return 0, ""
    match = closest[0]
    percent = round(100 * SequenceMatcher(None, match, note_key).ratio(), 2)
    return percent, match


def _snb_manual_evaluations(model_upload):
    """Decisions persist across reruns, but NEVER across a changed model."""
    if model_upload is None:
        model_bytes = _snb_secret("SNB_MODEL_ENCRYPTED").encode("ascii")
    else:
        model_bytes = model_upload.getvalue()
    model_digest = hashlib.sha256(model_bytes).hexdigest()
    storage = st.session_state.get("snb_activity_manual_evaluations")
    if not isinstance(storage, dict) or storage.get("model_digest") != model_digest:
        storage = {"model_digest": model_digest, "ratings": {}}
        st.session_state.snb_activity_manual_evaluations = storage
    return storage["ratings"]


def _snb_review_new_notes(notes, mapping, manual, conflicted=0):
    """Ask the user to classify each Note missing from the model before export.

    The text similarity is advisory ONLY. Until every distinct note is
    classified, the processed report and download button are not shown.
    """
    normalized = notes.map(_snb_normalize_note)
    unresolved_keys = set(normalized.unique()) - set(mapping) - set(manual)
    similarity = {}
    for key in normalized.unique():
        similarity[key] = _snb_model_similarity(key, list(mapping))

    if not unresolved_keys:
        return similarity, True

    counts = normalized.value_counts(dropna=False)
    labels = (
        pd.DataFrame({"label": notes.fillna("").astype(str), "key": normalized})
        .drop_duplicates(subset=["key"], keep="first")
        .set_index("key")["label"]
        .to_dict()
    )
    pending = sorted(unresolved_keys, key=lambda key: (-counts.get(key, 0), key))
    review = []
    for key in pending:
        score, closest = similarity[key]
        review.append({
            "الإفادة": labels.get(key, "") or "(إفادة فارغة)",
            "عدد المكالمات": int(counts.get(key, 0)),
            "أقرب إفادة في الموديل": closest or "لا يوجد",
            "نسبة التشابه": f"{score}%",
            "التقييم المطلوب": {1: "ناجحة", 0: "غير ناجحة"}.get(
                mapping.get(closest) if key and closest and score > 0 else None, "اختاري التقييم"
            )
        })

    st.warning(
        f"يوجد {len(pending):,} إفادة جديدة أو غير محسومة في موديل النشاط. "
        "حددي هل كل إفادة ناجحة أم غير ناجحة قبل اعتماد التقرير."
    )
    if conflicted:
        st.caption(
            "الإفادات التي لها تقييمان متعارضان في الموديل تعرض للمراجعة أيضًا."
        )
    st.caption(
        "التقييم مقترح من أقرب إفادة في الموديل. راجعي الاقتراح وعدّليه قبل الاعتماد."
    )
    with st.form("snb_activity_new_notes_review_form"):
        edited = st.data_editor(
            pd.DataFrame(review),
            use_container_width=True,
            hide_index=True,
            disabled=[
                "الإفادة", "عدد المكالمات",
                "أقرب إفادة في الموديل", "نسبة التشابه"
            ],
            column_config={
                "التقييم المطلوب": st.column_config.SelectboxColumn(
                    "هل الإفادة ناجحة أم غير ناجحة؟",
                    options=["اختاري التقييم", "ناجحة", "غير ناجحة"],
                    required=True
                )
            },
            key="snb_activity_new_notes_editor"
        )
        approved = st.form_submit_button(
            "اعتماد تقييم الإفادات المحددة"
        )
    if approved:
        saved = 0
        for i, row in edited.iterrows():
            answer = row["التقييم المطلوب"]
            if answer in ("ناجحة", "غير ناجحة"):
                manual[pending[i]] = 1 if answer == "ناجحة" else 0
                saved += 1
        if saved:
            st.rerun()
        st.error("اختاري ناجحة أو غير ناجحة لإفادة واحدة على الأقل.")
    return similarity, False



def _snb_review_trigger_reasons(note):
    """Return reasons requiring row-by-row manual classification.

    The word «تم» by itself requires review, EXCEPT for payment or debit
    phrases that should be evaluated by exact match with the activity model.
    A slash or «معالج» still forces review, even in those phrases.
    """
    value = "" if pd.isna(note) else str(note)
    reasons = []
    if "/" in value:
        reasons.append("/ أو //")
    if re.search(r"(?<!\w)معالج(?!\w)", value):
        reasons.append("معالج")

    # These established payment messages must use their exact model rating,
    # not a new manual review just because they contain the word «تم».
    # The discount/withdrawal exemption is specifically for text starting
    # with «تم خصم» or «تم سحب», including later details or amounts.
    normalized = re.sub(r"\s+", " ", value).strip()
    payment_from_model = bool(
        re.search(r"(?<!\w)تم\s+سداد\s+جزء(?!\w)", normalized)
        or re.search(r"(?<!\w)تم\s+السداد(?!\w)", normalized)
        or re.match(r"تم\s+(?:خصم|سحب)(?!\w)", normalized)
    )
    if not payment_from_model and re.search(r"(?<!\w)تم(?!\w)", value):
        reasons.append("تم")
    return reasons

def _snb_special_review_mask(notes):
    """Rows requiring explicit evaluation even when their Note matches the model."""
    return notes.map(lambda note: bool(_snb_review_trigger_reasons(note)))


def _snb_slash_row_decisions(activity_upload):
    """Store explicit special-review decisions for the current activity upload."""
    digest = hashlib.sha256(activity_upload.getvalue()).hexdigest()
    storage = st.session_state.get("snb_activity_slash_row_decisions")
    if not isinstance(storage, dict) or storage.get("digest") != digest:
        storage = {"digest": digest, "ratings": {}}
        st.session_state["snb_activity_slash_row_decisions"] = storage
    return storage["ratings"]


def _snb_review_slash_rows(frame, notes_col, activity_upload, model_mapping, manual_mapping=None):
    """Prefill exact-model ratings for review; commit ONLY explicit row approval."""
    decisions = _snb_slash_row_decisions(activity_upload)
    marked = _snb_special_review_mask(frame[notes_col])
    flagged = frame.loc[marked].copy()
    if flagged.empty:
        return decisions, True

    original_indices = flagged.index.tolist()
    # Full original customer data, plus a very visible first-column decision.
    full_rows = flagged.drop(
        columns=["__ejada_parsed_created"], errors="ignore"
    ).copy()
    reasons = [
        "، ".join(_snb_review_trigger_reasons(note))
        for note in flagged[notes_col]
    ]
    manual_mapping = manual_mapping or {}
    # Model values are ONLY editable suggestions here, never approvals.
    # Previously approved row choices override the suggestion.
    model_keys = list(model_mapping)
    matches = {
        key: _snb_model_similarity(key, model_keys)
        for key in flagged[notes_col].map(_snb_normalize_note).unique()
    }
    def _suggestion(note, original_index):
        key = _snb_normalize_note(note)
        prior = decisions.get(str(original_index))
        if prior in (0, 1):
            return prior, "معتمد سابقًا"
        if key in manual_mapping:
            return manual_mapping[key], "تقييم معتمد للإفادة"
        if key in model_mapping:
            return model_mapping[key], "موديل النشاط"
        score, nearest = matches[key]
        if score > 0 and nearest in model_mapping:
            return model_mapping[nearest], "أقرب إفادة في الموديل"
        return None, "لا توجد مطابقة"

    suggestions = [
        _suggestion(note, original_index)
        for note, original_index in zip(flagged[notes_col], original_indices)
    ]
    review = pd.DataFrame(index=full_rows.index)
    review["احسب المكالمة"] = [
        {1: "ناجحة", 0: "غير ناجحة"}.get(suggested, "اختاري التقييم")
        for suggested, source in suggestions
    ]
    review["مصدر التقييم"] = [source for suggested, source in suggestions]
    review["رقم الصف في الملف"] = [int(i) + 2 for i in original_indices]
    review["الإفادة"] = flagged[notes_col].fillna("").astype(str).tolist()
    note_matches = [matches[_snb_normalize_note(note)] for note in flagged[notes_col]]
    review["نسبة التقارب"] = [score for score, _ in note_matches]
    review["أقرب إفادة في الموديل"] = [nearest for _, nearest in note_matches]
    review["سبب المراجعة"] = reasons
    # Preserve *every* uploaded source field in the same editable table.
    # Avoid repeating the Notes column, already displayed as "الإفادة".
    for column in full_rows.columns:
        if column not in review.columns and column != notes_col:
            review[column] = full_rows[column].values
    locked = [c for c in review.columns if c != "احسب المكالمة"]

    pending = sum(str(i) not in decisions for i in original_indices)
    with st.expander(
        f"📋 مراجعة الإفادات وبيانات العملاء الكاملة — {len(flagged):,} صف "
        f"(متبقي {pending:,})",
        expanded=True
    ):
        st.warning(
            "تم تحديد الصفوف التي تحتوي على / أو // أو «معالج» أو «تم» "
            "(باستثناء الإفادات الخاصة بالسداد والخصم والسحب التي تُقيَّم من الموديل). "
            "راجعي كل صف ظاهر حتى لو طابق موديل النشاط."
        )
        st.caption("التقييم مقترح من أقرب إفادة في الموديل. الأصفر يعني تقارب أقل من 50%. تقدري تعدّلي التقييم قبل الاعتماد.")
        order = st.selectbox("ترتيب حسب اللون · Sort by color",
                             ["ترتيب الملف", "الأصفر أولًا", "بدون لون أولًا"])
        if order != "ترتيب الملف":
            low = review["نسبة التقارب"].lt(50)
            review = review.loc[low.sort_values(ascending=order == "بدون لون أولًا",
                                                kind="stable").index]
        original_indices = review.index.tolist()
        styles = pd.DataFrame("", index=review.index, columns=["الإفادة"])
        styles.loc[review["نسبة التقارب"].lt(50), "الإفادة"] = "background-color: #FFFF00; color: #111111"
        styled = review.style.apply(lambda _: styles, axis=None, subset=["الإفادة"])
        with st.form("snb_activity_slash_rows_review_form"):
            edited = st.data_editor(
                styled,
                use_container_width=True,
                hide_index=True,
                disabled=locked,
                column_config={
                    "نسبة التقارب": st.column_config.NumberColumn(format="%.2f%%"),
                    "احسب المكالمة": st.column_config.SelectboxColumn(
                        "احسب المكالمة",
                        options=["اختاري التقييم", "ناجحة", "غير ناجحة"],
                        required=True,
                        width="medium",
                    ),
                    "مصدر التقييم": st.column_config.TextColumn(
                        "مصدر التقييم", width="medium"
                    ),
                    "رقم الصف في الملف": st.column_config.NumberColumn(
                        "رقم الصف", width="small"
                    ),
                    "الإفادة": st.column_config.TextColumn(
                        "الإفادة", width="large"
                    ),
                    "سبب المراجعة": st.column_config.TextColumn(
                        "سبب المراجعة", width="medium"
                    ),
                },
                height=min(max(200, len(review) * 36 + 44), 620),
                key="snb_activity_slash_rows_editor_" + hashlib.sha256(
                    (hashlib.sha256(activity_upload.getvalue()).hexdigest()
                     + ":" + ",".join(map(str, original_indices))).encode("utf-8")
                ).hexdigest()[:16],
            )
            approved = st.form_submit_button(
                "✅ اعتماد تقييم الإفادات المحددة للمراجعة"
            )

    if approved:
        for position, original_index in enumerate(original_indices):
            choice = edited.iloc[position]["احسب المكالمة"]
            if choice in ("ناجحة", "غير ناجحة"):
                decisions[str(original_index)] = 1 if choice == "ناجحة" else 0
            else:
                decisions.pop(str(original_index), None)
        if all(str(i) in decisions for i in original_indices):
            st.rerun()
        st.error("اختاري ناجحة أو غير ناجحة لكل صف قبل استخراج التقرير.")
    pending = sum(str(i) not in decisions for i in original_indices)
    if pending:
        st.info(
            f"متبقي {pending:,} صف لم يتم تحديد تقييمه. "
            "لا يمكن اعتماد التقرير حتى اكتمال المراجعة."
        )
    return decisions, pending == 0


def _snb_add_evaluation(frame, notes_col, mapping, manual, similarity, row_decisions=None):
    """Place call evaluation and percentage match in adjacent columns after Notes."""
    output = frame.copy()
    for name in ("تقييم المكالمات", "نسبة تطابق التقييم"):
        if name in output.columns:
            output = output.drop(columns=[name])
    keys = output[notes_col].map(_snb_normalize_note)
    # Only exact model matches or user-approved new Notes set the evaluation.
    values = keys.map(mapping).combine_first(keys.map(manual))
    # Explicit row-level decisions take precedence over model and note-level ratings.
    if row_decisions:
        overrides = pd.Series([row_decisions.get(str(i)) for i in output.index], index=output.index)
        values = overrides.combine_first(values)
    labels = values.map({1: "ناجحة", 0: "غير ناجحة"}).fillna("غير مصنفة")
    index = output.columns.get_loc(notes_col) + 1
    output.insert(index, "تقييم المكالمات", labels)
    output.insert(
        index + 1, "نسبة تطابق التقييم",
        keys.map(lambda key: f"{similarity.get(key, (0, ''))[0]}%")
    )
    return output


def _snb_add_call_duration(frame, collector):
    """Gap in MINUTES from prior chronological call by SAME collector."""
    output = frame.copy()
    name = "مده المكالمه"
    if name in output.columns:
        output = output.drop(columns=[name])
    stamp = "__ejada_parsed_created"
    ordered = output.sort_values(
        [collector, stamp], kind="stable"
    )
    # Never treat an overnight interval as a gap inside the same workday.
    collector_keys = (
        ordered[collector].fillna("").astype(str).str.strip()
    )
    days = ordered[stamp].dt.normalize()
    gaps = (
        ordered.groupby([collector_keys, days], dropna=False)[stamp]
        .diff().dt.total_seconds().div(60)
        .fillna(0).clip(lower=0).round(2)
    )
    output.insert(
        output.columns.get_loc("Time") + 1,
        name, gaps.reindex(output.index).fillna(0)
    )
    return output


def _ejada_activity_dashboard_is_open(bank):
    """Track whether the Activity KPI and chart dashboard is open per bank."""
    return bool(st.session_state.get(f"ejada_activity_dashboard_open_{bank}", False))


def _ejada_toggle_activity_dashboard(bank):
    key = f"ejada_activity_dashboard_open_{bank}"
    st.session_state[key] = not bool(st.session_state.get(key, False))


def _ejada_activity_kpis(bank, covered, successful=None,
                           unsuccessful=None, unresolved=0):
    """Dashboard section for call KPIs and future Activity charts."""
    if not _ejada_activity_dashboard_is_open(bank):
        return
    with st.container(key=f"ejada_activity_dashboard_{bank}"):
        st.markdown(
            f'<div style="color:{label_color};font-weight:900;'
            'font-size:20px;text-align:center;margin:12px 0 8px;">'
            f'📊 لوحة مؤشرات النشاط | {bank}</div>',
            unsafe_allow_html=True
        )
        with st.container(key="ejada_activity_kpi_panel"):
            left, center, right = st.columns(3, gap="small")
            with left:
                st.metric("📞 إجمالي المكالمات المغطاة", f"{int(covered):,}")
            with center:
                st.metric(
                    "✅ إجمالي المكالمات الناجحة",
                    f"{int(successful):,}" if successful is not None else "—"
                )
            with right:
                st.metric(
                    "❌ إجمالي المكالمات غير الناجحة",
                    f"{int(unsuccessful):,}" if unsuccessful is not None else "—"
                )
        if unresolved:
            st.caption(
                f"يوجد {int(unresolved):,} مكالمة لم تُعتمد نتيجتها بعد؛ "
                "لا تُحتسب ناجحة ولا غير ناجحة."
            )
        # Future chart widgets are added beneath the KPI panel here.


def _snb_kpi_counts(selected, notes_col, mapping, manual, row_decisions):
    """Compute current SNB KPIs without auto-rating any special/unknown note."""
    keys = selected[notes_col].map(_snb_normalize_note)
    values = keys.map(mapping).combine_first(keys.map(manual))
    special = _snb_special_review_mask(selected[notes_col])
    # Special Notes ALWAYS require a per-row explicit decision.
    values.loc[special] = None
    if row_decisions:
        overrides = pd.Series(
            [row_decisions.get(str(index)) for index in selected.index],
            index=selected.index, dtype=object
        )
        values = overrides.combine_first(values)
    successful = int((values == 1).sum())
    unsuccessful = int((values == 0).sum())
    return len(selected), successful, unsuccessful, (
        len(selected) - successful - unsuccessful
    )


def _ejada_stc_activity_kpi_counts(uploaded):
    """Coverage from STC uploaded records; success only from explicit call rating."""
    if uploaded is None:
        return 0, 0, 0, 0
    name = uploaded.name.lower()
    raw = io.BytesIO(uploaded.getvalue())
    if name.endswith(".csv"):
        try:
            df = pd.read_csv(raw, dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            raw.seek(0)
            df = pd.read_csv(raw, dtype=str, encoding="cp1256")
    else:
        df = pd.read_excel(raw, dtype=str)
    df.columns = df.columns.astype(str).str.strip()
    id_col = _snb_find_column(
        df, "Debitor", "Debtor", "ID Numer", "ID Number", "ID",
        "رقم العميل", "رقم الهوية"
    )
    notes_col = _snb_find_column(
        df, "Notes", "Note", "الإفادة", "ملاحظات", "Feedback"
    )
    collector_col = _snb_find_column(
        df, "Collector", "Collecter", "المحصل", "Agent"
    )
    if all(c is not None for c in (id_col, notes_col, collector_col)):
        df = df.drop_duplicates(
            subset=[id_col, notes_col, collector_col], keep="first"
        ).copy()
    status_col = _snb_find_column(
        df, "تقييم المكالمات", "تقييم المكالمة", "التقيم",
        "التقييم", "Evaluation", "Call_Status", "Call Status",
        "Call Result", "Result"
    )
    if status_col is None:
        return len(df), None, None, len(df)
    raw_values = df[status_col].fillna("").astype(str).str.strip().str.casefold()
    # Explicit statuses only: unknown words must NEVER become failures.
    success_values = {"1", "ناجحة", "ناجح", "successful", "success", "yes"}
    unsuccessful_values = {
        "0", "غير ناجحة", "غير ناجح", "unsuccessful", "failed",
        "failure", "no"
    }
    successful = int(raw_values.isin(success_values).sum())
    unsuccessful = int(raw_values.isin(unsuccessful_values).sum())
    return len(df), successful, unsuccessful, (
        len(df) - successful - unsuccessful
    )


def _style_snb_activity_workbook(workbook):
    """Apply consistent navy headers, bold collector names and full table borders."""
    header_fill = PatternFill(fill_type="solid", fgColor="1F4373")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    collector_font = Font(name="Calibri", size=11, bold=True, color="173452")
    inner = Side(style="thin", color="B6C4D2")
    outer = Side(style="medium", color="1F4373")

    for sheet in workbook.worksheets:
        last_row, last_col = sheet.max_row, sheet.max_column
        if not last_row or not last_col:
            continue
        sheet.sheet_view.showGridLines = False
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:{get_column_letter(last_col)}{last_row}"
        sheet.row_dimensions[1].height = 34

        collector_columns = set()
        for col_idx in range(1, last_col + 1):
            heading = str(sheet.cell(row=1, column=col_idx).value or "").strip()
            if heading.casefold() in ("المحصل", "اسم المحصل", "collector"):
                collector_columns.add(col_idx)
                width = 44
            elif heading in ("التاريخ", "الأسبوع"):
                width = 30 if heading == "الأسبوع" else 17
            elif heading in ("الإفادة", "notes"):
                width = 54
            else:
                width = min(46, max(16, len(heading) + 4))
            sheet.column_dimensions[get_column_letter(col_idx)].width = width

        for row in sheet.iter_rows(min_row=1, max_row=last_row,
                                   min_col=1, max_col=last_col):
            for cell in row:
                row_idx, col_idx = cell.row, cell.column
                cell.border = Border(
                    left=outer if col_idx == 1 else inner,
                    right=outer if col_idx == last_col else inner,
                    top=outer if row_idx == 1 else inner,
                    bottom=outer if row_idx == last_row else inner,
                )
                if row_idx == 1:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=True
                    )
                elif col_idx in collector_columns:
                    cell.font = collector_font
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(
                        horizontal="right" if isinstance(cell.value, str) else "center",
                        vertical="center",
                    )


def _snb_wasted_time_reports(result, collector_col, notes_col, activity_kind="يومي"):
    """Return audited per-call waste and daily or weekly collector totals.

    Time gaps remain within each collector/day (never overnight). Subtract
    overlap with the fixed 12:00–12:30 break, then apply the current call's
    outcome threshold: successful = 20 min, unsuccessful = 5 min. Only the
    excess is wasted; unclassified calls receive no invented threshold.
    """
    detail_cols = [
        "التاريخ", "المحصل", "العميل", "وقت الإفادة", "الإفادة",
        "تقييم المكالمات", "مده المكالمه (دقائق)",
        "البريك المستبعد (دقائق)", "المدة بعد استبعاد البريك (دقائق)",
        "الحد المسموح (دقائق)", "الوقت المهدر (دقائق)", "الحالة"
    ]
    weekly = activity_kind == "أسبوعي"
    period_heading = "الأسبوع" if weekly else "التاريخ"
    summary_cols = [
        period_heading, "المحصل", "إجمالي المكالمات",
        "المكالمات الناجحة", "المكالمات غير الناجحة",
        "عدد المكالمات بوقت مهدر",
        "إجمالي الوقت المهدر (دقائق)",
        "إجمالي الوقت المهدر (ساعات)"
    ]
    stamp = "__ejada_parsed_created"
    if result.empty:
        return (
            pd.DataFrame(columns=detail_cols),
            pd.DataFrame(columns=summary_cols)
        )
    frame = result.dropna(subset=[stamp]).copy()
    if frame.empty:
        return (
            pd.DataFrame(columns=detail_cols),
            pd.DataFrame(columns=summary_cols)
        )
    frame["__ejada_day"] = frame[stamp].dt.normalize()
    frame["__ejada_collector"] = (
        frame[collector_col].fillna("").astype(str).str.strip()
        .replace("", "(بدون اسم)")
    )
    frame = frame.sort_values(
        ["__ejada_day", "__ejada_collector", stamp], kind="stable"
    )
    prev_time = frame.groupby(
        ["__ejada_day", "__ejada_collector"], dropna=False
    )[stamp].shift(1)
    # The exported Activity 'مده المكالمه' is computed from the same
    # per-collector/per-day timestamps, and is the source for this report.
    raw_gap = (
        pd.to_numeric(frame["مده المكالمه"], errors="coerce")
        .fillna(0).clip(lower=0).round(2)
    )
    break_deduction = pd.Series(
        [
            _break_overlap_minutes_for_report(previous, current)
            for previous, current in zip(prev_time, frame[stamp])
        ],
        index=frame.index, dtype="float64"
    ).round(2)
    effective_gap = (
        raw_gap - break_deduction
    ).clip(lower=0).round(2)
    statuses = frame["تقييم المكالمات"].fillna("").astype(str).str.strip()
    thresholds = statuses.map({
        "ناجحة": 20.0, "غير ناجحة": 5.0
    })
    wasted = (
        (effective_gap - thresholds.fillna(0))
        .clip(lower=0)
        .where(thresholds.notna(), 0)
        .round(2)
    )
    debtor_col = _snb_find_column(
        frame, "Debitor", "Debtor", "رقم العميل", "رقم الهوية"
    )
    customers = (
        frame[debtor_col].fillna("").astype(str)
        if debtor_col else pd.Series("", index=frame.index)
    )
    detail = pd.DataFrame({
        "التاريخ": frame[stamp].dt.strftime("%Y-%m-%d"),
        "المحصل": frame["__ejada_collector"],
        "العميل": customers,
        "وقت الإفادة": frame[stamp].dt.strftime("%I:%M:%S %p"),
        "الإفادة": frame[notes_col].fillna("").astype(str),
        "تقييم المكالمات": statuses,
        "مده المكالمه (دقائق)": raw_gap,
        "البريك المستبعد (دقائق)": break_deduction,
        "المدة بعد استبعاد البريك (دقائق)": effective_gap,
        "الحد المسموح (دقائق)": thresholds,
        "الوقت المهدر (دقائق)": wasted,
        "الحالة": [
            (
                "غير مصنفة - لا يحتسب وقت مهدر"
                if pd.isna(limit)
                else "وقت مهدر" if minutes > 0
                else "لا يوجد وقت مهدر"
            )
            for limit, minutes in zip(thresholds, wasted)
        ]
    }, index=frame.index)
    detail = detail[detail_cols].reset_index(drop=True)

    rows = []
    if weekly:
        week_start = frame[stamp].min().strftime("%Y-%m-%d")
        week_end = frame[stamp].max().strftime("%Y-%m-%d")
        report_period = f"{week_start} إلى {week_end}"
        grouped = (
            (report_period, collector, group)
            for collector, group in detail.groupby(
                "المحصل", sort=True, dropna=False
            )
        )
    else:
        grouped = (
            (day, collector, group)
            for (day, collector), group in detail.groupby(
                ["التاريخ", "المحصل"], sort=True, dropna=False
            )
        )
    for period_label, collector, group in grouped:
        total_waste = round(
            float(group["الوقت المهدر (دقائق)"].sum()), 2
        )
        rows.append({
            period_heading: period_label,
            "المحصل": collector,
            "إجمالي المكالمات": int(len(group)),
            "المكالمات الناجحة": int(
                (group["تقييم المكالمات"] == "ناجحة").sum()
            ),
            "المكالمات غير الناجحة": int(
                (group["تقييم المكالمات"] == "غير ناجحة").sum()
            ),
            "عدد المكالمات بوقت مهدر": int(
                (group["الوقت المهدر (دقائق)"] > 0).sum()
            ),
            "إجمالي الوقت المهدر (دقائق)": total_waste,
            "إجمالي الوقت المهدر (ساعات)": round(
                total_waste / 60, 2
            ),
        })
    return detail, pd.DataFrame(rows, columns=summary_cols)


st.markdown("""<style>[data-testid="stExpander"] table :is(td,th){font-size:16px!important;border:1px solid #9eafbf!important}[data-testid="stExpander"] table th{text-align:center!important;font-weight:700!important}</style>""", unsafe_allow_html=True)

def _expander_table(frame, key, show_index=False):
    page_size = 100
    page = 1
    if len(frame) > page_size:
        pages = (len(frame) + page_size - 1) // page_size
        page = st.number_input("صفحة الجدول", 1, pages, key=f"table_page_{key}")
    start = (page - 1) * page_size
    table = frame.iloc[start:start + page_size].to_html(
        index=show_index, escape=True, border=0, na_rep="",
        float_format=lambda value: f"{value:.2f}"
    )
    st.markdown(
        """<style>.rana-table{overflow:auto;max-height:450px}
        .rana-table table{border-collapse:collapse;width:100%}
        .rana-table table :is(th,td){border:1px solid #9eafbf;padding:7px 10px;font-size:16px}
        .rana-table table th{text-align:center!important;font-weight:700}
        .rana-table thead th{position:sticky;top:0;background:#eaf0f6;color:#173452}
        </style><div class="rana-table">""" + table + "</div>", unsafe_allow_html=True
    )


def _activity_period(stamps):
    valid = stamps.dropna()
    if valid.empty:
        return None, "رنا، الفترة مش واضحة لأن مفيش تاريخ صالح في Created On."
    first, last = valid.min(), valid.max()
    days = (last.date() - first.date()).days + 1
    if days > 7:
        return "شهري", "رنا إحنا في النشاط الشهري"
    if days > 1:
        return "أسبوعي", "رنا إحنا في النشاط الأسبوعي"
    if last.time() <= time(12):
        return "يومي", "رنا إحنا في الفترة الأولى"
    if first.time() > time(12):
        return "يومي", "رنا إحنا في الفترة التانية"
    return "يومي", "رنا إحنا في النشاط اليومي"


def render_snb_activity_report(model_upload, activity_upload, activity_kind="يومي", error_slot=None):
    """Report screen opens with both upload boxes still visible above."""
    if activity_upload is None:
        return
    if model_upload is None and not _snb_secret("SNB_MODEL_ENCRYPTED"):
        return

    with st.container(key="snb_activity_work_area"):
        # The activity kind is selected in the compact header row above.
        try:
            prepared, stats = _prepare_snb_activity_file(activity_upload)
            if _snb_secret("SNB_MODEL_ENCRYPTED"):
                model_mapping, conflicted_notes, invalid_model = (
                    _snb_secret_model_evaluations()
                )
            else:
                model_mapping, conflicted_notes, invalid_model = (
                    _snb_model_evaluations(model_upload)
                )
        except Exception as exc:
            message = f"تعذر تجهيز التقرير: {exc}"
            if error_slot is not None:
                error_slot.error(message)
            else:
                st.error(message)
            return
        # Persistent model sync failure must never disable the Activity report.
        # Display a warning rather than silently pretending remote notes loaded.
        try:
            saved_model_mapping, _, _ = _snb_load_saved_model_notes(
                _snb_model_git_settings()
            )
            model_mapping.update(saved_model_mapping)
        except Exception as exc:
            st.warning(
                "تعذر تحميل إضافات موديل النشاط المحفوظة على GitHub؛ "
                "التقرير سيعتمد مؤقتًا على الموديل الأساسي فقط. "
                f"التفاصيل: {exc}"
            )

        detected_kind, message = _activity_period(prepared["__ejada_parsed_created"])
        st.info(message)
        automatic = activity_kind == "تلقائي"
        if automatic:
            if detected_kind is None:
                return
            activity_kind = detected_kind
        period = "full"
        exclude_after_shift = True
        if not automatic and activity_kind == "يومي":
            suggested_period, unsure = _infer_snb_daily_period(prepared)
            if suggested_period == "full" and not unsure:
                # Automatically honor a file containing both shift windows:
                # no manual shift selector and no overtime checkbox.
                period = "full"
                st.caption(
                    "✅ تم اكتشاف ملف اليوم الكامل (الفترتان الأولى والثانية). "
                    "تم عرض جميع مكالمات اليوم بدون اختيار فترة أو حذف ساعة."
                )
            else:
                shift_labels = [
                    "الفترة الأولى (2:00 ص – 12:00 ظ)",
                    "الفترة الثانية (12:30 م – 4:00 م)",
                ]
                with st.container(key="snb_activity_daily_options"):
                    shift_col, last_hour_col, _controls_space = st.columns(
                        [1.8, 1.3, 5.0], gap="small", vertical_alignment="bottom"
                    )
                    with shift_col:
                        shift_choice = st.selectbox(
                            "اختاري الفترة اليومية",
                            shift_labels,
                            index=1 if suggested_period == "second" else 0,
                            key="snb_activity_daily_shift"
                        )
                    with last_hour_col:
                        exclude_after_shift = st.checkbox(
                            "استبعاد الساعة الزيادة بعد نهاية الفترة",
                            value=True,
                            key="snb_activity_exclude_after_shift"
                        )
                period = "first" if shift_choice == shift_labels[0] else "second"
                st.caption(
                    "البريك من 12:00 إلى 12:30. "
                    "لو فعلتي الاستبعاد: هتتحذف تسجيلات الساعة 12:00–1:00 ظهرًا "
                    "للفترة الأولى، أو 4:00–5:00 عصرًا للفترة الثانية. "
                    "آخر ساعة داخل الشيفت نفسه هتفضل محسوبة."
                )

        if automatic:
            selected = prepared.dropna(subset=["__ejada_parsed_created"]).copy()
            excluded_after_shift = 0
        else:
            selected, excluded_after_shift = _filter_snb_activity_period(
                prepared, activity_kind, period,
                exclude_after_shift=exclude_after_shift
            )
        manual_decisions = _snb_manual_evaluations(model_upload)
        slash_row_decisions = _snb_slash_row_decisions(activity_upload)
        if _ejada_activity_dashboard_is_open("SNB"):
            counts = _snb_kpi_counts(
                selected, stats["notes"], model_mapping,
                manual_decisions, slash_row_decisions
            )
            _ejada_activity_kpis("SNB", *counts)
        slash_row_decisions, slash_reviewed = _snb_review_slash_rows(
            selected, stats["notes"], activity_upload,
            model_mapping, manual_decisions
        )
        if not slash_reviewed:
            return
        # Special-review rows use their own row-level manual decisions.
        # Skip repeat note-level review for those rows.
        ordinary_notes = selected.loc[
            ~_snb_special_review_mask(selected[stats["notes"]]),
            stats["notes"]
        ]
        similarity, fully_reviewed = _snb_review_new_notes(
            ordinary_notes,
            model_mapping,
            manual_decisions,
            conflicted=conflicted_notes
        )
        if not fully_reviewed:
            # Don't issue a report or download with unknown call evaluations.
            return

        # Preserve a meaningful model-similarity percentage for reviewed notes.
        for key in selected[stats["notes"]].map(_snb_normalize_note).unique():
            if key not in similarity:
                similarity[key] = _snb_model_similarity(key, list(model_mapping))

        result = _snb_add_evaluation(
            selected, stats["notes"], model_mapping,
            manual_decisions, similarity, row_decisions=slash_row_decisions
        )
        result = _snb_add_call_duration(
            result, stats["collector"]
        )
        successful = int(
            (result["تقييم المكالمات"] == "ناجحة").sum()
        )
        unsuccessful = int(
            (result["تقييم المكالمات"] == "غير ناجحة").sum()
        )
        unknown = len(result) - successful - unsuccessful

        st.markdown(
            f'<p style="color:{label_color};font-weight:800;direction:rtl;'
            'text-align:right;font-size:15px;line-height:1.9;">'
            f'صفوف الملف: {stats["original"]:,} | '
            f'التكرارات المحذوفة: {stats["duplicates"]:,} | '
            f'بعد الفترة: {len(result):,} | '
            f'المستبعد بعد نهاية الفترة: {excluded_after_shift:,} | '
            f'ناجحة: {successful:,} | غير ناجحة: {unsuccessful:,} | '
            f'غير مصنفة: {unknown:,}'
            '</p>',
            unsafe_allow_html=True
        )

        if stats["bad_created"]:
            st.warning(
                f'عدد السجلات ذات Created on غير صالح: '
                f'{stats["bad_created"]:,}'
            )
        if stats["product"] is None:
            st.warning(
                "لم أتعرف على عمود Claim/نوع المتنج-التمويل؛ "
                "يرجى مراجعة أسماء أعمدة ملف النشاط."
            )
        else:
            st.caption(
                f'عمود المنتج المعتمد: {stats["product"]} | '
                f'عدد القيم التي تم تحويلها إلى CC+PF (من CC أو PF): '
                f'{stats["combined_products"]:,} (قبل إزالة التكرارات).'
            )
        if conflicted_notes:
            st.warning(
                f'موديل النشاط يحتوي على {conflicted_notes} إفادة لها '
                'تقييمان متعارضان (0 و1)؛ تركت هذه الإفادات غير مصنفة.'
            )
        if invalid_model:
            st.caption(
                f"تم تجاهل {invalid_model:,} صف غير صالح في موديل النشاط."
            )
        model_matched = int(
            selected[stats["notes"]].map(_snb_normalize_note)
            .isin(model_mapping).sum()
        )
        model_coverage = (
            100 * model_matched / len(selected) if len(selected) else 0
        )
        st.caption(
            f"نسبة الإفادات المطابقة تمامًا لموديل النشاط: "
            f"{model_coverage:.1f}% ({model_matched:,} من {len(selected):,}). "
            "عمود «نسبة تطابق التقييم» يقيس تشابه نص الإفادة بأقرب نص "
            "في الموديل، ولا يقيس صحة التقييم اليدوي."
        )
        st.caption(
            "Notes = الإفادة. 0 = غير ناجحة، 1 = ناجحة. "
            "مده المكالمه محسوبة بالدقائق بين مكالمتين متتاليتين "
            "لنفس المحصل؛ أول مكالمة = صفر."
        )
        final = result.drop(
            columns=["__ejada_parsed_created"], errors="ignore"
        )
        if final.empty:
            st.warning("لا توجد بيانات صالحة للفترة المكتشفة.")
            return
        st.dataframe(
            final, use_container_width=True, hide_index=True,
            height=460
        )
        after_break = _snb_after_break_notes(
            prepared, selected, stats["notes"], stats["collector"]
        )
        wasted_details, wasted_summary = _snb_wasted_time_reports(
            result, stats["collector"], stats["notes"],
            activity_kind=activity_kind
        )
        with st.expander(
            "🕧 أول إفادة لكل محصل بعد البريك + مدة الدوام", expanded=False
        ):
            st.caption(
                "صف واحد لكل محصل في كل يوم. مدة الدوام = وقت آخر إفادة "
                "ناقص وقت أول إفادة في الملف المرفوع (وتشمل فترة البريك)."
            )
            _expander_table(after_break, "after_break")
        with st.expander(
            f"⏱️ تفاصيل الوقت المهدر وإجماليه {'الأسبوعي' if activity_kind == 'أسبوعي' else 'اليومي'}",
            expanded=False
        ):
            st.caption(
                "الوقت المهدر = الفاصل بين الإفادات بعد خصم البريك "
                "ناقص 20 دقيقة للناجحة أو 5 دقائق لغير الناجحة؛ "
                "ولا يمكن أن يقل عن صفر."
            )
            _expander_table(wasted_details, "wasted_details")
            _expander_table(wasted_summary, "wasted_summary")
        _snb_render_model_management(
            selected[stats["notes"]], model_mapping,
            manual_decisions, slash_row_decisions
        )
        weekly_report = activity_kind == "أسبوعي"
        details_sheet = (
            "Wasted_Time_Details_Weekly" if weekly_report
            else "Wasted_Time_Details"
        )
        summary_sheet = (
            "Wasted_Time_Weekly" if weekly_report
            else "Wasted_Time_Daily"
        )
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final.to_excel(writer, sheet_name="Activity", index=False)
            after_break.to_excel(
                writer, sheet_name="After_Break_Notes", index=False
            )
            wasted_details.to_excel(
                writer, sheet_name=details_sheet, index=False
            )
            wasted_summary.to_excel(
                writer, sheet_name=summary_sheet, index=False
            )
            _style_snb_activity_workbook(writer.book)
        activity_only = io.BytesIO()
        with pd.ExcelWriter(activity_only, engine="openpyxl") as writer:
            final.to_excel(writer, sheet_name="Activity", index=False)
            _style_snb_activity_workbook(writer.book)
        st.download_button(
            "📥 تحميل شيت Activity منفصل", data=activity_only.getvalue(),
            file_name=f"SNB_Activity_{activity_kind}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_snb_activity_only"
        )
        st.download_button(
            "📥 تحميل نشاط SNB بعد المعالجة",
            data=output.getvalue(),
            file_name=f"SNB_Activity_{activity_kind}_Processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_snb_activity_processed"
        )


selected_bank = st.session_state.selected_bank
bank_files = st.session_state.page_files[selected_bank]

curr_page = st.session_state.current_page

def render_promise_filter_diagnostics(bank, kind):
    """Display each bank's unchanged filters, with two diagnostic tables side-by-side."""
    payload = (
        st.session_state.get('promise_filter_debug', {})
        .get(bank, {}).get(kind)
    )
    if not payload:
        return

    def _table_html(items, name_col):
        html_parts = [
            '<table class="ejada-diagnostic-table" dir="rtl">',
            f'<thead><tr><th>{name_col}</th><th>العدد</th></tr></thead><tbody>'
        ]
        for item in items:
            label = _xml_escape(str(item.get(name_col, '')))
            try:
                quantity = f"{int(item.get('العدد', 0)):,}"
            except (TypeError, ValueError):
                quantity = _xml_escape(str(item.get('العدد', '')))
            html_parts.append(f'<tr><td>{label}</td><td>{quantity}</td></tr>')
        html_parts.append('</tbody></table>')
        return ''.join(html_parts)

    kind_label = 'القائمة' if kind == 'active' else 'المكسورة'
    with st.container(key=f'promise_diagnostic_{bank}_{kind}'):
        with st.expander(
            f'🔎 تشخيص فلترة الوعود {kind_label} {bank}',
            expanded=True
        ):
            st.markdown(
                "<div style='font-size:12px;font-weight:800;"
                "text-align:right;color:#586F8C;margin-bottom:7px;'>"
                "الأرقام للتشخيص فقط، ولا تغيّر نتيجة التقرير.</div>",
                unsafe_allow_html=True
            )
            rows = payload.get('steps') or []
            class_rows = payload.get('classification_counts') or []

            if rows and class_rows:
                # Matched caption rows keep both green table headers level.
                # The Classification caption belongs to its own table only.
                diagnostic_tables = (
                    '<div class="ejada-diagnostic-dual" dir="rtl">'
                    '<div class="ejada-diagnostic-steps">'
                    '<div class="ejada-diagnostic-caption-spacer" aria-hidden="true"></div>'
                    + _table_html(rows, 'الخطوة')
                    + '</div><div class="ejada-diagnostic-classes">'
                    '<div class="ejada-diagnostic-small-caption">'
                    'قيم Classification بعد فلترة حالة المعالجة'
                    '</div>'
                    + _table_html(class_rows, 'Classification')
                    + '</div></div>'
                )
                st.markdown(diagnostic_tables, unsafe_allow_html=True)
            elif rows:
                st.markdown(_table_html(rows, 'الخطوة'), unsafe_allow_html=True)
            elif class_rows:
                st.markdown(
                    _table_html(class_rows, 'Classification'),
                    unsafe_allow_html=True
                )


def render_report_upload_bar(page, promise_metrics=None):
    """Compact report header: file details and stacked X/+ alongside the title."""
    titles = {
        'Promises': '🤝 إدارة ومتابعة الوعود (القائمة والمكسورة)',
        'Neglect': '🚨 تقرير الإهمال',
        'Activity': '📈 تقرير نشاط المحصلين',
    }
    current_files = bank_files.get(page) or []
    uploader_version = st.session_state.report_uploader_versions[selected_bank][page]

    with st.container(key=f"report_upload_bar_{selected_bank}_{page}"):
        if current_files:
            names = "، ".join(x.name for x in current_files)
            status_message = f"الملف الحالي: {names}"
        else:
            status_message = "مفيش ملف مرفوع لسه يا رنا"

        # A separate status line ensures filenames never push down either box.
        status_col, _, _ = st.columns([1.7, 3.9, 4.4], gap='small')
        with status_col:
            st.markdown(
                f'<div class="ejada-upload-name" title="{_xml_escape(status_message)}" '
                f'style="color:{label_color};">{_xml_escape(status_message)}</div>',
                unsafe_allow_html=True
            )

        # Top-aligned row used by all three reports for both banks.
        upload_col, title_col, _left_balance = st.columns(
            [1.7, 3.9, 4.4], gap='small', vertical_alignment='top'
        )

        with upload_col:
            with st.container(border=True, key=f"ejada_shared_upload_card_{selected_bank}_{page}"):
                # The first (RTL/rightmost) subcolumn contains X above +.
                actions_col, file_box_col = st.columns(
                    [0.16, 0.84], gap='small', vertical_alignment='top'
                )

                with actions_col:
                    with st.container(key=f"mini_actions_{selected_bank}_{page}"):
                        if current_files:
                            st.button(
                                "×",
                                key=f"small_clear_report_{selected_bank}_{page}",
                                on_click=clear_report_page_file,
                                args=(selected_bank, page),
                                help="مسح الملف المرفوع",
                                use_container_width=False
                            )
                        else:
                            # Reserve the X position so the + stays directly below it.
                            st.markdown(
                                '<div class="ejada-mini-x-placeholder"></div>',
                                unsafe_allow_html=True
                            )

                        new_files = st.file_uploader(
                            f"رفع ملف {titles[page]}",
                            type=['xlsx', 'csv'],
                            accept_multiple_files=True,
                            label_visibility='collapsed',
                            key=f"report_uploader_{selected_bank}_{page}_{uploader_version}"
                        )
                        if new_files:
                            new_sig = [(x.name, x.size) for x in new_files]
                            current_sig = [(x.name, x.size) for x in current_files]
                            if new_sig != current_sig:
                                bank_files[page] = new_files
                                st.rerun()
                        elif current_files:
                            # Preserve native uploader clearing behavior and page state.
                            bank_files[page] = None
                            st.session_state.report_uploader_versions[selected_bank][page] += 1
                            st.rerun()

                with file_box_col:
                    if current_files:
                        main_file = current_files[0]
                        file_name = _xml_escape(main_file.name)
                        file_size = f"{main_file.size / (1024 * 1024):.1f} MB"
                        extra_files = (
                            f" · +{len(current_files) - 1} ملفات"
                            if len(current_files) > 1 else ""
                        )
                        info_html = (
                            f'<div class="ejada-file-name" title="{file_name}">{file_name}</div>'
                            f'<div class="ejada-file-size">{file_size}{extra_files}</div>'
                        )
                    else:
                        info_html = (
                            '<div class="ejada-no-file">📄 لم يتم رفع ملف</div>'
                            '<div class="ejada-file-size">XLSX / CSV</div>'
                        )
                    st.markdown(
                        '<div class="ejada-current-file-box">' + info_html + '</div>',
                        unsafe_allow_html=True
                    )

        with title_col:
            if page == 'Promises':
                _n_total, _n_active, _n_broken = promise_metrics or (0, 0, 0)
                st.markdown(
                    f"""
                    <div class="ejada-promises-header-kpis" dir="rtl">
                        <div class="ejada-mini-kpi">
                            <span>📁 إجمالي الحسابات</span>
                            <strong>{int(_n_total):,}</strong>
                        </div>
                        <div class="ejada-mini-kpi ejada-mini-kpi-active">
                            <span>🤝 الوعود القائمة</span>
                            <strong>{int(_n_active):,}</strong>
                        </div>
                        <div class="ejada-mini-kpi">
                            <span>⚠️ الوعود المكسورة</span>
                            <strong>{int(_n_broken):,}</strong>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            else:
                title_font_size = 20
                st.markdown(
                    f"""
                    <div class="ejada-report-title-holder">
                    <div class="ejada-report-title"
                         style="background:{card_color};border-color:{border_color};
                                color:{text_color};font-size:{title_font_size}px;">
                        {titles[page]}
                    </div>
                </div>
                    """,
                    unsafe_allow_html=True
                )

# دالة معالجة بيانات أي صفحة

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

            # حماية من رفع محفظة بنك في خانة البنك الآخر.
            # نعتمد أولاً على Portfolio Type لأنه أوضح بصمة داخل ملف المحفظة.
            portfolio_type_col = find_col_in(
                temp_df,
                ['Portfolio Type', 'portfolio type', 'Portfolio_Type']
            )
            if portfolio_type_col:
                portfolio_values = (
                    temp_df[portfolio_type_col]
                    .dropna()
                    .astype(str)
                    .str.strip()
                )

                normalized_portfolio_values = (
                    portfolio_values
                    .str.lower()
                    .str.split()
                    .str.join(' ')
                )
                compact_portfolio_values = normalized_portfolio_values.str.replace(
                    ' ', '', regex=False
                )

                is_stc_portfolio = compact_portfolio_values.eq('stcbank').any()

                is_snb_portfolio = (
                    normalized_portfolio_values.str.contains('snb', regex=False, na=False) |
                    normalized_portfolio_values.str.contains('saudi national bank', regex=False, na=False) |
                    normalized_portfolio_values.str.contains('البنك الأهلي السعودي', regex=False, na=False) |
                    normalized_portfolio_values.str.contains('البنك الاهلي السعودي', regex=False, na=False)
                ).any()

                if bank == 'SNB' and is_stc_portfolio:
                    st.error("⚠️ رنا إنتِ في البنك الغلط — الملف ده خاص بـ STC Bank.")
                    st.caption(f"الملف المرفوع: {f.name}")
                    return (
                        pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(),
                        None, 0, 0, 0, 0, 0, 0
                    )

                if bank == 'STC' and is_snb_portfolio:
                    st.error("⚠️ رنا إنتِ في البنك الغلط — الملف ده خاص بالبنك الأهلي السعودي SNB.")
                    st.caption(f"الملف المرفوع: {f.name}")
                    return (
                        pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(),
                        None, 0, 0, 0, 0, 0, 0
                    )

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
    # Diagnostic only; never changes either bank's filter results.
    st.session_state.setdefault('promise_filter_debug', {})[bank] = {
        'active': None, 'broken': None
    }

    if not df_promises.empty and all(c in df_promises.columns for c in ['Follow up Due Date', 'Follow up Last Date', 'حالة المعالجة - التمويل', 'تحديث فئة العميل - التمويل', 'Classification']):
        due_date = pd.to_datetime(df_promises['Follow up Due Date'], errors='coerce').dt.normalize()
        last_date = pd.to_datetime(df_promises['Follow up Last Date'], errors='coerce').dt.normalize()
        base_condition = (
            (df_promises['حالة المعالجة - التمويل'].astype(str).str.strip() == 'غير معالج') & 
            (df_promises['تحديث فئة العميل - التمويل'].astype(str).str.strip().str.lower() != 'writeoff') & 
            (df_promises['Classification'].astype(str).str.strip() == 'وعود السداد قيد التنفيذ')
        )
        if bank == 'SNB':
            _snb_status = (df_promises['حالة المعالجة - التمويل'].astype(str).str.strip() == 'غير معالج')
            _snb_not_writeoff = (
                df_promises['تحديث فئة العميل - التمويل']
                .astype(str).str.strip().str.lower() != 'writeoff'
            )
            _snb_class = (
                df_promises['Classification'].astype(str).str.strip()
                == 'وعود السداد قيد التنفيذ'
            )
            _snb_after_status = _snb_status
            _snb_after_category = _snb_after_status & _snb_not_writeoff
            _snb_after_class = _snb_after_category & _snb_class
            _snb_active_date = _snb_after_class & (due_date == today)
            _snb_broken_date = _snb_after_class & (due_date < today)
            _snb_negative_delta = (last_date - due_date).dt.days < 0
            _snb_diag = st.session_state['promise_filter_debug'][bank]
            _snb_diag['active'] = {'steps': [
                {'الخطوة': 'إجمالي الصفوف الأصلية', 'العدد': int(len(df_promises))},
                {'الخطوة': 'بعد حالة المعالجة = غير معالج', 'العدد': int(_snb_after_status.sum())},
                {'الخطوة': 'بعد استبعاد فئة writeoff', 'العدد': int(_snb_after_category.sum())},
                {'الخطوة': 'بعد Classification = وعود السداد قيد التنفيذ',
                 'العدد': int(_snb_after_class.sum())},
                {'الخطوة': 'بعد Follow up Due Date = اليوم', 'العدد': int(_snb_active_date.sum())},
                {'الخطوة': 'بعد استبعاد Follow up Last Date = اليوم',
                 'العدد': int((_snb_active_date & (last_date != today)).sum())},
            ]}
            _snb_diag['broken'] = {'steps': [
                {'الخطوة': 'إجمالي الصفوف الأصلية', 'العدد': int(len(df_promises))},
                {'الخطوة': 'بعد حالة المعالجة = غير معالج', 'العدد': int(_snb_after_status.sum())},
                {'الخطوة': 'بعد استبعاد فئة writeoff', 'العدد': int(_snb_after_category.sum())},
                {'الخطوة': 'بعد Classification = وعود السداد قيد التنفيذ',
                 'العدد': int(_snb_after_class.sum())},
                {'الخطوة': 'بعد Follow up Due Date قبل اليوم', 'العدد': int(_snb_broken_date.sum())},
                {'الخطوة': 'بعد فرق التاريخ السالب (Last Date - Due Date)',
                 'العدد': int((_snb_broken_date & _snb_negative_delta).sum())},
            ]}
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

    # STC active-promises logic:
    # 1) exclude rows whose processing status is exactly "معالج"
    # 2) keep monthly installments / positive / promise-in-progress classifications
    # 3) fill blank Expected Closing from Follow up Due Date
    # 4) keep Expected Closing = today
    # 5) exclude rows whose Follow up Last Date = today
    if bank == "STC" and not df_promises.empty:
        stc_processing_col = find_col_in(
            df_promises,
            ['حالة المعالجة - التمويل', 'حالة المعالجة', 'processing status', 'processing_status']
        )
        stc_classification_col = find_col_in(df_promises, ['Classification', 'classification'])
        stc_expected_col = find_col_in(df_promises, ['Expected Closing', 'expected closing', 'Expected_Closing'])
        stc_due_col = find_col_in(df_promises, ['Follow up Due Date', 'follow up due date'])
        stc_last_col = find_col_in(df_promises, ['Follow up Last Date', 'follow up last date'])

        if all([stc_processing_col, stc_classification_col, stc_expected_col, stc_due_col, stc_last_col]):
            stc_promises = df_promises.copy()
            stc_debug_source = stc_promises.copy()

            def _normalize_stc_filter_text(value):
                text = '' if pd.isna(value) else str(value).strip()
                text = re.sub(r'\s+', ' ', text)
                return (
                    text
                    .replace('أ', 'ا')
                    .replace('إ', 'ا')
                    .replace('آ', 'ا')
                    .replace('ى', 'ي')
                    .replace('ة', 'ه')
                )

            def _stc_blank_mask(series):
                cleaned = series.astype(str).str.strip().str.lower()
                return series.isna() | cleaned.isin(['', 'nan', 'nat', 'none'])

            # Expected Closing: fill ONLY blank cells from Follow up Due Date.
            # Existing dates remain unchanged.
            stc_expected_blank = _stc_blank_mask(stc_promises[stc_expected_col])
            stc_promises.loc[stc_expected_blank, stc_expected_col] = (
                stc_promises.loc[stc_expected_blank, stc_due_col]
            )

            # Follow up Last Date: exclude blank rows completely before any
            # STC active-promises filtering.
            stc_last_blank = _stc_blank_mask(stc_promises[stc_last_col])
            stc_promises = stc_promises.loc[~stc_last_blank].copy()

            stc_due_date = pd.to_datetime(
                stc_promises[stc_due_col], errors='coerce'
            ).dt.normalize()
            stc_expected_date = pd.to_datetime(
                stc_promises[stc_expected_col], errors='coerce'
            ).dt.normalize()
            stc_last_date = pd.to_datetime(
                stc_promises[stc_last_col], errors='coerce'
            ).dt.normalize()

            stc_processing = stc_promises[stc_processing_col].apply(
                _normalize_stc_filter_text
            )
            stc_classification = stc_promises[stc_classification_col].apply(
                _normalize_stc_filter_text
            )

            stc_allowed_classifications = {
                _normalize_stc_filter_text('اقساط شهريه'),
                _normalize_stc_filter_text('ايجابي'),
                _normalize_stc_filter_text('وعود سداد قيد التنفيذ'),
                _normalize_stc_filter_text('وعود السداد قيد التنفيذ'),
                _normalize_stc_filter_text('وعد سداد قيد التنفيذ'),
            }

            # Diagnostic trace for STC active promises. This does not change
            # the result; it only shows how many rows survive each filter.
            _dbg = stc_debug_source.copy()
            _dbg_processing = _dbg[stc_processing_col].apply(_normalize_stc_filter_text)
            _dbg_classification = _dbg[stc_classification_col].apply(_normalize_stc_filter_text)

            _dbg_expected_blank = _stc_blank_mask(_dbg[stc_expected_col])
            _dbg_expected_filled = _dbg[stc_expected_col].copy()
            _dbg_expected_filled.loc[_dbg_expected_blank] = _dbg.loc[_dbg_expected_blank, stc_due_col]
            _dbg_expected_date = pd.to_datetime(_dbg_expected_filled, errors='coerce').dt.normalize()

            _dbg_last_blank = _stc_blank_mask(_dbg[stc_last_col])
            _dbg_last_date = pd.to_datetime(_dbg[stc_last_col], errors='coerce').dt.normalize()

            _dbg_not_processed = _dbg_processing != _normalize_stc_filter_text('معالج')
            _dbg_allowed_class = _dbg_classification.isin(stc_allowed_classifications)
            _dbg_after_processing = _dbg_not_processed
            _dbg_after_classification = _dbg_after_processing & _dbg_allowed_class
            _dbg_after_last_not_blank = _dbg_after_classification & (~_dbg_last_blank)
            _dbg_after_expected_today = _dbg_after_last_not_blank & (_dbg_expected_date == today)
            _dbg_final = _dbg_after_expected_today & (_dbg_last_date != today)

            _dbg_steps = [
                {'الخطوة': 'إجمالي الصفوف الأصلية', 'العدد': int(len(_dbg))},
                {'الخطوة': 'بعد استبعاد حالة المعالجة = معالج', 'العدد': int(_dbg_after_processing.sum())},
                {'الخطوة': 'بعد فلتر Classification', 'العدد': int(_dbg_after_classification.sum())},
                {'الخطوة': 'Expected Closing blank تم ملؤه من Due Date', 'العدد': int((_dbg_after_classification & _dbg_expected_blank).sum())},
                {'الخطوة': 'بعد استبعاد Follow up Last Date الـ blank', 'العدد': int(_dbg_after_last_not_blank.sum())},
                {'الخطوة': 'بعد Expected Closing = اليوم', 'العدد': int(_dbg_after_expected_today.sum())},
                {'الخطوة': 'بعد استبعاد Follow up Last Date = اليوم', 'العدد': int(_dbg_final.sum())},
            ]

            _dbg_class_counts = (
                _dbg.loc[_dbg_after_processing, stc_classification_col]
                .fillna('(blank)')
                .astype(str)
                .str.strip()
                .replace('', '(blank)')
                .value_counts(dropna=False)
                .rename_axis('Classification')
                .reset_index(name='العدد')
            )

            st.session_state['stc_active_promises_debug'] = {
                'steps': _dbg_steps,
                'classification_counts': _dbg_class_counts.to_dict('records'),
            }
            st.session_state['promise_filter_debug'][bank]['active'] = (
                st.session_state['stc_active_promises_debug']
            )


            stc_base_condition = (
                (stc_processing != _normalize_stc_filter_text('معالج')) &
                (stc_classification.isin(stc_allowed_classifications))
            )

            active_promises_df = stc_promises[
                stc_base_condition &
                (stc_expected_date == today) &
                (stc_last_date != today)
            ].copy()

            # STC broken-promises logic:
            # 1) exclude processing status = معالج
            # 2) keep only the three requested Classification values
            # 3) Expected Closing blanks were already filled from Follow up Due Date
            # 4) keep Expected Closing dates before today
            # 5) calculate Follow up Last Date - Expected Closing
            # 6) keep only negative differences
            stc_broken_allowed_classifications = {
                _normalize_stc_filter_text('اقساط شهريه'),
                _normalize_stc_filter_text('ايجابي'),
                _normalize_stc_filter_text('وعود السداد قيد التنفيذ'),
            }

            stc_broken_condition = (
                (stc_processing != _normalize_stc_filter_text('معالج')) &
                (stc_classification.isin(stc_broken_allowed_classifications)) &
                (stc_expected_date < today)
            )

            # Diagnostic steps use STC's own broken-promise conditions.
            _dbg_broken_status = (
                _dbg_processing != _normalize_stc_filter_text('معالج')
            )
            _dbg_broken_class = _dbg_classification.isin(
                stc_broken_allowed_classifications
            )
            _dbg_broken_after_class = _dbg_broken_status & _dbg_broken_class
            _dbg_broken_has_last = _dbg_broken_after_class & (~_dbg_last_blank)
            _dbg_broken_before = _dbg_broken_has_last & (_dbg_expected_date < today)
            _dbg_broken_negative = (
                (_dbg_last_date - _dbg_expected_date).dt.days < 0
            )
            st.session_state['promise_filter_debug'][bank]['broken'] = {
                'steps': [
                    {'الخطوة': 'إجمالي الصفوف الأصلية',
                     'العدد': int(len(_dbg))},
                    {'الخطوة': 'بعد استبعاد حالة المعالجة = معالج',
                     'العدد': int(_dbg_broken_status.sum())},
                    {'الخطوة': 'بعد فلتر Classification للوعود المكسورة',
                     'العدد': int(_dbg_broken_after_class.sum())},
                    {'الخطوة': 'Expected Closing blank تم ملؤه من Due Date',
                     'العدد': int((_dbg_broken_after_class & _dbg_expected_blank).sum())},
                    {'الخطوة': 'بعد استبعاد Follow up Last Date الفارغ',
                     'العدد': int(_dbg_broken_has_last.sum())},
                    {'الخطوة': 'بعد Expected Closing قبل اليوم',
                     'العدد': int(_dbg_broken_before.sum())},
                    {'الخطوة': 'بعد فرق التاريخ السالب (Last Date - Expected Closing)',
                     'العدد': int((_dbg_broken_before & _dbg_broken_negative).sum())},
                ]
            }

            broken_promises_df = stc_promises[stc_broken_condition].copy()
            if not broken_promises_df.empty:
                stc_broken_diff = (
                    stc_last_date.loc[broken_promises_df.index] -
                    stc_expected_date.loc[broken_promises_df.index]
                ).dt.days

                negative_diff_mask = stc_broken_diff < 0
                broken_promises_df = broken_promises_df.loc[
                    negative_diff_mask[negative_diff_mask].index
                ].copy()
                stc_broken_diff = stc_broken_diff.loc[broken_promises_df.index]

                broken_promises_df = broken_promises_df.drop(
                    columns=['الفرق '],
                    errors='ignore'
                )
                insert_at = (
                    broken_promises_df.columns.get_loc(stc_last_col) + 1
                    if stc_last_col in broken_promises_df.columns
                    else len(broken_promises_df.columns)
                )
                broken_promises_df.insert(
                    insert_at,
                    'الفرق ',
                    stc_broken_diff.astype(int)
                )

            total_unhandled = int(stc_base_condition.sum())

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
if curr_page == 'Overview':
    st.markdown("""<style>
[data-testid="stAppViewContainer"]{background:#f1f3f2!important}
[data-testid="stMainBlockContainer"]{max-width:1460px;margin:auto}
[data-testid="stVerticalBlockBorderWrapper"],[data-testid="stPlotlyChart"]{background:white;border-radius:22px;padding:10px}
.kpi-grid-4{display:grid!important;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:16px;direction:rtl}
.kpi-grid-4 .kpi-card{padding:22px!important;border:0!important;border-radius:22px!important;background:white!important;box-shadow:0 3px 16px #17261f09!important;text-align:right}
.kpi-grid-4 .kpi-card:first-child{background:#116044!important}
.kpi-grid-4 .kpi-card:first-child *{color:white!important}
.kpi-grid-4 .kpi-value{font-size:34px!important}
.overview-title{direction:rtl;text-align:right;color:#162d25;font-size:32px;font-weight:750;margin:10px 0 14px}
</style><div class="overview-title">نظرة عامة على التحصيل</div>""", unsafe_allow_html=True)
    overview_snb_col, overview_stc_col, _overview_bank_spacer = st.columns([0.82, 0.82, 8.36], gap='small')
    with overview_snb_col:
        st.button(
            '🏦 SNB',
            key='overview_choose_SNB',
            type='primary' if selected_bank == 'SNB' else 'secondary',
            use_container_width=False,
            on_click=set_bank,
            args=('SNB',)
        )
        if selected_bank == 'SNB':
            st.markdown(
                f'<div style="color:{snb_green};font-size:19px;font-weight:900;'
                'width:116px;text-align:center;direction:ltr;white-space:nowrap;'
                'margin-top:-12px;">(30-60-NPL)</div>',
                unsafe_allow_html=True
            )
    with overview_stc_col:
        st.button(
            '🏦 STC',
            key='overview_choose_STC',
            type='primary' if selected_bank == 'STC' else 'secondary',
            use_container_width=False,
            on_click=set_bank,
            args=('STC',)
        )

    # Widgets live here only once. Their screen position is the earlier header.
    with _overview_upload_header:
        overview_versions = st.session_state.overview_uploader_versions[selected_bank]
        # Three separate, equal-size upload panels aligned right in RTL.
        upload_cols = st.columns([1, 1, 1, 2.7], gap='small')
    
        with upload_cols[0]:
            with st.container(border=True, key='overview_activity_upload_card'):
                st.markdown(
                    f"<div style='text-align:center;font-weight:900;color:{label_color};"
                    "font-size:14px;margin:0 0 3px;'>📈 1) تقرير النشاط</div>"
                    "",
                    unsafe_allow_html=True
                )
                activity_widget_files = st.file_uploader(
                    'ارفعي ملف تقرير النشاط في هذه الخانة',
                    type=['xlsx', 'csv'],
                    label_visibility='collapsed',
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
    
        with upload_cols[1]:
            with st.container(border=True, key='overview_model_upload_card'):
                st.markdown(
                    f"<div style='text-align:center;font-weight:900;color:{label_color};"
                    "font-size:14px;margin:0 0 3px;'>🧩 2) موديل النشاط</div>"
                    "",
                    unsafe_allow_html=True
                )
                activity_model_file = st.file_uploader(
                    'ارفعي ملف موديل النشاط في هذه الخانة',
                    type=['xlsx'],
                    label_visibility='collapsed',
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
        with upload_cols[2]:
            with st.container(border=True, key='overview_promises_upload_card'):
                st.markdown(f"<div style='text-align:center;font-weight:900;color:{label_color};font-size:14px;margin-bottom:3px;'>ارفعي المحفظة</div>", unsafe_allow_html=True)
                promises_widget_files = st.file_uploader(
                    'ارفعي تقرير الوعود والإهمال هنا',
                    type=['xlsx', 'csv'],
                    label_visibility='collapsed',
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

# Process Promise data once before rendering the header KPIs.
# Reuse the same result below; no duplicate parsing or filtering.
_promise_header_result = None
if curr_page == 'Promises' and bank_files.get('Promises'):
    _promise_header_result = process_page_data(
        bank_files['Promises'], bank=selected_bank
    )

if curr_page in ('Promises', 'Neglect'):
    _promise_metrics = None
    if curr_page == 'Promises' and _promise_header_result is not None:
        _promise_metrics = (
            _promise_header_result[5],
            _promise_header_result[7],
            _promise_header_result[8],
        )
    render_report_upload_bar(curr_page, promise_metrics=_promise_metrics)

# Activity is intentionally a blank design canvas until the new report is specified.
# Keep other pages and their upload state completely unchanged.
page_uploaded = None if curr_page == 'Activity' else bank_files[curr_page]

if page_uploaded is None:
    if curr_page == 'Home':
        st.markdown(f"""
        <div style="background-color:{card_color};border:1px solid {border_color};border-radius:15px;padding:42px 24px;margin:28px 0 18px;text-align:center;direction:rtl;box-shadow:0 4px 14px rgba(0,0,0,0.08);">
            <h2 style="color:{text_color};margin:0 0 12px;font-weight:900;font-size:30px;">أهلاً بيكي يا رنا في تقارير شركة إجادة لتحصيل الديون</h2>
            <p style="color:{kpi_title_color};margin:0 0 10px;font-weight:800;font-size:18px;line-height:1.7;">تابعي أداء التحصيل والوعود والنشاط بسهولة من مكان واحد.</p>
            <p style="color:{snb_green};margin:0;font-weight:900;font-size:17px;line-height:1.7;">🏦 اختاري البنك لعرض التقارير الخاصة به.</p>
        </div>
        """, unsafe_allow_html=True)
    elif curr_page == 'Overview':
        st.markdown(f"""
        <div style="background-color:{card_color};border:1px solid {border_color};border-radius:15px;padding:22px 18px;margin:8px 0 18px;text-align:center;direction:rtl;">
            <h2 style="color:{text_color};margin:0 0 8px;font-weight:900;">📊 لوحة التحكم الرئيسية | EJADA Overview</h2>
            <p style="color:{kpi_title_color};margin:0;font-weight:700;">اختاري البنك، ثم ارفعي تقرير النشاط وتقرير الوعود والإهمال لبناء لوحة المؤشرات.</p>
        </div>
        """, unsafe_allow_html=True)
    elif curr_page == 'Activity':
        # Put file uploader, title and compact activity-type selector on ONE row.
        activity_error_slot = st.empty()
        _snb_secret_model_ready = (
            selected_bank == "SNB" and bool(_snb_secret("SNB_MODEL_ENCRYPTED"))
        )
        with st.container(key="snb_activity_header_row"):
            # Rightmost uploader, compact adjacent activity type, free space left.
            activity_col, kind_col, dashboard_col, _left_space = st.columns(
                [1.8, 1.15, 1.25, 5.4], gap="small",
                vertical_alignment="center"
            )
            with activity_col:
                with st.container(border=True, key="activity_file_upload_card"):
                    st.markdown(
                        f'<div style="color:{label_color};direction:rtl;text-align:center;'
                        'font-size:18px;font-weight:900;margin-bottom:8px;">'
                        'رفع ملف النشاط</div>',
                        unsafe_allow_html=True
                    )
                    activity_key = (
                        f"activity_data_file_{selected_bank}_"
                        f"{st.session_state.activity_upload_versions[selected_bank]['activity']}"
                    )
                    st.file_uploader(
                        "رفع ملف النشاط", type=["xlsx", "csv"],
                        accept_multiple_files=False,
                        label_visibility="collapsed", key=activity_key,
                        on_change=save_activity_page_upload,
                        args=(selected_bank, "activity", activity_key)
                    )
            with kind_col:
                with st.container(key="snb_activity_kind_compact"):
                    activity_kind = st.selectbox(
                        "نوع النشاط", options=["تلقائي", "يومي", "أسبوعي", "شهري"],
                        key="snb_activity_kind_auto"
                    )
            with dashboard_col:
                is_open = _ejada_activity_dashboard_is_open(selected_bank)
                st.button(
                    "📊 إخفاء الداشبورد" if is_open else "📊 عرض الداشبورد",
                    key=f"activity_dashboard_toggle_{selected_bank}",
                    help="فتح وإغلاق لوحة مؤشرات النشاط والشارتات القادمة",
                    use_container_width=True,
                    type="primary" if is_open else "secondary",
                    on_click=_ejada_toggle_activity_dashboard,
                    args=(selected_bank,),
                )

        # Legacy fallback: no encrypted model configured, allow a model upload
        # below the main header without moving the centred title.
        if not _snb_secret_model_ready:
            model_col, _model_space = st.columns([1.8, 7.8], gap="small")
            with model_col:
                with st.container(border=True, key="activity_model_upload_card"):
                    st.markdown(
                        f'<div style="color:{label_color};direction:rtl;'
                        'text-align:center;font-size:18px;font-weight:900;'
                        'margin-bottom:8px;">خانة رفع موديل النشاط</div>',
                        unsafe_allow_html=True
                    )
                    model_key = (
                        f"activity_model_file_{selected_bank}_"
                        f"{st.session_state.activity_upload_versions[selected_bank]['model']}"
                    )
                    st.file_uploader(
                        "خانة رفع موديل النشاط", type=["xlsx"],
                        accept_multiple_files=False,
                        label_visibility="collapsed", key=model_key,
                        on_change=save_activity_page_upload,
                        args=(selected_bank, "model", model_key)
                    )

        # Bank KPIs render only when the dashboard button is open.
        if selected_bank == "STC" and _ejada_activity_dashboard_is_open("STC"):
            _stc_file = st.session_state.activity_page_uploads["STC"]["activity"]
            try:
                _stc_metrics = _ejada_stc_activity_kpi_counts(_stc_file)
                _ejada_activity_kpis("STC", *_stc_metrics)
                if _stc_file is not None and _stc_metrics[1] is None:
                    st.caption(
                        "ملف STC لا يحتوي على عمود تقييم مكالمة صريح؛ "
                        "المغطاة محسوبة، أما الناجحة وغير الناجحة فتظهران "
                        "بعلامة — بدل تخمين التصنيف."
                    )
            except Exception as exc:
                st.error(f"تعذر قراءة بيانات نشاط STC: {exc}")
                _ejada_activity_kpis("STC", 0)
        elif (
            selected_bank == "SNB"
            and _ejada_activity_dashboard_is_open("SNB")
            and not _snb_secret_model_ready
            and st.session_state.activity_page_uploads["SNB"]["model"] is None
        ):
            # No baseline model available yet: count coverage, but never guess.
            _waiting = st.session_state.activity_page_uploads["SNB"]["activity"]
            if _waiting is None:
                _ejada_activity_kpis("SNB", 0, 0, 0)
            else:
                try:
                    _waiting_data, _ = _prepare_snb_activity_file(_waiting)
                    _ejada_activity_kpis(
                        "SNB", len(_waiting_data), None, None, len(_waiting_data)
                    )
                    st.caption(
                        "حمّلي موديل النشاط لتحديد الناجحة وغير الناجحة "
                        "بدون تخمين."
                    )
                except Exception as exc:
                    st.warning(f"تعذر قراءة ملف النشاط لعرض المؤشرات: {exc}")
                    _ejada_activity_kpis("SNB", 0, None, None)

        # SNB report is rendered only when an activity file and model exist.
        if selected_bank == 'SNB':
            _uploaded = st.session_state.activity_page_uploads['SNB']
            if (
                (_snb_secret_model_ready or _uploaded['model'] is not None)
                and _uploaded['activity'] is not None
            ):
                st.session_state.snb_activity_screen = 'report'
            if st.session_state.get('snb_activity_screen') == 'report':
                render_snb_activity_report(
                    _uploaded['model'], _uploaded['activity'],
                    activity_kind=activity_kind,
                    error_slot=activity_error_slot
                )
    else:
        st.markdown(f"<p style='color:{label_color};font-weight:800;font-size:15px;text-align:right;direction:rtl;margin-top:8px;'>ارفعي الملف من الخانة اللي فوق لعرض التقرير.</p>", unsafe_allow_html=True)
else:
    if curr_page == 'Overview':
        active_promises_df, broken_promises_df, neglect_df, activity_base_df, report_bytes, total_rows, total_unhandled, active_count, broken_count, neglect_count, total_activity_count = process_page_data(
            bank=selected_bank,
            activity_files=st.session_state.overview_uploads[selected_bank]['Activity'],
            promises_files=st.session_state.overview_uploads[selected_bank]['Promises']
        )
    elif curr_page == 'Promises' and _promise_header_result is not None:
        (active_promises_df, broken_promises_df, neglect_df,
         activity_base_df, report_bytes, total_rows, total_unhandled,
         active_count, broken_count, neglect_count,
         total_activity_count) = _promise_header_result
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
                <div class="kpi-title">إجمالي الوعود</div>
                <div class="kpi-value">{active_count + broken_count:,}</div>
                <small>القائمة {active_count:,} · المكسورة {broken_count:,}</small>
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

        overview_row1 = st.columns([1.6, 1], gap='large')

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
        tab_active, tab_broken = st.tabs([f"✅ الوعود القائمة ({active_count:,})", f"⚠️ الوعود المكسورة ({broken_count:,})"])
        
        with tab_active:
            # Diagnostics and sheet details are displayed together below the KPIs.

            st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

            # Preserve SNB's original Pivot summary; STC has no active Pivot.
            if selected_bank != 'STC':
                st.markdown(
                    f"<h4 style='text-align:right; direction:rtl; color:{label_color}; font-weight:900;'>📊 ملخص تقرير الوعود القائمة (Pivot)</h4>",
                    unsafe_allow_html=True
                )
                st.dataframe(
                    style_pivot_frame(pivot_to_frame(build_pivot(active_promises_df, order='first'))),
                    use_container_width=True,
                    hide_index=True
                )

            # Matching horizontal cards: diagnostic on the right, sheet on the left.
            with st.container(key=f"promise_details_row_{selected_bank}"):
                diag_col, details_col = st.columns([3.1, 1], gap='small')
                with diag_col:
                    render_promise_filter_diagnostics(selected_bank, 'active')
                with details_col:
                    with st.container(key=f"active_sheet_detail_{selected_bank}"):
                        with st.expander(
                            "📋 تفاصيل شيت الوعود القائمة",
                            expanded=(selected_bank != 'STC')
                        ):
                            _expander_table(active_promises_df, "active_promises_df")
            
        with tab_broken:
            render_promise_filter_diagnostics(selected_bank, 'broken')
            # في STC نخفي ملخص الـ Pivot الخاص بالوعود المكسورة.
            if selected_bank != 'STC':
                st.markdown(
                    f"<h4 style='text-align:right; direction:rtl; color:{label_color}; font-weight:900;'>📊 ملخص تقرير الوعود المكسورة (Pivot)</h4>",
                    unsafe_allow_html=True
                )
                st.dataframe(
                    style_pivot_frame(pivot_to_frame(build_pivot(broken_promises_df, order='desc'))),
                    use_container_width=True,
                    hide_index=True
                )

            st.markdown(
                f"<h4 style='text-align:right; direction:rtl; color:{label_color}; font-weight:900; margin-top:20px;'>📋 تفاصيل شيت الوعود المكسورة</h4>",
                unsafe_allow_html=True
            )
            st.dataframe(broken_promises_df, use_container_width=True, height=400)
            
        st.markdown("<br>", unsafe_allow_html=True)
        if selected_bank == 'STC':
            stc_active_report = build_single_promise_report_xlsx(
                active_promises_df,
                report_kind='active',
                bank='STC'
            ).getvalue()
            stc_broken_report = build_single_promise_report_xlsx(
                broken_promises_df,
                report_kind='broken',
                bank='STC'
            ).getvalue()

            with st.container(key="promise_downloads_STC"):
                # Full-width RTL container: buttons stacked at the far right.
                st.download_button(
                    "📥 تحميل تقرير الوعود القائمة STC.xlsx",
                    data=stc_active_report,
                    file_name="STC_Active_Promises_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_stc_active_promises",
                    use_container_width=False
                )
                st.download_button(
                    "📥 تحميل تقرير الوعود المكسورة STC.xlsx",
                    data=stc_broken_report,
                    file_name="STC_Broken_Promises_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_stc_broken_promises",
                    use_container_width=False
                )
        else:
            with st.container(key="promise_downloads_SNB"):
                _left_spacer, center_download, _right_spacer = st.columns(
                    [1, 2.2, 1], gap='medium'
                )
                with center_download:
                    st.download_button(
                        "📥 تحميل التقرير الشامل (الإكسيل بالملخصات).xlsx",
                        data=report_bytes,
                        file_name=f"{selected_bank}_Daily_Promises_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_promises_all_{selected_bank}",
                        use_container_width=False
                    )

    elif curr_page == 'Neglect':
        default_sub_states = ['وعد بسداد مبلغ المعالجة', 'واعد بسداد مبلغ المعالجه', 'رافض السداد', 'الجوال مغلق', 'مماطل', 'تم سداد جزء وليس مبلغ المعالجه', 'لا يجيب', 'متقاعد الراتب علي بنك آخر', 'يرغب بجدوله المديونيه']
        selected_sub_states = st.multiselect("فلاتر الـ Sub State المشمولة:", options=default_sub_states, default=default_sub_states, key=f"neglect_states_{selected_bank}")
        sub_col = 'Sub State' if 'Sub State' in neglect_df.columns else 'sub state' if 'sub state' in neglect_df.columns else None
        final_neglect_df = neglect_df[neglect_df[sub_col].isin(selected_sub_states)] if sub_col else neglect_df
        st.markdown(f"<p style='text-align:right; direction:rtl; color:{label_color}; font-weight: bold;'>العدد الإجمالي: <b style='color:#586F8C; font-size:18px;'>{len(final_neglect_df):,}</b> حساب مهمل</p>", unsafe_allow_html=True)
        st.dataframe(final_neglect_df, use_container_width=True, height=450)