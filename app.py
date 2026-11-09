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
    st.rerun()

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
