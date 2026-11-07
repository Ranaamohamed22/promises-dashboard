import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date
import io

# ==========================================
# إعدادات الصفحة الأساسية
# ==========================================
st.set_page_config(page_title="SNB Promises Dashboard", layout="wide", initial_sidebar_state="expanded")

# حفظ الصفحة الحالية والثيم في الذاكرة
if 'current_page' not in st.session_state:
    st.session_state.current_page = 'Dashboard'
if 'theme_choice' not in st.session_state:
    st.session_state.theme_choice = 'التصميم الأصلي (Dark)'

# ==========================================
# إعدادات الثيم والألوان بناءً على الاختيار
# ==========================================
if st.session_state.theme_choice == "البنك الأهلي السعودي (Light)":
    bg_color = "#F4F7F6"
    card_color = "#FFFFFF"
    border_color = "#D1E0D7"
    primary_color = "#2CA549" # الأخضر الفاتح للبنك
    primary_text = "#FFFFFF"
    text_color = "#111D16" # أسود/أخضر غامق للنصوص
    sub_text_color = "#666666"
    download_btn = "#004B36" # الأخضر الغامق للبنك
    bar_colors = ['#2CA549', '#F2C8ED', '#FF6B6B', '#20AEF3']
    gauge_bg = "#E8F0EA"
    
    # ألوان القائمة الجانبية في الفاتح
    sidebar_bg = "#FFFFFF"
    sidebar_border = "#D1E0D7"
    sidebar_text = "#111D16"
    sidebar_title_bg = "#F4F7F6"
else:
    bg_color = "#171821"
    card_color = "#21222D"
    border_color = "#2C2D3A"
    primary_color = "#A9DFD8" # السماوي الأصلي
    primary_text = "#171821"
    text_color = "#FFFFFF" # أبيض للنصوص
    sub_text_color = "#A0A0A0"
    download_btn = "#A9DFD8"
    bar_colors = ['#A9DFD8', '#F2C8ED', '#FF6B6B', '#20AEF3']
    gauge_bg = "#2C2D3A"
    
    # ألوان القائمة الجانبية في الغامق
    sidebar_bg = "#171821"
    sidebar_border = "#2C2D3A"
    sidebar_text = "#E0E0E0"
    sidebar_title_bg = "#21222D"

# ==========================================
# شريط البحث العلوي
# ==========================================
col_search, col_space, col_dash, col_theme_btn, col_profile = st.columns([7.5, 0.5, 1.2, 0.4, 0.6], gap="small")

with col_search:
    search_query = st.text_input("", placeholder="🔍 ابحثي هنا (برقم العميل، الاسم، أو أي بيان)...", label_visibility="collapsed")

with col_space:
    st.empty() 

with col_dash:
    if st.button("🏠 Dashboard", key="top_dashboard_btn", use_container_width=True):
        st.session_state.current_page = 'Dashboard'
        st.rerun()

with col_theme_btn:
    current_icon = "🌗" if st.session_state.theme_choice == 'التصميم الأصلي (Dark)' else "🎨"
    if st.button(current_icon, key="theme_toggle"):
        if st.session_state.theme_choice == 'التصميم الأصلي (Dark)':
            st.session_state.theme_choice = 'البنك الأهلي السعودي (Light)'
        else:
            st.session_state.theme_choice = 'التصميم الأصلي (Dark)'
        st.rerun()

with col_profile:
    st.markdown(f"<div style='text-align: right; padding-top: 5px; font-size: 18px; color:{text_color};'>👩‍💻 ⌄</div>", unsafe_allow_html=True)
