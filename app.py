import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date, datetime, time
import io
import base64
import os

# ==========================================
# إعدادات الصفحة الأساسية
# ==========================================
st.set_page_config(page_title="SNB Promises Dashboard", layout="wide", initial_sidebar_state="collapsed")

if 'current_page' not in st.session_state:
    st.session_state.current_page = 'Overview'
if 'theme_choice' not in st.session_state:
    st.session_state.theme_choice = 'ثيم إجادة (Navy)'

# ==========================================
# إعدادات الثيم والألوان بالترتيب المطلوب (بدون أخضر)
# ==========================================
ejada_navy_dark = "#1F4373"
ejada_navy_light = "#304D73"
ejada_grey_blue = "#586F8C"
ejada_border_light = "#99AABF"
ejada_bg_light = "#F2F2F2"
snb_green = "#2CA549"  # يُستخدم فقط لشعار SNB والأزرار النشطة الأساسية

if st.session_state.theme_choice == "ثيم إجادة (Navy)":
    bg_color = ejada_navy_dark 
    card_color = "#FFFFFF" 
    border_color = ejada_grey_blue
    primary_color = ejada_navy_light 
    primary_text = "#FFFFFF"
    text_color = ejada_navy_dark
    sub_text_color = ejada_grey_blue
    gauge_bg = "#304D73"
    nav_btn_bg = "#FFFFFF"
    nav_btn_text = ejada_navy_dark
    nav_btn_border = ejada_border_light
    nav_btn_hover_border = ejada_grey_blue
    kpi_title_color = "#586F8C"
    header_logo_color = "#FFFFFF"
    bar_chart_text = "#FFFFFF"
    # ترتيب الألوان الجديد للشارتات بناءً على طلبك (بدون أخضر)
    chart_colors = ["#1F4373", "#304D73", "#586F8C", "#99AABF", "#F2F2F2"]
else:
    bg_color = "#13131A"
    card_color = "#1C1C24"
    border_color = "#2D2D3B"
    primary_color = ejada_navy_light
    primary_text = "#FFFFFF"
    text_color = "#FFFFFF" 
    sub_text_color = "#A0AEC0"
    gauge_bg = "#2D2D3B"
    nav_btn_bg = "#1A1A1D"
    nav_btn_text = "#E0E0E0"
    nav_btn_border = "#333333"
    nav_btn_hover_border = ejada_grey_blue
    kpi_title_color = "#A0AEC0"
    header_logo_color = "#FFFFFF"
    bar_chart_text = "#FFFFFF"
    chart_colors = ["#1F4373", "#304D73", "#586F8C", "#99AABF", "#F2F2F2"]

# ==========================================
# تصميم الـ UI الديناميكي (CSS)
# ==========================================
st.markdown(f"""
<style>
    [data-testid="collapsedControl"] {{display: none !important;}}
    [data-testid="stSidebar"] {{display: none !important;}}
    
    .stApp {{ background-color: {bg_color} !important; }}

    .block-container {{ padding-top: 3.8rem !important; padding-left: 1rem !important; padding-right: 1rem !important; max-width: 100% !important; }}
    
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
    div[data-testid="stButton"] > button[kind="secondary"]:hover, 
    div[data-testid="stButton"] > button[kind="secondary"]:focus {{ border-color: {ejada_grey_blue} !important; color: {ejada_navy_light} !important; }}
    
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
    div[data-testid="stButton"] > button[kind="primary"]:focus {{ box-shadow: none !important; }}

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
    div.st-key-theme_toggle div[data-testid="stButton"] > button:hover {{ transform: scale(1.1) !important; }}

    .stTextInput input {{ background-color: {nav_btn_bg} !important; color: {text_color} !important; border: 1px solid {border_color} !important; border-radius: 8px !important; }}
    .stTextInput input::placeholder {{ color: {sub_text_color} !important; }}
    [data-testid="stFileUploader"] > section {{ background-color: {nav_btn_bg} !important; border: 1px dashed {border_color} !important; border-radius: 8px !important; }}
    [data-testid="stFileUploader"] section * {{ color: {text_color} !important; }}

    .kpi-card {{ background-color: {nav_btn_bg}; border-radius: 12px; padding: 20px; box-shadow: 0 4px 10px rgba(0,0,0,0.2); border: 1px solid {border_color}; border-left-width: 6px; border-left-style: solid; text-align: left; direction: ltr; position: relative; }}
    .kpi-title {{ color: {kpi_title_color}; font-size: 13px; text-transform: uppercase; font-weight: 800; margin-bottom: 10px; }}
    .kpi-value {{ color: {text_color}; font-size: 30px; font-weight: 900; margin-bottom: 5px; }}
    .kpi-sub {{ color: {sub_text_color}; font-size: 12px; font-weight: 600; }}
    .kpi-icon {{ position: absolute; top: 20px; right: 20px; font-size: 24px; opacity: 0.9; }}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 0. الصف العلوي: القمر + Overview
# ==========================================
top_cols = st.columns([0.4, 1.3, 8], gap="small")

with top_cols[0]:
    if st.button(" ", key="theme_toggle"): 
        if st.session_state.theme_choice == 'التصميم الأصلي (Dark)':
            st.session_state.theme_choice = 'ثيم إجادة (Navy)'
        else:
            st.session_state.theme_choice = 'التصميم الأصلي (Dark)'
        st.rerun()

with top_cols[1]:
    if st.button("📊 Overview", type="primary" if st.session_state.current_page == 'Overview' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Overview'
        st.rerun()

# ==========================================
# 1. الهيدر الرئيسي (إجادة + SNB)
# ==========================================
st.markdown(f"""
<div style="position: relative; width: 100%; min-height: 120px; margin-top: -60px; margin-bottom: 25px; pointer-events: none;">
    <div style="text-align: center; line-height: 1.1;">
        <div style="color: {header_logo_color}; font-size: 45px; font-weight: 900; font-family: Tahoma, sans-serif;">إجـــادة</div>
        <div style="color: {header_logo_color}; font-size: 20px; letter-spacing: 10px; font-weight: 900; margin-top: -8px; margin-right:-10px;">EJADA</div>
        <div style="color: {header_logo_color}; font-size: 18px; font-weight: bold; margin-top: 5px;">لتحصيل الديون</div>
    </div>
    <div style="position: absolute; top: 0; right: -15px; text-align: right; line-height: 1.2; direction: ltr;">
        <div style="white-space: nowrap;">
            <span style="color: {snb_green}; font-weight: 900; font-size: 26px; letter-spacing: 1px; vertical-align: middle;">SNB</span>
            <span style="color: {snb_green}; font-size: 22px; font-weight: bold; font-family: Tahoma, sans-serif; margin-left: 8px; vertical-align: middle;">البنك الأهلي السعودي</span>
        </div>
        <div style="color: {snb_green}; font-size: 17px; font-weight: 900; text-align: center; letter-spacing: 1.5px; margin-top: 2px;">
            (30-60-NPL)
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 2. شريط التنقل
# ==========================================
nav_cols = st.columns([1.3, 1.3, 1.3, 1.3, 5], gap="small") 

with nav_cols[0]:
    if st.button("✅ الوعود القائمة", type="primary" if st.session_state.current_page == 'Active' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Active'
        st.rerun()

with nav_cols[1]:
    if st.button("⚠️ الوعود المكسورة", type="primary" if st.session_state.current_page == 'Broken' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Broken'
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

st.markdown(f"<hr style='margin: 10px 0 10px 0; border-color: {border_color}; clear: both;'>", unsafe_allow_html=True)

# ==========================================
# 3. شريط التحكم (رفع الملفات المتعددة والبحث)
# ==========================================
filter_cols = st.columns([3, 1, 6])
with filter_cols[0]:
    uploaded_files = st.file_uploader("", type=["xlsx", "xls", "csv"], accept_multiple_files=True, label_visibility="collapsed")
with filter_cols[2]:
    st.markdown("<div style='padding-top: 15px;'>", unsafe_allow_html=True)
    search_query = st.text_input("", placeholder="🔍 ابحثي هنا (برقم العميل، الاسم، أو أي بيان)...", label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown(f"<hr style='margin: 10px 0 25px 0; border-color: {border_color};'>", unsafe_allow_html=True)

# ==========================================
# 4. معالجة البيانات المجمعة
# ==========================================
if uploaded_files:
    today = pd.Timestamp(date.today()).normalize()
    
    promises_dfs = []
    activity_dfs = []
    all_dfs = []

    for f in uploaded_files:
        try:
            temp_df = pd.read_csv(f) if f.name.endswith('.csv') else pd.read_excel(f)
            temp_df.columns = temp_df.columns.astype(str).str.strip()
            all_dfs.append(temp_df)
            
            if 'Follow up Due Date' in temp_df.columns and 'Follow up Last Date' in temp_df.columns:
                promises_dfs.append(temp_df)
            
            has_created = any(c.lower().strip() in ['created on', 'created_on', 'createdon', 'تاريخ الإنشاء', 'وقت المعالجة', 'date', 'التاريخ'] for c in temp_df.columns)
            if has_created or not ('Follow up Due Date' in temp_df.columns):
                activity_dfs.append(temp_df)
        except Exception as e:
            st.error(f"حدث خطأ في قراءة الملف {f.name}: {e}")

    df_all = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
    df_promises = pd.concat(promises_dfs, ignore_index=True) if promises_dfs else pd.DataFrame()
    df_act_raw = pd.concat(activity_dfs, ignore_index=True) if activity_dfs else df_all.copy()

    def find_col_in(target_df, candidates):
        if target_df.empty:
            return None
        for c in candidates:
            for col in target_df.columns:
                if col.lower().strip() == c.lower().strip():
                    return col
        return None

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
    if len(dedup_cols) == 3:
        activity_base_df = df_act_raw.drop_duplicates(subset=dedup_cols).copy()
    else:
        activity_base_df = df_act_raw.copy()

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
        broken_promises_df['Date_Diff'] = (last_date[broken_promises_df.index] - due_date[broken_promises_df.index]).dt.days
        broken_promises_df = broken_promises_df[broken_promises_df['Date_Diff'] < 0]

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

        if search_query:
            active_promises_df = active_promises_df[active_promises_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)]
            broken_promises_df = broken_promises_df[broken_promises_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)]
            neglect_df = neglect_df[neglect_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)]

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        active_promises_df.to_excel(writer, sheet_name='الوعود القائمة', index=False)
        broken_promises_df.to_excel(writer, sheet_name='الوعود المكسورة', index=False)
        neglect_df.to_excel(writer, sheet_name='تقرير الإهمال', index=False)
    output_buffer.seek(0)

    total_rows = len(df_promises) if not df_promises.empty else len(df_all)
    active_count = len(active_promises_df)
    broken_count = len(broken_promises_df)
    neglect_count = len(neglect_df)
    total_activity_count = len(activity_base_df)

    # ==========================================
    # صفحة الـ Overview الشاملة
    # ==========================================
    if st.session_state.current_page == 'Overview':
        st.markdown(f"<h2 style='text-align:center; margin:0; color:{ejada_navy_light}; font-weight:900;'>📊 لوحة القيادة الشاملة (Executive Overview)</h2>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align:center; color:#FFFFFF; font-weight: bold; margin-bottom: 30px;'>ملخص أداء محفظة الديون ونشاط المحصلين - {today.date()}</p>", unsafe_allow_html=True)
        
        st.markdown(f"""
        <div style="display: grid; grid-template-columns: repeat(6, 1fr); gap: 15px; margin-bottom: 30px;">
            <div class="kpi-card" style="border-left-color: {border_color}; padding: 15px;">
                <div class="kpi-icon">📋</div>
                <div class="kpi-title">إجمالي الحسابات</div>
                <div class="kpi-value">{total_rows:,}</div>
            </div>
            <div class="kpi-card" style="border-left-color: #586F8C; padding: 15px;">
                <div class="kpi-icon">⏳</div>
                <div class="kpi-title">وعود قيد التنفيذ</div>
                <div class="kpi-value">{total_unhandled:,}</div>
            </div>
            <div class="kpi-card" style="border-left-color: #304D73; padding: 15px;">
                <div class="kpi-icon">✅</div>
                <div class="kpi-title">وعود قائمة</div>
                <div class="kpi-value">{active_count:,}</div>
            </div>
            <div class="kpi-card" style="border-left-color: #586F8C; padding: 15px;">
                <div class="kpi-icon">⚠️</div>
                <div class="kpi-title">وعود مكسورة</div>
                <div class="kpi-value">{broken_count:,}</div>
            </div>
            <div class="kpi-card" style="border-left-color: #99AABF; padding: 15px;">
                <div class="kpi-icon">🚨</div>
                <div class="kpi-title">حالات إهمال</div>
                <div class="kpi-value">{neglect_count:,}</div>
            </div>
            <div class="kpi-card" style="border-left-color: {border_color}; padding: 15px;">
                <div class="kpi-icon">📞</div>
                <div class="kpi-title">إجمالي النشاط</div>
                <div class="kpi-value">{total_activity_count:,}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        chart_col1, chart_col2, chart_col3 = st.columns(3)

        # 1. Donut Chart بالترتيب الجديد للألوان (بدون أخضر)
        with chart_col1:
            fig_donut = go.Figure(go.Pie(
                labels=['الوعود القائمة', 'الوعود المكسورة', 'حالات الإهمال'],
                values=[active_count, broken_count, neglect_count],
                hole=0.6,
                marker_colors=["#1F4373", "#304D73", "#586F8C"],
                textinfo='percent+label',
                textfont_color="#FFFFFF"
            ))
            fig_donut.update_layout(
                title_text="توزيع الوعود وحالات الإهمال", 
                title_x=0.5,
                paper_bgcolor='rgba(0,0,0,0)', 
                font_color="#FFFFFF", 
                margin=dict(t=50, b=10, l=10, r=10),
                showlegend=False
            )
            st.plotly_chart(fig_donut, use_container_width=True, config={'displayModeBar': False})

        # 2. Gauge Chart بألوان درجات إجادة
        with chart_col2:
            total_extracted = active_count + broken_count
            gauge_val = int((active_count / total_extracted * 100)) if total_extracted > 0 else 0
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number",
                value=gauge_val,
                number={'suffix': "%", 'font': {'color': "#FFFFFF", 'size': 36, 'weight': 'bold'}},
                title={'text': "نسبة الوعود القائمة من الإجمالي", 'font': {'color': "#FFFFFF", 'size': 14}},
                gauge={
                    'axis': {'range': [0, 100], 'visible': False},
                    'bar': {'color': "#304D73"},
                    'bgcolor': "#1F4373",
                    'borderwidth': 0,
                    'steps': [
                        {'range': [0, 50], 'color': '#586F8C'},
                        {'range': [50, 100], 'color': '#99AABF'}
                    ]
                }
            ))
            fig_gauge.update_layout(paper_bgcolor='rgba(0,0,0,0)', font_color="#FFFFFF", margin=dict(l=20, r=20, t=50, b=20))
            st.plotly_chart(fig_gauge, use_container_width=True, config={'displayModeBar': False})

        # 3. Bar Chart بالترتيب الدقيق للألوان (بدون أخضر)
        with chart_col3:
            if collector_col and collector_col in activity_base_df.columns and total_activity_count > 0:
                top_colls = activity_base_df[collector_col].value_counts().head(5)
                fig_bar = go.Figure(go.Bar(
                    x=top_colls.index,
                    y=top_colls.values,
                    marker_color=["#1F4373", "#304D73", "#586F8C", "#99AABF", "#F2F2F2"],
                    text=top_colls.values,
                    textposition='auto',
                    textfont=dict(color="#FFFFFF", weight='bold')
                ))
                fig_bar.update_layout(
                    title_text="أكثر 5 محصلين نشاطاً", 
                    title_x=0.5,
                    paper_bgcolor='rgba(0,0,0,0)', 
                    plot_bgcolor='rgba(0,0,0,0)', 
                    font_color="#FFFFFF", 
                    margin=dict(t=50, b=10, l=10, r=10)
                )
                fig_bar.update_yaxes(showgrid=False, visible=False)
                st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar': False})
            else:
                st.info("لا توجد بيانات نشاط كافية لعرض رسم بياني.")
                
        st.markdown(f"""
        <div style="background-color: {nav_btn_bg}; border-radius: 12px; padding: 25px; border: 1px solid {border_color}; box-shadow: 0 4px 10px rgba(0,0,0,0.2);">
            <h5 style="margin:0 0 20px 0; color:{text_color}; font-weight:900;">Breakdown Overview</h5>
            <div style="display: flex; align-items: center; justify-content: space-between; padding: 12px 0; border-bottom: 1px solid {border_color}; direction: rtl;">
                <span style="width:30%; color:{text_color}; font-weight: bold; text-align: right;">الوعود القائمة اليوم</span>
                <div style="width: 55%; background-color: {gauge_bg}; height: 8px; border-radius: 4px; overflow: hidden;"><div style="width:{gauge_val}%; background-color:#1F4373; height:100%;"></div></div>
                <span style="color:#1F4373; font-weight: bold; font-size: 14px;">{gauge_val}%</span>
            </div>
            <div style="display: flex; align-items: center; justify-content: space-between; padding: 12px 0; border-bottom: 1px solid {border_color}; direction: rtl;">
                <span style="width:30%; color:{text_color}; font-weight: bold; text-align: right;">الوعود المكسورة</span>
                <div style="width: 55%; background-color: {gauge_bg}; height: 8px; border-radius: 4px; overflow: hidden;"><div style="width:{100-gauge_val if total_extracted > 0 else 0}%; background-color:#304D73; height:100%;"></div></div>
                <span style="color:#304D73; font-weight: bold; font-size: 14px;">{100-gauge_val if total_extracted > 0 else 0}%</span>
            </div>
            <div style="display: flex; align-items: center; justify-content: space-between; padding: 12px 0; direction: rtl;">
                <span style="width:30%; color:{text_color}; font-weight: bold; text-align: right;">حالات الإهمال</span>
                <div style="width: 55%; background-color: {gauge_bg}; height: 8px; border-radius: 4px; overflow: hidden;"><div style="width:{int(neglect_count/total_rows*100) if total_rows > 0 else 0}%; background-color:#586F8C; height:100%;"></div></div>
                <span style="color:#586F8C; font-weight: bold; font-size: 14px;">{int(neglect_count/total_rows*100) if total_rows > 0 else 0}%</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(label=f"📥 تحميل التقرير الكامل - {today.date()}.xlsx", data=output_buffer, file_name=f"Daily_Promises_Report_{today.date()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    # ==========================================
    # باقي الصفحات 
    # ==========================================
    elif st.session_state.current_page == 'Active':
        st.markdown(f"""
        <div style="background-color: {nav_btn_bg}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; margin-bottom: 20px;">
            <h3 style="margin:0; color:#1F4373; font-weight:900;">✅ شيت الوعود القائمة ({active_count:,} حساب)</h3>
        </div>
        """, unsafe_allow_html=True)
        st.dataframe(active_promises_df, use_container_width=True, height=450)
        
    elif st.session_state.current_page == 'Broken':
        st.markdown(f"""
        <div style="background-color: {nav_btn_bg}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; margin-bottom: 20px;">
            <h3 style="margin:0; color:#304D73; font-weight:900;">⚠️ شيت الوعود المكسورة ({broken_count:,} حساب)</h3>
        </div>
        """, unsafe_allow_html=True)
        st.dataframe(broken_promises_df, use_container_width=True, height=450)
        
    elif st.session_state.current_page == 'Neglect':
        st.markdown(f"""
        <div style="background-color: {nav_btn_bg}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; margin-bottom: 20px;">
            <h3 style="margin:0; color:#586F8C; font-weight:900;">🚨 تقرير الإهمال (تأخير أكثر من 3 أيام)</h3>
        </div>
        """, unsafe_allow_html=True)
        default_sub_states = ['وعد بسداد مبلغ المعالجة', 'واعد بسداد مبلغ المعالجه', 'رافض السداد', 'الجوال مغلق', 'مماطل', 'تم سداد جزء وليس مبلغ المعالجه', 'لا يجيب', 'متقاعد الراتب علي بنك آخر', 'يرغب بجدوله المديونيه']
        selected_sub_states = st.multiselect("فلاتر الـ Sub State المشمولة:", options=default_sub_states, default=default_sub_states)
        sub_col = 'Sub State' if 'Sub State' in neglect_df.columns else 'sub state' if 'sub state' in neglect_df.columns else None
        final_neglect_df = neglect_df[neglect_df[sub_col].isin(selected_sub_states)] if sub_col else neglect_df
        st.markdown(f"<p style='text-align:right; direction:rtl; color:#FFFFFF; font-weight: bold;'>العدد الإجمالي: <b style='color:#586F8C; font-size:18px;'>{len(final_neglect_df):,}</b> حساب مهمل</p>", unsafe_allow_html=True)
        st.dataframe(final_neglect_df, use_container_width=True, height=450)

    # ==========================================
    # صفحة تقرير النشاط (Activity Report) 
    # ==========================================
    elif st.session_state.current_page == 'Activity':
        st.markdown(f"""
        <div style="background-color: {nav_btn_bg}; border-radius: 12px; padding: 20px; border: 1px solid {border_color}; direction:rtl; margin-bottom: 20px;">
            <h3 style="margin:0; color:{text_color}; font-weight:900;">📈 تقرير النشاط وموديل التقييم (ناجحة / مغطاة)</h3>
            <p style="margin: 5px 0 0 0; color: {sub_text_color}; font-size: 14px;">تم حذف التكرار تلقائياً بناءً على ({collector_col or 'collecter'}, {id_col or 'id numer'}, {notes_col or 'notes'}).</p>
        </div>
        """, unsafe_allow_html=True)

        with st.expander("📁 ارفعي موديل النشاط (شيت التقييم 1 و 0) - اختياري", expanded=True):
            st.markdown(f"<p style='font-size: 13px; color: {sub_text_color}; direction: rtl;'>التقييم (1) يعني 'ناجحة'، والتقييم (0) يعني 'مغطاة'.</p>", unsafe_allow_html=True)
            mapping_file = st.file_uploader("شيت موديل النشاط (اختياري)", type=["xlsx"], key="mapping_uploader")

        filter_act_col1, filter_act_col2, filter_act_col3 = st.columns([2, 1.5, 1.5])
        with filter_act_col1:
            shift_option = st.selectbox("⏰ اختاري فترة النشاط (الشفت):", ["كل اليوم بدون تحديد", "من 2:00 م إلى 12:00 منتصف الليل (14:00 - 00:00)", "من 12:00 ظهراً إلى 4:00 عصراً (12:00 - 16:00)", "من 4:00 عصراً إلى آخر اليوم (16:00 - 23:59)"])
        with filter_act_col2:
            exclude_last_hour = st.checkbox("⏳ استبعاد آخر ساعة من الفترة", value=False)
        with filter_act_col3:
            exclude_break = st.checkbox("☕ استبعاد البريك (12:00 - 12:30)", value=False)

        if created_col and 'Created_DateTime' in activity_base_df.columns and activity_base_df['Created_DateTime'].notna().any():
            act_df = activity_base_df.dropna(subset=['Created_DateTime']).copy()
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
        else:
            st.info("ℹ️ يتم عرض كافة سجلات النشاط (لم يتم العثور على عمود وقت Created On لتطبيق فلتر الساعات).")
            act_df = activity_base_df.copy()

        if mapping_file is not None:
            map_df = pd.read_excel(mapping_file)
            map_df.columns = map_df.columns.str.strip()
            common_cols = [c for c in map_df.columns if c in act_df.columns]
            if common_cols:
                join_key = common_cols[0] 
                target_status_col = [c for c in map_df.columns if c != join_key]
                if target_status_col:
                    act_df = act_df.merge(map_df[[join_key, target_status_col[0]]], on=join_key, how='left')
                    def classify_call(val):
                        val_str = str(val).strip()
                        if val_str in ['1', '1.0', '1,0']: return 'ناجحة'
                        elif val_str in ['0', '0.0', '0,0']: return 'مغطاة'
                        else: return 'غير مصنف'
                    act_df['Call_Status'] = act_df[target_status_col[0]].apply(classify_call)
            else:
                st.warning("⚠️ لم يتم العثور على عمود مشترك بين شيت النشاط الأساسي وموديل النشاط المرفوع.")
                act_df['Call_Status'] = 'غير مصنف'
        else:
            target_col = find_col_in(act_df, ['التقييم', 'التقيمي', 'evaluation', 'eval', 'تقييم', 'status', 'الحالة', 'sub state', 'result', 'نتيجة المكالمة'])
            target_col = target_col if target_col else act_df.columns[0]
            available_vals = act_df[target_col].dropna().astype(str).unique().tolist()
            selected_success_vals = st.multiselect(f"حددي القيم التي تعتبر (ناجحة = 1) من عمود [{target_col}]:", options=available_vals, default=[v for v in available_vals if str(v).strip() in ['1', '1.0', 'ناجح', 'وعد', 'سداد', 'تم']])
            act_df['Call_Status'] = act_df[target_col].astype(str).apply(lambda x: 'ناجحة' if x in selected_success_vals else 'مغطاة')

        if search_query:
            act_df = act_df[act_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)]

        total_activity = len(act_df)
        success_count = len(act_df[act_df['Call_Status'] == 'ناجحة'])
        covered_count = len(act_df[act_df['Call_Status'] == 'مغطاة'])
        success_rate = (success_count / total_activity * 100) if total_activity > 0 else 0

        if collector_col and collector_col in act_df.columns and total_activity > 0:
            top_collector_series = act_df[collector_col].value_counts()
            if not top_collector_series.empty:
                top_name = top_collector_series.index[0]
                top_calls = top_collector_series.values[0]
                top_pct = (top_calls / total_activity * 100) if total_activity > 0 else 0
                
                st.markdown(f"""
                <div style="background-color: {nav_btn_bg}; border-radius: 12px; padding: 20px 30px; border: 2px solid #586F8C; direction:rtl; margin-bottom: 20px; display: flex; align-items: center; justify-content: space-between;">
                    <div>
                        <h4 style="margin:0; color: #586F8C; font-weight:900;">🏆 بطل الشفت (الأكثر نشاطاً)</h4>
                        <p style="margin: 5px 0 0 0; color: {text_color}; font-size: 18px;">المحصل: <b style="color:#1F4373; font-size: 22px;">{top_name}</b></p>
                    </div>
                    <div style="text-align: left; background-color: {bg_color}; padding: 10px 20px; border-radius: 8px; border: 1px solid #586F8C;">
                        <h3 style="margin:0; color: #586F8C; font-weight:900;">{top_calls:,} مكالمة</h3>
                        <p style="margin: 0; color: #FFFFFF; font-size: 14px; font-weight: bold;">استحوذ على <span style="color:#99AABF;">{top_pct:.1f}%</span> من إجمالي شغل اليوم</p>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        st.markdown(f"""
        <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 20px;">
            <div class="kpi-card" style="border-left-color: {border_color};">
                <div class="kpi-icon">📞</div>
                <div class="kpi-title">إجمالي المحاولات</div>
                <div class="kpi-value">{total_activity:,}</div>
                <div class="kpi-sub" style="color: {sub_text_color};">بعد إزالة التكرار</div>
            </div>
            <div class="kpi-card" style="border-left-color: #1F4373;">
                <div class="kpi-icon">✅</div>
                <div class="kpi-title">المكالمات الناجحة (1)</div>
                <div class="kpi-value">{success_count:,}</div>
                <div class="kpi-sub" style="color: #1F4373;">تقييم ناجح</div>
            </div>
            <div class="kpi-card" style="border-left-color: {border_color};">
                <div class="kpi-icon">🎧</div>
                <div class="kpi-title">المكالمات المغطاة (0)</div>
                <div class="kpi-value">{covered_count:,}</div>
                <div class="kpi-sub" style="color: {sub_text_color};">تقييم مغطاة</div>
            </div>
            <div class="kpi-card" style="border-left-color: #586F8C;">
                <div class="kpi-icon">🎯</div>
                <div class="kpi-title">نسبة النجاح الإجمالية</div>
                <div class="kpi-value">{success_rate:.1f}%</div>
                <div class="kpi-sub" style="color: #586F8C;">معدل الوصول للعملاء</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if collector_col and collector_col in act_df.columns:
            pivot_summary = act_df.groupby([collector_col, 'Call_Status']).size().unstack(fill_value=0)
            if 'ناجحة' not in pivot_summary.columns: pivot_summary['ناجحة'] = 0
            if 'مغطاة' not in pivot_summary.columns: pivot_summary['مغطاة'] = 0
            
            pivot_summary['إجمالي المحاولات'] = pivot_summary['ناجحة'] + pivot_summary['مغطاة']
            pivot_summary['نسبة النجاح %'] = (pivot_summary['ناجحة'] / pivot_summary['إجمالي المحاولات'] * 100).round(1).astype(str) + '%'
            pivot_summary = pivot_summary.sort_values(by='إجمالي المحاولات', ascending=False).reset_index()

            with st.expander("📊 تفصيل أداء المحصلين (ناجحة مقابل مغطاة)", expanded=True):
                st.dataframe(pivot_summary, use_container_width=True)

        extra_time_cols = [c for c in ['Date', 'Time', 'Call_Status'] if c in act_df.columns]
        display_cols = extra_time_cols + [c for c in act_df.columns if c not in ['Created_DateTime', 'Date', 'Time', 'Call_Status']]
        st.dataframe(act_df[display_cols], use_container_width=True, height=450)

        act_buf = io.BytesIO()
        act_df[display_cols].to_excel(act_buf, index=False, sheet_name='تقرير النشاط والتقييم')
        act_buf.seek(0)
        st.download_button("📥 تحميل تقرير النشاط مفلتر.xlsx", data=act_buf, file_name=f"Activity_Success_Report_{today.date()}.xlsx")

else:
    st.markdown(f"""
    <div style="background-color: {nav_btn_bg}; border-radius: 15px; text-align: center; padding: 100px 20px; direction: rtl; border: 1px solid {border_color}; margin-top: 50px;">
        <h2 style="color: {text_color}; margin-bottom: 15px; font-weight:900;">أهلاً بيكي يا رنا في منصة تقارير شركة إجادة لتحصيل الديون</h2>
        <p style="color: {sub_text_color}; font-size: 18px; font-weight: bold;">من شريط التحكم بالأعلى، ارفعي شيت الإكسيل اليومي لعرض الداشبورد والتقارير.</p>
    </div>
    """, unsafe_allow_html=True)
