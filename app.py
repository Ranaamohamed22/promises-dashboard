import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date
import io

# إعدادات الصفحة
st.set_page_config(page_title="Promises Management Dashboard", layout="wide", initial_sidebar_state="expanded")

# حفظ الصفحة الحالية للتحكم في التنقل
if 'current_page' not in st.session_state:
    st.session_state.current_page = 'Dashboard'

# ==========================================
# تصميم الـ UI (CSS)
# ==========================================
st.markdown("""
<style>
    .stApp {
        background-color: #171821;
        color: #FFFFFF;
    }
    [data-testid="stSidebar"] {
        background-color: #171821;
        border-right: 1px solid #2C2D3A;
    }
    .section-card {
        background-color: #21222D;
        border-radius: 15px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.25);
    }
    .metric-box {
        background-color: #171821;
        border-radius: 12px;
        padding: 18px;
        border: 1px solid #2C2D3A;
        text-align: right;
        direction: rtl;
    }
    .metric-icon {
        font-size: 24px;
        margin-bottom: 10px;
        display: inline-block;
        padding: 8px;
        border-radius: 10px;
        background-color: #21222D;
    }
    .metric-val {
        font-size: 26px;
        font-weight: bold;
        color: #FFFFFF;
        margin: 5px 0;
    }
    .metric-lbl {
        font-size: 13px;
        color: #A0A0A0;
        margin-bottom: 6px;
    }
    .metric-sub {
        font-size: 11px;
        font-weight: 600;
    }
    .prog-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 12px 0;
        border-bottom: 1px solid #2C2D3A;
        direction: rtl;
    }
    .prog-bar-bg {
        width: 45%;
        background-color: #2C2D3A;
        height: 6px;
        border-radius: 4px;
        overflow: hidden;
    }
    .badge-pct {
        padding: 3px 10px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: bold;
        background-color: #171821;
    }
    .sidebar-title {
        color: #FFFFFF;
        font-size: 22px;
        font-weight: 800;
        text-align: center;
        margin: 20px 0 15px 0;
        padding: 10px;
        background-color: #21222D;
        border-radius: 12px;
        border: 1px solid #2C2D3A;
    }
    [data-testid="stSidebar"] .stButton > button {
        width: 100% !important;
        text-align: center !important;
        justify-content: center !important;
        direction: rtl;
        padding: 16px 20px !important;
        border-radius: 12px !important;
        margin-bottom: 12px !important;
        display: flex !important;
        align-items: center !important;
    }
    [data-testid="stSidebar"] .stButton > button p {
        font-size: 20px !important;
        font-weight: bold !important;
        margin: 0 !important;
    }
    [data-testid="stSidebar"] .stButton > button[kind="primary"] {
        background-color: #A9DFD8 !important;
        color: #171821 !important;
        border: none !important;
    }
    [data-testid="stSidebar"] .stButton > button[kind="primary"] p {
        color: #171821 !important;
    }
    [data-testid="stSidebar"] .stButton > button[kind="secondary"] {
        background-color: #21222D !important;
        color: #E0E0E0 !important;
        border: 1px solid #2C2D3A !important;
    }
    [data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover {
        background-color: #2C2D3A !important;
        color: #FFFFFF !important;
        border: 1px solid #A9DFD8 !important;
    }
    .stDownloadButton > button {
        background-color: #A9DFD8 !important;
        color: #171821 !important;
        font-weight: bold !important;
        border-radius: 10px !important;
        border: none !important;
        padding: 14px 20px !important;
        width: 100%;
    }
    .stDownloadButton > button p {
        font-size: 17px !important;
        font-weight: bold !important;
        color: #171821 !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# القائمة الجانبية (Sidebar)
# ==========================================
with st.sidebar:
    if st.button("🏠  Dashboard", type="primary" if st.session_state.current_page == 'Dashboard' else "secondary", use_container_width=True):
        st.session_state.current_page = 'Dashboard'
        st.rerun()

    st.markdown("<p style='color:#E0E0E0; font-size:16px; font-weight:bold; text-align:center; direction:rtl; margin-top:15px;'>📂 ارفعي الشيت اليومي من هنا:</p>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")
    
    st.markdown("<hr style='border-color: #2C2D3A; margin: 20px 0;'>", unsafe_allow_html=True)
    st.markdown("<div class='sidebar-title'>📑 قوائم التقارير</div>", unsafe_allow_html=True)

    reports_menu = [
        {"id": "Active", "label": "✅  الوعود القائمة"},
        {"id": "Broken", "label": "⚠️  الوعود المكسورة"},
    ]

    for report in reports_menu:
        if st.button(
            report["label"],
            type="primary" if st.session_state.current_page == report["id"] else "secondary",
            use_container_width=True
        ):
            st.session_state.current_page = report["id"]
            st.rerun()

# ==========================================
# شريط البحث العلوي
# ==========================================
col_search, col_profile = st.columns([4, 1])
with col_search:
    search_query = st.text_input("", placeholder="🔍 ابحثي هنا (برقم العميل، الاسم، أو أي بيان)...", label_visibility="collapsed")
with col_profile:
    st.markdown("<div style='text-align: right; padding-top: 8px; font-size: 18px;'>🔔 &nbsp;&nbsp; 👩‍💻 ⌄</div>", unsafe_allow_html=True)

# ==========================================
# معالجة البيانات
# ==========================================
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()
    
    today = pd.Timestamp(date.today()).normalize()

    due_date = pd.to_datetime(df['Follow up Due Date'], errors='coerce').dt.normalize()
    last_date = pd.to_datetime(df['Follow up Last Date'], errors='coerce').dt.normalize()

    base_condition = (
        (df['حالة المعالجة - التمويل'].astype(str).str.strip() == 'غير معالج') &
        (df['تحديث فئة العميل - التمويل'].astype(str).str.strip().str.lower() != 'writeoff') &
        (df['Classification'].astype(str).str.strip() == 'وعود السداد قيد التنفيذ')
    )

    active_promises_df = df[
        base_condition &
        (due_date == today) &
        (last_date != today)
    ].copy()

    broken_promises_df = df[
        base_condition &
        (due_date < today)
    ].copy()

    broken_promises_df['Date_Diff'] = (last_date[broken_promises_df.index] - due_date[broken_promises_df.index]).dt.days
    broken_promises_df = broken_promises_df[broken_promises_df['Date_Diff'] < 0]

    if search_query:
        active_promises_df = active_promises_df[
            active_promises_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
        ]
        broken_promises_df = broken_promises_df[
            broken_promises_df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
        ]

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        active_promises_df.to_excel(writer, sheet_name='الوعود القائمة', index=False)
        broken_promises_df.to_excel(writer, sheet_name='الوعود المكسورة', index=False)
    output_buffer.seek(0)

    total_rows = len(df)
    total_unhandled = len(df[base_condition])
    active_count = len(active_promises_df)
    broken_count = len(broken_promises_df)

    if st.session_state.current_page == 'Dashboard':
        row1_col1, row1_col2 = st.columns([2.2, 1])

        with row1_col1:
            st.markdown(
                f"""
                <div class="section-card">
                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:15px;">
                        <div>
                            <h4 style="margin:0; color:#FFFFFF;">Today's Promises</h4>
                            <span style="font-size:12px; color:#A0A0A0;">Promises Summary ({today.date()})</span>
                        </div>
                    </div>
                    <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px;">
                        <div class="metric-box">
                            <div class="metric-icon" style="color:#FEB95A;">📊</div>
                            <div class="metric-val">{total_rows:,}</div>
                            <div class="metric-lbl">إجمالي الشيت</div>
                            <div class="metric-sub" style="color:#FEB95A;">كل الحسابات المرفوعة</div>
                        </div>
                        <div class="metric-box">
                            <div class="metric-icon" style="color:#A9DFD8;">⏳</div>
                            <div class="metric-val">{total_unhandled:,}</div>
                            <div class="metric-lbl">قيد التنفيذ (غير معالج)</div>
                            <div class="metric-sub" style="color:#A9DFD8;">بعد استبعاد WriteOff</div>
                        </div>
                        <div class="metric-box">
                            <div class="metric-icon" style="color:#F2C8ED;">✅</div>
                            <div class="metric-val">{active_count:,}</div>
                            <div class="metric-lbl">الوعود القائمة</div>
                            <div class="metric-sub" style="color:#F2C8ED;">مستحقة اليوم</div>
                        </div>
                        <div class="metric-box">
                            <div class="metric-icon" style="color:#20AEF3;">⚠️</div>
                            <div class="metric-val">{broken_count:,}</div>
                            <div class="metric-lbl">الوعود المكسورة</div>
                            <div class="metric-sub" style="color:#20AEF3;">متأخرة (فرق سالب)</div>
                        </div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with row1_col2:
            st.markdown('<div class="section-card"><h5 style="margin:0 0 10px 0;">Level</h5>', unsafe_allow_html=True)
            fig_bar = go.Figure(data=[
                go.Bar(name='العدد', x=['قائمة', 'مكسورة', 'إجمالي وعود'], y=[active_count, broken_count, total_unhandled],
                       marker_color=['#A9DFD8', '#F2C8ED', '#20AEF3'], width=0.4)
            ])
            fig_bar.update_layout(
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                font_color='#A0A0A0', height=185, margin=dict(l=0, r=0, t=0, b=0),
                xaxis=dict(showgrid=False), yaxis=dict(showgrid=False, showticklabels=False)
            )
            st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar': False})
            st.markdown('</div>', unsafe_allow_html=True)

        row2_col1, row2_col2 = st.columns([2.2, 1])

        with row2_col1:
            active_pct = int((active_count / total_unhandled * 100)) if total_unhandled > 0 else 0
            broken_pct = int((broken_count / total_unhandled * 100)) if total_unhandled > 0 else 0
            other_pct = max(0, 100 - active_pct - broken_pct)

            st.markdown(
                f"""
                <div class="section-card">
                    <h5 style="margin:0 0 15px 0;">Promises Breakdown</h5>
                    <div class="prog-row">
                        <span style="width:8%; color:#A0A0A0;">01</span>
                        <span style="width:30%;">الوعود القائمة اليوم</span>
                        <div class="prog-bar-bg"><div style="width:{active_pct}%; background-color:#FEB95A; height:100%;"></div></div>
                        <span class="badge-pct" style="color:#FEB95A; border:1px solid #FEB95A;">{active_pct}%</span>
                    </div>
                    <div class="prog-row">
                        <span style="width:8%; color:#A0A0A0;">02</span>
                        <span style="width:30%;">الوعود المكسورة المتأخرة</span>
                        <div class="prog-bar-bg"><div style="width:{broken_pct}%; background-color:#A9DFD8; height:100%;"></div></div>
                        <span class="badge-pct" style="color:#A9DFD8; border:1px solid #A9DFD8;">{broken_pct}%</span>
                    </div>
                    <div class="prog-row" style="border-bottom:none;">
                        <span style="width:8%; color:#A0A0A0;">03</span>
                        <span style="width:30%;">وعود أخرى قيد التنفيذ</span>
                        <div class="prog-bar-bg"><div style="width:{other_pct}%; background-color:#20AEF3; height:100%;"></div></div>
                        <span class="badge-pct" style="color:#20AEF3; border:1px solid #20AEF3;">{other_pct}%</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

            st.download_button(
                label=f"📥 تحميل التقرير الكامل (القائمة + المكسورة) - {today.date()}.xlsx",
                data=output_buffer,
                file_name=f"Daily_Promises_Report_{today.date()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with row2_col2:
            st.markdown('<div class="section-card"><h5 style="margin:0;">Fulfilment Ratio</h5><p style="font-size:12px; color:#A0A0A0;">نسبة القائمة إلى إجمالي الوعود المستخرجة</p>', unsafe_allow_html=True)
            total_extracted = active_count + broken_count
            gauge_val = int((active_count / total_extracted * 100)) if total_extracted > 0 else 0
            
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number",
                value=gauge_val,
                number={'suffix': "%", 'font': {'color': '#FFFFFF', 'size': 28}},
                gauge={
                    'axis': {'range': [0, 100], 'visible': False},
                    'bar': {'color': "#A9DFD8"},
                    'bgcolor': "#2C2D3A",
                    'borderwidth': 0
                }
            ))
            fig_gauge.update_layout(
                paper_bgcolor='rgba(0,0,0,0)', font_color='#FFFFFF',
                height=160, margin=dict(l=10, r=10, t=10, b=10)
            )
            st.plotly_chart(fig_gauge, use_container_width=True, config={'displayModeBar': False})
            st.markdown('</div>', unsafe_allow_html=True)

    elif st.session_state.current_page == 'Active':
        st.markdown(
            f"""
            <div class="section-card" style="direction:rtl;">
                <h3 style="margin:0; color:#A9DFD8;">✅ شيت الوعود القائمة ({active_count:,} حساب)</h3>
            </div>
            """,
            unsafe_allow_html=True
        )
        st.dataframe(active_promises_df, use_container_width=True, height=450)
        
        active_buf = io.BytesIO()
        active_promises_df.to_excel(active_buf, index=False, sheet_name='الوعود القائمة')
        active_buf.seek(0)
        st.download_button("📥 تحميل شيت الوعود القائمة فقط", data=active_buf, file_name=f"Active_Promises_{today.date()}.xlsx")

    elif st.session_state.current_page == 'Broken':
        st.markdown(
            f"""
            <div class="section-card" style="direction:rtl;">
                <h3 style="margin:0; color:#F2C8ED;">⚠️ شيت الوعود المكسورة ({broken_count:,} حساب)</h3>
            </div>
            """,
            unsafe_allow_html=True
        )
        st.dataframe(broken_promises_df, use_container_width=True, height=450)
        
        broken_buf = io.BytesIO()
        broken_promises_df.to_excel(broken_buf, index=False, sheet_name='الوعود المكسورة')
        broken_buf.seek(0)
        st.download_button("📥 تحميل شيت الوعود المكسورة فقط", data=broken_buf, file_name=f"Broken_Promises_{today.date()}.xlsx")

else:
    st.markdown("""
    <div class="section-card" style="text-align: center; padding: 60px 20px; direction: rtl;">
        <h2 style="color: #A9DFD8; margin-bottom: 10px;">👋 أهلاً بيكي في منصة تقارير الوعود اليومية</h2>
        <p style="color: #A0A0A0; font-size: 16px;">من القائمة الجانبية على الشمال، ارفعي شيت الإكسيل اليومي لعرض الداشبورد والتقارير.</p>
    </div>
    """, unsafe_allow_html=True)