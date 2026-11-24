"""Isolated, functional Streamlit preview of the SNB performance redesign.

Launch with: streamlit run performance_preview.py
This is NOT the production entrypoint app.py. It only calls the unchanged
SNB performance report builder, and never rewrites existing exported workbooks.
"""
from hashlib import sha256
from html import escape

import plotly.graph_objects as go
import streamlit as st

from preview_reporting import prepare_performance_preview
from snb_comprehensive import render_snb_comprehensive_page


st.set_page_config(
    page_title="EJADA | Performance Design Preview",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
:root {
    --ejada-navy: #071f3c;
    --ejada-navy-light: #0c2b50;
    --ejada-green: #04852c;
    --ejada-text: #122744;
    --ejada-border: #d6e2f0;
    --ejada-bg: #f4f8fd;
}
.stApp {background:var(--ejada-bg);}
[data-testid="stMainBlockContainer"] {
    max-width: 1580px; padding-top:1.3rem; padding-left:1.6rem;
    padding-right:1.6rem; padding-bottom:3rem;
}
[data-testid="stSidebar"] {background:var(--ejada-navy)!important; border-right:1px solid #d5e1ee;}
[data-testid="stSidebar"] * {color:white!important;}
[data-testid="stSidebar"] [data-testid="stButton"] button {
    background:transparent!important;
    border:1px solid transparent!important;
    border-radius:10px!important;
    justify-content:flex-start!important;
    min-height:41px!important;
    text-align:right!important;
}
[data-testid="stSidebar"] [data-testid="stButton"] button:hover {
    background:#113b69!important;
}
[data-testid="stSidebar"] [data-testid="stButton"] button[kind="primary"] {
    background:linear-gradient(90deg,#006e2f,#05a13e)!important;
    border-color:#1a9c4b!important;
    font-weight:800!important;
}
.preview-logo {
    display:flex; align-items:center; gap:12px; padding:12px 5px 20px;
    font-size:20px; font-weight:900; line-height:1.15; direction:ltr;
    color:white;
}
.preview-logo-icon {
    display:flex; align-items:center; justify-content:center; width:41px;
    height:41px; border-radius:8px; background:#009339;
    font-size:24px;
}
.preview-section-label {
    margin:18px 0 7px; padding:7px 4px 5px; border-top:1px solid #31516f;
    text-align:right; direction:rtl; font-size:13px; font-weight:700;
    color:#dbe8f5!important;
}
.preview-muted-nav {font-size:12px;color:#a6bbd2!important;line-height:1.9;direction:rtl;}
.preview-banner {
    background:#e9f8ee; border:1px solid #b6ddc2; border-radius:10px;
    padding:8px 14px; margin:0 0 16px; font-size:12px;
    direction:rtl; color:#175436;
}
.preview-header {
    background:white; border:1px solid var(--ejada-border);
    padding:23px 22px 18px; border-radius:16px; direction:rtl;
    display:flex; align-items:center; justify-content:space-between;
    gap:18px; margin-bottom:16px;
}
.preview-header h1 {font-size:29px; color:var(--ejada-text); margin:0 0 5px;font-weight:900;}
.preview-header p {color:#64768c; margin:0;font-size:14px;}
.preview-bank {
    white-space:nowrap; background:#f9fbff; border:1px solid #c8d6e5;
    border-radius:9px; padding:9px 16px; color:#1c3c5c; font-weight:750;
}
.preview-section-heading {
    text-align:right;direction:rtl;font-size:22px;font-weight:850;
    color:var(--ejada-text);margin:18px 0 8px;
}
.preview-kpis {
    display:grid; grid-template-columns:repeat(4,minmax(0,1fr));
    gap:12px; direction:rtl; margin:14px 0 22px;
}
.preview-kpi {
    background:linear-gradient(180deg,#f9fcff,#edf4fd);
    border:1px solid #c8daf0;border-radius:14px;min-height:128px;
    padding:15px 11px;text-align:center;color:var(--ejada-text);
}
.preview-kpi.green {
    background:linear-gradient(180deg,#f5fff8,#e6f9ed);
    border-color:#b5e5c7;
}
.preview-kpi-icon {font-size:24px;margin-bottom:4px;line-height:1;}
.preview-kpi-label {font-size:14px;font-weight:750;margin:6px auto;}
.preview-kpi-value {font-size:32px;line-height:1.2;font-weight:900;color:#061d34;}
.preview-kpi.green .preview-kpi-value {color:#007a29;}
.preview-panel-title {
    direction:rtl;text-align:right;font-size:17px;font-weight:800;
    margin:6px 0 12px;color:var(--ejada-text);
}
.preview-caption {direction:rtl;text-align:right;color:#60758b;font-size:12px;}
.preview-download-row {direction:rtl;text-align:right;margin:8px 0;}
[data-testid="stFileUploader"] {
    direction:rtl;
}
[data-testid="stFileUploaderDropzone"] {
    background:#f1fcf5!important;
    border:1px dashed #89c69e!important;
    border-radius:12px!important;
    min-height:90px!important;
}
[data-testid="stTabs"] [role="tablist"] {justify-content:space-between;gap:2px;direction:rtl;}
[data-testid="stTabs"] [role="tab"] {font-weight:700;color:#344965;}
[data-testid="stTabs"] [aria-selected="true"] {color:var(--ejada-green)!important;}
[data-testid="stTabs"] [data-baseweb="tab-highlight"] {
    background-color:var(--ejada-green)!important;
}
@media (max-width:1000px) {
    .preview-kpis {grid-template-columns:repeat(2,minmax(0,1fr));}
    .preview-header {flex-direction:column;align-items:stretch;}
    [data-testid="stMainBlockContainer"] {padding-left:.6rem;padding-right:.6rem;}
}
@media (max-width:520px) {
    .preview-kpis {grid-template-columns:1fr;}
    .preview-header h1 {font-size:23px;}
}
</style>
""",
    unsafe_allow_html=True,
)


def set_page(page):
    st.session_state["ejada_preview_page"] = page


def render_sidebar():
    with st.sidebar:
        st.markdown(
            '<div class="preview-logo"><span class="preview-logo-icon">📈</span>'
            '<span>Promises<br>Dashboard</span></div>',
            unsafe_allow_html=True,
        )
        page = st.session_state.get("ejada_preview_page", "تقرير الأداء")
        st.button(
            "⌂  الرئيسية", use_container_width=True,
            type="primary" if page == "الرئيسية" else "secondary",
            on_click=set_page, args=("الرئيسية",),
        )
        st.markdown('<div class="preview-section-label">لوحات المتابعة</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="preview-muted-nav">نظرة عامة · الوعود · الإهمال · النشاط'
            '<div style="font-size:11px;">متاحة في التطبيق الأساسي</div></div>',
            unsafe_allow_html=True,
        )
        st.markdown('<div class="preview-section-label">التقارير</div>', unsafe_allow_html=True)
        st.button(
            "▥  تقرير الأداء", use_container_width=True,
            type="primary" if page == "تقرير الأداء" else "secondary",
            on_click=set_page, args=("تقرير الأداء",),
        )
        st.button(
            "▤  التقرير الشامل", use_container_width=True,
            type="primary" if page == "التقرير الشامل" else "secondary",
            on_click=set_page, args=("التقرير الشامل",),
        )
        st.markdown('<div class="preview-section-label">الإدارة</div>', unsafe_allow_html=True)
        st.caption("إدارة موديل النشاط والإعدادات موجودة في التطبيق الأساسي.")
        st.caption("نسخة تصميم تجريبية — لا تغيّر البيانات أو إعدادات الإنتاج.")


def render_header(title, subtitle):
    st.markdown(
        '<div class="preview-header">'
        '<div><h1>▥ ' + escape(title) + '</h1><p>' + escape(subtitle) + '</p></div>'
        '<div class="preview-bank">البنك: &nbsp; SNB ▾</div>'
        '</div>',
        unsafe_allow_html=True,
    )


def render_kpis(view=None):
    if view is None:
        values = ["—", "—", "—", "—"]
    else:
        report = view["report"]
        values = [
            f'{view["total_rows"]:,}',
            f'{report["original_count"]:,}',
            f'{report["payment_count"]:,}',
            f'{view["rate"]:.1f}%' if view["rate"] is not None else "—",
        ]
    items = (
        ("♙", "إجمالي صفوف المحفظة"),
        ("✓", "معالج من المحفظة"),
        ("▱", "معالج بالسداد"),
        ("%", "نسبة الصفوف المعالجة"),
    )
    cards = "".join(
        '<div class="preview-kpi' + (' green' if i == 3 else '') + '">'
        '<div class="preview-kpi-icon">' + icon + '</div>'
        '<div class="preview-kpi-label">' + label + '</div>'
        '<div class="preview-kpi-value">' + escape(values[i]) + '</div></div>'
        for i, (icon, label) in enumerate(items)
    )
    st.markdown('<div class="preview-kpis">' + cards + '</div>', unsafe_allow_html=True)


def render_donut(view):
    treated = view["processed_rows"]
    remainder = view["remaining_rows"]
    if view["total_rows"] == 0:
        st.info("المحفظة لا تحتوي على صفوف.")
        return
    chart = go.Figure(go.Pie(
        labels=["الصفوف المعالجة", "صفوف غير مدرجة ضمن المعالج"],
        values=[treated, remainder],
        hole=0.68,
        marker={"colors": ["#079a42", "#b3bdc9"]},
        textinfo="none",
        sort=False,
    ))
    chart.update_layout(
        height=290, margin=dict(l=4, r=4, b=15, t=10),
        paper_bgcolor="rgba(0,0,0,0)", font=dict(family="Arial"),
        annotations=[dict(text=f'{view["rate"]:.1f}%', x=0.5, y=0.53,
                          showarrow=False, font=dict(size=27, color="#102b45", family="Arial"))],
        legend=dict(orientation="h", y=-0.05, x=0.5, xanchor="center", font=dict(size=11)),
    )
    st.plotly_chart(chart, use_container_width=True, key="preview_pie")


def render_bar(breakdown, dimension, key):
    if breakdown.empty:
        st.info("لا يوجد عمود مناسب للتحليل في الملف؛ لم نفترض تصنيفًا غير موجود.")
        return
    data = breakdown.copy()
    chart = go.Figure()
    chart.add_bar(
        x=data[dimension].astype(str), y=data["إجمالي الصفوف"],
        name="إجمالي الصفوف", marker_color="#b6c6d7"
    )
    chart.add_bar(
        x=data[dimension].astype(str), y=data["الصفوف المعالجة"],
        name="الصفوف المعالجة", marker_color="#03973c"
    )
    chart.update_layout(
        barmode="group", height=290,
        margin=dict(l=5, r=5, b=20, t=10),
        legend=dict(orientation="h", y=-0.28, x=0.5, xanchor="center"),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Arial"),
        yaxis=dict(rangemode="tozero", gridcolor="#e5edf7", title=""),
        xaxis=dict(title=""),
    )
    st.plotly_chart(chart, use_container_width=True, key=key)


def render_visualizations(view):
    tabs = st.tabs(["ملخص الأداء", "تحليل المنتجات", "تحليل المحصلين", "اتجاه السدادات"])
    with tabs[0]:
        left, right = st.columns([1.05, 1.75], gap="medium")
        with left:
            with st.container(border=True):
                st.markdown('<div class="preview-panel-title">توزيع المعالجة</div>',
                            unsafe_allow_html=True)
                render_donut(view)
        with right:
            with st.container(border=True):
                st.markdown('<div class="preview-panel-title">المعالجة حسب المنتج</div>',
                            unsafe_allow_html=True)
                render_bar(view["products"], "المنتج", "preview_products_summary")
    with tabs[1]:
        st.markdown("**تفصيل الصفوف حسب المنتج (بدون تغيير أرقام التقرير):**")
        render_bar(view["products"], "المنتج", "preview_products_detail")
        if not view["products"].empty:
            st.dataframe(view["products"], use_container_width=True, hide_index=True)
    with tabs[2]:
        st.markdown("**تفصيل الصفوف حسب المحصل:**")
        render_bar(view["collectors"], "المحصل", "preview_collectors")
        if not view["collectors"].empty:
            st.dataframe(view["collectors"], use_container_width=True, hide_index=True)
    with tabs[3]:
        st.caption("هذا اتجاه مبالغ السدادات حسب تاريخ إنشائها، وليس اتجاه معالجة الحسابات.")
        daily = view["payments_by_date"]
        if daily.empty:
            st.info("لا توجد تواريخ سدادات صالحة لعرض الاتجاه.")
        else:
            chart = go.Figure(go.Scatter(
                x=daily["التاريخ"], y=daily["مبلغ السداد"],
                mode="lines+markers", line=dict(color="#049139", width=3),
                marker=dict(size=6), name="مبلغ السداد",
            ))
            chart.update_layout(
                height=320, margin=dict(l=12, r=15, b=20, t=10),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                yaxis=dict(title="مبلغ السداد", gridcolor="#e5edf7"),
                xaxis=dict(title="التاريخ"),
            )
            st.plotly_chart(chart, use_container_width=True, key="preview_payments_trend")


def render_performance():
    render_header("تقرير الأداء", "تحليل أداء المحفظة والمعالجة والسدادات — معاينة التصميم الجديد")
    st.markdown('<div class="preview-banner">نسخة تجريبية مستقلة: '
                'كل الحسابات والتصدير تتم من خلال كود التقرير الأصلي دون تعديل.</div>',
                unsafe_allow_html=True)
    col_portfolio, col_payments = st.columns(2, gap="medium")
    with col_portfolio, st.container(border=True):
        st.markdown("**📄 ملف المحفظة**")
        portfolio_file = st.file_uploader(
            "ارفعي ملف المحفظة Excel",
            type=["xlsx"], key="preview_portfolio",
        )
    with col_payments, st.container(border=True):
        st.markdown("**📄 ملف السدادات**")
        payments_file = st.file_uploader(
            "ارفعي ملف السدادات Excel",
            type=["xlsx"], key="preview_payments",
        )

    if portfolio_file is None or payments_file is None:
        render_kpis()
        st.info("ارفعي ملف المحفظة وملف السدادات لعرض المؤشرات والرسومات والتحميل. لا توجد أرقام افتراضية.")
        return

    p_bytes = portfolio_file.getvalue()
    s_bytes = payments_file.getvalue()
    try:
        view = prepare_performance_preview(p_bytes, s_bytes)
    except Exception as exc:
        render_kpis()
        st.error(f"تعذر تجهيز التقرير: {exc}")
        return

    report = view["report"]
    st.caption(
        f"المحفظة: {escape(portfolio_file.name)} · {sha256(p_bytes).hexdigest()[:8]} | "
        f"السدادات: {escape(payments_file.name)} · {sha256(s_bytes).hexdigest()[:8]} | "
        f"عدد سدادات اليوم: {report['daily_count']:,}"
    )
    if report["invalid_treatment_rows"]:
        st.warning(
            f"هناك {len(report['invalid_treatment_rows']):,} صف غير معالج "
            "به قيمة مبلغ معالجة غير صالحة. لم نعتبر هذه القيم صفرًا."
        )
    render_kpis(view)
    render_visualizations(view)

    st.markdown('<div class="preview-section-heading">تفاصيل التقرير</div>', unsafe_allow_html=True)
    st.dataframe(view["details"], use_container_width=True, hide_index=True, height=270)
    with st.expander("عرض كل أعمدة تقرير المعالج الأصلي"):
        st.dataframe(view["processed_full"], use_container_width=True, hide_index=True, height=300)
    download_one, download_two = st.columns(2, gap="small")
    today = report["today"].isoformat()
    with download_one:
        st.download_button(
            "⬇ تصدير تقرير المعالج Excel (الأصلي)",
            data=report["processed_bytes"],
            file_name=f"SNB_Processed_{today}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with download_two:
        st.download_button(
            "⬇ تصدير سدادات اليوم Excel (الأصلي)",
            data=report["daily_bytes"],
            file_name=f"SNB_Daily_Payments_{today}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


render_sidebar()
page = st.session_state.get("ejada_preview_page", "تقرير الأداء")
if page == "الرئيسية":
    render_header("الرئيسية", "هذه نسخة مستقلة لمراجعة تصميم SNB قبل دمجه.")
    st.info("اختاري «تقرير الأداء» من الشريط الجانبي لمشاهدة التصميم التفاعلي. التطبيق الأصلي غير متأثر.")
elif page == "التقرير الشامل":
    render_header("التقرير الشامل", "النسخة الحالية من التقرير الشامل، دون تعديل لمنطق الحسابات.")
    render_snb_comprehensive_page()
else:
    render_performance()
