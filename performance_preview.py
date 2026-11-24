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
from preview_theme import resolve_theme, THEME_OPTIONS
from snb_comprehensive import render_snb_comprehensive_page


st.set_page_config(
    page_title="EJADA | Performance Design Preview",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Resolve the selected bank/theme before drawing any widgets.
# This affects colors only; STC uploads never use SNB business logic.
if "ejada_preview_bank" not in st.session_state:
    st.session_state["ejada_preview_bank"] = "SNB"
if "ejada_preview_theme_mode" not in st.session_state:
    st.session_state["ejada_preview_theme_mode"] = "حسب البنك"
preview_bank = st.session_state["ejada_preview_bank"]
preview_theme_mode = st.session_state["ejada_preview_theme_mode"]
palette = resolve_theme(preview_bank, preview_theme_mode)

_PREVIEW_CSS = """
.stApp {background:var(--ejada-background)!important;color:var(--ejada-text)!important;}
[data-testid="stMainBlockContainer"] {
    max-width:1580px;padding-top:0.75rem;padding-left:1.4rem;
    padding-right:1.4rem;padding-bottom:3rem;
}
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"] {display:none!important;}
.stApp h1,.stApp h2,.stApp h3,.stApp h4,
.stApp [data-testid="stMarkdownContainer"] h1,
.stApp [data-testid="stMarkdownContainer"] h2 {
    color:var(--ejada-text)!important;
}
.stApp [data-testid="stWidgetLabel"] p,
.stApp [data-testid="stCaptionContainer"] p,
.stApp [data-testid="stFileUploader"] label,
.stApp [data-testid="stFileUploader"] span,
.stApp [data-testid="stFileUploader"] small {
    color:var(--ejada-muted)!important;
    opacity:1!important;
}
.stApp [data-testid="stWidgetLabel"] p {font-weight:750!important;}
.stApp [data-testid="stMainBlockContainer"] [data-testid="stMarkdownContainer"] p {
    color:var(--ejada-text);
}
.stApp [data-testid="stMainBlockContainer"] [data-testid="stCaptionContainer"] p {
    color:var(--ejada-muted)!important;
}
.stApp [data-testid="stFileUploaderDropzone"] [data-testid="stFileUploaderDropzoneInstructions"] p {
    color:var(--ejada-text)!important;
}
.stApp [data-testid="stSelectbox"] [data-baseweb="select"]>div {
    background:var(--ejada-surface)!important;
    border:1px solid var(--ejada-border)!important;
    color:var(--ejada-text)!important;
    border-radius:10px!important;
    min-height:38px!important;
}
.stApp [data-testid="stSelectbox"] [data-baseweb="select"] * {
    color:var(--ejada-text)!important;
}
.stApp [data-testid="stSelectbox"] [data-baseweb="select"] svg {
    fill:var(--ejada-text)!important;
}
.stApp [data-testid="stAlert"] {
    border-radius:10px;
}
.st-key-preview_navbar {
    background:var(--ejada-surface);
    border:1px solid var(--ejada-border);
    border-top:3px solid var(--ejada-accent);
    border-radius:12px;
    padding:9px 12px!important;
    margin-bottom:16px!important;
}
.st-key-preview_navbar [data-testid="stHorizontalBlock"]{
    align-items:center!important;
}
.preview-bank-chip {
    direction:ltr;display:flex;align-items:center;gap:8px;
    white-space:nowrap;font-size:18px;font-weight:900;color:var(--ejada-text);
    min-height:38px;
}
.preview-bank-chip:before {
    content:"";width:10px;height:10px;border-radius:100%;
    background:var(--ejada-accent);display:inline-block;
}
.st-key-preview_report_pills {
    direction:rtl!important;
    display:flex;align-items:center;justify-content:flex-start;
    flex-wrap:wrap!important;gap:4px!important;
    padding:4px!important;border-radius:30px;
    background:var(--ejada-surface-alt);
}
.st-key-preview_report_pills [data-testid="stButton"] button {
    white-space:nowrap!important;border:1px solid transparent!important;
    border-radius:23px!important;
    min-height:36px!important;
    padding:5px 11px!important;
    background:transparent!important;
    box-shadow:none!important;
    font-size:13px!important;
}
.st-key-preview_report_pills [data-testid="stButton"] button p {
    color:var(--ejada-text)!important;
    font-weight:650!important;
    white-space:nowrap!important;
}
.st-key-preview_report_pills [data-testid="stButton"] button[kind="primary"] {
    background:var(--ejada-accent-soft)!important;
    border-color:var(--ejada-border)!important;
    box-shadow:0 1px 4px #101e2f12!important;
}
.st-key-preview_report_pills [data-testid="stButton"] button[kind="primary"] p {
    color:var(--ejada-accent-text)!important;font-weight:900!important;
}
.preview-header {
    background:var(--ejada-surface);
    border:1px solid var(--ejada-border);
    padding:19px 23px 16px;
    border-radius:15px;
    direction:rtl;
    margin-bottom:13px;
}
.preview-header h1 {
    font-size:26px;color:var(--ejada-text);margin:0 0 5px;font-weight:900;
}
.preview-header p {color:var(--ejada-muted);margin:0;font-size:14px;}
.preview-banner {
    background:var(--ejada-accent-soft);
    border:1px solid var(--ejada-border);
    border-radius:10px;
    padding:8px 14px;margin:0 0 16px;
    font-size:12px;direction:rtl;color:var(--ejada-text);
}
.preview-section-heading {
    text-align:right;direction:rtl;
    font-size:22px;font-weight:850;color:var(--ejada-text);
    margin:18px 0 8px;
}
.st-key-preview_upload_row [data-testid="stHorizontalBlock"] {
    direction:rtl!important;
    align-items:flex-start!important;
}
.st-key-preview_portfolio_field,
.st-key-preview_payments_field {
    width:min(285px,100%)!important;
    max-width:285px!important;
    padding:0!important;
    border:none!important;
    background:transparent!important;
}
.preview-upload-heading {
    color:var(--ejada-text)!important;
    font-size:15px;font-weight:850;
    text-align:right;direction:rtl;
    margin:0 0 5px;
}
.st-key-preview_portfolio_field [data-testid="stFileUploader"],
.st-key-preview_payments_field [data-testid="stFileUploader"] {
    width:min(285px,100%)!important;max-width:285px!important;
    direction:rtl!important;
}
.st-key-preview_portfolio_field [data-testid="stFileUploaderDropzone"],
.st-key-preview_payments_field [data-testid="stFileUploaderDropzone"] {
    min-height:54px!important;
    padding:6px 9px!important;
    background:var(--ejada-accent-soft)!important;
    border:1px dashed var(--ejada-accent)!important;
    border-radius:9px!important;
    display:flex!important;
    align-items:center!important;
    justify-content:center!important;
}
.st-key-preview_upload_row [data-testid="stFileUploaderDropzoneInstructions"] {
    display:none!important;
}
.st-key-preview_upload_row [data-testid="stFileUploaderDropzone"] button {
    background:var(--ejada-accent)!important;
    border:1px solid var(--ejada-accent)!important;
    border-radius:8px!important;
    padding:5px 14px!important;
    min-height:34px!important;
    color:white!important;
}
.st-key-preview_upload_row [data-testid="stFileUploaderDropzone"] button p {
    color:white!important;
}
.st-key-preview_upload_row [data-testid="stFileUploaderDropzone"] button:hover {
    background:var(--ejada-accent-hover)!important;
}
.preview-kpis {
    display:grid;grid-template-columns:repeat(4,minmax(0,1fr));
    gap:12px;direction:rtl;margin:14px 0 22px;
}
.preview-kpi {
    background:var(--ejada-surface-alt);
    border:1px solid var(--ejada-border);
    border-radius:14px;min-height:118px;
    padding:14px 10px;text-align:center;color:var(--ejada-text);
}
.preview-kpi.highlight {
    background:var(--ejada-accent-soft);
    border-color:var(--ejada-accent);
}
.preview-kpi-icon {
    font-size:24px;margin-bottom:4px;line-height:1;
    color:var(--ejada-accent);
}
.preview-kpi-label {
    font-size:14px;font-weight:750;margin:6px auto;
    color:var(--ejada-text);
}
.preview-kpi-value {
    font-size:32px;line-height:1.2;font-weight:900;color:var(--ejada-text);
}
.preview-kpi.highlight .preview-kpi-value {color:var(--ejada-accent-text);}
.preview-panel-title {
    direction:rtl;text-align:right;font-size:17px;font-weight:800;
    margin:6px 0 12px;color:var(--ejada-text);
}
.stApp [data-testid="stVerticalBlockBorderWrapper"] {
    border-color:var(--ejada-border)!important;
    background:var(--ejada-surface)!important;
    border-radius:12px!important;
}
[data-testid="stTabs"] [role="tablist"] {
    justify-content:space-between;gap:2px;direction:rtl;
    background:var(--ejada-surface-alt);border-radius:9px;
}
[data-testid="stTabs"] [role="tab"] {
    font-weight:750!important;color:var(--ejada-text)!important;
}
[data-testid="stTabs"] [aria-selected="true"] {
    color:var(--ejada-accent-text)!important;
}
[data-testid="stTabs"] [data-baseweb="tab-highlight"] {
    background-color:var(--ejada-accent)!important;
}
.stApp [data-testid="stDownloadButton"] button {
    border-color:var(--ejada-accent)!important;
    background:var(--ejada-accent-soft)!important;
}
.stApp [data-testid="stDownloadButton"] button p{
    color:var(--ejada-accent-text)!important;
}
.stApp [data-testid="stDataFrame"] {
    border:1px solid var(--ejada-border)!important;
    border-radius:9px;
}
@media (max-width:1200px) {
    .st-key-preview_navbar [data-testid="stHorizontalBlock"]{flex-wrap:wrap!important;}
}
@media (max-width:940px) {
    .preview-kpis {grid-template-columns:repeat(2,minmax(0,1fr));}
    [data-testid="stMainBlockContainer"]{
        padding-left:.7rem;padding-right:.7rem;
    }
}
@media (max-width:520px) {
    .preview-kpis {grid-template-columns:1fr;}
    .preview-header h1 {font-size:22px;}
    .st-key-preview_report_pills [data-testid="stButton"] button{
        font-size:12px!important;
    }
}
"""


def apply_preview_theme():
    variable_css = ":root,.stApp{" + "".join(
        f"--ejada-{name.replace('_', '-')}:{value};"
        for name, value in palette.items()
    ) + "}"
    st.markdown("<style>" + variable_css + _PREVIEW_CSS + "</style>",
                unsafe_allow_html=True)


def set_page(page):
    st.session_state["ejada_preview_page"] = page


def render_top_navigation():
    """Match the existing application's compact, horizontal report row."""
    pages = [
        ("نظرة عامة", "نظرة عامة"),
        ("الوعود القائمة والمكسورة", "الوعود القائمة والمكسورة"),
        ("الإهمال", "الإهمال"),
        ("النشاط", "النشاط"),
    ]
    if preview_bank == "SNB":
        pages.extend([
            ("تقرير الأداء", "تقرير الأداء"),
            ("التقرير الشامل", "التقرير الشامل"),
        ])
    page = st.session_state.get("ejada_preview_page", "تقرير الأداء")
    with st.container(key="preview_navbar"):
        bank_col, nav_col = st.columns([3.9, 8.1], gap="small", vertical_alignment="center")
        with bank_col:
            label_col, bank_select_col, theme_col = st.columns(
                [0.8, 1.05, 1.2], gap="small", vertical_alignment="center"
            )
            with label_col:
                st.markdown(
                    '<div class="preview-bank-chip">' + escape(preview_bank) + '</div>',
                    unsafe_allow_html=True,
                )
            with bank_select_col:
                st.selectbox(
                    "البنك", options=["SNB", "STC"],
                    key="ejada_preview_bank",
                    label_visibility="collapsed",
                    help="اختيار البنك يغير ألوان المعاينة حسب هوية البنك.",
                )
            with theme_col:
                st.selectbox(
                    "ثيم الألوان", options=THEME_OPTIONS,
                    key="ejada_preview_theme_mode",
                    label_visibility="collapsed",
                    help="الأسود الهادي يحتفظ بلون البنك المميز.",
                )
        with nav_col:
            with st.container(
                key="preview_report_pills", horizontal=True,
                horizontal_alignment="right", gap="small",
            ):
                for name, label in pages:
                    st.button(
                        label,
                        key="preview_nav_" + {
                            "نظرة عامة": "overview",
                            "الوعود القائمة والمكسورة": "promises",
                            "الإهمال": "neglect",
                            "النشاط": "activity",
                            "تقرير الأداء": "performance",
                            "التقرير الشامل": "comprehensive",
                        }[name],
                        type="primary" if page == name else "secondary",
                        on_click=set_page, args=(name,),
                    )


def render_header(title, subtitle):
    st.markdown(
        '<div class="preview-header">'
        '<div><h1>▥ ' + escape(title) + '</h1><p>' + escape(subtitle) + '</p></div>'
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
        '<div class="preview-kpi' + (' highlight' if i == 3 else '') + '">'
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
        marker={"colors": [palette["accent"], palette["neutral_chart"]]},
        textinfo="none",
        sort=False,
    ))
    chart.update_layout(
        height=290, margin=dict(l=4, r=4, b=15, t=10),
        paper_bgcolor="rgba(0,0,0,0)", font=dict(family="Arial", color=palette["text"]),
        annotations=[dict(text=f'{view["rate"]:.1f}%', x=0.5, y=0.53,
                          showarrow=False, font=dict(size=27, color=palette["text"], family="Arial"))],
        legend=dict(orientation="h", y=-0.05, x=0.5, xanchor="center", font=dict(size=11, color=palette["text"])),
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
        name="إجمالي الصفوف", marker_color=palette["neutral_chart"]
    )
    chart.add_bar(
        x=data[dimension].astype(str), y=data["الصفوف المعالجة"],
        name="الصفوف المعالجة", marker_color=palette["accent"]
    )
    chart.update_layout(
        barmode="group", height=290,
        margin=dict(l=5, r=5, b=20, t=10),
        legend=dict(orientation="h", y=-0.28, x=0.5, xanchor="center"),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Arial", color=palette["text"]),
        yaxis=dict(rangemode="tozero", gridcolor=palette["grid"], title=""),
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
                mode="lines+markers", line=dict(color=palette["accent"], width=3),
                marker=dict(size=6), name="مبلغ السداد",
            ))
            chart.update_layout(
                height=320, margin=dict(l=12, r=15, b=20, t=10),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color=palette["text"]),
                yaxis=dict(title="مبلغ السداد", gridcolor=palette["grid"]),
                xaxis=dict(title="التاريخ"),
            )
            st.plotly_chart(chart, use_container_width=True, key="preview_payments_trend")


def render_performance():
    render_header("تقرير الأداء", "تحليل أداء المحفظة والمعالجة والسدادات — معاينة التصميم الجديد")
    st.markdown('<div class="preview-banner">نسخة تجريبية مستقلة: '
                'كل الحسابات والتصدير تتم من خلال كود التقرير الأصلي دون تعديل.</div>',
                unsafe_allow_html=True)
    # Preserve the original portfolio-then-payments order in RTL display.
    # Compact native upload controls only; no large surrounding upload cards.
    if preview_bank != "SNB":
        render_kpis()
        st.info(
            "تم تطبيق ألوان STC للمعاينة فقط. تقرير الأداء هنا مبني على "
            "حسابات SNB، لذلك رفع ملفات STC غير متاح في هذه النسخة التجريبية."
        )
        return
    with st.container(key="preview_upload_row"):
        col_portfolio, col_payments, _spacer = st.columns(
            [1.25, 1.25, 2.2], gap="small", vertical_alignment="top"
        )
        with col_portfolio, st.container(key="preview_portfolio_field"):
            st.markdown(
                '<div class="preview-upload-heading">📄 ملف المحفظة</div>',
                unsafe_allow_html=True,
            )
            portfolio_file = st.file_uploader(
                "ارفعي ملف المحفظة Excel", type=["xlsx"],
                key="preview_portfolio", label_visibility="collapsed",
            )
        with col_payments, st.container(key="preview_payments_field"):
            st.markdown(
                '<div class="preview-upload-heading">📄 ملف السدادات</div>',
                unsafe_allow_html=True,
            )
            payments_file = st.file_uploader(
                "ارفعي ملف السدادات Excel", type=["xlsx"],
                key="preview_payments", label_visibility="collapsed",
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


apply_preview_theme()
render_top_navigation()
page = st.session_state.get("ejada_preview_page", "تقرير الأداء")
if page == "تقرير الأداء":
    render_performance()
elif page == "التقرير الشامل" and preview_bank == "SNB":
    render_header("التقرير الشامل", "تقرير SNB الحالي بحساباته الأصلية — معاينة التصميم فقط.")
    render_snb_comprehensive_page()
else:
    render_header(page, "معاينة ثيم الواجهة والتنقل — دون تغيير التقارير الأصلية.")
    st.info(
        "هذه الصفحة موجودة في التطبيق الأساسي، ولم ننقل بياناتها أو "
        "نغيّر منطقها في نسخة التصميم التجريبية."
    )
