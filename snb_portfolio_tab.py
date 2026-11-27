"""Read-only SNB portfolio summary tab, independent of processed/payments rules.

Mirrors the supplied Excel layout (collector / product / DPD30 / DPD60 / NPL).
Only the currently uploaded portfolio is read; no customer records are stored.
Existing build_snb_performance and its exported workbooks are not modified.
"""
from html import escape as html_escape
from io import BytesIO
import re
import unicodedata

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


STAGES = ("DPD30", "DPD60", "NPL")
COLLECTOR_CANDIDATES = ("Salesperson", "Collector", "Collecter", "Collector Name",
                        "COLLECTOR_NAME", "اسم المحصل", "المحصل")
PRODUCT_CANDIDATES = ("PRODUCT_CATEGORY", "Product Category", "Product", "PRODUCT",
                      "Product_Type", "المنتج", "نوع المنتج",
                      "نوع المنتج-التمويل", "نوع المنتج التمويل",
                      "نوع المتنج-التمويل")  # Header in current SNB CRM export
STAGE_CANDIDATES = ("ملاحظات-التمويل", "ملاحظات التمويل", "DPD_BUCKET", "Stage", "المرحلة")
DEBT_CANDIDATES = ("Net Amount", "NET_AMOUNT", "Net_Amount",
                   "Total Debt", "Outstanding Amount", "إجمالي المديونية")
ACCOUNT_CANDIDATES = ("Account Number", "ACCOUNT_NUMBER", "Account No", "رقم الحساب")

COLUMNS = (
    "اسم المحصل", "المنتج",
    "DPD30 | عدد الحسابات", "DPD30 | إجمالي المديونية",
    "DPD60 | عدد الحسابات", "DPD60 | إجمالي المديونية",
    "NPL | عدد الحسابات", "NPL | إجمالي المديونية",
)


def _header_key(value):
    text = unicodedata.normalize("NFKC", str(value)).casefold()
    return "".join(
        char for char in text
        if unicodedata.category(char) not in {"Cf", "Mn", "Pd"}
        and not char.isspace() and char not in "_ـ−"
    )


def _required_column(frame, candidates, label):
    exact = {str(col).strip().casefold(): col for col in frame.columns}
    normalized = {}
    for col in frame.columns:
        normalized.setdefault(_header_key(col), []).append(col)
    for candidate in candidates:
        match = exact.get(candidate.casefold())
        if match is not None:
            return match
    for candidate in candidates:
        matches = normalized.get(_header_key(candidate), ())
        match = matches[0] if len(matches) == 1 else None
        if match is not None:
            return match
    raise ValueError(
        f"لا يمكن تكوين جدول المحفظة: عمود {label} غير موجود. "
        f"الأسماء المدعومة: {', '.join(candidates)}."
    )


def _normalize_stage(value):
    if pd.isna(value):
        return ""
    canonical = re.sub(r"[\s_\-]+", "", str(value)).upper()
    if canonical in ("DPD30", "30DPD"):
        return "DPD30"
    if canonical in ("DPD60", "60DPD"):
        return "DPD60"
    if canonical == "NPL":
        return "NPL"
    return ""


def build_portfolio_summary(portfolio_bytes):
    """Return an Excel-copyable eight-column report and data-quality counts.

    Account count follows existing pivot convention: number of nonblank
    Account Number rows (NOT distinct IDs). Debt sums the Net Amount values,
    with blanks treated as zero. Unrecognized stages are disclosed and never
    silently included in one of the three columns.
    """
    required = (
        (COLLECTOR_CANDIDATES, "اسم المحصل"),
        (PRODUCT_CANDIDATES, "المنتج"),
        (STAGE_CANDIDATES, "مرحلة المحفظة"),
        (DEBT_CANDIDATES, "إجمالي المديونية"),
        (ACCOUNT_CANDIDATES, "رقم الحساب"),
    )
    with pd.ExcelFile(BytesIO(portfolio_bytes)) as workbook:
        first_sheet = workbook.sheet_names[0]
        source = workbook.parse(first_sheet, dtype=object)
        try:
            columns = [_required_column(source, candidates, label)
                       for candidates, label in required]
        except ValueError as first_error:
            matches = []
            for name in workbook.sheet_names[1:]:
                headers = workbook.parse(name, nrows=0)
                try:
                    columns = [_required_column(headers, candidates, label)
                               for candidates, label in required]
                except ValueError:
                    continue
                matches.append((name, columns))
            if len(matches) > 1:
                names = "، ".join(name for name, _ in matches)
                raise ValueError(
                    f"أكثر من شيت فيه أعمدة المحفظة ({names}). "
                    "خلي شيت المحفظة المطلوب أول شيت في الملف."
                ) from first_error
            if not matches:
                raise ValueError(
                    f"لم أجد شيت فيه كل أعمدة المحفظة المطلوبة. "
                    f"أول شيت اتقرأ: «{first_sheet}». {first_error}"
                ) from first_error
            name, columns = matches[0]
            source = workbook.parse(name, dtype=object)

    collector_col, product_col, stage_col, debt_col, account_col = columns

    collector = source[collector_col].astype("string").fillna("").str.strip()
    product = source[product_col].astype("string").fillna("").str.strip().str.upper()
    stage = source[stage_col].map(_normalize_stage)
    accounts = source[account_col].astype("string").fillna("").str.strip()
    debt_text = (
        source[debt_col].astype("string").str.strip()
        .str.replace(",", "", regex=False)
        .str.replace("٬", "", regex=False)
    )
    blank_debt = debt_text.isna() | debt_text.eq("")
    debts = pd.to_numeric(debt_text.mask(blank_debt), errors="coerce")
    invalid_debt = (~blank_debt & debts.isna()).fillna(False)
    recognized = stage.isin(STAGES)
    if (invalid_debt & recognized).any():
        bad = int((invalid_debt & recognized).sum())
        raise ValueError(
            f"عمود {debt_col} يحتوي على {bad} مبلغ غير رقمي ضمن مراحل "
            "DPD30/DPD60/NPL. لم نعتبر هذه القيم صفرًا."
        )

    raw = pd.DataFrame({
        "اسم المحصل": collector,
        "المنتج": product,
        "stage": stage,
        "account_count": accounts.ne("").astype(int),
        "debt": debts.fillna(0).astype(float),
    })
    excluded = int((~recognized).sum())
    # Disclose unassigned rows as well; never treat them as assigned.
    missing_labels = int(
        (recognized & (collector.eq("") | product.eq(""))).sum()
    )
    valid = recognized & collector.ne("") & product.ne("")
    filtered = raw.loc[valid].copy()
    if filtered.empty:
        return pd.DataFrame(columns=COLUMNS), {
            "source_rows": len(source),
            "included_rows": 0,
            "excluded_stages": excluded,
            "missing_labels": missing_labels,
            "blank_debts": int((blank_debt & valid).sum()),
        }

    grouped = (
        filtered.groupby(["اسم المحصل", "المنتج", "stage"], sort=False, dropna=False)
        .agg(account_count=("account_count", "sum"), debt=("debt", "sum"))
    )
    ordered_pairs = filtered[["اسم المحصل", "المنتج"]].drop_duplicates()
    product_rank = {"CC": 0, "PF": 1, "RF": 2}
    # Preserve collector's first appearance; within each collector show CC/PF/RF.
    rows = []
    for collector_name in ordered_pairs["اسم المحصل"].drop_duplicates():
        products = ordered_pairs.loc[
            ordered_pairs["اسم المحصل"].eq(collector_name), "المنتج"
        ].tolist()
        products.sort(key=lambda name: (product_rank.get(name, 99), name))
        for product_name in products:
            row = [str(collector_name), str(product_name)]
            for stage_name in STAGES:
                key = (collector_name, product_name, stage_name)
                if key in grouped.index:
                    row.extend([
                        int(grouped.loc[key, "account_count"]),
                        float(grouped.loc[key, "debt"]),
                    ])
                else:
                    row.extend([0, 0.0])
            rows.append(row)
    return pd.DataFrame(rows, columns=COLUMNS), {
        "source_rows": len(source),
        "included_rows": int(valid.sum()),
        "excluded_stages": excluded,
        "missing_labels": missing_labels,
        "blank_debts": int((blank_debt & valid).sum()),
    }


def write_portfolio_sheet(sheet, table):
    """Write the existing portfolio grid unchanged into any Excel worksheet."""
    sheet.title = "المحفظة"
    sheet.sheet_view.rightToLeft = False
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "C3"

    dark = PatternFill("solid", fgColor="075143")
    light = PatternFill("solid", fgColor="E0EFE9")
    white = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    forest = Font(name="Arial", size=10, bold=True, color="064B3D")
    thin = Side(style="thin", color="A8B8B3")
    edge = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for columns, label in (
        ("A1:A2", "اسم المحصل"), ("B1:B2", "المنتج"),
        ("C1:D1", "DPD30"), ("E1:F1", "DPD60"), ("G1:H1", "NPL"),
    ):
        sheet.merge_cells(columns)
        top = sheet[columns.split(":")[0]]
        top.value = label
    for col in range(1, 9):
        header = sheet.cell(1, col)
        header.fill = dark
        header.border = edge
        header.font = white
        header.alignment = center
        sub = sheet.cell(2, col)
        sub.fill = dark if col <= 2 else light
        sub.border = edge
        sub.font = white if col <= 2 else forest
        sub.alignment = center
    for col in (3, 5, 7):
        sheet.cell(2, col, "عدد الحسابات")
        sheet.cell(2, col + 1, "إجمالي المديونية")
        for c in (col, col + 1):
            sheet.cell(2, c).font = forest
            sheet.cell(2, c).alignment = center

    sheet.row_dimensions[1].height = 31
    sheet.row_dimensions[2].height = 34
    for record in table.itertuples(index=False, name=None):
        sheet.append(list(record))
        ridx = sheet.max_row
        for col in range(1, 9):
            cell = sheet.cell(ridx, col)
            cell.border = edge
            cell.alignment = Alignment(
                horizontal="right" if col == 1 else "center",
                vertical="center",
            )
            if col in (3, 5, 7):
                cell.number_format = "#,##0"
            elif col in (4, 6, 8):
                cell.number_format = "#,##0.00"
        sheet.row_dimensions[ridx].height = 23
    for col, width in {"A": 41, "B": 13, "C": 18, "D": 22,
                       "E": 18, "F": 22, "G": 18, "H": 22}.items():
        sheet.column_dimensions[col].width = width



def portfolio_summary_to_excel(table):
    """Export the same original standalone portfolio workbook and its styles."""
    book = Workbook()
    write_portfolio_sheet(book.active, table)
    output = BytesIO()
    book.save(output)
    return output.getvalue()


def _portfolio_html(table, rows_to_show=100):
    """A safe selectable table with grouped headings; all text is HTML-escaped."""
    shown = table.head(rows_to_show)
    cells = []
    for values in shown.itertuples(index=False, name=None):
        row = []
        for col, value in enumerate(values):
            if col >= 2:
                rendered = f"{int(value):,}" if col in (2, 4, 6) else f"{float(value):,.2f}"
            else:
                rendered = html_escape(str(value))
            row.append(f"<td>{rendered}</td>")
        cells.append("<tr>" + "".join(row) + "</tr>")
    return """<div class="ejada-portfolio-scroll" dir="ltr">
    <table class="ejada-portfolio-table">
      <thead>
        <tr><th rowspan="2">اسم المحصل</th><th rowspan="2">المنتج</th>
            <th colspan="2">DPD30</th><th colspan="2">DPD60</th>
            <th colspan="2">NPL</th></tr>
        <tr><th>عدد الحسابات</th><th>إجمالي المديونية</th>
            <th>عدد الحسابات</th><th>إجمالي المديونية</th>
            <th>عدد الحسابات</th><th>إجمالي المديونية</th></tr>
      </thead>
      <tbody>""" + "".join(cells) + """</tbody></table></div>"""


def render_portfolio_summary(portfolio_bytes):
    """Display a read-only, copyable portfolio view after upload alone."""
    try:
        table, diagnostics = build_portfolio_summary(portfolio_bytes)
    except Exception as exc:
        st.error(f"تعذر عرض جدول المحفظة: {exc}")
        return

    if diagnostics["excluded_stages"]:
        st.warning(
            f"يوجد {diagnostics['excluded_stages']:,} صف بمرحلة أخرى/فارغة "
            "خارج DPD30 وDPD60 وNPL؛ لم نضفه لهذه المراحل."
        )
    if diagnostics["missing_labels"]:
        st.warning(
            f"يوجد {diagnostics['missing_labels']:,} صف ضمن المراحل المطلوبة "
            "بدون اسم محصل أو منتج، ولم ننسبه إلى محصل آخر."
        )
    if diagnostics["blank_debts"]:
        st.caption(
            f"تم التعامل مع {diagnostics['blank_debts']:,} مبلغ فارغ كصفر "
            "وفق قاعدة Pivot المستخدمة حاليًا."
        )
    if table.empty:
        st.info("لا توجد صفوف مكتملة ضمن المراحل الثلاث لعرض الجدول.")
        return

    st.markdown(
        """<style>
        .ejada-portfolio-scroll {overflow:auto;max-width:100%;border:1px solid #A9BDB6;border-radius:8px;}
        .ejada-portfolio-table {width:100%;border-collapse:collapse;min-width:950px;
            background:#FFFFFF;color:#101820;font:15px Arial,Tahoma,sans-serif;}
        .ejada-portfolio-table th,.ejada-portfolio-table td {
            border:1px solid #A9BDB6;padding:10px 13px;text-align:center;white-space:nowrap;
        }
        .ejada-portfolio-table thead tr:first-child th {
            background:#075143;color:white;font-weight:800;height:45px;
        }
        .ejada-portfolio-table thead tr:nth-child(2) th {
            background:#E0EFE9;color:#064B3D;font-weight:800;height:38px;
        }
        .ejada-portfolio-table tbody td:first-child {text-align:right;direction:rtl;}
        .ejada-portfolio-table tbody tr:hover {background:#EDF8F1;}
        </style>""",
        unsafe_allow_html=True,
    )
    st.markdown(_portfolio_html(table), unsafe_allow_html=True)
    if len(table) > 100:
        st.caption(
            "المعاينة أعلى الصفحة تعرض أول 100 صف لحماية سرعة التطبيق. "
            "كل الصفوف موجودة في ملف Excel للتحميل والنسخ."
        )

    st.download_button(
        "📥 تحميل جدول المحفظة Excel بنفس شكل الأعمدة",
        data=portfolio_summary_to_excel(table),
        file_name="SNB_Portfolio_By_Collector_Product_Stages.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="snb_perf_portfolio_export",
        use_container_width=False,
    )
    return table
