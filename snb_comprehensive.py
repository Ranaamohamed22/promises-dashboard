"""SNB comprehensive report: upload a source workbook and apply ordered bucket steps."""
from copy import copy
from io import BytesIO
import re

import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# Add future report steps here in sequence; existing reports are independent.
COMPREHENSIVE_STEPS = (
    ("BN", "latest_bucket"),
    ("B-1", "previous_1"),
    ("B-2", "previous_2"),
    ("B-3", "previous_3"),
    ("B1", "first_bucket"),
    ("B1 - BN", "bucket_difference"),
)


def _find_bucket_columns(sheet):
    headers = [
        str(sheet.cell(1, index).value or "").strip()
        for index in range(1, sheet.max_column + 1)
    ]
    first = next(
        (i + 1 for i, title in enumerate(headers) if title.upper() == "DPD_BUCKET"),
        None,
    )
    if first is None:
        first = next(
            (i + 1 for i, title in enumerate(headers)
             if "BUCKET" in title.upper() and not re.fullmatch(r"BUCKET\s*\d+", title, re.I)),
            None,
        )
    historical = [
        i + 1 for i, title in enumerate(headers)
        if re.fullmatch(r"BUCKET\s*\d+", title, re.I)
    ]
    if first is None or not historical:
        raise ValueError("لازم الملف يحتوي على أول عمود DPD_BUCKET وعمود تاريخي BUCKET برقم.")
    return first, historical[-1]


def build_comprehensive_report(source_bytes):
    """Apply the ordered steps afresh to the uploaded workbook; no stored data."""
    workbook = load_workbook(BytesIO(source_bytes))
    sheet = workbook.active
    first_bucket, last_bucket = _find_bucket_columns(sheet)
    titles = [step[0] for step in COMPREHENSIVE_STEPS]
    # If the user uploads a previously processed export, update its six columns
    # instead of appending a new copy on each run.
    headers = [
        str(sheet.cell(1, idx).value or "").strip()
        for idx in range(1, sheet.max_column + 1)
    ]
    if headers[-len(titles):] == titles:
        start = sheet.max_column - len(titles) + 1
    else:
        start = sheet.max_column + 1
    last_bucket_ref = get_column_letter(last_bucket)
    first_bucket_ref = get_column_letter(first_bucket)
    letters = [get_column_letter(start + i) for i in range(len(titles))]
    source_style = sheet.cell(1, last_bucket)
    for offset, (title, _) in enumerate(COMPREHENSIVE_STEPS):
        col = start + offset
        header = sheet.cell(1, col, title)
        if source_style.has_style:
            header._style = copy(source_style._style)
        else:
            header.fill = PatternFill("solid", fgColor="1F4373")
            header.font = Font(name="Arial", color="FFFFFF", bold=True)
        header.alignment = Alignment(horizontal="center", vertical="center")
        sheet.column_dimensions[letters[offset]].width = max(12, len(title) + 5)
    for row in range(2, sheet.max_row + 1):
        formulas = (
            f'=IFERROR(LEFT({last_bucket_ref}{row},2),"")',
            f'=IF({letters[0]}{row}="","",VALUE({letters[0]}{row})-1)',
            f'=IF({letters[1]}{row}="","",{letters[1]}{row}-1)',
            f'=IF({letters[2]}{row}="","",{letters[2]}{row}-1)',
            f'=IFERROR(LEFT({first_bucket_ref}{row},2),"")',
            f'=IF(OR({letters[4]}{row}="",{letters[0]}{row}=""),"",VALUE({letters[4]}{row})-VALUE({letters[0]}{row}))',
        )
        for offset, formula in enumerate(formulas):
            target = sheet.cell(row, start + offset)
            target.value = formula
            if sheet.cell(row, last_bucket).has_style:
                target._style = copy(sheet.cell(row, last_bucket)._style)
            target.number_format = "General" if offset in (0, 4) else "0;-0;0"
            target.alignment = Alignment(horizontal="center", vertical="center")
    # Excel evaluates formulas when the downloaded file is opened.
    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
    except AttributeError:
        pass
    sheet.auto_filter.ref = f"A1:{get_column_letter(max(sheet.max_column, start + 5))}{sheet.max_row}"
    output = BytesIO()
    workbook.save(output)
    return output.getvalue(), sheet.max_row - 1, (
        sheet.cell(1, first_bucket).value, sheet.cell(1, last_bucket).value
    )


def render_snb_comprehensive_page():
    st.markdown("### 📋 التقرير الشامل — البنك الأهلي السعودي")
    st.caption("ارفعي التقرير الأصلي لتطبيق خطوات BUCKET بالترتيب. يمكن إضافة خطوات جديدة للتقرير لاحقًا.")
    uploaded = st.file_uploader(
        "رفع ملف التقرير الشامل (Excel)", type=["xlsx"], key="snb_comprehensive_source"
    )
    st.markdown("**الخطوات الحالية:**")
    st.markdown(
        "1. BN: أول خانتين من آخر BUCKET  \n"
        "2. B-1: قيمة BN ناقص 1  \n"
        "3. B-2: قيمة B-1 ناقص 1  \n"
        "4. B-3: قيمة B-2 ناقص 1  \n"
        "5. B1: أول خانتين من أول DPD_BUCKET  \n"
        "6. B1 - BN: طرح BN من B1"
    )
    if uploaded is None:
        st.info("ارفعي الملف الأصلي لبدء معالجة التقرير الشامل.")
        return
    try:
        output, count, bucket_names = build_comprehensive_report(uploaded.getvalue())
    except Exception as error:
        st.error(f"تعذر معالجة التقرير الشامل: {error}")
        return
    st.success(f"تم تجهيز التقرير لـ {count:,} صف.")
    st.caption(f"أول BUCKET: {bucket_names[0]} | آخر BUCKET: {bucket_names[1]}")
    st.download_button(
        "📥 تحميل التقرير الشامل بعد المعالجة",
        data=output,
        file_name="SNB_Comprehensive_Processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="snb_comprehensive_download",
        use_container_width=True,
    )
