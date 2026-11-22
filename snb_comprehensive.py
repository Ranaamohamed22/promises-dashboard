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
    ("BN - B1", "bucket_difference"),
    ("حالة المعالجة", "processing_status"),
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
        raise ValueError("لازم الملف يحتوي على أول عمود DPD_BUCKET.")
    # Only numbered BUCKET periods with a space belong to the historical
    # tracking sequence (BUCKET 07 ... BUCKET 21).
    # Bucket_1 ... Bucket_6 are a DIFFERENT set of fields and must never
    # participate in the treatment status comparison.
    historical = [
        (index + 1, int(re.search(r"\d+$", title).group()))
        for index, title in enumerate(headers)
        if re.fullmatch(r"BUCKET\s+\d+", title, re.I)
    ]
    if not historical:
        raise ValueError("لازم الملف يحتوي على أعمدة BUCKET التاريخية.")
    ordered = [col for col, _ in sorted(historical, key=lambda item: (item[1], item[0]))]
    return first, ordered[-1], [first] + ordered


def _bucket_number(value):
    """Take the initial two-digit BUCKET code, preserving invalid data as invalid."""
    if value is None:
        return None
    match = re.match(r"^\s*(\d{2})(?:\s*[-–]|\b)", str(value))
    return int(match.group(1)) if match else None


def _is_writeoff(value):
    text = str(value or "").strip()
    return bool(
        re.match(r"^29\s*[-–]\s*999\b", text, re.I)
        or re.match(r"^writeoff\b", text, re.I)
    )


def _is_treated(sheet, row, history_columns, first_number, latest_number):
    """First-to-last improvement OR any intermediate chronological decrease.

    Ignore invalid intermediate codes. An invalid first or latest BUCKET
    (including Writeoff) yields #VALUE! in the difference and 'غير معالج'.
    """
    if first_number is None or latest_number is None:
        return False
    if latest_number - first_number < 0:
        return True
    highest_prior = None
    for col in history_columns:
        number = _bucket_number(sheet.cell(row, col).value)
        if number is None:
            continue
        if highest_prior is not None and number < highest_prior:
            return True
        highest_prior = number if highest_prior is None else max(highest_prior, number)
    return False


def build_comprehensive_report(source_bytes):
    """Generate a fresh SNB workbook with original data and six calculated columns."""
    workbook = load_workbook(BytesIO(source_bytes))
    sheet = workbook.active
    first_bucket, last_bucket, bucket_history = _find_bucket_columns(sheet)
    titles = [title for title, _ in COMPREHENSIVE_STEPS]

    # Repeated exports must not accumulate extra columns. Older outputs used a
    # visible B1 and a reversed difference; drop those legacy tail columns.
    headers = [
        str(sheet.cell(1, i).value or "").strip()
        for i in range(1, sheet.max_column + 1)
    ]
    legacy_with_status = ["BN", "B-1", "B-2", "B-3", "B1", "B1 - BN", "حالة المعالجة"]
    legacy_without_status = legacy_with_status[:-1]
    if headers[-len(titles):] == titles:
        start = sheet.max_column - len(titles) + 1
    elif headers[-len(legacy_with_status):] == legacy_with_status:
        start = sheet.max_column - len(legacy_with_status) + 1
        sheet.delete_cols(start, len(legacy_with_status))
    elif headers[-len(legacy_without_status):] == legacy_without_status:
        start = sheet.max_column - len(legacy_without_status) + 1
        sheet.delete_cols(start, len(legacy_without_status))
    else:
        start = sheet.max_column + 1

    letters = [get_column_letter(start + offset) for offset in range(len(titles))]
    last_ref = get_column_letter(last_bucket)
    first_ref = get_column_letter(first_bucket)
    header_style = sheet.cell(1, last_bucket)
    status_counts = {"معالج": 0, "غير معالج": 0}

    for offset, (title, _) in enumerate(COMPREHENSIVE_STEPS):
        cell = sheet.cell(1, start + offset, title)
        if header_style.has_style:
            cell._style = copy(header_style._style)
        else:
            cell.fill = PatternFill("solid", fgColor="1F4373")
            cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        sheet.column_dimensions[letters[offset]].width = max(14, len(title) + 5)

    for row in range(2, sheet.max_row + 1):
        latest_cell = sheet.cell(row, last_bucket)
        if _is_writeoff(latest_cell.value):
            latest_cell.value = "Writeoff"

        b1 = _bucket_number(sheet.cell(row, first_bucket).value)
        bn = None if _is_writeoff(latest_cell.value) else _bucket_number(latest_cell.value)
        status = (
            "معالج"
            if _is_treated(sheet, row, bucket_history, b1, bn)
            else "غير معالج"
        )
        status_counts[status] += 1
        # B1 is intentionally NOT a visible column: refer to the original
        # first bucket directly in the BN-B1 formula.
        formulas = (
            f"=LEFT({last_ref}{row},2)",
            f"=VALUE({letters[0]}{row})-1",
            f"={letters[1]}{row}-1",
            f"={letters[2]}{row}-1",
            f"=VALUE({letters[0]}{row})-VALUE(LEFT({first_ref}{row},2))",
        )
        for offset in range(5):
            target = sheet.cell(row, start + offset)
            # Do not leave blanks or misleading cached numbers when BN is Wr:
            # Excel must show #VALUE! for invalid arithmetic.
            if offset == 4 and _is_writeoff(latest_cell.value):
                # The difference displays Writeoff only when the final bucket
                # is Writeoff; do not change the status or other columns.
                target.value = "Writeoff"
            elif (offset in (1, 2, 3) and bn is None) or (
                offset == 4 and (bn is None or b1 is None)
            ):
                target.value = "#VALUE!"
            else:
                target.value = formulas[offset]
            if latest_cell.has_style:
                target._style = copy(latest_cell._style)
            target.number_format = "General" if offset == 0 else "0;-0;0"
            target.alignment = Alignment(horizontal="center", vertical="center")
        status_cell = sheet.cell(row, start + 5, status)
        if latest_cell.has_style:
            status_cell._style = copy(latest_cell._style)
        status_cell.alignment = Alignment(horizontal="center", vertical="center")
        status_cell.number_format = "General"

    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
    except AttributeError:
        pass
    sheet.auto_filter.ref = (
        f"A1:{get_column_letter(max(sheet.max_column, start + 5))}{sheet.max_row}"
    )
    output = BytesIO()
    workbook.save(output)
    return output.getvalue(), sheet.max_row - 1, (
        sheet.cell(1, first_bucket).value, sheet.cell(1, last_bucket).value
    ), status_counts

def render_snb_comprehensive_page():
    st.markdown("### 📋 التقرير الشامل — البنك الأهلي السعودي")
    st.caption("ارفعي التقرير الأصلي لتطبيق خطوات BUCKET بالترتيب. يمكن إضافة خطوات جديدة للتقرير لاحقًا.")
    uploaded = st.file_uploader(
        "رفع ملف التقرير الشامل (Excel)", type=["xlsx"], key="snb_comprehensive_source"
    )
    st.markdown("**الخطوات الحالية:**")
    st.markdown(
        "1. BN: أول خانتين من آخر BUCKET، مع تحويل 29 - 999 إلى Writeoff  \n"
        "2. B-1: قيمة BN ناقص 1  \n"
        "3. B-2: قيمة B-1 ناقص 1  \n"
        "4. B-3: قيمة B-2 ناقص 1  \n"
        "5. BN - B1: آخر BUCKET ناقص أول BUCKET؛ B1 غير ظاهر، وWr ينتج #VALUE!  \n"
        "6. حالة المعالجة: معالج إذا الفرق سالب أو بوكت أحدث انخفض عن بوكت سابق"
    )
    if uploaded is None:
        st.info("ارفعي الملف الأصلي لبدء معالجة التقرير الشامل.")
        return
    try:
        output, count, bucket_names, statuses = build_comprehensive_report(uploaded.getvalue())
    except Exception as error:
        st.error(f"تعذر معالجة التقرير الشامل: {error}")
        return
    st.success(f"تم تجهيز التقرير لـ {count:,} صف. معالج: {statuses['معالج']:,} | غير معالج: {statuses['غير معالج']:,}")
    st.caption(f"أول BUCKET: {bucket_names[0]} | آخر BUCKET: {bucket_names[1]}")
    st.download_button(
        "📥 تحميل التقرير الشامل بعد المعالجة",
        data=output,
        file_name="SNB_Comprehensive_Processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="snb_comprehensive_download",
        use_container_width=True,
    )
