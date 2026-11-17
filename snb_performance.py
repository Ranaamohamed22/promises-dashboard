"""Independent SNB performance report: portfolio treatment and daily payments."""
from datetime import datetime
from hashlib import sha256
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


PORTFOLIO_REQUIRED = (
    "Account Number", "حالة المعالجة - التمويل", "مبلغ المعالجة"
)
PAYMENTS_REQUIRED = (
    "Customer Account Number", "Payment", "Date of Creation"
)


def _account_key(value):
    """Normalize matching keys without stripping meaningful leading zeros."""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    value = str(value).strip()
    if value.endswith(".0") and value[:-2].isdigit():
        return value[:-2]
    return value


def _amounts(series, label):
    text = series.astype("string").str.strip()
    text = text.str.replace(",", "", regex=False).str.replace("٬", "", regex=False)
    empty = text.isna() | text.eq("")
    numeric = pd.to_numeric(text.mask(empty), errors="coerce")
    bad = ~empty & numeric.isna()
    if bad.any():
        raise ValueError(
            f"عمود {label} يحتوي على {int(bad.sum())} قيمة غير رقمية. "
            "راجعي القيم قبل تشغيل التقرير."
        )
    return numeric.fillna(0).astype(float)


def _creation_dates(series):
    """Handle Excel dates, Excel serials, and common text date formats."""
    def parse(value):
        if pd.isna(value) or str(value).strip() == "":
            return pd.NaT
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            if 20000 <= float(value) <= 90000:
                return pd.Timestamp("1899-12-30") + pd.to_timedelta(
                    float(value), unit="D"
                )
        raw = str(value).strip()
        if raw.isdigit() and len(raw) == 5 and 20000 <= int(raw) <= 90000:
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(
                int(raw), unit="D"
            )
        return pd.to_datetime(value, errors="coerce", dayfirst=True)
    return series.map(parse)


def _export_excel(df, sheet_name, highlight_payment_treated=False):
    stream = BytesIO()
    df.to_excel(stream, sheet_name=sheet_name, index=False, engine="openpyxl")
    stream.seek(0)
    book = load_workbook(stream)
    sheet = book.active
    sheet.freeze_panes = "A2"
    if sheet.max_column:
        sheet.auto_filter.ref = (
            f"A1:{get_column_letter(sheet.max_column)}{max(sheet.max_row, 1)}"
        )
    header_fill = PatternFill("solid", fgColor="1F4373")
    green_fill = PatternFill("solid", fgColor="E2F3E5")
    for cell in sheet[1]:
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        # pandas assigns placeholder names to blank Excel header cells.
        if isinstance(cell.value, str) and cell.value.startswith("Unnamed: "):
            cell.value = None
    sheet.row_dimensions[1].height = 26
    for col_index in range(1, sheet.max_column + 1):
        col_letter = get_column_letter(col_index)
        values = [
            len(str(sheet.cell(row, col_index).value or ""))
            for row in range(1, min(sheet.max_row, 50) + 1)
        ]
        sheet.column_dimensions[col_letter].width = min(
            42, max(14, (max(values) if values else 14) + 3)
        )
    if highlight_payment_treated and "نوع المعالجة" in df.columns:
        type_column = list(df.columns).index("نوع المعالجة") + 1
        for row_index in range(2, sheet.max_row + 1):
            cell = sheet.cell(row_index, type_column)
            if cell.value == "المعالج بالسداد":
                cell.fill = green_fill
    output = BytesIO()
    book.save(output)
    return output.getvalue()


def build_snb_performance(portfolio_bytes, payments_bytes):
    """Build solely from the latest uploaded byte snapshots; never read old files."""
    portfolio = pd.read_excel(BytesIO(portfolio_bytes), dtype=object)
    payments = pd.read_excel(BytesIO(payments_bytes), dtype=object)
    for name, frame, required in (
        ("المحفظة", portfolio, PORTFOLIO_REQUIRED),
        ("السدادات", payments, PAYMENTS_REQUIRED),
    ):
        missing = [col for col in required if col not in frame.columns]
        if missing:
            raise ValueError(
                f"ملف {name} ناقص منه الأعمدة الأساسية: " + "، ".join(missing)
            )
    if "نوع المعالجة" in portfolio.columns:
        raise ValueError(
            "المحفظة تحتوي بالفعل على عمود «نوع المعالجة»؛ "
            "ارفعي المحفظة الأصلية قبل إنشاء التقرير."
        )

    account_keys = portfolio["Account Number"].map(_account_key)
    payment_keys = payments["Customer Account Number"].map(_account_key)
    paid = _amounts(payments["Payment"], "Payment")
    status = (
        portfolio["حالة المعالجة - التمويل"]
        .astype("string").fillna("").str.strip()
    )
    original_mask = status.isin(("معالج", "المعالج"))
    untreated_mask = status.isin(("غير معالج", "غير المعالج"))

    # Original processed accounts are included regardless of their amount cells.
    # Untreated rows with missing/non-numeric amounts cannot be safely compared:
    # never treat them as zero, and report their actual values from this upload.
    untreated_amount_text = (
        portfolio.loc[untreated_mask, "مبلغ المعالجة"]
        .astype("string").str.strip()
        .str.replace(",", "", regex=False)
        .str.replace("٬", "", regex=False)
    )
    untreated_amount_numeric = pd.to_numeric(
        untreated_amount_text, errors="coerce"
    )
    invalid_treatment_mask = untreated_amount_numeric.isna()
    invalid_treatment_rows = [
        int(idx) + 2
        for idx in untreated_amount_numeric.index[invalid_treatment_mask]
    ]
    invalid_treatment_values = [
        str(value) if pd.notna(value) else "(فارغ)"
        for value in portfolio.loc[
            untreated_amount_numeric.index[invalid_treatment_mask],
            "مبلغ المعالجة",
        ].head(5)
    ]
    treatment_amount = pd.Series(float("nan"), index=portfolio.index)
    treatment_amount.loc[untreated_mask] = untreated_amount_numeric.astype(float)
    totals = (
        pd.DataFrame({"account": payment_keys, "amount": paid})
        .query("account != ''")
        .groupby("account", sort=False)["amount"].sum()
    )
    collected_by_account = account_keys.map(totals).fillna(0)
    paid_mask = (
        untreated_mask
        & collected_by_account.gt(0)
        & collected_by_account.ge(treatment_amount)
    )
    original_rows = portfolio.loc[original_mask].copy()
    original_rows["نوع المعالجة"] = "المعالج الأصلي"
    payment_rows = portfolio.loc[paid_mask].copy()
    payment_rows["نوع المعالجة"] = "المعالج بالسداد"
    # concat preserves the original relative row order inside each group.
    processed = pd.concat([original_rows, payment_rows], ignore_index=True)

    today = datetime.now(ZoneInfo("Africa/Cairo")).date()
    today_mask = _creation_dates(payments["Date of Creation"]).dt.date.eq(today)
    # Boolean filtering preserves every matching payment row and original order.
    daily_payments = payments.loc[today_mask].copy()
    daily_total = float(paid.loc[today_mask].sum())
    return {
        "today": today,
        "original_count": len(original_rows),
        "payment_count": len(payment_rows),
        "processed_count": len(processed),
        "daily_total": daily_total,
        "daily_count": len(daily_payments),
        "invalid_treatment_rows": invalid_treatment_rows,
        "invalid_treatment_values": invalid_treatment_values,
        "processed_bytes": _export_excel(
            processed, "المعالج", highlight_payment_treated=True
        ),
        "daily_bytes": _export_excel(daily_payments, "سدادات اليوم"),
    }


def render_snb_performance_page():
    # Resetting the uploader generation forces Streamlit to discard the
    # previously selected files, even if the replacement has the same name.
    if "snb_perf_upload_generation" not in st.session_state:
        st.session_state["snb_perf_upload_generation"] = 0
    if st.button("🔄 مسح الملفين ورفع ملفات جديدة", key="snb_perf_reset_uploads"):
        st.session_state["snb_perf_upload_generation"] += 1
        st.rerun()
    generation = st.session_state["snb_perf_upload_generation"]

    # Match the Activity page: keep the two upload controls in the top row,
    # rather than pushing them down with a full-width title and caption.
    portfolio_col, payments_col, heading_col = st.columns(
        [1.2, 1.2, 1.0], gap="small", vertical_alignment="center"
    )
    with portfolio_col:
        portfolio_file = st.file_uploader(
            "رفع ملف المحفظة", type=["xlsx"], key=f"snb_perf_portfolio_{generation}"
        )
    with payments_col:
        payments_file = st.file_uploader(
            "رفع ملف السدادات", type=["xlsx"], key=f"snb_perf_payments_{generation}"
        )
    with heading_col:
        st.markdown("**📊 تقرير الأداء — SNB**")
        st.caption("ارفعي المحفظة والسدادات معًا لتحديث التقرير تلقائيًا.")
    if portfolio_file is None or payments_file is None:
        st.info("ارفعي ملف المحفظة وملف السدادات معًا لتشغيل تقرير الأداء.")
        return
    # Snapshot the *currently selected* uploader content on every Streamlit
    # rerun. This path uses no old sample files, filesystem reads or data cache.
    portfolio_bytes = portfolio_file.getvalue()
    payments_bytes = payments_file.getvalue()
    portfolio_version = sha256(portfolio_bytes).hexdigest()[:10]
    payments_version = sha256(payments_bytes).hexdigest()[:10]
    st.caption(
        f"الملفات المستخدمة الآن: المحفظة «{portfolio_file.name}» "
        f"(نسخة {portfolio_version}) | السدادات «{payments_file.name}» "
        f"(نسخة {payments_version})"
    )
    try:
        result = build_snb_performance(portfolio_bytes, payments_bytes)
    except Exception as exc:
        st.error(f"تعذر تجهيز تقرير الأداء: {exc}")
        return

    # Show exactly the three requested count KPIs. Non-numeric thresholds
    # remain excluded from comparisons without an on-page warning banner.
    cols = st.columns(3, gap="small")
    cols[0].metric("عدد المعالج من المحفظة", f"{result['original_count']:,}")
    cols[1].metric("عدد المعالج من السداد", f"{result['payment_count']:,}")
    cols[2].metric("عدد سدادات اليوم", f"{result['daily_count']:,}")
    today_text = result["today"].strftime("%d/%m/%Y")
    today_filename = result["today"].isoformat()
    download_a, download_b = st.columns(2, gap="medium")
    with download_a:
        st.download_button(
            "📥 تحميل تقرير المعالج",
            data=result["processed_bytes"],
            file_name=f"SNB_Processed_{today_filename}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="snb_perf_download_processed",
            use_container_width=True,
        )
    with download_b:
        st.markdown(f"**تاريخ سدادات اليوم بتوقيت مصر: {today_text}**")
        st.download_button(
            "📥 تحميل سدادات اليوم",
            data=result["daily_bytes"],
            file_name=f"SNB_Daily_Payments_{today_filename}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="snb_perf_download_daily",
            use_container_width=True,
        )
        if result["daily_count"] == 0:
            st.info("لا توجد سدادات بتاريخ اليوم")
