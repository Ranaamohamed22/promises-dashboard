"""Presentation-only metrics for the isolated SNB performance design preview.

The existing build_snb_performance function remains the sole authority for
treatment classification, payment matching and exported workbooks.
"""
from io import BytesIO

import pandas as pd

from snb_performance import _amounts, _creation_dates, build_snb_performance


PRODUCT_COLUMNS = ("PRODUCT_CATEGORY", "Product", "PRODUCT", "المنتج")
COLLECTOR_COLUMNS = ("Collector", "COLLECTOR", "Collector Name", "COLLECTOR_NAME", "اسم المحصل", "المحصل")
DISPLAY_COLUMNS = {
    "Account No": ("Account Number", "Account No", "ACCOUNT_NUMBER"),
    "Product": PRODUCT_COLUMNS,
    "Collector": COLLECTOR_COLUMNS,
    "Amount": ("مبلغ المعالجة", "Treatment Amount", "Amount"),
    "Status": ("نوع المعالجة", "حالة المعالجة - التمويل"),
}


def _first_column(frame, candidates):
    """Only select columns actually present; never guess their contents."""
    return next((name for name in candidates if name in frame.columns), None)


def _breakdown(source, processed, candidates, label):
    name = _first_column(source, candidates)
    if name is None or name not in processed.columns:
        return pd.DataFrame(columns=[label, "إجمالي الصفوف", "الصفوف المعالجة"])
    all_groups = source[name].astype("string").fillna("غير محدد").str.strip().replace("", "غير محدد")
    done_groups = processed[name].astype("string").fillna("غير محدد").str.strip().replace("", "غير محدد")
    all_counts = all_groups.value_counts(dropna=False)
    done_counts = done_groups.value_counts(dropna=False)
    labels = list(dict.fromkeys(list(all_counts.index) + list(done_counts.index)))
    return pd.DataFrame({
        label: labels,
        "إجمالي الصفوف": [int(all_counts.get(name, 0)) for name in labels],
        "الصفوف المعالجة": [int(done_counts.get(name, 0)) for name in labels],
    })


def _daily_payments(payments):
    """Summarize actual payment totals by creation date, not treatment events."""
    if payments.empty:
        return pd.DataFrame(columns=["التاريخ", "مبلغ السداد"])
    created = _creation_dates(payments["Date of Creation"])
    paid = _amounts(payments["Payment"], "Payment")
    valid = created.notna()
    if not valid.any():
        return pd.DataFrame(columns=["التاريخ", "مبلغ السداد"])
    summary = pd.DataFrame({
        "التاريخ": created.loc[valid].dt.date.to_numpy(),
        "مبلغ السداد": paid.loc[valid].to_numpy(),
    })
    return summary.groupby("التاريخ", as_index=False, sort=True)["مبلغ السداد"].sum()


def _detail_table(processed):
    selected = {}
    for label, candidates in DISPLAY_COLUMNS.items():
        found = _first_column(processed, candidates)
        if found is not None:
            selected[found] = label
    if selected:
        return processed[list(selected)].rename(columns=selected).copy()
    return processed.copy()


def prepare_performance_preview(portfolio_bytes, payments_bytes):
    """Return a view over the unchanged business report and its exact exports.

    Counts and percentages here are explicitly ROW-based, not unique accounts.
    No uploaded content is saved to disk or to persistent application storage.
    """
    result = build_snb_performance(portfolio_bytes, payments_bytes)
    portfolio = pd.read_excel(BytesIO(portfolio_bytes), dtype=object)
    payments = pd.read_excel(BytesIO(payments_bytes), dtype=object)
    processed = pd.read_excel(BytesIO(result["processed_bytes"]), dtype=object)
    total_rows = len(portfolio)
    processed_rows = int(result["processed_count"])
    return {
        "report": result,
        "total_rows": total_rows,
        "processed_rows": processed_rows,
        "remaining_rows": max(0, total_rows - processed_rows),
        "rate": (100.0 * processed_rows / total_rows) if total_rows else None,
        "products": _breakdown(portfolio, processed, PRODUCT_COLUMNS, "المنتج"),
        "collectors": _breakdown(portfolio, processed, COLLECTOR_COLUMNS, "المحصل"),
        "payments_by_date": _daily_payments(payments),
        "details": _detail_table(processed),
        "processed_full": processed,
    }
