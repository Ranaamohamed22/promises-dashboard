"""Shared SNB payment parsing helpers. Keep business semantics unchanged.

Extracted without edits from snb_performance.py so calculation behavior
remains identical while the module becomes smaller and easier to test.
"""

import pandas as pd

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

