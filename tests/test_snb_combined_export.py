"""Regression checks for the optional two-sheet SNB performance download.

Synthetic figures only: never commit bank or customer data.
"""
from datetime import datetime
from io import BytesIO
from zoneinfo import ZoneInfo
import inspect
import unittest

import pandas as pd
from openpyxl import load_workbook, Workbook

from snb_performance import (
    build_snb_performance,
    build_combined_snb_performance_workbook,
)
from snb_portfolio_tab import (
    build_portfolio_summary,
    portfolio_summary_to_excel,
    render_portfolio_summary,
)


def excel_bytes(frame):
    stream = BytesIO()
    frame.to_excel(stream, index=False, engine="openpyxl")
    return stream.getvalue()


class CombinedSNBPerformanceTests(unittest.TestCase):
    def setUp(self):
        today = datetime.now(ZoneInfo("Africa/Cairo")).date().isoformat()
        self.portfolio = pd.DataFrame({
            "Account Number": ["0001", "0002", "0003"],
            "حالة المعالجة - التمويل": ["معالج", "غير معالج", "غير معالج"],
            "مبلغ المعالجة": [100, 200, 400],
            "Salesperson": ["Collector A", "Collector A", "Collector B"],
            "نوع المتنج-التمويل": ["CC", "PF", "RF"],
            "ملاحظات-التمويل": ["DPD30", "DPD60", "NPL"],
            "Net Amount": [1234.56, 2500.75, 4800.00],
        })
        self.payments = pd.DataFrame({
            "Customer Account Number": ["0002", "0002", "0003"],
            "Payment": [150, 50, 25],
            "Date of Creation": [today, today, today],
        })

    def test_one_workbook_has_two_tabs_and_identical_original_values(self):
        source_portfolio = excel_bytes(self.portfolio)
        original = build_snb_performance(
            source_portfolio, excel_bytes(self.payments)
        )
        summary, diagnostics = build_portfolio_summary(source_portfolio)
        self.assertEqual(diagnostics["included_rows"], 3)
        self.assertEqual(original["original_count"], 1)
        self.assertEqual(original["payment_count"], 1)

        combined_bytes = build_combined_snb_performance_workbook(
            original["processed_bytes"], summary
        )
        combined_book = load_workbook(BytesIO(combined_bytes))
        self.assertEqual(combined_book.sheetnames, ["المعالج", "المحفظة"])

        old_processed = load_workbook(BytesIO(original["processed_bytes"])).active
        new_processed = combined_book["المعالج"]
        self.assertEqual(old_processed.max_row, new_processed.max_row)
        self.assertEqual(old_processed.max_column, new_processed.max_column)
        for old_row, new_row in zip(old_processed.iter_rows(), new_processed.iter_rows()):
            for old, new in zip(old_row, new_row):
                self.assertEqual(old.value, new.value)
                self.assertEqual(old.number_format, new.number_format)
                self.assertEqual(old.fill.fgColor.rgb, new.fill.fgColor.rgb)

        old_portfolio = load_workbook(BytesIO(
            portfolio_summary_to_excel(summary)
        )).active
        new_portfolio = combined_book["المحفظة"]
        self.assertEqual(old_portfolio.max_row, new_portfolio.max_row)
        self.assertEqual(old_portfolio.max_column, new_portfolio.max_column)
        for old_row, new_row in zip(old_portfolio.iter_rows(),
                                    new_portfolio.iter_rows()):
            for old, new in zip(old_row, new_row):
                self.assertEqual(old.value, new.value)
                self.assertEqual(old.number_format, new.number_format)
                self.assertEqual(old.fill.fgColor.rgb, new.fill.fgColor.rgb)
        self.assertEqual(new_portfolio["C1"].value, "DPD30")
        self.assertEqual(new_portfolio["E1"].value, "DPD60")
        self.assertEqual(new_portfolio["G1"].value, "NPL")
        self.assertFalse(new_portfolio.sheet_view.rightToLeft)
        self.assertEqual(new_portfolio.freeze_panes, "C3")

        # Old daily-payments output remains a separate, unchanged workbook.
        daily = load_workbook(BytesIO(original["daily_bytes"]))
        self.assertEqual(daily.sheetnames, ["سدادات اليوم"])

    def test_combining_wrong_processed_sheet_fails_safely(self):
        wb = Workbook()
        wb.active.title = "Wrong Sheet"
        file = BytesIO()
        wb.save(file)
        with self.assertRaisesRegex(ValueError, "المعالج"):
            build_combined_snb_performance_workbook(
                file.getvalue(), pd.DataFrame()
            )

    def test_deleted_full_table_expander_is_not_rendered(self):
        source = inspect.getsource(render_portfolio_summary)
        self.assertNotIn("جدول المحفظة كامل — تحديد ونسخ للخلايا", source)
        self.assertNotIn("with st.expander(", source)


if __name__ == "__main__":
    unittest.main()
