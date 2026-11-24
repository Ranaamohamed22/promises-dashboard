"""Synthetic, no-client-data checks for the isolated Streamlit performance preview."""
import unittest
from datetime import datetime
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd

from preview_reporting import prepare_performance_preview


def to_excel_bytes(dataframe):
    output = BytesIO()
    dataframe.to_excel(output, index=False, engine="openpyxl")
    return output.getvalue()


class PreviewTest(unittest.TestCase):
    def setUp(self):
        today = datetime.now(ZoneInfo("Africa/Cairo")).date().isoformat()
        self.portfolio = pd.DataFrame({
            "Account Number": ["0001", "0002", "0003", "0004"],
            "حالة المعالجة - التمويل": [
                "معالج", "غير معالج", "غير معالج", "غير معالج"
            ],
            "مبلغ المعالجة": [100, 120, 100, 100],
            "PRODUCT_CATEGORY": ["CC", "CC", "PF", "RF"],
            "Collector": ["X", "Y", "X", "Z"],
        })
        self.payments = pd.DataFrame({
            "Customer Account Number": ["0002", "0002", "0003"],
            "Payment": [50, 70, 20],
            "Date of Creation": [today, today, "2020-01-01"],
        })

    def test_preview_uses_original_business_results(self):
        view = prepare_performance_preview(
            to_excel_bytes(self.portfolio), to_excel_bytes(self.payments)
        )
        original = view["report"]
        self.assertEqual(view["total_rows"], 4)
        self.assertEqual(original["original_count"], 1)
        self.assertEqual(original["payment_count"], 1)
        self.assertEqual(original["processed_count"], 2)
        self.assertEqual(view["processed_rows"], 2)
        self.assertEqual(view["remaining_rows"], 2)
        self.assertEqual(view["rate"], 50.0)
        self.assertEqual(original["daily_count"], 2)
        self.assertAlmostEqual(original["daily_total"], 120.0)
        self.assertEqual(view["processed_full"]["Account Number"].astype(str).tolist(),
                         ["0001", "0002"])
        # All exported Excel files come directly from the existing report builder.
        exported = pd.read_excel(BytesIO(original["processed_bytes"]), dtype=object)
        pd.testing.assert_frame_equal(view["processed_full"], exported, check_dtype=False)
        self.assertTrue(original["daily_bytes"])
        self.assertEqual(view["details"]["Account No"].astype(str).tolist(), ["0001", "0002"])

    def test_breakdowns_reflect_uploaded_rows_not_guesses(self):
        view = prepare_performance_preview(
            to_excel_bytes(self.portfolio), to_excel_bytes(self.payments)
        )
        rows = view["products"].set_index("المنتج")
        self.assertEqual(int(rows.loc["CC", "إجمالي الصفوف"]), 2)
        self.assertEqual(int(rows.loc["CC", "الصفوف المعالجة"]), 2)
        self.assertEqual(int(rows.loc["PF", "الصفوف المعالجة"]), 0)
        collectors = view["collectors"].set_index("المحصل")
        self.assertEqual(int(collectors.loc["Y", "الصفوف المعالجة"]), 1)
        self.assertEqual(len(view["payments_by_date"]), 2)

    def test_optional_columns_do_not_invent_classification(self):
        without_optional = self.portfolio.drop(columns=["PRODUCT_CATEGORY", "Collector"])
        view = prepare_performance_preview(
            to_excel_bytes(without_optional), to_excel_bytes(self.payments)
        )
        self.assertTrue(view["products"].empty)
        self.assertTrue(view["collectors"].empty)
        self.assertEqual(view["rate"], 50.0)

    def test_original_validation_errors_are_preserved(self):
        broken = self.payments.copy()
        broken["Payment"] = broken["Payment"].astype(object)
        broken.loc[0, "Payment"] = "not numeric"
        with self.assertRaisesRegex(ValueError, "Payment"):
            prepare_performance_preview(
                to_excel_bytes(self.portfolio), to_excel_bytes(broken)
            )

    def test_empty_portfolio_has_no_false_percentage(self):
        view = prepare_performance_preview(
            to_excel_bytes(self.portfolio.iloc[:0]), to_excel_bytes(self.payments.iloc[:0])
        )
        self.assertEqual(view["total_rows"], 0)
        self.assertIsNone(view["rate"])
        self.assertEqual(view["processed_rows"], 0)

    def test_preview_renders_before_any_file_upload(self):
        # Smoke test of the actual Streamlit entrypoint, not only the data layer.
        from streamlit.testing.v1 import AppTest
        at = AppTest.from_file("performance_preview.py", default_timeout=30).run()
        self.assertEqual(len(at.exception), 0)


if __name__ == "__main__":
    unittest.main()
