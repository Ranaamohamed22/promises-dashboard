"""Regression tests for SNB performance reporting (synthetic data only)."""
import unittest
from datetime import datetime
from io import BytesIO
from zoneinfo import ZoneInfo
import pandas as pd
from snb_performance import _account_key, build_snb_performance

def to_excel_bytes(frame):
    output = BytesIO()
    frame.to_excel(output, index=False, engine="openpyxl")
    return output.getvalue()

class TestSnbPerformance(unittest.TestCase):
    def setUp(self):
        today = datetime.now(ZoneInfo("Africa/Cairo")).date().isoformat()
        self.portfolio = pd.DataFrame({
            "Account Number": ["0001", "0002", "0003", "0004", "0005"],
            "حالة المعالجة - التمويل": [
                "معالج", "غير معالج", "غير معالج", "غير معالج", "غير معالج"
            ],
            "مبلغ المعالجة": [100, 120, 100, "not numeric", 0],
        })
        self.payments = pd.DataFrame({
            "Customer Account Number": ["0002", "0002", "0003", "0004", "0005"],
            "Payment": [50, 70, 80, 200, 0],
            "Date of Creation": [
                today, today, "2020-01-01", "2020-01-01", "2020-01-01"
            ],
        })

    def test_aggregated_payments_and_processed_export(self):
        result = build_snb_performance(
            to_excel_bytes(self.portfolio), to_excel_bytes(self.payments))
        self.assertEqual(result["original_count"], 1)
        self.assertEqual(result["payment_count"], 1)
        self.assertEqual(result["processed_count"], 2)
        self.assertEqual(result["invalid_treatment_rows"], [5])
        self.assertEqual(result["daily_count"], 2)
        self.assertAlmostEqual(result["daily_total"], 120)
        processed = pd.read_excel(BytesIO(result["processed_bytes"]), dtype=str)
        self.assertEqual(processed["Account Number"].tolist(), ["0001", "0002"])
        self.assertEqual(processed["نوع المعالجة"].tolist(),
                         ["المعالج الأصلي", "المعالج بالسداد"])

    def test_matching_keys_preserves_leading_zeros(self):
        self.assertEqual(_account_key("000123"), "000123")
        self.assertEqual(_account_key("123.0"), "123")
        self.assertEqual(_account_key(123.0), "123")
        self.assertEqual(_account_key(None), "")

    def test_missing_payment_column_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "Payment"):
            build_snb_performance(
                to_excel_bytes(self.portfolio),
                to_excel_bytes(self.payments.drop(columns=["Payment"])))

    def test_nonnumeric_payment_is_rejected(self):
        altered = self.payments.copy()
        altered.loc[0, "Payment"] = "not a number"
        with self.assertRaisesRegex(ValueError, "Payment"):
            build_snb_performance(
                to_excel_bytes(self.portfolio), to_excel_bytes(altered))

if __name__ == "__main__":
    unittest.main()
