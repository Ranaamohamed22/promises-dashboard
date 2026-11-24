"""Synthetic tests for the helpers extracted from SNB performance."""
import unittest
from datetime import date
import pandas as pd
from analytics_core.snb_performance_utils import _account_key, _amounts, _creation_dates
from snb_performance import (
    _account_key as _legacy_account_key,
    _amounts as _legacy_amounts,
    _creation_dates as _legacy_creation_dates,
)

class TestSnbHelpers(unittest.TestCase):
    def test_legacy_names_reference_extracted_functions(self):
        self.assertIs(_legacy_account_key, _account_key)
        self.assertIs(_legacy_amounts, _amounts)
        self.assertIs(_legacy_creation_dates, _creation_dates)

    def test_account_keys_preserve_string_leading_zeros(self):
        pairs = [
            (None, ""),
            (123, "123"),
            (123.0, "123"),
            (" 000123 ", "000123"),
            ("000123.0", "000123"),
            ("", ""),
        ]
        for original, expected in pairs:
            with self.subTest(original=original):
                self.assertEqual(_account_key(original), expected)

    def test_amount_normalization_and_blanks(self):
        values = pd.Series([" 1,200.50 ", "2٬000", "", None, "-3"])
        self.assertEqual(_amounts(values, "Payment").tolist(), [1200.5, 2000, 0, 0, -3])

    def test_invalid_amount_fails_closed(self):
        with self.assertRaisesRegex(ValueError, "Payment"):
            _amounts(pd.Series(["100", "invalid"]), "Payment")

    def test_excel_serials_text_dates_and_blank_dates(self):
        dates = _creation_dates(pd.Series([45292, "45292", "31/12/2023", None]))
        self.assertEqual(dates.iloc[0].date(), date(2024, 1, 1))
        self.assertEqual(dates.iloc[1].date(), date(2024, 1, 1))
        self.assertEqual(dates.iloc[2].date(), date(2023, 12, 31))
        self.assertTrue(pd.isna(dates.iloc[3]))

if __name__ == "__main__":
    unittest.main()
