"""Synthetic portfolio-grid regression; never reads real customer records."""
from io import BytesIO
import unittest

import pandas as pd
from openpyxl import load_workbook

from snb_portfolio_tab import (
    build_portfolio_summary,
    portfolio_summary_to_excel,
    _normalize_stage,
    STAGES,
    COLUMNS,
)


def workbook_bytes(frame):
    stream = BytesIO()
    frame.to_excel(stream, index=False, engine="openpyxl")
    return stream.getvalue()


class PortfolioSummaryTests(unittest.TestCase):
    def setUp(self):
        self.sample = pd.DataFrame({
            "Salesperson": ["Collector A", "Collector A", "Collector A",
                            "Collector B", "Collector B", "Collector C"],
            "PRODUCT_CATEGORY": ["CC", "CC", "PF", "CC", "CC", "PF"],
            "ملاحظات-التمويل": ["Dpd30", "DPD60", "NPL", "NPL", "npl", "OTHER"],
            "Account Number": ["001", "002", "003", "004", "005", "006"],
            "Net Amount": [2280401.12, 791736.35, 1967203.42,
                           102.5, 200.25, 90],
        })

    def test_grouping_and_product_order_match_reference(self):
        table, diag = build_portfolio_summary(workbook_bytes(self.sample))
        self.assertEqual(table.columns.tolist(), list(COLUMNS))
        self.assertEqual(STAGES, ("DPD30", "DPD60", "NPL"))
        self.assertEqual(table["اسم المحصل"].tolist(),
                         ["Collector A", "Collector A", "Collector B"])
        self.assertEqual(table["المنتج"].tolist(), ["CC", "PF", "CC"])
        alice_cc = table.iloc[0]
        self.assertEqual(alice_cc["DPD30 | عدد الحسابات"], 1)
        self.assertAlmostEqual(alice_cc["DPD30 | إجمالي المديونية"], 2280401.12)
        self.assertEqual(alice_cc["DPD60 | عدد الحسابات"], 1)
        self.assertAlmostEqual(alice_cc["DPD60 | إجمالي المديونية"], 791736.35)
        self.assertEqual(alice_cc["NPL | عدد الحسابات"], 0)
        self.assertEqual(table.iloc[1]["NPL | عدد الحسابات"], 1)
        self.assertEqual(table.iloc[2]["NPL | عدد الحسابات"], 2)
        self.assertAlmostEqual(table.iloc[2]["NPL | إجمالي المديونية"], 302.75)
        self.assertEqual(diag["source_rows"], 6)
        self.assertEqual(diag["included_rows"], 5)
        self.assertEqual(diag["excluded_stages"], 1)

    def test_real_crm_product_header_typo_is_supported(self):
        # Synthetic fixture with the exact column names in the uploaded
        # Sheet1. Never commit any real customer or bank account records.
        source = pd.DataFrame({
            "Salesperson": ["Collector A", "Collector A", "Collector B"],
            "نوع المتنج-التمويل": ["RF", "CC", "PF"],
            "ملاحظات-التمويل": ["DPD30", "NPL", "DPD60"],
            "Account Number": ["TEST001", "TEST002", "TEST003"],
            "Net Amount": [120.25, 340.50, 200.00],
        })
        table, diag = build_portfolio_summary(workbook_bytes(source))
        self.assertEqual(diag["included_rows"], 3)
        self.assertEqual(diag["excluded_stages"], 0)
        self.assertEqual(table["المنتج"].tolist(), ["CC", "RF", "PF"])
        self.assertEqual(table["اسم المحصل"].tolist(),
                         ["Collector A", "Collector A", "Collector B"])
        self.assertEqual(table["DPD30 | عدد الحسابات"].sum(), 1)
        self.assertEqual(table["DPD60 | عدد الحسابات"].sum(), 1)
        self.assertEqual(table["NPL | عدد الحسابات"].sum(), 1)
        total_debt = sum(table[f"{stage} | إجمالي المديونية"].sum()
                         for stage in STAGES)
        self.assertAlmostEqual(total_debt, 660.75)

    def test_count_is_nonblank_rows_not_distinct_ids(self):
        repeated = self.sample.iloc[:2].copy()
        repeated["ملاحظات-التمويل"] = ["Dpd30", "Dpd30"]
        repeated["Account Number"] = ["001", "001"]
        table, _ = build_portfolio_summary(workbook_bytes(repeated))
        self.assertEqual(table.iloc[0]["DPD30 | عدد الحسابات"], 2)

    def test_missing_column_is_clear_error(self):
        sample = self.sample.drop(columns=["PRODUCT_CATEGORY"])
        with self.assertRaisesRegex(ValueError, "المنتج"):
            build_portfolio_summary(workbook_bytes(sample))

    def test_invalid_debt_never_silently_becomes_zero(self):
        sample = self.sample.copy()
        sample["Net Amount"] = sample["Net Amount"].astype(object)
        sample.loc[0, "Net Amount"] = "invalid"
        with self.assertRaisesRegex(ValueError, "غير رقمي"):
            build_portfolio_summary(workbook_bytes(sample))

    def test_blank_debt_is_disclosed(self):
        sample = self.sample.copy()
        sample.loc[0, "Net Amount"] = None
        table, info = build_portfolio_summary(workbook_bytes(sample))
        self.assertEqual(info["blank_debts"], 1)
        self.assertEqual(table.iloc[0]["DPD30 | إجمالي المديونية"], 0)

    def test_excel_has_exact_grouped_two_row_headers(self):
        table, _ = build_portfolio_summary(workbook_bytes(self.sample))
        export_bytes = portfolio_summary_to_excel(table)
        sheet = load_workbook(BytesIO(export_bytes)).active
        self.assertEqual(sheet.title, "المحفظة")
        self.assertEqual(sheet["A1"].value, "اسم المحصل")
        self.assertEqual(sheet["B1"].value, "المنتج")
        self.assertEqual([sheet.cell(1, c).value for c in (3, 5, 7)],
                         ["DPD30", "DPD60", "NPL"])
        self.assertEqual(sheet["C2"].value, "عدد الحسابات")
        self.assertEqual(sheet["D2"].value, "إجمالي المديونية")
        self.assertEqual(sheet["A3"].value, "Collector A")
        self.assertEqual(sheet["B3"].value, "CC")
        self.assertAlmostEqual(sheet["D3"].value, 2280401.12)
        self.assertIn("A1:A2", [str(r) for r in sheet.merged_cells.ranges])
        self.assertIn("C1:D1", [str(r) for r in sheet.merged_cells.ranges])
        self.assertEqual(sheet.max_column, 8)
        self.assertEqual(sheet.freeze_panes, "C3")

    def test_normalization_is_explicit(self):
        self.assertEqual(_normalize_stage("30Dpd"), "DPD30")
        self.assertEqual(_normalize_stage("DPD_60"), "DPD60")
        self.assertEqual(_normalize_stage("NPL"), "NPL")
        self.assertEqual(_normalize_stage("unknown"), "")


if __name__ == "__main__":
    unittest.main()
