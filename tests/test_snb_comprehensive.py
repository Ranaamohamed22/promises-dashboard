"""Regression tests for SNB comprehensive reporting (synthetic data only)."""
import unittest
from io import BytesIO
from openpyxl import Workbook, load_workbook
from snb_comprehensive import build_comprehensive_report

def make_workbook(headers, rows):
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()

class TestSnbComprehensive(unittest.TestCase):
    def setUp(self):
        self.headers = ["Account Number", "DPD_BUCKET", "BUCKET_07", "BUCKET_08", "BUCKET_09"]
        self.rows = [
            ["A", "05 - 60", "05 - 60", "04 - 30", "03 - 20"],
            ["B", "03 - 20", "03 - 20", "02 - 10", "04 - 40"],
            ["C", "02 - 10", "02 - 10", "03 - 20", "04 - 40"],
            ["D", "05 - 60", "05 - 60", "04 - 30", "29 - 999 Writeoff"],
        ]
        self.source = make_workbook(self.headers, self.rows)

    def test_categories_and_writeoff_are_preserved(self):
        output, count, bucket_names, totals = build_comprehensive_report(self.source)
        sheet = load_workbook(BytesIO(output)).active
        self.assertEqual(count, 4)
        self.assertEqual(bucket_names, ("DPD_BUCKET", "BUCKET_09"))
        self.assertEqual(totals, {"معالج": 2, "غير معالج": 2, "direct": 1, "middle": 1})
        self.assertEqual([sheet.cell(1, col).value for col in range(6, 12)],
                         ["BN", "B-1", "B-2", "B-3", "BN - B1", "حالة المعالجة"])
        self.assertEqual([sheet.cell(row, 11).value for row in range(2, 6)],
                         ["معالج", "معالج", "غير معالج", "غير معالج"])
        self.assertEqual(sheet["E5"].value, "Writeoff")
        self.assertEqual(sheet["J5"].value, "Writeoff")
        self.assertTrue(sheet["J2"].value.startswith("="))
        self.assertEqual(sheet.max_column, 11)

    def test_reupload_does_not_duplicate_columns(self):
        once, _, _, first_counts = build_comprehensive_report(self.source)
        twice, _, _, second_counts = build_comprehensive_report(once)
        sheet = load_workbook(BytesIO(twice)).active
        self.assertEqual(sheet.max_column, 11)
        self.assertEqual(first_counts, second_counts)
        self.assertEqual(sheet["J5"].value, "Writeoff")

    def test_missing_historical_buckets_is_rejected(self):
        source = make_workbook(["Account", "DPD_BUCKET", "BUCKET_01"],
                               [["A", "03 - 20", "02 - 10"]])
        with self.assertRaisesRegex(ValueError, "BUCKET"):
            build_comprehensive_report(source)

    def test_invalid_first_bucket_cannot_be_treated(self):
        source = make_workbook(self.headers,
                               [["X", None, "04 - 30", "03 - 20", "02 - 10"]])
        _, count, _, totals = build_comprehensive_report(source)
        self.assertEqual(count, 1)
        self.assertEqual(totals["معالج"], 0)
        self.assertEqual(totals["غير معالج"], 1)

if __name__ == "__main__":
    unittest.main()
