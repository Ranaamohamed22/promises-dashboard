"""End-to-end synthetic regression for the existing daily Excel report script."""
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path
from unittest.mock import patch
import pandas as pd
import run

class TestDailyPromisesReport(unittest.TestCase):
    def test_generates_same_active_and_broken_sheets(self):
        today = date.today()
        yesterday = today - timedelta(days=1)
        before_yesterday = today - timedelta(days=2)
        base = {
            "حالة المعالجة - التمويل": "غير معالج",
            "تحديث فئة العميل - التمويل": "normal",
            "Classification": "وعود السداد قيد التنفيذ",
        }
        rows = [
            dict(base, **{"Account Number": "active", "Follow up Due Date": today,
                          "Follow up Last Date": yesterday}),
            dict(base, **{"Account Number": "broken", "Follow up Due Date": yesterday,
                          "Follow up Last Date": before_yesterday}),
            dict(base, **{"Account Number": "not_broken", "Follow up Due Date": yesterday,
                          "Follow up Last Date": today}),
            dict(base, **{"Account Number": "writeoff", "Follow up Due Date": today,
                          "Follow up Last Date": yesterday,
                          "تحديث فئة العميل - التمويل": "Writeoff"}),
            dict(base, **{"Account Number": "processed", "Follow up Due Date": today,
                          "Follow up Last Date": yesterday,
                          "حالة المعالجة - التمويل": "معالج"}),
        ]
        with tempfile.TemporaryDirectory() as root:
            source_dir, output_dir = Path(root) / "input", Path(root) / "output"
            source_dir.mkdir()
            pd.DataFrame(rows).to_excel(source_dir / "daily.xlsx", index=False)
            with patch.object(run, "input_folder", str(source_dir)), patch.object(
                run, "output_folder", str(output_dir)
            ):
                run.process_daily_sheets()
            filename = output_dir / f"Daily_Promises_Report_{today}.xlsx"
            self.assertTrue(filename.is_file())
            active = pd.read_excel(filename, sheet_name="الوعود القائمة", dtype={"Account Number": str})
            broken = pd.read_excel(filename, sheet_name="الوعود المكسورة", dtype={"Account Number": str})
            self.assertEqual(active["Account Number"].tolist(), ["active"])
            self.assertEqual(broken["Account Number"].tolist(), ["broken"])
            self.assertEqual(broken["Date_Diff"].tolist(), [-1])

if __name__ == "__main__":
    unittest.main()
