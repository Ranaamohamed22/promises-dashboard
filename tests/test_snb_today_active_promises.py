"""SNB active-promises rules: synthetic cases, never real bank records.

Extract only the pure calculation functions from app.py so Streamlit does
not start or need banking secrets during regression testing.
"""
import ast
from datetime import date
from pathlib import Path
import re
import unittest

import pandas as pd


def load_promise_calculation():
    source = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    names = {"_clean_snb_promise_text", "_snb_promise_data"}
    nodes = [node for node in tree.body
             if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
             and node.name in names]
    if {node.name for node in nodes} != names:
        raise AssertionError("SNB promise calculation functions not found")
    scope = {"pd": pd, "re": re}
    mod = ast.Module(body=nodes, type_ignores=[])
    exec(compile(ast.fix_missing_locations(mod), "app.py", "exec"), scope)
    return scope, source


class TestSnbTodayActivePromises(unittest.TestCase):
    def setUp(self):
        self.fn, self.source = load_promise_calculation()
        self.today = date(2026, 11, 26)
        self.frame = pd.DataFrame([
            # Current active: due today, unprocessed and last != today.
            {"account": "A", "sub": "وعد بسداد مبلغ المعالجة", "due": "2026-11-26",
             "last": "2026-11-25", "state": "غير معالج", "category": "WriteOff"},
            # Today's last follow-up means EXCLUDE.
            {"account": "B", "sub": "وعد بسداد قسط", "due": "2026-11-26",
             "last": "2026-11-26", "state": "غير معالج", "category": "Normal"},
            # Already processed means EXCLUDE.
            {"account": "C", "sub": "وعد بسداد قسط", "due": "2026-11-26",
             "last": "2026-11-25", "state": "معالج", "category": "Normal"},
            # Future promise must not be counted under today's active count.
            {"account": "D", "sub": "وعد بسداد قسط", "due": "2026-11-27",
             "last": "2026-11-25", "state": "غير معالج", "category": "Normal"},
            # Excluding today's last follow-up does not exclude blank last dates.
            {"account": "E", "sub": "وعد بسداد قسط", "due": "2026-11-26",
             "last": None, "state": "غير معالج", "category": "Normal"},
            # Not a promise type.
            {"account": "F", "sub": "لا يجيب", "due": "2026-11-26",
             "last": "2026-11-25", "state": "غير معالج", "category": "Normal"},
            # Broken rule stays independent/unmodified.
            {"account": "G", "sub": "وعد بسداد قسط", "due": "2026-11-25",
             "last": "2026-11-24", "state": "غير معالج", "category": "Normal"},
            {"account": "H", "sub": "وعد بسداد قسط", "due": "2026-11-25",
             "last": "2026-11-24", "state": "معالج", "category": "Normal"},
            {"account": "I", "sub": "وعد بسداد قسط", "due": "bad_date",
             "last": "2026-11-25", "state": "غير معالج", "category": "Normal"},
        ]).rename(columns={
            "sub": "Sub State",
            "due": "Follow up Due Date",
            "last": "Follow up Last Date",
            "state": "حالة المعالجة - التمويل",
            "category": "تحديث فئة العميل - التمويل",
        })

    def test_exact_active_condition(self):
        active, broken, types, promise_count, missing_dates = (
            self.fn["_snb_promise_data"](self.frame, self.today))
        self.assertEqual(active["account"].tolist(), ["A", "E"])
        self.assertEqual(broken["account"].tolist(), ["G", "H"])
        self.assertEqual(promise_count, 8)
        self.assertEqual(missing_dates, 1)
        self.assertEqual(int(types.sum()), 8)

    def test_processed_and_future_never_count_as_active(self):
        active, *_ = self.fn["_snb_promise_data"](self.frame, self.today)
        self.assertFalse({"B", "C", "D"} & set(active["account"].tolist()))

    def test_writeoff_not_silently_excluded_without_requested_rule(self):
        active, *_ = self.fn["_snb_promise_data"](self.frame, self.today)
        self.assertIn("A", active["account"].tolist())

    def test_uploader_requires_last_date_and_processing_status(self):
        self.assertIn(
            '"Sub State", "Follow up Due Date", "Follow up Last Date",',
            self.source,
        )
        self.assertIn(
            '"حالة المعالجة - التمويل"\n    }.issubset(raw.columns)',
            self.source,
        )


if __name__ == "__main__":
    unittest.main()
