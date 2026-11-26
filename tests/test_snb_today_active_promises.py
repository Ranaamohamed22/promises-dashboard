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
            # This WriteOff is excluded despite due today, unprocessed and last != today.
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
        self.frame["Classification"] = "وعود السداد قيد التنفيذ"
        self.frame.loc[self.frame["account"].eq("F"), "Classification"] = "محاولة التوصل للعميل"
        self.frame = pd.concat([self.frame, pd.DataFrame([
            # Sub State is not a 'وعد بسداد' string, but classification says promise.
            {"account": "J", "Sub State": "يرغب بجدولة المديونية",
             "Follow up Due Date": "2026-11-26", "Follow up Last Date": "2026-11-25",
             "حالة المعالجة - التمويل": "غير معالج",
             "تحديث فئة العميل - التمويل": "Normal",
             "Classification": "وعود السداد قيد التنفيذ"},
            # Sub State says 'وعد بسداد', but classification does NOT: exclude.
            {"account": "K", "Sub State": "وعد بسداد قسط",
             "Follow up Due Date": "2026-11-26", "Follow up Last Date": "2026-11-25",
             "حالة المعالجة - التمويل": "غير معالج",
             "تحديث فئة العميل - التمويل": "Normal",
             "Classification": "محاولة التوصل للعميل"},
            # A classified promise from a non-'وعد بسداد' Sub State is broken.
            {"account": "L", "Sub State": "يرغب بجدولة المديونية",
             "Follow up Due Date": "2026-11-25", "Follow up Last Date": "2026-11-24",
             "حالة المعالجة - التمويل": "غير معالج",
             "تحديث فئة العميل - التمويل": "Normal",
             "Classification": "وعود السداد قيد التنفيذ"},
        ])], ignore_index=True)

    def test_exact_active_condition(self):
        active, broken, types, promise_count, missing_dates = (
            self.fn["_snb_promise_data"](self.frame, self.today))
        self.assertEqual(active["account"].tolist(), ["E", "J"])
        self.assertEqual(broken["account"].tolist(), ["G", "H", "L"])
        self.assertEqual(promise_count, 10)
        self.assertEqual(missing_dates, 1)
        self.assertEqual(int(types.sum()), 10)

    def test_processed_and_future_never_count_as_active(self):
        active, *_ = self.fn["_snb_promise_data"](self.frame, self.today)
        self.assertFalse({"B", "C", "D", "K"} & set(active["account"].tolist()))

    def test_writeoff_is_excluded_from_active_only(self):
        active, broken, *_ = self.fn["_snb_promise_data"](self.frame, self.today)
        self.assertNotIn("A", active["account"].tolist())
        self.assertEqual(broken["account"].tolist(), ["G", "H", "L"])

    def test_writeoff_spelling_and_whitespace_are_normalized(self):
        extra = self.frame.loc[self.frame["account"].eq("A")].copy()
        extra["account"] = "J"
        extra["تحديث فئة العميل - التمويل"] = "  WRITEoff  "
        data = pd.concat([self.frame, extra], ignore_index=True)
        active, *_ = self.fn["_snb_promise_data"](data, self.today)
        self.assertEqual(active["account"].tolist(), ["E", "J"])

    def test_classification_is_authoritative_over_sub_state(self):
        active, broken, types, *_ = self.fn["_snb_promise_data"](self.frame, self.today)
        self.assertIn("J", active["account"].tolist())
        self.assertIn("L", broken["account"].tolist())
        self.assertNotIn("K", active["account"].tolist())
        self.assertEqual(types.loc["يرغب بجدولة المديونية"], 2)

    def test_hidden_classification_chars_are_normalized(self):
        data = self.frame.copy()
        data.loc[data["account"].eq("J"), "Classification"] = "\u2060 وعود  السداد قيد التنفيذ "
        active, *_ = self.fn["_snb_promise_data"](data, self.today)
        self.assertIn("J", active["account"].tolist())

    def test_active_filter_order_is_status_writeoff_classification_dates(self):
        fn_start = self.source.index("def _snb_promise_data(raw, today):")
        fn_end = self.source.index("def _render_snb_promises(files):", fn_start)
        code = self.source[fn_start:fn_end]
        steps = (
            "after_status = unprocessed",
            "after_writeoff = after_status & not_writeoff",
            "after_classification = after_writeoff & is_promise",
            "after_due = after_classification & valid & dates.eq(today)",
            "current = after_due & last_dates.ne(today)",
        )
        indices = [code.index(step) for step in steps]
        self.assertEqual(indices, sorted(indices))
        # The separate broken-promise path must retain its existing rule.
        self.assertIn("broken = is_promise & valid & dates.lt(today)", code)

    def test_uploader_requires_last_date_and_processing_status(self):
        self.assertIn(
            '"Classification", "Sub State", "Follow up Due Date", "Follow up Last Date",',
            self.source,
        )
        self.assertIn(
            '"حالة المعالجة - التمويل", "تحديث فئة العميل - التمويل"\n    }.issubset(raw.columns)',
            self.source,
        )


if __name__ == "__main__":
    unittest.main()
