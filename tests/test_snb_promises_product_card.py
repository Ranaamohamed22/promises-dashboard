"""Synthetic tests for the single SNB 'الوعود لكل منتج' UI addition."""
import ast
from datetime import date
from html import escape
import io
from pathlib import Path
import re
import unittest

import pandas as pd


FUNCTIONS = {
    "_clean_snb_promise_text",
    "_snb_promise_data",
    "_snb_promise_product_summary",
    "_snb_promise_product_cards_html",
    "_snb_promises_export",
}


def load_selected_functions():
    source = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    extracted = [
        node for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name in FUNCTIONS
    ]
    if {node.name for node in extracted} != FUNCTIONS:
        raise AssertionError("SNB product card/source functions not found")
    scope = {"pd": pd, "re": re, "io": io, "_xml_escape": escape}
    exec(compile(ast.fix_missing_locations(
        ast.Module(body=extracted, type_ignores=[])), "app.py", "exec"), scope)
    return scope, source


def row(product, *, due="2026-11-26", last="2026-11-25",
        status="غير معالج", category="Normal",
        classification="وعود السداد قيد التنفيذ"):
    return {
        "نوع المتنج-التمويل": product,
        "Classification": classification,
        "Sub State": "وعد بسداد قسط",
        "حالة المعالجة - التمويل": status,
        "تحديث فئة العميل - التمويل": category,
        "Follow up Due Date": due,
        "Follow up Last Date": last,
        "Salesperson": "Test collector",
    }


class TestSnbPromiseProducts(unittest.TestCase):
    def setUp(self):
        self.fn, self.source = load_selected_functions()
        self.today = date(2026, 11, 26)
        self.raw = pd.DataFrame([
            row("Alpha"),
            row("Alpha"),
            row("Alpha", due="2026-11-25", last="2026-11-24"),
            row("Beta"),
            row("Beta", due="2026-11-25", last="2026-11-24"),
            row("  Beta\u2060 "),
            row(None),
            row("Gamma", status="معالج"),
            row("Gamma", category="WriteOff"),
            row("Gamma", classification="not a promise"),
            row("Gamma", due="2026-11-27"),
            row("Gamma", due="2026-11-25", last="2026-11-25"),  # zero delta
            row("Gamma", due="not-a-date"),
        ])

    def calculate(self, raw=None):
        data = self.raw if raw is None else raw
        active, broken, types, *_ = self.fn["_snb_promise_data"](data, self.today)
        per = self.fn["_snb_promise_product_summary"](data, active, broken)
        return active, broken, types, per

    def test_counts_come_only_from_exact_existing_kpi_frames(self):
        active, broken, _, per = self.calculate()
        self.assertEqual(len(active), 5)
        self.assertEqual(len(broken), 2)
        self.assertEqual(int(per["قائمة"].sum()), len(active))
        self.assertEqual(int(per["مكسورة"].sum()), len(broken))
        self.assertEqual(int(per["الإجمالي"].sum()), len(active) + len(broken))
        self.assertEqual(per["المنتج"].tolist(), [
            "Alpha", "Beta", "غير محدد", "Gamma",
        ])
        self.assertEqual(per.loc[0, ["قائمة", "مكسورة", "الإجمالي"]].tolist(), [2, 1, 3])
        self.assertEqual(per.loc[1, ["قائمة", "مكسورة", "الإجمالي"]].tolist(), [2, 1, 3])
        self.assertEqual(per.loc[2, ["قائمة", "مكسورة", "الإجمالي"]].tolist(), [1, 0, 1])
        self.assertEqual(per.loc[3, ["قائمة", "مكسورة", "الإجمالي"]].tolist(), [0, 0, 0])
        self.assertAlmostEqual(per.loc[0, "نسبة الكسر (محسوبة)"], 1 / 3)
        self.assertTrue(pd.isna(per.loc[3, "نسبة الكسر (محسوبة)"]))

    def test_card_gauge_calculated_percent_color_and_zero_cases(self):
        _, _, _, per = self.calculate()
        html = self.fn["_snb_promise_product_cards_html"](per, True, True, True)
        self.assertIn("الوعود لكل منتج", html)
        self.assertIn("33.3%", html)
        self.assertIn("conic-gradient(from 270deg, var(--broken) 0 16.666667%", html)
        self.assertIn("-webkit-mask:radial-gradient(circle,transparent 80px,#000 81px)", html)
        self.assertIn("rgba(255,255,255,0.15)", html)
        self.assertIn("0.0%", html)  # product with zero broken but non-zero standing
        self.assertIn("—", html)  # product with total 0
        self.assertEqual(html.count('<article class="snb-pr-product-item">'), len(per))
        self.assertNotIn("color:red", html)
        self.assertIn("var(--accent)", html)
        self.assertIn("var(--broken)", html)
        self.assertIn("var(--text)", html)
        self.assertIn("var(--muted)", html)

    def test_no_upload_three_unlabeled_placeholders(self):
        html = self.fn["_snb_promise_product_cards_html"](None, False, False, False)
        self.assertEqual(html.count('<article class="snb-pr-product-item">'), 3)
        self.assertEqual(html.count('<div class="snb-pr-product-name">—</div>'), 3)
        self.assertIn("ارفعي الملف علشان تظهر الأرقام", html)
        self.assertNotIn("Alpha", html)

    def test_missing_product_column_is_only_a_card_message(self):
        without = self.raw.drop(columns=["نوع المتنج-التمويل"])
        active, broken, *_ = self.fn["_snb_promise_data"](without, self.today)
        per = self.fn["_snb_promise_product_summary"](without, active, broken)
        self.assertIsNone(per)
        html = self.fn["_snb_promise_product_cards_html"](per, True, True, False)
        self.assertIn("العمود نوع المتنج-التمويل غير موجود في الملف المرفوع", html)
        self.assertIn("الوعود لكل منتج", html)
        self.assertEqual(html.count('<article class="snb-pr-product-item">'), 0)

    def test_excel_adds_one_sheet_and_uses_same_product_table(self):
        active, broken, types, per = self.calculate()
        summary = [
            ("إجمالي الحسابات", len(self.raw)),
            ("الوعود القائمة", len(active)),
            ("الوعود المكسورة", len(broken)),
            ("نسبة الكسر (محسوبة)", len(broken) / (len(active) + len(broken))),
        ]
        payload = self.fn["_snb_promises_export"](
            summary, types, active, broken, [], per,
        )
        sheets = pd.read_excel(io.BytesIO(payload), sheet_name=None)
        self.assertEqual(list(sheets), [
            "ملخص", "الوعود القائمة", "الوعود المكسورة", "لكل محصل", "لكل منتج",
        ])
        exported = sheets["لكل منتج"]
        self.assertEqual(list(exported.columns), [
            "المنتج", "قائمة", "مكسورة", "الإجمالي", "نسبة الكسر (محسوبة)",
        ])
        pd.testing.assert_frame_equal(
            exported[["المنتج", "قائمة", "مكسورة", "الإجمالي"]],
            per[["المنتج", "قائمة", "مكسورة", "الإجمالي"]],
        )
        self.assertEqual(len(sheets["الوعود القائمة"]), len(active))
        self.assertEqual(len(sheets["الوعود المكسورة"]), len(broken))

    def test_empty_product_card_survives_all_three_streamlit_themes(self):
        from streamlit.testing.v1 import AppTest
        path = Path(__file__).resolve().parents[1] / "app.py"
        ui = AppTest.from_file(str(path), default_timeout=60).run()
        for theme in ("navy", "dark", "light"):
            ui.button(key=f"ejada_theme_{theme}").click().run()
            self.assertEqual(ui.session_state["theme"], theme)
            ui.button(key="home_report_open_Promises").click().run()
            self.assertEqual(len(ui.exception), 0)
            self.assertEqual(ui.session_state["current_page"], "Promises")
            product_html = [
                markdown.value for markdown in ui.markdown
                if "snb-pr-product-section" in markdown.value
                and "الوعود لكل منتج" in markdown.value
            ]
            self.assertEqual(len(product_html), 1)
            self.assertIn("ارفعي الملف علشان تظهر الأرقام", product_html[0])
            self.assertEqual(
                product_html[0].count('<article class="snb-pr-product-item">'), 3,
            )
            ui.button(key="snb_pr_change_bank").click().run()
            self.assertEqual(len(ui.exception), 0)

    def test_card_is_between_existing_sections_and_old_sections_unchanged(self):
        start = self.source.index('def _render_snb_promises(files):')
        end = self.source.index('if _snb_overview_page:', start)
        code = self.source[start:end]
        self.assertIn('with st.container(key="snb_glass_promises"):', code)
        self.assertIn("snb-pr-metrics", code)
        self.assertIn("أنواع الوعود", code)
        self.assertIn("+ product_section", code)
        self.assertIn("الوعود لكل محصل · أعلى 10 محصلين", code)
        self.assertLess(code.index("أنواع الوعود"), code.index("+ product_section"))
        self.assertLess(code.index("+ product_section"), code.index("الوعود لكل محصل · أعلى 10 محصلين"))
        self.assertIn('key="snb_pr_download"', code)


if __name__ == "__main__":
    unittest.main()
