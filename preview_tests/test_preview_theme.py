"""Checks that the Streamlit design preview themes switch without report changes."""
import unittest
from pathlib import Path

from preview_theme import BANK_PALETTES, THEME_OPTIONS, resolve_theme


class TestPreviewBrandThemes(unittest.TestCase):
    def test_bank_palettes_have_distinct_brand_colors(self):
        snb = resolve_theme("SNB")
        stc = resolve_theme("STC")
        self.assertNotEqual(snb["accent"], stc["accent"])
        self.assertNotEqual(snb["background"], stc["background"])
        self.assertEqual(set(snb), set(stc))
        self.assertTrue(snb["text"].startswith("#"))
        self.assertTrue(stc["text"].startswith("#"))

    def test_soft_charcoal_is_readable_and_keeps_bank_accent(self):
        for bank in BANK_PALETTES:
            with self.subTest(bank=bank):
                normal = resolve_theme(bank)
                charcoal = resolve_theme(bank, "أسود هادي")
                self.assertEqual(normal["accent"], charcoal["accent"])
                self.assertNotEqual(normal["background"], charcoal["background"])
                self.assertNotEqual(normal["text"], charcoal["text"])
                self.assertEqual(charcoal["text"], "#F4F8FC")
                self.assertNotEqual(charcoal["surface"], charcoal["background"])
        self.assertIn("أسود هادي", THEME_OPTIONS)

    def test_palettes_are_independent_copies(self):
        changed = resolve_theme("SNB")
        changed["accent"] = "#000000"
        self.assertNotEqual(resolve_theme("SNB")["accent"], "#000000")

    def test_unknown_bank_or_theme_rejected(self):
        with self.assertRaises(ValueError):
            resolve_theme("OTHER")
        with self.assertRaises(ValueError):
            resolve_theme("SNB", "not a theme")

    def test_streamlit_bank_and_theme_controls_switch_without_errors(self):
        from streamlit.testing.v1 import AppTest
        path = Path(__file__).resolve().parents[1] / "performance_preview.py"
        at = AppTest.from_file(str(path), default_timeout=40).run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(at.session_state["ejada_preview_bank"], "SNB")
        at.selectbox(key="ejada_preview_bank").set_value("STC").run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(at.session_state["ejada_preview_bank"], "STC")
        at.selectbox(key="ejada_preview_theme_mode").set_value("أسود هادي").run()
        self.assertEqual(len(at.exception), 0)
        self.assertEqual(at.session_state["ejada_preview_theme_mode"], "أسود هادي")

if __name__ == "__main__":
    unittest.main()
