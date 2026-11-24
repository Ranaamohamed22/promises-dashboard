"""Smoke test of the production Streamlit home and theme selector.

Uses Streamlit's in-memory test runner without real client files.
"""
from pathlib import Path
import unittest

class TestProductionUi(unittest.TestCase):
    def test_home_bank_selection_and_theme_controls(self):
        from streamlit.testing.v1 import AppTest
        app_path = Path(__file__).resolve().parents[1] / "app.py"
        ui = AppTest.from_file(str(app_path), default_timeout=60).run()
        self.assertEqual(len(ui.exception), 0)
        self.assertEqual(ui.session_state["theme_choice"], "حسب البنك (تلقائي)")

        ui.button(key="home_open_bank_SNB").click().run()
        self.assertEqual(len(ui.exception), 0)
        self.assertEqual(ui.session_state["selected_bank"], "SNB")
        self.assertEqual(ui.session_state["current_page"], "Overview")

        ui.selectbox(key="theme_choice").set_value("أسود هادي (Charcoal)").run()
        self.assertEqual(len(ui.exception), 0)
        self.assertEqual(ui.session_state["theme_choice"], "أسود هادي (Charcoal)")
        self.assertEqual(ui.session_state["selected_bank"], "SNB")

        # SNB performance must still render the exact original upload widgets.
        ui.button(key="nav_perf_btn").click().run()
        self.assertEqual(len(ui.exception), 0)
        self.assertEqual(ui.session_state["current_page"], "Performance")
        # File uploaders are not exposed by Streamlit AppTest. Verify that
        # the Performance page's own reset control rendered successfully.
        self.assertIsNotNone(ui.button(key="snb_perf_reset_uploads"))

        # Change bank through the original home flow; STC remains independent.
        ui.button(key="ejada_change_bank_btn").click().run()
        self.assertEqual(len(ui.exception), 0)
        ui.button(key="home_open_bank_STC").click().run()
        self.assertEqual(len(ui.exception), 0)
        self.assertEqual(ui.session_state["selected_bank"], "STC")
        ui.selectbox(key="theme_choice").set_value("حسب البنك (تلقائي)").run()
        self.assertEqual(len(ui.exception), 0)
        self.assertEqual(ui.session_state["theme_choice"], "حسب البنك (تلقائي)")
        self.assertEqual(ui.session_state["selected_bank"], "STC")

if __name__ == "__main__":
    unittest.main()
