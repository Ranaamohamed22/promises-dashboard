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
        self.assertEqual(ui.session_state["theme"], "navy")
        ui.button(key="ejada_theme_dark").click().run()
        self.assertEqual(len(ui.exception), 0)
        self.assertEqual(ui.session_state["theme"], "dark")

        ui.button(key="home_report_open_Overview").click().run()
        self.assertEqual(len(ui.exception), 0)
        self.assertEqual(ui.session_state["selected_bank"], "SNB")
        self.assertEqual(ui.session_state["current_page"], "Overview")

        # SNB performance must still render the exact original upload widgets.
        ui.button(key="nav_perf_btn").click().run()
        self.assertEqual(len(ui.exception), 0)
        self.assertEqual(ui.session_state["current_page"], "Performance")

        self.assertEqual(ui.session_state["theme"], "dark")
        self.assertEqual(ui.session_state["selected_bank"], "SNB")

        # Keep native upload controls and tabs, without the custom reset button.
        # Streamlit's AppTest does not expose file_uploader directly.
        self.assertNotIn(
            "snb_perf_reset_uploads",
            [widget.key for widget in ui.button],
        )
        self.assertEqual(
            [tab.label for tab in ui.tabs],
            ["المعالج", "المحفظة"],
        )

        # Change bank through the original home flow; STC remains independent.
        ui.button(key="ejada_change_bank_btn").click().run()
        self.assertEqual(len(ui.exception), 0)
        ui.button(key="home_switch_STC").click().run()
        ui.button(key="home_report_open_Overview").click().run()
        self.assertEqual(len(ui.exception), 0)
        self.assertEqual(ui.session_state["selected_bank"], "STC")
        self.assertEqual(ui.session_state["selected_bank"], "STC")
        self.assertEqual(ui.session_state["theme"], "dark")

if __name__ == "__main__":
    unittest.main()
