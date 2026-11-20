"""Headless check that CleanSheet_gui cleans on a background thread (non-blocking) and marshals results back."""
import sys, tempfile, unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import openpyxl
try:
    import tkinter as tk
    import clean_sheet_gui as g
    TK = True
except Exception:
    TK = False


@unittest.skipIf(not TK, "tkinter/clean_sheet_gui not importable headless")
class GuiThread(unittest.TestCase):
    def make_file(self, d):
        p = Path(d) / "in.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
        ws.append(["Client", "Phone"]); ws.append(["Ali", "0555123456"]); wb.save(p); return p

    def setUp(self):
        self.app = g.App()
        self.app.withdraw()
        self.calls = []

    def tearDown(self):
        try: self.app.destroy()
        except Exception: pass

    def test_clean_is_non_blocking_and_delivers_result(self):
        with tempfile.TemporaryDirectory() as d:
            p = self.make_file(d)
            self.app.path = p
            self.app.out_var.set(str(Path(d) / "o"))
            self.app.sheet = None; self.app.header_override = None; self.app.levels_override = None
            self.app.type_vars = {"Client": tk.StringVar(value="Name"), "Phone": tk.StringVar(value="Phone")}
            self.app.date_var = tk.StringVar(value=list(g.DATE_CHOICES)[0])
            self.app.csv_var = tk.BooleanVar(value=False)
            self.app.skip_var = tk.BooleanVar(value=True)
            self.app.locale_var = tk.StringVar(value=g.LOCALE_DEFAULT)

            def fake_run_job(path, columns, date_order="dmy", out_dir=None, also_csv=False, sheet=None,
                             header_row=None, levels=None, skip_totals=True, locale=g.LOCALE_DEFAULT):
                self.calls.append((columns, locale))
                import time; time.sleep(0.05)
                from collections import Counter
                return {"xlsx": Path(d) / "out" / "in_cleaned.xlsx", "report": Path(d) / "out" / "in_report.md",
                        "out": Path(d) / "out", "rows_in": 1, "rows_out": 1, "removed": 0,
                        "verify": 0, "unresolved": 0, "missing": 0, "note": "", "skipped": 0, "warnings": []}

            g.run_job = fake_run_job
            started = []
            self.app.clean()
            # must return immediately (non-blocking) and have disabled the button / started progress
            self.assertTrue(self.app.run_btn.instate(["disabled"]), "Clean button should be disabled while running")
            self.assertEqual(self.app._polling, True)
            # deliver the queued result by polling on the Tk thread (give the worker time to finish)
            import time
            for _ in range(200):
                if not self.app._polling:
                    break
                self.app._poll()
                time.sleep(0.01)
            self.assertFalse(self.app._polling)
            self.assertIsNotNone(self.app.result)
            self.assertEqual(self.app.result["rows_in"], 1)
            self.assertTrue(self.app.run_btn.instate(["!disabled"]))
            self.assertEqual(self.calls[0][1], g.LOCALE_DEFAULT)


if __name__ == "__main__": unittest.main(verbosity=1)
