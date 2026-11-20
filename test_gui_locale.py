"""Headless checks that the GUI honors the selected locale (phone country code)."""
import sys, tempfile, unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import openpyxl
try:
    import clean_sheet_gui as g
except ModuleNotFoundError:
    g = None


@unittest.skipIf(g is None, "tkinter/clean_sheet_gui not importable")
class GuiLocale(unittest.TestCase):
    def make_phone_file(self, d, number):
        p = Path(d) / "phones.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
        ws.append(["Client", "Phone"]); ws.append(["Ali", number]); wb.save(p); return p

    def test_locale_fr_uses_33(self):
        with tempfile.TemporaryDirectory() as d:
            p = self.make_phone_file(d, "0612345678")
            r = g.run_job(p, {"Client": "name", "Phone": "phone"}, "dmy", Path(d) / "o", locale="fr")
            val = openpyxl.load_workbook(r["xlsx"])["Clean"]["B2"].value
            self.assertEqual(val, "+33612345678")

    def test_locale_dz_uses_213(self):
        with tempfile.TemporaryDirectory() as d:
            p = self.make_phone_file(d, "0612345678")
            r = g.run_job(p, {"Client": "name", "Phone": "phone"}, "dmy", Path(d) / "o", locale="dz")
            val = openpyxl.load_workbook(r["xlsx"])["Clean"]["B2"].value
            self.assertEqual(val, "+213612345678")

    def test_locale_ma_uses_212(self):
        with tempfile.TemporaryDirectory() as d:
            p = self.make_phone_file(d, "0612345678")
            r = g.run_job(p, {"Client": "name", "Phone": "phone"}, "dmy", Path(d) / "o", locale="ma")
            val = openpyxl.load_workbook(r["xlsx"])["Clean"]["B2"].value
            self.assertEqual(val, "+212612345678")

    def test_default_locale_is_dz(self):
        self.assertEqual(g.LOCALE_DEFAULT, "dz")


if __name__ == "__main__": unittest.main(verbosity=1)
