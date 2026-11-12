"""Headless checks for clean_sheet_gui (no display needed for the logic; GUI smoke test needs a display)."""
import sys, tempfile, unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
import openpyxl
try:
    import clean_sheet_gui as g
except ModuleNotFoundError as e:
    g = None; print("tkinter missing:", e)

@unittest.skipIf(g is None, "tkinter not installed")
class GuiLogic(unittest.TestCase):
    def make(self, d):
        p = Path(d) / "in.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
        ws.append(["Client", "Prix", "Date", "Téléphone"]); ws.append([" ali ", "1 450 DA", "2026-03-04", "0555123456"]); ws.append([" ali ", "1 450 DA", "2026-03-04", "0555123456"]); ws.append(["sara", "900", "03/04/2026", "0661001122"])
        wb.save(p); return p
    def test_detect_and_run(self):
        with tempfile.TemporaryDirectory() as d:
            p = self.make(d); res = g.inspect_file(p); cols = res["types"]
            self.assertEqual(cols, {"Client": "name", "Prix": "price", "Date": "date", "Téléphone": "phone"}); self.assertEqual(res["header_row"], 1)
            r = g.run_job(p, cols, "mdy", Path(d) / "o")
            self.assertEqual((r["rows_in"], r["rows_out"], r["removed"], r["verify"]), (3, 2, 1, 1))
            self.assertTrue(r["xlsx"].exists() and (Path(d) / "o" / "in_used_config.json").exists())
            ws = openpyxl.load_workbook(r["xlsx"])["Clean"]; self.assertEqual(ws["C3"].value.date().isoformat(), "2026-03-04")  # ambiguous 03/04 read month-first as chosen
    def test_skip_type_leaves_column(self):
        with tempfile.TemporaryDirectory() as d:
            p = self.make(d); r = g.run_job(p, {"Client": "skip", "Prix": "price", "Date": "date", "Téléphone": "phone"}, "dmy", Path(d) / "o")
            self.assertEqual(openpyxl.load_workbook(r["xlsx"])["Clean"]["A2"].value, " ali ")
    def test_offset_arabic_table_with_title(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "m.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
            ws["C1"] = "كشف النقاط"; ws.merge_cells("C1:F1")
            for j, h in enumerate(["الاسم", "الهاتف", "النقطة", "المبلغ"]): ws.cell(row=4, column=3 + j, value=h)
            for j, v in enumerate(["ماريا", "٠٥٥٥١٢٣٤٥٦", "١٤٫٥", "١٢٠٠ دج"]): ws.cell(row=5, column=3 + j, value=v)
            for j, v in enumerate(["أحمد", "0661001122", "۱۲", "٩٠٠ دج"]): ws.cell(row=6, column=3 + j, value=v)
            wb.save(p); res = g.inspect_file(p)
            self.assertEqual((res["header_row"], res["how"]), (4, "auto"))
            self.assertEqual(res["types"], {"الاسم": "name", "الهاتف": "phone", "النقطة": "number", "المبلغ": "price"})
            self.assertTrue(res["samples"]["النقطة"].startswith("14.5"))
            r = g.run_job(p, res["types"], "dmy", Path(d) / "o")
            rows = list(openpyxl.load_workbook(r["xlsx"])["Clean"].values)
            self.assertEqual(rows[1][:4], ("ماريا", "+213555123456", 14.5, 1200.0)); self.assertEqual(rows[2][2], 12)
    def test_header_override(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "h.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
            ws.append(["Titre", None]); ws.append(["A", "B"]); ws.append([1, 2]); wb.save(p)
            self.assertEqual(g.inspect_file(p, None, 2)["headers"], ["A", "B"])
    def test_screenshot_layout_two_level_header_and_total_row(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "s.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
            ws["A1"] = "Mois"; ws["C1"] = "Produit1"; ws["F1"] = "Produits2"; ws["H1"] = "Total Génération"
            for k, h in enumerate([None, "Qté", "prix unit", "mongtant", "Qter", "prix unit", "montant", "Qté", "mis"]):
                if h: ws.cell(row=2, column=k + 1, value=h)
            for i, m in enumerate(["janvier", "février", "mars"], start=3):
                for k, v in enumerate([m, 1000 + i, 15.5, 15500, 1200, 15.0, 18000, 1200, 1000]): ws.cell(row=i, column=k + 1, value=v)
            ws["A6"] = "Totale"; ws["B6"] = 3000; wb.save(p)
            res = g.inspect_file(p)
            self.assertEqual((res["levels"], res["skipped"], res["rows"]), (2, 1, 3))
            self.assertEqual(list(res["headers"])[:4], ["Mois", "Produit1 - Qté", "Produit1 - prix unit", "Produit1 - mongtant"])
            self.assertEqual([res["types"][h] for h in res["headers"]], ["category", "qty", "price", "price", "qty", "price", "price", "qty", "number"])
            r = g.run_job(p, res["types"], "dmy", Path(d) / "o"); self.assertEqual((r["rows_out"], r["skipped"]), (3, 1))
            self.assertEqual(g.inspect_file(p, None, None, 1)["levels"], 1)
if __name__ == "__main__": unittest.main()
