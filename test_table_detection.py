import sys, tempfile, unittest
from datetime import datetime, date
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import openpyxl
from openpyxl.worksheet.table import Table
import clean_sheet as cs

AR_HEAD = ["الاسم", "الكمية", "السعر", "الهاتف", "التاريخ", "ملاحظة"]
def fill(ws, top, left, header, rows):
    for j, h in enumerate(header): ws.cell(row=top, column=left + j, value=h)
    for i, r in enumerate(rows, start=1):
        for j, v in enumerate(r): ws.cell(row=top + i, column=left + j, value=v)

class Offset(unittest.TestCase):
    def test_title_rows_and_offset_table(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "a.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
            ws["C1"] = "Rapport des ventes 2026"; ws.merge_cells("C1:H1")
            ws["C2"] = "Généré le"; ws["D2"] = "01/10/2026"
            fill(ws, 4, 3, AR_HEAD, [["أحمد", "٣", "١٢٠٠ دج", "٠٥٥٥١٢٣٤٥٦", "٢٠٢٦/٠٣/٠١", "طلب ٢ قطع"],
                                      ["سارة", "۲", "١٬٥٠٠ د.ج", "0661001122", "٢٠٢٦/٠٣/٠٢", "عادي"]])
            wb.save(p); info = cs.read_table_info(p)
            self.assertEqual(info["header_row"], 4); self.assertEqual(info["headers"], AR_HEAD); self.assertEqual(info["rows"][0]["src"], 5)
            r = cs.run(p, out_dir=Path(d) / "o")
            self.assertEqual(r["cfg"]["columns"], {"الاسم": "name", "الكمية": "qty", "السعر": "price", "الهاتف": "phone", "التاريخ": "date", "ملاحظة": "text"})
            rows = list(openpyxl.load_workbook(r["xlsx"])["Clean"].values)
            self.assertEqual(rows[1][:2], ("أحمد", 3)); self.assertEqual(rows[1][3], "+213555123456")
            self.assertEqual(rows[1][2], 1200.0); self.assertEqual(rows[2][2], 1500.0)   # Arabic currency + digits
            self.assertEqual(rows[1][5], "طلب 2 قطع")                                      # digits inside text converted
            self.assertEqual(rows[2][1], 2)                                                # Persian digit
            self.assertIn("header at row 4", openpyxl.load_workbook(r["xlsx"])["Summary"]["B3"].value)
    def test_first_row_empty_header_on_row_3(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "b.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
            fill(ws, 3, 2, ["Nom", "Note", "Note finale"], [["Ali", 12, 14], ["Sara", 15, 16]]); wb.save(p)
            info = cs.read_table_info(p); self.assertEqual((info["header_row"], info["headers"]), (3, ["Nom", "Note", "Note finale"]))
    def test_excel_table_object_ignores_notes_below(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "c.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
            ws["A1"] = "note"; fill(ws, 5, 2, ["Client", "Prix"], [["a", 10], ["b", 20]])
            t = Table(displayName="T1", ref="B5:C7"); ws.add_table(t); ws["B10"] = "Total"; ws["C10"] = 30; wb.save(p)
            info = cs.read_table_info(p); self.assertEqual((info["header_row"], len(info["rows"]), info["how"]), (5, 2, "Excel table"))
    def test_empty_header_cell_gets_column_letter_not_number(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "e.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
            fill(ws, 1, 1, ["Client", None, "Prix"], [["a", "x", 10], ["b", "y", 20]]); wb.save(p)
            self.assertEqual(cs.read_table_info(p)["headers"], ["Client", "Column B", "Prix"])
    def test_manual_header_row_and_sheet(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "m.xlsx"; wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Vide"; ws2 = wb.create_sheet("Data")
            fill(ws2, 2, 1, ["A", "B"], [[1, 2], [3, 4]]); wb.save(p)
            self.assertEqual(cs.list_sheets(p), ["Vide", "Data"])
            info = cs.read_table_info(p, "Data", 2); self.assertEqual((info["headers"], info["how"], info["sheet"]), (["A", "B"], "manual", "Data"))
            with self.assertRaises(SystemExit): cs.read_table_info(p, "Nope")
    def test_normal_file_still_row_1(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "n.xlsx"; wb = openpyxl.Workbook(); ws = wb.active; fill(ws, 1, 1, ["Customer", "Qty", "Price"], [["a", 1, 5], ["b", 2, 6]]); wb.save(p)
            self.assertEqual(cs.read_table_info(p)["header_row"], 1)

class Digits(unittest.TestCase):
    def test_text_and_id_digits_converted_and_logged(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "x.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
            fill(ws, 1, 1, ["رقم", "ملاحظة"], [["١٢٣", "قطعة ٥"], ["۴۵۶", "عادي"]]); wb.save(p)
            r = cs.run(p, out_dir=Path(d) / "o"); wb2 = openpyxl.load_workbook(r["xlsx"])
            rows = list(wb2["Clean"].values); self.assertEqual((rows[1][0], rows[1][1], rows[2][0]), ("123", "قطعة 5", "456"))
            log = list(wb2["Change log"].values)[1:]; self.assertTrue(any(l[2] == "١٢٣" and l[3] == "123" for l in log))
    def test_skip_column_left_untouched(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "s.xlsx"; wb = openpyxl.Workbook(); ws = wb.active; fill(ws, 1, 1, ["A"], [["٣"], ["٤"]]); wb.save(p)
            import json; c = Path(d) / "c.json"; c.write_text(json.dumps({"columns": {"A": "skip"}}))
            r = cs.run(p, c, Path(d) / "o"); self.assertEqual(list(openpyxl.load_workbook(r["xlsx"])["Clean"].values)[1][0], "٣")

class Sniff(unittest.TestCase):
    def test_unknown_headers_typed_by_content(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "k.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
            fill(ws, 1, 1, ["A", "B", "C", "D"], [["0555 12 34 56", "2026-03-01", "a@b.com", "xx"], ["0661 00 11 22", "2026-03-02", "c@d.org", "yy"], ["0770 11 22 33", "2026-03-03", "e@f.net", "zz"]]); wb.save(p)
            h, rows, _ = cs.read_table(p); cfg = cs.resolve_config({}, h, rows)
            self.assertEqual(cfg["columns"], {"A": "phone", "B": "date", "C": "email", "D": "text"})

if __name__ == "__main__": unittest.main()

class Numbers(unittest.TestCase):
    def test_number_type_and_logging(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "n.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
            fill(ws, 1, 1, ["الاسم", "النقطة", "الكمية"], [["ماريا", "١٤٫٥", "٣"], ["أحمد", "۱۲", "٢"], ["سارة", "16", "1"]]); wb.save(p)
            r = cs.run(p, out_dir=Path(d) / "o"); self.assertEqual(r["cfg"]["columns"]["النقطة"], "number")
            wb2 = openpyxl.load_workbook(r["xlsx"]); rows = list(wb2["Clean"].values)
            self.assertEqual([x[1] for x in rows[1:]], [14.5, 12, 16]); self.assertEqual([x[2] for x in rows[1:]], [3, 2, 1])
            log = [(l[1], l[2], l[3]) for l in list(wb2["Change log"].values)[1:]]
            self.assertIn(("الكمية", "٣", "3"), log); self.assertIn(("النقطة", "١٤٫٥", "14.5"), log)   # Arabic digits are logged as changes
    def test_sniff_number_but_not_ids_with_leading_zero(self):
        cfg = cs.resolve_config({}, ["A"])
        self.assertEqual(cs.sniff_type(["12", "14.5", "16"], cfg), "number")
        self.assertEqual(cs.sniff_type(["007", "008", "009"], cfg), "text")
    def test_number_parser(self):
        cfg = cs.resolve_config({}, ["A"])
        self.assertEqual(cs.parse_number("١٤٫٥", cfg)[:2], (14.5, "ok")); self.assertEqual(cs.parse_number("0", cfg)[:2], (0, "ok"))
        self.assertEqual(cs.parse_number("abc", cfg)[1], "fail"); self.assertEqual(cs.parse_number("-3", cfg)[:2], (-3, "ok"))
