import sys, tempfile, unittest
from datetime import date
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import openpyxl
import clean_sheet as cs

MONTHS = ["janvier", "février", "mars", "avril", "juin", "juillet", "octobre", "décembre"]
SUBS = [None, "Qté", "prix unit", "mongtant", "Qter", "prix unit", "montant", "Qté", "mis"]
EXPECT = ["Mois", "Produit1 - Qté", "Produit1 - prix unit", "Produit1 - mongtant", "Produits2 - Qter", "Produits2 - prix unit",
          "Produits2 - montant", "Total Génération - Qté", "Total Génération - mis"]

def build(path, merged=False, top=1, left=1):
    wb = openpyxl.Workbook(); ws = wb.active
    c = lambda r, k, v: ws.cell(row=top + r, column=left + k, value=v)
    c(0, 0, "Mois")
    if merged:   # label in the first cell of a merged range
        for k0, k1, lab in ((1, 3, "Produit1"), (4, 6, "Produits2"), (7, 8, "Total Génération")):
            c(0, k0, lab); ws.merge_cells(start_row=top, start_column=left + k0, end_row=top, end_column=left + k1)
    else:        # labels as in the screenshot: one cell, blanks around it
        c(0, 2, "Produit1"); c(0, 5, "Produits2"); c(0, 7, "Total Génération")
    for k, h in enumerate(SUBS):
        if h: c(1, k, h)
    for i, m in enumerate(MONTHS, start=2):
        q = 1000 + 100 * i; pu = 15.5 - i * 0.2
        for k, v in enumerate([m, q, pu, q * pu, q + 200, pu, (q + 200) * pu, q + 200, q]): c(i, k, v)
        ws.cell(row=top + i, column=left + 2).number_format = '#,##0.00 "DZD"'
    r = 2 + len(MONTHS)
    c(r, 0, "Totale")
    for k in range(1, 9): c(r, k, 12345 + k)
    wb.save(path)

class Multi(unittest.TestCase):
    def check(self, p):
        info = cs.read_table_info(p)
        self.assertEqual(info["headers"], EXPECT); self.assertEqual(info["levels"], 2)
        self.assertEqual(len(info["rows"]), len(MONTHS)); self.assertEqual([k["reason"] for k in info["skipped"]], ["total row"])
        cfg = cs.resolve_config({}, info["headers"], info["rows"], info["detect_names"])["columns"]
        self.assertEqual([cfg[h] for h in EXPECT], ["category", "qty", "price", "price", "qty", "price", "price", "qty", "number"])
        return info
    def test_screenshot_layout_unmerged(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "a.xlsx"; build(p); self.check(p)
            r = cs.run(p, out_dir=Path(d) / "o"); wb = openpyxl.load_workbook(r["xlsx"])
            self.assertIn("Skipped rows", wb.sheetnames); self.assertEqual(list(wb["Skipped rows"].values)[1][:2], (11, "total row"))
            clean = list(wb["Clean"].values); self.assertNotIn("Total (qty x price)", clean[0])            # ambiguous: several qty/price columns
            self.assertEqual(len(clean) - 1, len(MONTHS)); self.assertIsInstance(clean[1][1], (int, float)); self.assertIsInstance(clean[1][8], (int, float))
            self.assertEqual(clean[1][0], "Janvier")                      # all-lowercase text is title-cased and logged in the Change log
    def test_merged_group_cells(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "b.xlsx"; build(p, merged=True); self.check(p)
    def test_offset_and_title_above(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "c.xlsx"; build(p, top=4, left=3)
            wb = openpyxl.load_workbook(p); ws = wb.active; ws["C1"] = "Rapport annuel"; ws.merge_cells("C1:K1"); wb.save(p)
            info = self.check(p); self.assertEqual(info["header_row"], 5)
    def test_levels_override_and_keep_totals(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "d.xlsx"; build(p)
            info = cs.read_table_info(p, levels=1); self.assertEqual(info["headers"][1], "Qté"); self.assertEqual(info["levels"], 1)
            info = cs.read_table_info(p, skip_totals=False); self.assertEqual((len(info["rows"]), info["skipped"]), (len(MONTHS) + 1, []))
    def test_single_title_is_not_a_group_row(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "e.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
            ws["A1"] = "Ventes 2026"; ws.merge_cells("A1:C1"); ws.append(["Client", "Qté", "Prix"]); ws.append(["a", 1, 5]); ws.append(["b", 2, 6]); wb.save(p)
            info = cs.read_table_info(p); self.assertEqual((info["headers"], info["levels"]), (["Client", "Qté", "Prix"], 1))
    def test_formulas_without_saved_results_warn(self):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "f.xlsx"; wb = openpyxl.Workbook(); ws = wb.active
            ws.append(["A", "B", "C"]); ws.append([2, 3, "=A2*B2"]); ws.append([4, 5, "=A3*B3"]); wb.save(p)
            info = cs.read_table_info(p); self.assertTrue(any("formula" in w for w in info["warnings"]))

class ArabicDetection(unittest.TestCase):
    def test_hamza_and_diacritics_in_headers(self):
        for h, t in [("الإسم", "name"), ("الاسم الكامل", "name"), ("الكميّة", "qty"), ("الكميه", "qty"), ("سعر الوحدة", "price"), ("الثمن", "price"),
                     ("رقم الهاتف", "phone"), ("المدينة", "category"), ("الولاية", "category"), ("الشهر", "category"), ("تاريخ الطلب", "date"), ("اللقب", "name")]:
            self.assertEqual(cs.detect_type(h), t, h)
    def test_header_typos(self):
        self.assertEqual(cs.detect_type("mongtant"), "price"); self.assertEqual(cs.detect_type("Qter"), "qty"); self.assertEqual(cs.detect_type("montant"), "price")
    def test_arabic_month_names(self):
        cfg = cs.resolve_config({}, ["D"])
        for raw, exp in [("15 جانفي 2026", date(2026, 1, 15)), ("٣ جوان ٢٠٢٦", date(2026, 6, 3)), ("5 كانون الثاني 2026", date(2026, 1, 5)),
                         ("12 أكتوبر 2026", date(2026, 10, 12)), ("1 فيفري 2026", date(2026, 2, 1)), ("20 جويلية 2026", date(2026, 7, 20))]:
            self.assertEqual(cs.parse_date(raw, cfg)[:2], (exp, "ok"), raw)
    def test_arabic_number_words_and_variants_flagged(self):
        cfg = cs.resolve_config({}, ["Q"]); self.assertEqual(cs.parse_qty("ثلاثة", cfg)[:2], (3, "ok")); self.assertEqual(cs.parse_qty("اثنين", cfg)[:2], (2, "ok"))
        res = cs.clean_text_column(["أحمد", "أحمد", "احمد", "سارة"], cfg, "N", "name")
        self.assertEqual(res[2][1], "verify"); self.assertEqual(res[2][0], "احمد")      # hamza variant flagged, not merged
    def test_native_numbers_stay_numbers_in_text_columns(self):
        cfg = cs.resolve_config({}, ["T"]); res = cs.clean_text_column([1000, 15.5, "abc"], cfg, "T", "text")
        self.assertEqual([r[0] for r in res], [1000, 15.5, "Abc"])

if __name__ == "__main__": unittest.main()
