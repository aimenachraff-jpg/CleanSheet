import sys, tempfile, unittest, hashlib, json
from datetime import date, datetime
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import openpyxl
import clean_sheet as cs

CFG = cs.resolve_config({}, ["Date"])  # defaults

class Parsers(unittest.TestCase):
    def test_price_clear(self):
        for raw, exp in [("2500 DA", 2500.0), ("3 900 DA", 3900.0), ("1800DA", 1800.0), ("2500.00", 2500.0),
                         ("1.234,50", 1234.5), ("1,234.50", 1234.5), ("2,50", 2.5), (2500, 2500.0), ("١٢٣٤", 1234.0)]:
            v, s, _ = cs.parse_price(raw, CFG); self.assertEqual((v, s), (exp, "ok"), raw)
    def test_price_ambiguous_flagged(self):
        v, s, m = cs.parse_price("2,500", CFG); self.assertEqual((v, s), (2500.0, "verify")); self.assertIn("ambiguous", m)
        self.assertEqual(cs.parse_price("2.500", CFG)[1], "verify")
    def test_price_bad(self):
        for raw in ["abc", "12 3", "1,23,4"]:
            v, s, _ = cs.parse_price(raw, CFG); self.assertEqual((v, s), (raw, "fail"), raw)
        self.assertEqual(cs.parse_price("", CFG)[1], "missing"); self.assertEqual(cs.parse_price("0", CFG)[1], "verify")
        self.assertEqual(cs.parse_price("-5", CFG)[1], "verify")
    def test_date(self):
        self.assertEqual(cs.parse_date("2026-03-01", CFG)[:2], (date(2026, 3, 1), "ok"))
        self.assertEqual(cs.parse_date("March 2, 2026", CFG)[:2], (date(2026, 3, 2), "ok"))
        self.assertEqual(cs.parse_date("2 mars 2026", CFG)[:2], (date(2026, 3, 2), "ok"))
        self.assertEqual(cs.parse_date("25/03/2026", CFG)[:2], (date(2026, 3, 25), "ok"))
        self.assertEqual(cs.parse_date("03/25/2026", CFG)[:2], (date(2026, 3, 25), "ok"))
        self.assertEqual(cs.parse_date(datetime(2026, 3, 1), CFG)[:2], (date(2026, 3, 1), "ok"))
    def test_date_ambiguous_and_bad(self):
        v, s, m = cs.parse_date("03/04/2026", CFG); self.assertEqual((v, s), (date(2026, 4, 3), "verify")); self.assertIn("ambiguous", m)
        v, s, _ = cs.parse_date("03/04/2026", {**CFG, "date_order": "mdy"}); self.assertEqual(v, date(2026, 3, 4))
        self.assertEqual(cs.parse_date("31/02/2026", CFG)[1], "fail"); self.assertEqual(cs.parse_date("soon", CFG)[1], "fail")
        self.assertEqual(cs.parse_date("01/02/26", CFG)[1], "verify"); self.assertEqual(cs.parse_date("", CFG)[1], "missing")
    def test_phone(self):
        for raw in ["0555 12 34 56", "+213555123456", "0555-12-34-56", "00213 555 12 34 56", "+213 (0)555123456"]:
            self.assertEqual(cs.parse_phone(raw, CFG)[:2], ("+213555123456", "ok"), raw)
        v, s, m = cs.parse_phone("555123456", CFG); self.assertEqual((v, s), ("+213555123456", "verify"))
        self.assertEqual(cs.parse_phone("12345", CFG)[1], "fail")
        self.assertEqual(cs.parse_phone("+33 6 12 34 56 78", CFG)[:2], ("+33612345678", "ok"))
    def test_qty(self):
        for raw, exp in [("2", 2), (" 3 ", 3), ("two", 2), ("deux", 2), (2.0, 2), ("٣", 3)]:
            self.assertEqual(cs.parse_qty(raw, CFG)[:2], (exp, "ok"), raw)
        self.assertEqual(cs.parse_qty("2.5", CFG)[1], "fail"); self.assertEqual(cs.parse_qty(-1, CFG)[1], "fail")
    def test_none_is_missing_not_text(self):
        for f in (cs.parse_price, cs.parse_date, cs.parse_phone, cs.parse_qty, cs.parse_email):
            self.assertEqual(f(None, CFG)[1], "missing", f.__name__)
    def test_email(self):
        self.assertEqual(cs.parse_email("  Foo@Bar.COM ", CFG)[:2], ("foo@bar.com", "ok")); self.assertEqual(cs.parse_email("foo@", CFG)[1], "fail")

class Text(unittest.TestCase):
    def test_unify_case_whitespace_and_flag_similar(self):
        res = cs.clean_text_column(["  ahmed  benali", "Ahmed Benali", "AHMED BENALI", "Ahmed Benali", "Sara"], CFG, "Customer", "name")
        self.assertEqual({r[0] for r in res[:4]}, {"Ahmed Benali"}); self.assertEqual(res[4][0], "Sara")
        res = cs.clean_text_column(["T-Shirt", "T-Shirt", "Tshirt", "Cap"], CFG, "P", "category")
        self.assertEqual(res[2][1], "verify"); self.assertEqual(res[2][0], "Tshirt")  # flagged, not merged
        self.assertEqual(res[3][1], "ok")
    def test_category_map(self):
        cfg = {**CFG, "category_maps": {"City": {"algiers": "Alger"}}}
        self.assertEqual(cs.clean_text_column(["Algiers", "Alger"], cfg, "City", "category")[0][0], "Alger")
    def test_lowercase_only_titlecased(self):
        self.assertEqual(cs.clean_text_column(["lina haddad", "LINA HADDAD"], CFG, "C", "name")[0][0], "Lina Haddad")
    def test_arabic_untouched(self):
        self.assertEqual(cs.clean_text_column(["محمد  علي"], CFG, "C", "name")[0][0], "محمد علي")

def make_xlsx(path, rows, headers=("OrderID", "Customer", "City", "Qty", "Price", "Date", "Phone")):
    wb = openpyxl.Workbook(); ws = wb.active; ws.append(list(headers))
    for r in rows: ws.append(r)
    wb.save(path)

class Workflow(unittest.TestCase):
    def test_end_to_end_xlsx(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d); src = d / "in.xlsx"
            make_xlsx(src, [
                [1, " ahmed benali", "oran ", "2", "2,500", "03/04/2026", "0555 12 34 56"],
                [2, "Ahmed Benali", "Oran", 1, 1800, datetime(2026, 3, 5), "+213555123456"],
                [2, "Ahmed Benali", "Oran", 1, 1800, datetime(2026, 3, 5), "+213555123456"],   # exact dup
                [3, "Sara", "Alger", "two", "abc", "", "12345"],
                [None, None, None, None, None, None, None],                                      # blank row
                [4, "Ahmed Benali", "Oran", 1, 1800, datetime(2026, 3, 5), "+213555123456"],     # same content, new ID
                [5, "Zed", "Oran", 1, 100, datetime(2026, 3, 5), "0661001122"],
                [5, "Yan", "Oran", 1, 100, datetime(2026, 3, 6), "0661001123"],                  # reused ID
            ])
            before = hashlib.md5(src.read_bytes()).hexdigest()
            r = cs.run(src, out_dir=d / "out")
            self.assertEqual(before, hashlib.md5(src.read_bytes()).hexdigest(), "input modified!")
            self.assertEqual((r["rows_in"], r["rows_out"], r["removed"]), (7, 6, 1))
            wb = openpyxl.load_workbook(r["xlsx"]); self.assertEqual(wb.sheetnames[:4], ["Summary", "Clean", "Review", "Change log"])
            clean = list(wb["Clean"].values); h = clean[0]; row1 = dict(zip(h, clean[1]))
            self.assertEqual(row1["Customer"], "Ahmed Benali"); self.assertEqual(row1["City"], "Oran")
            self.assertEqual(row1["Price"], 2500.0); self.assertEqual(row1["Phone"], "+213555123456")
            self.assertEqual(row1["Total (qty x price)"], 5000.0)
            self.assertEqual(row1["Date"].date() if isinstance(row1["Date"], datetime) else row1["Date"], date(2026, 4, 3))
            row3 = dict(zip(h, clean[3]))   # Sara row: unresolved price/phone kept as original, missing date
            self.assertEqual((row3["Price"], row3["Phone"], row3["Date"]), ("abc", "12345", None)); self.assertIsNone(row3["Total (qty x price)"])
            review = list(wb["Review"].values)[1:]; notes = " | ".join(str(x[5]) for x in review)
            for needle in ["ambiguous", "could not read price", "date missing", "phone has", "possible duplicate", "ID also used"]:
                self.assertIn(needle, notes, needle)
            self.assertTrue(r["report"].exists())
            self.assertEqual(len(list(wb["Before-After (20 rows)"].values)), 7)
    def test_config_french_headers(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d); src = d / "fr.xlsx"
            make_xlsx(src, [["A1", "karim", "ORAN", "2", "1 200 DA", "5 mars 2026", "0770 11 22 33"]],
                      headers=("N°", "Client", "Ville", "Quantité", "Prix", "Date commande", "Téléphone"))
            cfgp = d / "c.json"; cfgp.write_text(json.dumps({"columns": {"Client": "name", "Quantité": "qty"}}))
            r = cs.run(src, cfgp, d / "o")
            self.assertEqual(r["cfg"]["columns"]["Prix"], "price"); self.assertEqual(r["cfg"]["columns"]["Téléphone"], "phone")
            self.assertEqual(r["cfg"]["columns"]["Date commande"], "date"); self.assertEqual(r["cfg"]["columns"]["N°"], "id")
            vals = list(openpyxl.load_workbook(r["xlsx"])["Clean"].values)[1]
            self.assertEqual((vals[1], vals[3], vals[4], vals[6]), ("Karim", 2, 1200.0, "+213770112233"))
    def test_bad_config_column_rejected(self):
        with self.assertRaises(SystemExit): cs.resolve_config({"columns": {"Nope": "price"}}, ["Price"])
        with self.assertRaises(SystemExit): cs.resolve_config({"columns": {"Price": "banana"}}, ["Price"])
    def test_csv_input_and_semicolon(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d); p = d / "x.csv"; p.write_text("Customer;Qty;Price\nali;1;1 000 DA\nali;1;1 000 DA\n", encoding="utf-8")
            r = cs.run(p, out_dir=d / "o", also_csv=True); self.assertEqual((r["rows_out"], r["removed"]), (1, 1))
            self.assertTrue((d / "o" / "x_cleaned.csv").exists())

if __name__ == "__main__": unittest.main(verbosity=1)
