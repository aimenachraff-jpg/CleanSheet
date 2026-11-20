"""Reliability and regression tests (Phase 2).

Covers: ambiguous duplicates / false positives, dates, decimal separators,
currencies, phone numbers, missing values, malformed and large files,
original-file-untouched, output reproducibility, config round-trip.
"""
import sys, tempfile, time, unittest, hashlib
from datetime import date, datetime
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import openpyxl
import clean_sheet as cs

CFG = cs.resolve_config({}, ["D"])


def make_xlsx(path, rows, headers=("OrderID", "Customer", "City", "Qty", "Price", "Date", "Phone")):
    wb = openpyxl.Workbook(); ws = wb.active; ws.append(list(headers))
    for r in rows: ws.append(r)
    wb.save(path)


class ParsersExtended(unittest.TestCase):
    def test_price_currency_tokens_per_locale(self):
        cases = [
            (dict(cs.LOCALES["dz"]), "1 450 DA", 1450.0, "ok"),
            (dict(cs.LOCALES["dz"]), "دج 900", 900.0, "ok"),
            (dict(cs.LOCALES["ma"]), "1 250 DH", 1250.0, "ok"),
            (dict(cs.LOCALES["ma"]), "1.250,50 درهم", 1250.5, "ok"),
            (dict(cs.LOCALES["tn"]), "12,5 DT", 12.5, "ok"),
            # "1,250 LE" is ambiguous-by-design (comma+3 digits): value correct, flagged verify
            (dict(cs.LOCALES["eg"]), "1,250 LE", 1250.0, "verify"),
            (dict(cs.LOCALES["fr"]), "1 250 €", 1250.0, "ok"),
        ]
        for cfg, raw, exp, st in cases:
            v, s, _ = cs.parse_price(raw, cfg)
            self.assertEqual((v, s), (exp, st), (cfg.get("country_code"), raw))

    def test_date_two_digit_year_and_variants(self):
        self.assertEqual(cs.parse_date("01/02/26", CFG)[:2], (date(2026, 2, 1), "verify"))
        self.assertEqual(cs.parse_date("15 جانفي 2026", CFG)[:2], (date(2026, 1, 15), "ok"))
        self.assertEqual(cs.parse_date("5 كانون الثاني 2026", CFG)[:2], (date(2026, 1, 5), "ok"))
        self.assertEqual(cs.parse_date("1 فيفري 2026", CFG)[:2], (date(2026, 2, 1), "ok"))

    def test_phone_landline_verify_and_international(self):
        # Morocco landline prefix is plausible -> verify, not fail
        self.assertEqual(cs.parse_phone("0522123456", cs.LOCALES["ma"])[1], "verify")
        self.assertEqual(cs.parse_phone("+212 5 22 12 34 56", cs.LOCALES["ma"])[:2], ("+212522123456", "verify"))
        # Moroccan mobile prefix -> ok
        self.assertEqual(cs.parse_phone("+212 6 12 34 56 78", cs.LOCALES["ma"])[:2], ("+212612345678", "ok"))
        # Tunisia: 8 digits, no leading-zero warning
        self.assertEqual(cs.parse_phone("20123456", cs.LOCALES["tn"])[:2], ("+21620123456", "ok"))
        # generic keeps international form trusted
        self.assertEqual(cs.parse_phone("+44 20 1234 5678", cs.LOCALES["generic"])[:2], ("+442012345678", "ok"))

    def test_missing_values_kept_not_guessed(self):
        self.assertEqual(cs.parse_price(None, CFG)[1], "missing")
        self.assertEqual(cs.parse_date(None, CFG)[1], "missing")
        self.assertEqual(cs.parse_phone(None, CFG)[1], "missing")
        self.assertEqual(cs.parse_qty(None, CFG)[1], "missing")
        self.assertEqual(cs.parse_email(None, CFG)[1], "missing")


class DuplicatesAndFalsePositives(unittest.TestCase):
    def test_distinct_names_not_merged_false_positive(self):
        # Two genuinely different names must not be merged just because they look similar
        res = cs.clean_text_column(["Ahmed Benali", "Ali Benali", "Sara"], CFG, "Customer", "name")
        self.assertEqual(res[0][0], "Ahmed Benali"); self.assertEqual(res[1][0], "Ali Benali")
        self.assertEqual(res[0][1], "ok"); self.assertEqual(res[1][1], "ok")

    def test_duplicate_first_occurrence_kept(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d); src = d / "dup.xlsx"
            make_xlsx(src, [
                [1, "Ali", "Oran", "2", "100", "2026-01-01", "0555123456"],
                [1, "Ali", "Oran", "2", "100", "2026-01-01", "0555123456"],   # exact dup
                [2, "Sara", "Alger", "1", "50", "2026-01-02", "0661001122"],
            ])
            r = cs.run(src, out_dir=d / "o")
            self.assertEqual((r["rows_in"], r["rows_out"], r["removed"]), (3, 2, 1))
            clean = list(openpyxl.load_workbook(r["xlsx"])["Clean"].values)
            # first Ali row kept (OrderID 1), Sara present
            ids = [row[0] for row in clean[1:]]
            self.assertEqual(ids, [1, 2])


class Integrity(unittest.TestCase):
    def test_input_file_untouched_sha256(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d); src = d / "in.xlsx"
            make_xlsx(src, [["Ali", "1 000 DA"]], headers=("Customer", "Price"))
            before = hashlib.sha256(src.read_bytes()).hexdigest()
            cs.run(src, out_dir=d / "o")
            after = hashlib.sha256(src.read_bytes()).hexdigest()
            self.assertEqual(before, after)

    def _clean_values(self, xlsx_path):
        wb = openpyxl.load_workbook(xlsx_path)
        return [list(row) for row in wb["Clean"].values]

    def test_output_reproducible(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d); src = d / "in.xlsx"
            make_xlsx(src, [
                [1, " ahmed ", "Oran", "2", "2,500", "03/04/2026", "0555 12 34 56"],
                [2, "Sara", "Alger", "1", "abc", "", "12345"],
            ], headers=("ID", "Customer", "City", "Qty", "Price", "Date", "Phone"))
            r1 = cs.run(src, out_dir=d / "o1")
            r2 = cs.run(src, out_dir=d / "o2")
            # xlsx zip timestamps differ between saves; compare logical content
            self.assertEqual(self._clean_values(r1["xlsx"]), self._clean_values(r2["xlsx"]))
            self.assertEqual(r1["report"].read_text(encoding="utf-8"), r2["report"].read_text(encoding="utf-8"))
            self.assertEqual(r1["rows_out"], r2["rows_out"])

    def test_config_round_trip_reproduces(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d); src = d / "in.xlsx"
            make_xlsx(src, [["Ali", "1 000 DA", "03/04/2026", "0555123456"]],
                      headers=("Customer", "Price", "Date", "Phone"))
            cfgp = d / "c.json"
            r1 = cs.run(src, write_config=cfgp, out_dir=d / "o1")
            r2 = cs.run(src, cfg_path=cfgp, out_dir=d / "o2")
            self.assertEqual(self._clean_values(r1["xlsx"]), self._clean_values(r2["xlsx"]))
            self.assertEqual(r1["report"].read_text(encoding="utf-8"), r2["report"].read_text(encoding="utf-8"))


class LargeAndMalformed(unittest.TestCase):
    def test_large_file_completes(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d); src = d / "big.xlsx"
            wb = openpyxl.Workbook(); ws = wb.active
            ws.append(["ID", "Customer", "Price", "Phone"])
            for i in range(5000):
                ws.append([i, "Client %d" % i, "1 000 DA", "0555123456"])
            wb.save(src)
            t0 = time.time()
            r = cs.run(src, out_dir=d / "o")
            elapsed = time.time() - t0
            self.assertEqual((r["rows_in"], r["rows_out"]), (5000, 5000))
            self.assertLess(elapsed, 60, "large file should clean in well under a minute")

    def test_malformed_tiny_csv_no_crash(self):
        with tempfile.TemporaryDirectory() as d:
            d = Path(d); p = d / "tiny.csv"; p.write_text("Ali\nSara\n", encoding="utf-8")
            r = cs.run(p, out_dir=d / "o")
            self.assertEqual(r["rows_out"], 1)


if __name__ == "__main__": unittest.main(verbosity=1)
