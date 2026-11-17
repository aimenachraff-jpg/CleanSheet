#!/usr/bin/env python3
"""Tests for the locale packs (clean_sheet_locales.py) and --locale wiring."""
import unittest
from clean_sheet import parse_phone, parse_price, merge_locale, resolve_config
from clean_sheet_locales import LOCALES


def phone(v, locale):
    return parse_phone(v, LOCALES[locale])


class TestLocalePhones(unittest.TestCase):
    def test_dz(self):
        self.assertEqual(phone("0555123456", "dz"), ("+213555123456", "ok", ""))
        self.assertEqual(phone("0555 12 34 56", "dz")[0], "+213555123456")

    def test_ma(self):
        self.assertEqual(phone("0612345678", "ma"), ("+212612345678", "ok", ""))
        self.assertEqual(phone("+212 6 12 34 56 78", "ma")[0], "+212612345678")
        # a landline prefix is plausible in Morocco, so it is verified, not failed
        st = phone("0522123456", "ma")[1]
        self.assertEqual(st, "verify")

    def test_tn(self):
        self.assertEqual(phone("20123456", "tn"), ("+21620123456", "ok", ""))
        self.assertEqual(phone("+216 20 123 456", "tn")[0], "+21620123456")
        # 7 digits is wrong for Tunisia either way
        self.assertEqual(phone("2012345", "tn")[1], "fail")

    def test_eg(self):
        self.assertEqual(phone("01012345678", "eg"), ("+201012345678", "ok", ""))
        self.assertEqual(phone("+20 10 1234 5678", "eg")[0], "+201012345678")

    def test_fr(self):
        self.assertEqual(phone("0612345678", "fr"), ("+33612345678", "ok", ""))
        self.assertEqual(phone("+33 6 12 34 56 78", "fr")[0], "+33612345678")

    def test_generic(self):
        # no country assumed: plausible numbers kept as written, marked verify
        out = phone("020 1234 5678", "generic")
        self.assertEqual(out[0], "02012345678")
        self.assertEqual(out[1], "verify")
        # international form is normalised and trusted
        self.assertEqual(phone("+44 20 1234 5678", "generic"), ("+442012345678", "ok", ""))
        # nonsense length fails
        self.assertEqual(phone("12345", "generic")[1], "fail")

    def test_same_number_different_country(self):
        self.assertEqual(phone("0612345678", "fr")[0], "+33612345678")
        self.assertEqual(phone("0612345678", "ma")[0], "+212612345678")


class TestLocalePrices(unittest.TestCase):
    def test_moroccan_dirham(self):
        cfg = dict(LOCALES["ma"])
        self.assertEqual(parse_price("1 250 DH", cfg)[0], 1250)
        self.assertEqual(parse_price("1.250,50 درهم", cfg)[0], 1250.5)

    def test_tunisian_dinar(self):
        cfg = dict(LOCALES["tn"])
        self.assertEqual(parse_price("12,5 DT", cfg), (12.5, "ok", ""))
        # ambiguous comma+3 digits: read as thousands but flagged for review, never silently chosen
        out = parse_price("12,500 DT", cfg)
        self.assertEqual(out[0], 12500)
        self.assertEqual(out[1], "verify")
        self.assertEqual(parse_price("12.500 د.ت", cfg)[0], 12500)

    def test_egyptian_pound(self):
        cfg = dict(LOCALES["eg"])
        self.assertEqual(parse_price("1,250 LE", cfg)[0], 1250)


class TestLocaleWiring(unittest.TestCase):
    def test_user_config_beats_locale(self):
        merged = merge_locale({"country_code": "999"}, "ma")
        self.assertEqual(merged["country_code"], "999")
        self.assertEqual(merged["national_length"], 9)  # from ma

    def test_locale_fills_defaults(self):
        merged = merge_locale({}, "tn")
        self.assertEqual(merged["country_code"], "216")
        self.assertEqual(merged["national_length"], 8)

    def test_unknown_locale_is_noop(self):
        self.assertEqual(merge_locale({}, "zz"), {})

    def test_all_locales_have_required_keys(self):
        for name, loc in LOCALES.items():
            self.assertIn("country_code", loc, name)
            self.assertIn("national_length", loc, name)
            self.assertIn("currency_tokens", loc, name)
            self.assertIn("date_order", loc, name)

    def test_resolve_config_accepts_locale_keys(self):
        merged = merge_locale({}, "ma")
        cfg = resolve_config(merged, ["Prix", "Téléphone"])
        self.assertEqual(cfg["country_code"], "212")


if __name__ == "__main__":
    unittest.main()
