#!/usr/bin/env python3
"""Locale packs for clean_sheet.py.

A locale is a set of country-specific facts the cleaner needs: how phone
numbers are written, which currency words appear in price cells, and any
month names the global month table does not already know.

Each pack only sets defaults. Anything passed in a --config file wins.

The default locale is "dz" (Algeria): that is where the project started.
Use --locale generic for files that follow no single country's conventions.
"""

LOCALES = {
    # Algeria: +213, 9 national digits (mobiles 05/06/07), DA/دج prices,
    # Maghrebi French-Arabic month names (janvier, جانفي, جويلية...).
    "dz": {
        "country_code": "213",
        "national_length": 9,
        "mobile_prefixes": "23456789",
        "leading_zero_required": True,
        "date_order": "dmy",
        "currency_tokens": ["dzd", "da", "dinars", "dinar", "eur", "usd", "€", "$",
                            "دينار جزائري", "دينار", "دنانير", "د.ج", "د ج", "دج"],
        "months": {},
    },
    # Morocco: +212, 9 national digits (mobiles 06/07), MAD/درهم/DH prices,
    # Moroccan month names (غشت for August, شتنبر for September, دجنبر for December).
    "ma": {
        "country_code": "212",
        "national_length": 9,
        "mobile_prefixes": "67",
        "leading_zero_required": True,
        "date_order": "dmy",
        "currency_tokens": ["mad", "dhs", "dh", "dirhams", "dirham", "eur", "usd", "€", "$",
                            "الدرهم", "درهم", "د.م", "د م"],
        "months": {"غشت": 8, "شتنبر": 9, "دجنبر": 12},
    },
    # Tunisia: +216, 8 national digits (fixed and mobile share one range),
    # TND/DT/دينار prices. No mobile-prefix check: prefixes say nothing here.
    "tn": {
        "country_code": "216",
        "national_length": 8,
        "mobile_prefixes": "",
        "leading_zero_required": False,
        "date_order": "dmy",
        "currency_tokens": ["tnd", "dt", "dinars", "dinar", "millimes", "millime", "eur", "usd", "€", "$",
                            "دينار", "دنانير", "د.ت", "د ت", "مليم"],
        "months": {},
    },
    # Egypt: +20, 10 national digits (mobiles 010/011/012/015), EGP/جنيه prices,
    # Eastern-Arabic month names (يناير, أغسطس...) already in the global table.
    "eg": {
        "country_code": "20",
        "national_length": 10,
        "mobile_prefixes": "101215",
        "leading_zero_required": True,
        "date_order": "dmy",
        "currency_tokens": ["egp", "le", "جنيه", "ج.م", "ج م", "usd", "$"],
        "months": {},
    },
    # France: +33, 9 national digits (mobiles 06/07), EUR prices, French months.
    "fr": {
        "country_code": "33",
        "national_length": 9,
        "mobile_prefixes": "67",
        "leading_zero_required": True,
        "date_order": "dmy",
        "currency_tokens": ["eur", "€", "usd", "$", "euros", "euro"],
        "months": {},
    },
    # Any other country: phones are only checked for a plausible length
    # (8 to 15 digits), kept in their original form when international,
    # and month names come from the shared global table
    # (English, French, standard Arabic, Maghrebi and Levantine variants).
    "generic": {
        "country_code": "",
        "national_length": None,
        "mobile_prefixes": "",
        "leading_zero_required": False,
        "date_order": "dmy",
        "currency_tokens": ["eur", "usd", "€", "$", "dinars", "dinar", "دينار"],
        "months": {},
    },
}

DESCRIPTIONS = {
    "dz": "Algeria (+213, DA/دج, جانفي)",
    "ma": "Morocco (+212, MAD/درهم, غشت)",
    "tn": "Tunisia (+216, DT/د.ت, 8-digit numbers)",
    "eg": "Egypt (+20, EGP/جنيه)",
    "fr": "France (+33, €)",
    "generic": "any other country (length checks only)",
}
