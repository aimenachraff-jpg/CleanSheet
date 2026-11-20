# Contributing

Thanks for wanting to help. The rules that keep this project what it is:

1. **Flag, never guess.** If the tool is unsure about a value, it highlights it and leaves the decision to the user. No feature should silently pick an interpretation.
2. **No new dependencies.** openpyxl is the only one. Standard library for everything else.
3. **No real data in the repo.** Sample files and tests use fictional names and phones only. Never paste a real client file into an issue; anonymize first.
4. **Every change gets a test.** The suite runs on every push.

## Getting set up

```
pip install openpyxl
python test_clean_sheet.py        # and the other test_*.py files
```

## Good first contributions

### A new locale pack

Copy the pattern in `clean_sheet_locales.py`. For a country you need three facts:

- `country_code` and `national_length`: how phone numbers are written (check a few real numbers online, do not guess)
- `currency_tokens`: the words that appear in price cells (English, French or Arabic, whatever locals actually type)
- `mobile_prefixes` and `leading_zero_required`: only if you are sure; leave empty otherwise

Then add a test in `test_locale_packs.py` following `test_ma` as the model, and add the locale to the README table. One pull request per country keeps review quick.

### A new column type

Column types live in `TYPE_RULES_SRC` in `clean_sheet.py` (header matching) and as `parse_*` functions (value parsing). A type that parses but cannot be decided should return `verify`, not pick a value.

### Fixing a broken file

See the "My file broke it" issue template. Reproduce with a fictional version of the file, add the case as a test, then fix.

## Before you open a pull request

```
python test_clean_sheet.py
python test_table_detection.py
python test_multilevel.py
python test_gui_logic.py
python test_locale_packs.py
python test_csv_encoding.py
python test_gui_locale.py
python test_gui_thread.py
python test_reliability.py
```

Keep the README honest: if you add a flag or a locale, document it there in the same plain style.
