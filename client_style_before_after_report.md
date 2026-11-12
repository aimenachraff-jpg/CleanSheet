# Cleaning report - client_style_sample_before.xlsx

- Rows read: 25 | exact duplicates removed: 1 | rows in cleaned file: 24
- Auto-fixed cells: 42 | to VERIFY: 5 | UNRESOLVED: 2 | MISSING: 1

## What was changed automatically

- Client: 1 cells (whitespace, casing, format)
- Ville: 2 cells (whitespace, casing, format)
- Quantité: 1 cells (whitespace, casing, format)
- Prix: 14 cells (whitespace, casing, format)
- Date: 1 cells (whitespace, casing, format)
- Téléphone: 23 cells (whitespace, casing, format)

## Needs a human decision

- Quantité `2.5` x1 -> UNRESOLVED: quantity is not a whole non-negative number (rows 19)
- Prix `1,450` x1 -> VERIFY: '1,450' is ambiguous (thousands vs decimal); read as 1450 (rows 7)
- Prix `abc` x1 -> UNRESOLVED: could not read price (rows 13)
- Date `02/09/2026` x1 -> VERIFY: ambiguous day/month in '02/09/2026'; read as DMY (rows 4)
- Date `05/09/2026` x1 -> VERIFY: ambiguous day/month in '05/09/2026'; read as DMY (rows 11)
- Date `10/09/2026` x1 -> VERIFY: ambiguous day/month in '10/09/2026'; read as DMY (rows 20)
- Téléphone `(blank)` x1 -> MISSING: phone missing (rows 8)
- Téléphone `770112233` x1 -> VERIFY: leading 0 missing in source; assumed local number (rows 9)

Nothing flagged above was silently guessed: yellow/orange cells in the Clean sheet mark each one. The original file was not modified.