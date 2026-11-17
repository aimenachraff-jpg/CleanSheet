#!/usr/bin/env python3
"""clean_sheet.py v2 - clean a messy CSV/XLSX table without silent guessing.

Usage:
  python clean_sheet.py input.xlsx [--config cfg.json] [--out-dir DIR] [--sheet NAME] [--header-row 1] [--also-csv]
  python clean_sheet.py input.csv --write-config cfg.json     # save auto-detected column mapping to edit

Principles: the input file is never modified; every change is logged; anything ambiguous
or unparseable is flagged for human review, never silently guessed. Free libs only: openpyxl.
"""
import argparse, csv, difflib, json, re, sys, unicodedata
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from clean_sheet_locales import LOCALES

DEFAULTS = {
    "columns": {}, "date_order": "dmy", "country_code": "213", "national_length": 9,
    "mobile_prefixes": "23456789",
    "leading_zero_required": True,
    "currency_tokens": ["dzd", "da", "dinars", "dinar", "eur", "usd", "€", "$", "دينار جزائري", "دينار", "دنانير", "د.ج", "د ج", "دج"],
    "category_maps": {}, "similarity_threshold": 0.88, "id_column": None,
}
NUM_WORDS = {"zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7,
             "eight": 8, "nine": 9, "ten": 10, "un": 1, "une": 1, "deux": 2, "trois": 3, "quatre": 4,
             "cinq": 5, "sept": 7, "huit": 8, "neuf": 9, "dix": 10, "واحد": 1, "اثنان": 2, "ثلاثة": 3}
MONTHS = {m: i + 1 for i, names in enumerate([
    ["jan", "january", "janvier"], ["feb", "february", "fev", "fevrier", "février", "févr"],
    ["mar", "march", "mars"], ["apr", "april", "avr", "avril"], ["may", "mai"], ["jun", "june", "juin"],
    ["jul", "july", "juil", "juillet"], ["aug", "august", "aout", "août"],
    ["sep", "sept", "september", "septembre"], ["oct", "october", "octobre"],
    ["nov", "november", "novembre"], ["dec", "december", "decembre", "décembre"]]) for m in names}
_AR_STRIP = re.compile("[\u064b-\u065f\u0670\u0640]")          # diacritics (tashkeel) and tatweel
_AR_MAP = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا", "ى": "ي", "ئ": "ي", "ؤ": "و", "ة": "ه"})
def ar_norm(s):
    """Normalise Arabic spelling variants for MATCHING only (hamza forms, ya/alef maqsura, ta marbuta, diacritics)."""
    return _AR_STRIP.sub("", s).translate(_AR_MAP)

AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹٫٬", "0123456789" "0123456789" ".,")
TYPE_RULES_SRC = [
    ("email", r"e-?mail|courriel|بريد"),
    ("phone", r"phone|mobile|gsm|whatsapp|^tel\b|^tél\b|téléphone|هاتف|جوال|نقال|موبايل"),
    ("date", r"date|تاريخ"),
    ("qty", r"qty|quantit|qt[eé]?r?\b|qte|^nbr|الكمية|كمية|العدد|عدد"),
    ("price", r"price|prix|mo?n?g?tant|amount|tarif|p\.?u\b|سعر|ثمن|تمن|مبلغ|المبلغ|قيمة"),
    ("number", r"^(score|grade|moyenne|mark|marks|points?)\b|نقطة|النقطة|المعدل|العلامة"),
    ("id", r"order ?id|^id$|^ref\b|^réf|^n°|^no$|^num\b|id$|^رقم$|^رقم\s|المرجع"),
    ("category", r"city|ville|wilaya|product|produit|item|article|categor|catégor|status|statut|^mois\b|^month\b|مدينة|ولاية|بلدية|منتج|سلعة|صنف|فئة|الحالة|شهر"),
    ("name", r"customer|client|name|nom|prénom|prenom|société|company|زبون|اسم|لقب|عميل"),
]
TYPE_RULES = [(t, re.compile(ar_norm(rx))) for t, rx in TYPE_RULES_SRC]
FUZZY_WORDS = {"price": ["montant", "prix", "price", "amount", "tarif"], "qty": ["quantite", "quantity"]}
TYPES = {"id", "name", "category", "text", "qty", "price", "number", "date", "phone", "email", "skip"}


MONTHS.update({ar_norm(k): v for k, v in {
    "جانفي": 1, "فيفري": 2, "مارس": 3, "أفريل": 4, "ماي": 5, "جوان": 6, "جويلية": 7, "أوت": 8, "سبتمبر": 9, "أكتوبر": 10, "نوفمبر": 11, "ديسمبر": 12,
    "يناير": 1, "فبراير": 2, "أبريل": 4, "مايو": 5, "يونيو": 6, "يوليو": 7, "أغسطس": 8,
    "كانون الثاني": 1, "شباط": 2, "آذار": 3, "نيسان": 4, "أيار": 5, "حزيران": 6, "تموز": 7, "آب": 8, "أيلول": 9,
    "تشرين الأول": 10, "تشرين الثاني": 11, "كانون الأول": 12,
    "دجنبر": 12}.items()})
NUM_WORDS.update({ar_norm(k): v for k, v in {"واحد": 1, "واحدة": 1, "اثنان": 2, "اثنين": 2, "ثلاثة": 3, "أربعة": 4, "خمسة": 5, "ستة": 6,
                                             "سبعة": 7, "ثمانية": 8, "تسعة": 9, "عشرة": 10}.items()})

def norm_text(v):
    if v is None: return ""
    s = unicodedata.normalize("NFC", str(v)).translate(AR_DIGITS)
    s = re.sub(r"[\u00a0\u202f\u200e\u200f\u200b]", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def norm_num(v):
    return "" if v is None else norm_text(str(v).translate(AR_DIGITS))

def detect_type(h):
    x = ar_norm(str(h).casefold().strip())
    for t, rx in TYPE_RULES:
        if rx.search(x): return t
    for w in re.findall(r"[^\W\d_]{4,}", x):                      # tolerate typos such as "mongtant"
        for t, words in FUZZY_WORDS.items():
            if any(difflib.SequenceMatcher(None, w, k).ratio() >= 0.8 for k in words): return t
    return "text"

# ---------- parsers: return (value, status, message). status: ok | verify | fail | missing ----------
def parse_qty(v, cfg):
    if isinstance(v, bool): return v, "fail", "unexpected boolean"
    if isinstance(v, (int, float)):
        if float(v).is_integer() and v >= 0: return int(v), "ok", ""
        return v, "fail", "quantity is not a whole non-negative number"
    s = norm_num(v).casefold()
    if not s: return None, "missing", "quantity missing"
    if re.fullmatch(r"\d+", s): return int(s), "ok", ""
    if re.fullmatch(r"\d+[.,]0+", s): return int(re.split(r"[.,]", s)[0]), "ok", ""
    if ar_norm(s) in NUM_WORDS: return NUM_WORDS[ar_norm(s)], "ok", ""
    return v, "fail", "could not read quantity"

def parse_price(v, cfg):
    if isinstance(v, bool): return v, "fail", "unexpected boolean"
    if isinstance(v, (int, float)):
        if v < 0: return float(v), "verify", "negative price"
        return float(v), "ok", ""
    s = norm_num(v)
    if not s: return None, "missing", "price missing"
    t = s
    for tok in sorted(cfg["currency_tokens"], key=len, reverse=True):
        t = re.sub(re.escape(tok), "", t, flags=re.I)
    t = t.strip().replace("'", "")
    if not re.fullmatch(r"-?[\d.,\s]+", t or "x"): return v, "fail", "could not read price"
    neg = t.startswith("-"); t = t.lstrip("-").strip()
    if " " in t:
        if not re.fullmatch(r"\d{1,3}( \d{3})+([.,]\d{1,2})?", t): return v, "fail", "odd spacing in price"
        t = t.replace(" ", "")
    msg, status = "", "ok"
    if "," in t and "." in t:
        dec = "," if t.rfind(",") > t.rfind(".") else "."
        thou = "." if dec == "," else ","
        t = t.replace(thou, "").replace(dec, ".")
    elif "," in t or "." in t:
        sep = "," if "," in t else "."
        parts = t.split(sep)
        if len(parts) > 2:
            if all(len(p) == 3 for p in parts[1:]) and 1 <= len(parts[0]) <= 3: t = "".join(parts)
            else: return v, "fail", "inconsistent separators in price"
        else:
            after = parts[1]
            if len(after) == 3 and 1 <= len(parts[0]) <= 3:
                t = "".join(parts); status = "verify"
                msg = f"'{s}' is ambiguous (thousands vs decimal); read as {t}"
            elif len(after) in (1, 2): t = parts[0] + "." + after
            else: return v, "fail", "could not read decimals in price"
    try: val = float(t)
    except ValueError: return v, "fail", "could not read price"
    if neg: val = -val; status, msg = "verify", "negative price"
    if val == 0: status, msg = "verify", "price is zero"
    return val, status, msg

def _mkdate(y, m, d):
    try: return date(y, m, d)
    except ValueError: return None

def parse_date(v, cfg):
    if isinstance(v, datetime): return v.date(), "ok", ""
    if isinstance(v, date): return v, "ok", ""
    s = norm_num(v)
    if not s: return None, "missing", "date missing"
    m = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?", s)
    if m:
        d = _mkdate(int(m[1]), int(m[2]), int(m[3]))
        return (d, "ok", "") if d else (v, "fail", "impossible date")
    m = re.fullmatch(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})", s)
    if m:
        a, b, y = int(m[1]), int(m[2]), int(m[3]); status, msg = "ok", ""
        if y < 100: y += 2000; status, msg = "verify", "two-digit year read as 20xx"
        if a > 12 and b <= 12: day, mon = a, b
        elif b > 12 and a <= 12: day, mon = b, a
        elif a == b: day, mon = a, b
        elif a <= 12 and b <= 12:
            day, mon = (a, b) if cfg["date_order"] == "dmy" else (b, a)
            status, msg = "verify", f"ambiguous day/month in '{s}'; read as {cfg['date_order'].upper()}"
        else: return v, "fail", "impossible date"
        d = _mkdate(y, mon, day)
        return (d, status, msg) if d else (v, "fail", "impossible date")
    s = ar_norm(s)
    for rx, order in ((r"(\d{1,2})\s+([^\W\d_]+(?:\s[^\W\d_]+)?)\.?,?\s+(\d{4})", "dmy"), (r"([^\W\d_]+(?:\s[^\W\d_]+)?)\.?\s+(\d{1,2}),?\s+(\d{4})", "mdy")):
        m = re.fullmatch(rx, s)
        if m:
            day, mon, y = (m[1], m[2], m[3]) if order == "dmy" else (m[2], m[1], m[3])
            mo = MONTHS.get(mon.casefold())
            d = _mkdate(int(y), mo, int(day)) if mo else None
            return (d, "ok", "") if d else (v, "fail", "could not read date")
    return v, "fail", "could not read date"

def parse_phone(v, cfg):
    s = norm_num(v)
    if not s: return None, "missing", "phone missing"
    cc, nl = cfg["country_code"], cfg["national_length"]
    digits = re.sub(r"\D", "", s)
    intl = s.startswith("+") or s.startswith("00")
    if s.startswith("00"): digits = digits[2:]
    status, msg = "ok", ""
    if not cc:  # generic locale: no country assumed, only plausibility
        if len(digits) < 8 or len(digits) > 15: return v, "fail", f"phone has {len(digits)} digits, expected between 8 and 15"
        if intl: return ("+" if s.startswith("+") else "") + digits, "ok", ""
        return digits, "verify", "no country code set (--locale); number kept as written"
    if intl:
        if digits.startswith(cc): nat = digits[len(cc):]
        elif 8 <= len(digits) <= 15: return "+" + digits, "ok", ""
        else: return v, "fail", "could not read phone number"
        if nat.startswith("0"): nat = nat[1:]
    elif digits.startswith("0"): nat = digits[1:]
    elif digits.startswith(cc) and len(digits) == len(cc) + nl: nat = digits[len(cc):]
    else:
        nat = digits
        if len(nat) == nl and cfg.get("leading_zero_required", True):
            status, msg = "verify", "leading 0 missing in source; assumed local number"
    if len(nat) != nl: return v, "fail", f"phone has {len(nat)} digits after country code, expected {nl}"
    prefixes = cfg.get("mobile_prefixes") or ""
    if prefixes and nat[0] not in prefixes: status, msg = "verify", f"unusual +{cc} number prefix"
    return f"+{cc}{nat}", status, msg

def parse_email(v, cfg):
    s = norm_text(v).lower()
    if not s: return None, "missing", "email missing"
    if re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]{2,}", s): return s, "ok", ""
    return v, "fail", "not a valid email format"

def parse_number(v, cfg):
    """Generic number (grade, score, measure): Arabic/Persian digits and decimal separators converted; ambiguous separators flagged."""
    val, st, msg = parse_price(v, cfg)
    if st in ("missing", "fail"): return val, st, msg.replace("price", "number")
    if msg in ("price is zero", "negative price"): st, msg = "ok", ""
    if isinstance(val, float) and val.is_integer(): val = int(val)
    return val, st, msg

PARSERS = {"qty": parse_qty, "price": parse_price, "number": parse_number, "date": parse_date, "phone": parse_phone, "email": parse_email}

def same(orig, new):
    if orig is None: return new in (None, "")
    if isinstance(new, date): return isinstance(orig, (date, datetime)) or str(orig).strip() == new.isoformat()
    if isinstance(new, (int, float)) and not isinstance(new, bool):
        if isinstance(orig, (int, float)): return float(orig) == float(new)
        if any(ch.isdigit() and not ch.isascii() for ch in str(orig)): return False
        try: return float(str(orig).strip()) == float(new)
        except ValueError: return False
    return str(orig).strip() == str(new)

def proper(v):
    return v.lower() == v.upper() or (v != v.lower() and v != v.upper())

def clean_text_column(raw, cfg, header, kind):
    maps = {k.casefold(): v for k, v in cfg["category_maps"].get(header, {}).items()}
    native = [isinstance(r, (int, float, datetime, date)) and not isinstance(r, bool) for r in raw]   # numbers/dates are never turned into text
    vals = []
    for r, nat in zip(raw, native):
        n = "" if nat else norm_text(r)
        if n and n.casefold() in maps: n = maps[n.casefold()]
        vals.append(n)
    freq = Counter(v for v in vals if v)
    groups = defaultdict(list)
    for v in freq: groups[v.casefold()].append(v)
    canon = {}
    for variants in groups.values():
        good = [x for x in variants if proper(x)]
        pick = max(good, key=lambda x: freq[x]) if good else max(variants, key=lambda x: freq[x]).title()
        for x in variants: canon[x] = pick
    cfreq = Counter()
    for v, c in freq.items(): cfreq[canon[v]] += c
    key = lambda x: re.sub(r"[\W_]+", "", ar_norm(x.casefold()))
    flagged = {}
    if 1 < len(cfreq) <= 300:
        items = sorted(cfreq, key=lambda x: -cfreq[x])
        for i, a_ in enumerate(items):
            for b_ in items[:i]:
                same_key = key(a_) == key(b_)
                if same_key or difflib.SequenceMatcher(None, key(a_), key(b_)).ratio() >= cfg["similarity_threshold"]:
                    flagged[a_] = f"may be the same as '{b_}' ({'differs only by punctuation, spacing or Arabic letter variants' if same_key else 'similar spelling'}); not merged"
                    break
    out = []
    for r, v, nat in zip(raw, vals, native):
        if nat: out.append((r, "ok", "")); continue
        if not v:
            out.append((None, "missing" if kind != "text" else "ok", "value missing" if kind != "text" else "")); continue
        c = canon[v]
        out.append((c, "verify", flagged[c]) if c in flagged else (c, "ok", ""))
    return out

def sniff_type(values, cfg):
    """Guess a column type from its CONTENT (used only when the header gives no clue)."""
    vals = [v for v in values if norm_text(v) != ""][:200]
    if len(vals) < 2: return "text"
    frac = lambda f: sum(1 for v in vals if f(v)) / len(vals)
    if frac(lambda v: isinstance(v, str) and re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]{2,}", norm_text(v).lower())) >= 0.8: return "email"
    def isphone(v):
        t = re.sub(r"[\s().\-]", "", norm_num(v))
        return bool(re.fullmatch(r"(\+|00)\d{8,14}|0\d{8,10}|[567]\d{8}", t))
    if frac(isphone) >= 0.8: return "phone"
    if frac(lambda v: isinstance(v, (datetime, date)) or (isinstance(v, str) and parse_date(v, cfg)[1] in ("ok", "verify"))) >= 0.8: return "date"
    toks = cfg["currency_tokens"]
    if frac(lambda v: isinstance(v, str) and any(re.search(re.escape(t), norm_num(v), re.I) for t in toks) and parse_price(v, cfg)[1] in ("ok", "verify")) >= 0.6: return "price"
    def isnum(v):
        if isinstance(v, bool): return False
        if isinstance(v, (int, float)): return True
        t = norm_num(v)
        return bool(re.fullmatch(r"-?\d+([.,]\d+)?", t)) and not re.fullmatch(r"0\d+", t)
    if frac(isnum) >= 0.8: return "number"
    return "text"

def merge_locale(user_cfg, locale):
    """Layer a locale pack under the user's config; explicit --config keys win."""
    loc = LOCALES.get(locale or "")
    if not loc: return user_cfg
    merged = {k: v for k, v in loc.items() if k != "months"}
    merged.update(user_cfg)
    MONTHS.update({ar_norm(k): v for k, v in loc.get("months", {}).items()})
    return merged

def resolve_config(user_cfg, headers, rows=None, detect_names=None):
    cfg = json.loads(json.dumps(DEFAULTS)); cfg.update({k: v for k, v in (user_cfg or {}).items() if k in DEFAULTS})
    declared = dict(cfg["columns"]); auto = {}
    for j, h in enumerate(headers):
        if h in declared:
            if declared[h] not in TYPES: raise SystemExit(f"Unknown type '{declared[h]}' for column '{h}'. Allowed: {sorted(TYPES)}")
            continue
        t = detect_type((detect_names or {}).get(h, h))
        if rows:
            vals = [r["raw"][j] for r in rows]; nonempty = [v for v in vals if norm_text(v) != ""]
            if t in ("name", "category") and nonempty and sum(1 for v in nonempty if isinstance(v, (int, float)) and not isinstance(v, bool)) / len(nonempty) >= 0.9: t = "number"
            elif t == "text": t = sniff_type(vals, cfg)
        auto[h] = t
    cfg["columns"] = {**auto, **{h: t for h, t in declared.items() if h in headers}}
    cfg["_auto_detected"] = sorted(auto)
    missing = [h for h in declared if h not in headers]
    if missing: raise SystemExit(f"Config names columns not in the file: {missing}. File has: {headers}")
    return cfg

def list_sheets(path):
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True); names = list(wb.sheetnames); wb.close(); return names
    return []

def _read_grid(path, sheet):
    p = Path(path); hint = None; sname = None; merges = []; fmiss = 0
    if p.suffix.lower() in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(p, data_only=True)
        if sheet and sheet not in wb.sheetnames: raise SystemExit(f"Sheet '{sheet}' not found. Sheets: {wb.sheetnames}")
        ws = wb[sheet] if sheet else wb.active; sname = ws.title
        grid = [list(r) for r in ws.iter_rows(values_only=True)]
        merges = [(m.min_row, m.min_col, m.max_row, m.max_col) for m in ws.merged_cells.ranges]
        best = 0
        for t in getattr(ws, "tables", {}).values():           # real Excel "Table" objects
            try: c1, r1, c2, r2 = openpyxl.utils.cell.range_boundaries(t.ref)
            except Exception: continue
            if (r2 - r1 + 1) * (c2 - c1 + 1) > best: best = (r2 - r1 + 1) * (c2 - c1 + 1); hint = (r1, c1, r2, c2)
        try:                                                    # formulas whose result was never saved read as empty
            wsf = openpyxl.load_workbook(p, data_only=False)[ws.title]
            for rf, rv in zip(wsf.iter_rows(), ws.iter_rows()):
                for cf, cv in zip(rf, rv):
                    if isinstance(cf.value, str) and cf.value.startswith("=") and cv.value is None: fmiss += 1
        except Exception: pass
    elif p.suffix.lower() in (".csv", ".tsv", ".txt"):
        raw = p.read_bytes()
        try: text = raw.decode("utf-8-sig")
        except UnicodeDecodeError: text = raw.decode("cp1252")
        dialect = csv.excel_tab if p.suffix.lower() == ".tsv" else (csv.Sniffer().sniff(text[:4096], delimiters=",;\t") if text.strip() else csv.excel)
        grid = [list(r) for r in csv.reader(text.splitlines(), dialect)]
    else: raise SystemExit("Unsupported input type (use .xlsx or .csv)")
    return grid, hint, sname, merges, fmiss

def _is_texty(c):
    if isinstance(c, (datetime, date)): return False
    v = norm_text(c)
    return v != "" and not re.fullmatch(r"[\d\s.,:/\-+%$€]+", v)

def detect_header_row(grid):
    """1-based number of the (lowest) header row: skips title rows, blank rows and notes above the table."""
    scan = grid[:60]; counts = [sum(1 for c in r if norm_text(c) != "") for r in scan]
    maxn = max(counts) if counts else 0
    if maxn == 0: raise SystemExit("The file/sheet is empty")
    need = 1 if maxn == 1 else max(2, 0.6 * maxn)
    for i, r in enumerate(scan):
        n = counts[i]
        if n < need: continue
        if sum(1 for c in r if _is_texty(c)) < 0.6 * n: continue
        nxt = counts[i + 1:i + 4]
        if not nxt or any(x >= max(1, 0.5 * n) for x in nxt): return i + 1
    return counts.index(maxn) + 1

TOTAL_RX = re.compile(ar_norm(r"^(total|totale|totaux|sous[- ]?total|grand total|subtotal|somme|sum|المجموع|مجموع|الاجمالي|اجمالي|الجملة|جملة)\b"), re.I)

def _is_total_row(vals):
    for v in vals:
        if norm_text(v) != "": return bool(TOTAL_RX.match(ar_norm(norm_text(v).casefold())))
    return False

def read_table_info(path, sheet=None, header_row=None, levels=None, skip_totals=True):
    """Find the table. header_row = the LOWEST header row (the one with the sub-headings). levels=None detects group rows above it."""
    grid, hint, sname, merges, fmiss = _read_grid(path, sheet)
    if not any(norm_text(c) != "" for r in grid for c in r): raise SystemExit("The file/sheet is empty")
    lo = hi = r_hi = None; how = "manual" if header_row else "auto"
    if header_row is None:
        if hint: header_row, lo, hi, r_hi = hint[0], hint[1] - 1, hint[3] - 1, hint[2]; how = "Excel table"
        else: header_row = detect_header_row(grid)
    if len(grid) < header_row: raise SystemExit("No header row found")
    width = max(len(r) for r in grid); grid = [list(r) + [None] * (width - len(r)) for r in grid]
    def filled(ri):                                              # merged header cells repeat their label across the merge
        row = list(grid[ri])
        for r1, c1, r2, c2, in merges:
            if r1 - 1 <= ri <= r2 - 1 and norm_text(grid[r1 - 1][c1 - 1]) != "":
                for c in range(c1 - 1, min(c2, width)):
                    if norm_text(row[c]) == "": row[c] = grid[r1 - 1][c1 - 1]
        return row
    hdr = filled(header_row - 1); data = grid[header_row:r_hi]
    lo = 0 if lo is None else lo; hi = width - 1 if hi is None else hi
    keep = [c for c in range(lo, hi + 1) if norm_text(hdr[c]) != "" or any(norm_text(r[c]) != "" for r in data)]
    if not keep: raise SystemExit("No data found below the header row")
    group_rows = []; want = 2 if levels is None else max(0, levels - 1); cand = header_row - 1
    while cand >= 1 and len(group_rows) < want:
        row = filled(cand - 1); labs = [(c, norm_text(row[c])) for c in keep if norm_text(row[c]) != ""]
        ok = bool(labs) if levels is not None else (len({l.casefold() for _, l in labs}) >= 2 and all(_is_texty(row[c]) for c, _ in labs))
        if not ok: break
        group_rows.append(cand); cand -= 1
    group_rows.reverse()
    subs = {c: norm_text(hdr[c]) for c in keep}; glabels = []
    for gr in group_rows:                                        # group labels: merged range, else nearest label above a sub-heading
        row = filled(gr - 1); lab = {c: norm_text(row[c]) for c in keep}
        cands = [c for c in keep if lab[c] and subs[c]]; res = {}
        for c in keep:
            if lab[c]: res[c] = lab[c]
            elif subs[c] and cands: res[c] = lab[min(cands, key=lambda k: (abs(k - c), k > c))]
            else: res[c] = ""
        glabels.append(res)
    headers, dnames, seen = [], {}, Counter()
    for c in keep:
        parts = []
        for x in [g[c] for g in glabels] + [subs[c]]:
            if x and (not parts or parts[-1].casefold() != x.casefold()): parts.append(x)
        h = " - ".join(parts) or f"Column {get_column_letter(c + 1)}"; seen[h] += 1
        h = h if seen[h] == 1 else f"{h} ({seen[h]})"; headers.append(h)
        dnames[h] = subs[c] or (parts[-1] if parts else h)
    rows, skipped, blanks = [], [], 0
    for off, r in enumerate(data, start=header_row + 1):
        vals = [r[c] for c in keep]
        if all(norm_text(x) == "" for x in vals): blanks += 1; continue
        if skip_totals and _is_total_row(vals): skipped.append({"src": off, "raw": vals, "reason": "total row"}); continue
        rows.append({"src": off, "raw": vals})
    top = group_rows[0] if group_rows else header_row
    note = (f"sheet '{sname}', " if sname else "") + (f"header at row {header_row}" if not group_rows else f"header at rows {top}-{header_row} ({len(group_rows) + 1} levels)") \
        + f" ({how}), columns {get_column_letter(keep[0] + 1)}-{get_column_letter(keep[-1] + 1)}"
    warnings = []
    if fmiss: warnings.append(f"{fmiss} formula cell(s) have no saved result and read as empty: open the file in Excel, save it, then try again")
    if skipped: warnings.append(f"{len(skipped)} total row(s) skipped, not cleaned (see sheet 'Skipped rows'): source row(s) " + ", ".join(str(k['src']) for k in skipped[:8]))
    return {"headers": headers, "detect_names": dnames, "rows": rows, "skipped": skipped, "blanks": blanks, "header_row": header_row,
            "levels": len(group_rows) + 1, "sheet": sname, "note": note, "how": how, "warnings": warnings}

def read_table(path, sheet=None, header_row=None):
    i = read_table_info(path, sheet, header_row); return i["headers"], i["rows"], i["blanks"]

def clean_table(headers, rows, cfg):
    n = len(rows); cells = {}; log = []
    for j, h in enumerate(headers):
        t = cfg["columns"][h]; raw = [r["raw"][j] for r in rows]
        if t == "skip": res = [(x, "ok", "") for x in raw]
        elif t in ("name", "category", "text"):
            res = clean_text_column(raw, cfg, h, t)
        elif t == "id":
            res = [((norm_text(x) or None), "missing" if norm_text(x) == "" else "ok", "ID missing" if norm_text(x) == "" else "") for x in raw]
            res = [(int(v) if isinstance(v, str) and v.isdigit() and not isinstance(x, str) else (x if isinstance(x, (int, float)) else v), s, m) for (v, s, m), x in zip(res, raw)]
        else: res = [PARSERS[t](x, cfg) for x in raw]
        for i, (v, s, m) in enumerate(res):
            cells[(i, j)] = (v, s, m)
            o = raw[i]
            if s == "ok" and not same(o, v) and not (o is None and v in (None, "")):
                log.append({"row": rows[i]["src"], "column": h, "before": o, "after": v, "status": "auto-fixed", "note": ""})
            elif s in ("verify", "fail", "missing"):
                st = {"verify": "VERIFY", "fail": "UNRESOLVED", "missing": "MISSING"}[s]
                log.append({"row": rows[i]["src"], "column": h, "before": o, "after": v, "status": st, "note": m})
    # exact duplicates (all columns identical after cleaning) -> removed, first kept
    keyrow, removed, keep = {}, [], []
    for i in range(n):
        k = tuple(repr(cells[(i, j)][0]) for j in range(len(headers)))
        if k in keyrow: removed.append((i, keyrow[k]))
        else: keyrow[k] = i; keep.append(i)
    for i, first in removed:
        log.append({"row": rows[i]["src"], "column": "(entire row)", "before": "duplicate", "after": "removed",
                    "status": "auto-fixed", "note": f"exact duplicate of source row {rows[first]['src']}"})
    # reused IDs / possible duplicates with different IDs (flag only)
    idcols = [j for j, h in enumerate(headers) if cfg["columns"][h] == "id"]
    if idcols:
        j0 = idcols[0]; byid = defaultdict(list)
        for i in keep:
            if cells[(i, j0)][0] not in (None, ""): byid[cells[(i, j0)][0]].append(i)
        for ident, idx in byid.items():
            if len(idx) > 1:
                for i in idx[1:]:
                    log.append({"row": rows[i]["src"], "column": headers[j0], "before": ident, "after": ident, "status": "VERIFY",
                                "note": f"ID also used by source row {rows[idx[0]]['src']} with different content"})
        others = [j for j, h in enumerate(headers) if cfg["columns"][h] in ("name", "category", "qty", "price", "date", "phone", "email")]
        seen = {}
        for i in keep:
            k = tuple(repr(cells[(i, j)][0]) for j in others)
            if sum(1 for j in others if cells[(i, j)][0] not in (None, "")) >= 3:
                if k in seen:
                    log.append({"row": rows[i]["src"], "column": "(entire row)", "before": "", "after": "kept", "status": "VERIFY",
                                "note": f"same content as source row {rows[seen[k]]['src']} but different ID; possible duplicate, not removed"})
                else: seen[k] = i
    return cells, keep, removed, log

def totals(headers, cfg, cells, keep):
    q = [j for j, h in enumerate(headers) if cfg["columns"][h] == "qty"]; p = [j for j, h in enumerate(headers) if cfg["columns"][h] == "price"]
    if not (len(q) == 1 and len(p) == 1): return None
    out = {}
    for i in keep:
        a, b = cells[(i, q[0])][0], cells[(i, p[0])][0]
        out[i] = round(a * b, 2) if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) else None
    return out

FILL_V = PatternFill("solid", fgColor="FFF2A8"); FILL_F = PatternFill("solid", fgColor="F8B88B")
HDR = PatternFill("solid", fgColor="1F3A5F")

def write_workbook(path, headers, rows, cfg, cells, keep, removed, log, blanks, src_name, table_note="", skipped=None, warnings=None):
    wb = openpyxl.Workbook(); tot = totals(headers, cfg, cells, keep)
    c = Counter(l["status"] for l in log); review = [l for l in log if l["status"] != "auto-fixed"]
    ws = wb.active; ws.title = "Summary"
    lines = [("Cleaning summary", ""), ("Source file", src_name), ("Table found", table_note), ("Rows read", len(rows) + 0), ("Blank rows skipped", blanks),
             ("Exact duplicate rows removed", len(removed)), ("Rows in cleaned output", len(keep)),
             ("Cells auto-fixed", c["auto-fixed"] - len(removed)), ("Cells to VERIFY (assumption made)", c["VERIFY"]),
             ("Cells UNRESOLVED (original kept)", c["UNRESOLVED"]), ("Cells MISSING", c["MISSING"]), ("", ""),
             ("Column types used", "")] + [(h, cfg["columns"][h] + (" (auto-detected)" if h in cfg["_auto_detected"] else "")) for h in headers]
    for w_ in (warnings or []): lines += [("Note", w_)]
    lines += [("", ""), ("Legend", "Yellow = changed with an assumption, please verify. Orange = could not be cleaned or missing; original kept.")]
    for r in lines: ws.append(list(r))
    ws["A1"].font = Font(bold=True, size=14); ws.column_dimensions["A"].width = 42; ws.column_dimensions["B"].width = 70
    wc = wb.create_sheet("Clean"); outh = headers + (["Total (qty x price)"] if tot is not None else []) + ["Source row"]
    wc.append(outh)
    for k, i in enumerate(keep, start=2):
        vals = [cells[(i, j)][0] for j in range(len(headers))]
        if tot is not None: vals.append(tot[i])
        wc.append(vals + [rows[i]["src"]])
        for j in range(len(headers)):
            v, s, _ = cells[(i, j)]; cell = wc.cell(row=k, column=j + 1)
            if isinstance(v, date): cell.number_format = "yyyy-mm-dd"
            if s == "verify": cell.fill = FILL_V
            elif s in ("fail", "missing"): cell.fill = FILL_F
    wr = wb.create_sheet("Review"); wr.append(["Source row", "Column", "Original", "Cleaned / kept", "Status", "What to check"])
    for l in review: wr.append([l["row"], l["column"], str(l["before"]) if l["before"] is not None else "", str(l["after"]) if l["after"] is not None else "", l["status"], l["note"]])
    wl = wb.create_sheet("Change log"); wl.append(["Source row", "Column", "Before", "After", "Status", "Note"])
    for l in sorted(log, key=lambda x: (x["row"], x["column"])): wl.append([l["row"], l["column"], str(l["before"]) if l["before"] is not None else "", str(l["after"]) if l["after"] is not None else "", l["status"], l["note"]])
    if skipped:
        wsk = wb.create_sheet("Skipped rows"); wsk.append(["Source row", "Reason"] + headers)
        for k in skipped: wsk.append([k["src"], k["reason"]] + [o.date() if isinstance(o, datetime) else o for o in k["raw"]])
    wd = wb.create_sheet("Before-After (20 rows)"); h2 = ["Source row"]
    for h in headers: h2 += [f"{h} (before)", f"{h} (after)"]
    wd.append(h2)
    for i in keep[:20]:
        row = [rows[i]["src"]]
        for j in range(len(headers)):
            o = rows[i]["raw"][j]; v = cells[(i, j)][0]
            row += [o.date() if isinstance(o, datetime) else o, v]
        wd.append(row)
    for w in [x for x in (wc, wr, wl, wd, wb["Skipped rows"] if "Skipped rows" in wb.sheetnames else None) if x is not None]:
        for cell in w[1]: cell.font = Font(bold=True, color="FFFFFF"); cell.fill = HDR; cell.alignment = Alignment(wrap_text=True, vertical="top")
        w.freeze_panes = "A2"
        for idx, col in enumerate(w.columns, start=1):
            width = max((len(str(x.value)) if x.value is not None else 0) for x in list(col)[:60])
            w.column_dimensions[get_column_letter(idx)].width = min(max(10, width + 2), 48)
    wb.save(path)
    return lines

def report_md(src_name, headers, cfg, rows, keep, removed, log, blanks, cells, table_note="", warnings=None):
    c = Counter(l["status"] for l in log); fixed = c["auto-fixed"] - len(removed)
    L = [f"# Cleaning report - {src_name}", "", f"- Table found: {table_note}",
         f"- Rows read: {len(rows)} | exact duplicates removed: {len(removed)} | rows in cleaned file: {len(keep)}",
         f"- Auto-fixed cells: {fixed} | to VERIFY: {c['VERIFY']} | UNRESOLVED: {c['UNRESOLVED']} | MISSING: {c['MISSING']}", "",
         *[f"- NOTE: {w_}" for w_ in (warnings or [])], "", "## What was changed automatically", ""]
    per = Counter(l["column"] for l in log if l["status"] == "auto-fixed" and l["column"] != "(entire row)")
    L += [f"- {h}: {n} cells (whitespace, casing, format)" for h, n in per.items()] or ["- none"]
    L += ["", "## Needs a human decision", ""]
    rev = [l for l in log if l["status"] != "auto-fixed"]
    groups = defaultdict(list)
    for l in rev:
        shown = "(blank)" if l["before"] in (None, "") else str(l["before"])
        note = re.sub(r"source row \d+", "source row N", l["note"]) if l["column"] == "(entire row)" or "also used" in l["note"] else l["note"]
        groups[(l["column"], l["status"], shown if l["status"] != "MISSING" else "(blank)", note)].append(l["row"])
    for (col, st, shown, note), rws in sorted(groups.items(), key=lambda kv: -len(kv[1]))[:20]:
        rs = ", ".join(map(str, rws[:8])) + (f" +{len(rws) - 8} more" if len(rws) > 8 else "")
        L.append(f"- {col} `{shown}` x{len(rws)} -> {st}: {note} (rows {rs})")
    if not rev: L.append("- nothing flagged")
    if len(groups) > 20: L.append(f"- ... {len(groups) - 20} more groups (see Review sheet)")
    L += ["", "Nothing flagged above was silently guessed: yellow/orange cells in the Clean sheet mark each one. The original file was not modified."]
    return "\n".join(L)

def run(path, cfg_path=None, out_dir=None, sheet=None, header_row=None, also_csv=False, write_config=None, levels=None, skip_totals=True, locale="dz"):
    src = Path(path); info = read_table_info(src, sheet, header_row, levels, skip_totals); headers, rows, blanks = info["headers"], info["rows"], info["blanks"]
    user_cfg = json.loads(Path(cfg_path).read_text(encoding="utf-8")) if cfg_path else {}
    cfg = resolve_config(merge_locale(user_cfg, locale), headers, rows, info["detect_names"])
    if write_config:
        Path(write_config).write_text(json.dumps({k: v for k, v in cfg.items() if not k.startswith("_")}, indent=2, ensure_ascii=False), encoding="utf-8")
    out = Path(out_dir) if out_dir else src.parent; out.mkdir(parents=True, exist_ok=True)
    target = out / f"{src.stem}_cleaned.xlsx"
    if target.resolve() == src.resolve(): raise SystemExit("Refusing to overwrite the input file")
    cells, keep, removed, log = clean_table(headers, rows, cfg)
    write_workbook(target, headers, rows, cfg, cells, keep, removed, log, blanks, src.name, info["note"], info["skipped"], info["warnings"])
    rep = out / f"{src.stem}_report.md"; rep.write_text(report_md(src.name, headers, cfg, rows, keep, removed, log, blanks, cells, info["note"], info["warnings"]), encoding="utf-8")
    if also_csv:
        with open(out / f"{src.stem}_cleaned.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f); w.writerow(headers)
            for i in keep: w.writerow([("" if cells[(i, j)][0] is None else cells[(i, j)][0]) for j in range(len(headers))])
    return {"xlsx": target, "report": rep, "rows_in": len(rows), "rows_out": len(keep), "removed": len(removed), "log": log, "cells": cells, "keep": keep, "headers": headers, "cfg": cfg, "info": info}

def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input"); ap.add_argument("--config"); ap.add_argument("--out-dir"); ap.add_argument("--sheet")
    ap.add_argument("--header-row", type=int, default=None, help="row number of the LOWEST header row"); ap.add_argument("--levels", type=int, default=None, help="header levels (1 = single header row); default: detect"); ap.add_argument("--keep-total-rows", action="store_true"); ap.add_argument("--also-csv", action="store_true"); ap.add_argument("--write-config")
    ap.add_argument("--locale", default="dz", choices=sorted(LOCALES), help="country pack for phone/price/date rules (default: dz)")
    a = ap.parse_args(argv)
    r = run(a.input, a.config, a.out_dir, a.sheet, a.header_row, a.also_csv, a.write_config, a.levels, not a.keep_total_rows, a.locale)
    c = Counter(l["status"] for l in r["log"])
    print(f"{r['rows_in']} rows in -> {r['rows_out']} out ({r['removed']} duplicates removed). VERIFY={c['VERIFY']} UNRESOLVED={c['UNRESOLVED']} MISSING={c['MISSING']}")
    print("Wrote:", r["xlsx"], "and", r["report"])

if __name__ == "__main__":
    main()
