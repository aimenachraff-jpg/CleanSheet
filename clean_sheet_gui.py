#!/usr/bin/env python3
"""CleanSheet - a simple window around clean_sheet.py (no command line needed).
Pick a file, check the detected table and column types, click Clean. Your original file is never modified."""
import json, os, queue, subprocess, sys, threading, traceback
from collections import Counter
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import clean_sheet as cs
from clean_sheet_locales import DESCRIPTIONS

DATE_CHOICES = {"Day first (31/12/2026)": "dmy", "Month first (12/31/2026)": "mdy"}
LOCALE_CHOICES = list(DESCRIPTIONS)  # locale codes, shown by description
LOCALE_DEFAULT = "dz"

# Optional drag-and-drop: tkinterdnd2 is NOT a hard dependency. When it is
# installed, files can be dropped on the window; otherwise the file picker is
# the only way to choose a file (no behavior change, no install required).
try:
    from tkinterdnd2 import TkinterDnD  # type: ignore
    HAS_DND = True
except Exception:  # pragma: no cover - tkinterdnd2 not installed in this env
    TkinterDnD = None
    HAS_DND = False

if HAS_DND:
    from tkinterdnd2 import DND_FILES  # type: ignore
else:
    DND_FILES = None
TYPE_LABELS = {  # friendly names shown in the window -> internal type
    "Name": "name", "Category (city, product...)": "category", "Quantity": "qty", "Price / amount": "price", "Number (score, grade...)": "number",
    "Date": "date", "Phone": "phone", "Email": "email", "ID / reference": "id", "Other text": "text", "Leave untouched": "skip"}
LABEL_OF = {v: k for k, v in TYPE_LABELS.items()}


def inspect_file(path, sheet=None, header_row=None, levels=None, skip_totals=True):
    """Find the table in the file. Returns headers, guessed types, example values, header row, sheets."""
    info = cs.read_table_info(path, sheet, header_row, levels, skip_totals)
    headers, rows = info["headers"], info["rows"]
    cfg = cs.resolve_config({}, headers, rows, info["detect_names"])
    samples = {}
    for j, h in enumerate(headers):
        vals = []
        for r in rows:
            v = cs.norm_text(r["raw"][j])
            if v: vals.append(v[:28])
            if len(vals) == 3: break
        samples[h] = "  |  ".join(vals)
    preview = [[cs.norm_text(r["raw"][j])[:24] for j in range(len(headers))] for r in rows[:3]]
    return {"headers": headers, "types": {h: cfg["columns"][h] for h in headers}, "samples": samples, "header_row": info["header_row"],
            "how": info["how"], "sheet": info["sheet"], "sheets": cs.list_sheets(path), "note": info["note"], "rows": len(rows),
            "levels": info["levels"], "warnings": info["warnings"], "skipped": len(info["skipped"]), "preview": preview}


def run_job(path, columns, date_order="dmy", out_dir=None, also_csv=False, sheet=None, header_row=None, levels=None, skip_totals=True, locale=LOCALE_DEFAULT):
    """Run the cleaner with an explicit column mapping. Writes the config used next to the output."""
    path = Path(path)
    out = Path(out_dir) if out_dir else path.parent / "cleaned_output"
    out.mkdir(parents=True, exist_ok=True)
    cfg_path = out / f"{path.stem}_used_config.json"
    cfg_path.write_text(json.dumps({"columns": columns, "date_order": date_order, "locale": locale}, indent=2, ensure_ascii=False), encoding="utf-8")
    locale_code = locale.split(" - ", 1)[0] if " - " in locale else locale
    r = cs.run(path, cfg_path, out, sheet, header_row, also_csv, None, levels, skip_totals, locale=locale_code)
    c = Counter(l["status"] for l in r["log"])
    return {"xlsx": r["xlsx"], "report": r["report"], "out": out, "rows_in": r["rows_in"], "rows_out": r["rows_out"],
            "removed": r["removed"], "verify": c["VERIFY"], "unresolved": c["UNRESOLVED"], "missing": c["MISSING"], "note": r["info"]["note"], "skipped": len(r["info"]["skipped"]), "warnings": r["info"]["warnings"]}


def open_path(p):
    p = str(p)
    if sys.platform.startswith("win"): os.startfile(p)
    elif sys.platform == "darwin": subprocess.Popen(["open", p])
    else: subprocess.Popen(["xdg-open", p])


class App((TkinterDnD.Tk if HAS_DND else tk.Tk)):
    def __init__(self):
        super().__init__()
        self.title("CleanSheet - Excel/CSV cleaner"); self.geometry("900x720"); self.minsize(780, 620)
        if HAS_DND:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)
        st = ttk.Style(self)
        for theme in ("vista", "clam", "xpnative", "winnative"):
            if theme in st.theme_names(): st.theme_use(theme); break
        base = ("Segoe UI", 10)
        ACCENT = "#1F3A5F"  # dark blue, matches the workbook header color
        st.configure(".", font=base)
        st.configure("H.TLabel", font=("Segoe UI", 11, "bold"), foreground=ACCENT)
        st.configure("Hint.TLabel", foreground="#555")
        st.configure("Big.TButton", font=("Segoe UI", 10, "bold"), padding=7)
        st.configure("Prev.TLabel", foreground="#333")
        self.path = None; self.sheet = None; self.header_override = None; self.levels_override = None; self.type_vars = {}; self.result = None
        self._queue = queue.Queue(); self._polling = False
        root = ttk.Frame(self, padding=14); root.pack(fill="both", expand=True)

        ttk.Label(root, text="1. Choose your file", style="H.TLabel").grid(row=0, column=0, sticky="w")
        top = ttk.Frame(root); top.grid(row=1, column=0, sticky="ew", pady=(4, 10)); top.columnconfigure(1, weight=1)
        ttk.Button(top, text="Choose Excel / CSV file...", command=self.pick_file).grid(row=0, column=0)
        self.file_lbl = ttk.Label(top, text="No file chosen", style="Hint.TLabel"); self.file_lbl.grid(row=0, column=1, sticky="w", padx=10)
        self.table_row = ttk.Frame(top); self.table_row.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Label(self.table_row, text="Sheet:").pack(side="left")
        self.sheet_var = tk.StringVar(); self.sheet_cb = ttk.Combobox(self.table_row, textvariable=self.sheet_var, state="readonly", width=22)
        self.sheet_cb.pack(side="left", padx=(4, 16)); self.sheet_cb.bind("<<ComboboxSelected>>", lambda e: self.reload(sheet_changed=True))
        ttk.Label(self.table_row, text="Header row:").pack(side="left")
        self.hdr_var = tk.StringVar(); self.hdr_sp = ttk.Spinbox(self.table_row, from_=1, to=500, width=5, textvariable=self.hdr_var, command=self.hdr_changed)
        self.hdr_sp.pack(side="left", padx=4); self.hdr_sp.bind("<Return>", lambda e: self.hdr_changed()); self.hdr_sp.bind("<FocusOut>", lambda e: self.hdr_changed(only_if_changed=True))
        ttk.Label(self.table_row, text="  Header levels:").pack(side="left")
        self.lev_var = tk.StringVar(); self.lev_sp = ttk.Spinbox(self.table_row, from_=1, to=3, width=3, textvariable=self.lev_var, command=self.lev_changed)
        self.lev_sp.pack(side="left", padx=4); self.lev_sp.bind("<Return>", lambda e: self.lev_changed())
        self.skip_var = tk.BooleanVar(value=True); ttk.Checkbutton(self.table_row, text="Skip total rows", variable=self.skip_var, command=self.reload).pack(side="left", padx=10)
        ttk.Button(self.table_row, text="Auto-detect again", command=self.auto_again).pack(side="left", padx=4)
        self.detect_lbl = ttk.Label(root, text="", style="Hint.TLabel", wraplength=840, justify="left"); self.detect_lbl.grid(row=2, column=0, sticky="w")

        ttk.Label(root, text="2. Check what each column contains", style="H.TLabel").grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Label(root, text="The window guessed from the column names and the values. Change a type if it is wrong; 'Leave untouched' skips a column.", style="Hint.TLabel", wraplength=840, justify="left").grid(row=4, column=0, sticky="w")
        box = ttk.Frame(root, relief="solid", borderwidth=1); box.grid(row=5, column=0, sticky="nsew", pady=6); root.rowconfigure(5, weight=1); root.columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(box, highlightthickness=0, background="white"); sb = ttk.Scrollbar(box, orient="vertical", command=self.canvas.yview)
        self.cols_frame = ttk.Frame(self.canvas); self.cols_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.cols_frame, anchor="nw"); self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-e.delta / 120), "units"))

        opt = ttk.Frame(root); opt.grid(row=6, column=0, sticky="ew", pady=(4, 2))
        ttk.Label(opt, text="Ambiguous dates like 03/04/2026:").pack(side="left")
        self.date_var = tk.StringVar(value=list(DATE_CHOICES)[0])
        ttk.Combobox(opt, textvariable=self.date_var, values=list(DATE_CHOICES), state="readonly", width=26).pack(side="left", padx=6)
        ttk.Label(opt, text="  Country/locale:").pack(side="left", padx=(14, 0))
        self.locale_var = tk.StringVar(value=LOCALE_DEFAULT)
        ttk.Combobox(opt, textvariable=self.locale_var, values=[f"{k} - {DESCRIPTIONS[k]}" for k in LOCALE_CHOICES],
                     state="readonly", width=34).pack(side="left", padx=6)
        self.csv_var = tk.BooleanVar(value=False); ttk.Checkbutton(opt, text="Also save a CSV copy", variable=self.csv_var).pack(side="left", padx=14)
        out = ttk.Frame(root); out.grid(row=7, column=0, sticky="ew", pady=4); out.columnconfigure(1, weight=1)
        ttk.Label(out, text="Save results in:").grid(row=0, column=0)
        self.out_var = tk.StringVar(); ttk.Entry(out, textvariable=self.out_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(out, text="Browse...", command=self.pick_out).grid(row=0, column=2)

        ttk.Label(root, text="3. Clean", style="H.TLabel").grid(row=8, column=0, sticky="w", pady=(8, 2))
        act = ttk.Frame(root); act.grid(row=9, column=0, sticky="w")
        self.run_btn = ttk.Button(act, text="Clean file", style="Big.TButton", command=self.clean, state="disabled"); self.run_btn.pack(side="left")
        self.open_file_btn = ttk.Button(act, text="Open cleaned file", command=lambda: open_path(self.result["xlsx"]), state="disabled"); self.open_file_btn.pack(side="left", padx=(14, 6))
        self.open_dir_btn = ttk.Button(act, text="Open output folder", command=lambda: open_path(self.result["out"]), state="disabled"); self.open_dir_btn.pack(side="left")
        self.progress = ttk.Progressbar(act, mode="indeterminate", length=140)
        self.progress.pack(side="left", padx=(16, 4))
        self.status_lbl = ttk.Label(act, text="", style="Hint.TLabel"); self.status_lbl.pack(side="left")
        self.result_lbl = ttk.Label(root, text="", wraplength=840, justify="left"); self.result_lbl.grid(row=10, column=0, sticky="w", pady=6)
        ttk.Label(root, text="Your original file is never changed. In the 'Clean' sheet: yellow = changed with an assumption, orange = unreadable or missing. Check the 'Review' sheet by hand before sending anything.",
                  style="Hint.TLabel", wraplength=840, justify="left").grid(row=11, column=0, sticky="w")
        self.table_row.grid_remove()

    # ---- loading -------------------------------------------------------------------------
    def pick_file(self):
        p = filedialog.askopenfilename(title="Choose a file", filetypes=[("Excel or CSV", "*.xlsx *.xlsm *.csv *.tsv"), ("All files", "*.*")])
        if p: self.load_file(p)

    def load_file(self, p):
        self.path = Path(p); self.sheet = None; self.header_override = None; self.levels_override = None
        self.file_lbl.config(text=str(self.path)); self.out_var.set(str(self.path.parent / "cleaned_output"))
        self.reload()

    def reload(self, sheet_changed=False):
        if not self.path: return
        if sheet_changed: self.sheet = self.sheet_var.get() or None; self.header_override = None; self.levels_override = None
        try: res = inspect_file(self.path, self.sheet, self.header_override, self.levels_override, self.skip_var.get())
        except SystemExit as e: messagebox.showerror("Cannot read this table", str(e)); return
        except Exception as e: messagebox.showerror("Cannot read file", f"{type(e).__name__}: {e}"); return
        if res["sheets"]:
            self.sheet_cb.config(values=res["sheets"]); self.sheet_var.set(res["sheet"]); self.sheet = res["sheet"]
        self.table_row.grid()
        self.hdr_var.set(str(res["header_row"])); self.last_hdr = str(res["header_row"]); self.lev_var.set(str(res["levels"]))
        how = {"auto": "found automatically", "manual": "set by you", "Excel table": "taken from the Excel table in the file"}[res["how"]]
        lev = "" if res["levels"] == 1 else f" The header has {res['levels']} levels (group title + sub-headings), shown joined like 'Group - Sub'."
        extra = "".join("\n" + w for w in res["warnings"])
        self.detect_lbl.config(text=f"Table: lowest header row {res['header_row']} ({how}); {res['rows']} data rows.{lev} If the names look wrong, change 'Header row' or 'Header levels' above.{extra}")
        for w in self.cols_frame.winfo_children(): w.destroy()
        for c, t in enumerate(("Column in your file", "What it contains", "Examples from your data")):
            ttk.Label(self.cols_frame, text=t, font=("Segoe UI", 9, "bold")).grid(row=0, column=c, sticky="w", padx=8, pady=(6, 4))
        self.type_vars = {}
        for i, h in enumerate(res["headers"], start=1):
            ttk.Label(self.cols_frame, text=h, width=28, anchor="w").grid(row=i, column=0, sticky="w", padx=8, pady=2)
            v = tk.StringVar(value=LABEL_OF[res["types"][h]]); self.type_vars[h] = v
            ttk.Combobox(self.cols_frame, textvariable=v, values=list(TYPE_LABELS), state="readonly", width=24).grid(row=i, column=1, padx=8, pady=2)
            ttk.Label(self.cols_frame, text=res["samples"][h], foreground="#555").grid(row=i, column=2, sticky="w", padx=8)
        # ---- preview of the first data rows -------------------------------------------
        prev_row = len(res["headers"]) + 1
        prev = ttk.Frame(self.cols_frame); prev.grid(row=prev_row, column=0, columnspan=3, sticky="ew", pady=(10, 4))
        prev.columnconfigure(1, weight=1)
        ttk.Label(prev, text="Preview — first rows of your data (read-only):", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w")
        prev_text = tk.Text(prev, height=4, wrap="none", font=("Consolas", 9))
        prev_text.grid(row=0, column=1, sticky="ew")
        hsb = ttk.Scrollbar(prev, orient="horizontal", command=prev_text.xview); hsb.grid(row=1, column=1, sticky="ew")
        prev_text.configure(xscrollcommand=hsb.set)
        lines = ["\t".join(res["headers"])]
        for row in res.get("preview", []):
            lines.append("\t".join(row))
        prev_text.insert("1.0", "\n".join(lines))
        prev_text.configure(state="disabled")
        # ------------------------------------------------------------------------------
        self.run_btn.config(state="normal"); self.result_lbl.config(text="")
        self.open_file_btn.config(state="disabled"); self.open_dir_btn.config(state="disabled")

    def hdr_changed(self, only_if_changed=False):
        v = self.hdr_var.get().strip()
        if only_if_changed and v == getattr(self, "last_hdr", v): return
        if not v.isdigit() or int(v) < 1: messagebox.showerror("Header row", "Enter a row number (1, 2, 3...)."); return
        self.header_override = int(v); self.reload()

    def lev_changed(self):
        v = self.lev_var.get().strip()
        if not v.isdigit() or not 1 <= int(v) <= 3: messagebox.showerror("Header levels", "Enter 1, 2 or 3."); return
        self.levels_override = int(v); self.reload()

    def auto_again(self):
        self.header_override = None; self.levels_override = None; self.reload()

    def pick_out(self):
        d = filedialog.askdirectory(title="Choose output folder")
        if d: self.out_var.set(d)

    def _on_drop(self, event):
        """Handle a file dropped onto the window (tkinterdnd2, when installed)."""
        if not HAS_DND: return
        p = (event.data or "").strip().strip("{}")
        if p and Path(p).is_file() and p.lower().endswith((".xlsx", ".xlsm", ".csv", ".tsv", ".txt")):
            self.load_file(p)

    # ---- running -------------------------------------------------------------------------
    def clean(self):
        if not self.path or self._polling: return
        cols = {h: TYPE_LABELS[v.get()] for h, v in self.type_vars.items()}
        kwargs = dict(
            date_order=DATE_CHOICES[self.date_var.get()],
            out_dir=self.out_var.get() or None,
            also_csv=self.csv_var.get(),
            sheet=self.sheet, header_row=self.header_override, levels=self.levels_override,
            skip_totals=self.skip_var.get(), locale=self.locale_var.get(),
        )
        self.run_btn.config(state="disabled"); self.progress.start(14)
        self.status_lbl.config(text="Cleaning… (original file is not modified)")
        self.result_lbl.config(text="")
        self._queue = queue.Queue()
        threading.Thread(target=self._worker, args=(self.path, cols, kwargs), daemon=True).start()
        self._polling = True; self.after(200, self._poll)

    def _worker(self, path, cols, kwargs):
        """Runs on a background thread; never touches Tk widgets."""
        try:
            result = run_job(path, cols, **kwargs)
            self._queue.put(("ok", result))
        except SystemExit as e:
            self._queue.put(("system_exit", str(e)))
        except Exception as e:  # pragma: no cover - exercised by GUI smoke tests only
            self._queue.put(("error", f"{type(e).__name__}: {e}"))

    def _poll(self):
        """Runs on the Tkinter thread; delivers worker results to the UI."""
        try:
            kind, payload = self._queue.get_nowait()
        except queue.Empty:
            if self._polling:
                self.after(200, self._poll)
            return
        self._polling = False
        self.progress.stop(); self.status_lbl.config(text="")
        self.run_btn.config(state="normal")
        if kind == "ok":
            r = payload; self.result = r
            self.result_lbl.config(text=f"Done: {r['rows_in']} rows in, {r['rows_out']} out ({r['removed']} exact duplicates removed).\n"
                                        f"To check by hand: {r['verify']} yellow (assumption), {r['unresolved']} orange (unreadable), {r['missing']} missing."
                                        + (f" {r['skipped']} total row(s) were skipped (sheet 'Skipped rows').\n" if r["skipped"] else "\n") +
                                        f"Saved in: {r['out']}")
            self.open_file_btn.config(state="normal"); self.open_dir_btn.config(state="normal")
        elif kind == "system_exit":
            messagebox.showerror("Cleaning stopped", payload)
        else:
            messagebox.showerror("Unexpected error", payload + "\n\n" + traceback.format_exc(limit=3))


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
