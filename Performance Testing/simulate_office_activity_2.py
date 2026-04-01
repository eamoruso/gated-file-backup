#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
simulate_office_activity.py  (Enhanced + Faculty-skew profile)

Windows-only. Automates Word & Excel to:
  - Create a new document/workbook with randomized, profile-aware content size
  - Optionally perform "burst" edits (multiple saves) before rename/copy
  - Save into a base folder (default: Documents, or --save-dir)
  - Randomize operation order: rename→copy OR copy→rename (per file, per iter)
  - Randomize destination subfolder depth; optionally move instead of copy
  - Optional small timing jitter and short-lived file locks to trigger retries
  - Optionally skip Word or Excel per iteration (with profile skew)
Repeat steps multiple times and record timings + plan to CSV and a .log.

Examples:
  python simulate_office_activity.py --iterations 15 --profile faculty --seed 42
  python simulate_office_activity.py --iterations 20 --profile faculty --visible 0
  python simulate_office_activity.py --iterations 25 --profile faculty --edit-bursts-prob 0.5

Dependencies:
  pip install pywin32
"""

import os
import sys
import time
import csv
import shutil
import argparse
from datetime import datetime, timezone
import random
from pathlib import Path
import threading
import string

# -------------------- Retry utilities --------------------
def retry(op, *, tries=6, base_delay=0.10, factor=1.8, jitter=0.05, exceptions=(Exception,)):
    """Retry helper with exponential backoff."""
    delay = base_delay
    last = None
    for _ in range(tries):
        try:
            return op()
        except exceptions as e:
            last = e
            time.sleep(delay + random.uniform(0, jitter))
            delay *= factor
    raise last

def wait_for_file(path: str, attempts=20, pause=0.05):
    """Wait until file exists and has non-zero size."""
    p = Path(path)
    for _ in range(attempts):
        if p.exists():
            try:
                if p.stat().st_size > 0:
                    return True
            except Exception:
                pass
        time.sleep(pause)
    return False

# -------------------- Windows & COM checks --------------------
if sys.platform != "win32":
    print("This script runs on Windows only (requires COM automation).")
    sys.exit(1)

try:
    import win32com.client as win32
except ImportError:
    print("Missing dependency: pywin32. Install with: pip install pywin32")
    sys.exit(1)

# Optional APIs for file attributes/locking
try:
    import win32con
    import win32file
    import win32api
except Exception:
    win32con = None
    win32file = None
    win32api = None

# COM FileFormat constants
WD_FORMAT_DOCUMENT_DEFAULT = 16  # .docx
XL_OPENXML_WORKBOOK = 51         # .xlsx

# -------------------- Helpers --------------------
def get_documents_folder() -> str:
    return os.path.join(os.path.expanduser("~"), "Documents")

def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

def now_tag() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S_%f")

def open_word(visible: bool):
    word = win32.Dispatch("Word.Application")
    word.Visible = bool(visible)
    try:
        word.DisplayAlerts = 0
    except Exception:
        pass
    return word

def open_excel(visible: bool):
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = bool(visible)
    try:
        excel.DisplayAlerts = False
    except Exception:
        pass
    return excel

def close_word(word):
    try:
        word.Quit()
    except Exception:
        pass

def close_excel(excel):
    try:
        excel.Quit()
    except Exception:
        pass

def excel_col_letter(n: int) -> str:
    """1 -> A, 26 -> Z, 27 -> AA ..."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

WORDS = [
    "analysis","system","backup","journal","dataset","method","random","stochastic","monitor",
    "event","trigger","snapshot","policy","agent","delta","latency","interval","incremental",
    "full","diff","consistency","version","recovery","retry","flush","lock","shadow","volume",
    "index","checkpoint","target","source","throughput","write","read","I/O","workload",
]

FACULTY_TOPICS = [
    "syllabus","rubric","lecture","slides","assignment","quiz","midterm","final","gradebook",
    "draft","revision","manuscript","submission","IRB","section","semester","office-hours",
    "capstone","lab","project","attendance","feedback","peer-review","outline","bibliography",
]

def rand_text(paragraphs: int, words_min: int, words_max: int) -> str:
    parts = []
    for _ in range(paragraphs):
        n = random.randint(words_min, words_max)
        tokens = [random.choice(WORDS + FACULTY_TOPICS) for __ in range(n)]
        parts.append(" ".join(tokens).capitalize() + ".")
    return "\n\n".join(parts)

def random_sleep_ms(ms_min: int, ms_max: int):
    if ms_max <= 0:
        return
    delay = random.uniform(ms_min, ms_max) / 1000.0
    time.sleep(delay)

def set_random_attrs(path: str, make_hidden: bool, make_readonly: bool):
    try:
        if make_readonly:
            os.chmod(path, 0o444)
    except Exception:
        pass
    try:
        if make_hidden and win32api and win32con:
            attrs = win32api.GetFileAttributes(path)
            win32api.SetFileAttributes(path, attrs | win32con.FILE_ATTRIBUTE_HIDDEN)
    except Exception:
        pass

def clear_readonly(path: str):
    try:
        os.chmod(path, 0o666)
    except Exception:
        pass

def inject_exclusive_lock_async(path: str, lock_ms: int):
    """
    Start a thread that opens the file with no sharing (exclusive), holds for lock_ms, then closes.
    Returns an Event that's set when the lock is active, plus the thread.
    """
    started = threading.Event()
    def _locker():
        handle = None
        try:
            if win32file and win32con:
                handle = win32file.CreateFile(
                    path,
                    win32con.GENERIC_READ,
                    0,  # no sharing
                    None,
                    win32con.OPEN_EXISTING,
                    win32con.FILE_ATTRIBUTE_NORMAL,
                    None
                )
                started.set()
                time.sleep(max(lock_ms, 1)/1000.0)
        except Exception:
            started.set()
            time.sleep(max(lock_ms, 1)/1000.0)
        finally:
            try:
                if handle:
                    win32file.CloseHandle(handle)
            except Exception:
                pass
    t = threading.Thread(target=_locker, daemon=True)
    t.start()
    started.wait(timeout=0.2)
    return started, t

def tri_int(low: int, high: int, mode: int) -> int:
    """Integer from triangular distribution."""
    return max(low, min(high, int(round(random.triangular(low, high, mode)))))

def random_rename_variant_generic(base_name: str) -> str:
    variants = [
        lambda s: s + "_renamed",
        lambda s: "Final_" + s,
        lambda s: s.replace("_", " ") + " (v2)",
        lambda s: s + "_" + datetime.now().strftime("%H%M%S"),
        lambda s: s.title(),
        lambda s: "résumé_" + s,
        lambda s: s + "_Δ",
    ]
    fn = random.choice(variants)
    return fn(base_name)

def random_rename_variant_faculty(base_name: str) -> str:
    seasons = ["Spring", "Summer", "Fall"]
    season = random.choice(seasons)
    year = datetime.now().year
    lecture_no = f"{random.randint(1, 15):02d}"
    variants = [
        lambda s: f"Syllabus_{season}{year}",
        lambda s: f"Lecture_{lecture_no}_{season}{year}",
        lambda s: f"{s}_Draft_v{random.randint(2,5)}",
        lambda s: f"{s}_Final",
        lambda s: f"Assignment_{random.randint(1,6)}_{season}{year}",
        lambda s: f"{s}_Revision_{random.randint(1,3)}",
        lambda s: f"Gradebook_{season}{year}",
    ]
    return random.choice(variants)(base_name)

# -------------------- Main --------------------
def main():
    parser = argparse.ArgumentParser(description="Simulate Word/Excel activity (randomized/profiled) and log timings.")
    parser.add_argument("--save-dir", type=str, default="",
                        help="Folder to save initial Word/Excel files (default: Documents). "
                             "Use a non-synced path like C:\\Temp\\OfficeSim to avoid sync locks.")
    parser.add_argument("--iterations", type=int, default=5, help="Number of repetitions (default: 5)")
    parser.add_argument("--delay", type=float, default=0.0, help="Delay (seconds) between iterations (default: 0)")
    parser.add_argument("--visible", type=int, default=0, choices=[0,1], help="Show Word/Excel UI? 0/1 (default: 0)")
    parser.add_argument("--relaunch-apps", type=int, default=0, choices=[0,1],
                        help="Relaunch Word/Excel each iteration? 0/1 (default: 0 - reuse instance)")
    parser.add_argument("--subfolder-name", type=str, default="OfficeSim_Subfolder",
                        help="Subfolder under save-dir/Documents to copy/move files into (default: OfficeSim_Subfolder)")
    parser.add_argument("--results-file", type=str, default="",
                        help="CSV results path (default: Documents/office_sim_results_<timestamp>.csv)")
    parser.add_argument("--log-file", type=str, default="",
                        help="Readable log path (default: same base name with .log)")

    # ---- Profile & randomness controls ----
    parser.add_argument("--profile", type=str, choices=["none","faculty"], default="none",
                        help="Skew randomness for a workload profile (default: none).")
    parser.add_argument("--seed", type=int, default=None, help="Random seed (int). If omitted, uses OS entropy.")
    parser.add_argument("--skip-word-p", type=float, default=None, help="Probability to skip Word in an iteration (0-1).")
    parser.add_argument("--skip-excel-p", type=float, default=None, help="Probability to skip Excel in an iteration (0-1).")
    parser.add_argument("--jitter-ms", type=int, nargs=2, default=None,
                        help="Min/Max milliseconds of random sleep inserted around ops.")
    parser.add_argument("--lock-prob", type=float, default=None,
                        help="Probability to inject a short exclusive file lock before rename/copy (0-1).")
    parser.add_argument("--lock-ms", type=int, nargs=2, default=None,
                        help="Min/Max ms for the transient lock.")
    parser.add_argument("--copy-move-prob", type=float, default=None,
                        help="Probability to MOVE (shutil.move) instead of COPY for the dest step (0-1).")
    parser.add_argument("--dest-depth", type=int, nargs=2, default=None,
                        help="Min/Max extra nested subfolder depth under --subfolder-name.")
    parser.add_argument("--word-paras", type=int, nargs=2, default=None,
                        help="Min/Max number of paragraphs to write into Word.")
    parser.add_argument("--word-para-len", type=int, nargs=2, default=None,
                        help="Min/Max words per paragraph in Word.")
    parser.add_argument("--excel-rows", type=int, nargs=2, default=None,
                        help="Min/Max Excel data rows.")
    parser.add_argument("--excel-cols", type=int, nargs=2, default=None,
                        help="Min/Max Excel data columns.")
    parser.add_argument("--randomize-ops-order", type=int, default=None, choices=[0,1],
                        help="If 1, randomly choose rename→copy or copy→rename per file.")
    parser.add_argument("--attr-hidden-p", type=float, default=None,
                        help="Probability to set Hidden attribute on saved file before ops (0-1).")
    parser.add_argument("--attr-readonly-p", type=float, default=None,
                        help="Probability to set Read-only attribute on saved file before ops (0-1).")
    # Burst edit controls
    parser.add_argument("--edit-bursts-prob", type=float, default=None,
                        help="Probability to perform burst edits (multiple saves) before rename/copy (0-1).")
    parser.add_argument("--word-bursts", type=int, nargs=2, default=None,
                        help="Min/Max extra saves in a Word burst (default depends on profile).")
    parser.add_argument("--excel-bursts", type=int, nargs=2, default=None,
                        help="Min/Max extra saves in an Excel burst (default depends on profile).")

    args = parser.parse_args()

    # ---- Apply profile defaults if not explicitly set ----
    def def_if_none(val, default):
        return default if val is None else val

    if args.profile == "faculty":
        # Skewed defaults
        args.skip_word_p      = def_if_none(args.skip_word_p, 0.05)
        args.skip_excel_p     = def_if_none(args.skip_excel_p, 0.35)
        args.jitter_ms        = def_if_none(args.jitter_ms, [5, 40])
        args.lock_prob        = def_if_none(args.lock_prob, 0.10)
        args.lock_ms          = def_if_none(args.lock_ms, [40, 120])
        args.copy_move_prob   = def_if_none(args.copy_move_prob, 0.30)
        args.dest_depth       = def_if_none(args.dest_depth, [1, 3])
        # Content sizes: choose broad ranges; we'll refine with triangular below
        args.word_paras       = def_if_none(args.word_paras, [1, 10])
        args.word_para_len    = def_if_none(args.word_para_len, [10, 45])
        args.excel_rows       = def_if_none(args.excel_rows, [15, 250])
        args.excel_cols       = def_if_none(args.excel_cols, [5, 24])
        args.randomize_ops_order = def_if_none(args.randomize_ops_order, 1)
        args.attr_hidden_p    = def_if_none(args.attr_hidden_p, 0.02)
        args.attr_readonly_p  = def_if_none(args.attr_readonly_p, 0.10)
        # Bursts
        args.edit_bursts_prob = def_if_none(args.edit_bursts_prob, 0.45)
        args.word_bursts      = def_if_none(args.word_bursts, [1, 3])
        args.excel_bursts     = def_if_none(args.excel_bursts, [1, 2])
    else:
        # Legacy-like neutral defaults if user didn't set them
        args.skip_word_p      = def_if_none(args.skip_word_p, 0.0)
        args.skip_excel_p     = def_if_none(args.skip_excel_p, 0.0)
        args.jitter_ms        = def_if_none(args.jitter_ms, [0, 25])
        args.lock_prob        = def_if_none(args.lock_prob, 0.0)
        args.lock_ms          = def_if_none(args.lock_ms, [40, 120])
        args.copy_move_prob   = def_if_none(args.copy_move_prob, 0.0)
        args.dest_depth       = def_if_none(args.dest_depth, [0, 2])
        args.word_paras       = def_if_none(args.word_paras, [1, 3])
        args.word_para_len    = def_if_none(args.word_para_len, [8, 20])
        args.excel_rows       = def_if_none(args.excel_rows, [5, 50])
        args.excel_cols       = def_if_none(args.excel_cols, [3, 12])
        args.randomize_ops_order = def_if_none(args.randomize_ops_order, 1)
        args.attr_hidden_p    = def_if_none(args.attr_hidden_p, 0.0)
        args.attr_readonly_p  = def_if_none(args.attr_readonly_p, 0.0)
        args.edit_bursts_prob = def_if_none(args.edit_bursts_prob, 0.0)
        args.word_bursts      = def_if_none(args.word_bursts, [1, 1])
        args.excel_bursts     = def_if_none(args.excel_bursts, [1, 1])

    # Seed RNG
    if args.seed is not None:
        random.seed(args.seed)

    documents = get_documents_folder()
    base_dir = os.path.abspath(args.save_dir) if args.save_dir else documents
    ensure_dir(base_dir)

    subfolder_root = os.path.join(base_dir, args.subfolder_name)
    ensure_dir(subfolder_root)

    tag = now_tag()
    results_csv = os.path.join(documents, f"office_sim_results_{tag}.csv") if not args.results_file else os.path.abspath(args.results_file)
    ensure_dir(os.path.dirname(results_csv))
    log_path = os.path.splitext(results_csv)[0] + ".log" if not args.log_file else os.path.abspath(args.log_file)
    ensure_dir(os.path.dirname(log_path))

    # Prepare CSV headers (extended)
    headers = [
        "iteration", "utc_iso", "profile",
        "plan_skip_word", "plan_skip_excel",
        "plan_ops_order_word", "plan_ops_order_excel",
        "plan_move_word", "plan_move_excel",
        "plan_dest_depth_word", "plan_dest_depth_excel",
        "plan_lock_word_ms", "plan_lock_excel_ms",
        "plan_attr_hidden_word", "plan_attr_readonly_word",
        "plan_attr_hidden_excel", "plan_attr_readonly_excel",
        "plan_word_paras", "plan_word_words_total",
        "plan_excel_rows", "plan_excel_cols",
        "word_burst_edits", "word_burst_s",
        "excel_burst_edits", "excel_burst_s",
        "word_create_save_s", "word_rename_s", "word_copy_or_move_s",
        "excel_create_save_s", "excel_rename_s", "excel_copy_or_move_s",
        "total_s",
        "word_saved_path", "word_renamed_path", "word_dest_path",
        "excel_saved_path", "excel_renamed_path", "excel_dest_path",
        "error"
    ]

    print(f"Base folder  : {base_dir}")
    print(f"Copy root    : {subfolder_root}")
    print(f"Results CSV  : {results_csv}")
    print(f"Readable log : {log_path}\n")

    with open(results_csv, mode="w", newline="", encoding="utf-8") as f_csv, \
         open(log_path,    mode="w", encoding="utf-8") as f_log:

        writer = csv.DictWriter(f_csv, fieldnames=headers)
        writer.writeheader()

        # Reuse vs relaunch apps
        word_app = None
        excel_app = None
        if not args.relaunch_apps:
            word_app = open_word(args.visible)
            excel_app = open_excel(args.visible)

        for i in range(1, args.iterations + 1):
            row = {h: "" for h in headers}
            row["iteration"] = i
            row["utc_iso"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
            row["profile"] = args.profile

            # Per-iteration plan
            do_word = random.random() >= args.skip_word_p
            do_excel = random.random() >= args.skip_excel_p
            ops_order_word = random.choice(["rename-then-copy", "copy-then-rename"]) if args.randomize_ops_order else "rename-then-copy"
            ops_order_excel = random.choice(["rename-then-copy", "copy-then-rename"]) if args.randomize_ops_order else "rename-then-copy"
            move_word = random.random() < args.copy_move_prob
            move_excel = random.random() < args.copy_move_prob
            dest_depth_word = random.randint(*args.dest_depth)
            dest_depth_excel = random.randint(*args.dest_depth)
            lock_ms_word = (random.randint(*args.lock_ms) if random.random() < args.lock_prob else 0)
            lock_ms_excel = (random.randint(*args.lock_ms) if random.random() < args.lock_prob else 0)
            attr_hidden_word = random.random() < args.attr_hidden_p
            attr_readonly_word = random.random() < args.attr_readonly_p
            attr_hidden_excel = random.random() < args.attr_hidden_p
            attr_readonly_excel = random.random() < args.attr_readonly_p
            jitter_min, jitter_max = args.jitter_ms
            # Faculty skew: sample sizes via triangular distributions
            if args.profile == "faculty":
                paras = tri_int(args.word_paras[0], args.word_paras[1], mode=3)
                # word length: mode favors shorter paragraphs but with tail for longer
                base_len = tri_int(args.word_para_len[0], args.word_para_len[1], mode=18)
                wlen_min = max(6, int(base_len * 0.7))
                wlen_max = min(60, int(base_len * 1.3))
                x_rows = tri_int(args.excel_rows[0], args.excel_rows[1], mode=60)
                x_cols = tri_int(args.excel_cols[0], args.excel_cols[1], mode=10)
            else:
                paras = random.randint(*args.word_paras)
                wlen_min, wlen_max = args.word_para_len
                x_rows = random.randint(*args.excel_rows)
                x_cols = random.randint(*args.excel_cols)

            row["plan_skip_word"] = int(not do_word)
            row["plan_skip_excel"] = int(not do_excel)
            row["plan_ops_order_word"] = ops_order_word
            row["plan_ops_order_excel"] = ops_order_excel
            row["plan_move_word"] = int(move_word)
            row["plan_move_excel"] = int(move_excel)
            row["plan_dest_depth_word"] = dest_depth_word
            row["plan_dest_depth_excel"] = dest_depth_excel
            row["plan_lock_word_ms"] = lock_ms_word
            row["plan_lock_excel_ms"] = lock_ms_excel
            row["plan_attr_hidden_word"] = int(attr_hidden_word)
            row["plan_attr_readonly_word"] = int(attr_readonly_word)
            row["plan_attr_hidden_excel"] = int(attr_hidden_excel)
            row["plan_attr_readonly_excel"] = int(attr_readonly_excel)

            # Open/close apps per-iteration if requested
            if args.relaunch_apps:
                word_app = open_word(args.visible)
                excel_app = open_excel(args.visible)

            # Unique base names
            tstamp = now_tag()
            word_base = f"SimWord_{i}_{tstamp}"
            excel_base = f"SimExcel_{i}_{tstamp}"

            # Rename variants (faculty vs generic)
            if args.profile == "faculty":
                word_renamed_base = random_rename_variant_faculty(word_base)
                excel_renamed_base = random_rename_variant_faculty(excel_base)
            else:
                word_renamed_base = random_rename_variant_generic(word_base)
                excel_renamed_base = random_rename_variant_generic(excel_base)

            word_saved = os.path.join(base_dir, word_base + ".docx")
            word_renamed = os.path.join(base_dir, word_renamed_base + ".docx")
            excel_saved = os.path.join(base_dir, excel_base + ".xlsx")
            excel_renamed = os.path.join(base_dir, excel_renamed_base + ".xlsx")

            # Destination folders
            def random_subdirs(base_dir: str, depth: int) -> str:
                p = Path(base_dir)
                for _ in range(depth):
                    # faculty-ish subfolders
                    seg = random.choice([
                        f"Course_{random.choice(['CEN','COP','CIS','CNT'])}{random.randint(2000,5999)}",
                        f"{random.choice(['Spring','Summer','Fall'])}{datetime.now().year}",
                        f"Week_{random.randint(1,16)}",
                        "".join(random.choices(string.ascii_letters + string.digits, k=random.randint(4,10))),
                    ])
                    p = p / seg
                ensure_dir(str(p))
                return str(p)

            dest_dir_word = random_subdirs(subfolder_root, dest_depth_word)
            dest_dir_excel = random_subdirs(subfolder_root, dest_depth_excel)
            word_dest_path_pre  = os.path.join(dest_dir_word, os.path.basename(word_saved))
            word_dest_path_post = os.path.join(dest_dir_word, os.path.basename(word_renamed))
            excel_dest_path_pre  = os.path.join(dest_dir_excel, os.path.basename(excel_saved))
            excel_dest_path_post = os.path.join(dest_dir_excel, os.path.basename(excel_renamed))

            t_total_start = time.perf_counter()
            error_text = ""

            try:
                # ---------- WORD ----------
                if do_word:
                    try:
                        random_sleep_ms(jitter_min, jitter_max)
                        # Make SaveAs synchronous
                        try:
                            word_app.Options.BackgroundSave = False
                            word_app.Options.SaveInterval = 0
                        except Exception:
                            pass

                        t0 = time.perf_counter()
                        doc = word_app.Documents.Add()
                        text = rand_text(paras, wlen_min, wlen_max)
                        words_total = sum(len(p.split()) for p in text.splitlines() if p.strip())
                        row["plan_word_paras"] = paras
                        row["plan_word_words_total"] = words_total
                        row["word_burst_edits"] = 0
                        row["word_burst_s"] = 0.0

                        word_app.Selection.TypeText(f"Run {i} at {tstamp}\n\n{text}")

                        # Initial save
                        def do_word_save():
                            doc.SaveAs2(word_saved, FileFormat=WD_FORMAT_DOCUMENT_DEFAULT, AddToRecentFiles=False)
                            if not wait_for_file(word_saved):
                                raise IOError(f"Saved file not visible yet: {word_saved}")
                        retry(do_word_save, tries=6, base_delay=0.08)

                        # Faculty burst edits (small appends + Save)
                        burst_s = 0.0
                        if random.random() < args.edit_bursts_prob:
                            n_bursts = random.randint(*args.word_bursts)
                            for _ in range(n_bursts):
                                bt0 = time.perf_counter()
                                # Append small chunk
                                word_app.Selection.TypeText("\n" + rand_text(1, 5, 12))
                                doc.Save()
                                wait_for_file(word_saved)
                                random_sleep_ms(5, 20)
                                burst_s += (time.perf_counter() - bt0)
                            row["word_burst_edits"] = n_bursts
                            row["word_burst_s"] = round(burst_s, 6)

                        doc.Close(SaveChanges=False)
                        row["word_create_save_s"] = round(time.perf_counter() - t0, 6)
                        row["word_saved_path"] = word_saved

                        # Attributes
                        set_random_attrs(word_saved, attr_hidden_word, attr_readonly_word)

                        # Lock
                        if lock_ms_word > 0:
                            _, _ = inject_exclusive_lock_async(word_saved, lock_ms_word)
                            random_sleep_ms(5, 15)

                        # Order
                        if ops_order_word == "rename-then-copy":
                            # RENAME
                            random_sleep_ms(jitter_min, jitter_max)
                            t1 = time.perf_counter()
                            if attr_readonly_word:
                                clear_readonly(word_saved)
                            retry(lambda: os.replace(word_saved, word_renamed), tries=6, base_delay=0.05)
                            row["word_rename_s"] = round(time.perf_counter() - t1, 6)
                            row["word_renamed_path"] = word_renamed

                            # COPY or MOVE
                            random_sleep_ms(jitter_min, jitter_max)
                            t2 = time.perf_counter()
                            if move_word:
                                retry(lambda: shutil.move(word_renamed, os.path.join(dest_dir_word, os.path.basename(word_renamed))),
                                      tries=6, base_delay=0.05)
                                row["word_dest_path"] = os.path.join(dest_dir_word, os.path.basename(word_renamed))
                            else:
                                retry(lambda: shutil.copy2(word_renamed, word_dest_path_post), tries=6, base_delay=0.05)
                                row["word_dest_path"] = word_dest_path_post
                            row["word_copy_or_move_s"] = round(time.perf_counter() - t2, 6)

                        else:  # copy-then-rename
                            random_sleep_ms(jitter_min, jitter_max)
                            t2 = time.perf_counter()
                            if move_word:
                                retry(lambda: shutil.move(word_saved, os.path.join(dest_dir_word, os.path.basename(word_saved))),
                                      tries=6, base_delay=0.05)
                                row["word_dest_path"] = os.path.join(dest_dir_word, os.path.basename(word_saved))
                                dest_current = row["word_dest_path"]
                                dest_renamed = os.path.join(dest_dir_word, os.path.basename(word_renamed))
                                random_sleep_ms(jitter_min, jitter_max)
                                t1 = time.perf_counter()
                                retry(lambda: os.replace(dest_current, dest_renamed), tries=6, base_delay=0.05)
                                row["word_rename_s"] = round(time.perf_counter() - t1, 6)
                                row["word_renamed_path"] = dest_renamed
                                row["word_dest_path"] = dest_renamed
                            else:
                                retry(lambda: shutil.copy2(word_saved, word_dest_path_pre), tries=6, base_delay=0.05)
                                row["word_dest_path"] = word_dest_path_pre
                                row["word_copy_or_move_s"] = round(time.perf_counter() - t2, 6)
                                random_sleep_ms(jitter_min, jitter_max)
                                t1 = time.perf_counter()
                                if attr_readonly_word:
                                    clear_readonly(word_saved)
                                retry(lambda: os.replace(word_saved, word_renamed), tries=6, base_delay=0.05)
                                row["word_rename_s"] = round(time.perf_counter() - t1, 6)
                                row["word_renamed_path"] = word_renamed

                    except Exception as e:
                        error_text = f"{type(e).__name__} (Word): {e}"

                else:
                    row["plan_word_paras"] = 0
                    row["plan_word_words_total"] = 0
                    row["word_burst_edits"] = 0
                    row["word_burst_s"] = 0.0

                # ---------- EXCEL ----------
                if do_excel:
                    try:
                        random_sleep_ms(jitter_min, jitter_max)
                        t0 = time.perf_counter()

                        wb = excel_app.Workbooks.Add()
                        ws = wb.Worksheets(1)

                        rows = x_rows
                        cols = x_cols
                        row["plan_excel_rows"] = rows
                        row["plan_excel_cols"] = cols

                        ws.Range("A1").Value = f"Run {i}"
                        ws.Range("A2").Value = f"Timestamp {tstamp}"

                        # Fill starting A4
                        start_row = 4
                        start_col = 1
                        end_row = start_row + rows - 1
                        end_col = start_col + cols - 1
                        top_left = excel_col_letter(start_col) + str(start_row)
                        bottom_right = excel_col_letter(end_col) + str(end_row)

                        matrix = []
                        for r in range(rows):
                            rowv = []
                            for c in range(cols):
                                if (r + c) % 3 == 0:
                                    rowv.append(r * c + random.randint(0, 1000))
                                else:
                                    rowv.append(random.choice(WORDS + FACULTY_TOPICS))
                            matrix.append(rowv)
                        ws.Range(f"{top_left}:{bottom_right}").Value = tuple(tuple(x) for x in matrix)

                        def do_excel_save():
                            wb.SaveAs(excel_saved, FileFormat=XL_OPENXML_WORKBOOK)
                            if not wait_for_file(excel_saved):
                                raise IOError(f"Saved file not visible yet: {excel_saved}")
                        retry(do_excel_save, tries=6, base_delay=0.06)

                        # Bursty small updates before close (simulate grading pass)
                        e_burst_s = 0.0
                        row["excel_burst_edits"] = 0
                        row["excel_burst_s"] = 0.0
                        if random.random() < args.edit_bursts_prob:
                            n_bursts = random.randint(*args.excel_bursts)
                            for _ in range(n_bursts):
                                bt0 = time.perf_counter()
                                rr = random.randint(start_row, end_row)
                                cc = random.randint(start_col, end_col)
                                cell = excel_col_letter(cc) + str(rr)
                                ws.Range(cell).Value = random.choice(["OK","Review","Late","✓", random.randint(50, 100)])
                                wb.Save()
                                wait_for_file(excel_saved)
                                random_sleep_ms(5, 20)
                                e_burst_s += (time.perf_counter() - bt0)
                            row["excel_burst_edits"] = n_bursts
                            row["excel_burst_s"] = round(e_burst_s, 6)

                        wb.Close(SaveChanges=False)

                        row["excel_create_save_s"] = round(time.perf_counter() - t0, 6)
                        row["excel_saved_path"] = excel_saved

                        set_random_attrs(excel_saved, attr_hidden_excel, attr_readonly_excel)

                        if lock_ms_excel > 0:
                            _, _ = inject_exclusive_lock_async(excel_saved, lock_ms_excel)
                            random_sleep_ms(5, 15)

                        if ops_order_excel == "rename-then-copy":
                            random_sleep_ms(jitter_min, jitter_max)
                            t1 = time.perf_counter()
                            if attr_readonly_excel:
                                clear_readonly(excel_saved)
                            retry(lambda: os.replace(excel_saved, excel_renamed), tries=6, base_delay=0.05)
                            row["excel_rename_s"] = round(time.perf_counter() - t1, 6)
                            row["excel_renamed_path"] = excel_renamed

                            random_sleep_ms(jitter_min, jitter_max)
                            t2 = time.perf_counter()
                            if move_excel:
                                retry(lambda: shutil.move(excel_renamed, os.path.join(dest_dir_excel, os.path.basename(excel_renamed))),
                                      tries=6, base_delay=0.05)
                                row["excel_dest_path"] = os.path.join(dest_dir_excel, os.path.basename(excel_renamed))
                            else:
                                retry(lambda: shutil.copy2(excel_renamed, excel_dest_path_post), tries=6, base_delay=0.05)
                                row["excel_dest_path"] = excel_dest_path_post
                            row["excel_copy_or_move_s"] = round(time.perf_counter() - t2, 6)
                        else:
                            random_sleep_ms(jitter_min, jitter_max)
                            t2 = time.perf_counter()
                            if move_excel:
                                retry(lambda: shutil.move(excel_saved, os.path.join(dest_dir_excel, os.path.basename(excel_saved))),
                                      tries=6, base_delay=0.05)
                                row["excel_dest_path"] = os.path.join(dest_dir_excel, os.path.basename(excel_saved))
                                dest_current = row["excel_dest_path"]
                                dest_renamed = os.path.join(dest_dir_excel, os.path.basename(excel_renamed))
                                random_sleep_ms(jitter_min, jitter_max)
                                t1 = time.perf_counter()
                                retry(lambda: os.replace(dest_current, dest_renamed), tries=6, base_delay=0.05)
                                row["excel_rename_s"] = round(time.perf_counter() - t1, 6)
                                row["excel_renamed_path"] = dest_renamed
                                row["excel_dest_path"] = dest_renamed
                            else:
                                retry(lambda: shutil.copy2(excel_saved, excel_dest_path_pre), tries=6, base_delay=0.05)
                                row["excel_dest_path"] = excel_dest_path_pre
                                row["excel_copy_or_move_s"] = round(time.perf_counter() - t2, 6)
                                random_sleep_ms(jitter_min, jitter_max)
                                t1 = time.perf_counter()
                                if attr_readonly_excel:
                                    clear_readonly(excel_saved)
                                retry(lambda: os.replace(excel_saved, excel_renamed), tries=6, base_delay=0.05)
                                row["excel_rename_s"] = round(time.perf_counter() - t1, 6)
                                row["excel_renamed_path"] = excel_renamed

                    except Exception as e:
                        et = f"{type(e).__name__} (Excel): {e}"
                        error_text = (error_text + " | " + et).strip(" |")

                # ---------- Totals & Log ----------
                row["total_s"] = round(time.perf_counter() - t_total_start, 6)
                row["error"] = error_text
                writer.writerow(row)

                f_log.write(
                    f"[Iter {i}] total={row['total_s']}s profile={args.profile} "
                    f"PLAN(word_skip={row['plan_skip_word']}, excel_skip={row['plan_skip_excel']}, "
                    f"word_order={row['plan_ops_order_word']}, excel_order={row['plan_ops_order_excel']}, "
                    f"move_w={row['plan_move_word']}, move_x={row['plan_move_excel']}, "
                    f"lock_w={row['plan_lock_word_ms']}ms, lock_x={row['plan_lock_excel_ms']}ms, "
                    f"burst_w={row['word_burst_edits']}, burst_x={row['excel_burst_edits']}) "
                    f"WORD(create+save={row.get('word_create_save_s','')}, "
                    f"rename={row.get('word_rename_s','')}, copy/move={row.get('word_copy_or_move_s','')}) "
                    f"EXCEL(create+save={row.get('excel_create_save_s','')}, "
                    f"rename={row.get('excel_rename_s','')}, copy/move={row.get('excel_copy_or_move_s','')}) "
                    f"ERROR={error_text}\n"
                )
                f_log.flush()

            finally:
                if args.relaunch_apps:
                    close_word(word_app)
                    close_excel(excel_app)
                if args.delay > 0 and i < args.iterations:
                    time.sleep(args.delay)

        if not args.relaunch_apps:
            close_word(word_app)
            close_excel(excel_app)

    print("\nDone.")
    print(f"CSV saved to: {results_csv}")
    print(f"Log saved to : {log_path}")

if __name__ == "__main__":
    main()
