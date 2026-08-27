#!/usr/bin/env python3
"""
End-to-end DOMIN8 reporting pipeline.

    STEP 1  fetch    pull Uniware reports  -> reports/input/uniware/
    STEP 2  check    verify the manual-drop folders have current files
    STEP 3  build    reconcile everything  -> reports/output/
    STEP 4  verify   tie the output back to the raw input files
    STEP 5  merch    build the Stock vs Sales workbook

Usage
-----
    # first time: create the folder tree, then stop
    python run_pipeline.py --init

    # DEFAULT: build from the files already in reports/input/uniware/
    python run_pipeline.py

    # pull fresh data from Uniware first
    set UNIWARE_USER=domin8@abstrabit.com
    set UNIWARE_PASS=...
    python run_pipeline.py --fetch

    # other useful flags
    python run_pipeline.py --days 90            # wider Uniware window
    python run_pipeline.py --start 2026-04-01 --end 2026-06-30
    python run_pipeline.py --skip-verify
    python run_pipeline.py --status             # what's on disk right now

Exit codes: 0 = clean, 1 = ran but with warnings, 2 = failed.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timedelta
from pathlib import Path

import pipeline_config as C

WARN: list[str] = []


def say(msg=""):
    print(msg, flush=True)


def head(title):
    say()
    say("=" * 74)
    say(title)
    say("=" * 74)


def warn(msg):
    WARN.append(msg)
    say(f"  !! {msg}")


def newest(folder: Path):
    """(path, mtime) of the most recently modified data file in a folder."""
    files = [p for p in folder.glob("*")
             if p.is_file() and p.suffix.lower() in (".csv", ".xlsx", ".xls", ".xlsm")
             and not p.name.startswith(("~$", "."))]
    if not files:
        return None, None
    p = max(files, key=lambda f: f.stat().st_mtime)
    return p, datetime.fromtimestamp(p.stat().st_mtime)


def data_files(folder: Path):
    return sorted(p for p in folder.glob("*")
                  if p.is_file() and p.suffix.lower() in (".csv", ".xlsx", ".xls", ".xlsm")
                  and not p.name.startswith(("~$", ".")))



def date_slices(days: int, cap: int):
    """Split a window into <=cap-day slices, newest first.

    Uniware's export presets top out at 90 days; a longer single request can
    come back truncated with no error. Slicing is the documented workaround.
    """
    end = datetime.now().date()
    out = []
    remaining = days
    while remaining > 0:
        span = min(cap, remaining)
        start = end - timedelta(days=span - 1)
        out.append((start.isoformat(), end.isoformat()))
        end = start - timedelta(days=1)
        remaining -= span
    return out


def upsert_po_history(slice_files: list[Path]) -> tuple[int, int]:
    """Merge PO slices into the persistent history file.

    A PO line changes as it is received, so the newest version of each
    (PO Code, Item SkuCode) wins. Returns (rows_before, rows_after).
    """
    import pandas as pd

    frames = []
    if C.PO_HISTORY_FILE.exists():
        try:
            frames.append(pd.read_csv(C.PO_HISTORY_FILE, low_memory=False))
        except Exception as exc:
            warn(f"could not read {C.PO_HISTORY_FILE.name}: {exc}")
    before = len(frames[0]) if frames else 0

    # oldest slice first, so later (newer) rows overwrite on drop_duplicates
    for f in sorted(slice_files, key=lambda x: x.stat().st_mtime):
        try:
            frames.append(pd.read_csv(f, low_memory=False))
        except Exception as exc:
            warn(f"could not read PO slice {f.name}: {exc}")
    if not frames:
        return before, before

    hist = pd.concat(frames, ignore_index=True)
    key = [k for k in C.PO_HISTORY_KEY if k in hist.columns]
    if key:
        hist = hist.drop_duplicates(subset=key, keep="last")
    C.PO_HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    hist.to_csv(C.PO_HISTORY_FILE, index=False)
    return before, len(hist)

# ---------------------------------------------------------------------------
# STEP 0 -- folder tree
# ---------------------------------------------------------------------------

def init_tree():
    head("STEP 0  folder tree")
    for d in [C.REPORTS, C.INPUT, C.OUTPUT] + C.ALL_INPUT_DIRS:
        existed = d.exists()
        d.mkdir(parents=True, exist_ok=True)
        say(f"  {'ok    ' if existed else 'made  '} {d.relative_to(C.ROOT)}")

    readme = C.INPUT / "READ ME - where to put files.txt"
    if not readme.exists():
        readme.write_text(
            "WHERE TO PUT EACH FILE\n"
            "======================\n\n"
            "uniware/          Filled automatically by run_pipeline.py.\n"
            "                  Do not edit by hand - it is cleared each run.\n\n"
            "amazon vc/        Drop the Amazon Vendor Central exports here:\n"
            "                    - Sales report   (ASIN level)\n"
            "                    - Inventory report (ASIN level)\n"
            "                  Keep Amazon's own filenames; they carry the dates.\n\n"
            "retail stores/    Drop every store's sale + stock-on-hand file here.\n"
            "                  Keep the store name in the filename.\n\n"
            "Master mapping table goes in this folder (input/), not in a subfolder.\n"
            "It must have a 'Sku Code' column plus the platform id columns.\n\n"
            "Files are matched by their CONTENTS, so date-stamped names are fine.\n"
            "Anything unrecognised is listed in the report's 'Skipped Files' sheet.\n",
            encoding="utf-8")
        say(f"  made   {readme.relative_to(C.ROOT)}")

    master = find_master()
    if master:
        say(f"\n  master mapping table: {master.name}")
    else:
        warn(f"no master mapping table in {C.INPUT.relative_to(C.ROOT)} — "
             f"drop the 'Marketplace product id Master.xlsx' there")


def find_master():
    for p in C.INPUT.glob("*.xls*"):
        if p.name.startswith(("~$", ".")):
            continue
        if C.MASTER_HINT in p.name.lower():
            return p
    # fall back: any xlsx directly in input/ that has a Sku Code column
    try:
        import pandas as pd
        for p in C.INPUT.glob("*.xls*"):
            if p.name.startswith(("~$", ".")):
                continue
            probe = pd.read_excel(p, header=None, nrows=4)
            flat = " ".join(str(v).lower() for v in probe.values.ravel())
            if "sku code" in flat:
                return p
    except Exception:
        pass
    return None



def preflight_outputs() -> list[str]:
    """Fail before the work, not after it.

    Windows locks any file open in Excel. Discovering that on the last write --
    after every file has been parsed -- is the most annoying possible place to
    stop, so every intended output is test-opened up front.
    """
    head("STEP 0b  output files writable?")
    C.OUTPUT.mkdir(parents=True, exist_ok=True)

    targets = ["Omnichannel_Report.xlsx", "dashboard.html", "fact_sales.csv",
               "fact_inventory.csv", "fact_purchase.csv", "exceptions.csv",
               "reconciliation_checks.csv"]
    targets += [f.name for f in C.OUTPUT.glob("Stock_vs_Sales_*.xlsx")]

    locked = []
    for name in dict.fromkeys(targets):
        f = C.OUTPUT / name
        if not f.exists():
            continue
        try:
            with open(f, "ab"):
                pass
        except OSError:
            locked.append(name)

    if locked:
        say(f"  {len(locked)} file(s) LOCKED — close them and re-run:")
        for n in locked:
            say(f"      {n}")
        say()
        say("  These are almost always open in Excel. The run stops here rather")
        say("  than doing all the work and failing on the last write.")
    else:
        say(f"  all clear ({C.OUTPUT.relative_to(C.ROOT)})")
    return locked

# ---------------------------------------------------------------------------
# STEP 1 -- Uniware pull
# ---------------------------------------------------------------------------

def archive_uniware():
    """Move the current uniware/ contents into _archive/<timestamp>/.

    The PO history file is deliberately NOT archived -- it accumulates across
    runs and is the whole point of the incremental pull.
    """
    files = [f for f in data_files(C.UNIWARE_DIR)
             if f.name != C.PO_HISTORY_FILE.name]
    if not files:
        return
    dest = C.ARCHIVE_DIR / f"uniware_{datetime.now():%Y%m%d_%H%M}"
    dest.mkdir(parents=True, exist_ok=True)
    for f in files:
        shutil.move(str(f), str(dest / f.name))
    say(f"  archived {len(files)} previous file(s) -> "
        f"{dest.relative_to(C.ROOT)}")

    olds = sorted([d for d in C.ARCHIVE_DIR.glob("uniware_*") if d.is_dir()])
    for d in olds[:-C.ARCHIVE_KEEP]:
        shutil.rmtree(d, ignore_errors=True)


def fetch_uniware(args) -> bool:
    head("STEP 1  fetch Uniware")

    if not C.UNIWARE_SCRIPT.exists():
        warn(f"{C.UNIWARE_SCRIPT.name} not found next to this script — skipping fetch")
        return False
    if not (os.environ.get("UNIWARE_USER") and os.environ.get("UNIWARE_PASS")):
        warn("UNIWARE_USER / UNIWARE_PASS not set — skipping fetch, "
             "will use whatever is already in reports/input/uniware/")
        return False

    # Let uniware_exports.py write into a scratch dir, then flatten its
    # timestamped subfolder into input/uniware with stable names. That keeps the
    # reconciler's input folder flat and predictable.
    staging = Path(tempfile.mkdtemp(prefix="uniware_"))
    env = dict(os.environ, UNIWARE_OUTDIR=str(staging))
    # Single facility. uniware_exports.py reads UNIWARE_FACILITY, and we never
    # pass --all-facilities / --combine, so no per-facility fan-out happens.
    if C.FACILITY:
        env["UNIWARE_FACILITY"] = C.FACILITY
        say(f"  facility scope: {C.FACILITY} (single-facility pull)")

    # Two passes. Sales reports only need the recent window; purchase orders
    # drive Ageing and Purchase qty, which need YEARS of history -- pulling them
    # on a 30-day window is why those columns come out blank.
    short = [r for r in C.UNIWARE_REPORTS if r not in C.LONG_WINDOW_REPORTS]
    long_ = [r for r in C.UNIWARE_REPORTS if r in C.LONG_WINDOW_REPORTS]

    cap = C.API_MAX_DAYS
    passes = []
    if short:
        if args.start and args.end:
            passes.append(("recent", short, ["--start", args.start, "--end", args.end]))
        elif args.days > cap:
            for i, (s0, e0) in enumerate(date_slices(args.days, cap), 1):
                passes.append((f"recent {i}/{-(-args.days // cap)}", short,
                               ["--start", s0, "--end", e0]))
        else:
            passes.append(("recent", short, ["--days", str(args.days)]))

    if long_ and not args.no_history:
        # First run backfills; later runs only need the recent slice, since the
        # history file persists. --po-days forces a re-backfill.
        if args.po_days is not None:
            po_days = args.po_days
        elif C.PO_HISTORY_FILE.exists():
            po_days = cap
        else:
            po_days = C.PO_BACKFILL_DAYS
            say(f"  no PO history yet — backfilling {po_days} days "
                f"(one-off; later runs pull {cap})")
        slices = date_slices(po_days, cap)
        for i, (s0, e0) in enumerate(slices, 1):
            passes.append((f"PO history {i}/{len(slices)}", long_,
                           ["--start", s0, "--end", e0]))

    t0 = proc = None
    import time as _t
    t0 = _t.time()
    for label, reports, win in passes:
        cmd = [sys.executable, str(C.UNIWARE_SCRIPT)] + win
        for r in reports:
            cmd += ["--only", r]
        say(f"  [{label}] {' '.join(win)}  ->  {', '.join(reports)}")
        try:
            proc = subprocess.run(cmd, cwd=str(C.ROOT), env=env, timeout=args.timeout)
        except subprocess.TimeoutExpired:
            warn(f"Uniware {label} fetch timed out after {args.timeout}s")
            continue
        if proc.returncode != 0:
            warn(f"uniware_exports.py ({label}) exited {proc.returncode}")
    say(f"  staging: {staging}")
    say()
    if proc is None:
        warn("no Uniware pass ran")
        shutil.rmtree(staging, ignore_errors=True)
        return False

    pulled = sorted(staging.rglob("*.csv"))
    if not pulled:
        warn(f"Uniware fetch produced no CSVs (exit {proc.returncode}) — "
             f"using existing files in reports/input/uniware/")
        shutil.rmtree(staging, ignore_errors=True)
        return False

    archive_uniware()
    C.UNIWARE_DIR.mkdir(parents=True, exist_ok=True)
    stamp = f"{datetime.now():%Y%m%d_%H%M}"

    # PO slices never land as separate files -- dedup would keep only the newest
    # and the history would be lost. They are upserted into one history file.
    po_slices = [f for f in pulled if "purchase_order" in f.name.lower()]
    if po_slices:
        before, after = upsert_po_history(po_slices)
        say(f"  PO history: {len(po_slices)} slice(s) merged — "
            f"{before:,} -> {after:,} unique PO lines "
            f"({C.PO_HISTORY_FILE.name})")
    pulled = [f for f in pulled if f not in po_slices]

    moved = 1 if po_slices else 0
    for src in pulled:
        if "__COMBINED" in src.name:      # prefer the merged file when present
            dest = C.UNIWARE_DIR / f"{src.stem.replace('__COMBINED', '')}_{stamp}.csv"
        elif "__" in src.stem:            # per-facility split; keep the facility tag
            dest = C.UNIWARE_DIR / f"{src.stem}_{stamp}.csv"
        else:
            dest = C.UNIWARE_DIR / f"{src.stem}_{stamp}.csv"
        shutil.copy2(src, dest)
        moved += 1
    shutil.rmtree(staging, ignore_errors=True)

    say(f"\n  landed {moved} file(s) in {C.UNIWARE_DIR.relative_to(C.ROOT)} "
        f"({_t.time() - t0:.0f}s)")
    return True


# ---------------------------------------------------------------------------
# STEP 2 -- manual-drop check
# ---------------------------------------------------------------------------

def check_inputs():
    head("STEP 2  check inputs")

    master = find_master()
    if master:
        say(f"  master mapping   {master.name}")
    else:
        warn("MISSING master mapping table — the build cannot run without it")

    rows = []
    for label, folder in [("Uniware (auto)", C.UNIWARE_DIR)] + list(C.MANUAL_DIRS.items()):
        files = data_files(folder)
        p, when = newest(folder)
        age = (datetime.now() - when).days if when else None
        rows.append((label, len(files), when, age))

        if not files:
            warn(f"{label}: EMPTY ({folder.relative_to(C.ROOT)}) — "
                 f"that channel will be missing from the report")
        elif age is not None and age > C.STALE_AFTER_DAYS:
            warn(f"{label}: newest file is {age} days old ({p.name}) — "
                 f"probably a stale cycle")

    say()
    say(f"  {'source':22}{'files':>7}  {'newest':<20}{'age':>6}")
    for label, n, when, age in rows:
        say(f"  {label:22}{n:>7}  "
            f"{when.strftime('%Y-%m-%d %H:%M') if when else '—':<20}"
            f"{(str(age) + 'd') if age is not None else '—':>6}")
    return master is not None


# ---------------------------------------------------------------------------
# STEP 3 / 4 -- build and verify
# ---------------------------------------------------------------------------

def run_script(script: Path, extra: list[str], label: str) -> int:
    head(label)
    if not script.exists():
        warn(f"{script.name} not found next to this script")
        return 2
    cmd = [sys.executable, str(script), "--input", str(C.INPUT)] + extra
    say(f"  {script.name} {' '.join(extra)}")
    say()
    return subprocess.run(cmd, cwd=str(C.ROOT)).returncode


def build():
    extra = ["--output", str(C.OUTPUT)]
    if C.FACILITY:
        extra += ["--facility", C.FACILITY]
    return run_script(C.ROOT / "reconcile.py", extra, "STEP 3  build report")


def verify():
    extra = ["--report", str(C.OUTPUT)]
    if C.FACILITY:
        extra += ["--facility", C.FACILITY]
    return run_script(C.ROOT / "check_reconcile.py", extra,
                      "STEP 4  verify against raw files")


def stock_vs_sales(a):
    extra = ["--report", str(C.OUTPUT),
             "--period-days", str(a.period_days),
             "--recent-days", str(a.recent_days)]
    if a.asof:
        extra += ["--asof", a.asof]
    return run_script(C.ROOT / "stock_vs_sales.py", extra,
                      "STEP 5  Stock vs Sales")


def show_status():
    head("STATUS")
    master = find_master()
    say(f"  root    {C.ROOT}")
    say(f"  master  {master.name if master else '*** MISSING ***'}")
    say()
    for label, folder in [("uniware", C.UNIWARE_DIR)] + list(C.MANUAL_DIRS.items()):
        files = data_files(folder)
        say(f"  {label} ({len(files)} file(s))")
        for f in files:
            say(f"      {datetime.fromtimestamp(f.stat().st_mtime):%Y-%m-%d %H:%M}  "
                f"{f.stat().st_size:>10,}  {f.name}")
        if not files:
            say("      (empty)")
    say()
    outs = sorted(p for p in C.OUTPUT.glob("*") if p.is_file()) if C.OUTPUT.exists() else []
    say(f"  output ({len(outs)} file(s))")
    for f in outs:
        say(f"      {datetime.fromtimestamp(f.stat().st_mtime):%Y-%m-%d %H:%M}  "
            f"{f.stat().st_size:>10,}  {f.name}")
    if not outs:
        say("      (empty — nothing built yet)")


# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--init", action="store_true",
                    help="create the folder tree and stop")
    ap.add_argument("--status", action="store_true",
                    help="show what is on disk and stop")
    ap.add_argument("--fetch", action="store_true",
                    help="pull fresh data from the Uniware API first. Without "
                         "this the pipeline uses the files already in "
                         "reports/input/uniware/ (the default).")
    ap.add_argument("--offline", action="store_true",
                    help="explicitly use the files on disk. This is already the "
                         "default; the flag is kept so existing commands and "
                         "scheduled tasks keep working.")
    ap.add_argument("--days", type=int, default=C.DEFAULT_DAYS,
                    help=f"Uniware window in days (default {C.DEFAULT_DAYS})")
    ap.add_argument("--start", help="Uniware window start, YYYY-MM-DD")
    ap.add_argument("--end", help="Uniware window end, YYYY-MM-DD")
    ap.add_argument("--po-days", type=int, default=None,
                    help=f"force a purchase-order backfill over this many days "
                         f"(sliced into {C.API_MAX_DAYS}-day requests). Default: "
                         f"{C.PO_BACKFILL_DAYS} on the first run, then "
                         f"{C.API_MAX_DAYS} incremental.")
    ap.add_argument("--no-history", action="store_true",
                    help="skip the long purchase-order pull (faster; leaves "
                         "Ageing and Purchase qty on their fallbacks)")
    ap.add_argument("--timeout", type=int, default=1800,
                    help="seconds to allow for the Uniware fetch (default 1800)")
    ap.add_argument("--force", action="store_true",
                    help="run even when an output file is locked; the affected "
                         "files are written with a _HHMM suffix instead")
    ap.add_argument("--skip-verify", action="store_true")
    ap.add_argument("--skip-merch", action="store_true",
                    help="skip the Stock vs Sales workbook")
    ap.add_argument("--asof", help="Stock vs Sales report date, YYYY-MM-DD "
                                   "(default: today)")
    ap.add_argument("--period-days", type=int, default=C.PERIOD_DAYS,
                    help=f"Stock vs Sales sales window (default {C.PERIOD_DAYS})")
    ap.add_argument("--recent-days", type=int, default=C.RECENT_DAYS,
                    help=f"Stock vs Sales short window (default {C.RECENT_DAYS})")
    ap.add_argument("--dry-run", action="store_true",
                    help="detect input files, print what would load, then stop")
    a = ap.parse_args()

    say(f"DOMIN8 reporting pipeline · {datetime.now():%Y-%m-%d %H:%M}")
    say(f"root: {C.ROOT}")

    init_tree()
    if a.init:
        say("\nTree ready. Drop the master mapping table and the manual reports in, "
            "then run:  python run_pipeline.py")
        return 0
    if a.status:
        show_status()
        return 0

    locked = preflight_outputs()
    if locked and not a.force:
        say(f"\nFAILED: {len(locked)} output file(s) are open elsewhere. "
            f"Close them and re-run, or pass --force to write timestamped copies.")
        return 2
    if locked:
        warn(f"--force: {len(locked)} file(s) were locked and got a _HHMM suffix "
             f"({', '.join(locked[:4])}{'...' if len(locked) > 4 else ''}). The "
             f"canonical files still hold the PREVIOUS run's data.")

    do_fetch = a.fetch or (C.FETCH_BY_DEFAULT and not a.offline)
    if do_fetch:
        fetch_uniware(a)
    else:
        head("STEP 1  Uniware data")
        files = data_files(C.UNIWARE_DIR)
        p, when = newest(C.UNIWARE_DIR)
        say(f"  using the {len(files)} file(s) already in "
            f"{C.UNIWARE_DIR.relative_to(C.ROOT)}  (pass --fetch to pull fresh)")
        if when:
            age = (datetime.now() - when).days
            say(f"  newest: {p.name}  ({when:%Y-%m-%d %H:%M}, {age}d old)")
        if not files:
            warn(f"{C.UNIWARE_DIR.relative_to(C.ROOT)} is empty and --fetch was "
                 f"not passed — every Uniware-sourced number will be missing")

    have_master = check_inputs()

    if a.dry_run:
        dargs = ["--dry-run"] + (["--facility", C.FACILITY] if C.FACILITY else [])
        rc = run_script(C.ROOT / "reconcile.py", dargs, "DRY RUN  detect input files")
        return rc

    if not have_master:
        say("\nFAILED: no master mapping table. Put it in "
            f"{C.INPUT.relative_to(C.ROOT)} and re-run.")
        return 2

    rc = build()
    if rc != 0:
        say(f"\nFAILED: reconcile.py exited {rc}")
        return 2

    if not a.skip_verify:
        vrc = verify()
        if vrc != 0:
            warn(f"check_reconcile.py exited {vrc} — inspect the failed checks above")

    if not a.skip_merch:
        mrc = stock_vs_sales(a)
        if mrc != 0:
            warn(f"stock_vs_sales.py exited {mrc} — the Stock vs Sales workbook "
                 f"may be missing")

    head("DONE")
    for f in sorted(C.OUTPUT.glob("*")):
        if f.is_file():
            say(f"  {f.name:34} {f.stat().st_size:>10,} bytes")
    if WARN:
        say(f"\n  {len(WARN)} warning(s):")
        for w in WARN:
            say(f"    - {w}")
        return 1
    say("\n  no warnings")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(130)
