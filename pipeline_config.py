"""Paths and settings for the DOMIN8 reporting pipeline.

Everything is relative to this file's own folder, so the whole tree can be moved
or cloned onto another machine without editing anything.

    <root>/
      pipeline_config.py        <- you are here
      run_pipeline.py           orchestrator
      uniware_exports.py        Uniware API puller (existing)
      reconcile.py              the reconciler
      check_reconcile.py        ties output back to the raw files
      reports/
        input/
          Marketplace product id Master.xlsx
          uniware/              <- filled automatically by uniware_exports.py
          amazon vc/            <- drop Amazon Vendor Central reports here
          retail stores/        <- drop store sale + SOH files here
          _archive/             <- previous cycles, kept automatically
        output/                 <- the report lands here
"""

from pathlib import Path

ROOT = Path(__file__).resolve().parent

REPORTS = ROOT / "reports"
INPUT = REPORTS / "input"
OUTPUT = REPORTS / "output"

UNIWARE_DIR = INPUT / "uniware"
AMAZON_DIR = INPUT / "amazon vc"
STORES_DIR = INPUT / "retail stores"
ARCHIVE_DIR = INPUT / "_archive"

# Folders the operator fills by hand each cycle. run_pipeline.py warns when one
# is empty or stale instead of quietly producing a report with a channel missing.
MANUAL_DIRS = {
    "Amazon Vendor Central": AMAZON_DIR,
    "Retail stores": STORES_DIR,
}

ALL_INPUT_DIRS = [UNIWARE_DIR, AMAZON_DIR, STORES_DIR, ARCHIVE_DIR]

# The master mapping table. Matched by content, so the filename can change.
MASTER_HINT = "master"

# ---------------------------------------------------------------------------
# Uniware pull
# ---------------------------------------------------------------------------

UNIWARE_SCRIPT = ROOT / "uniware_exports.py"

# Fetch from the Uniware API, or just use whatever is already in
# reports/input/uniware/ ?
#
# Default is FALSE: the pipeline runs on the files on disk. That keeps every run
# reproducible, lets you work against the sample data, and means nobody
# accidentally overwrites a good input set. Pass --fetch to pull.
FETCH_BY_DEFAULT = False

# ---------------------------------------------------------------------------
# Facility scope
#
# Everything is DOMIN8-only for now. This is enforced in three places:
#   1. the Uniware pull is single-facility (never --all-facilities)
#   2. per-facility export files for any OTHER facility are dropped at load
#   3. rows carrying a different Facility value are filtered and COUNTED, so
#      they show up in the report rather than disappearing
#
# To go multi-facility later: set FACILITY = None and the filter turns off.
# ---------------------------------------------------------------------------

FACILITY = "DOMIN8"

# ---------------------------------------------------------------------------
# Timezone for anything a human reads
#
# The pipeline runs on a GitHub runner, whose clock is UTC. Left alone, every
# timestamp in STATUS.txt and every archive folder name comes out in UTC — so a
# run at 01:06 on Tuesday morning IST is filed as 19:36 the previous day, and
# the team reading it is off by five and a half hours and sometimes a date.
#
# Reports are read in India, so they are stamped in India.
# ---------------------------------------------------------------------------

REPORT_TZ = "Asia/Kolkata"

# Reports to pull. These are the `name` values in uniware_exports.py's REPORTS
# dict -- not the dropdown labels.
UNIWARE_REPORTS = [
    "Tally GST Report",
    "Tally Return GST Report",
    "Purchase Orders",
    "Inventory Snapshot",
    "Item Master",
]

# Default window for the Uniware pull.
DEFAULT_DAYS = 90

# Hard API ceiling. uniware_exports.py warns above this ("the UI's date presets
# cap at LAST_90_DAYS ... chunk into 90-day slices if the export comes back
# truncated"), so any longer window is requested as a series of 90-day slices
# rather than one call that may silently truncate.
API_MAX_DAYS = 90

# Warn if a manual-drop folder's newest file is older than this.
STALE_AFTER_DAYS = 10

# Keep this many archived cycles under reports/input/_archive.
ARCHIVE_KEEP = 12

# ---------------------------------------------------------------------------
# Stock vs Sales (the merchandising workbook)
# ---------------------------------------------------------------------------

PERIOD_DAYS = 60     # the main sales window  (the client's "Jun-Jul")
RECENT_DAYS = 15     # the short window       (the client's "Last 15 days")

# Merchandiser judgement columns live here and are joined in each run, so notes
# survive regeneration. Seeded automatically on the first run.
REORDER_OVERRIDES = INPUT / "reorder_status.csv"

# Purchase-order history drives Ageing and Purchase qty. A 30-day pull is far
# too short for those - pull POs over a long window at least once.
# Backfill window, used only on the FIRST run (when the history file does not
# exist yet) or when --po-days is passed explicitly. After that each run only
# needs the recent slice, because the history file persists and is upserted.
PO_BACKFILL_DAYS = 730

# These reports are pulled on PO_BACKFILL_DAYS instead of DEFAULT_DAYS, because
# they are cumulative rather than a recent-activity feed.
LONG_WINDOW_REPORTS = ["Purchase Orders"]

# Long-window reports are pulled in API_MAX_DAYS slices and upserted into this
# file, which accumulates across runs. It is what the reconciler actually reads,
# so Ageing and Purchase qty keep improving instead of resetting every cycle.
PO_HISTORY_FILE = UNIWARE_DIR / "Purchase_Orders_history.csv"

# Natural key for the upsert. A PO line changes over time (pending -> received),
# so the newest version of each key wins.
PO_HISTORY_KEY = ["PO Code", "Item SkuCode"]
