"""The purchase-order master file.

One CSV, living in `reports/input/uniware/`, that outlives every run and
accumulates every PO line the business has ever raised.

Why it exists
-------------
Uniware's Purchase Orders export is a *window*. Ask for 90 days and you get the
POs raised in those 90 days — so a SKU bought eighteen months ago has no PO on
file, its `Purchase. qty` comes out 0, and `Overall Sell-through` has no
denominator. Pulling a longer window is not a fix either: the API's presets cap
at 90 days, so a long window is really a series of 90-day slices, and each
slice on its own is just as partial as the last.

The answer is to stop treating the export as the source of truth and keep a
master instead. Every slice — fetched or hand-uploaded — is merged into this
file, keyed on (PO Code, Item SkuCode), newest version of a line winning. The
file only ever grows.

Where it lives, and why that matters
------------------------------------
`reports/input/uniware/Purchase_Orders_history.csv`, which in Drive is the
folder the client already uploads to. That is deliberate:

  * they can SEED it — drop in a spreadsheet of every PO to date and the next
    run absorbs it, which is exactly how the history gets its back catalogue;
  * they can SEE it — a master they cannot open is a master they cannot trust;
  * it survives, because `input/` is pulled at the start of every run and the
    grown file is written back at the end.

An earlier version kept it in the hidden `_state/` folder. That worked for the
machine and was useless to the client, who could neither seed it nor check it.

Seed files are not Uniware exports
----------------------------------
A hand-made seed will have headers like "SKU" or "PO Number" rather than
Uniware's "Item SkuCode" and "PO Code". `normalise()` maps the common spellings
onto the canonical ones so a reasonable seed just works, and `merge()` says
plainly what it could not understand rather than silently keeping nothing.
"""

from __future__ import annotations

import re
from pathlib import Path

# The natural key. A PO line is not immutable — it goes from pending to
# received, quantities get corrected — so the same key can arrive many times
# and the newest version has to win.
KEY = ["PO Code", "Item SkuCode"]

# Ordering column. Uniware stamps every line with when it last changed, which
# is a far better ordering than file mtime: files pulled from Drive are written
# fresh on every run, so their mtimes say when we downloaded them, not when the
# data changed. Falling back to file order only when this is absent.
UPDATED_COL = "Updated"

# Columns a frame must have to be a PO table at all.
REQUIRED = KEY
REQUIRED_SET = set(KEY)

# Spellings seen in hand-made seed files, mapped onto Uniware's own headers.
# Keys are matched after casefolding and collapsing runs of non-alphanumerics,
# so "PO  Code", "po_code" and "PO-Code" all land on the same entry.
ALIASES = {
    "pocode": "PO Code",
    "ponumber": "PO Code",
    "pono": "PO Code",
    "purchaseordercode": "PO Code",
    "purchaseorderno": "PO Code",
    "purchaseordernumber": "PO Code",
    "itemskucode": "Item SkuCode",
    "skucode": "Item SkuCode",
    "sku": "Item SkuCode",
    "itemsku": "Item SkuCode",
    "childsku": "Item SkuCode",
    "orderquantity": "Order Quantity",
    "orderedquantity": "Order Quantity",
    "orderedunits": "Order Quantity",
    "qtyordered": "Order Quantity",
    "recievedquantity": "Recieved Quantity",   # Uniware's own spelling
    "receivedquantity": "Recieved Quantity",
    "qtyreceived": "Recieved Quantity",
    "pendingquantity": "Pending Quantity",
    "rejectedquantity": "Rejected Quantity",
    "created": "Created",
    "createddate": "Created",
    "podate": "Created",
    "purchaseorderdate": "Created",
    "updated": UPDATED_COL,
    "lastupdated": UPDATED_COL,
    "modified": UPDATED_COL,
}

# What a Purchase Orders export is called. Uniware's own downloader writes
# "Purchase_Orders.csv" (it replaces spaces with underscores); the client
# downloads "Purchase Orders_18082026151654.csv" from the UI by hand. Both are
# the same report, so the pattern has to be blind to the separator — matching
# only "purchase_order" silently ignored every file the client uploaded.
PO_FILE_RE = re.compile(r"purchase[\s_\-]*orders?", re.I)

HISTORY_NAME = "Purchase_Orders_history.csv"


def _canon(name: str) -> str:
    """Casefold a header down to something aliases can be matched on."""
    return re.sub(r"[^a-z0-9]", "", str(name).strip().lower())


def normalise(df):
    """Rename known column spellings onto Uniware's, and tidy the key columns.

    Only renames where the canonical name is not already present, so a real
    Uniware export passes through untouched and a seed carrying both "SKU" and
    "Item SkuCode" keeps the real one.
    """
    import pandas as pd  # noqa: F401  (imported for the caller's benefit)

    have = set(df.columns)
    rename = {}
    for col in df.columns:
        target = ALIASES.get(_canon(col))
        if target and target not in have and target != col:
            rename[col] = target
            have.add(target)
    if rename:
        df = df.rename(columns=rename)

    # Keys join against the master mapping table and against each other, so a
    # trailing space is the difference between a PO line counting and not.
    for k in KEY:
        if k in df.columns:
            df[k] = df[k].astype("string").str.strip()
    return df


def looks_like_po(df) -> bool:
    """Does this frame carry the columns that make a PO line identifiable?"""
    cols = set(normalise(df.head(0)).columns)
    return all(k in cols for k in REQUIRED)


def _header(path: Path) -> set[str]:
    """The column names, read without pulling the rows in."""
    import pandas as pd

    try:
        if path.suffix.lower() in (".xlsx", ".xls"):
            head = pd.read_excel(path, nrows=0)
        else:
            head = pd.read_csv(path, nrows=0)
    except Exception:                                            # noqa: BLE001
        return set()
    return set(normalise(head).columns)


def po_files(folder: Path, exclude: str = HISTORY_NAME) -> list[Path]:
    """Every purchase-order table in `folder`, the history file excluded.

    Two ways to qualify, and the second one matters more than it looks.

    By name, which catches Uniware's own exports whatever the separator:
    "Purchase_Orders.csv" from the API downloader, "Purchase Orders_1808...csv"
    from a person clicking Download in the UI.

    Or by shape — any table carrying both a PO Code and an Item SkuCode
    column. That is what makes seeding work in practice. The client was asked
    to upload "the file with all the purchase orders to date"; expecting them
    to also guess a filename is how a feature ends up unused. A spreadsheet
    called "PO master 2024-25.xlsx" is unmistakably a PO table by its columns,
    so it is treated as one. Nothing else in the input folder has that pair of
    columns, so the test cannot pull in a sales or inventory export by mistake.
    """
    if not folder.exists():
        return []
    out = []
    for f in sorted(folder.iterdir()):
        if not f.is_file() or f.name == exclude:
            continue
        if f.suffix.lower() not in (".csv", ".xlsx", ".xls"):
            continue
        if PO_FILE_RE.search(f.stem) or REQUIRED_SET <= _header(f):
            out.append(f)
    return out


def _parse_when(series):
    """Parse an Updated column that may be ISO, may be Indian day-first.

    Two passes, and the order matters. Uniware stamps ISO
    ("2026-08-03 16:22:52"); a spreadsheet a person maintains is far more
    likely to be "03/08/2026". Parsing everything with `dayfirst=True` looks
    accommodating and is a trap: pandas applies it to the ISO strings too, so
    2026-08-03 is read as the 8th of March. That silently inverts the ordering
    and a superseded PO line beats the correction that replaced it — which is
    precisely the bug this column exists to prevent.

    So: ISO first, strictly. Only what fails that is retried day-first.
    """
    import pandas as pd

    s = series.astype("string").str.strip()
    out = pd.to_datetime(s, errors="coerce", format="ISO8601")
    missing = out.isna() & s.notna()
    if missing.any():
        out.loc[missing] = pd.to_datetime(s[missing], errors="coerce",
                                          dayfirst=True)
    return out


def _read(path: Path):
    import pandas as pd

    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str)
    return pd.read_csv(path, low_memory=False, dtype=str)


def read_history(path: Path):
    """The master as it stands, or an empty frame.

    Tolerates a seed uploaded as .xlsx next to the canonical .csv name, since
    "upload the file you already have" is the whole point of seeding.
    """
    import pandas as pd

    for cand in (path, path.with_suffix(".xlsx"), path.with_suffix(".xls")):
        if cand.exists():
            try:
                return normalise(_read(cand)), cand
            except Exception:                                    # noqa: BLE001
                return pd.DataFrame(), cand
    return pd.DataFrame(), None


def merge(history_path: Path, slices: list[Path], log=print) -> dict:
    """Fold `slices` into the master at `history_path`. Returns a summary.

    The master is rewritten in full every time. That is deliberate: an append
    would let a superseded version of a PO line sit below its replacement, and
    the whole file is small enough (tens of thousands of rows) that rewriting
    it costs nothing.
    """
    import pandas as pd

    hist, found_at = read_history(history_path)
    before = len(hist)
    if found_at is not None and found_at.name != history_path.name:
        log(f"      seeded from {found_at.name}")

    frames = [hist] if before else []
    unreadable, ignored, absorbed = [], [], []

    for f in slices:
        try:
            df = normalise(_read(f))
        except Exception as exc:                                 # noqa: BLE001
            unreadable.append(f"{f.name}: {exc}")
            continue
        missing = [k for k in REQUIRED if k not in df.columns]
        if missing:
            ignored.append(f"{f.name} (no {', '.join(missing)} column)")
            continue
        frames.append(df)
        absorbed.append(f)

    # `absorbed` is the important half of this return value. A caller that
    # assumes every file it handed over was taken will delete or decline to
    # upload the ones that were not, and a purchase-order file that could not
    # be parsed would then simply cease to exist. Anything not absorbed has to
    # survive as itself.
    if not frames:
        return {"before": before, "after": before, "added": 0,
                "files": 0, "unreadable": unreadable, "ignored": ignored,
                "absorbed": [], "written": False}

    all_rows = pd.concat(frames, ignore_index=True)

    # Newest version of each line wins. Sort by the Updated stamp when there is
    # one, so the result does not depend on which order the files happened to
    # be read in; `kind="stable"` keeps the existing order among ties, which
    # means later files (the fresh slices) still beat the stored history.
    if UPDATED_COL in all_rows.columns:
        order = _parse_when(all_rows[UPDATED_COL])
        all_rows = all_rows.assign(_ord=order).sort_values(
            "_ord", kind="stable", na_position="first").drop(columns="_ord")

    key = [k for k in KEY if k in all_rows.columns]
    if key:
        all_rows = all_rows.dropna(subset=key)
        all_rows = all_rows.drop_duplicates(subset=key, keep="last")

    history_path.parent.mkdir(parents=True, exist_ok=True)
    all_rows.to_csv(history_path, index=False)

    return {"before": before, "after": len(all_rows),
            "added": len(all_rows) - before, "files": len(absorbed),
            "unreadable": unreadable, "ignored": ignored,
            "absorbed": absorbed, "written": True}


def describe(result: dict) -> str:
    """One line a human can read in the run log."""
    if not result.get("written"):
        return "no purchase-order rows found — history left as it was"
    return (f"{result['files']} file(s) merged · "
            f"{result['before']:,} -> {result['after']:,} PO lines "
            f"(+{result['added']:,} new)")
