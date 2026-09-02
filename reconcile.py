#!/usr/bin/env python3
"""
Omnichannel sales & inventory reconciliation.

Reads the raw reports you collect each cycle (Amazon Vendor Central, Uniware,
retail-store Excel/CSV), resolves every platform identifier to your master SKU,
and writes one reconciled Excel workbook + an HTML dashboard.

    pip install pandas openpyxl
    python reconcile.py --input ./Sample --output ./report

Folder layout expected (folder names are free-form; files are matched by content):

    Sample/
      Marketplace product id Master.xlsx     <- master mapping table (required)
      Amazon VC/                             <- any Amazon VC report CSVs
      Uniware/                               <- any Uniware export CSVs
      Retail Stores/                         <- store sale / SOH files

Adding a new store: add one entry to STORE_PROFILES near the bottom. No other
code changes. Run with --dry-run to see what was detected before building.

KEY CONCEPT -- sell-in vs sell-out
  sell_in  = what you invoiced TO a channel   (Uniware Tally GST)
  sell_out = what the end customer bought     (Amazon VC, store sale reports, D2C)
These are different measures and are NEVER summed. Every fact row carries a
`flow` column so a report can't accidentally add them.
"""

from __future__ import annotations

import argparse
import json
import hashlib
import re
import sys
from datetime import datetime
import unicodedata
import warnings
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

# Windows consoles default to cp1252, which cannot encode the rupee sign
# this pipeline prints. That killed a run on a developer machine while
# working fine in CI, where the console is UTF-8. Force UTF-8 and replace
# anything unprintable rather than raising: a report must not die over a
# currency symbol.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


# ===========================================================================
# CONFIG 1 -- master mapping table
# Wide column name in your master sheet  ->  canonical identifier type.
# Add a column here when you onboard a new marketplace.
# ===========================================================================

MASTER_ID_COLUMNS = {
    "Sku Code":          "sku_code",
    "Item Name":         "style_code",
    "Myntra article no": "myntra_article",
    "Myntra style id":   "myntra_style_id",
    "AZ sku":            "az_sku",
    "AZ ASIN":           "az_asin",
    "Nykaa vendor sku":  "nykaa_vendor_sku",
    "Nykaa SKU":         "nykaa_sku",
    "FK SKU":            "fk_sku",
    "FK FSN":            "fk_fsn",
    "TCQ SKU":           "tcq_sku",
    "TCQ ID":            "tcq_id",
    "AJIO SKU":          "ajio_sku",
    "Jio code":          "jio_code",
    "FLIPKART SKU":      "flipkart_sku",
    "FSN":               "fsn",
}

# Identifier spaces tried, in order, when a declared type doesn't hit.
FALLBACK_ID_TYPES = ["sku_code", "az_asin", "az_sku", "nykaa_vendor_sku",
                     "fk_sku", "tcq_sku", "ajio_sku", "myntra_article",
                     "flipkart_sku", "jio_code"]

# ===========================================================================
# CONFIG 2 -- channel ledgers
# Uniware "Channel Ledger" value -> (canonical channel, channel type, flow).
# THIS TABLE IS WHAT KEEPS SELL-IN AND SELL-OUT APART. Any ledger not listed
# here is reported as `unknown` in the audit rather than silently guessed.
# ===========================================================================

CHANNEL_MAP = {
    "MYNTRAPPMP":              ("Myntra",                "marketplace",  "sell_in"),
    "NYKAA_FASHION":           ("Nykaa Fashion",         "marketplace",  "sell_in"),
    "NYKAA_COM":               ("Nykaa",                 "marketplace",  "sell_in"),
    "TATACLIQ":                ("Tata Cliq",             "marketplace",  "sell_in"),
    "AJIO":                    ("AJIO",                  "marketplace",  "sell_in"),
    "FLIPKART":                ("Flipkart",              "marketplace",  "sell_in"),
    # Cocoblu Retail is Amazon's Indian vendor entity -- these are Amazon POs.
    "COCOBLU RETAIL LIMITED.": ("Amazon Vendor Central", "marketplace",  "sell_in"),
    "COCOBLU P.O":             ("Amazon Vendor Central", "marketplace",  "sell_in"),
    "COCOBLU PO":              ("Amazon Vendor Central", "marketplace",  "sell_in"),
    # Own website + its checkout/financing partners are direct to consumer.
    "DOMIN8ACTIVE":            ("D2C Website",           "d2c",          "sell_out"),
    "SNAPMINT":                ("D2C Website",           "d2c",          "sell_out"),
    "POPCLUB":                 ("D2C Website",           "d2c",          "sell_out"),
    # Retail stores: Uniware side is the B2B invoice = sell-in.
    "ALL SPORTS":              ("All Sports",            "retail_store", "sell_in"),
    "INCS KOC":                ("INCS",                  "retail_store", "sell_in"),
    "J&J_KOHIMA_B2B":          ("Jack & Jill",           "retail_store", "sell_in"),
    "J&J_5TH_MILE_B2B":        ("Jack & Jill",           "retail_store", "sell_in"),
    "J&J_DIMAPUR_B2B":         ("Jack & Jill",           "retail_store", "sell_in"),
    "J&J_CHUMOUKEDIMA_B2B":    ("Jack & Jill",           "retail_store", "sell_in"),
    "PROMO":                   ("Promotions",            "internal",     "exclude"),
}

# ===========================================================================
# Canonical schemas
# ===========================================================================

SALES_COLS = ["source_file", "source_type", "channel", "channel_type", "location",
              "flow", "grain", "period_start", "period_end",
              "external_id", "external_id_type", "matched_via",
              "master_sku", "style_code",
              "qty_sold", "qty_ordered", "qty_returned",
              "gross_value", "net_value", "discount_value", "returns_value",
              "value_basis"]

INV_COLS = ["source_file", "source_type", "channel", "channel_type", "location",
            "snapshot_date", "external_id", "external_id_type", "matched_via",
            "master_sku", "style_code", "qty_on_hand", "value_on_hand"]

# Purchase orders are INBOUND -- a third flow, kept in its own table so it can
# never be mistaken for sales. "purchase" is not sell_in and not sell_out.
PO_COLS = ["source_file", "source_type", "po_code", "po_status", "po_type",
           "created_date", "approved_date", "delivery_date", "vendor", "vendor_code",
           "facility", "external_id", "external_id_type", "matched_via",
           "master_sku", "style_code",
           "qty_ordered", "qty_received", "qty_pending", "qty_rejected",
           "unit_price", "sub_total", "total", "ageing_days"]

# Folder names never scanned for input files.
EXCLUDE_DIRS = {"output", "_archive", "archive", "exports", ".git",
                "__pycache__", "old", "backup"}

# Files the pipeline itself owns. They live in the input tree but are config,
# not data, so they must never be classified as a source.
EXCLUDE_FILES = {"reorder_status.csv", "channel_map.csv", "reconciliation_checks.csv"}

MONTHS = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
          "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}


# ===========================================================================
# Helpers
# ===========================================================================

def norm_id(s) -> pd.Series:
    """Uppercase/trim an identifier column; blanks become None."""
    out = (pd.Series(s).astype("string").astype(object).fillna("")
           .map(lambda v: unicodedata.normalize("NFKC", str(v)))
           .str.strip().str.upper())
    return out.replace({"NAN": None, "": None, "NONE": None, "NAT": None})


def money(s) -> pd.Series:
    """'₹10,422.42' / '0.00%' / '' -> float. Amazon exports are full of these."""
    return pd.to_numeric(
        pd.Series(s).astype(str)
        .str.replace(r"[₹$€,%\s]", "", regex=True)
        .str.replace(" ", "", regex=False)
        .replace({"": None, "nan": None, "-": None, "None": None, "NaN": None}),
        errors="coerce")


def num(s) -> pd.Series:
    return pd.to_numeric(pd.Series(s), errors="coerce")


def read_csv(path: Path, **kw) -> pd.DataFrame:
    """CSV read that survives the encodings these exports arrive in."""
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return pd.read_csv(path, encoding=enc, low_memory=False, **kw)
        except UnicodeDecodeError:
            continue
    return pd.read_csv(path, encoding="latin-1", engine="python",
                       on_bad_lines="skip", low_memory=False, **kw)


def clean_cols(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).replace(" ", " ").strip().strip('"') for c in df.columns]
    return df


def dates_in_name(name: str):
    """Pull a (start, end) period out of a filename, if present."""
    pats = [(r"(\d{1,2}-\d{1,2}-\d{4})[_-](\d{1,2}-\d{1,2}-\d{4})", "%d-%m-%Y"),
            (r"(\d{4}-\d{2}-\d{2})[_-](\d{4}-\d{2}-\d{2})", "%Y-%m-%d"),
            (r"(\d{1,2}-\d{1,2}-\d{4})", "%d-%m-%Y"),
            (r"(\d{4}-\d{2}-\d{2})", "%Y-%m-%d")]
    for pat, fmt in pats:
        m = re.search(pat, name)
        if not m:
            continue
        g = m.groups()
        a = pd.to_datetime(g[0], format=fmt, errors="coerce")
        b = pd.to_datetime(g[1], format=fmt, errors="coerce") if len(g) > 1 else a
        if pd.notna(a):
            return a, b
    return pd.NaT, pd.NaT


def parse_dates(series) -> pd.Series:
    """Day-first date parsing that leaves already-typed columns alone."""
    s = pd.Series(series)
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    txt = s.astype(str).str.strip()
    # ISO first (Uniware timestamps), then day-first (Indian dd-mm-yyyy). Trying
    # ISO up front keeps pandas from emitting a dayfirst warning on every load.
    iso = pd.to_datetime(txt, format="ISO8601", errors="coerce")
    if iso.notna().sum() >= txt.ne("").sum() * 0.9:
        return iso
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return pd.to_datetime(txt, dayfirst=True, errors="coerce")


def drop_total_rows(df: pd.DataFrame, col=0) -> pd.DataFrame:
    """Pivot exports carry Grand Total rows that would double-count."""
    key = df.columns[col] if isinstance(col, int) else col
    keep = df[key].notna()
    keep &= ~df[key].astype(str).str.contains(r"grand\s*total|^total$|^sum\b",
                                              case=False, regex=True, na=False)
    return df[keep]



def writable(path: Path) -> bool:
    """Can we actually write here? Windows locks any file open in Excel."""
    try:
        with open(path, "ab"):
            return True
    except OSError:
        return False


def safe_excel_path(path: Path) -> Path:
    """Windows locks a workbook that is open in Excel, and a network share can
    refuse an in-place overwrite. Either way, write beside it rather than dying
    on the last step after all the work is done."""
    if writable(path):
        return path
    alt = path.with_name(f"{path.stem}_{datetime.now():%H%M}{path.suffix}")
    print(f"  !! {path.name} is locked or read-only (open in Excel?) "
          f"— writing {alt.name} instead")
    return alt


def safe_to_csv(df, path: Path, **kw) -> Path:
    """Same protection for the CSVs. A locked exceptions.csv used to kill the
    run AFTER all the parsing was done, which is the most annoying possible
    place to fail."""
    target = path if writable(path) else path.with_name(
        f"{path.stem}_{datetime.now():%H%M}{path.suffix}")
    if target != path:
        print(f"  !! {path.name} is locked or read-only (open in Excel?) "
              f"— writing {target.name} instead. Downstream steps will read the "
              f"OLD {path.name}; close it and re-run for a clean set.")
    df.to_csv(target, index=False, **kw)
    return target

# ===========================================================================
# Identity resolution
# ===========================================================================

class Resolver:
    """Resolves platform identifiers to master SKUs. Tracks every miss."""

    def __init__(self, master: pd.DataFrame):
        self.master = master
        rows = []
        for col, id_type in MASTER_ID_COLUMNS.items():
            if col not in master.columns:
                continue
            vals = master[col]
            if pd.api.types.is_numeric_dtype(vals):     # Jio code / style id -> kill the .0
                vals = vals.map(lambda v: "" if pd.isna(v) else f"{int(v)}")
            sub = pd.DataFrame({
                "identifier_type": id_type,
                "identifier_value": norm_id(vals).values,
                "master_sku": master["Sku Code"].values,
                "style_code": master["Item Name"].values,
            }).dropna(subset=["identifier_value"])
            rows.append(sub)

        self.bridge = (pd.concat(rows, ignore_index=True)
                       .drop_duplicates(["identifier_type", "identifier_value", "master_sku"]))

        uniq = (self.bridge[self.bridge.identifier_type != "style_code"]
                .drop_duplicates(["identifier_type", "identifier_value"]))
        self._lookup = {(r.identifier_type, r.identifier_value): (r.master_sku, r.style_code)
                        for r in uniq.itertuples()}
        self._styles = set(self.bridge.loc[self.bridge.identifier_type == "sku_code",
                                          "style_code"].dropna())
        self.exceptions: list[dict] = []

    def resolve(self, df: pd.DataFrame, id_type: str, source_file: str) -> pd.DataFrame:
        df = df.copy()
        df["external_id"] = norm_id(df["external_id"]).values
        df["external_id_type"] = id_type

        skus, styles, via = [], [], []
        for val in df["external_id"]:
            hit = self._lookup.get((id_type, val))
            used = id_type
            if hit is None:                                  # try the other id spaces
                for t in FALLBACK_ID_TYPES:
                    hit = self._lookup.get((t, val))
                    if hit:
                        used = t
                        break
            if hit is None and val in self._styles:          # style-level, not size-level
                skus.append(None); styles.append(val); via.append("style_code")
                continue
            if hit is None:
                skus.append(None); styles.append(None); via.append(None)
                continue
            skus.append(hit[0]); styles.append(hit[1]); via.append(used)

        df["master_sku"], df["style_code"], df["matched_via"] = skus, styles, via

        qty_col = "qty_sold" if "qty_sold" in df else ("qty_on_hand" if "qty_on_hand" in df else None)
        val_col = next((c for c in ("net_value", "gross_value", "value_on_hand") if c in df), None)
        for r in df[df.matched_via.isna()].itertuples():
            self.exceptions.append({
                "source_file": source_file,
                "external_id": r.external_id,
                "declared_id_type": id_type,
                "qty": getattr(r, qty_col, None) if qty_col else None,
                "value": getattr(r, val_col, None) if val_col else None,
                "reason": "identifier not present in master mapping table",
            })
        return df

    def exception_frame(self) -> pd.DataFrame:
        cols = ["source_file", "external_id", "declared_id_type", "reason", "qty", "value"]
        if not self.exceptions:
            return pd.DataFrame(columns=cols)
        return (pd.DataFrame(self.exceptions)
                .groupby(["source_file", "external_id", "declared_id_type", "reason"],
                         as_index=False)
                .agg(qty=("qty", "sum"), value=("value", "sum"))
                .sort_values("value", ascending=False, na_position="last")[cols])


# ===========================================================================
# Source detection -- match files by their CONTENT, not their filename
# ===========================================================================

@dataclass
class Detected:
    path: Path
    kind: str
    sheet: str | None = None
    header: int = 0
    note: str = ""
    profile: dict = field(default_factory=dict)


def _peek_csv(path: Path, skiprows=0) -> list[str]:
    try:
        return list(clean_cols(read_csv(path, nrows=3, skiprows=skiprows)).columns)
    except Exception:
        return []


def _sheet_header(path: Path, sheet: str, probe: list[str]) -> int | None:
    """Find the row index that holds the real header for a messy Excel sheet."""
    try:
        raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=25)
    except Exception:
        return None
    for i in range(len(raw)):
        cells = {str(v).replace(" ", " ").strip().lower()
                 for v in raw.iloc[i].tolist() if pd.notna(v)}
        if sum(any(p in c for c in cells) for p in probe) >= 2:
            return i
    return None


def detect(folder: Path, store_profiles: list[dict]) -> tuple[Path | None, list[Detected]]:
    """Walk the input folder and classify every file."""
    master, found = None, []

    for path in sorted(folder.rglob("*")):
        if not path.is_file() or path.name.startswith(("~$", ".")):
            continue
        if path.name.lower() in EXCLUDE_FILES:
            continue
        # Skip archives, scratch and previous outputs. Without this, an
        # _archive/ folder of last cycle's files would double-count everything.
        if any(part.startswith("_") or part.lower() in EXCLUDE_DIRS
               for part in path.relative_to(folder).parts[:-1]):
            continue
        suffix = path.suffix.lower()
        if suffix not in (".csv", ".xlsx", ".xls", ".xlsm"):
            continue

        # ---- master mapping table ----
        if suffix in (".xlsx", ".xlsm", ".xls"):
            try:
                probe = pd.read_excel(path, header=None, nrows=4)
                flat = " ".join(str(v).lower() for v in probe.values.ravel() if pd.notna(v))
                if "sku code" in flat and ("asin" in flat or "fsn" in flat or "myntra" in flat):
                    master = path
                    continue
            except Exception:
                pass

        # ---- CSVs ----
        if suffix == ".csv":
            head0 = _peek_csv(path)
            head1 = _peek_csv(path, skiprows=1)          # Amazon VC banner line
            h0 = {c.lower() for c in head0}
            h1 = {c.lower() for c in head1}

            # Inventory is tested BEFORE sales: the Amazon inventory report also
            # contains "Unfilled Customer Ordered Units", which would otherwise
            # make it look like a sales report.
            def _amz(h):
                inv = any("on hand" in c or "on-hand" in c for c in h)
                sale = any(c in ("shipped units", "ordered revenue", "shipped cogs") for c in h)
                return "inventory" if ("asin" in h and inv) else ("sales" if ("asin" in h and sale) else None)

            for h, note in ((h1, "skiprows=1"), (h0, "")):
                got = _amz(h)
                if got:
                    found.append(Detected(path, f"amazon_vc_{got}", note=note))
                    break
            else:
                got = None
            if got:
                continue

            if {"product sku code", "channel ledger", "qty"} <= h0:
                found.append(Detected(path, "uniware_tally_gst")); continue
            if "channel ledger" in h0 and "product sku code" not in h0:
                found.append(Detected(path, "uniware_returns")); continue
            if {"po code", "item skucode"} <= h0 or (
                    "po code" in h0 and "order quantity" in h0):
                found.append(Detected(path, "uniware_purchase_orders")); continue
            if {"item skucode", "facility"} <= h0:
                found.append(Detected(path, "uniware_inventory")); continue
            if {"product code", "category code"} <= h0 and "mrp" in h0:
                found.append(Detected(path, "uniware_item_master")); continue

            # store CSVs -> profile match on filename
            prof = next((p for p in store_profiles
                         if p.get("match") and re.search(p["match"], path.name, re.I)
                         and p["file_type"] == "csv"), None)
            if prof:
                found.append(Detected(path, prof["kind"], profile=prof,
                                      header=prof.get("header", 0))); continue

            found.append(Detected(path, "UNRECOGNISED", note="no rule matched"))
            continue

        # ---- Excel workbooks: classify each sheet ----
        try:
            xl = pd.ExcelFile(path)
        except Exception as e:
            found.append(Detected(path, "UNREADABLE", note=str(e)[:80])); continue

        for sheet in xl.sheet_names:
            prof = next((p for p in store_profiles
                         if p["file_type"] == "xlsx"
                         and (not p.get("match") or re.search(p["match"], path.name, re.I))
                         and (not p.get("sheet") or re.search(p["sheet"], sheet, re.I))), None)
            if prof:
                hdr = _sheet_header(path, sheet, prof["probe"])
                if hdr is not None:
                    found.append(Detected(path, prof["kind"], sheet=sheet,
                                          header=hdr, profile=prof))
                    continue
            found.append(Detected(path, "UNRECOGNISED", sheet=sheet,
                                  note="no profile matched this sheet"))

    return master, found



# Point-in-time sources: two files = two snapshots of the SAME stock, so summing
# them double-counts. Transactional sources overlap when two pulls share days.
# Either way only the newest file per (kind, facility) is loaded.
SNAPSHOT_KINDS = {"uniware_inventory", "amazon_vc_inventory", "uniware_item_master",
                  "uniware_purchase_orders", "store_flat_soh", "store_matrix_soh"}


def _facility_tag(name: str) -> str:
    """uniware_exports.py --all-facilities writes Report__FACILITY.csv."""
    m = re.search(r"__([A-Za-z0-9_]+?)(?:_\d{8}_\d{4})?\.csv$", name)
    return (m.group(1) or "").upper() if m else ""


def dedupe(found: list[Detected], facility: str | None = None
           ) -> tuple[list[Detected], list[dict]]:
    """Keep one file per (kind, sheet, facility). Newest wins.

    This is what stops a second pull on the same day from doubling every number.
    A __COMBINED file supersedes the per-facility splits of the same report.
    """
    superseded: list[dict] = []

    want = (facility or "").upper()
    if want:
        # A per-facility export for a DIFFERENT facility never belongs in a
        # single-facility run. Files with no facility tag are kept -- most
        # exports carry the facility as a column instead, filtered below.
        kept = []
        for d in found:
            tag = _facility_tag(d.path.name)
            if tag and tag != want and "COMBINED" not in tag:
                superseded.append({"file": d.path.name, "kind": d.kind,
                                   "reason": f"facility {tag} — this run is scoped "
                                             f"to {want}"})
                continue
            kept.append(d)
        found = kept

    combined_kinds = {d.kind for d in found if "__COMBINED" in d.path.name.upper()}
    keep_stage = []
    for d in found:
        tag = _facility_tag(d.path.name)
        is_comb = "__COMBINED" in d.path.name.upper()
        if d.kind in combined_kinds and tag and not is_comb:
            superseded.append({"file": d.path.name, "kind": d.kind,
                               "reason": "per-facility split; a __COMBINED file for "
                                         "this report is present"})
            continue
        keep_stage.append((d, "" if is_comb else tag))

    # byte-identical files (a re-download, or a manual copy)
    seen_hash: dict[tuple, str] = {}
    staged = []
    for d, tag in keep_stage:
        try:
            # keyed on the sheet too: one workbook legitimately yields several
            # Detected entries (a Sale sheet and a Soh sheet) from one file.
            h = (hashlib.sha256(d.path.read_bytes()).hexdigest(), d.sheet or "")
        except OSError:
            h = None
        if h and h in seen_hash:
            superseded.append({"file": d.path.name, "kind": d.kind,
                               "reason": f"byte-identical duplicate of "
                                         f"{seen_hash[h]}"})
            continue
        if h:
            seen_hash[h] = d.path.name
        staged.append((d, tag))

    groups: dict[tuple, list[Detected]] = {}
    for d, tag in staged:
        groups.setdefault((d.kind, d.sheet or "", tag), []).append(d)

    keep = []
    for (kind, _sheet, tag), items in groups.items():
        if len(items) == 1 or kind in ("UNRECOGNISED", "UNREADABLE"):
            keep.extend(items)
            continue
        items.sort(key=lambda x: x.path.stat().st_mtime)
        winner = items[-1]
        keep.append(winner)
        why = ("point-in-time snapshot — summing two would double-count stock"
               if kind in SNAPSHOT_KINDS
               else "overlapping window — summing two would double-count sales")
        for loser in items[:-1]:
            note = f"{why}; superseded by {winner.path.name}"
            # A newer pull with a much smaller file usually means a shorter window
            # was requested. Newest-wins would then silently drop history.
            try:
                if loser.path.stat().st_size > winner.path.stat().st_size * 1.2:
                    note += (f"  ** WARNING: the superseded file is larger "
                             f"({loser.path.stat().st_size:,} vs "
                             f"{winner.path.stat().st_size:,} bytes) — the newer "
                             f"pull may cover a shorter window **")
            except OSError:
                pass
            superseded.append({"file": loser.path.name, "kind": kind, "reason": note})
    return keep, superseded

# ===========================================================================
# Adapters -- one per file shape. Each returns canonical columns.
# ===========================================================================

def a_amazon_sales(d: Detected) -> tuple[pd.DataFrame, dict]:
    skip = 1 if "skiprows" in d.note else 0
    df = clean_cols(read_csv(d.path, skiprows=skip))
    start, end = dates_in_name(d.path.name)
    out = pd.DataFrame({"external_id": df["ASIN"]})
    out["qty_ordered"] = num(df.get("Ordered Units"))
    out["qty_sold"] = num(df.get("Shipped Units"))
    out["qty_returned"] = num(df.get("Customer Returns"))
    out["gross_value"] = money(df.get("Ordered Revenue"))
    out["net_value"] = money(df.get("Shipped COGS"))
    out["period_start"], out["period_end"] = start, end
    return out, dict(channel="Amazon Vendor Central", channel_type="marketplace",
                     location="Amazon FC (IN)", flow="sell_out", grain="period",
                     id_type="az_asin", kind="sales")


def a_amazon_inv(d: Detected) -> tuple[pd.DataFrame, dict]:
    skip = 1 if "skiprows" in d.note else 0
    df = clean_cols(read_csv(d.path, skiprows=skip))
    start, _ = dates_in_name(d.path.name)
    col = next((c for c in df.columns if "sellable on hand units" in c.lower()
                or "sellable on-hand units" in c.lower()), None)
    valcol = next((c for c in df.columns if "sellable on-hand inventory" in c.lower()), None)
    out = pd.DataFrame({"external_id": df["ASIN"]})
    out["qty_on_hand"] = num(df[col]) if col else None
    out["value_on_hand"] = money(df[valcol]) if valcol else None
    out["snapshot_date"] = start
    return out, dict(channel="Amazon Vendor Central", channel_type="marketplace",
                     location="Amazon FC (IN)", grain="snapshot",
                     id_type="az_asin", kind="inventory")


def a_uniware_gst(d: Detected) -> tuple[pd.DataFrame, dict]:
    df = clean_cols(read_csv(d.path))
    out = pd.DataFrame({
        "external_id": df["Product SKU Code"],
        "_ledger": df["Channel Ledger"].astype(str).str.strip(),
        "qty_sold": num(df["Qty"]),
        "gross_value": money(df["Total"]),
        "net_value": money(df.get("Sales")),
    })
    dt = parse_dates(df["Date"])
    out["period_start"] = out["period_end"] = dt
    return out, dict(grain="day", id_type="sku_code", kind="sales", ledger_col="_ledger")


def a_uniware_returns(d: Detected) -> tuple[pd.DataFrame, dict]:
    """No SKU column in this export -> channel grain only, can never join to SKU."""
    df = clean_cols(read_csv(d.path))
    out = pd.DataFrame({
        "_ledger": df["Channel Ledger"].astype(str).str.strip(),
        "qty_returned": num(df["Qty"]),
        "returns_value": money(df["Total"]),
    })
    dt = parse_dates(df["Date"])
    out["period_start"] = out["period_end"] = dt
    return out, dict(grain="day", id_type=None, kind="sales", ledger_col="_ledger")


def a_uniware_inv(d: Detected) -> tuple[pd.DataFrame, dict]:
    df = clean_cols(read_csv(d.path))
    out = pd.DataFrame({
        "external_id": df["Item SkuCode"],
        "location": df["Facility"].astype(str).str.strip(),
        "qty_on_hand": num(df["Inventory"]),
    })
    out["value_on_hand"] = out["qty_on_hand"] * money(df.get("MRP"))
    m = re.search(r"(\d{14})", d.path.stem)
    out["snapshot_date"] = (pd.to_datetime(m.group(1), format="%d%m%Y%H%M%S", errors="coerce")
                            if m else dates_in_name(d.path.name)[0])
    return out, dict(channel="Warehouse", channel_type="warehouse", grain="snapshot",
                     id_type="sku_code", kind="inventory")


def a_uniware_po(d: Detected) -> tuple[pd.DataFrame, dict]:
    """Uniware Purchase Orders export -- inbound supply, its own flow.

    Note Uniware ships the column as 'Recieved Quantity' (sic); both spellings
    are accepted so a future fix upstream doesn't break the load.
    """
    df = clean_cols(read_csv(d.path))

    def col(*names):
        for n in names:
            for c in df.columns:
                if c.strip().lower() == n.lower():
                    return df[c]
        return None

    out = pd.DataFrame({"external_id": col("Item SkuCode", "Vendor SkuCode")})
    out["po_code"] = col("PO Code")
    out["po_status"] = col("Purchase Order Status")
    out["po_type"] = col("Type")
    out["created_date"] = parse_dates(col("Created"))
    out["approved_date"] = parse_dates(col("PO Approved Date"))
    out["delivery_date"] = parse_dates(col("Delivery Date"))
    out["vendor"] = col("Vendor Name")
    out["vendor_code"] = col("Vendor Code")
    out["facility"] = col("Facility")
    out["qty_ordered"] = num(col("Order Quantity"))
    out["qty_received"] = num(col("Recieved Quantity", "Received Quantity"))
    out["qty_rejected"] = num(col("Rejected Quantity"))
    out["qty_pending"] = num(col("Pending Quantity"))
    out["unit_price"] = money(col("Unit Price"))
    out["sub_total"] = money(col("Sub Total"))
    out["total"] = money(col("Total"))
    out["ageing_days"] = num(col("PO Ageing (Days)"))
    return out, dict(id_type="sku_code", kind="purchase")


def a_store_flat(d: Detected) -> tuple[pd.DataFrame, dict]:
    """Row-per-line store export (All Sports sale/SOH, INCS). Profile-driven."""
    p = d.profile
    if d.path.suffix.lower() == ".csv":
        df = clean_cols(read_csv(d.path, header=d.header))
    else:
        df = clean_cols(pd.read_excel(d.path, sheet_name=d.sheet, header=d.header))

    def col(key):
        want = p.get(key)
        if not want:
            return None
        for c in df.columns:
            if c.strip().lower() == want.strip().lower():
                return c
        for c in df.columns:
            if want.strip().lower() in c.strip().lower():
                return c
        return None

    idc = col("id_col")
    if idc is None:
        raise KeyError(f"{p['kind']}: id column {p.get('id_col')!r} not in {list(df.columns)[:12]}")
    df = df[df[idc].notna()]
    df = drop_total_rows(df, idc)

    out = pd.DataFrame({"external_id": df[idc]})
    for canon, key in [("qty_sold", "qty_col"), ("qty_on_hand", "soh_col"),
                       ("gross_value", "gross_col"), ("net_value", "net_col"),
                       ("discount_value", "disc_col")]:
        c = col(key)
        if c is not None:
            out[canon] = num(df[c])

    loc = col("location_col")
    out["location"] = (df[loc].astype(str).str.strip() if loc is not None
                       else p.get("location", p["channel"]))

    kind = "inventory" if p["kind"].endswith("soh") else "sales"
    if kind == "inventory":
        out["snapshot_date"] = _snapshot_date(d, p)
    else:
        dc = col("date_col")
        if dc is not None:
            dt = parse_dates(df[dc])
            out["period_start"] = out["period_end"] = dt
        elif p.get("month_col") and col("month_col") is not None:
            mo = df[col("month_col")].astype(str).str[:3].str.upper().map(MONTHS)
            # No year in the file. Take the most recent occurrence of that month
            # at or before the file's own timestamp -- never a future month.
            ref = pd.Timestamp(d.path.stat().st_mtime, unit="s")
            yr = p.get("year") or [
                (ref.year if (pd.notna(m) and m <= ref.month) else ref.year - 1)
                for m in mo]
            out["period_start"] = pd.to_datetime(
                dict(year=pd.Series(yr, index=mo.index) if not isinstance(yr, int)
                          else pd.Series(yr, index=mo.index),
                     month=mo.fillna(1).astype(int), day=1), errors="coerce")
            out["period_end"] = out["period_start"] + pd.offsets.MonthEnd(0)
        else:
            s, e = dates_in_name(d.path.name)
            if pd.notna(s) and p.get("period_days"):
                s = s - pd.Timedelta(days=p["period_days"] - 1)
            out["period_start"], out["period_end"] = s, e

    return out, dict(channel=p["channel"], channel_type="retail_store",
                     flow=p.get("flow", "sell_out"), grain=p.get("grain", "period"),
                     id_type=p.get("id_type", "sku_code"), kind=kind)


def _snapshot_date(d: Detected, p: dict):
    """Snapshot date from the filename, else from a header line like 'Soh As on 17th Aug 2026'."""
    s, _ = dates_in_name(d.path.name)
    if pd.notna(s):
        return s
    if d.path.suffix.lower() != ".csv":
        try:
            hdr = pd.read_excel(d.path, sheet_name=d.sheet, header=None, nrows=d.header)
            flat = " ".join(str(v) for v in hdr.values.ravel() if pd.notna(v))
            m = re.search(r"(\d{1,2})\w*\s+([A-Za-z]{3})\w*\s+(\d{4})", flat)
            if m:
                return pd.Timestamp(int(m.group(3)), MONTHS[m.group(2)[:3].upper()],
                                    int(m.group(1)))
        except Exception:
            pass
    return pd.NaT


def a_store_matrix(d: Detected) -> tuple[pd.DataFrame, dict]:
    """SKU rows x store columns, one measure (Jack & Jill SOH pivot)."""
    p = d.profile
    df = clean_cols(read_csv(d.path, header=d.header))
    df = drop_total_rows(df, df.columns[0])
    keep = [c for c in df.columns[1:]
            if not re.search(r"grand\s*total|^total$|^unnamed", str(c), re.I)]
    out = (df.melt(id_vars=[df.columns[0]], value_vars=keep,
                   var_name="location", value_name="qty_on_hand")
             .rename(columns={df.columns[0]: "external_id"}))
    out["qty_on_hand"] = num(out["qty_on_hand"])
    out = out[out.qty_on_hand.fillna(0) != 0]
    out["snapshot_date"] = _snapshot_date(d, p)
    return out, dict(channel=p["channel"], channel_type="retail_store", grain="snapshot",
                     id_type=p.get("id_type", "sku_code"), kind="inventory")


def a_store_blocks(d: Detected) -> tuple[pd.DataFrame, dict]:
    """Two header rows, then repeating N-column blocks per store (J&J sale report)."""
    p = d.profile
    raw = read_csv(d.path, header=None)
    store_row, sub_row = raw.iloc[0], raw.iloc[1]
    body = raw.iloc[2:].reset_index(drop=True)

    start, end = dates_in_name(d.path.name)
    if pd.notna(start) and p.get("period_days"):
        start = start - pd.Timedelta(days=p["period_days"] - 1)

    want = p.get("measure_label", "sold")
    blocks, cur = [], None
    for c in range(raw.shape[1]):                     # find each store's column span
        label = str(store_row.iloc[c]).strip()
        if label and label.lower() not in ("nan", "") \
                and not re.search(r"grand\s*total|row labels", label, re.I):
            cur = label
        if cur and want in str(sub_row.iloc[c]).strip().lower():
            blk = pd.DataFrame({"external_id": body.iloc[:, 0],
                                "location": cur,
                                "qty_sold": num(body.iloc[:, c])})
            blocks.append(blk)

    out = pd.concat(blocks, ignore_index=True) if blocks else pd.DataFrame(
        columns=["external_id", "location", "qty_sold"])
    out = drop_total_rows(out, "external_id")
    out = out[out.qty_sold.fillna(0) != 0]
    out["period_start"], out["period_end"] = start, end
    return out, dict(channel=p["channel"], channel_type="retail_store",
                     flow="sell_out", grain="period",
                     id_type=p.get("id_type", "sku_code"), kind="sales")


# Recognised but deliberately not loaded as facts (product dimensions).
KNOWN_UNUSED = {"uniware_item_master"}

ADAPTERS = {
    "amazon_vc_sales":     a_amazon_sales,
    "amazon_vc_inventory": a_amazon_inv,
    "uniware_tally_gst":   a_uniware_gst,
    "uniware_returns":     a_uniware_returns,
    "uniware_inventory":   a_uniware_inv,
    "uniware_purchase_orders": a_uniware_po,
    "store_flat":          a_store_flat,
    "store_flat_soh":      a_store_flat,
    "store_matrix_soh":    a_store_matrix,
    "store_blocks":        a_store_blocks,
}

# ===========================================================================
# CONFIG 3 -- STORE PROFILES.  *** ADD A NEW STORE HERE. ***
#
#   match        regex against the FILENAME (omit to match any)
#   sheet        regex against the Excel SHEET name (xlsx only)
#   probe        2+ header words used to locate the real header row
#   kind         store_flat | store_flat_soh | store_matrix_soh | store_blocks
#   *_col        source column names -> canonical measures
#   id_type      sku_code (size level) or style_code (no size -> needs allocation)
# ===========================================================================

STORE_PROFILES = [
    # --- All Sports: clean row-per-invoice-line export ---
    dict(kind="store_flat", channel="All Sports", file_type="xlsx",
         match=r"all\s*sports", sheet=r"sale", probe=["principal code", "invoice"],
         id_col="Principal Code", qty_col="Total Sales Qty",
         gross_col="Invoice MRP Value", net_col="Invoice Basic Value",
         disc_col="Invoice Discount Value", date_col="Invoice Date",
         grain="day", flow="sell_out", location="All Sports"),
    dict(kind="store_flat_soh", channel="All Sports", file_type="xlsx",
         match=r"all\s*sports", sheet=r"soh|stock", probe=["principal code", "stock on hand"],
         id_col="Principal Code", soh_col="Total Stock On Hand",
         gross_col="Total MRP Value", location="All Sports"),

    # --- INCS: monthly, STYLE level (no size) ---
    dict(kind="store_flat", channel="INCS", file_type="xlsx",
         match=r"incs", sheet=r".*", probe=["item name", "qty"],
         id_col="Item Name", qty_col="Qty", net_col="Total",
         location_col="INCS", month_col="Month",   # year inferred, not hardcoded
         grain="month", flow="sell_out", id_type="style_code"),

    # --- Jack & Jill: pivot exports ---
    dict(kind="store_blocks", channel="Jack & Jill", file_type="csv",
         match=r"salereport", probe=[], measure_label="sold qty",
         period_days=7, id_type="sku_code"),
    dict(kind="store_matrix_soh", channel="Jack & Jill", file_type="csv",
         match=r"sohreport", probe=[], header=1, id_type="sku_code"),
]

for _p in STORE_PROFILES:                      # store_matrix_soh needs header row 1
    _p.setdefault("header", 0)


# ===========================================================================
# Pipeline
# ===========================================================================

def load_master(path: Path) -> pd.DataFrame:
    """Master sheet has a decorative first row; the real header is row 2."""
    for hdr in (1, 0, 2):
        df = clean_cols(pd.read_excel(path, header=hdr))
        if "Sku Code" in df.columns:
            df["Sku Code"] = norm_id(df["Sku Code"]).values
            if "Item Name" in df.columns:
                df["Item Name"] = norm_id(df["Item Name"]).values
            df = df[df["Sku Code"].notna()].reset_index(drop=True)
            dupes = df.loc[df["Sku Code"].duplicated(keep=False), "Sku Code"].unique()
            if len(dupes):
                print(f"  !! master sheet has {len(dupes)} duplicated Sku Code(s), "
                      f"keeping the first of each: {list(dupes)[:5]}")
                df = df.drop_duplicates("Sku Code", keep="first").reset_index(drop=True)
            return df
    raise SystemExit(f"Could not find a 'Sku Code' column in {path.name}")


def run(input_dir: Path, output_dir: Path, dry_run=False, facility=None):
    output_dir.mkdir(parents=True, exist_ok=True)
    master_path, detected = detect(input_dir, STORE_PROFILES)
    detected, superseded = dedupe(detected, facility)

    print(f"\nScanned {input_dir}"
          + (f"   [facility scope: {facility}]" if facility else ""))
    print(f"  master mapping : {master_path.name if master_path else '*** NOT FOUND ***'}")
    for d in detected:
        tag = f"{d.path.name}" + (f" [{d.sheet}]" if d.sheet else "")
        print(f"  {d.kind:20s} {tag}" + (f"   ({d.note})" if d.note else ""))

    for sup in superseded:
        print(f"  {'superseded':20s} {sup['file']}   ({sup['reason']})")

    if dry_run:
        return
    if not master_path:
        raise SystemExit("\nNo master mapping table found. It must be an .xlsx containing a "
                         "'Sku Code' column alongside platform id columns (ASIN / FSN / Myntra).")

    master = load_master(master_path)
    resolver = Resolver(master)
    print(f"\nMaster: {len(master):,} SKUs · {master['Item Name'].nunique():,} styles · "
          f"{len(resolver.bridge):,} identifier rows")

    sales, inventory, purchases, audit, skipped = [], [], [], [], []
    facility_drops: list[dict] = []

    for d in detected:
        fn = ADAPTERS.get(d.kind)
        if fn is None:
            reason = ("dimension file, not a fact table - not loaded by design"
                      if d.kind in KNOWN_UNUSED else
                      (d.note or f"no adapter for kind '{d.kind}'"))
            skipped.append({"file": d.path.name, "sheet": d.sheet or "",
                            "kind": d.kind, "reason": reason})
            continue
        try:
            df, meta = fn(d)
        except Exception as e:
            skipped.append({"file": d.path.name, "sheet": d.sheet or "",
                            "reason": f"{type(e).__name__}: {e}"})
            print(f"  !! {d.path.name}: {type(e).__name__}: {e}")
            continue

        df["source_file"] = d.path.name + (f"::{d.sheet}" if d.sheet else "")
        df["source_type"] = d.kind

        # channel + flow: from the row's ledger where present, else from the adapter
        lc = meta.get("ledger_col")
        if lc and lc in df.columns:
            key = df[lc].astype(str).str.strip().str.upper()
            trip = key.map(lambda v: CHANNEL_MAP.get(v))
            df["channel"] = [t[0] if t else raw for t, raw in zip(trip, df[lc])]
            df["channel_type"] = [t[1] if t else "unknown" for t in trip]
            df["flow"] = [t[2] if t else "unknown" for t in trip]
            df.drop(columns=[lc], inplace=True)
        else:
            df["channel"] = meta.get("channel", "Unknown")
            df["channel_type"] = meta.get("channel_type", "unknown")
            df["flow"] = meta.get("flow")

        if "location" not in df.columns:
            df["location"] = meta.get("location", df["channel"])
        df["grain"] = meta.get("grain")

        # Facility scope. Applied to whichever column actually carries it, and
        # always counted so a filtered-out facility is visible in the report.
        dropped_rows = 0
        if facility:
            fcol = next((c for c in ("facility", "Facility") if c in df.columns), None)
            if fcol is None and meta["kind"] == "inventory" \
                    and d.kind == "uniware_inventory":
                fcol = "location"
            if fcol is not None:
                keep_mask = norm_id(df[fcol]).eq(facility.upper()) | df[fcol].isna()
                dropped_rows = int((~keep_mask).sum())
                if dropped_rows:
                    other = sorted(set(norm_id(df.loc[~keep_mask, fcol]).dropna()))
                    facility_drops.append({
                        "source_file": d.path.name, "kind": d.kind,
                        "column": fcol, "rows_dropped": dropped_rows,
                        "facilities_found": ", ".join(other[:6]),
                        "scope": facility})
                    df = df[keep_mask]

        if meta.get("id_type") and "external_id" in df.columns:
            df = resolver.resolve(df, meta["id_type"], df["source_file"].iloc[0]
                                  if len(df) else d.path.name)
        else:
            for c in ("external_id", "external_id_type", "matched_via",
                      "master_sku", "style_code"):
                df[c] = df.get(c)

        target, cols = {"inventory": (inventory, INV_COLS),
                        "purchase": (purchases, PO_COLS)}.get(
                            meta["kind"], (sales, SALES_COLS))
        for c in cols:
            if c not in df.columns:
                df[c] = None
        target.append(df[cols])

        matched = int(df["matched_via"].notna().sum())
        audit.append({
            "source_file": df["source_file"].iloc[0] if len(df) else d.path.name,
            "kind": d.kind, "rows": len(df),
            "matched": matched, "unmatched": len(df) - matched,
            "match_pct": round(matched / len(df) * 100, 1) if len(df) else None,
            "qty": float(num(df.get("qty_sold",
                            df.get("qty_on_hand", df.get("qty_ordered")))).sum()),
            "value": float(num(df.get("net_value",
                              df.get("value_on_hand", df.get("total")))).sum()),
            "rows_other_facility": dropped_rows,
        })
        print(f"  {d.kind:20s} {len(df):6,} rows  {matched:6,} matched")

    if not sales and not inventory and not purchases:
        raise SystemExit("Nothing could be parsed. Run with --dry-run to see what was detected.")

    fact_sales = (pd.concat(sales, ignore_index=True) if sales
                  else pd.DataFrame(columns=SALES_COLS))
    fact_inv = (pd.concat(inventory, ignore_index=True) if inventory
                else pd.DataFrame(columns=INV_COLS))
    fact_po = (pd.concat(purchases, ignore_index=True) if purchases
               else pd.DataFrame(columns=PO_COLS))
    audit_df, skipped_df = pd.DataFrame(audit), pd.DataFrame(skipped)

    for c in ("qty_sold", "qty_ordered", "qty_returned", "gross_value",
              "net_value", "discount_value", "returns_value"):
        fact_sales[c] = num(fact_sales[c])
    for c in ("qty_on_hand", "value_on_hand"):
        fact_inv[c] = num(fact_inv[c])
    for c in ("period_start", "period_end"):
        fact_sales[c] = pd.to_datetime(fact_sales[c], errors="coerce")
    fact_inv["snapshot_date"] = pd.to_datetime(fact_inv["snapshot_date"], errors="coerce")
    for c in ("qty_ordered", "qty_received", "qty_pending", "qty_rejected",
              "unit_price", "sub_total", "total", "ageing_days"):
        fact_po[c] = num(fact_po[c])
    for c in ("created_date", "approved_date", "delivery_date"):
        fact_po[c] = pd.to_datetime(fact_po[c], errors="coerce")

    # ---- value basis -------------------------------------------------------
    # Never invent a number when the row already carries one. Amazon's Shipped
    # COGS arrives as 0.00 while Ordered Revenue is populated; some stores send
    # units with no money at all.
    mrp = {}
    for d in detected:
        if d.kind == "uniware_inventory":
            src = clean_cols(read_csv(d.path))
            mrp = dict(zip(norm_id(src["Item SkuCode"]), money(src.get("MRP"))))
            break

    fact_sales["value_basis"] = "as_reported"
    blank = fact_sales.net_value.isna() | fact_sales.net_value.eq(0)
    use_gross = blank & fact_sales.gross_value.fillna(0).ne(0)
    fact_sales.loc[use_gross, "net_value"] = fact_sales.loc[use_gross, "gross_value"]
    fact_sales.loc[use_gross, "value_basis"] = "gross_fallback"

    if mrp:
        fact_sales["_mrp"] = fact_sales.master_sku.map(mrp)
        still = fact_sales.net_value.isna() | fact_sales.net_value.eq(0)
        imp = still & fact_sales._mrp.notna() & fact_sales.qty_sold.notna()
        fact_sales.loc[imp, "net_value"] = fact_sales.loc[imp, "qty_sold"] * fact_sales.loc[imp, "_mrp"]
        fact_sales.loc[imp, "value_basis"] = "imputed_mrp"
        fact_sales.drop(columns=["_mrp"], inplace=True)

    # ---- reports ----------------------------------------------------------
    mapped = fact_sales[fact_sales.master_sku.notna()
                        & fact_sales.flow.isin(["sell_in", "sell_out"])]

    channel = (mapped.groupby(["flow", "channel_type", "channel"], dropna=False)
               .agg(skus=("master_sku", "nunique"), units=("qty_sold", "sum"),
                    net_value=("net_value", "sum"), gross_value=("gross_value", "sum"))
               .reset_index().sort_values(["flow", "net_value"], ascending=[True, False]))

    store = (mapped[mapped.channel_type == "retail_store"]
             .pivot_table(index=["channel", "location"], columns="flow",
                          values=["qty_sold", "net_value"], aggfunc="sum"))
    store.columns = [f"{a}_{b}" for a, b in store.columns]
    store = store.reset_index()
    soh_loc = (fact_inv[fact_inv.channel_type == "retail_store"]
               .groupby(["channel", "location"])["qty_on_hand"].sum())
    store["closing_stock"] = store.set_index(["channel", "location"]).index.map(soh_loc)
    for c in ("qty_sold_sell_out", "qty_sold_sell_in", "net_value_sell_out", "net_value_sell_in"):
        store[c] = store.get(c, 0.0)
    store = store.sort_values("qty_sold_sell_out", ascending=False)

    sku = mapped.pivot_table(index="master_sku", columns="flow",
                             values=["qty_sold", "net_value"], aggfunc="sum").fillna(0)
    sku.columns = [f"{a}_{b}" for a, b in sku.columns]
    sku = sku.reset_index()
    for c in ("qty_sold_sell_in", "qty_sold_sell_out", "net_value_sell_in", "net_value_sell_out"):
        sku[c] = sku.get(c, 0.0)
    sku["style_code"] = sku.master_sku.map(master.set_index("Sku Code")["Item Name"])
    sku["stock_on_hand"] = sku.master_sku.map(
        fact_inv.groupby("master_sku")["qty_on_hand"].sum()).fillna(0)
    denom = (sku.qty_sold_sell_out + sku.stock_on_hand).replace(0, float("nan"))
    sku["sell_through_pct"] = (sku.qty_sold_sell_out / denom * 100).astype(float).round(1)
    sku["channels"] = sku.master_sku.map(mapped.groupby("master_sku")["channel"].nunique())
    sku = sku.sort_values("net_value_sell_out", ascending=False)
    sku = sku[["master_sku", "style_code", "channels", "qty_sold_sell_in", "qty_sold_sell_out",
               "net_value_sell_in", "net_value_sell_out", "stock_on_hand", "sell_through_pct"]]

    # ---- purchase orders: inbound supply, never mixed with sales ----------
    OPEN = ~fact_po.po_status.astype(str).str.upper().isin(
        ["COMPLETE", "COMPLETED", "CANCELLED", "CANCELED", "CLOSED"])
    po_open = fact_po[OPEN & fact_po.qty_pending.fillna(0).gt(0)]

    po_vendor = (fact_po.groupby(["vendor", "po_status"], dropna=False)
                 .agg(pos=("po_code", "nunique"), lines=("po_code", "size"),
                      ordered=("qty_ordered", "sum"), received=("qty_received", "sum"),
                      pending=("qty_pending", "sum"), rejected=("qty_rejected", "sum"),
                      value=("total", "sum"),
                      avg_ageing_days=("ageing_days", "mean"))
                 .reset_index().sort_values("pending", ascending=False))
    if len(po_vendor):
        po_vendor["avg_ageing_days"] = po_vendor.avg_ageing_days.round(1)
        po_vendor["fill_rate_pct"] = (po_vendor.received /
                                      po_vendor.ordered.replace(0, float("nan"))
                                      * 100).round(1)

    # Inbound cover: what is on order against what is in stock and what is
    # actually selling. This is the join that makes the PO file worth loading.
    po_sku = (fact_po[fact_po.master_sku.notna()].groupby("master_sku")
              .agg(on_order=("qty_pending", "sum"), ordered_total=("qty_ordered", "sum"),
                   received_total=("qty_received", "sum"), po_value=("total", "sum"))
              .reset_index())
    if len(po_sku):
        po_sku["style_code"] = po_sku.master_sku.map(master.set_index("Sku Code")["Item Name"])
        po_sku["stock_on_hand"] = po_sku.master_sku.map(
            fact_inv.groupby("master_sku")["qty_on_hand"].sum()).fillna(0)
        so = mapped[mapped.flow == "sell_out"].groupby("master_sku")["qty_sold"].sum()
        po_sku["sell_out_units"] = po_sku.master_sku.map(so).fillna(0)
        po_sku["total_supply"] = po_sku.stock_on_hand + po_sku.on_order
        po_sku = po_sku.sort_values("on_order", ascending=False)
        po_sku = po_sku[["master_sku", "style_code", "on_order", "stock_on_hand",
                         "total_supply", "sell_out_units", "ordered_total",
                         "received_total", "po_value"]]

    inv_pos = (fact_inv[fact_inv.master_sku.notna()]
               .groupby(["channel", "channel_type", "location"], dropna=False)
               .agg(skus=("master_sku", "nunique"), units=("qty_on_hand", "sum"),
                    value=("value_on_hand", "sum"))
               .reset_index().sort_values("units", ascending=False))

    coverage = (fact_sales.groupby(["source_file", "channel"], dropna=False)
                .agg(grain=("grain", "first"), flow=("flow", "first"),
                     period_start=("period_start", "min"),
                     period_end=("period_end", "max"), rows=("qty_sold", "size"))
                .reset_index())
    snap_cov = (fact_inv.groupby(["source_file", "channel"], dropna=False)
                .agg(snapshot_date=("snapshot_date", "max"), rows=("qty_on_hand", "size"))
                .reset_index())

    returns = (fact_sales[fact_sales.qty_returned.fillna(0) != 0]
               .groupby(["channel", "channel_type"], dropna=False)
               .agg(units_returned=("qty_returned", "sum"),
                    returns_value=("returns_value", "sum"),
                    sku_level=("master_sku", lambda s: s.notna().any()))
               .reset_index().sort_values("units_returned", ascending=False))

    basis = (fact_sales.groupby("value_basis")
             .agg(rows=("net_value", "size"), value=("net_value", "sum"))
             .reset_index())

    unknown_ledgers = sorted(set(fact_sales.loc[fact_sales.flow == "unknown", "channel"]
                                 .dropna().astype(str)))

    summary = pd.DataFrame([
        ("Sell-out units (end customers)", mapped.loc[mapped.flow == "sell_out", "qty_sold"].sum()),
        ("Sell-out value", mapped.loc[mapped.flow == "sell_out", "net_value"].sum()),
        ("Sell-in units (invoiced to channels)", mapped.loc[mapped.flow == "sell_in", "qty_sold"].sum()),
        ("Sell-in value", mapped.loc[mapped.flow == "sell_in", "net_value"].sum()),
        ("Channels", mapped.channel.nunique()),
        ("Locations", mapped.location.nunique()),
        ("SKUs with activity", mapped.master_sku.nunique()),
        ("Stock on hand (units)", fact_inv.qty_on_hand.sum()),
        ("Units on order (open POs)", po_open.qty_pending.sum() if len(fact_po) else 0),
        ("Open PO value", po_open.total.sum() if len(fact_po) else 0),
        ("Identifier match rate %", round(audit_df.matched.sum() /
                                          max(audit_df.rows.sum(), 1) * 100, 1)),
        ("Unmapped identifiers", len(resolver.exception_frame())),
        ("Files skipped", len(skipped_df)),
        ("Unmapped channel ledgers", len(unknown_ledgers)),
        ("Files superseded (deduped)", len(superseded)),
        ("Facility scope", facility or "all facilities"),
        ("Rows dropped (other facility)",
         sum(x["rows_dropped"] for x in facility_drops)),
    ], columns=["metric", "value"])

    # ---- write -------------------------------------------------------------
    xlsx = safe_excel_path(output_dir / "Omnichannel_Report.xlsx")
    with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
        summary.to_excel(w, sheet_name="00 Summary", index=False)
        channel.to_excel(w, sheet_name="01 Channel", index=False)
        store.to_excel(w, sheet_name="02 Stores", index=False)
        sku.to_excel(w, sheet_name="03 SKU Performance", index=False)
        inv_pos.to_excel(w, sheet_name="04 Inventory", index=False)
        if len(fact_po):
            po_vendor.to_excel(w, sheet_name="04a PO by Vendor", index=False)
            po_sku.to_excel(w, sheet_name="04b Inbound by SKU", index=False)
            po_open.to_excel(w, sheet_name="04c Open PO Lines", index=False)
        returns.to_excel(w, sheet_name="05 Returns", index=False)
        resolver.exception_frame().to_excel(w, sheet_name="06 Exceptions", index=False)
        audit_df.to_excel(w, sheet_name="07 Load Audit", index=False)
        coverage.to_excel(w, sheet_name="08 Period Coverage", index=False)
        snap_cov.to_excel(w, sheet_name="09 Snapshot Coverage", index=False)
        basis.to_excel(w, sheet_name="10 Value Basis", index=False)
        if len(skipped_df):
            skipped_df.to_excel(w, sheet_name="11 Skipped Files", index=False)
        if superseded:
            pd.DataFrame(superseded).to_excel(w, sheet_name="12 Superseded Files",
                                              index=False)
        if facility_drops:
            pd.DataFrame(facility_drops).to_excel(w, sheet_name="13 Facility Filter",
                                                  index=False)
        fact_sales.head(100_000).to_excel(w, sheet_name="90 Fact Sales", index=False)
        fact_inv.head(100_000).to_excel(w, sheet_name="91 Fact Inventory", index=False)
        if len(fact_po):
            fact_po.head(100_000).to_excel(w, sheet_name="92 Fact Purchase", index=False)

    safe_to_csv(fact_sales, output_dir / "fact_sales.csv")
    safe_to_csv(fact_inv, output_dir / "fact_inventory.csv")
    if len(fact_po):
        safe_to_csv(fact_po, output_dir / "fact_purchase.csv")
    safe_to_csv(resolver.exception_frame(), output_dir / "exceptions.csv")
    write_dashboard(output_dir / "dashboard.html", summary, channel, store, sku,
                    audit_df, coverage, resolver.exception_frame(), basis, returns,
                    po_vendor if len(fact_po) else pd.DataFrame(),
                    po_sku if len(fact_po) else pd.DataFrame())

    # ---- console -----------------------------------------------------------
    print("\n" + "=" * 74)
    print(summary.to_string(index=False))
    print("=" * 74)
    print("\nCHANNEL (sell-in and sell-out are separate measures, never summed)")
    print(channel.to_string(index=False))
    if len(fact_po):
        print("\nPURCHASE ORDERS (inbound supply -- a third flow, not sales)")
        print(po_vendor.head(12).to_string(index=False))
    if unknown_ledgers:
        print(f"\n!! {len(unknown_ledgers)} channel ledger(s) not in CHANNEL_MAP -> "
              f"flow='unknown', excluded from reports: {unknown_ledgers}")
    if facility_drops:
        tot = sum(x["rows_dropped"] for x in facility_drops)
        print(f"\n!! facility scope {facility}: dropped {tot:,} row(s) belonging to "
              f"other facilities (see the Facility Filter sheet)")
        print(pd.DataFrame(facility_drops)[
            ["source_file", "column", "rows_dropped", "facilities_found"]
        ].to_string(index=False))
    if superseded:
        print(f"\n!! {len(superseded)} file(s) superseded to avoid double counting "
              f"(see the Superseded Files sheet)")
    if len(skipped_df):
        print(f"\n!! {len(skipped_df)} file(s)/sheet(s) skipped:")
        print(skipped_df.to_string(index=False))
    exc = resolver.exception_frame()
    if len(exc):
        print(f"\n!! {len(exc)} unmapped identifiers "
              f"(₹{exc.value.sum():,.0f}) -> add to the master sheet. Top 5:")
        print(exc.head(5).to_string(index=False))

    print("\nWrote:")
    for f in sorted(output_dir.iterdir()):
        if f.is_file():
            print(f"  {f}  ({f.stat().st_size:,} bytes)")


# ===========================================================================
# Dashboard
# ===========================================================================

def write_dashboard(path: Path, summary, channel, store, sku, audit,
                    coverage, exceptions, basis, returns,
                    po_vendor=None, po_sku=None):
    def recs(df, n=None):
        d = df.head(n) if n else df
        return json.loads(d.to_json(orient="records", date_format="iso"))

    ch = channel.pivot_table(index="channel", columns="flow",
                             values=["units", "net_value"], aggfunc="sum").fillna(0)
    ch.columns = [f"{a}_{b}" for a, b in ch.columns]
    ch = ch.reset_index()
    for c in ("net_value_sell_in", "net_value_sell_out", "units_sell_in", "units_sell_out"):
        ch[c] = ch.get(c, 0.0)
    ch = ch.sort_values("net_value_sell_in", ascending=False)

    payload = json.dumps({
        "summary": recs(summary), "channel": recs(ch), "store": recs(store),
        "sku": recs(sku, 40), "audit": recs(audit), "coverage": recs(coverage),
        "exceptions": recs(exceptions, 25), "basis": recs(basis), "returns": recs(returns),
        "po_vendor": recs(po_vendor if po_vendor is not None else pd.DataFrame(), 20),
        "po_sku": recs(po_sku if po_sku is not None else pd.DataFrame(), 25),
    }, default=str)

    path.write_text(_HTML.replace("__PAYLOAD__", payload), encoding="utf-8")


_HTML = r"""<!DOCTYPE html><html lang="en" data-theme="light"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Omnichannel Reconciliation</title><style>
:root{--surface:#fcfcfb;--plane:#f9f9f7;--ink:#0b0b0b;--ink2:#52514e;--muted:#898781;
 --grid:#e1e0d9;--axis:#c3c2b7;--bd:rgba(11,11,11,.10);--s1:#2a78d6;--s2:#eb6834;--warn:#fab219;--ok:#006300}
:root[data-theme=dark],@media (prefers-color-scheme:dark){:root:where(:not([data-theme=light])){
 --surface:#1a1a19;--plane:#0d0d0d;--ink:#fff;--ink2:#c3c2b7;--muted:#898781;--grid:#2c2c2a;
 --axis:#383835;--bd:rgba(255,255,255,.10);--s1:#3987e5;--s2:#d95926;--ok:#0ca30c}}
:root[data-theme=dark]{--surface:#1a1a19;--plane:#0d0d0d;--ink:#fff;--ink2:#c3c2b7;--grid:#2c2c2a;
 --axis:#383835;--bd:rgba(255,255,255,.10);--s1:#3987e5;--s2:#d95926;--ok:#0ca30c}
*{box-sizing:border-box}body{margin:0;background:var(--plane);color:var(--ink);
 font:14px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif}
.w{max-width:1200px;margin:0 auto;padding:30px 22px 70px}
header{display:flex;justify-content:space-between;align-items:flex-start;gap:14px;flex-wrap:wrap}
h1{font-size:21px;font-weight:650;margin:0 0 4px}.sub{color:var(--ink2);font-size:13px;max-width:70ch}
h2{font-size:15px;font-weight:600;margin:34px 0 4px}
.note{color:var(--muted);font-size:12.5px;margin:0 0 13px;max-width:82ch}
.btn{background:var(--surface);border:1px solid var(--bd);border-radius:8px;padding:7px 13px;
 font:inherit;font-size:13px;color:var(--ink2);cursor:pointer}
.card{background:var(--surface);border:1px solid var(--bd);border-radius:12px;padding:19px}
.g{display:grid;gap:13px}.g2{grid-template-columns:1fr 1fr}.g4{grid-template-columns:repeat(5,1fr)}
@media(max-width:880px){.g4{grid-template-columns:repeat(2,1fr)}.g2{grid-template-columns:1fr}}
.lb{color:var(--ink2);font-size:12.5px}.hv{font-size:46px;font-weight:650;letter-spacing:-.02em;margin-top:3px}
.tv{font-size:25px;font-weight:600;margin-top:3px}.cp{color:var(--muted);font-size:12px;margin-top:4px}
.lg{display:flex;gap:17px;margin:0 0 13px;font-size:12.5px;color:var(--ink2)}
.sw{width:11px;height:11px;border-radius:3px;display:inline-block;margin-right:6px;vertical-align:-1px}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th{text-align:left;color:var(--muted);font-weight:500;padding:7px 9px;border-bottom:1px solid var(--grid);
 white-space:nowrap;font-size:11px;text-transform:uppercase;letter-spacing:.04em}
td{padding:7px 9px;border-bottom:1px solid var(--grid);color:var(--ink2)}
tr:last-child td{border-bottom:none}td.n{text-align:right;font-variant-numeric:tabular-nums;color:var(--ink)}
td.k{color:var(--ink);font-weight:500}.sc{overflow-x:auto}
.pill{display:inline-block;padding:2px 8px;border-radius:99px;font-size:11px;border:1px solid var(--bd)}
svg{display:block;overflow:visible}.tk{fill:var(--muted);font-size:11px}
.al{fill:var(--ink);font-size:12px}.vl{fill:var(--ink2);font-size:11px;font-variant-numeric:tabular-nums}
.tip{position:fixed;pointer-events:none;opacity:0;transition:opacity .1s;background:var(--surface);
 border:1px solid var(--bd);border-radius:8px;padding:9px 11px;font-size:12px;z-index:9;min-width:165px;
 box-shadow:0 6px 20px rgba(0,0,0,.15)}.tip b{color:var(--ink);display:block;margin-bottom:5px}
.tip .r{display:flex;justify-content:space-between;gap:15px;color:var(--ink2)}
.tip .r span:last-child{color:var(--ink);font-variant-numeric:tabular-nums}
.bar{cursor:pointer}.bar:hover{opacity:.83}
.flag{border-left:2px solid var(--warn);padding-left:12px;margin:14px 0;color:var(--ink2);font-size:12.5px}
footer{margin-top:40px;padding-top:16px;border-top:1px solid var(--grid);color:var(--muted);font-size:12px}
</style></head><body><div class="w">
<header><div><h1>Omnichannel reconciliation</h1><div class="sub">Every channel and store report resolved to
one master SKU, split into <b>sell-in</b> (invoiced by the brand) and <b>sell-out</b> (bought by an end
customer). The two are never added together.</div></div><button class="btn" id="tt">Dark</button></header>
<div class="g g2" style="margin-top:20px">
<div class="card"><div class="lb">Sell-out units — end customers</div><div class="hv" id="h1">—</div><div class="cp" id="h1c"></div></div>
<div class="card"><div class="lb">Sell-in units — invoiced to channels</div><div class="hv" id="h2">—</div><div class="cp" id="h2c"></div></div></div>
<div class="g g4" style="margin-top:13px">
<div class="card"><div class="lb">Channels</div><div class="tv" id="t1">—</div><div class="cp" id="t1c"></div></div>
<div class="card"><div class="lb">Active SKUs</div><div class="tv" id="t2">—</div><div class="cp">resolved to master</div></div>
<div class="card"><div class="lb">Stock on hand</div><div class="tv" id="t3">—</div><div class="cp">all locations</div></div>
<div class="card"><div class="lb">On order</div><div class="tv" id="t5">—</div><div class="cp" id="t5c"></div></div>
<div class="card"><div class="lb">Match rate</div><div class="tv" id="t4">—</div><div class="cp" id="t4c"></div></div></div>
<h2>Channel performance</h2><p class="note">A channel on both sides is expected, not double counting:
sell-in is what you invoiced it, sell-out is what its customers bought. Sell-in only means no sell-out
report has been collected for that channel yet.</p>
<div class="card"><div class="lg"><span><span class="sw" style="background:var(--s1)"></span>Sell-in</span>
<span><span class="sw" style="background:var(--s2)"></span>Sell-out</span></div><div id="ch"></div></div>
<h2>Store performance</h2><div class="card sc"><table id="tStore"></table></div>
<h2>Top SKUs</h2><p class="note">Sell-through = sell-out ÷ (sell-out + stock on hand).</p>
<div class="card sc"><table id="tSku"></table></div>
<div id="poWrap" style="display:none">
<h2>Inbound supply — purchase orders</h2><p class="note">A third flow. Purchases are inbound
from vendors and are never mixed into sell-in or sell-out. Total supply = stock on hand + open PO
quantity; read it against sell-out to see what is over- or under-bought.</p>
<div class="card sc"><table id="tPoV"></table></div>
<p class="note" style="margin-top:16px">Total supply against sell-out, worst-covered first.</p>
<div class="card sc"><table id="tPoS"></table></div></div>
<h2>Data quality</h2><p class="note">The queue to work each cycle. Nothing here is silently dropped.</p>
<div class="g g2"><div class="card sc"><table id="tAudit"></table></div>
<div class="card sc"><table id="tExc"></table></div></div>
<div class="flag" id="flag"></div>
<h2>Period coverage</h2><p class="note">Sources cover different windows at different grains — comparing
across them without saying so is how these numbers get misread.</p>
<div class="card sc"><table id="tCov"></table></div>
<h2>Value basis</h2><p class="note"><code>as_reported</code> came from the file. <code>gross_fallback</code>
used gross because the net column was zero. <code>imputed_mrp</code> is units × MRP where the source sent
no money at all — not invoiced revenue.</p>
<div class="card sc"><table id="tBasis"></table></div>
<footer id="ft"></footer></div><div class="tip" id="tip"></div>
<script id="p" type="application/json">__PAYLOAD__</script><script>
const D=JSON.parse(document.getElementById('p').textContent),tip=document.getElementById('tip');
const S=Object.fromEntries(D.summary.map(r=>[r.metric,r.value]));
const nf=n=>(n==null||isNaN(n))?'—':Math.round(n).toLocaleString('en-IN');
const inr=n=>(n==null||isNaN(n))?'—':'₹'+Math.round(n).toLocaleString('en-IN');
const cmp=n=>n>=1e7?'₹'+(n/1e7).toFixed(2)+' Cr':n>=1e5?'₹'+(n/1e5).toFixed(2)+' L':inr(n);
const tt=document.getElementById('tt');
tt.onclick=()=>{const d=document.documentElement.getAttribute('data-theme')==='dark';
 document.documentElement.setAttribute('data-theme',d?'light':'dark');tt.textContent=d?'Dark':'Light';draw();};
document.getElementById('h1').textContent=nf(S['Sell-out units (end customers)']);
document.getElementById('h1c').textContent=cmp(S['Sell-out value'])+' · consumer demand';
document.getElementById('h2').textContent=nf(S['Sell-in units (invoiced to channels)']);
document.getElementById('h2c').textContent=cmp(S['Sell-in value'])+' · invoiced revenue';
document.getElementById('t1').textContent=nf(S['Channels']);
document.getElementById('t1c').textContent=nf(S['Locations'])+' locations';
document.getElementById('t2').textContent=nf(S['SKUs with activity']);
document.getElementById('t3').textContent=nf(S['Stock on hand (units)']);
document.getElementById('t4').textContent=S['Identifier match rate %']+'%';
document.getElementById('t5').textContent=nf(S['Units on order (open POs)']||0);
document.getElementById('t5c').textContent=cmp(S['Open PO value']||0)+' open';
document.getElementById('t4c').textContent=nf(S['Unmapped identifiers'])+' ids unmapped';
function draw(){const el=document.getElementById('ch'),cs=getComputedStyle(document.documentElement),
 C1=cs.getPropertyValue('--s1').trim(),C2=cs.getPropertyValue('--s2').trim(),d=D.channel;
 if(!d.length){el.innerHTML='<p class="note">No channel data.</p>';return;}
 const W=Math.min(el.clientWidth||980,1140),rh=45,L=170,R=110,T=8,H=d.length*rh+T+26,
 mx=Math.max(...d.map(r=>Math.max(r.net_value_sell_in,r.net_value_sell_out)),1)*1.02,
 x=v=>L+(v/mx)*(W-L-R),st=Math.pow(10,Math.floor(Math.log10(mx))),
 af=v=>v===0?'0':mx>=1e7?'₹'+(v/1e7).toFixed(2)+' Cr':mx>=1e5?'₹'+(v/1e5).toFixed(2)+' L':inr(v);
 let s=`<svg viewBox="0 0 ${W} ${H}" width="100%" height="${H}" role="img" aria-label="Net value by channel, sell-in versus sell-out">`;
 for(let v=0;v<=mx;v+=st/2){s+=`<line x1="${x(v)}" y1="${T}" x2="${x(v)}" y2="${H-26}" stroke="var(--grid)"/>`
  +`<text class="tk" x="${x(v)}" y="${H-9}" text-anchor="middle">${af(v)}</text>`;}
 s+=`<line x1="${L}" y1="${T}" x2="${L}" y2="${H-26}" stroke="var(--axis)"/>`;
 d.forEach((r,i)=>{const yt=T+i*rh+5,bh=15;
  s+=`<text class="al" x="${L-11}" y="${yt+bh+1}" text-anchor="end">${r.channel}</text>`;
  [[r.net_value_sell_in,C1,'Sell-in',r.units_sell_in],[r.net_value_sell_out,C2,'Sell-out',r.units_sell_out]]
  .forEach(([v,c,lb,q],j)=>{if(!(v>0))return;const y=yt+j*(bh+2),w=Math.max(3,x(v)-L);
   s+=`<rect class="bar" x="${L}" y="${y}" width="${w}" height="${bh}" fill="${c}" rx="4"`
    +` data-c="${r.channel}" data-l="${lb}" data-v="${v}" data-q="${q}"/>`
    +`<rect x="${L}" y="${y}" width="${Math.min(4,w)}" height="${bh}" fill="${c}"/>`
    +`<text class="vl" x="${L+w+8}" y="${y+bh-3}">${inr(v)}</text>`;});});
 el.innerHTML=s+'</svg>';
 el.querySelectorAll('.bar').forEach(b=>{b.onmousemove=e=>{tip.innerHTML=`<b>${b.dataset.c} · ${b.dataset.l}</b>`
  +`<div class="r"><span>Net value</span><span>${inr(+b.dataset.v)}</span></div>`
  +`<div class="r"><span>Units</span><span>${nf(+b.dataset.q)}</span></div>`;
  tip.style.opacity=1;tip.style.left=Math.min(e.clientX+14,innerWidth-205)+'px';tip.style.top=(e.clientY+14)+'px';};
  b.onmouseleave=()=>tip.style.opacity=0;});}
function tbl(id,cols,rows){const t=document.getElementById(id);if(!t)return;
 t.innerHTML='<thead><tr>'+cols.map(c=>`<th${c.n?' style="text-align:right"':''}>${c.h}</th>`).join('')
 +'</tr></thead><tbody>'+(rows.length?rows.map(r=>'<tr>'+cols.map(c=>
  `<td class="${c.n?'n':(c.k?'k':'')}">${c.f(r)}</td>`).join('')+'</tr>').join('')
 :`<tr><td colspan="${cols.length}" style="color:var(--muted)">Nothing to show.</td></tr>`)+'</tbody>';}
tbl('tStore',[{h:'Store',k:1,f:r=>r.channel},{h:'Location',f:r=>r.location},
 {h:'Sell-out units',n:1,f:r=>nf(r.qty_sold_sell_out)},{h:'Sell-out value',n:1,f:r=>inr(r.net_value_sell_out)},
 {h:'Sell-in units',n:1,f:r=>nf(r.qty_sold_sell_in)},{h:'Closing stock',n:1,f:r=>nf(r.closing_stock)}],D.store);
tbl('tSku',[{h:'Master SKU',k:1,f:r=>r.master_sku},{h:'Style',f:r=>r.style_code||'—'},
 {h:'Ch',n:1,f:r=>nf(r.channels)},{h:'Sell-in',n:1,f:r=>nf(r.qty_sold_sell_in)},
 {h:'Sell-out',n:1,f:r=>nf(r.qty_sold_sell_out)},{h:'Value out',n:1,f:r=>inr(r.net_value_sell_out)},
 {h:'Stock',n:1,f:r=>nf(r.stock_on_hand)},
 {h:'Sell-through',n:1,f:r=>r.sell_through_pct==null?'—':r.sell_through_pct+'%'}],D.sku);
tbl('tAudit',[{h:'Source file',k:1,f:r=>String(r.source_file).slice(0,34)},
 {h:'Rows',n:1,f:r=>nf(r.rows)},{h:'Matched',n:1,f:r=>nf(r.matched)},
 {h:'Match %',n:1,f:r=>r.match_pct==null?'n/a':r.match_pct+'%'}],D.audit);
tbl('tExc',[{h:'Unmapped id',k:1,f:r=>r.external_id},{h:'Qty',n:1,f:r=>nf(r.qty)},
 {h:'Value',n:1,f:r=>inr(r.value)}],D.exceptions);
tbl('tCov',[{h:'Source',k:1,f:r=>String(r.source_file).slice(0,40)},{h:'Channel',f:r=>r.channel||'—'},
 {h:'Flow',f:r=>`<span class="pill">${r.flow||'—'}</span>`},
 {h:'Grain',f:r=>`<span class="pill">${r.grain||'—'}</span>`},
 {h:'From',f:r=>String(r.period_start||'—').slice(0,10)},{h:'To',f:r=>String(r.period_end||'—').slice(0,10)},
 {h:'Rows',n:1,f:r=>nf(r.rows)}],D.coverage);
tbl('tBasis',[{h:'Basis',k:1,f:r=>`<code>${r.value_basis}</code>`},{h:'Rows',n:1,f:r=>nf(r.rows)},
 {h:'Value',n:1,f:r=>inr(r.value)}],D.basis);
if((D.po_vendor||[]).length||(D.po_sku||[]).length){
 document.getElementById('poWrap').style.display='';
 tbl('tPoV',[{h:'Vendor',k:1,f:r=>String(r.vendor||'—').slice(0,26)},
  {h:'Status',f:r=>`<span class="pill">${r.po_status||'—'}</span>`},
  {h:'POs',n:1,f:r=>nf(r.pos)},{h:'Ordered',n:1,f:r=>nf(r.ordered)},
  {h:'Received',n:1,f:r=>nf(r.received)},{h:'Pending',n:1,f:r=>nf(r.pending)},
  {h:'Fill %',n:1,f:r=>r.fill_rate_pct==null?'—':r.fill_rate_pct+'%'},
  {h:'Value',n:1,f:r=>inr(r.value)}],D.po_vendor||[]);
 tbl('tPoS',[{h:'Master SKU',k:1,f:r=>r.master_sku},
  {h:'On order',n:1,f:r=>nf(r.on_order)},{h:'In stock',n:1,f:r=>nf(r.stock_on_hand)},
  {h:'Total supply',n:1,f:r=>nf(r.total_supply)},
  {h:'Sell-out',n:1,f:r=>nf(r.sell_out_units)},
  {h:'Received',n:1,f:r=>nf(r.received_total)},
  {h:'PO value',n:1,f:r=>inr(r.po_value)}],D.po_sku||[]);}
const chOnly=D.returns.filter(r=>!r.sku_level),ru=chOnly.reduce((a,b)=>a+(b.units_returned||0),0);
document.getElementById('flag').innerHTML='<b>Read before quoting these numbers.</b> '
 +(S['Unmapped identifiers']>0?`${nf(S['Unmapped identifiers'])} identifiers did not resolve and are excluded from every figure above — see the exception queue. `:'')
 +(ru>0?`${nf(ru)} returned units come from a source with no SKU column, so they can only be reported by channel. `:'')
 +(D.basis.some(b=>b.value_basis==='imputed_mrp')?'Some rows are valued at MRP because the source sent units only — flagged <code>imputed_mrp</code>, not invoiced revenue. ':'')
 +'Sell-in and sell-out are separate measures and are never summed.';
document.getElementById('ft').textContent=`${nf(S['Channels'])} channels · ${nf(S['Locations'])} locations · `
 +`${nf(S['SKUs with activity'])} active SKUs · all values in INR`
 +(S['Files skipped']>0?` · ${nf(S['Files skipped'])} file(s) skipped — see the workbook`:'');
draw();addEventListener('resize',draw);
</script></body></html>"""


def main():
    ap = argparse.ArgumentParser(
        description="Reconcile multi-channel sales/inventory reports into one report.")
    ap.add_argument("--input", "-i", required=True, help="folder holding the raw reports")
    ap.add_argument("--output", "-o", default="./report", help="output folder (default ./report)")
    ap.add_argument("--facility", default=None,
                    help="load only this Uniware facility (e.g. DOMIN8). Rows and "
                         "per-facility files belonging to other facilities are "
                         "dropped and counted, never silently ignored.")
    ap.add_argument("--dry-run", action="store_true",
                    help="show which files were detected, then stop")
    a = ap.parse_args()

    src = Path(a.input).expanduser().resolve()
    if not src.is_dir():
        sys.exit(f"Not a folder: {src}")
    run(src, Path(a.output).expanduser().resolve(), a.dry_run, a.facility)


if __name__ == "__main__":
    main()
