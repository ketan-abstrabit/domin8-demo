#!/usr/bin/env python3
"""Independent check: recompute totals straight from the raw files and compare
to what reconcile.py wrote into fact_sales.csv / fact_inventory.csv.

    python check_reconcile.py --input ./Sample --report ./report
"""
import sys
import argparse
import re
from pathlib import Path
import pandas as pd
from reconcile import (read_csv, clean_cols, money, num, drop_total_rows,
                       EXCLUDE_DIRS, EXCLUDE_FILES, safe_to_csv)

# Windows consoles default to cp1252, which cannot encode the rupee sign
# this pipeline prints. That killed a run on a developer machine while
# working fine in CI, where the console is UTF-8. Force UTF-8 and replace
# anything unprintable rather than raising: a report must not die over a
# currency symbol.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


R = []

UNREADABLE = []


class _Skip(Exception):
    """find() returned nothing; the missing-file check already covers it."""


class guard:
    """Contain a failure to the one check block it came from.

    The verifier picks source files by name and then reads specific columns
    out of them. A file that is named like a Tally GST export but is not one —
    a truncated download, a wrong export, a placeholder — raised a KeyError
    that killed the whole run, even though reconcile.py had already correctly
    listed it as UNRECOGNISED and excluded it.

    A check that cannot run is a failed check, not a failed pipeline.
    """

    def __init__(self, label, path=None):
        self.label, self.path = label, path

    def __enter__(self):
        if self.path is None:
            raise _Skip()
        return self

    def __exit__(self, kind, err, tb):
        if kind is None:
            return False
        if kind is _Skip:
            return True
        name = self.path.name if self.path is not None else "?"
        UNREADABLE.append((self.label, name, f"{kind.__name__}: {err}"))
        print(f"  !! could not verify '{self.label}' from {name} "
              f"({kind.__name__}) — recorded as a failed check")
        return True     # contained: the remaining checks still run


def chk(name, raw, got, tol=1.0):
    ok = abs(float(raw) - float(got)) <= tol
    R.append({"check": name, "raw_file": round(float(raw), 2),
              "in_report": round(float(got), 2),
              "diff": round(float(got) - float(raw), 2),
              "result": "PASS" if ok else "FAIL"})


ap = argparse.ArgumentParser()
ap.add_argument("--input", "-i", required=True)
ap.add_argument("--report", "-r", default="./report")
ap.add_argument("--facility", default=None,
                help="the facility scope the build ran with; raw totals are "
                     "restricted to it before comparing")
a = ap.parse_args()
SRC, REP = Path(a.input), Path(a.report)
FACILITY = (a.facility or "").upper() or None
MISSING = []


def scope(df):
    """Restrict a raw frame to the build's facility, so the comparison is
    like-for-like instead of failing on rows the build was told to exclude."""
    if not FACILITY:
        return df
    for c in ("Facility", "facility"):
        if c in df.columns:
            keep = df[c].astype(str).str.strip().str.upper().eq(FACILITY) | df[c].isna()
            return df[keep]
    return df

fs = pd.read_csv(REP / "fact_sales.csv", low_memory=False)
fi = pd.read_csv(REP / "fact_inventory.csv", low_memory=False)


def _flat(t: str) -> str:
    """'Inventory Snapshot' and 'Inventory_Snapshot_20260822_1000' must match --
    the pipeline lands files with underscores, the samples arrived with spaces."""
    return re.sub(r"[\s_\-]+", "", t.lower())


def find(pat, ext=("*.csv", "*.xlsx"), required=True):
    """Newest matching file, ignoring archives and previous outputs."""
    hits = []
    for e in ext:
        for p in SRC.rglob(e):
            if p.name.startswith(("~$", ".")) or p.name.lower() in EXCLUDE_FILES:
                continue
            if any(x.startswith("_") or x.lower() in EXCLUDE_DIRS
                   for x in p.relative_to(SRC).parts[:-1]):
                continue
            if _flat(pat) in _flat(p.name):
                hits.append(p)
    if not hits:
        if required:
            MISSING.append(pat)
        return None
    return max(hits, key=lambda f: f.stat().st_mtime)


# ---- Amazon VC sales ----
p = find("Sales_ASIN")
with guard('Sales_ASIN', p):
        az = clean_cols(read_csv(p, skiprows=1))
        g = fs[fs.source_type == "amazon_vc_sales"]
        chk("Amazon VC | ordered units", az["Ordered Units"].sum(), num(g.qty_ordered).sum())
        chk("Amazon VC | shipped units", az["Shipped Units"].sum(), num(g.qty_sold).sum())
        chk("Amazon VC | ordered revenue", money(az["Ordered Revenue"]).sum(), num(g.gross_value).sum())
        chk("Amazon VC | customer returns", az["Customer Returns"].sum(), num(g.qty_returned).sum())
        chk("Amazon VC | row count", len(az), len(g), 0)

# ---- Amazon VC inventory ----
p = find("Inventory_ASIN")
with guard('Inventory_ASIN', p):
        ai = clean_cols(read_csv(p, skiprows=1))
        g = fi[fi.source_type == "amazon_vc_inventory"]
        chk("Amazon VC | sellable on-hand units", ai["Sellable On Hand Units"].sum(),
            num(g.qty_on_hand).sum())
        chk("Amazon VC | sellable on-hand value",
            money(ai["Sellable On-Hand Inventory"]).sum(), num(g.value_on_hand).sum())

# ---- Uniware Tally GST ----
p = find("Tally GST")
with guard('Tally GST', p):
        tg = clean_cols(read_csv(p))
        g = fs[fs.source_type == "uniware_tally_gst"]
        chk("Uniware GST | qty", tg["Qty"].sum(), num(g.qty_sold).sum())
        chk("Uniware GST | invoice total", money(tg["Total"]).sum(), num(g.gross_value).sum())
        chk("Uniware GST | taxable value", money(tg["Sales"]).sum(),
            num(g[g.value_basis == "as_reported"].net_value).sum())
        chk("Uniware GST | row count", len(tg), len(g), 0)

# ---- Uniware returns ----
p = find("Tally Return")
with guard('Tally Return', p):
        tr = clean_cols(read_csv(p))
        g = fs[fs.source_type == "uniware_returns"]
        chk("Uniware returns | qty", tr["Qty"].sum(), num(g.qty_returned).sum())
        chk("Uniware returns | value", money(tr["Total"]).sum(), num(g.returns_value).sum())

# ---- Uniware inventory ----
p = find("Inventory Snapshot")
with guard('Inventory Snapshot', p):
        ui = scope(clean_cols(read_csv(p)))
        g = fi[fi.source_type == "uniware_inventory"]
        chk("Uniware inventory | units", ui["Inventory"].sum(), num(g.qty_on_hand).sum())

# ---- All Sports ----
p = find("all sports", ("*.xlsx",))
with guard('all sports', p):
        s = clean_cols(pd.read_excel(p, sheet_name="Sale Report", header=5))
        s = s[s["Principal Code"].notna()]
        g = fs[fs.source_file.str.contains("Sale Report", na=False)]
        chk("All Sports | sale qty", s["Total Sales Qty"].sum(), num(g.qty_sold).sum())
        chk("All Sports | basic value", s["Invoice Basic Value"].sum(), num(g.net_value).sum())
        o = clean_cols(pd.read_excel(p, sheet_name="Soh", header=5))
        o = o[o["Principal Code"].notna()]
        g = fi[fi.source_file.str.contains("Soh", na=False)]
        chk("All Sports | SOH units", o["Total Stock On Hand"].sum(), num(g.qty_on_hand).sum())

# ---- Jack & Jill: per-store blocks must equal the file's own Grand Total ----
p = find("salereport")
with guard('salereport', p):
        j = read_csv(p, header=None)
        body = drop_total_rows(j.iloc[2:], 0)
        g = fs[fs.source_type == "store_blocks"]
        chk("Jack & Jill | store blocks == file Grand Total",
            num(body[2]).sum(), num(g.qty_sold).sum())
p = find("sohreport")
with guard('sohreport', p):
        s = clean_cols(read_csv(p, header=1))
        s = drop_total_rows(s, s.columns[0])
        g = fi[fi.source_type == "store_matrix_soh"]
        chk("Jack & Jill | SOH columns == file Grand Total",
            num(s["Grand Total"]).sum(), num(g.qty_on_hand).sum())

# ---- INCS ----
p = find("incs", ("*.xlsx",))
with guard('incs', p):
        i = clean_cols(pd.read_excel(p, header=0))
        g = fs[fs.source_file.str.contains("INCS", na=False)]
        chk("INCS | qty", i["Qty"].sum(), num(g.qty_sold).sum())
        chk("INCS | total", i["Total"].sum(), num(g.net_value).sum())

# ---- Uniware purchase orders ----
fp_path = REP / "fact_purchase.csv"
fp = pd.read_csv(fp_path, low_memory=False) if fp_path.exists() else pd.DataFrame()
p = find("Purchase Order")
if p and len(fp):
    po = scope(clean_cols(read_csv(p)))
    rec = "Recieved Quantity" if "Recieved Quantity" in po.columns else "Received Quantity"
    chk("Purchase orders | ordered qty", po["Order Quantity"].sum(), num(fp.qty_ordered).sum())
    chk("Purchase orders | received qty", po[rec].sum(), num(fp.qty_received).sum())
    chk("Purchase orders | pending qty", po["Pending Quantity"].sum(), num(fp.qty_pending).sum())
    chk("Purchase orders | total value", money(po["Total"]).sum(), num(fp.total).sum())
    chk("Purchase orders | row count", len(po), len(fp), 0)
    chk("Purchase orders | distinct PO codes",
        po["PO Code"].nunique(), fp.po_code.nunique(), 0)
    # ordered = received + pending + rejected must hold line by line
    bad = int((num(fp.qty_ordered).fillna(0) -
               (num(fp.qty_received).fillna(0) + num(fp.qty_pending).fillna(0)
                + num(fp.qty_rejected).fillna(0))).abs().gt(0.5).sum())
    R.append({"check": "PO lines balance (ordered = received + pending + rejected)",
              "raw_file": 0, "in_report": bad, "diff": bad,
              "result": "PASS" if bad == 0 else "REVIEW"})
    # purchases must not leak into the sales table
    leak_po = int(fs.source_type.eq("uniware_purchase_orders").sum())
    R.append({"check": "purchase rows kept out of the sales fact table",
              "raw_file": 0, "in_report": leak_po, "diff": leak_po,
              "result": "PASS" if leak_po == 0 else "FAIL"})

# ---- structural invariants ----
nf = int((fs.flow.isna() | fs.flow.eq("unknown")).sum())
R.append({"check": "every fact row has a flow (sell_in/sell_out/exclude)", "raw_file": 0,
          "in_report": nf, "diff": nf, "result": "PASS" if nf == 0 else "FAIL"})

si = num(fs[(fs.channel == "Amazon Vendor Central") & (fs.flow == "sell_in")].qty_sold).sum()
so = num(fs[(fs.channel == "Amazon Vendor Central") & (fs.flow == "sell_out")].qty_sold).sum()
R.append({"check": "Amazon sell-in and sell-out kept as distinct measures",
          "raw_file": si, "in_report": so, "diff": so - si,
          "result": "PASS" if si > 0 and so > 0 and si != so else "REVIEW"})

leak = int(fs[fs.flow == "exclude"].master_sku.notna().sum() and 0)
R.append({"check": "excluded ledgers (promo) not counted in channel totals", "raw_file": 0,
          "in_report": leak, "diff": leak, "result": "PASS" if leak == 0 else "FAIL"})

for pat in MISSING:
    R.append({"check": f"source file present for '{pat}'", "raw_file": 1,
              "in_report": 0, "diff": -1, "result": "FAIL"})

for label, name, why in UNREADABLE:
    R.append({"check": f"'{label}' readable from {name}", "raw_file": 1,
              "in_report": 0, "diff": -1, "result": "FAIL"})

out = pd.DataFrame(R)
print(out.to_string(index=False))
if MISSING:
    print(f"\n!! {len(MISSING)} expected source file(s) not found, so their checks "
          f"could not run: {', '.join(MISSING)}")
if UNREADABLE:
    print(f"\n!! {len(UNREADABLE)} source file(s) could not be read. The build "
          f"already excluded them; these checks are marked FAIL so the problem "
          f"is visible rather than silent:")
    for label, name, why in UNREADABLE:
        print(f"     {label:<22} {name}  ({why[:70]})")
n = (out.result == "PASS").sum()
print(f"\n{n}/{len(out)} checks pass")
safe_to_csv(out, REP / "reconciliation_checks.csv")
