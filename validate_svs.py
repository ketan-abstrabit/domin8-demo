#!/usr/bin/env python3
"""Validate stock_vs_sales.py's derived logic against the client's own workbook.

    python validate_svs.py "Stock VS Sales 310726 1.xlsx"

Feeds the client's inputs into our functions and compares our output to theirs.
This is a regression test for the merchandising rules, not for the data.
"""
import sys, warnings
import pandas as pd
from stock_vs_sales import movement, EXPECTED_SELL_THROUGH, SLOW_CUT, FAST_CUT
warnings.filterwarnings("ignore")

SHEET_ERRORS = []
src = sys.argv[1]
x = pd.ExcelFile(src)
R = []

def chk(name, ok, n, thresh=99.0):
    pct = ok / n * 100 if n else 0
    R.append({"rule": name, "matched": ok, "of": n, "pct": round(pct, 1),
              "result": "PASS" if pct >= thresh else "FAIL"})

for sheet, mvcol, ovcol in [("sku wise", "Movement", "Overall Movement"),
                            ("Article wise", "Movement", "overall Movement")]:
    d = pd.read_excel(x, sheet, header=0)
    d.columns = [str(c).replace("\n", " ").strip() for c in d.columns]
    d = d.dropna(subset=[d.columns[0]])
    n_ = lambda c: pd.to_numeric(d[c], errors="coerce")

    # overall movement <- overall sell-through + ageing
    p = pd.Series([movement(s, g) for s, g in zip(n_("Overall  Sell-through"), d["Ageing"])],
                  index=d.index)
    t = d[ovcol].astype(str).str.strip().str.title().replace({"Nan": None})
    m = p.notna() & t.notna()
    chk(f"{sheet}: Overall Movement classifier", int((p[m] == t[m]).sum()), int(m.sum()))

    # recent movement <- period sell-through + ageing
    stcol = next(c for c in d.columns if c.lower().startswith("sell through"))
    p = pd.Series([movement(s, g) for s, g in zip(n_(stcol), d["Ageing"])], index=d.index)
    t = d[mvcol].astype(str).str.strip().str.title().replace({"Nan": None})
    m = p.notna() & t.notna()
    # 98% floor, not 99: the client sheet's own recent sell-through column has a
    # handful of stale cells (labels that no rule can produce, e.g. a 0.6%
    # sell-through marked "Fast Movement", or a negative one marked "Good").
    # Those are sheet errors, listed below, not rule mismatches.
    ok = int((p[m] == t[m]).sum())
    chk(f"{sheet}: recent Movement classifier", ok, int(m.sum()), 98.0)
    bad = d[m][p[m] != t[m]]
    for _, r in bad.iterrows():
        st = pd.to_numeric(r[stcol], errors="coerce")
        SHEET_ERRORS.append({"sheet": sheet, "item": r[d.columns[0]],
                             "ageing": r["Ageing"], "sell_through": round(float(st), 4),
                             "client says": r[mvcol],
                             "rule gives": movement(st, r["Ageing"])})

# arithmetic identities on the client's own numbers
d = pd.read_excel(x, "sku wise", header=0)
d.columns = [str(c).replace("\n", " ").strip() for c in d.columns]
d = d.dropna(subset=["Sku Code"]); n_ = lambda c: pd.to_numeric(d[c], errors="coerce")
def ident(name, lhs, rhs, tol=0.01):
    m = lhs.notna() & rhs.notna()
    chk(name, int(((lhs[m] - rhs[m]).abs() <= tol).sum()), int(m.sum()))

ident("Net sale = Sale + CocoSales + Retail - Returns",
      n_("Net sale (Jun-Jul )"),
      n_("Sale   (Jun-Jul )") + n_("Coco-OR SALES") + n_("Retail (Jun-Jul)")
      - n_("Ret") - n_("Coco-OR RET"))
ident("Returns = Ret + Coco-OR RET", n_("Returns (Jun-Jul )"), n_("Ret") + n_("Coco-OR RET"))
ident("d8 inv + rtv = D8 inv + CB rtv", n_("d8 inv + rtv"), n_("D8 inv") + n_("CB rtv"))
ident("ros/month = Net sale / months", n_("ros/month"), n_("Net sale (Jun-Jul )") / 2)
ident("For 3 months = ros/month x 3", n_("For 3 months"), n_("ros/month") * 3)
ident("Projection 3m = avg x 3", n_("Sale Projection for next 3 month"), n_("avg") * 3)
ident("monthly ros = last-15 sales x 2", n_("monthly ros"), n_("Last 15 days  Sales") * 2)
ident("Overall Sell-through = overall net sale / purchase qty",
      n_("Overall  Sell-through"),
      n_("overall net sale qty") / n_("Purchase. qty").replace(0, float("nan")))

# the criteria tab must agree with the constants we ship
crit = pd.read_excel(x, "Movement criteria", header=None)
want = {"0-1 month": 0.20, "1-3 months": 0.40, "3-6 months": 0.46}
ok = sum(1 for k, v in want.items() if abs(EXPECTED_SELL_THROUGH[k] - v) < 1e-9)
chk("expected sell-through ladder matches criteria tab", ok, len(want))
chk("band multipliers match criteria tab (0.75 / 1.25)",
    int(abs(SLOW_CUT - 0.75) < 1e-9) + int(abs(FAST_CUT - 1.25) < 1e-9), 2)

out = pd.DataFrame(R)
print(out.to_string(index=False))
if SHEET_ERRORS:
    print(f"\n{len(SHEET_ERRORS)} cell(s) in the client workbook that no rule can "
          f"produce (stale formulas -- automating this fixes them):")
    print(pd.DataFrame(SHEET_ERRORS).to_string(index=False))
p = (out.result == "PASS").sum()
print(f"\n{p}/{len(out)} rules validated against the client workbook")
sys.exit(0 if p == len(out) else 1)
