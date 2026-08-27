# DOMIN8 — Omnichannel Reporting

Reconciles sales, inventory and purchase data from Uniware, Amazon Vendor
Central and the retail stores into one report, keyed on a master SKU table.

**Live app:** _(add the Streamlit URL once deployed)_

---

## What it does

Three separate flows, never added together:

| Flow | Source | Measures |
|---|---|---|
| `sell_in` | Uniware Tally GST | what the brand invoiced **to** a channel |
| `sell_out` | Amazon VC, store sale reports, D2C | what the **end customer** bought |
| `purchase` | Uniware Purchase Orders | what was ordered **from vendors** |

Uniware's Tally GST report is the invoice ledger and already covers every
channel — including Amazon as `Cocoblu Retail Limited.`. So the store and Amazon
files are not duplicates of it; they measure the other side of the same goods.
Union them into one revenue column and the numbers roughly double.

Every row carries a `flow` column, purchases live in their own table, and the
verifier asserts none of them leak into the others.

## Output

| | |
|---|---|
| **Action list** | REORDER / SIZE BREAK / CLEAR, each with a plain-English reason |
| **Stock vs Sales** | `sku wise` and `Article wise`, matching the team's existing workbook column-for-column |
| **Omnichannel report** | 18 sheets — channel, stores, inventory, POs, returns, exceptions, audit, coverage |
| **Dashboard** | self-contained HTML |
| **Fact tables** | `fact_sales`, `fact_inventory`, `fact_purchase` |

## Files

```
app.py                 Streamlit UI (upload → run → view → download)
run_pipeline.py        CLI orchestrator, for scheduled/local runs
pipeline_config.py     paths, facility scope, windows
reconcile.py           the reconciler — detection, dedup, identity resolution
stock_vs_sales.py      the merchandising workbook
check_reconcile.py     ties output back to the raw files (32 checks)
validate_svs.py        re-checks the movement rules against a hand-built workbook
uniware_exports.py     Uniware API puller
```

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Or CLI, if `reports/input/` exists beside the code:

```bash
python run_pipeline.py            # build from files on disk
python run_pipeline.py --fetch    # pull Uniware first
```

## Deploy

1. Push to GitHub (this repo).
2. [share.streamlit.io](https://share.streamlit.io) → **New app** → pick this
   repo → main file `app.py` → Deploy.
3. App settings → make it private, add the viewers' email addresses.
4. Secrets → paste from `.streamlit/secrets.toml.example` (only needed once the
   Uniware fetch is enabled).

Every `git push` redeploys in about a minute.

### Known limitation of the hosted pilot

The container has **no persistent disk**. Merchandiser reorder notes and the
purchase-order history do not survive a redeploy or a new session. Adding
Supabase (`storage.py`) fixes this and is the intended next step.

## Data

**No client data is in this repo.** See [SAMPLE_DATA.md](SAMPLE_DATA.md).

## Behaviour worth knowing

**Files are matched by content, not filename**, so date-stamped exports work
unrenamed.

**Duplicates are caught before load.** Two pulls of the same report would double
every number — verified: two inventory snapshots took stock from 31,074 to
58,088. Point-in-time sources take one file only; transactional sources take the
newest per report and facility; `__COMBINED` supersedes per-facility splits.
Everything dropped is listed with its reason.

**Facility scope is `DOMIN8`.** Enforced in the pull, at file level, and at row
level — foreign rows are counted and reported, never silently dropped.

**Nothing is silently discarded.** Unresolved identifiers go to an exception
queue and are excluded from every figure. Unrecognised files are listed.

**Values are never invented.** `value_basis` on every sales row is
`as_reported`, `gross_fallback` (net was zero, used gross — Amazon's Shipped
COGS arrives as ₹0.00) or `imputed_mrp` (source sent units only).

**Movement classification** was reverse-engineered from the client's own
workbook and validates at 100% against both its tabs (1,716 SKU rows, 378
article rows):

```
No Movement     sell-through <= 0
Slow Movement   0 < st < Avg × 0.75
Good Movement   Avg × 0.75 <= st <= Avg × 1.25
Fast Movement   st > Avg × 1.25
```

Avg is the expected sell-through for the ageing bucket — 20% at 0–1 month, 40%
at 1–3 months, 46% beyond.

## Verify

```bash
python check_reconcile.py --input reports/input --report reports/output --facility DOMIN8
```

32/32 checks pass on the reference data.
