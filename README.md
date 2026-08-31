# DOMIN8 — Omnichannel Reporting

Reconciles sales, inventory and purchase data from Uniware, Amazon Vendor
Central and the retail stores into one report, keyed on a master SKU table —
then alerts on what it finds.

**The client's interface is a Google Drive folder.** They drop files into
`input/`; the finished workbooks appear in `output/latest/` and a digest lands
in their inbox. No app, no login, no code on their machine.
Setup: **[DRIVE_SETUP.md](DRIVE_SETUP.md)**.

```
Drive input/  ──▶  GitHub Actions  ──▶  Drive output/latest/
                   (every 2h, free)     + email digest
```

`run_drive.py` is the only piece that knows Drive exists. The reconciler and
the merchandising workbook read a folder and write a folder, exactly as they do
locally, so everything stays testable offline.

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
| **Alerts** | `Alerts.xlsx` — 8 alert types, one sheet each, plus what's new |
| **Digest** | `alert_digest.html` — the email, also kept as a file |

## Alerts

Eight alerts, all from numbers the pipeline already computed: out of stock, low
stock, reorder level reached, best sellers, high returns, slow moving, dead
stock, stock ageing.

Three things make them readable rather than noise:

**Digest, not per-SKU.** 421 reorder emails is a spreadsheet with extra steps.
One message per cycle, grouped, top 15 rows each, everything in the attachment.

**State.** Each alert is classified NEW / ONGOING / RESOLVED against the
previous run, so the same 401 out-of-stock SKUs do not shout every fortnight.
The digest leads with what changed.

**Money, not counts.** Dead, slow and ageing stock rank on capital sitting
still. Out-of-stock, low-stock and reorder rank on the sales the shortage costs
per month — ranking those on stock value would sort by a number that is zero by
definition.

Every threshold is in [`alert_rules.yaml`](alert_rules.yaml). Edit, commit,
push; the next run uses it.

## Files

```
run_drive.py           Drive mode — pull, build, publish, alert  ← what CI runs
drive_sync.py          the only Drive-aware module
alerts.py              rules → events → digest
alert_rules.yaml       every threshold, commented
run_pipeline.py        CLI orchestrator, for local runs
pipeline_config.py     paths, facility scope, windows
reconcile.py           the reconciler — detection, dedup, identity resolution
stock_vs_sales.py      the merchandising workbook
check_reconcile.py     ties output back to the raw files (32 checks)
validate_svs.py        re-checks the movement rules against a hand-built workbook
uniware_exports.py     Uniware API puller
app.py                 Streamlit UI — kept for internal debugging, not the client
tests/                 offline end-to-end test of the whole Drive cycle
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

See **[DRIVE_SETUP.md](DRIVE_SETUP.md)** — a Shared Drive, a service account,
six GitHub secrets. About 30 minutes, no cost.

Every `git push` updates the client's pipeline. They do nothing.

### Verify it without touching Drive

```bash
python tests/test_drive_cycle.py
```

Runs the full pull → build → publish cycle against an in-memory fake Drive:
change detection, in-place overwrite, archive retention, deleted files, alert
state transitions, and both failure paths. 26 checks, about two minutes,
offline.

### State between runs

`_state/` in the Drive folder holds the input fingerprint, the accumulated
purchase-order history and the open-alert set. That is what lets the pipeline
skip unchanged inputs, keep improving Ageing, and know which alerts are new.
CI runners are disposable; Drive is the disk.

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
