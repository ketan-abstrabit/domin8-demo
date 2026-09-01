# Setting up Drive mode

One-time setup, about 30 minutes. After this the client's entire interface is a
Drive folder: they drop files in `input/`, the reports appear in
`output/latest/`, and a digest lands in their inbox when something changes.

Nothing runs on their machine. Nothing costs anything.

---

## 1. Make the Shared Drive

**This has to be a Shared Drive, not a folder in someone's My Drive.** A service
account has no storage quota of its own, so it cannot own files; uploading into
a personal My Drive folder fails with `403 storageQuotaExceeded` even when the
folder is shared with full edit rights. Files in a Shared Drive are owned by the
drive, so the upload succeeds. `run_drive.py` checks this before doing any work
and refuses with an explanation rather than dying half way.

1. Drive → **Shared drives** → **New** → call it `DOMIN8 Reporting`.
2. Inside it create one folder: `input`.

That is all you need to create by hand. The first run creates `output/`,
`output/latest/`, `output/archive/`, `_state/` and the three input subfolders.

Open `input` and copy the folder ID out of the address bar — the part after
`/folders/`:

```
https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQrStUvWxYz
                                        └──────── this ────────┘
```

You want the ID of the **top-level `DOMIN8 Reporting` drive**, not `input`.

## 2. Make the service account

1. [console.cloud.google.com](https://console.cloud.google.com) → new project,
   e.g. `domin8-reporting`. No billing account needed.
2. **APIs & Services → Library → Google Drive API → Enable.**
3. **IAM & Admin → Service Accounts → Create.** Name it `domin8-runner`.
   Skip the optional role and user steps.
4. Open it → **Keys → Add key → Create new key → JSON.** A file downloads.
   Keep it; it is shown once.
5. Copy the account's email — it looks like
   `domin8-runner@domin8-reporting.iam.gserviceaccount.com`.

## 3. Let it into the drive

Shared drive → **Manage members** → paste the service-account email → role
**Content manager**.

Not Viewer, not Contributor. Content manager is the lowest role that can
overwrite an existing file, which is how `output/latest/` keeps stable links.

## 4. Wire up GitHub

Repo → **Settings → Secrets and variables → Actions → New repository secret**.

**Required — nothing runs without these two:**

| Secret | Value |
|---|---|
| `GOOGLE_SA_KEY` | the entire contents of the JSON key file, pasted as-is |
| `DRIVE_ROOT_ID` | the folder ID from step 1 |

**Optional — email only. Skip the whole block until you want the digest sent:**

| Secret | Value |
|---|---|
| `ALERT_RECIPIENTS` | who gets the digest, comma-separated |
| `ALERT_ADMINS` | who gets told when a *run fails* — you, not the client |
| `SMTP_USER` | the mailbox that sends |
| `SMTP_PASS` | its app password (see below) |
| `SMTP_FROM` | defaults to `SMTP_USER` |
| `SMTP_HOST` / `SMTP_PORT` | default to `smtp.gmail.com` / `587` |

With the SMTP secrets absent, everything still builds and publishes; the run
logs `email not sent — SMTP_USER / SMTP_PASS / recipients not all set` and
carries on. The digest is still written to `output/latest/alert_digest.html`,
so you can open it and see exactly what the email would have said. Adding the
secrets later needs no code change and no redeploy.

For a Workspace mailbox, `SMTP_PASS` is an **app password**, not the account
password: Google Account → Security → 2-Step Verification → App passwords. If
the admin has app passwords disabled, use any mailbox that allows them, or
switch `SMTP_HOST` to your own provider.

## 5. Put the first cycle's files in

The run needs real inputs — an empty `input/` fails on purpose rather than
publishing an empty report. Upload into the shared drive:

```
input/
  Marketplace product id Master.xlsx     ← loose in input/, not in a subfolder
  uniware/          Tally GST, Tally Return GST, Purchase Orders,
                    Inventory Snapshot, Item Master
  amazon vc/        the Sales and Inventory exports
  retail stores/    the store sale and SOH files
```

Create the three subfolders yourself, or run the workflow once first — it
creates them (and then fails on the empty input, which is expected).

Files are matched **by content, not filename**, so date-stamped exports work
unrenamed and the subfolder names are the only thing that has to be right.

## 6. First run

Repo → **Actions → DOMIN8 report → Run workflow**, tick **force**.

Watch the log. If it fails, it is almost always one of the mistakes above and
the error message names which:

| Message | Means |
|---|---|
| `storageQuotaExceeded` / `SHARED DRIVE` | the folder is in a My Drive, not a shared drive |
| `cannot write to it` | service account has Viewer, needs Content manager |
| `input/ is empty in Drive` | step 5 not done |
| `No Google credentials` | `GOOGLE_SA_KEY` missing or not valid JSON |

Then check the drive: `output/latest/` should hold ten files, `output/bi/`
three Google Sheets, `STATUS.txt` should say `Result : OK`, and `input/` should
have gained a `reorder_status` Google Sheet.

## 7. Give the client the button

The report runs when someone asks for it. That someone is the client, pressing
a button on a Google web app — see **[apps_script/README.md](apps_script/README.md)**
for the 15-minute setup.

They get one URL. Behind it: Run now, the last run's status, and a link to the
dashboard. They never see GitHub and never hold a credential; the token lives
in a script you own that runs as you.

## 8. Hand over

Give the client edit access to the shared drive and tell them three things:

- Put files in `input/` — the subfolder names are `uniware`, `amazon vc`,
  `retail stores`. The master mapping table goes loose in `input/`.
- Press **Run the report** on the web app. Three to four minutes later the
  reports are in `output/latest/`. Previous cycles are in `output/archive/`.
- If something looks wrong, the page shows it, and `STATUS.txt` in the folder
  says when the last run happened, what it read, and whether it worked.

---

## Connecting Looker Studio

Looker Studio cannot read the `.xlsx` and `.csv` files in `output/latest/`. Its
Sheets connector opens native Google Sheets only, and the "Google Drive"
connector in the gallery is a third-party one that lists files and folders —
metadata, not their contents. So each run also republishes the three fact
tables to `output/bi/` as Google Sheets, which the connector does read.

One-time, per table:

1. Looker Studio → **Create → Data source → Google Sheets**.
2. **Shared drives** tab → `DOMIN8 Reporting` → `output` → `bi` →
   `fact_sales`, `fact_inventory` or `fact_purchase`.
3. Tick **Use first row as headers**. Connect.

A data source binds to a Drive file ID, and the run overwrites these sheets in
place rather than replacing them, so the IDs hold and the dashboard survives
every run without re-linking. Deleting a sheet from `bi/` breaks that — the
next run creates a new file with a new ID and the data source has to be
repointed.

Set **Data freshness** to 15 minutes on each source. The sheets only change
when a run happens, so anything shorter just adds queries.

The data source runs on the credentials of whoever creates it. Use an account
that will keep access to the shared drive — if it is a personal one and that
person leaves, every report built on it stops.

## When it runs

**When the client presses the button.** That is the primary path — the report
is built when someone actually wants it.

There is also a forced run on the 1st and 16th at 08:00 IST, matching their
twice-monthly cycle, so a cycle still goes out if nobody remembers to press
anything. Edit the `cron:` line in `.github/workflows/domin8-report.yml` to
change it; GitHub cron is UTC and IST is UTC+5:30.

There is deliberately **no frequent poll**. Polling only ever existed to notice
uploads, and the button notices them the moment the client does. Dropping it
takes Actions usage to near zero — a handful of runs a month instead of ~200.

You can still run it yourself from **Actions → DOMIN8 report → Run workflow**
when debugging.

## Changing the thresholds

`alert_rules.yaml`. Edit, commit, push. The next run uses the new numbers.

Every rule can be muted with `enabled: false` without losing its settings, and
each has a comment explaining what the number does and how it behaves against
the real data.

## Running it by hand

```bash
export GOOGLE_APPLICATION_CREDENTIALS=/path/to/key.json
export DRIVE_ROOT_ID=1AbCdEf...

python run_drive.py --dry-run     # pull and compare, build nothing
python run_drive.py --force       # rebuild and publish
python run_drive.py --no-email    # publish, send nothing
```

## Testing without touching Drive

```bash
python tests/test_drive_cycle.py
```

Runs the whole pull → build → publish cycle against an in-memory fake Drive,
offline, in about two minutes: change detection, in-place overwrite, archive
retention, deleted-file handling, alert state transitions, the My Drive
failure, and the failure-path `STATUS.txt`. 26 checks.

## What lives where

```
input/                    the client fills this
  <master mapping table>
  uniware/
  amazon vc/
  retail stores/
  reorder_status          Google Sheet — merchandiser overrides, edited in place
output/
  latest/                 overwritten each run, file IDs preserved
  archive/<date_time>/    last 12 cycles
  bi/                     the three fact tables as Google Sheets
STATUS.txt                last run: when, what it read, pass/fail
_state/                   fingerprints, PO history, alert state — leave alone
```

`_state/` is how the pipeline remembers things between runs: which alerts were
already open (so they do not re-fire), and the accumulated purchase-order
history that drives Ageing. Deleting it is not fatal — the next run treats
every alert as new and rebuilds PO history from whatever window it can see.
