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

Repo → **Settings → Secrets and variables → Actions → New repository secret**:

| Secret | Value |
|---|---|
| `GOOGLE_SA_KEY` | the entire contents of the JSON key file, pasted as-is |
| `DRIVE_ROOT_ID` | the folder ID from step 1 |
| `ALERT_RECIPIENTS` | who gets the digest, comma-separated |
| `ALERT_ADMINS` | who gets told when a *run fails* — you, not the client |
| `SMTP_USER` | the mailbox that sends |
| `SMTP_PASS` | its app password (see below) |
| `SMTP_FROM` | optional, defaults to `SMTP_USER` |
| `SMTP_HOST` / `SMTP_PORT` | optional, defaults to `smtp.gmail.com` / `587` |

For a Workspace mailbox, `SMTP_PASS` is an **app password**, not the account
password: Google Account → Security → 2-Step Verification → App passwords. If
the admin has app passwords disabled, use any mailbox that allows them, or
switch `SMTP_HOST` to your own provider.

Leave the SMTP secrets out entirely and everything still works — the digest is
written to `output/latest/alert_digest.html` and simply is not emailed.

## 5. First run

Repo → **Actions → DOMIN8 report → Run workflow**, tick **force**.

Watch the log. If it fails, the reason is almost always one of the two
mistakes above, and the error message names which.

Then check the drive: `output/latest/` should hold ten files, `STATUS.txt`
should say `Result : OK`, and `input/` should have gained a `reorder_status`
Google Sheet.

## 6. Hand over

Give the client edit access to the shared drive and tell them three things:

- Put files in `input/` — the subfolder names are `uniware`, `amazon vc`,
  `retail stores`. The master mapping table goes loose in `input/`.
- Reports appear in `output/latest/` within about two hours. Previous cycles
  are in `output/archive/`.
- If something looks wrong, open `STATUS.txt` — it says when the last run
  happened, what it read, and whether it worked.

---

## How often it runs

Every two hours through the Indian working day, plus a forced run on the 1st
and 16th to match their twice-monthly cycle.

Most firings cost about 30 seconds: `run_drive.py` fingerprints the input
folder against the last run and stops if nothing changed. Only a real upload
triggers a build. On a private repo that is roughly 2 hours of the free
2,000-minute monthly allowance; on a public repo, nothing.

To change the cadence, edit the two `cron:` lines in
`.github/workflows/domin8-report.yml`. They are UTC — IST is UTC+5:30.

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
STATUS.txt                last run: when, what it read, pass/fail
_state/                   fingerprints, PO history, alert state — leave alone
```

`_state/` is how the pipeline remembers things between runs: which alerts were
already open (so they do not re-fire), and the accumulated purchase-order
history that drives Ageing. Deleting it is not fatal — the next run treats
every alert as new and rebuilds PO history from whatever window it can see.
