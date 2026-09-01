# The Run now button

A Google Apps Script web app. The client opens one URL, presses one button, and
the report builds. They never see GitHub, never see the code, and never touch a
credential.

```
client presses Run  →  Apps Script (holds the token)  →  GitHub repository_dispatch
                                                              ↓
                              Drive output/latest/  ←  the workflow
```

The script is **standalone and owned by you**, deployed to run as you. That is
the whole security design: `google.script.run` executes on Google's servers
under your identity, so the GitHub token never reaches the client's browser.
Bind this to a Sheet the client can edit and they could read the token from the
Apps Script editor — so don't.

---

## Setup, about 15 minutes

### 1. Make the GitHub token

GitHub → **Settings → Developer settings → Personal access tokens →
Fine-grained tokens → Generate new token**.

- **Repository access:** Only select repositories → `domin8-demo`
- **Permissions:**
  - **Contents: Read and write** — required to fire `repository_dispatch`
  - **Actions: Read** — required to show run status on the page
- **Expiration:** whatever your policy allows. Put a reminder in the calendar;
  when it lapses the button stops and `selftest()` will say exactly why.

Copy the token. GitHub shows it once.

> If either permission turns out to be wrong, you won't have to guess — the
> page and `selftest()` print GitHub's own error, which names what's missing.

### 2. Create the script

[script.google.com](https://script.google.com) → **New project**. Name it
`DOMIN8 runner`.

- Replace the contents of `Code.gs` with this folder's `Code.gs`.
- **+ → HTML → name it exactly `Index`**, and paste this folder's `Index.html`.

The filename must be `Index` — `Code.gs` loads the template by that name.

### 3. Script properties

**Project Settings → Script Properties → Add script property**, three of them:

| Property | Value |
|---|---|
| `GITHUB_REPO` | `ketan-abstrabit/domin8-demo` |
| `GITHUB_TOKEN` | the token from step 1 |
| `DRIVE_ROOT_ID` | the shared-drive folder id — same value as the GitHub secret |

### 4. Check it before deploying

In the editor, select **`selftest`** from the function dropdown and press Run.
The first run asks for authorisation — approve it; that's Google asking whether
your script may reach Drive and the internet on your behalf.

Then read the execution log. Eight checks, each naming what it verified:

```
  PASS  GITHUB_REPO is set  — ketan-abstrabit/domin8-demo
  PASS  token can read the repo  — ketan-abstrabit/domin8-demo (private)
  PASS  token can read workflow runs  [Actions: read]  — 7 run(s) in history
  PASS  DRIVE_ROOT_ID resolves  — DOMIN8 Reporting
  PASS  output/latest exists  — 10 published file(s)
```

`selftest` never starts a run, so it is safe to re-run any time something looks
wrong. It is the only test that can exist for this — Apps Script cannot be
exercised from outside Google, so the checks live inside it.

### 5. Deploy

**Deploy → New deployment → Web app**

| Setting | Value |
|---|---|
| Execute as | **Me** — this is what keeps the token server-side |
| Who has access | **Anyone within `<your Workspace domain>`** |

If the client is on a different Workspace domain than you, "Anyone with a
Google account" is the only option that will let them in. That is a wide door:
anyone who guesses the URL could press the button. It cannot leak data (the
page shows only run status) and it cannot leak the token, but it can burn
Actions minutes. The two-minute cooldown in `triggerRun` limits the damage.
Prefer sharing a domain if you can.

Copy the web app URL and send it to the client. That URL is the entire product
as far as they are concerned.

### 6. Redeploying after a change

Editing the code does **not** update the live web app. **Deploy → Manage
deployments → the pencil icon → Version: New version → Deploy.** Same URL, new
code. Forgetting this step is the most common reason a fix appears to do
nothing.

---

## What the client sees

- **Run the report** — starts a build. Status updates every six seconds while
  it runs; three to four minutes end to end. Closing the tab doesn't stop it.
- **Rebuild even if nothing has changed** — off by default. Normally the run
  skips in about 30 seconds when the inputs are identical to last time, which
  is the correct answer and a fast one. Tick this to force a full rebuild.
- **Details of the last run** — `STATUS.txt` verbatim: when, what it read, what
  it found, who asked for it.
- **Dashboard** — serves `dashboard.html` from `output/latest/` as a real web
  page. Drive itself will only offer to download an HTML file; this renders it,
  behind the same Google sign-in, always showing the newest run.

## Guards worth knowing about

**Two-minute cooldown.** A second press inside two minutes is refused with an
explanation rather than queueing a duplicate run.

**Concurrency group.** If a run somehow starts while another is publishing,
GitHub queues it rather than cancelling — a double-click can't corrupt a
half-written `output/latest/`.

**Who pressed it** is recorded in `client_payload.requested_by` and lands in
`STATUS.txt`, so a cycle can be traced to a person.

## When something breaks

| Symptom | Cause |
|---|---|
| "GitHub rejected the token (401)" | token expired or revoked — regenerate, update `GITHUB_TOKEN` |
| "GitHub returned 403/404" | missing permission, or the repo isn't selected on the token |
| Button works, nothing happens | you edited the code but didn't redeploy a **new version** |
| "Cannot reach GitHub" | usually the token again; `selftest()` will say which |
| Page loads but links are missing | `DRIVE_ROOT_ID` wrong, or no successful run yet |

Run `selftest()` first for any of these. It checks each assumption separately
and reports GitHub's own words rather than a guess.
