"""Drive-mode entry point — what the scheduler actually runs.

The client's whole interface is a Drive folder. They drop reports into
`input/`; some minutes later the finished workbooks appear in `output/latest/`
and, if anything changed, a digest lands in their inbox. Nobody opens an app
and nobody sees the code.

One cycle:

    0  preflight     is the folder reachable, writable, and on a Shared Drive?
    1  pull          mirror Drive input/ onto local disk, restore saved state
    2  fingerprint   unchanged inputs? stop here, cheaply
    3  build         run_pipeline.py, exactly as it runs locally
    4  alerts        rules -> events -> NEW/ONGOING/RESOLVED -> digest
    5  push          output/latest/ (in place), output/archive/<stamp>/, state
    6  status        STATUS.txt, so a failure is visible instead of silent

Step 3 shells out to run_pipeline.py on purpose. The reconciler and the
merchandising workbook do not know Drive exists and should not learn — they
read a folder and write a folder, and stay identically testable offline.

    python run_drive.py --root-id <folder id> [--key-file sa.json]
                        [--force] [--no-email] [--dry-run]
"""

from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
from datetime import datetime, timezone
from pathlib import Path

import drive_sync as DS
import pipeline_config as C

# Windows consoles default to cp1252, which cannot encode the rupee sign
# this pipeline prints. That killed a run on a developer machine while
# working fine in CI, where the console is UTF-8. Force UTF-8 and replace
# anything unprintable rather than raising: a report must not die over a
# currency symbol.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


HERE = Path(__file__).resolve().parent
STATE_DIR = C.REPORTS / "_state"
RUN_STATE = "last_run.json"
ALERT_STATE = "alert_state.json"

_log_lines: list[str] = []


def log(msg: str = ""):
    print(msg, flush=True)
    _log_lines.append(str(msg))


def notes() -> list[str]:
    """The tail of the run log, for STATUS.txt. Enough to diagnose, not a dump."""
    return [ln for ln in _log_lines if ln.strip()][-40:]


# ---------------------------------------------------------------------------

def clear_local_inputs():
    """Drive is the source of truth, so start from an empty local mirror.

    Without this a report the client *deleted* from Drive would live on in the
    runner's checkout and keep appearing in the output — the pipeline would be
    reconciling a file nobody can see any more.
    """
    for folder in [C.INPUT] + [C.INPUT / k for k in DS.FOLDER_ALIASES]:
        if folder.exists():
            for f in folder.iterdir():
                if f.is_file():
                    f.unlink()
    C.INPUT.mkdir(parents=True, exist_ok=True)


def rules_fingerprint() -> str:
    """Hash of alert_rules.yaml, so editing a threshold counts as a change."""
    path = HERE / "alert_rules.yaml"
    if not path.exists():
        return "norules"
    return hashlib.sha256(path.read_bytes()).hexdigest()[:8]


def run_pipeline(args) -> int:
    cmd = [sys.executable, str(HERE / "run_pipeline.py"), "--force"]
    if args.fetch:
        cmd += ["--fetch", "--days", str(args.days)]
    if args.asof:
        cmd += ["--asof", args.asof]
    log(f"    {' '.join(cmd[1:])}")
    proc = subprocess.run(cmd, cwd=str(HERE), capture_output=True, text=True)
    tail = (proc.stdout or "").strip().splitlines()
    for ln in tail[-25:]:
        log("    " + ln)
    if proc.returncode != 0:
        for ln in (proc.stderr or "").strip().splitlines()[-15:]:
            log("    ! " + ln)
    return proc.returncode


# ---------------------------------------------------------------------------
# fetch — a separate button, on purpose
#
# Folding the Uniware pull into Run would put a third-party API in the path of
# the one button the client presses, so an ERP outage would read to them as
# "the report is broken". Kept apart, a failed fetch changes nothing: the
# previous pull is still in Drive and Run still works on it.
#
# It also means the client can skip it entirely and upload the five Uniware
# exports by hand, exactly as they do today.
# ---------------------------------------------------------------------------

PULL_STATE = "uniware_pull.json"
FETCH_STATE = "last_fetch.json"

UNIWARE_SCRIPT = "uniware_exports.py"


def _explain_uniware(stderr: str, attempts: int) -> str:
    """Turn a stack trace into something the person reading it can act on."""
    tail = (stderr or "").strip().splitlines()
    last = tail[-1] if tail else "no output"

    if "503" in last or "502" in last or "504" in last:
        return (
            f"Uniware returned a server error on all {attempts} attempts: "
            f"{last[:160]}\n"
            "This is Uniware's side, not the credentials — the same host "
            "refuses unrelated requests too. Either it is down, or it is "
            "refusing datacenter traffic: the same pull may well work from an "
            "office machine while failing from a cloud runner. Nothing in "
            "Drive was changed; the previous pull is still in input/uniware/.")
    if "401" in last or "invalid_grant" in last or "Bad credentials" in last:
        return (f"Uniware rejected the credentials: {last[:160]}\n"
                "Check UNIWARE_USER and UNIWARE_PASS, and that the account is "
                "not locked or password-expired.")
    if "timed out" in last.lower() or "ConnectionError" in last:
        return (f"Could not reach Uniware: {last[:160]}\n"
                "Network or DNS, not credentials.")
    return f"uniware_exports.py failed after {attempts} attempts: {last[:200]}"


def fetch_cycle(args) -> int:
    started = datetime.now(timezone.utc)
    local = DS.to_local(started)
    run_id = local.strftime("%d %b %Y, %H:%M")
    who = (getattr(args, "requested_by", "") or "").strip()

    log(f"DOMIN8 · Uniware fetch · {run_id}"
        + (f" · requested by {who}" if who else ""))

    if not (os.environ.get("UNIWARE_USER") and os.environ.get("UNIWARE_PASS")):
        log("\nFAILED: UNIWARE_USER / UNIWARE_PASS are not set. Add them as "
            "repository secrets before using the fetch button.")
        return 1

    log("\n[0] connecting to Drive")
    fs = DS.DriveFS(DS.build_service(key_file=args.key_file), args.root_id, log=log)
    DS.preflight(fs)
    ws = DS.Workspace(fs, log=log)

    tmp = Path(tempfile.mkdtemp(prefix="uniware_"))
    try:
        # -- 1  pull from Uniware ----------------------------------------
        log(f"\n[1] pulling {args.days} days from Uniware")
        env = dict(os.environ, UNIWARE_OUTDIR=str(tmp))
        if C.FACILITY:
            env["UNIWARE_FACILITY"] = C.FACILITY
        cmd = [sys.executable, str(HERE / UNIWARE_SCRIPT), "--days", str(args.days)]
        if C.FACILITY:
            cmd += ["--facility", C.FACILITY]

        # Uniware is a third party having its own day. A 5xx at the token
        # endpoint is transient often enough that failing the client's button
        # on the first one is the wrong answer — but not so often that we
        # should retry forever, so: three tries, widening gaps, then give up
        # honestly. Each attempt starts a fresh export job, so retrying is safe.
        attempts, proc, last = 3, None, ""
        for attempt in range(1, attempts + 1):
            try:
                proc = subprocess.run(cmd, cwd=str(HERE), env=env, text=True,
                                      capture_output=True,
                                      timeout=args.fetch_timeout)
            except subprocess.TimeoutExpired:
                raise RuntimeError(
                    f"Uniware did not finish within {args.fetch_timeout}s. "
                    f"Nothing was changed — the previous pull is still in "
                    f"input/uniware/.")

            for ln in (proc.stdout or "").strip().splitlines()[-20:]:
                log("    " + ln)
            if proc.returncode == 0:
                break

            last = (proc.stderr or "").strip()
            for ln in last.splitlines()[-8:]:
                log("    ! " + ln)
            if attempt < attempts:
                wait = 30 * attempt
                log(f"    attempt {attempt} of {attempts} failed — "
                    f"retrying in {wait}s")
                time.sleep(wait)

        if proc is None or proc.returncode != 0:
            raise RuntimeError(_explain_uniware(last, attempts))

        pulled = sorted(p for p in tmp.rglob("*.csv") if p.is_file())
        if not pulled:
            raise RuntimeError("Uniware returned no files")

        # -- 2  publish into input/uniware/ ------------------------------
        #
        # The files go to Drive, not just to this runner's disk. Reports built
        # from data nobody can see are unauditable, and the report run reads
        # Drive — so a pull that stayed local would be invisible to it.
        log(f"\n[2] uploading {len(pulled)} file(s) to input/uniware/")
        uni = ws.sub["uniware"]
        uploaded = []
        for f in pulled:
            got = fs.upload(f, uni)
            uploaded.append({"id": got["id"], "name": got["name"]})
            log(f"      {f.stat().st_size:>10,}  {got['name']}")

        # -- 3  retire the previous pull ---------------------------------
        #
        # Uniware stamps its filenames, so every pull would otherwise pile up.
        # The reconciler would still pick the newest, but the folder becomes
        # unreadable and the client cannot tell what is current. Only files
        # this tool uploaded before are removed — anything they put there by
        # hand is left alone.
        previous = ws.load_state(PULL_STATE, {}) or {}
        keep = {u["id"] for u in uploaded}
        retired = 0
        for old in previous.get("files", []):
            if old.get("id") in keep:
                continue
            try:
                fs.trash(old["id"], uni)
                retired += 1
            except Exception:                                   # noqa: BLE001
                pass
        if retired:
            log(f"      retired {retired} file(s) from the previous pull")

        ws.save_state(PULL_STATE, {"files": uploaded, "run_id": run_id})
        ws.save_state(FETCH_STATE, {
            "ok": True, "run_id": run_id, "files": len(uploaded),
            "days": args.days, "requested_by": who,
            "names": [u["name"] for u in uploaded],
        })
        log(f"\nDone. {len(uploaded)} Uniware report(s) in input/uniware/. "
            f"Press Run to build.")
        return 0

    except Exception as exc:                                    # noqa: BLE001
        log("\nFAILED: " + str(exc))
        try:
            ws.save_state(FETCH_STATE, {
                "ok": False, "run_id": run_id, "error": str(exc)[:400],
                "requested_by": who,
            })
        except Exception:                                       # noqa: BLE001
            pass
        _email_failure("Uniware fetch failed: " + str(exc), run_id)
        return 1
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------------------

def cycle(args) -> int:
    started = datetime.now(timezone.utc)
    # Stamped in the timezone the reports are read in, not the runner's UTC —
    # otherwise archive folders sort under the wrong date for evening runs.
    local = DS.to_local(started)
    stamp = local.strftime("%Y-%m-%d_%H%M")
    run_id = local.strftime("%d %b %Y, %H:%M")

    who = (getattr(args, "requested_by", "") or "").strip()
    log(f"DOMIN8 · Drive run · {run_id}"
        + (f" · requested by {who}" if who else " · scheduled"))

    # -- 0  connect ------------------------------------------------------
    log("\n[0] connecting to Drive")
    svc = DS.build_service(key_file=args.key_file)
    fs = DS.DriveFS(svc, args.root_id, log=log)
    meta = DS.preflight(fs)
    log(f"    folder '{meta['name']}' on shared drive {meta.get('driveId')}")
    ws = DS.Workspace(fs, log=log)

    manifest: dict = {"files": []}
    try:
        # -- 1  pull -----------------------------------------------------
        log("\n[1] pulling input/ from Drive")
        clear_local_inputs()
        manifest = DS.pull_inputs(ws, C.INPUT, overrides=C.REORDER_OVERRIDES,
                                  log=log)
        if not manifest["files"]:
            raise RuntimeError(
                "input/ is empty in Drive. Nothing to build. Drop the master "
                "mapping table and this cycle's reports in and it will run.")

        STATE_DIR.mkdir(parents=True, exist_ok=True)
        if ws.load_state_file(C.PO_HISTORY_FILE.name, C.PO_HISTORY_FILE):
            log(f"    restored {C.PO_HISTORY_FILE.name} from _state/")
        ws.load_state_file(ALERT_STATE, STATE_DIR / ALERT_STATE)

        # -- 2  has anything changed? ------------------------------------
        #
        # "Inputs unchanged" is only a reason to skip if last time's reports
        # are still there to stand in for this run. The fingerprint says
        # nothing about that: delete output/latest, or lose a publish halfway,
        # and the pipeline would skip for ever, insisting reports are current
        # while the folder sits empty. So the outputs get a vote.
        #
        # The thresholds count as an input too. They live in the repo rather
        # than in Drive, so changing one and pushing would leave the client
        # pressing Run and getting a skip — the same alerts as before, because
        # nothing in Drive had moved. Folding the rules into the fingerprint
        # means a threshold change rebuilds on its own, and removes the last
        # honest reason for anyone to reach for the force checkbox.
        fp = DS.fingerprint(manifest) + "+" + rules_fingerprint()
        last = ws.load_state(RUN_STATE, {}) or {}
        published = [f for f in fs.children(ws.latest, refresh=True)
                     if f["mimeType"] != DS.FOLDER_MIME]
        # The BI sheets get the same vote as the reports. They are published
        # in step 5, so a skipped run never creates them — which meant turning
        # this feature on for a folder whose inputs had not changed left
        # output/bi/ empty for ever, with the run reporting success. Whatever
        # downstream artefact a skip claims is current has to actually exist.
        bi_now = [f for f in fs.children(ws.bi, refresh=True)
                  if f["mimeType"] != DS.FOLDER_MIME]
        bi_wanted = DS.BI_EXPECTED
        log(f"\n[2] input fingerprint {fp}  (last run {last.get('fingerprint', '—')})")
        log(f"    output/latest holds {len(published)} file(s), "
            f"output/bi holds {len(bi_now)}")

        if last.get("ok") is False:
            log("    the last run failed — rebuilding rather than skipping")
        elif fp == last.get("fingerprint") and not args.force and bi_now and len(bi_now) < bi_wanted:
            log(f"    output/bi has {len(bi_now)} of {bi_wanted} sheets — "
                "rebuilding to complete it")
        elif fp == last.get("fingerprint") and not args.force:
            if published and bi_now:
                log("    unchanged since the last run — nothing to do")
                DS.write_status(ws, ok=True, started=started, manifest=manifest,
                                lines=["Inputs are identical to the last run. "
                                       "Reports in output/latest/ are current.",
                                       "Add or replace a file in input/ to trigger "
                                       "a rebuild, or press Run with 'rebuild "
                                       "even if nothing has changed'."],
                                skipped=True)
                return 0
            log("    inputs unchanged, but output/latest or output/bi is "
                "empty — rebuilding rather than claiming they are current")

        if args.dry_run:
            log("\n--dry-run: would rebuild here. Stopping.")
            return 0

        # -- 3  build ----------------------------------------------------
        log("\n[3] building")
        if C.OUTPUT.exists():
            shutil.rmtree(C.OUTPUT)
        C.OUTPUT.mkdir(parents=True, exist_ok=True)
        rc = run_pipeline(args)
        # A non-zero exit from run_pipeline.py means "something is worth your
        # attention", not necessarily "there is no report". It returns 1 for
        # warnings too — stale inputs, a locked file, a check that could not be
        # verified — after building everything successfully. Throwing that away
        # would mean one aged Amazon export costs the client their whole cycle.
        #
        # So the outputs decide. If the two headline workbooks are there, the
        # build worked; the warnings are carried into STATUS.txt where someone
        # can act on them. If they are not, it really did fail.
        # Every headline workbook, not merely one of them. A crash part-way
        # through leaves some of them on disk, and "at least one exists" was
        # lenient enough to publish a half-built cycle as a warning.
        built = [p.name for p in C.OUTPUT.glob("*") if p.is_file()]
        patterns = getattr(C, "PIPELINE_OUTPUTS", None) or getattr(C, "MAIN_OUTPUTS", [])
        missing = [pat for pat in patterns
                   if not any(DS.is_main_output(n, [pat]) for n in built)]
        if rc != 0 and missing:
            raise RuntimeError(
                f"run_pipeline.py exited {rc} and did not produce "
                f"{', '.join(missing)} — see the notes above")
        warned = rc != 0
        if warned:
            log(f"    run_pipeline.py exited {rc} but produced "
                f"{len(built)} file(s) — treating as warnings, see STATUS.txt")

        # -- 4  alerts ---------------------------------------------------
        log("\n[4] alerts")
        import alerts
        summary = alerts.run(
            report_dir=C.OUTPUT,
            rules_path=HERE / "alert_rules.yaml",
            state_path=STATE_DIR / ALERT_STATE,
            run_id=run_id,
            email=not args.no_email,
            log=log,
        )

        # -- 5  push -----------------------------------------------------
        log("\n[5] publishing to Drive")
        links = DS.push_outputs(ws, C.OUTPUT, stamp,
                                main_patterns=getattr(C, "MAIN_OUTPUTS", []),
                                extras_dir=getattr(C, "EXTRAS_DIR", "extras"),
                                log=log)
        DS.push_bi_tables(ws, C.OUTPUT, log=log)
        ws.save_state_file(STATE_DIR / ALERT_STATE)
        if C.PO_HISTORY_FILE.exists():
            ws.save_state_file(C.PO_HISTORY_FILE)
        DS.push_reorder_sheet(ws, C.REORDER_OVERRIDES, log=log)
        ws.save_state(RUN_STATE, {
            "fingerprint": fp,
            "ok": True,
            "run_id": run_id,
            "stamp": stamp,
            "files": len(manifest["files"]),
            "alerts_open": summary["open"],
            "alerts_new": summary["new"],
        })

        # -- 6  status ---------------------------------------------------
        DS.write_status(ws, ok=True, started=started, manifest=manifest, lines=[
            f"{summary['open']:,} open alerts, {summary['new']:,} new, "
            f"{summary['resolved']:,} cleared.",
            f"Rs {summary['capital_at_risk']:,.0f} of stock is flagged as "
            f"slow, ageing or dead.",
            ("Digest emailed." if summary["emailed"]
             else "No email sent this run (nothing changed, or SMTP not configured)."),
            f"Archived as output/archive/{stamp}/.",
            (f"Requested by {who}." if who else "Started by the schedule."),
        ] + ([
            "",
            "WARNINGS — the report was built, but something wants attention:",
        ] + [f"  {ln.strip()}" for ln in _log_lines
             if ln.strip().startswith("- ")][-8:] if warned else []))
        log(f"\nDone. {len(links)} file(s) published to output/latest/.")
        return 0

    except Exception as exc:                                    # noqa: BLE001
        log("\nFAILED: " + str(exc))
        tb = traceback.format_exc().strip().splitlines()[-6:]
        for ln in tb:
            log("    " + ln)
        # Record the failure, so the next press retries instead of comparing
        # fingerprints and skipping. A run that died between publishing some
        # files and publishing the rest leaves output/latest non-empty, so the
        # "are there reports?" guard would wave it through — pressing Run again
        # has to be a retry, always.
        try:
            ws.save_state(RUN_STATE, {"fingerprint": None, "ok": False,
                                      "run_id": run_id, "error": str(exc)[:400]})
        except Exception:                                       # noqa: BLE001
            log("    (could not record the failure in _state either)")
        try:
            DS.write_status(ws, ok=False, started=started, manifest=manifest,
                            lines=notes())
        except Exception:                                       # noqa: BLE001
            log("    (could not write STATUS.txt either — check the share)")
        _email_failure(str(exc), run_id)
        return 1


def _email_failure(reason: str, run_id: str):
    """A run that dies quietly is worse than one that dies loudly.

    With no dashboard to look at, silence is indistinguishable from success —
    the client would assume the reports are current when they are a fortnight
    stale. So a failure emails the operator, not the client.
    """
    to = [a for a in os.environ.get("ALERT_ADMINS", "").replace(";", ",").split(",")
          if a.strip()]
    if not to:
        return
    try:
        import alerts
        body = (f"<p>The DOMIN8 Drive run failed on {run_id}.</p>"
                f"<pre style='background:#f6f8fa;padding:10px;border-radius:6px'>"
                f"{reason}</pre><p>Last lines of the run log:</p>"
                f"<pre style='font-size:12px'>" + "\n".join(notes()[-20:]) + "</pre>")
        alerts.send_email(f"[DOMIN8] report run FAILED — {run_id}", body,
                          [a.strip() for a in to], log=log)
    except Exception:                                           # noqa: BLE001
        pass


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root-id", default=os.environ.get("DRIVE_ROOT_ID"),
                    help="Drive folder ID of the DOMIN8 reporting folder "
                         "(or set DRIVE_ROOT_ID)")
    ap.add_argument("--key-file", default=None,
                    help="path to the service-account JSON (or set GOOGLE_SA_KEY "
                         "to the JSON itself, which is what CI does)")
    ap.add_argument("--force", action="store_true",
                    help="rebuild even when the inputs have not changed")
    ap.add_argument("--no-email", action="store_true")
    ap.add_argument("--dry-run", action="store_true",
                    help="pull and compare, then stop before building")
    ap.add_argument("--fetch", action="store_true",
                    help="pull Uniware over the API before building")
    ap.add_argument("--days", type=int, default=C.DEFAULT_DAYS)
    ap.add_argument("--asof", help="Stock vs Sales report date, YYYY-MM-DD")
    ap.add_argument("--fetch-only", action="store_true",
                    help="pull Uniware into Drive input/uniware/ and stop. "
                         "This is what the client's 'Fetch Uniware' button does; "
                         "it never touches the reports.")
    ap.add_argument("--fetch-timeout", type=int, default=1500,
                    help="seconds to allow the Uniware pull (default 1500)")
    ap.add_argument("--requested-by", default="",
                    help="who pressed Run now, recorded in STATUS.txt so the "
                         "client can see who asked for a cycle and when")
    a = ap.parse_args()

    if not a.root_id:
        ap.error("--root-id (or DRIVE_ROOT_ID) is required")
    return fetch_cycle(a) if a.fetch_only else cycle(a)


if __name__ == "__main__":
    sys.exit(main())
