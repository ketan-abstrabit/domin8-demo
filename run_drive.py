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
import os
import shutil
import subprocess
import sys
import traceback
from datetime import datetime, timezone
from pathlib import Path

import drive_sync as DS
import pipeline_config as C

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
        fp = DS.fingerprint(manifest)
        last = ws.load_state(RUN_STATE, {}) or {}
        published = [f for f in fs.children(ws.latest, refresh=True)
                     if f["mimeType"] != DS.FOLDER_MIME]
        log(f"\n[2] input fingerprint {fp}  (last run {last.get('fingerprint', '—')})")
        log(f"    output/latest holds {len(published)} file(s)")

        if fp == last.get("fingerprint") and not args.force:
            if published:
                log("    unchanged since the last run — nothing to do")
                DS.write_status(ws, ok=True, started=started, manifest=manifest,
                                lines=["Inputs are identical to the last run. "
                                       "Reports in output/latest/ are current.",
                                       "Add or replace a file in input/ to trigger "
                                       "a rebuild, or press Run with 'rebuild "
                                       "even if nothing has changed'."],
                                skipped=True)
                return 0
            log("    inputs unchanged, but output/latest is empty — "
                "rebuilding rather than reporting reports that are not there")

        if args.dry_run:
            log("\n--dry-run: would rebuild here. Stopping.")
            return 0

        # -- 3  build ----------------------------------------------------
        log("\n[3] building")
        if C.OUTPUT.exists():
            shutil.rmtree(C.OUTPUT)
        C.OUTPUT.mkdir(parents=True, exist_ok=True)
        rc = run_pipeline(args)
        if rc != 0:
            raise RuntimeError(f"run_pipeline.py exited {rc} — see the notes above")

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
        links = DS.push_outputs(ws, C.OUTPUT, stamp, log=log)
        ws.save_state_file(STATE_DIR / ALERT_STATE)
        if C.PO_HISTORY_FILE.exists():
            ws.save_state_file(C.PO_HISTORY_FILE)
        DS.push_reorder_sheet(ws, C.REORDER_OVERRIDES, log=log)
        ws.save_state(RUN_STATE, {
            "fingerprint": fp,
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
        ])
        log(f"\nDone. {len(links)} file(s) published to output/latest/.")
        return 0

    except Exception as exc:                                    # noqa: BLE001
        log("\nFAILED: " + str(exc))
        tb = traceback.format_exc().strip().splitlines()[-6:]
        for ln in tb:
            log("    " + ln)
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
    ap.add_argument("--requested-by", default="",
                    help="who pressed Run now, recorded in STATUS.txt so the "
                         "client can see who asked for a cycle and when")
    a = ap.parse_args()

    if not a.root_id:
        ap.error("--root-id (or DRIVE_ROOT_ID) is required")
    return cycle(a)


if __name__ == "__main__":
    sys.exit(main())
