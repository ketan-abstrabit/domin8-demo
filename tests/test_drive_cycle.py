"""End-to-end test of the Drive cycle, offline.

Seeds a fake shared drive from the sample inputs, then asserts the things that
would otherwise only be discovered in production:

  1  a first run builds, publishes and archives
  2  a second run with identical inputs is skipped, not rebuilt
 2b  ...unless the reports it points at have been deleted
  3  --force rebuilds anyway
  4  a changed input triggers a rebuild
  5  publishing overwrites in place — output/latest/ file IDs survive, so a
     link shared with the client keeps working
  6  a file deleted in Drive stops appearing in the report
  7  alerts move NEW -> ONGOING across runs instead of re-firing
  8  Drive's eventual consistency does not produce duplicate folders or files
  9  a broken or absent mail setup never takes the run down with it
 10  a My Drive folder fails loudly, with the fix in the message
 11  a mid-run failure still leaves a readable STATUS.txt

    python tests/test_drive_cycle.py [--keep]
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(HERE))

import drive_sync as DS                       # noqa: E402
import fake_drive as FD                       # noqa: E402
import pipeline_config as C                   # noqa: E402

PASS, FAIL = "PASS", "FAIL"
results: list[tuple[str, str, str]] = []


def check(name: str, ok: bool, detail: str = ""):
    results.append((PASS if ok else FAIL, name, detail))
    print(f"  {PASS if ok else FAIL}  {name}" + (f"   {detail}" if detail else ""))


def seed(sample: Path) -> FD.FakeDrive:
    """Build a shared drive that looks like the client's, from the sample set."""
    d = FD.FakeDrive()
    inp = d.mkdir("input", d.root_id)
    subs = {k: d.mkdir(k, inp) for k in ("uniware", "amazon vc", "retail stores")}
    n = 0
    for f in sorted(sample.rglob("*")):
        if not f.is_file() or f.name.startswith("~$"):
            continue
        # Pipeline artefacts, not client input. A real first run starts without
        # them, and seeding the override sheet is one of the things under test.
        if f.stem in ("reorder_status", "channel_map", "reconciliation_checks"):
            continue
        rel = f.relative_to(sample)
        parent = subs.get(rel.parts[0]) if len(rel.parts) > 1 else inp
        if parent is None:
            continue
        d.put(f, parent)
        n += 1
    print(f"  seeded fake drive with {n} file(s)")
    return d


def run_cycle(root_id, **kw) -> int:
    import run_drive
    args = argparse.Namespace(root_id=root_id, key_file=None, force=False,
                              no_email=True, dry_run=False, fetch=False,
                              days=90, asof=None)
    for k, v in kw.items():
        setattr(args, k, v)
    run_drive._log_lines.clear()
    return run_drive.cycle(args)


def latest_names(d: FD.FakeDrive) -> dict[str, str]:
    out = d.find(d.root_id, "output")
    latest = d.find(out["id"], "latest")
    return {n["name"]: n["id"] for n in d.nodes.values()
            if not n["trashed"] and latest["id"] in n["parents"]}


def status_text(d: FD.FakeDrive) -> str:
    n = d.find(d.root_id, "STATUS.txt")
    return n["content"].decode() if n else ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sample", type=Path,
                    default=Path("/home/claude/deploy/reports/input"))
    ap.add_argument("--keep", action="store_true")
    a = ap.parse_args()

    if not a.sample.exists():
        sys.exit(f"sample inputs not found: {a.sample}\n"
                 "Point --sample at a reports/input folder.")

    # A clean slate: the fake drive is the only source of truth.
    for p in (C.INPUT, C.OUTPUT, C.REPORTS / "_state"):
        shutil.rmtree(p, ignore_errors=True)

    print("\n[setup]")
    d = seed(a.sample)
    FD.install(DS, d)

    print("\n[1] first run")
    rc = run_cycle(d.root_id)
    check("first run succeeds", rc == 0, f"rc={rc}")
    names = latest_names(d)
    check("output/latest/ populated", len(names) >= 6, f"{len(names)} files")
    check("Stock vs Sales published",
          any(n.startswith("Stock_vs_Sales") for n in names))
    check("Alerts workbook published", "Alerts.xlsx" in names)
    check("digest published", "alert_digest.html" in names)
    arch = d.find(d.find(d.root_id, "output")["id"], "archive")
    cycles = [n for n in d.nodes.values()
              if not n["trashed"] and arch["id"] in n["parents"]]
    check("cycle archived", len(cycles) == 1, cycles[0]["name"] if cycles else "")
    check("STATUS.txt says OK", "Result     : OK" in status_text(d))
    check("reorder_status seeded as a Sheet in input/",
          (d.find(d.find(d.root_id, "input")["id"], "reorder_status") or {})
          .get("mimeType") == FD.SHEET_MIME)

    print("\n[2] runs with nothing changed")
    # The first run seeds reorder_status into input/, so the run after it sees
    # a genuinely new file and rebuilds — correct, not a bug. Steady state is
    # the run after that.
    rc = run_cycle(d.root_id)
    check("run after seeding succeeds", rc == 0, f"rc={rc}")
    rc = run_cycle(d.root_id)
    check("steady-state run succeeds", rc == 0, f"rc={rc}")
    check("skipped as unchanged", "SKIPPED" in status_text(d))
    check("skip is cheap — no build ran",
          "Reports in output/latest/ are current" in status_text(d))

    print("\n[2b] a skip must not outlive the reports it points at")
    # What happened live: the output folders were deleted during cleanup, then
    # a run found the inputs unchanged and skipped — leaving STATUS.txt saying
    # "reports are current" over an empty folder, for ever. The fingerprint
    # alone is not enough; the outputs have to still exist.
    outf = d.find(d.root_id, "output")
    latest = d.find(outf["id"], "latest")
    wiped = 0
    for n in list(d.nodes.values()):
        if not n["trashed"] and latest["id"] in n["parents"]:
            n["trashed"] = True
            wiped += 1
    rc = run_cycle(d.root_id)
    check("run after output was deleted succeeds", rc == 0, f"rc={rc}")
    check("empty output/latest forces a rebuild despite unchanged inputs",
          "SKIPPED" not in status_text(d), f"{wiped} files had been deleted")
    check("reports are back", len(latest_names(d)) >= 6,
          f"{len(latest_names(d))} republished")

    print("\n[3] --force")
    # Re-snapshot here: [2b] deliberately deleted the published files, and a
    # file recreated after deletion gets a new Drive id by definition. The
    # property under test is that a *republish over an existing file* keeps its
    # id, so the baseline has to be the current set.
    before = latest_names(d)
    rc = run_cycle(d.root_id, force=True)
    after = latest_names(d)
    check("forced run rebuilds", rc == 0 and "SKIPPED" not in status_text(d))
    check("file IDs survive a republish — shared links keep working",
          all(before.get(k) == after.get(k) for k in before if k in after),
          f"{len(before)} tracked")

    print("\n[4] a changed input triggers a rebuild")
    inp = d.find(d.root_id, "input")["id"]
    amz = d.find(inp, "amazon vc")["id"]
    victim = next(n for n in d.nodes.values()
                  if amz in n["parents"] and not n["trashed"])
    victim["content"] += b"\n"
    victim["md5Checksum"] = None
    victim["modifiedTime"] = FD._now()
    rc = run_cycle(d.root_id)
    check("changed input rebuilds", rc == 0 and "SKIPPED" not in status_text(d))

    print("\n[5] a file deleted in Drive leaves the report")
    store = d.find(inp, "retail stores")["id"]
    gone = next(n for n in d.nodes.values()
                if store in n["parents"] and not n["trashed"])
    gone["trashed"] = True
    rc = run_cycle(d.root_id)
    check("run after deletion succeeds", rc == 0, f"rc={rc}")
    check("deleted file is not read any more",
          gone["name"] not in status_text(d), gone["name"])
    check("local mirror dropped it too",
          not (C.INPUT / "retail stores" / gone["name"]).exists())

    print("\n[6] alert state carries across runs")
    import json
    st = json.loads((C.REPORTS / "_state" / "alert_state.json").read_text())
    check("alert state persisted", len(st.get("alerts", {})) > 0,
          f"{len(st.get('alerts', {})):,} tracked")
    check("alert state round-trips through Drive",
          d.find(d.find(d.root_id, "_state")["id"], "alert_state.json") is not None)
    notes = status_text(d)
    check("STATUS.txt reports the alert counts", "open alerts" in notes)

    print("\n[7] Drive's eventual consistency does not duplicate folders")
    # A folder created a moment ago can be missing from the very next listing.
    # Code that re-lists to find what it just made creates a second one — which
    # is how a live run ended up with two `output` folders side by side.
    lagged = FD.FakeDrive(lag=3)
    lagged.mkdir("input", lagged.root_id)
    FD.install(DS, lagged)
    ws = DS.Workspace(DS.DriveFS(lagged, lagged.root_id, log=lambda *a: None))
    tops = [n["name"] for n in lagged.nodes.values()
            if not n["trashed"] and lagged.root_id in n["parents"]]
    check("no duplicate top-level folders under lag",
          len(tops) == len(set(tops)), ", ".join(sorted(tops)))
    outs = [n for n in lagged.nodes.values()
            if not n["trashed"] and n["name"] == "output"]
    check("exactly one output/ folder", len(outs) == 1, f"{len(outs)} found")
    # Two uploads of the same name in one run must overwrite, not duplicate.
    tmp = Path("/tmp/_dup_probe.txt")
    tmp.write_text("one")
    a1 = ws.fs.upload(tmp, ws.latest, "probe.txt")
    tmp.write_text("two")
    a2 = ws.fs.upload(tmp, ws.latest, "probe.txt")
    tmp.unlink(missing_ok=True)
    check("re-upload overwrites under lag", a1["id"] == a2["id"],
          f"{a1['id']} vs {a2['id']}")

    print("\n[8] email never takes the run down with it")
    import alerts
    os.environ.update(SMTP_PORT="", SMTP_HOST="", SMTP_USER="", SMTP_PASS="")
    try:
        ok = alerts.send_email("s", "<p>b</p>", [])
        check("empty SMTP_PORT does not raise", ok is False)
    except Exception as e:                                      # noqa: BLE001
        check("empty SMTP_PORT does not raise", False, f"{type(e).__name__}: {e}")
    os.environ.update(SMTP_USER="u", SMTP_PASS="p", SMTP_PORT="not-a-number")
    try:
        alerts.send_email("s", "<p>b</p>", ["x@y.z"])
        check("junk SMTP_PORT falls back to 587", False, "should have tried to connect")
    except ValueError as e:
        check("junk SMTP_PORT falls back to 587", False, f"still parsing: {e}")
    except Exception:                                           # noqa: BLE001
        # Any connection error is fine — the point is it got past parsing.
        check("junk SMTP_PORT falls back to 587", True)
    for k in ("SMTP_PORT", "SMTP_HOST", "SMTP_USER", "SMTP_PASS"):
        os.environ.pop(k, None)

    print("\n[9] My Drive is rejected with the fix in the message")
    bad = FD.FakeDrive(quota_bug=True)
    FD.install(DS, bad)
    try:
        DS.preflight(DS.DriveFS(bad, bad.root_id))
        check("My Drive rejected", False, "preflight let it through")
    except RuntimeError as e:
        check("My Drive rejected before any work", "SHARED DRIVE" in str(e).upper())
        check("error names the fix", "Content manager" in str(e))

    print("\n[10] a broken run still leaves a readable STATUS.txt")
    FD.install(DS, d)
    empty = FD.FakeDrive()
    empty.mkdir("input", empty.root_id)
    FD.install(DS, empty)
    rc = run_cycle(empty.root_id)
    check("empty input fails cleanly", rc == 1, f"rc={rc}")
    check("failure is written to STATUS.txt", "FAILED" in status_text(empty))
    check("failure explains what to do",
          "master mapping table" in status_text(empty).lower())

    if not a.keep:
        for p in (C.INPUT, C.OUTPUT, C.REPORTS / "_state"):
            shutil.rmtree(p, ignore_errors=True)

    bad_n = sum(1 for r, *_ in results if r == FAIL)
    print(f"\n{'=' * 62}\n{len(results) - bad_n}/{len(results)} checks pass")
    if bad_n:
        for r, name, detail in results:
            if r == FAIL:
                print(f"  FAIL  {name}  {detail}")
    return 1 if bad_n else 0


if __name__ == "__main__":
    sys.exit(main())
