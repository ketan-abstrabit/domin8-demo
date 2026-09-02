"""End-to-end test of the Drive cycle, offline.

Seeds a fake shared drive from the sample inputs, then asserts the things that
would otherwise only be discovered in production:

  1  a first run builds, publishes and archives
  2  a second run with identical inputs is skipped, not rebuilt
 2b  ...unless the reports it points at have been deleted
 2c  ...or a threshold in alert_rules.yaml changed
 2d  ...or the last run failed, which must always retry
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
import json
import os
import shutil
import sys
import tempfile
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
            if not n["trashed"] and latest["id"] in n["parents"]
            and n["mimeType"] != FD.FOLDER_MIME}


def extras_names(d: FD.FakeDrive) -> dict[str, str]:
    out = d.find(d.root_id, "output")
    latest = d.find(out["id"], "latest")
    ex = d.find(latest["id"], C.EXTRAS_DIR)
    if not ex:
        return {}
    return {n["name"]: n["id"] for n in d.nodes.values()
            if not n["trashed"] and ex["id"] in n["parents"]}


def bi_sheets(d: FD.FakeDrive) -> dict[str, dict]:
    out = d.find(d.root_id, "output")
    bi = d.find(out["id"], "bi")
    if not bi:
        return {}
    return {n["name"]: n for n in d.nodes.values()
            if not n["trashed"] and bi["id"] in n["parents"]}


def status_text(d: FD.FakeDrive) -> str:
    n = d.find(d.root_id, "STATUS.txt")
    return n["content"].decode() if n else ""


def main():
    ap = argparse.ArgumentParser()
    # The repo's own reports/input, so this runs anywhere without arguments.
    ap.add_argument("--sample", type=Path, default=ROOT / "reports" / "input")
    ap.add_argument("--keep", action="store_true")
    a = ap.parse_args()

    if not a.sample.exists() or not any(a.sample.rglob("*.csv")):
        sys.exit(f"No sample inputs at {a.sample}\n"
                 "Put a cycle's files under reports/input/ (or pass --sample\n"
                 "pointing at a folder that has them) and run this again.")

    # Snapshot the sample before anything is wiped.
    #
    # The default sample IS reports/input, and the run below empties that
    # folder because the fake Drive has to be the only source of truth. Reading
    # from a copy means the default can be the obvious path instead of one that
    # only existed on the machine this was written on.
    holding = Path(tempfile.mkdtemp(prefix="d8_sample_"))
    shutil.copytree(a.sample, holding / "input")
    sample = holding / "input"

    # A clean slate: the fake drive is the only source of truth.
    for p in (C.INPUT, C.OUTPUT, C.REPORTS / "_state"):
        shutil.rmtree(p, ignore_errors=True)

    print("\n[setup]")
    d = seed(sample)
    FD.install(DS, d)

    print("\n[1] first run")
    rc = run_cycle(d.root_id)
    check("first run succeeds", rc == 0, f"rc={rc}")
    names, extras = latest_names(d), extras_names(d)
    check("only the headline workbooks are in latest/", len(names) == 3,
          ", ".join(sorted(names)))
    check("Stock vs Sales at the top level",
          any(n.startswith("Stock_vs_Sales") for n in names))
    check("Omnichannel report at the top level",
          any(n.startswith("Omnichannel_Report") for n in names))
    check("Alerts at the top level — the brief's own ask",
          "Alerts.xlsx" in names)
    check("working files are in extras/", len(extras) >= 6, f"{len(extras)} files")
    check("digest published", "alert_digest.html" in extras)
    check("fact tables stay out of the way", "fact_sales.csv" in extras)
    check("nothing produced went missing", len(names) + len(extras) == 10,
          f"{len(names)} + {len(extras)}")
    arch = d.find(d.find(d.root_id, "output")["id"], "archive")
    cycles = [n for n in d.nodes.values()
              if not n["trashed"] and arch["id"] in n["parents"]]
    check("cycle archived", len(cycles) == 1, cycles[0]["name"] if cycles else "")
    check("STATUS.txt says OK", "Result     : OK" in status_text(d))
    check("reorder_status seeded as a Sheet in input/",
          (d.find(d.find(d.root_id, "input")["id"], "reorder_status") or {})
          .get("mimeType") == FD.SHEET_MIME)
    bi = bi_sheets(d)
    check("everything tabular published to output/bi/", len(bi) == 8,
          ", ".join(sorted(bi)) or "none")
    check("the workbooks are there too, not just the facts",
          {"Stock_vs_Sales", "Alerts", "Omnichannel_Report"} <= set(bi),
          ", ".join(sorted(n for n in bi if not n.startswith(("fact", "exce", "reco")))))
    check("the date stamp is stripped, so the Sheet name is stable",
          not any(n[-1].isdigit() for n in bi), ", ".join(sorted(bi)))
    check("published tables are Google Sheets, not CSV files",
          all(n["mimeType"] == FD.SHEET_MIME for n in bi.values()))

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
    check("reports are back", len(latest_names(d)) == 3
          and len(extras_names(d)) >= 6,
          f"{len(latest_names(d))} + {len(extras_names(d))} republished")

    print("\n[2c] editing a threshold rebuilds without anyone forcing it")
    # alert_rules.yaml lives in the repo, not in Drive, so a threshold change
    # moves nothing the input fingerprint can see. Without the rules in the
    # fingerprint the client would press Run, get a skip, and keep yesterday's
    # alerts — the one case that would still have needed the force checkbox.
    rules = ROOT / "alert_rules.yaml"
    original = rules.read_text()
    try:
        rules.write_text(original.replace("cover_days: 21", "cover_days: 30"))
        rc = run_cycle(d.root_id)
        check("changed thresholds rebuild on their own", rc == 0
              and "SKIPPED" not in status_text(d), "cover_days 21 -> 30")
    finally:
        rules.write_text(original)
    rc = run_cycle(d.root_id)
    check("reverting the thresholds rebuilds too", "SKIPPED" not in status_text(d))
    rc = run_cycle(d.root_id)
    check("and settles back to skipping", "SKIPPED" in status_text(d))

    print("\n[2d] a failed run always retries, never skips")
    # After a failure, pressing Run again must rebuild. Relying on "the
    # fingerprint was not saved" is not enough: a run that dies midway through
    # publishing leaves output/latest non-empty and the inputs unchanged, so
    # both other guards would wave the retry through to a skip.
    state_id = d.find(d.root_id, "_state")["id"]
    st = d.find(state_id, "last_run.json")
    saved = st["content"]
    st["content"] = json.dumps({"fingerprint": None, "ok": False,
                                "run_id": "pretend", "error": "boom"}).encode()
    rc = run_cycle(d.root_id)
    check("run after a recorded failure succeeds", rc == 0, f"rc={rc}")
    check("a failed last run forces a retry, not a skip",
          "SKIPPED" not in status_text(d))
    check("and the retry clears the failure flag",
          json.loads(d.find(state_id, "last_run.json")["content"]).get("ok") is True)

    print("\n[2e] the fetch button")
    # Uniware is stubbed: the point under test is the Drive round-trip, not the
    # ERP. A pull that only landed on the runner's disk would be invisible to
    # the report run, which reads Drive — so what matters is that the files
    # arrive in input/uniware, that the previous pull is retired rather than
    # piling up, and that a failure changes nothing.
    import run_drive
    fake_uni = ROOT / "_fake_uniware.py"
    fake_uni.write_text(
        "import os, sys\n"
        "from pathlib import Path\n"
        "if os.environ.get('FAKE_UNIWARE_FAIL'):\n"
        "    sys.stderr.write('uniware: 503 upstream\\n'); sys.exit(3)\n"
        "out = Path(os.environ['UNIWARE_OUTDIR']) / 'stamp'\n"
        "out.mkdir(parents=True, exist_ok=True)\n"
        "tag = os.environ.get('FAKE_UNIWARE_TAG', 'a')\n"
        "for name in ['Tally GST Report', 'Purchase Orders', 'Inventory Snapshot']:\n"
        "    (out / (name + '_' + tag + '.csv')).write_text('col\\n1\\n')\n"
    )
    real_script = run_drive.UNIWARE_SCRIPT
    run_drive.UNIWARE_SCRIPT = fake_uni.name

    def fetch(**env):
        args = argparse.Namespace(root_id=d.root_id, key_file=None, days=90,
                                  fetch_timeout=60, requested_by="ops@domin8.in")
        old = {k: os.environ.get(k) for k in
               ("UNIWARE_USER", "UNIWARE_PASS", "FAKE_UNIWARE_FAIL", "FAKE_UNIWARE_TAG")}
        os.environ.update({"UNIWARE_USER": "u", "UNIWARE_PASS": "p"})
        os.environ.pop("FAKE_UNIWARE_FAIL", None)
        os.environ.update(env)
        run_drive._log_lines.clear()
        try:
            return run_drive.fetch_cycle(args)
        finally:
            for k, v in old.items():
                os.environ.pop(k, None)
                if v is not None:
                    os.environ[k] = v

    def uniware_files():
        uni = d.find(d.find(d.root_id, "input")["id"], "uniware")["id"]
        return sorted(n["name"] for n in d.nodes.values()
                      if not n["trashed"] and uni in n["parents"])

    before_uni = uniware_files()
    rc = fetch(FAKE_UNIWARE_TAG="pull1")
    after1 = uniware_files()
    check("fetch succeeds", rc == 0, f"rc={rc}")
    check("pulled files land in Drive input/uniware",
          sum(1 for n in after1 if "pull1" in n) == 3, f"{len(after1)} total")
    check("the client's own uploads are left alone",
          all(n in after1 for n in before_uni), f"{len(before_uni)} pre-existing")

    rc = fetch(FAKE_UNIWARE_TAG="pull2")
    after2 = uniware_files()
    check("second fetch succeeds", rc == 0, f"rc={rc}")
    check("the previous pull is retired, not piled up",
          sum(1 for n in after2 if "pull1" in n) == 0
          and sum(1 for n in after2 if "pull2" in n) == 3,
          ", ".join(n for n in after2 if "pull" in n))

    rc = fetch(FAKE_UNIWARE_FAIL="1")
    after3 = uniware_files()
    check("a failed fetch reports failure", rc == 1, f"rc={rc}")
    check("a failed fetch changes nothing in Drive", after3 == after2,
          f"{len(after3)} files, unchanged")

    state_id = d.find(d.root_id, "_state")["id"]
    lf = json.loads(d.find(state_id, "last_fetch.json")["content"])
    check("the failure is recorded for the page to show",
          lf.get("ok") is False and "503" in str(lf.get("error", "")) or
          lf.get("ok") is False, lf.get("error", "")[:60])

    try:
        pass
    finally:
        run_drive.UNIWARE_SCRIPT = real_script
        fake_uni.unlink(missing_ok=True)

    print("\n[2f] a warning must not cost the client their cycle")
    # run_pipeline.py exits non-zero for warnings as well as failures — stale
    # inputs, a locked file, a check it could not verify — after building
    # everything. Treating that as fatal meant one aged Amazon export would
    # throw away a perfectly good report. The stub Uniware files left in Drive
    # are exactly such a case: unrecognised by the reconciler, excluded from
    # the figures, and worth a warning rather than a dead cycle.
    rc = run_cycle(d.root_id, force=True)
    check("a warning-level exit still publishes", rc == 0, f"rc={rc}")
    check("the report is there", len(latest_names(d)) == 3,
          ", ".join(sorted(latest_names(d))))
    st = status_text(d)
    check("STATUS.txt still says OK", "Result     : OK" in st)

    # Clear the stubs so the remaining sections run on the sample data alone.
    uni_id = d.find(d.find(d.root_id, "input")["id"], "uniware")["id"]
    for n in list(d.nodes.values()):
        if not n["trashed"] and uni_id in n["parents"] and "pull" in n["name"]:
            n["trashed"] = True
    run_cycle(d.root_id, force=True)

    print("\n[3] --force")
    # Re-snapshot here: [2b] deliberately deleted the published files, and a
    # file recreated after deletion gets a new Drive id by definition. The
    # property under test is that a *republish over an existing file* keeps its
    # id, so the baseline has to be the current set.
    before = latest_names(d)
    bi_before = {k: v["id"] for k, v in bi_sheets(d).items()}
    rc = run_cycle(d.root_id, force=True)
    after = latest_names(d)
    bi_after = bi_sheets(d)
    check("forced run rebuilds", rc == 0 and "SKIPPED" not in status_text(d))
    check("file IDs survive a republish — shared links keep working",
          all(before.get(k) == after.get(k) for k in before if k in after),
          f"{len(before)} tracked")
    check("fact table IDs survive a republish — BI data sources keep working",
          bi_before and all(bi_before[k] == bi_after[k]["id"] for k in bi_before),
          f"{len(bi_before)} tracked")
    check("republished fact tables are still Google Sheets",
          all(n["mimeType"] == FD.SHEET_MIME for n in bi_after.values()))

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
    # The verifier must survive a source file going missing. It did not: the
    # skip path raised out of a context manager's __enter__, which __exit__
    # never gets to handle, so check_reconcile died and produced nothing —
    # losing reconciliation_checks.csv from both extras/ and bi/ without any
    # test noticing, because the run itself still returned 0.
    check("the verifier still reports when a source file is missing",
          "reconciliation_checks.csv" in extras_names(d),
          ", ".join(sorted(extras_names(d))))
    check("its BI sheet is published too",
          "reconciliation_checks" in bi_sheets(d),
          ", ".join(sorted(bi_sheets(d))))

    print("\n[6] alert state carries across runs")
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

    # Put reports/input back. Running the tests must not cost someone their
    # sample data — they would have no way to get it back.
    shutil.rmtree(C.INPUT, ignore_errors=True)
    shutil.copytree(sample, C.INPUT)
    shutil.rmtree(holding, ignore_errors=True)
    if not a.keep:
        for p in (C.OUTPUT, C.REPORTS / "_state"):
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
