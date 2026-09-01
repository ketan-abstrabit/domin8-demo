"""Google Drive as the interface.

The client never sees the code and never opens an app. They drop files into a
Drive folder; the finished reports appear in the same folder. This module is the
only piece that knows Drive exists — everything downstream still reads and
writes ordinary files on disk, so `reconcile.py` and `stock_vs_sales.py` are
untouched.

Expected layout in Drive (created on first run if missing):

    <root>/
      input/
        <master mapping table>          any spreadsheet sitting loose in input/
        uniware/
        amazon vc/
        retail stores/
        reorder_status                  Google Sheet, seeded once, edited by hand
      output/
        latest/                         overwritten every run, file IDs kept
        archive/2026-08-31_1430/        every previous run
        bi/                             fact tables as Sheets, for BI tools
      STATUS.txt                        last run: when, what it read, pass/fail
      _state/                           fingerprints, PO history, alert state

SHARED DRIVE, NOT MY DRIVE
--------------------------
A service account has no Drive storage quota of its own. Put this tree in a
*Shared Drive* and files are owned by the drive, so the account can write. In a
personal My Drive the same call fails with 403 storageQuotaExceeded even though
the folder is shared with edit rights. `preflight()` checks this and says so in
plain words rather than letting the run die three minutes in.
"""

from __future__ import annotations

import hashlib
import io
import json
import os
import re
import shutil
import time
from datetime import datetime, timezone
from pathlib import Path

FOLDER_MIME = "application/vnd.google-apps.folder"
SHEET_MIME = "application/vnd.google-apps.spreadsheet"
DOC_MIME = "application/vnd.google-apps.document"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
CSV_MIME = "text/csv"

SCOPES = ["https://www.googleapis.com/auth/drive"]

# Google-native files have no bytes to download; they are exported instead.
EXPORT_AS = {
    SHEET_MIME: (XLSX_MIME, ".xlsx"),
    DOC_MIME: ("application/pdf", ".pdf"),
}

# What the client's folders may plausibly be called. Matched on a normalised
# name so "Amazon VC", "amazon_vc" and "Amazon Vendor Central" all land in the
# same place. The first spelling is the one we create if nothing matches.
FOLDER_ALIASES = {
    "uniware": ["uniware", "unicommerce", "uniware exports"],
    "amazon vc": ["amazon vc", "amazon", "amazonvc", "amazon vendor central",
                  "vendor central", "amazon vendor"],
    "retail stores": ["retail stores", "retail", "stores", "store",
                      "retail store", "offline stores"],
}

MAX_UPLOAD_RETRIES = 4


def _norm(name: str) -> str:
    """Fold case, punctuation and spacing so folder names match loosely."""
    return re.sub(r"[^a-z0-9]+", " ", str(name).lower()).strip()


def _now() -> datetime:
    return datetime.now(timezone.utc)


def to_local(dt: datetime) -> datetime:
    """Render a moment in the timezone the reports are read in.

    `astimezone()` with no argument uses the machine's clock, which on a CI
    runner is UTC — so every timestamp a human sees would be shifted, and after
    18:30 IST also dated to the previous day. Falls back to a fixed +05:30 if
    the zone database is missing from a slim image.
    """
    name = getattr(_cfg(), "REPORT_TZ", None)
    if not name:
        return dt.astimezone()
    try:
        from zoneinfo import ZoneInfo
        return dt.astimezone(ZoneInfo(name))
    except Exception:                                       # noqa: BLE001
        from datetime import timedelta
        return dt.astimezone(timezone(timedelta(hours=5, minutes=30), "IST"))


def _cfg():
    try:
        import pipeline_config
        return pipeline_config
    except Exception:                                       # noqa: BLE001
        return None


# ---------------------------------------------------------------------------
# auth
# ---------------------------------------------------------------------------

def build_service(key_json: str | None = None, key_file: str | None = None):
    """Authenticate as the service account and return a Drive v3 client.

    Credentials come from, in order: the argument, GOOGLE_SA_KEY (the JSON
    itself, which is how GitHub Actions passes a secret), or
    GOOGLE_APPLICATION_CREDENTIALS / --key-file (a path, for local runs).
    """
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    raw = key_json or os.environ.get("GOOGLE_SA_KEY")
    if raw and raw.strip().startswith("{"):
        info = json.loads(raw)
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=SCOPES)
    else:
        path = key_file or os.environ.get("GOOGLE_APPLICATION_CREDENTIALS") or raw
        if not path:
            raise RuntimeError(
                "No Google credentials. Set GOOGLE_SA_KEY to the service-account "
                "JSON, or GOOGLE_APPLICATION_CREDENTIALS to a path to it.")
        if not Path(path).exists():
            raise RuntimeError(f"Service-account key not found: {path}")
        creds = service_account.Credentials.from_service_account_file(
            path, scopes=SCOPES)

    return build("drive", "v3", credentials=creds, cache_discovery=False)


def media_upload(path: Path, mimetype: str | None = None):
    """Wrapped so the offline test harness can swap in its own uploader."""
    from googleapiclient.http import MediaFileUpload
    return MediaFileUpload(str(path), mimetype=mimetype, resumable=False)


# ---------------------------------------------------------------------------
# a thin, shared-drive-aware wrapper over the Drive API
# ---------------------------------------------------------------------------

class DriveFS:
    """Path-like access to one Drive subtree.

    Every call carries supportsAllDrives / includeItemsFromAllDrives, without
    which a Shared Drive looks empty and every write fails.
    """

    def __init__(self, service, root_id: str, log=print):
        self.svc = service
        self.root_id = root_id
        self.log = log
        self._children: dict[str, list[dict]] = {}   # parent id -> listing

    # -- reading ----------------------------------------------------------

    def children(self, parent_id: str, refresh: bool = False) -> list[dict]:
        if refresh or parent_id not in self._children:
            out, token = [], None
            while True:
                resp = self.svc.files().list(
                    q=f"'{parent_id}' in parents and trashed = false",
                    fields=("nextPageToken, files(id, name, mimeType, size, "
                            "modifiedTime, md5Checksum)"),
                    pageSize=1000,
                    pageToken=token,
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True,
                ).execute()
                out.extend(resp.get("files", []))
                token = resp.get("nextPageToken")
                if not token:
                    break
            self._children[parent_id] = out
        return self._children[parent_id]

    def find(self, parent_id: str, name: str, folder: bool | None = None):
        """Find one child by name, tolerant of case and punctuation."""
        target = _norm(name)
        for f in self.children(parent_id):
            if folder is True and f["mimeType"] != FOLDER_MIME:
                continue
            if folder is False and f["mimeType"] == FOLDER_MIME:
                continue
            if _norm(f["name"]) == target:
                return f
        return None

    def find_folder_aliased(self, parent_id: str, key: str):
        """Find a folder by any of its known spellings (see FOLDER_ALIASES)."""
        wanted = {_norm(a) for a in FOLDER_ALIASES.get(key, [key])}
        for f in self.children(parent_id):
            if f["mimeType"] == FOLDER_MIME and _norm(f["name"]) in wanted:
                return f
        return None

    # -- writing ----------------------------------------------------------

    def _remember(self, parent_id: str, node: dict):
        """Record a just-written child in the cache.

        Drive's file list is eventually consistent: a folder created a moment
        ago may be missing from the very next list call. Dropping the cache and
        re-listing therefore does not find it, and the caller creates a second
        one — which is exactly how a run ended up with two `output` folders.
        Adding the node we already hold is both correct and one call cheaper.
        """
        kids = self._children.get(parent_id)
        if kids is None:
            return
        for i, existing in enumerate(kids):
            if existing["id"] == node["id"]:
                kids[i] = {**existing, **node}
                return
        kids.append(node)

    def ensure_folder(self, parent_id: str, name: str) -> str:
        hit = self.find(parent_id, name, folder=True)
        if hit:
            return hit["id"]
        made = self.svc.files().create(
            body={"name": name, "mimeType": FOLDER_MIME, "parents": [parent_id]},
            fields="id, name, mimeType",
            supportsAllDrives=True,
        ).execute()
        self._remember(parent_id, {"id": made["id"], "name": name,
                                   "mimeType": FOLDER_MIME})
        self.log(f"      created folder  {name}")
        return made["id"]

    def ensure_path(self, *parts: str) -> str:
        """ensure_path('output', 'latest') -> folder id, creating as needed."""
        node = self.root_id
        for p in parts:
            node = self.ensure_folder(node, p)
        return node

    def download(self, meta: dict, dest_dir: Path,
                 export_mime: str | None = None) -> Path | None:
        """Fetch one file. Google-native files are exported to a real format.

        `export_mime` overrides the default for one call — the override sheet
        comes back as CSV rather than xlsx, which is what the pipeline reads
        and skips a needless spreadsheet round-trip.
        """
        dest_dir.mkdir(parents=True, exist_ok=True)
        name, mime = meta["name"], meta["mimeType"]

        if mime in EXPORT_AS:
            out_mime, ext = EXPORT_AS[mime]
            if export_mime:
                out_mime = export_mime
                ext = {CSV_MIME: ".csv", XLSX_MIME: ".xlsx"}.get(export_mime, ext)
            if not name.lower().endswith(ext):
                name += ext
            data = self.svc.files().export_media(
                fileId=meta["id"], mimeType=out_mime).execute()
        elif mime == FOLDER_MIME:
            return None
        else:
            data = self.svc.files().get_media(
                fileId=meta["id"], supportsAllDrives=True).execute()

        if isinstance(data, io.BytesIO):
            data = data.getvalue()
        dest = dest_dir / name
        dest.write_bytes(data)
        return dest

    def upload(self, local: Path, parent_id: str, name: str | None = None,
               convert_to: str | None = None) -> dict:
        """Create or overwrite a file by name.

        Overwrites in place with files().update so the Drive file ID survives —
        anyone who bookmarked last week's report or embedded its link keeps a
        working link instead of a 404 every fortnight.

        `convert_to` applies on update as well as create: without the mimeType
        on the update body, re-uploading CSV bytes over a Google Sheet turns it
        back into a plain CSV and every data source bound to it breaks.
        """
        name = name or local.name
        existing = self.find(parent_id, name, folder=False)
        media = media_upload(local, mimetype=_guess_mime(local))

        for attempt in range(MAX_UPLOAD_RETRIES):
            try:
                if existing:
                    got = self.svc.files().update(
                        fileId=existing["id"],
                        body={"mimeType": convert_to} if convert_to else None,
                        media_body=media,
                        fields="id, name, modifiedTime",
                        supportsAllDrives=True,
                    ).execute()
                    mime = convert_to or existing["mimeType"]
                else:
                    body = {"name": name, "parents": [parent_id]}
                    if convert_to:
                        body["mimeType"] = convert_to
                    got = self.svc.files().create(
                        body=body, media_body=media,
                        fields="id, name, modifiedTime",
                        supportsAllDrives=True,
                    ).execute()
                    mime = convert_to or _guess_mime(local) or "application/octet-stream"
                # Same consistency trap as folders: a second upload of the same
                # name in one run must find this file, not create a duplicate.
                self._remember(parent_id, {
                    "id": got["id"], "name": name, "mimeType": mime,
                    "modifiedTime": got.get("modifiedTime"),
                })
                return got
            except Exception as e:                       # noqa: BLE001
                msg = str(e)
                if "storageQuotaExceeded" in msg:
                    raise RuntimeError(QUOTA_HELP) from e
                if attempt == MAX_UPLOAD_RETRIES - 1:
                    raise
                time.sleep(2 ** attempt)
        raise RuntimeError("unreachable")

    def write_text(self, parent_id: str, name: str, text: str) -> dict:
        tmp = Path(os.environ.get("TMPDIR", "/tmp")) / f"_ds_{os.getpid()}_{name}"
        tmp.parent.mkdir(parents=True, exist_ok=True)
        tmp.write_text(text, encoding="utf-8")
        try:
            return self.upload(tmp, parent_id, name)
        finally:
            tmp.unlink(missing_ok=True)

    def read_text(self, parent_id: str, name: str) -> str | None:
        hit = self.find(parent_id, name, folder=False)
        if not hit:
            return None
        data = self.svc.files().get_media(
            fileId=hit["id"], supportsAllDrives=True).execute()
        if isinstance(data, io.BytesIO):
            data = data.getvalue()
        return data.decode("utf-8", errors="replace")

    def trash(self, file_id: str, parent_id: str | None = None):
        self.svc.files().update(fileId=file_id, body={"trashed": True},
                                supportsAllDrives=True).execute()
        if parent_id:
            self._children.pop(parent_id, None)


QUOTA_HELP = """
Drive refused the upload: storageQuotaExceeded.

This almost never means the drive is full. A service account has no storage
quota of its own, so it cannot own files. Uploading into a folder that lives in
somebody's My Drive fails this way even when the folder is shared with edit
rights, because the new file would be owned by the service account.

The fix: move the reporting folder into a SHARED DRIVE (Drive > Shared drives >
New), then add the service-account email as a Content manager. Files in a
shared drive are owned by the drive, so the upload succeeds.
""".strip()


def _guess_mime(path: Path) -> str | None:
    import mimetypes
    ext = path.suffix.lower()
    fixed = {
        ".xlsx": XLSX_MIME,
        ".xls": "application/vnd.ms-excel",
        ".csv": CSV_MIME,
        ".html": "text/html",
        ".txt": "text/plain",
        ".json": "application/json",
        ".yaml": "text/yaml",
        ".yml": "text/yaml",
    }
    return fixed.get(ext) or mimetypes.guess_type(path.name)[0]


# ---------------------------------------------------------------------------
# the tree
# ---------------------------------------------------------------------------

class Workspace:
    """The DOMIN8 folder tree, resolved once and reused."""

    def __init__(self, fs: DriveFS, log=print):
        self.fs = fs
        self.log = log
        self.input = fs.ensure_path("input")
        self.output = fs.ensure_path("output")
        self.latest = fs.ensure_path("output", "latest")
        self.archive = fs.ensure_path("output", "archive")
        self.bi = fs.ensure_path("output", "bi")
        self.state = fs.ensure_path("_state")
        self.sub = {}
        for key in FOLDER_ALIASES:
            hit = fs.find_folder_aliased(self.input, key)
            self.sub[key] = hit["id"] if hit else fs.ensure_folder(self.input, key)

    # -- state ------------------------------------------------------------

    def load_state(self, name: str, default=None):
        raw = self.fs.read_text(self.state, name)
        if raw is None:
            return default
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            self.log(f"  ! {name} in _state is not valid JSON — starting fresh")
            return default

    def save_state(self, name: str, obj):
        self.fs.write_text(self.state, name,
                           json.dumps(obj, indent=2, default=str))

    def load_state_file(self, name: str, dest: Path) -> bool:
        """Pull a binary/CSV state file (PO history) down to disk."""
        hit = self.fs.find(self.state, name, folder=False)
        if not hit:
            return False
        got = self.fs.download(hit, dest.parent)
        if got and got != dest:
            shutil.move(str(got), str(dest))
        return True

    def save_state_file(self, local: Path, name: str | None = None):
        if local.exists():
            self.fs.upload(local, self.state, name or local.name)


def preflight(fs: DriveFS) -> dict:
    """Confirm we can see the folder and that it lives on a Shared Drive.

    Runs before anything expensive, because the two ways this setup goes wrong
    (folder not shared, folder in My Drive) both produce confusing errors much
    later otherwise.
    """
    meta = fs.svc.files().get(
        fileId=fs.root_id,
        fields="id, name, mimeType, driveId, capabilities(canAddChildren)",
        supportsAllDrives=True,
    ).execute()

    if meta["mimeType"] != FOLDER_MIME:
        raise RuntimeError(f"DRIVE_ROOT_ID points at a file, not a folder: {meta['name']}")
    if not meta.get("capabilities", {}).get("canAddChildren", False):
        raise RuntimeError(
            f"The service account can see '{meta['name']}' but cannot write to it. "
            "Give it Content manager (not Viewer/Commenter) on the shared drive.")
    if not meta.get("driveId"):
        raise RuntimeError(QUOTA_HELP)
    return meta


# ---------------------------------------------------------------------------
# pull
# ---------------------------------------------------------------------------

# Anything the client leaves lying around that is not a report.
IGNORE_NAMES = {"read me - where to put files.txt", "readme.txt", "read me.txt"}


def pull_inputs(ws: Workspace, local_input: Path, overrides: Path | None = None,
                log=print) -> dict:
    """Mirror Drive's input/ onto local disk. Returns a manifest for fingerprinting.

    The merchandiser override sheet is the one exception. It lives in input/ so
    the team can find and edit it, but it is not a report: exported from a
    Google Sheet it arrives as `reorder_status.xlsx`, a name the reconciler
    does not recognise and would list as an unknown file. So it is diverted out
    of the input folder and converted to the CSV the pipeline expects — while
    still appearing in the manifest, so that editing it triggers a rebuild.
    """
    fs = ws.fs
    manifest: list[dict] = []

    def take(meta: dict, dest_dir: Path, label: str):
        if meta["mimeType"] == FOLDER_MIME:
            return
        if _norm(meta["name"]) in {_norm(n) for n in IGNORE_NAMES}:
            return
        if overrides is not None and _norm(Path(meta["name"]).stem) == _norm(REORDER_SHEET):
            got = _take_overrides(fs, meta, overrides, log)
            label = "overrides"
        else:
            got = fs.download(meta, dest_dir)
        if got is None:
            return
        manifest.append({
            "folder": label,
            "name": got.name,
            "drive_name": meta["name"],
            "id": meta["id"],
            "modified": meta.get("modifiedTime"),
            "md5": meta.get("md5Checksum"),
            "bytes": got.stat().st_size,
        })
        log(f"      {label:<16} {got.stat().st_size:>10,}  {got.name}")

    local_input.mkdir(parents=True, exist_ok=True)

    log("    input/  (master mapping table and anything loose)")
    for meta in fs.children(ws.input):
        take(meta, local_input, "input")

    for key, folder_id in ws.sub.items():
        kids = [m for m in fs.children(folder_id) if m["mimeType"] != FOLDER_MIME]
        log(f"    {key}/  ({len(kids)} file(s))")
        dest = local_input / key
        for meta in kids:
            take(meta, dest, key)
        if not kids:
            dest.mkdir(parents=True, exist_ok=True)
            log("      (empty)")

    return {"files": sorted(manifest, key=lambda m: (m["folder"], m["name"]))}


def fingerprint(manifest: dict) -> str:
    """A stable hash of what we just read, so an unchanged folder can be skipped.

    Keyed on Drive's own file ID and modifiedTime rather than on the bytes, so
    it is cheap, and on md5 where Drive gives us one, so a re-upload of an
    identical file does not force a rebuild.
    """
    parts = []
    for f in manifest["files"]:
        parts.append(f"{f['folder']}/{f['name']}|{f['id']}|"
                     f"{f.get('md5') or f.get('modified')}|{f['bytes']}")
    return hashlib.sha256("\n".join(parts).encode()).hexdigest()[:16]


# ---------------------------------------------------------------------------
# push
# ---------------------------------------------------------------------------

ARCHIVE_KEEP = 12


def is_main_output(name: str, patterns: list[str]) -> bool:
    from fnmatch import fnmatch
    return any(fnmatch(name, p) for p in patterns)


def push_outputs(ws: Workspace, local_output: Path, stamp: str,
                 main_patterns: list[str] | None = None,
                 extras_dir: str = "extras", log=print) -> dict:
    """Publish output/ to Drive: latest/ overwritten, archive/<stamp>/ added.

    The two headline workbooks sit at the top of latest/; the supporting files
    go one level down. A folder holding ten files makes the client hunt for the
    two they came for, and the other eight are evidence rather than deliverables.
    """
    fs = ws.fs
    files = sorted(p for p in local_output.glob("*") if p.is_file())
    if not files:
        raise RuntimeError("the pipeline produced no output files")

    patterns = main_patterns or []
    arch = fs.ensure_folder(ws.archive, stamp)
    latest_extras = arch_extras = None
    links = {}

    for f in files:
        if is_main_output(f.name, patterns):
            dest, adest, where = ws.latest, arch, ""
        else:
            if latest_extras is None:
                latest_extras = fs.ensure_folder(ws.latest, extras_dir)
                arch_extras = fs.ensure_folder(arch, extras_dir)
            dest, adest, where = latest_extras, arch_extras, f"{extras_dir}/"
        got = fs.upload(f, dest)
        links[f.name] = f"https://drive.google.com/file/d/{got['id']}/view"
        fs.upload(f, adest)
        log(f"      {f.stat().st_size:>10,}  {where}{f.name}")

    tidy_latest(ws, files, patterns, extras_dir, log=log)
    prune_archive(ws, log=log)
    return links


def tidy_latest(ws: Workspace, files, patterns, extras_dir: str, log=print):
    """Remove anything in latest/ that this run did not produce.

    Without this, a file that stops being generated — or one that used to live
    at the top level before the split — lingers for ever, and the client reads
    a stale workbook believing it is current.
    """
    produced = {f.name for f in files}
    main_now = {f.name for f in files if is_main_output(f.name, patterns)}
    for node in list(ws.fs.children(ws.latest, refresh=True)):
        if node["mimeType"] == FOLDER_MIME:
            continue
        if node["name"] not in main_now:
            ws.fs.trash(node["id"], ws.latest)
            reason = ("moved to " + extras_dir if node["name"] in produced
                      else "no longer produced")
            log(f"      removed from latest/: {node['name']}  ({reason})")

# Republished as Google Sheets for BI tools, which read native Sheets rather
# than the .xlsx and .csv files in latest/.
#
# Everything tabular goes: the flat CSVs one worksheet each, and the workbooks
# converted whole — a Google Sheet keeps every tab, and a BI data source binds
# to one worksheet, so `Stock_vs_Sales` arrives with `sku wise` and
# `Article wise` each selectable. That is the richest table in the pipeline and
# leaving it out would have meant rebuilding it from the facts by hand.
#
# The two HTML files are not tabular and are not republished; they are already
# readable as they are.
BI_CSV = ("fact_sales.csv", "fact_inventory.csv", "fact_purchase.csv",
          "exceptions.csv", "reconciliation_checks.csv")

BI_WORKBOOKS = ("Stock_vs_Sales*.xlsx", "Alerts*.xlsx", "Omnichannel_Report*.xlsx")

# Kept for the tests and callers that predate the workbook split.
BI_TABLES = BI_CSV


def _bi_name(path: Path) -> str:
    """A stable Sheet name, so a date-stamped file keeps one data source.

    Stock_vs_Sales_310826.xlsx becomes `Stock_vs_Sales`. Without this every
    cycle would publish a new name, create a new file, and orphan every chart
    built on the last one.
    """
    stem = path.stem
    for marker in ("_",):
        parts = stem.split(marker)
        if len(parts) > 1 and parts[-1].isdigit():
            return marker.join(parts[:-1])
    return stem


def push_bi_tables(ws: Workspace, local_output: Path, log=print) -> dict:
    """Publish everything tabular to output/bi/ as Google Sheets.

    Created once, then overwritten in place. A BI data source binds to a file
    ID, so holding the ID steady is what keeps a dashboard working across runs
    without anyone re-linking it.
    """
    from fnmatch import fnmatch

    wanted = [local_output / n for n in BI_CSV]
    for pattern in BI_WORKBOOKS:
        wanted += sorted(p for p in local_output.glob("*.xlsx")
                         if fnmatch(p.name, pattern))

    links, seen = {}, set()
    for f in wanted:
        if not f.exists():
            log(f"      {f.name} not produced this run — skipped")
            continue
        sheet = _bi_name(f)
        if sheet in seen:
            continue
        seen.add(sheet)
        got = ws.fs.upload(f, ws.bi, sheet, convert_to=SHEET_MIME)
        links[sheet] = f"https://docs.google.com/spreadsheets/d/{got['id']}/edit"
        log(f"      {sheet}  (Google Sheet)")
    return links


def prune_archive(ws: Workspace, keep: int = ARCHIVE_KEEP, log=print):
    """Keep the last `keep` runs. Old cycles are trashed, not hard-deleted."""
    folders = [f for f in ws.fs.children(ws.archive, refresh=True)
               if f["mimeType"] == FOLDER_MIME]
    folders.sort(key=lambda f: f["name"], reverse=True)
    for old in folders[keep:]:
        ws.fs.trash(old["id"], ws.archive)
        log(f"      archived cycle removed: {old['name']}")


# ---------------------------------------------------------------------------
# the merchandiser's editable sheet
# ---------------------------------------------------------------------------

REORDER_SHEET = "reorder_status"


def _take_overrides(fs: DriveFS, meta: dict, dest: Path, log=print) -> Path | None:
    """Land the override sheet at `dest` as CSV, whatever shape it arrived in."""
    import tempfile
    dest.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        got = fs.download(meta, Path(tmp), export_mime=CSV_MIME)
        if got is None:
            return None
        if got.suffix.lower() in (".xlsx", ".xls"):
            import pandas as pd
            pd.read_excel(got).to_csv(dest, index=False)
        else:
            shutil.copy2(got, dest)
    log(f"    merchandiser overrides -> {dest.name} "
        f"(kept out of input/, it is not a report)")
    return dest


def push_reorder_sheet(ws: Workspace, local: Path, log=print):
    """Seed the sheet on the first run only; never overwrite their edits."""
    if not local.exists():
        return
    if (ws.fs.find(ws.input, REORDER_SHEET, folder=False)
            or ws.fs.find(ws.input, REORDER_SHEET + ".csv", folder=False)):
        return
    ws.fs.upload(local, ws.input, REORDER_SHEET, convert_to=SHEET_MIME)
    log("    seeded 'reorder_status' as a Google Sheet in input/ — "
        "merchandisers edit it there and the next run picks it up")


# ---------------------------------------------------------------------------
# STATUS.txt — the only thing the client reads when something looks wrong
# ---------------------------------------------------------------------------

def write_status(ws: Workspace, *, ok: bool, started: datetime, manifest: dict,
                 lines: list[str], skipped: bool = False):
    dur = (_now() - started).total_seconds()
    verdict = ("SKIPPED — inputs unchanged since the last run" if skipped
               else "OK" if ok else "FAILED")
    body = [
        "DOMIN8 omnichannel reporting",
        "=" * 60,
        f"Last run   : {to_local(started).strftime('%d %b %Y, %H:%M %Z')}",
        f"Result     : {verdict}",
        f"Duration   : {dur:,.0f}s",
        "",
        f"Files read ({len(manifest.get('files', []))}):",
    ]
    for f in manifest.get("files", []):
        body.append(f"  {f['folder']:<16} {f['bytes']:>10,}  {f['drive_name']}")
    if not manifest.get("files"):
        body.append("  (none — input folders are empty)")
    body += ["", "Notes:"] + [f"  {ln}" for ln in lines]
    body += ["",
             "Reports are in  output/latest/ .",
             "Previous cycles are kept in  output/archive/ ."]
    ws.fs.write_text(ws.fs.root_id, "STATUS.txt", "\n".join(body) + "\n")
