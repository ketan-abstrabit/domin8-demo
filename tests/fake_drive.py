"""An in-memory stand-in for the Drive v3 API.

The Drive layer is the one piece that cannot be exercised by running the
pipeline locally, and it is also the piece most likely to be wrong in ways that
only show up at 6am on a schedule — a missing supportsAllDrives, a file that
gets duplicated instead of overwritten, a fingerprint that never matches.

This implements just enough of the API for drive_sync to run against: the same
call signatures, the same response shapes, the same eventual-consistency-free
behaviour. It is not a Drive emulator and does not try to be; it exists so the
whole pull -> build -> push cycle can be asserted on, offline, in seconds.
"""

from __future__ import annotations

import hashlib
import itertools
import re
from datetime import datetime, timezone
from pathlib import Path

FOLDER_MIME = "application/vnd.google-apps.folder"
SHEET_MIME = "application/vnd.google-apps.spreadsheet"


def _now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%fZ")


class _Req:
    """Drive's API objects are lazy: build the call, then .execute() it."""

    def __init__(self, fn):
        self._fn = fn

    def execute(self):
        return self._fn()


class FakeMedia:
    def __init__(self, path, mimetype=None, resumable=False):
        self.path = Path(path)
        self.mimetype = mimetype


class FakeDrive:
    """One shared drive, one tree, no network."""

    def __init__(self, drive_id="fakeDrive01", quota_bug=False):
        self.drive_id = drive_id
        # quota_bug reproduces the My Drive failure mode, so the error path and
        # its help text are covered too rather than only the happy one.
        self.quota_bug = quota_bug
        self._ids = itertools.count(1)
        self.nodes: dict[str, dict] = {}
        self.calls: list[str] = []
        self.root_id = self._add("DOMIN8 Reporting", FOLDER_MIME, None)

    # -- construction helpers used by tests -------------------------------

    def _add(self, name, mime, parent, content=b"") -> str:
        fid = f"id{next(self._ids):04d}"
        self.nodes[fid] = {
            "id": fid, "name": name, "mimeType": mime,
            "parents": [parent] if parent else [],
            "trashed": False, "content": content,
            "modifiedTime": _now(),
            "md5Checksum": hashlib.md5(content).hexdigest() if content else None,
        }
        return fid

    def mkdir(self, name, parent=None) -> str:
        return self._add(name, FOLDER_MIME, parent or self.root_id)

    def put(self, local: Path, parent: str, name=None, mime=None) -> str:
        local = Path(local)
        return self._add(name or local.name,
                         mime or "application/octet-stream",
                         parent, local.read_bytes())

    def find(self, parent: str, name: str):
        for n in self.nodes.values():
            if not n["trashed"] and parent in n["parents"] and n["name"] == name:
                return n
        return None

    def tree(self, node=None, depth=0) -> list[str]:
        node = node or self.root_id
        out = []
        for n in sorted((x for x in self.nodes.values()
                         if not x["trashed"] and node in x["parents"]),
                        key=lambda x: (x["mimeType"] != FOLDER_MIME, x["name"])):
            mark = "/" if n["mimeType"] == FOLDER_MIME else ""
            size = "" if mark else f"  ({len(n['content']):,}b)"
            out.append("  " * depth + n["name"] + mark + size)
            if mark:
                out += self.tree(n["id"], depth + 1)
        return out

    # -- the API surface --------------------------------------------------

    def files(self):
        return _Files(self)


class _Files:
    def __init__(self, drive: FakeDrive):
        self.d = drive

    # ---- list ----
    def list(self, q="", fields=None, pageSize=None, pageToken=None,
             supportsAllDrives=False, includeItemsFromAllDrives=False, **kw):
        self.d.calls.append("list")
        assert supportsAllDrives and includeItemsFromAllDrives, (
            "a shared-drive listing without supportsAllDrives/"
            "includeItemsFromAllDrives silently returns nothing")
        m = re.search(r"'([^']+)' in parents", q or "")
        parent = m.group(1) if m else None

        def go():
            files = [
                {k: v for k, v in n.items() if k != "content"} | (
                    {"size": str(len(n["content"]))} if n["mimeType"] != FOLDER_MIME else {})
                for n in self.d.nodes.values()
                if not n["trashed"] and parent in n["parents"]
            ]
            return {"files": files}
        return _Req(go)

    # ---- get ----
    def get(self, fileId=None, fields=None, supportsAllDrives=False, **kw):
        self.d.calls.append("get")

        def go():
            n = self.d.nodes[fileId]
            out = {"id": n["id"], "name": n["name"], "mimeType": n["mimeType"]}
            if "driveId" in (fields or "") and not self.d.quota_bug:
                out["driveId"] = self.d.drive_id
            if "capabilities" in (fields or ""):
                out["capabilities"] = {"canAddChildren": True}
            return out
        return _Req(go)

    def get_media(self, fileId=None, supportsAllDrives=False, **kw):
        self.d.calls.append("get_media")
        return _Req(lambda: self.d.nodes[fileId]["content"])

    def export_media(self, fileId=None, mimeType=None, **kw):
        self.d.calls.append("export_media")

        def go():
            raw = self.d.nodes[fileId]["content"]
            if mimeType and "spreadsheetml" in mimeType:
                # Native Sheets are stored here as CSV; a real export returns
                # xlsx bytes, so produce them. Without this the "client dropped
                # a Google Sheet in" path would never actually be tested.
                import io
                import pandas as pd
                buf = io.BytesIO()
                pd.read_csv(io.BytesIO(raw)).to_excel(buf, index=False)
                return buf.getvalue()
            return raw
        return _Req(go)

    # ---- write ----
    def create(self, body=None, media_body=None, fields=None,
               supportsAllDrives=False, **kw):
        self.d.calls.append("create")
        assert supportsAllDrives, "create without supportsAllDrives fails on a shared drive"

        def go():
            if self.d.quota_bug and media_body is not None:
                raise RuntimeError(
                    "<HttpError 403 ... returned \"Service Accounts do not have "
                    "storage quota\". Details: storageQuotaExceeded>")
            content = media_body.path.read_bytes() if media_body else b""
            fid = self.d._add(body["name"],
                              body.get("mimeType") or _mime_of(media_body),
                              body["parents"][0], content)
            n = self.d.nodes[fid]
            return {"id": n["id"], "name": n["name"],
                    "mimeType": n["mimeType"], "modifiedTime": n["modifiedTime"]}
        return _Req(go)

    def update(self, fileId=None, body=None, media_body=None, fields=None,
               supportsAllDrives=False, **kw):
        self.d.calls.append("update")
        assert supportsAllDrives, "update without supportsAllDrives fails on a shared drive"

        def go():
            n = self.d.nodes[fileId]
            if body:
                n.update({k: v for k, v in body.items()})
            if media_body is not None:
                n["content"] = media_body.path.read_bytes()
                n["md5Checksum"] = hashlib.md5(n["content"]).hexdigest()
            n["modifiedTime"] = _now()
            return {"id": n["id"], "name": n["name"],
                    "modifiedTime": n["modifiedTime"]}
        return _Req(go)


def _mime_of(media) -> str:
    return getattr(media, "mimetype", None) or "application/octet-stream"


def install(drive_sync_module, drive: FakeDrive):
    """Point drive_sync at the fake: no credentials, no network."""
    drive_sync_module.media_upload = lambda path, mimetype=None: FakeMedia(path, mimetype)
    drive_sync_module.build_service = lambda *a, **k: drive
    return drive
