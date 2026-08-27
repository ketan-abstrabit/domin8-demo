#!/usr/bin/env python3
"""
Automates the four Uniware reports currently pulled by hand from
Dashboard > Other Reports.

  Tally GST Report
  Tally Return GST Report
  Purchase Orders          (dropdown label: "Purchase Orders Export")
  Inventory Snapshot

Flow per report:  auth -> create export job -> poll status -> download CSV

Usage:
    export UNIWARE_USER='domin8@abstrabit.com' or $env:UNIWARE_USER = 'domin8@abstrabit.com'
    export UNIWARE_PASS='...' or $env:UNIWARE_PASS = '...'
    python uniware_exports.py                       # last 30 days
    python uniware_exports.py --days 90
    python uniware_exports.py --only "Inventory Snapshot"
    python uniware_exports.py --start 2026-04-01 --end 2026-06-30
"""

import argparse
import json
import os
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

import requests

TENANT = "domin8"
BASE = f"https://{TENANT}.unicommerce.com"
CLIENT_ID = "my-trusted-client"          # fixed literal, not a credential
FACILITY = os.environ.get("UNIWARE_FACILITY", "DOMIN8")
OUTDIR = Path(os.environ.get("UNIWARE_OUTDIR", "./exports"))

POLL_INTERVAL = 10                        # seconds between status checks
POLL_TIMEOUT = 900                        # give up after 15 min per job


# ---------------------------------------------------------------------------
# Report definitions.
#
# `job_type` MUST be the config's `name`, NOT the dropdown's displayName.
# The only difference in this set is Purchase Orders, but the trap is real:
#   name = "Purchase Orders"   displayName = "Purchase Orders Export"
#
# Column ids below were harvested from:
#   GET /data/tasks/export/config/get?exportConfigName=<name>
# ...which is the UI's own endpoint. Re-run it any time the schema changes;
# see discover_columns() at the bottom of this file.
#
# `date_filter` differs per report - Inventory Snapshot uses "updatedIn",
# the others use "addedOn". Passing the wrong id is silently ignored, which
# means you get a full-table dump instead of a date-bounded one. Check row
# counts on first run.
# ---------------------------------------------------------------------------

REPORTS = {
    "Tally GST Report": {
        "job_type": "Tally GST Report",
        "date_filter": "addedOn",
        "facility_level": True,
        "has_pii": True,
        "columns": [
            "invoiceDate", "saleOrderCode", "invoiceCode", "channelName",
            "channelLedgerName", "productCode", "productSKU", "QTY",
            "unitPrice", "Currency", "currencyConversionRate", "total",
            "customerName", "shippingAddressName",
            "shippingAddressLine1", "shippingAddressLine2",
            "shippingAddressCity", "shippingAddressState",
            "shippingAddressCountry", "shippingAddressPincode",
            "shippingAddressPhone", "shippingProvider", "trackingNumber",
            "sales", "salesLedger",
            "cgst", "cgstRate", "sgst", "sgstRate", "igst", "igstRate",
            "utgst", "utgstRate", "cess", "cessRate",
            "OtherCharges", "OtherChargesLedger",
            "OtherCharges1", "OtherChargesLedger1",
            "Servicetax", "ServicetaxLedger",
            "discountLedger", "discountAmount", "imei", "godDown",
            "dispatchdate", "narration", "entity", "voucherTypeName", "tin",
            "original", "original1", "channelInvoiceDate", "channelState",
            "customerGSTIN", "channelPartyGSTIN", "billingPartyCode",
            "taxVerification", "gstregistrationtype", "TCSAmount",
            "ajustmentInSellingPrice", "ajustmentInDiscount",
            "OtherChargesLedger2", "OtherCharges2",
            "storeCredit", "prepaidAmount",
            "taxOnOtherCharges", "taxOnOtherCharges1", "taxOnOtherCharges2",
            "irn", "acknowledgementNumber", "hsnCode", "paymentMethod",
            "billingAddressLine1", "billingAddressLine2",
            "isbn", "bundleSkuCode",
        ],
        # All 77 exportable columns. hiddenColumns (igstA, cgstA, sgstA,
        # utgstA, salesLedgerA/B, invItemId, taxLedgerA/B) are deliberately
        # absent - they are exportable:false server-side and feed the
        # calculated columns. Requesting them will be rejected.
    },

    "Tally Return GST Report": {
        "job_type": "Tally Return GST Report",
        # NOT addedOn. This config's filter id is "dateRange", and it is
        # "required": true (the UI labels it "Returned in Date Range*").
        # A create without it will be rejected.
        "date_filter": "dateRange",
        "date_required": True,
        "facility_level": True,
        "has_pii": True,
        "columns": [
            "invoiceDate", "saleOrderCode", "invoiceCode", "channelName",
            "channelLedgerName", "productCode", "productSKU", "QTY",
            "unitPrice", "Currency", "currencyConversionRate", "total",
            "customerName", "shippingAddressName",
            "shippingAddressLine1", "shippingAddressLine2",
            "shippingAddressCity", "shippingAddressState",
            "shippingAddressCountry", "shippingAddressPincode",
            "shippingAddressPhone", "shippingProvider", "trackingNumber",
            "sales", "salesLedger",
            "cgst", "cgstRate", "sgst", "sgstRate", "igst", "igstRate",
            "utgst", "utgstRate", "cess", "cessRate",
            "OtherCharges", "OtherChargesLedger",
            "OtherCharges1", "OtherChargesLedger1",
            "Servicetax", "ServicetaxLedger",
            "discountLedger", "discountAmount", "imei", "godDown",
            "dispatchdate", "narration", "entity", "voucherTypeName", "tin",
            "original", "original1", "channelInvoiceDate", "channelState",
            "channelPartyGSTIN", "customerGSTIN", "billingPartyCode",
            "taxVerification", "gstregistrationtype", "rpcode",
            "irn", "acknowledgementNumber", "productHSNCode", "returnType",
        ],
        # All 64 exportable columns. Note this config has "returnType" and
        # "rpcode" which the forward Tally GST Report does not, and lacks
        # TCSAmount / storeCredit / paymentMethod / billingAddress* which it
        # does. The two reports are NOT the same shape - do not share a schema.
    },

    "Purchase Orders": {
        "job_type": "Purchase Orders",
        "date_filter": "addedOn",
        "facility_level": True,
        "has_pii": False,
        "columns": [
            "purchaseOrderCode", "created", "type", "purchaseOrderCreatedBy",
            "approvedDate", "deliveryDate", "itemTypeName", "itemtypeSku",
            "category", "hsnCode", "gstTaxTypeCode", "vendor", "vendorCode",
            "vendorSku", "quantity", "recieveQuantity", "rejectedQuantity",
            "pendingQuantity", "ageingDays", "percentagePending",
            "percentageRejection", "facility", "unitPrice", "subTotal",
            "cgst", "igst", "sgst", "utgst", "cess",
            "cgstrate", "igstrate", "sgstrate", "utgstrate", "cessrate",
            "total", "purchaseOrderStatus", "rejectionReason", "updated",
            "POtolerance", "ToleranceQtyInwarded", "RejectedQtyInwarded",
            "purchaseOrder_Remarks",
        ],
        # All 42 exportable columns.
    },

    "Inventory Snapshot": {
        "job_type": "Inventory Snapshot",
        "date_filter": "updatedIn",     # NOT addedOn
        "facility_level": True,
        "has_pii": False,
        # POINT-IN-TIME REPORT - do NOT date-filter it by default.
        #
        # "updatedIn" filters on when the inventory record was last touched,
        # not on a reporting period. Applying a 30-day window silently drops
        # every SKU that has not moved in 30 days - i.e. exactly the
        # slow-movers and dead stock you most want to see. The result looks
        # like a successful snapshot but is a partial one.
        #
        # Leave this True for a true current snapshot. Set --force-date-filter
        # only if you specifically want "SKUs that changed recently".
        "point_in_time": True,
        "columns": [
            "facility", "itemTypeName", "itemtypeSku", "ean", "upc", "isbn",
            "color", "size", "brand", "categoryName", "MRP",
            "openSale", "inventory", "quantityNotFound", "excessQuantity",
            "quarantinedInventory", "inventoryNotSynced", "inventoryBlocked",
            "badInventory", "putawayPending", "pendingInventoryAssessment",
            "pendingStockTransfer", "openPurchase",
            "enabled", "updated", "costPrice",
            # 22 tenant-specific custom fields on ItemType. These are the
            # bulk of what the dashboard export shows and the earlier version
            # of this script was missing.
            "itemType_AdditionalProductDetails", "itemType_Closure",
            "itemType_Collection", "itemType_ColourFamily",
            "itemType_DisplayName", "itemType_FabricComposition",
            "itemType_FabricType", "itemType_Fit", "itemType_Gender",
            "itemType_ImageFolderLink", "itemType_KeyFeatures",
            "itemType_KeyWords", "itemType_Neck", "itemType_Occasion",
            "itemType_Pattern", "itemType_Pocket", "itemType_Print",
            "itemType_ProductTitle", "itemType_Sleeve",
            "itemType_SubCategory", "itemType_VideoLink", "itemType_Year",
        ],
        # All 48 exportable columns.
        # Note: this export exposes buckets the REST snapshot endpoint does
        # NOT return - quantityNotFound, excessQuantity, quarantinedInventory,
        # inventoryNotSynced, badInventory, and costPrice. For reconciliation
        # this export is strictly richer than /inventory/inventorySnapshot/get.
    },

    "Item Master": {
        "job_type": "Item Master",
        # Filter id is "dateRange" = "Added in Date Range". There is also
        # "updatedOn" = "Updated in Date Range", plus a deprecated
        # "updatedSince" (hours). None are required.
        "date_filter": "dateRange",
        # Catalogue-level, not transactional. A date window would return only
        # items ADDED in that window, dropping the entire legacy catalogue -
        # same trap as Inventory Snapshot. Default to no filter for a complete
        # master; use --force-date-filter to override.
        "point_in_time": True,
        "facility_level": True,   # header is harmless; this config has no
                                  # facility column and no facility filter
        "has_pii": False,         # hasPiiData:false - nothing gets masked
        "columns": [
            "categoryCode", "skuCode", "itemName", "description",
            "scanIdentifier", "requireCustomization",
            "length", "width", "height", "weight",
            "ean", "upc", "isbn", "color", "size", "brand",
            "itemDetailFields", "tags", "imageUrl", "productPageUrl",
            "taxTypeCode", "gstTaxTypeCode",
            "basePrice", "costPrice", "tat", "MRP",
            "updated", "category", "enabled", "type",
            # Bundle/kit composition lives here: `type` is SIMPLE or BUNDLE
            # and these three describe the components. Needed before any
            # sell-through or replenishment maths, or bundles double-count.
            "componentProductCode", "componentQuantity", "componentPrice",
            "hsn", "taxCalculationType",
            "batchGroupCode", "grnExpiryTolerance",
            "dispatchExpiryTolerance", "returnExpiryTolerance",
            "expirable", "determineExpiryFrom", "shelfLife", "ExpiryDate",
            "skuType", "fragile", "dangerousGood",
            # same 22 ItemType custom fields as Inventory Snapshot
            "itemType_AdditionalProductDetails", "itemType_Closure",
            "itemType_Collection", "itemType_ColourFamily",
            "itemType_DisplayName", "itemType_FabricComposition",
            "itemType_FabricType", "itemType_Fit", "itemType_Gender",
            "itemType_ImageFolderLink", "itemType_KeyFeatures",
            "itemType_KeyWords", "itemType_Neck", "itemType_Occasion",
            "itemType_Pattern", "itemType_Pocket", "itemType_Print",
            "itemType_ProductTitle", "itemType_Sleeve",
            "itemType_SubCategory", "itemType_VideoLink", "itemType_Year",
        ],
        # All 68 exportable columns (46 standard + 22 custom).
        # hiddenColumns is empty for this config.
        #
        # Other optional filters, if you ever need them:
        #   name          TEXT        item name contains
        #   enabledFilter BOOLEAN     enabled / disabled
        #   updatedOn     DATERANGE   updated in range
        #   categoryIds   MULTISELECT 47 categories on this tenant
        #   skuType       MULTISELECT GOODS | SERVICE
    },
}


# ---------------------------------------------------------------------------


# DO NOT hardcode facility codes from the UI dropdown - that dropdown shows
# DISPLAY NAMES, not codes. On the domin8 tenant they differ:
#
#   dropdown "AMAZON_FBA_MAA4TN"  -> real code "AMAZON_FBA_IN"
#   dropdown "AMAZON_FBA_IN_CJB1" -> real code "AMAZON_FBA_IN_NEW"
#   dropdown "DOMIN8"             -> real code "domin8"
#
# (MAA4/CJB1 are Amazon's own warehouse identifiers.) Using a display name as
# the Facility header returns HTTP 403, which looks exactly like a permissions
# problem and is not. Always resolve codes via facility/search.
#
# Same trap as exportJobTypeName: use `name`/`facilityCode`, never displayName.
FACILITY_SEARCH_BODY = {
    "facilityStatus": "ALL",
    "fromDate": "2010-01-01T00:00:00.000Z",
    "toDate": "2035-12-31T00:00:00.000Z",
    "dateType": "CREATED",
}


def fetch_facility_codes(session, token, enabled_only=True):
    """Resolve real facility codes from the API."""
    body = uniware_post(
        session, token, "/services/rest/v1/facility/search",
        FACILITY_SEARCH_BODY,
    )
    codes = []
    for p in body.get("parties") or []:
        if enabled_only and p.get("facilityStatus") != "ENABLED":
            continue
        codes.append(p["facilityCode"])
    return codes


def inspect_csv(path, sample_col="facility", max_distinct=25):
    """Report what actually came back, so scoping assumptions get verified
    against the file rather than assumed.

    Returns (row_count, col_count, {col_value: count} or None).
    """
    import csv as _csv

    with open(path, "r", newline="", encoding="utf-8", errors="replace") as fh:
        reader = _csv.reader(fh)
        try:
            header = next(reader)
        except StopIteration:
            return 0, 0, None

        idx = None
        for i, h in enumerate(header):
            if h.strip().lower() == sample_col.lower():
                idx = i
                break

        rows = 0
        counts = {}
        for row in reader:
            rows += 1
            if idx is not None and idx < len(row):
                val = (row[idx] or "").strip() or "(blank)"
                if len(counts) < max_distinct or val in counts:
                    counts[val] = counts.get(val, 0) + 1

    return rows, len(header), (counts if idx is not None else None)


def combine_csvs(paths, dest):
    """Concatenate same-schema CSVs into one file, keeping a single header.

    Verifies headers match before merging - a silent mismatch would misalign
    columns and produce a file that looks fine and is wrong.

    Also guards against the opposite failure: if an export ignores the
    Facility header, every per-facility file is the SAME dump and merging
    them multiplies rows with no error. We detect that by hashing each file's
    body and refusing to merge identical ones.
    """
    import hashlib

    header = None
    rows_written = 0
    seen_hashes = {}
    dest.parent.mkdir(parents=True, exist_ok=True)

    bodies = []
    for p in paths:
        with open(p, "r", newline="", encoding="utf-8",
                  errors="replace") as fh:
            first = fh.readline()
            body = fh.read()
        if header is None:
            header = first
        elif first.strip() != header.strip():
            raise RuntimeError(
                f"header mismatch in {p.name}; refusing to merge")

        digest = hashlib.sha256(body.encode("utf-8", "replace")).hexdigest()
        if digest in seen_hashes:
            raise RuntimeError(
                f"{p.name} is byte-identical to {seen_hashes[digest]} - the "
                f"export is not facility-scoped, so merging would duplicate "
                f"rows. Use a single call instead of --all-facilities.")
        seen_hashes[digest] = p.name
        bodies.append(body)

    with open(dest, "w", newline="", encoding="utf-8") as out:
        out.write(header)
        for body in bodies:
            out.write(body)
            rows_written += sum(1 for ln in body.splitlines() if ln.strip())

    return rows_written


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def get_token(session, username, password):
    """Query-param auth variant. The header variant (Authentication2) returns
    'invalid email:NONE_PROVIDED' on this tenant."""
    url = (
        f"{BASE}/oauth/token"
        f"?grant_type=password&client_id={CLIENT_ID}"
        f"&username={quote(username)}&password={quote(password, safe='')}"
    )
    r = session.get(url, timeout=60)
    r.raise_for_status()
    data = r.json()
    if "access_token" not in data:
        raise RuntimeError(f"No access_token in response: {data}")
    log(f"authenticated, token valid ~{data.get('expires_in', '?')}s")
    return data["access_token"]


def uniware_post(session, token, path, payload, facility=None,
                 raise_on_error=True):
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"bearer {token}",
    }
    if facility:
        headers["Facility"] = facility

    body_json = json.dumps(payload)
    if os.environ.get("UNIWARE_DEBUG"):
        log(f"  --> POST {path}")
        log(f"  --> {body_json[:2000]}")

    r = session.post(f"{BASE}{path}", headers=headers,
                     data=body_json, timeout=120)

    # A bare 404 means the path is wrong. Anything else non-2xx: Uniware
    # usually explains itself in the response body, so ALWAYS print that
    # before raising. errors[0].message names the exact Java class and field.
    if r.status_code == 404:
        raise RuntimeError(f"404 - check the endpoint path: {path}")

    if not r.ok:
        # 403 = the API user has no access to this facility. Worth surfacing
        # distinctly: it is a permissions gap, not a payload problem.
        if r.status_code == 403:
            raise PermissionError(
                f"403 Forbidden"
                + (f" for facility '{facility}'" if facility else "")
                + " - the API user lacks access to this facility/resource")
        log(f"  HTTP {r.status_code} from {path}")
        try:
            err_body = r.json()
            log(f"  response: {json.dumps(err_body, indent=2)[:3000]}")
            for err in err_body.get("errors") or []:
                log(f"  >> field={err.get('fieldName')} "
                    f"code={err.get('code')}")
                log(f"  >> {err.get('message')}")
            if not raise_on_error:
                return err_body
        except Exception:
            log(f"  raw response: {r.text[:3000]}")
        raise RuntimeError(f"{path} returned HTTP {r.status_code}")

    body = r.json()
    if not body.get("successful", False):
        # errors[0].description is the human message ("frequency can not be
        # empty"). errors[0].message is the machine code
        # ("MISSING_REQUIRED_PARAMETERS"). Print both - they are swapped
        # relative to what you'd expect.
        for err in body.get("errors") or []:
            log(f"  ERROR code={err.get('code')} "
                f"desc={err.get('description')}")
            log(f"  detail: {err.get('message')}")
        if not raise_on_error:
            return body
        raise RuntimeError(f"{path} returned successful:false")
    return body


# `frequency` is MANDATORY on this build even for a one-off export, despite
# the docs presenting it as optional next to cronExpression.
#
# CONFIRMED valid values (from error 200204 on the domin8 tenant):
#   ONETIME    - run once, now
#   RECURRENT  - scheduled; pair with cronExpression
#
# Note ONETIME, not ONE_TIME.
FREQUENCY_CANDIDATES = ["ONETIME", "RECURRENT"]

# The exportFilters serialisation is NOT documented and my first guess
# ({"name": ..., "dateRange": {...}}) was wrong. Uniware's deserializer gives
# very specific errors, so we try candidate shapes in order and keep the one
# that is accepted. The winning shape is logged and cached for the run.
#
# Note the configs identify filters by "id" (e.g. addedOn, dateRange,
# updatedIn), so id-keyed shapes are tried before name-keyed ones.
def _iso(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%S.000Z")


def _ymd(dt):
    return dt.strftime("%Y-%m-%d")


FILTER_SHAPES = [
    # CONFIRMED via --probe-filters against com.unifier.core.api.export
    # .WsExportFilter on the domin8 tenant:
    #
    #   id        IS a field (omitting it -> "exportFilters[0].id can not be
    #             empty"; supplying it -> "DateRange is mandatory for
    #             daterange value type", i.e. it advanced past parsing)
    #   dateRange IS a field (parsed cleanly on its own)
    #
    #   NOT fields: name, filterId, code, column, value
    #
    # So shape 1 below is the right one. The rest are kept only as fallbacks
    # in case a differently-typed filter (TEXT, BOOLEAN, MULTISELECT) needs
    # another key alongside id - date filters do not.
    ("id+dateRange", lambda f, s, e: [
        {"id": f, "dateRange": {"start": _iso(s), "end": _iso(e)}}]),
    ("id+start/end", lambda f, s, e: [
        {"id": f, "start": _iso(s), "end": _iso(e)}]),
    ("id+from/to", lambda f, s, e: [
        {"id": f, "fromDate": _iso(s), "toDate": _iso(e)}]),
    ("id+textRange", lambda f, s, e: [
        {"id": f, "dateRange": {"textRange": "LAST_30_DAYS"}}]),
]

# Diagnostic payloads for --probe-filters. The trick: an EMPTY filter object
# deserializes cleanly (no unknown fields to reject), so the server moves on
# to validation and tells us which fields it actually wanted. Each subsequent
# probe adds one candidate field name; "Unrecognized field X" means X is wrong,
# anything else means X is real.
FILTER_PROBES = [
    ("empty object", [{}]),
    ("id only", [{"id": "PLACEHOLDER"}]),
    ("filterId only", [{"filterId": "PLACEHOLDER"}]),
    ("code only", [{"code": "PLACEHOLDER"}]),
    ("column only", [{"column": "PLACEHOLDER"}]),
    ("value only", [{"value": "PLACEHOLDER"}]),
    ("dateRange only", [{"dateRange": {"start": "2026-01-01T00:00:00.000Z",
                                       "end": "2026-01-31T00:00:00.000Z"}}]),
]


def probe_filters(session, token, spec, facility, filter_id):
    """Interrogate WsExportFilter to learn its real field names."""
    log("PROBING exportFilters structure "
        f"(report='{spec['job_type']}', filter id='{filter_id}')")
    log("-" * 68)

    for label, filters in FILTER_PROBES:
        payload = {
            "exportJobTypeName": spec["job_type"],
            "exportColums": spec["columns"][:3],   # keep it small
            "frequency": "ONETIME",
            "exportFilters": json.loads(
                json.dumps(filters).replace("PLACEHOLDER", filter_id)),
        }
        body = uniware_post(
            session, token, "/services/rest/v1/export/job/create",
            payload, facility=facility, raise_on_error=False,
        )
        ok = body.get("successful")
        msgs = [f"{e.get('description')} | {e.get('message')}"
                for e in (body.get("errors") or [])]
        verdict = "ACCEPTED" if ok else "rejected"
        log(f"  {label:18} -> {verdict}")
        for m in msgs:
            log(f"     {m[:300]}")
        if ok:
            log(f"  >>> '{label}' was accepted. Note the job it created.")
    log("-" * 68)
    log("Read the messages above: 'Unrecognized field X' means X is NOT a "
        "field on WsExportFilter. Any other complaint means X IS valid and "
        "something else is missing.")


_filter_cache = {}

_frequency_cache = {}


def _is_frequency_error(body):
    for err in (body.get("errors") or []):
        blob = f"{err.get('description', '')} {err.get('message', '')}".lower()
        if "frequency" in blob:
            return True
    return False


def create_job(session, token, spec, start, end, opts, facility_code=None):
    """Build and submit the create payload.

    Required fields on this build: exportJobTypeName, exportColums, frequency.

    Bisect flags for when something is rejected:
      --no-columns   omit exportColums   (will fail: it is required)
      --no-filters   omit exportFilters  (valid - unbounded export)
      --columns-as-string  send columns as CSV instead of an array
      --frequency X  force a specific frequency value
    """
    base_payload = {"exportJobTypeName": spec["job_type"]}

    if spec.get("columns") and not opts.no_columns:
        # exportColums MUST be a JSON array. Sending a comma-separated string
        # returns: "Can not deserialize instance of java.util.ArrayList out of
        # VALUE_STRING token ... CreateExportJobRequest[\"exportColums\"]".
        # The docs describe it as "column names" which reads like a string -
        # it isn't. (Note the API's own spelling: exportColums, one 'n'.)
        if opts.columns_as_string:
            base_payload["exportColums"] = ",".join(spec["columns"])
        else:
            base_payload["exportColums"] = spec["columns"]

    skip_dates = (
        opts.no_filters
        or (spec.get("point_in_time") and not opts.force_date_filter)
    )

    facility = None
    if spec.get("facility_level"):
        facility = facility_code or FACILITY

    # Which filter shapes to attempt. If one already worked this run, reuse it.
    if skip_dates:
        shapes = [("none", None)]
    elif opts.filter_shape:
        shapes = [(s, fn) for s, fn in FILTER_SHAPES
                  if s == opts.filter_shape] or [FILTER_SHAPES[0]]
    elif _filter_cache.get("shape"):
        want = _filter_cache["shape"]
        shapes = [(s, fn) for s, fn in FILTER_SHAPES if s == want]
    else:
        shapes = FILTER_SHAPES

    freq = opts.frequency or _frequency_cache.get("value") or "ONETIME"
    last_body = None

    for shape_name, builder in shapes:
        payload = dict(base_payload, frequency=freq)
        if builder is not None:
            payload["exportFilters"] = builder(
                spec["date_filter"], start, end)

        body = uniware_post(
            session, token, "/services/rest/v1/export/job/create",
            payload, facility=facility, raise_on_error=False,
        )
        last_body = body

        if body.get("successful"):
            if builder is not None and _filter_cache.get("shape") != shape_name:
                log(f"  exportFilters shape '{shape_name}' accepted")
                _filter_cache["shape"] = shape_name
            _frequency_cache["value"] = freq
            code = (body.get("jobCode") or body.get("exportJobCode")
                    or body.get("code"))
            if not code:
                raise RuntimeError(f"No job code in response: {body}")
            return code

        if _is_frequency_error(body):
            # frequency wrong rather than filters - swap and retry this shape
            alt = "RECURRENT" if freq == "ONETIME" else "ONETIME"
            log(f"  frequency '{freq}' rejected, retrying as '{alt}'")
            freq = alt
            payload = dict(base_payload, frequency=freq)
            if builder is not None:
                payload["exportFilters"] = builder(
                    spec["date_filter"], start, end)
            body = uniware_post(
                session, token, "/services/rest/v1/export/job/create",
                payload, facility=facility, raise_on_error=False,
            )
            last_body = body
            if body.get("successful"):
                _frequency_cache["value"] = freq
                return (body.get("jobCode") or body.get("exportJobCode")
                        or body.get("code"))

        if builder is None:
            break   # nothing left to vary
        log(f"  exportFilters shape '{shape_name}' rejected, trying next")

    raise RuntimeError(f"create failed; last response: {last_body}")


def wait_for_job(session, token, job_code):
    deadline = time.time() + POLL_TIMEOUT
    while time.time() < deadline:
        body = uniware_post(
            session, token, "/services/rest/v1/export/job/status",
            {"jobCode": job_code},
        )
        status = (body.get("status") or body.get("jobStatus") or "").upper()
        log(f"  status={status or '(none)'}")

        if status in ("SUCCESSFUL", "COMPLETE", "COMPLETED"):
            path = body.get("filePath") or body.get("url")
            if not path:
                raise RuntimeError(f"Job done but no filePath: {body}")
            return path
        if status in ("FAILED", "ERROR", "CANCELLED"):
            raise RuntimeError(f"Job {job_code} failed: {body}")

        time.sleep(POLL_INTERVAL)

    raise TimeoutError(f"Job {job_code} still running after {POLL_TIMEOUT}s")


def download(session, token, file_path, dest):
    url = file_path if file_path.startswith("http") else f"{BASE}{file_path}"
    r = session.get(url, headers={"Authorization": f"bearer {token}"},
                    timeout=300, stream=True)
    r.raise_for_status()
    dest.parent.mkdir(parents=True, exist_ok=True)
    with open(dest, "wb") as fh:
        for chunk in r.iter_content(65536):
            fh.write(chunk)
    return dest.stat().st_size


def discover_columns(session, name):
    """Re-harvest column and filter ids straight from the UI's own endpoint.

    Requires a logged-in browser session cookie, so this is a helper for
    manual use rather than part of the automated run. Easiest route: open
    the URL below in a browser where you're already logged into Uniware.
    """
    url = (f"{BASE}/data/tasks/export/config/get"
           f"?exportConfigName={quote(name)}")
    print(f"Open this in a logged-in browser:\n  {url}")
    print("Then read exportColumns[].id and exportFilters[].id")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--days", type=int, default=30)
    ap.add_argument("--start")
    ap.add_argument("--end")
    ap.add_argument("--only", action="append",
                    help="run only these reports (repeatable)")
    ap.add_argument("--no-columns", action="store_true",
                    help="omit exportColums (let Uniware use defaults)")
    ap.add_argument("--no-filters", action="store_true",
                    help="omit exportFilters (unbounded export)")
    ap.add_argument("--columns-as-string", action="store_true",
                    help="send exportColums as a CSV string (default: array)")
    ap.add_argument("--frequency",
                    help="force a frequency value instead of auto-detecting")
    ap.add_argument("--notify-email",
                    help="optional notificationEmail on the export job")
    ap.add_argument("--all-facilities", action="store_true",
                    help="run facility-level reports for all 5 facilities")
    ap.add_argument("--facility", action="append",
                    help="specific facility code (repeatable)")
    ap.add_argument("--force-date-filter", action="store_true",
                    help="apply the date window even to point-in-time reports")
    ap.add_argument("--combine", action="store_true",
                    help="also write one merged CSV per report across "
                         "facilities")
    ap.add_argument("--filter-shape",
                    choices=[s for s, _ in FILTER_SHAPES],
                    help="force one exportFilters shape instead of probing")
    ap.add_argument("--probe-filters", action="store_true",
                    help="diagnostic: interrogate WsExportFilter field names "
                         "and exit")
    args = ap.parse_args()

    if args.facility:
        opts_facilities = args.facility
    else:
        opts_facilities = None      # resolved after auth

    username = os.environ.get("UNIWARE_USER")
    password = os.environ.get("UNIWARE_PASS")
    if not username or not password:
        sys.exit("Set UNIWARE_USER and UNIWARE_PASS environment variables.")

    if args.start and args.end:
        start = datetime.strptime(args.start, "%Y-%m-%d").replace(
            tzinfo=timezone.utc)
        end = datetime.strptime(args.end, "%Y-%m-%d").replace(
            tzinfo=timezone.utc)
    else:
        end = datetime.now(timezone.utc)
        start = end - timedelta(days=args.days)

    span = (end - start).days
    if span > 90:
        log(f"WARNING: {span}-day window requested. The UI's date presets cap "
            f"at LAST_90_DAYS. Explicit start/end may or may not be honoured "
            f"beyond that - verify row counts, and chunk into 90-day slices "
            f"if the export comes back truncated.")

    targets = args.only or list(REPORTS.keys())
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    session = requests.Session()
    token = get_token(session, username, password)

    if opts_facilities is None:
        opts_facilities = fetch_facility_codes(session, token)
        log(f"resolved {len(opts_facilities)} enabled facility code(s): "
            f"{', '.join(opts_facilities)}")
    if not opts_facilities:
        sys.exit("No facility codes available for this user.")
    if not args.all_facilities:
        log(f"single-call mode: sending Facility='{opts_facilities[0]}' and "
            f"checking whether the output is scoped or complete")

    results, failures = [], []
    denied_facilities = set()

    if args.probe_filters:
        name = (args.only or ["Purchase Orders"])[0]
        spec = REPORTS[name]
        probe_filters(session, token, spec, FACILITY, spec["date_filter"])
        return 0

    # PHASE 1: submit every job. Uniware processes them server-side in
    # parallel, so submitting all up front and polling afterwards is much
    # faster than create-wait-download-repeat.
    pending = []
    for name in targets:
        spec = REPORTS.get(name)
        if not spec:
            log(f"SKIP unknown report: {name}")
            continue

        # DEFAULT: one call per report. The Facility header is mandatory on
        # facility-level endpoints, but the export configs for these four
        # reports expose NO facility filter - and Inventory Snapshot carries a
        # `facility` COLUMN. That strongly suggests the export returns every
        # facility the user can see, with the column identifying each row, and
        # the header serves only to satisfy the endpoint.
        #
        # inspect_csv() checks this against the downloaded file and prints the
        # verdict. Use --all-facilities to loop only if the check shows the
        # header really does scope the output.
        if spec.get("facility_level") and args.all_facilities:
            fac_list = opts_facilities
        elif spec.get("facility_level"):
            fac_list = [opts_facilities[0]]
        else:
            fac_list = [None]

        for fac in fac_list:
            label = name + (f" @ {fac}" if (fac and args.all_facilities)
                            else "")
            if fac and fac in denied_facilities:
                log(f"{label}: skipped (facility previously returned 403)")
                continue
            if spec.get("has_pii"):
                log(f"{label}: hasPiiData=true; all columns requested, "
                    f"server may mask per user PII role")
            if spec.get("point_in_time") and not args.force_date_filter:
                log(f"{label}: point-in-time, no date filter "
                    f"(full current snapshot)")
            try:
                code = create_job(session, token, spec, start, end, args, fac)
                log(f"{label}: submitted {code}")
                pending.append((label, name, fac, code))
            except PermissionError as exc:
                if fac:
                    denied_facilities.add(fac)
                log(f"{label}: {exc}")
                failures.append((label, str(exc)))
            except Exception as exc:
                log(f"{label}: SUBMIT FAILED: {exc}")
                failures.append((label, str(exc)))

    if not pending:
        log("nothing submitted successfully")
        return 1

    log(f"submitted {len(pending)} job(s); waiting for completion")

    # PHASE 2: poll and download each.
    for label, name, fac, code in pending:
        try:
            file_path = wait_for_job(session, token, code)
            safe = name.replace(" ", "_")
            if fac and args.all_facilities:
                safe = f"{safe}__{fac}"
            dest = OUTDIR / stamp / f"{safe}.csv"
            size = download(session, token, file_path, dest)

            rows, cols, fac_counts = inspect_csv(dest)
            log(f"{label}: saved {dest.name} "
                f"({rows:,} rows x {cols} cols, {size:,} bytes)")

            if fac_counts is not None:
                distinct = sorted(fac_counts)
                if len(distinct) == 1:
                    log(f"  facility column: 1 value ({distinct[0]}) "
                        f"-> Facility header DOES scope this export")
                    if fac and distinct[0].lower() != str(fac).lower():
                        log(f"  note: requested code '{fac}', column shows "
                            f"'{distinct[0]}' (display name vs code)")
                else:
                    log(f"  facility column: {len(distinct)} values "
                        f"-> Facility header does NOT scope this export")
                    for v in distinct:
                        log(f"    {v}: {fac_counts[v]:,} rows")
            else:
                # No facility column: we CANNOT tell from the file whether
                # this export was scoped to one facility or covered all of
                # them. Matters most for the Tally GST reports - if they are
                # scoped and you only pull one facility, the GST return is
                # missing every marketplace-fulfilled invoice.
                log(f"  no facility column: coverage NOT verifiable from the "
                    f"file. Compare row counts across --facility runs before "
                    f"trusting this for filing.")

            results.append((label, dest, size))
        except Exception as exc:
            log(f"{label}: FAILED: {exc}")
            failures.append((label, str(exc)))

    # PHASE 3: optionally merge per-facility files into one CSV per report.
    if args.combine:
        by_report = {}
        for label, dest, _size in results:
            base = label.split(" @ ")[0]
            by_report.setdefault(base, []).append(dest)

        for base, paths in by_report.items():
            if len(paths) < 2:
                continue
            safe = base.replace(" ", "_")
            merged = OUTDIR / stamp / f"{safe}__COMBINED.csv"
            try:
                n = combine_csvs(sorted(paths), merged)
                size = merged.stat().st_size
                log(f"{base}: combined {len(paths)} files -> {merged.name} "
                    f"({n:,} rows, {size:,} bytes)")
                results.append((f"{base} [COMBINED]", merged, size))
            except Exception as exc:
                log(f"{base}: combine failed: {exc}")

    print()
    log(f"done: {len(results)} succeeded, {len(failures)} failed")
    for name, dest, size in results:
        print(f"  OK    {name:28} {size:>12,} bytes  {dest}")
    for name, err in failures:
        print(f"  FAIL  {name:28} {err}")

    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())