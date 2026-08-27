"""
DOMIN8 — Omnichannel Reporting
Streamlit front end over the existing pipeline.

    pip install streamlit plotly pandas openpyxl
    streamlit run app.py

Nothing in reconcile.py / stock_vs_sales.py is modified. This file uploads the
per-cycle files into a scratch workspace, calls the same functions the CLI calls,
and renders the results.
"""
from __future__ import annotations

import contextlib
import io
import shutil
import zipfile
import sys
import tempfile
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))
import pipeline_config as C
from reconcile import run as reconcile_run
from stock_vs_sales import (build as svs_build, action_list, period_label,
                            write_override_template)

warnings.filterwarnings("ignore")

BRAND = "DOMIN8"
ROOT = Path(__file__).parent

# Validated categorical palette (see the data-viz palette reference).
INK, MUTED = "#0b0b0b", "#898781"
SELL_IN, SELL_OUT = "#2a78d6", "#eb6834"

st.set_page_config(page_title=f"{BRAND} · Omnichannel Reports",
                   page_icon="📊", layout="wide",
                   initial_sidebar_state="expanded")

st.markdown("""
<style>
  .block-container {padding-top: 2.2rem; max-width: 1400px;}
  .brandbar {background:#0b0b0b; color:#fff; padding:18px 26px; border-radius:12px;
             margin-bottom:22px; display:flex; align-items:baseline; gap:16px;}
  .brandbar h1 {margin:0; font-size:26px; font-weight:700; letter-spacing:.14em;}
  .brandbar span {color:#c3c2b7; font-size:13.5px;}
  div[data-testid="stMetric"] {background:#fcfcfb; border:1px solid rgba(11,11,11,.10);
             border-radius:12px; padding:16px 18px;}
  div[data-testid="stMetricLabel"] p {font-size:12.5px; color:#52514e;}
  section[data-testid="stSidebar"] {background:#f9f9f7;}
  .stTabs [data-baseweb="tab-list"] {gap: 22px;}
</style>
""", unsafe_allow_html=True)

st.markdown(f'<div class="brandbar"><h1>{BRAND}</h1>'
            f'<span>Omnichannel reporting · sell-in, sell-out and inbound '
            f'reconciled to one master SKU</span></div>', unsafe_allow_html=True)


# ===========================================================================
# Pipeline plumbing
# ===========================================================================

def workspace() -> Path:
    """A scratch copy of the input tree, so an upload never touches the real one."""
    if "ws" not in st.session_state:
        ws = Path(tempfile.mkdtemp(prefix="domin8_"))
        (ws / "input" / "uniware").mkdir(parents=True)
        (ws / "input" / "amazon vc").mkdir(parents=True)
        (ws / "input" / "retail stores").mkdir(parents=True)
        (ws / "output").mkdir(parents=True)

        # Seed from a local tree IF one exists. On a hosted deploy there is no
        # such tree -- the repo carries no data -- so everything arrives by
        # upload and lives in this session's workspace only.
        if C.INPUT.exists():
            for src in C.INPUT.glob("*.xls*"):
                if not src.name.startswith(("~$", ".")):
                    shutil.copy2(src, ws / "input" / src.name)
            if C.REORDER_OVERRIDES.exists():
                shutil.copy2(C.REORDER_OVERRIDES,
                             ws / "input" / C.REORDER_OVERRIDES.name)
            for real, sub in ((C.UNIWARE_DIR, "uniware"),
                              (C.AMAZON_DIR, "amazon vc"),
                              (C.STORES_DIR, "retail stores")):
                if not real.exists():
                    continue
                for src in real.glob("*"):
                    if src.is_file() and not src.name.startswith(("~$", ".")):
                        shutil.copy2(src, ws / "input" / sub / src.name)
        st.session_state.ws = ws
    return st.session_state.ws


def stage(uploads, folder: Path) -> int:
    for f in uploads or []:
        (folder / f.name).write_bytes(f.getbuffer())
    return len(uploads or [])


def has_master(inp: Path) -> Path | None:
    """The master mapping table, by content -- filename is not load-bearing."""
    for f in inp.glob("*.xls*"):
        if f.name.startswith(("~$", ".")):
            continue
        try:
            probe = pd.read_excel(f, header=None, nrows=4)
        except Exception:
            continue
        flat = " ".join(str(v).lower() for v in probe.values.ravel())
        if "sku code" in flat:
            return f
    return None


ZIP_ROUTES = {"uniware": "uniware", "amazon": "amazon vc",
              "retail": "retail stores", "store": "retail stores"}


def absorb_zip(upload, inp: Path) -> tuple[int, list[str]]:
    """Take one zip of the whole input set and file each member by its path.

    Twelve uploads a cycle is a chore; one drag-and-drop is not. Members are
    routed by the folder name inside the archive, and anything with a Sku Code
    column lands at the top level as the master table.
    """
    placed, skipped = 0, []
    with zipfile.ZipFile(io.BytesIO(upload.getbuffer())) as z:
        for m in z.namelist():
            name = Path(m).name
            if m.endswith("/") or name.startswith(("~$", ".", "__MACOSX")):
                continue
            if Path(name).suffix.lower() not in (".csv", ".xlsx", ".xls", ".xlsm"):
                skipped.append(name)
                continue
            low = m.lower()
            dest = next((v for k, v in ZIP_ROUTES.items() if k in low), None)
            target = (inp / dest / name) if dest else (inp / name)
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(z.read(m))
            placed += 1
    return placed, skipped


def data_files(folder: Path):
    return sorted(p for p in folder.glob("*")
                  if p.is_file() and p.suffix.lower() in (".csv", ".xlsx", ".xls", ".xlsm")
                  and not p.name.startswith(("~$", ".")))


@st.cache_data(show_spinner=False)
def sheet(path_str: str, name: str) -> pd.DataFrame:
    return pd.read_excel(path_str, sheet_name=name)


def execute(ws: Path, asof, period_days, recent_days):
    """Run the real pipeline and collect everything the UI needs."""
    inp, out = ws / "input", ws / "output"
    log = io.StringIO()
    with contextlib.redirect_stdout(log):
        reconcile_run(inp, out, facility=C.FACILITY)

    xlsx = out / "Omnichannel_Report.xlsx"
    res = {
        "log": log.getvalue(),
        "summary":  sheet(str(xlsx), "00 Summary"),
        "channel":  sheet(str(xlsx), "01 Channel"),
        "stores":   sheet(str(xlsx), "02 Stores"),
        "inventory": sheet(str(xlsx), "04 Inventory"),
        "exceptions": sheet(str(xlsx), "06 Exceptions"),
        "audit":    sheet(str(xlsx), "07 Load Audit"),
        "coverage": sheet(str(xlsx), "08 Period Coverage"),
        "xlsx": xlsx.read_bytes(),
        "dashboard": (out / "dashboard.html").read_bytes(),
    }

    svs_log = io.StringIO()
    with contextlib.redirect_stdout(svs_log):
        sku, art, art_raw, meta = svs_build(inp, out, pd.Timestamp(asof),
                                            period_days, recent_days)
        acts = action_list(art_raw)
    res.update(sku=sku, article=art, actions=acts, meta=meta,
               plabel=period_label(meta["per_start"], meta["asof"]))

    from stock_vs_sales import write_workbook, safe_excel_path
    svs_path = safe_excel_path(out / f"Stock_vs_Sales_{pd.Timestamp(asof):%d%m%y}.xlsx")
    write_workbook(svs_path, sku, art, meta, 0, art_raw)
    res["svs_xlsx"] = svs_path.read_bytes()
    res["svs_name"] = svs_path.name

    # Seed the merchandiser-notes template on first run, exactly as the CLI does,
    # so the editor has rows to show instead of an empty state.
    notes_path = inp / C.REORDER_OVERRIDES.name
    if not notes_path.exists():
        with contextlib.redirect_stdout(io.StringIO()):
            write_override_template(inp, sku)
    res["notes"] = (pd.read_csv(notes_path, dtype=str).fillna("")
                    if notes_path.exists() else pd.DataFrame())
    res["ran_at"] = datetime.now()
    return res


def kpi(df: pd.DataFrame, metric: str, default=0):
    if df is None or "metric" not in df.columns:
        return default
    hit = df.loc[df.metric == metric, "value"]
    return hit.iat[0] if len(hit) else default


def inr(v) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return "—"
    if v >= 1e7:
        return f"₹{v/1e7:.2f} Cr"
    if v >= 1e5:
        return f"₹{v/1e5:.2f} L"
    return f"₹{v:,.0f}"


def num(v) -> str:
    try:
        return f"{float(v):,.0f}"
    except (TypeError, ValueError):
        return "—"


# ===========================================================================
# Sidebar — inputs and run
# ===========================================================================

ws = workspace()

with st.sidebar:
    inp = ws / "input"

    st.subheader("Input files")
    with st.expander("Upload everything at once (.zip)"):
        z = st.file_uploader("Zip of the input folder", type=["zip"],
                             label_visibility="collapsed",
                             help="Folders named uniware / amazon / retail inside "
                                  "the zip are filed automatically.")
        if z is not None and st.button("Unpack", use_container_width=True):
            n, skipped = absorb_zip(z, inp)
            st.success(f"Filed {n} file(s)")
            if skipped:
                st.caption(f"ignored: {', '.join(skipped[:4])}")
            st.rerun()

    master = has_master(inp)
    m_up = st.file_uploader("Master mapping table", type=["xlsx", "xls"],
                            help="The sheet with Sku Code plus every platform id. "
                                 "Upload once per session.")
    if m_up is not None:
        (inp / m_up.name).write_bytes(m_up.getbuffer())
        master = has_master(inp)

    uni = st.file_uploader("Uniware reports", type=["csv"],
                           accept_multiple_files=True,
                           help="Tally GST, Returns, Inventory Snapshot, "
                                "Item Master, Purchase Orders")
    stage(uni, inp / "uniware")

    amazon = st.file_uploader("Amazon Vendor Central", type=["csv", "xlsx"],
                              accept_multiple_files=True,
                              help="Sales and Inventory reports, ASIN level")
    stores = st.file_uploader("Retail stores", type=["csv", "xlsx", "xls"],
                              accept_multiple_files=True,
                              help="Each store's sale and stock-on-hand files")

    st.divider()
    st.caption("READY TO BUILD?")
    checks = [
        ("Master mapping table", 1 if master else 0, True),
        ("Uniware reports", len(data_files(inp / "uniware")), False),
        ("Amazon Vendor Central", len(data_files(inp / "amazon vc")), False),
        ("Retail stores", len(data_files(inp / "retail stores")), False),
    ]
    for label, n, required in checks:
        if n:
            st.markdown(f"✅ &nbsp;{label} · **{n}**", unsafe_allow_html=True)
        elif required:
            st.markdown(f"⛔ &nbsp;{label} — **required**", unsafe_allow_html=True)
        else:
            st.markdown(f"⚪ &nbsp;{label} — not provided", unsafe_allow_html=True)

    st.button("Refresh from Uniware", disabled=True, use_container_width=True,
              help="Coming soon — Uniware will be pulled automatically each night.")

    st.divider()
    with st.expander("Report settings"):
        asof = st.date_input("Report date", value=datetime.now().date())
        period_days = st.number_input("Sales window (days)", 7, 365,
                                      C.PERIOD_DAYS, step=1)
        recent_days = st.number_input("Short window (days)", 3, 90,
                                      C.RECENT_DAYS, step=1)

    go = st.button("Build report", type="primary", use_container_width=True,
                   disabled=not master)
    if not master:
        st.caption("The master mapping table is needed before anything can run.")
    if st.session_state.get("res"):
        st.caption(f"Last run {st.session_state.res['ran_at']:%d %b %H:%M}")

if go:
    stage(amazon, ws / "input" / "amazon vc")
    stage(stores, ws / "input" / "retail stores")
    missing = [lbl for lbl, sub in (("Amazon", "amazon vc"),
                                    ("retail store", "retail stores"),
                                    ("Uniware", "uniware"))
               if not data_files(ws / "input" / sub)]
    if missing:
        st.warning("No " + ", ".join(missing) + " files — those channels will be "
                   "missing from the report rather than wrong.")
    with st.spinner("Reconciling every channel to the master SKU…"):
        st.session_state.res = execute(ws, asof, int(period_days), int(recent_days))
    st.rerun()

res = st.session_state.get("res")

if res is None:
    st.info("Upload this cycle's files in the sidebar — or drop one zip of the "
            "whole input folder — then **Build report**.")
    c1, c2, c3 = st.columns(3)
    c1.markdown("#### Sell-in\nWhat you invoiced **to** each channel.")
    c2.markdown("#### Sell-out\nWhat the **end customer** actually bought.")
    c3.markdown("#### Inbound\nWhat you ordered **from vendors**.")
    st.caption("Three separate flows. The report never adds them together.")
    st.stop()


# ===========================================================================
# Summary
# ===========================================================================

s = res["summary"]
so_u, so_v = kpi(s, "Sell-out units (end customers)"), kpi(s, "Sell-out value")
si_u, si_v = kpi(s, "Sell-in units (invoiced to channels)"), kpi(s, "Sell-in value")

a, b, c, d = st.columns(4)
a.metric("Sell-out units", num(so_u), help="Bought by end customers")
a.caption(inr(so_v))
b.metric("Sell-in units", num(si_u), help="Invoiced to channels")
b.caption(inr(si_v))
c.metric("Stock on hand", num(kpi(s, "Stock on hand (units)")))
c.caption(f"{num(kpi(s,'Units on order (open POs)'))} on order")
d.metric("Active SKUs", num(kpi(s, "SKUs with activity")))
d.caption(f"{num(kpi(s,'Channels'))} channels · {num(kpi(s,'Locations'))} locations")

match = kpi(s, "Identifier match rate %")
unmapped = int(float(kpi(s, "Unmapped identifiers") or 0))
if unmapped:
    st.warning(f"**{match}% of identifiers resolved.** {unmapped:,} did not and are "
               f"excluded from every figure above — see **Data quality**.")

tabs = st.tabs(["Action list", "Stock vs Sales", "Channels & stores",
                "Reorder notes", "Data quality", "Downloads"])

# ---------------------------------------------------------------- actions --
with tabs[0]:
    acts = res["actions"]
    st.caption("What to do this week. REORDER — selling and running out. "
               "SIZE BREAK — a moving style missing half its sizes. "
               "CLEAR — no sales, sitting on stock.")
    if not len(acts):
        st.success("Nothing flagged this cycle.")
    else:
        counts = acts["Action"].value_counts()
        cols = st.columns(len(counts))
        for col, (k, v) in zip(cols, counts.items()):
            col.metric(k.title(), int(v))
        pick = st.multiselect("Show", list(counts.index), list(counts.index))
        st.dataframe(acts[acts.Action.isin(pick)], use_container_width=True,
                     hide_index=True, height=520)

# --------------------------------------------------------- stock vs sales --
with tabs[1]:
    grain = st.radio("Grain", ["sku wise", "Article wise"],
                     horizontal=True, label_visibility="collapsed")
    df = res["sku"] if grain == "sku wise" else res["article"]

    f1, f2, f3 = st.columns(3)
    cats = sorted(x for x in df["Category"].dropna().unique())
    pick_cat = f1.multiselect("Category", cats)
    movs = sorted(x for x in df["Movement"].dropna().unique())
    pick_mov = f2.multiselect("Movement", movs)
    invs = sorted(x for x in df["Inventory status"].dropna().unique())
    pick_inv = f3.multiselect("Inventory status", invs)

    view = df
    if pick_cat:
        view = view[view["Category"].isin(pick_cat)]
    if pick_mov:
        view = view[view["Movement"].isin(pick_mov)]
    if pick_inv:
        view = view[view["Inventory status"].isin(pick_inv)]

    st.caption(f"{len(view):,} of {len(df):,} rows · "
               f"period {res['meta']['per_start']:%d %b} – "
               f"{res['meta']['asof']:%d %b}")
    st.dataframe(view, use_container_width=True, hide_index=True, height=520)

# ------------------------------------------------------ channels & stores --
with tabs[2]:
    ch = res["channel"].copy()
    wide = ch.pivot_table(index="channel", columns="flow",
                          values="net_value", aggfunc="sum").fillna(0)
    for col in ("sell_in", "sell_out"):
        if col not in wide:
            wide[col] = 0.0
    wide = wide.sort_values("sell_in", ascending=False)

    st.caption("A channel on both sides is expected, not double counting: "
               "sell-in is what you invoiced it, sell-out is what its customers bought.")
    try:
        import plotly.graph_objects as go
        fig = go.Figure()
        fig.add_bar(y=wide.index, x=wide["sell_in"], name="Sell-in",
                    orientation="h", marker_color=SELL_IN)
        fig.add_bar(y=wide.index, x=wide["sell_out"], name="Sell-out",
                    orientation="h", marker_color=SELL_OUT)
        fig.update_layout(barmode="group", height=44 * len(wide) + 120,
                          margin=dict(l=0, r=0, t=10, b=10),
                          plot_bgcolor="rgba(0,0,0,0)",
                          paper_bgcolor="rgba(0,0,0,0)",
                          legend=dict(orientation="h", y=1.08, x=0),
                          xaxis=dict(gridcolor="#e1e0d9", zeroline=False,
                                     title=None, tickprefix="₹"),
                          yaxis=dict(autorange="reversed", title=None))
        st.plotly_chart(fig, use_container_width=True)
    except ImportError:
        st.bar_chart(wide)

    st.markdown("##### Channel detail")
    st.dataframe(ch, use_container_width=True, hide_index=True)
    st.markdown("##### Retail stores")
    st.dataframe(res["stores"], use_container_width=True, hide_index=True)

# ------------------------------------------------------------ note editor --
with tabs[3]:
    st.caption("Merchandiser judgement — these two columns are not calculated. "
               "Edit here and they are joined into every future report.")
    notes = res["notes"]
    if not len(notes):
        st.info("No notes file yet. It is seeded on the first full run.")
    else:
        edited = st.data_editor(
            notes, use_container_width=True, hide_index=True, height=520,
            disabled=["Sku Code", "Item Name"],
            column_config={
                "Style Re-order Status": st.column_config.SelectboxColumn(
                    options=["", "Active", "Dropped", "Under review"]),
                "Reorder Status": st.column_config.TextColumn(width="large"),
            })
        if st.button("Save notes", type="primary"):
            edited.to_csv(ws / "input" / C.REORDER_OVERRIDES.name, index=False)
            st.success("Saved. Build the report again to apply them.")

# ---------------------------------------------------------- data quality ---
with tabs[4]:
    q1, q2 = st.columns(2)
    q1.metric("Identifier match rate", f"{match}%")
    q2.metric("Unmapped identifiers", f"{unmapped:,}")
    st.caption("Unmapped identifiers are excluded from every figure and listed "
               "here. Add them to the master mapping table to bring them in.")
    st.dataframe(res["exceptions"], use_container_width=True, hide_index=True,
                 height=280)
    st.markdown("##### What loaded")
    st.dataframe(res["audit"], use_container_width=True, hide_index=True)
    st.markdown("##### Period coverage")
    st.caption("Each source covers a different window at a different grain. "
               "Read this before comparing channels.")
    st.dataframe(res["coverage"], use_container_width=True, hide_index=True)
    with st.expander("Run log"):
        st.code(res["log"], language="text")

# ------------------------------------------------------------- downloads ---
with tabs[5]:
    stamp = f"{res['meta']['asof']:%d%m%y}"
    d1, d2, d3 = st.columns(3)
    d1.download_button("Stock vs Sales (.xlsx)", res["svs_xlsx"],
                       file_name=res["svs_name"], use_container_width=True)
    d2.download_button("Omnichannel report (.xlsx)", res["xlsx"],
                       file_name=f"Omnichannel_Report_{stamp}.xlsx",
                       use_container_width=True)
    d3.download_button("Dashboard (.html)", res["dashboard"],
                       file_name=f"dashboard_{stamp}.html",
                       use_container_width=True)
    st.caption("The Stock vs Sales workbook keeps the exact column layout of the "
               "sheet the team already uses.")
