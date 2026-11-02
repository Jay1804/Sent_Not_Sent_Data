"""Streamlit app: Report Sent tracker (month-to-date).

Credentials come from .env (DB_HOST, DB_NAME, DB_USER, DB_PASSWORD, optional
DB_PORT). Window: first of the month 00:00:00 up to today 07:59:59 by default.
"""

import io
import os
import threading
import time as time_mod
import zipfile
from pathlib import Path
from datetime import date, datetime, time

import pandas as pd
import streamlit as st
from streamlit.runtime.scriptrunner import get_script_run_ctx
from dotenv import load_dotenv
from sqlalchemy import create_engine, event
from sqlalchemy.engine import URL
from sqlalchemy.exc import SQLAlchemyError

from sent_report import build_xlsb_fast
from sent_tracker_query import fetch_range_cached

load_dotenv()

st.set_page_config(page_title="Report Sent Tracker", page_icon="\U0001F4CB", layout="wide")

CUTOFF = time(7, 59, 59)


@st.cache_resource
def get_engine(workers: int = 6):
    missing = [k for k in ("DB_HOST", "DB_NAME", "DB_USER", "DB_PASSWORD") if not os.environ.get(k)]
    if missing:
        raise RuntimeError(f"Missing in .env: {', '.join(missing)}")
    url = URL.create(
        "mysql+pymysql",
        username=os.environ["DB_USER"],
        password=os.environ["DB_PASSWORD"],
        host=os.environ["DB_HOST"],
        port=int(os.environ.get("DB_PORT", 3306)),
        database=os.environ["DB_NAME"],
    )
    engine = create_engine(url, pool_pre_ping=True, pool_recycle=1800,
                           pool_size=workers, max_overflow=0,
                           connect_args={"read_timeout": 1800, "connect_timeout": 30})

    @event.listens_for(engine, "connect")
    def _session_settings(dbapi_conn, _):
        with dbapi_conn.cursor() as cur:
            # default 1024 silently truncates the Check Names list
            cur.execute("SET SESSION group_concat_max_len = 1000000")

    return engine


def load_data(start: datetime, end: datetime, workers: int, bar, use_cache: bool) -> pd.DataFrame:
    return fetch_range_cached(get_engine(workers), start, end, workers,
                              lambda frac, msg: bar.progress(min(max(frac, 0.0), 1.0), text=msg),
                              use_cache=use_cache)


@st.cache_resource
def _registry() -> dict:
    return {"jobs": {}, "excel_lock": threading.Lock()}


def _session_id() -> str:
    ctx = get_script_run_ctx()
    return ctx.session_id if ctx else "default"


def job_state() -> dict:
    """This browser session's latest workbook job (other sessions never see it)."""
    return _registry()["jobs"].setdefault(_session_id(), {"status": "idle"})


def start_workbook_job(df: pd.DataFrame, path: Path, label: str):
    """Start a background build for exactly this data. Builds from all sessions
    share one Excel, so they queue instead of fighting over it."""
    reg = _registry()
    job = {"status": "queued", "frac": 0.0, "msg": "Waiting for Excel...", "path": None,
           "error": None, "label": label, "started": time_mod.time()}
    reg["jobs"][_session_id()] = job  # latest job replaces the displayed one

    def run():
        try:
            with reg["excel_lock"]:
                job.update(status="running", started=time_mod.time())

                def cb(frac, msg):
                    job["frac"], job["msg"] = frac, msg
                build_xlsb_fast(df, path, cb)
            job.update(status="done", path=str(path))
        except Exception as exc:
            job.update(status="error", error=str(exc))

    threading.Thread(target=run, daemon=True).start()


def zip_bytes(name: str, payload: bytes) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        zf.writestr(name, payload)
    return buf.getvalue()


def to_xlsx(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter",
                        engine_kwargs={"options": {"constant_memory": False, "nan_inf_to_errors": True}},
                        datetime_format="yyyy-mm-dd hh:mm:ss", date_format="yyyy-mm-dd") as xw:
        df.to_excel(xw, sheet_name="Sent Tracker", index=False)
        ws = xw.sheets["Sent Tracker"]
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(df), len(df.columns) - 1)
    return buf.getvalue()


st.title("Report Sent Tracker")

today = date.today()
with st.sidebar:
    st.header("Report window")
    start_date = st.date_input("From (00:00:00)", value=today.replace(day=1))
    end_date = st.date_input("To", value=today)
    workers = st.slider("Parallel DB connections", 1, 32, 16,
                        help="Days are fetched concurrently. Raise if the DB handles it; lower if it struggles.")
    use_cache = st.checkbox("Use cached past days", value=True,
                            help="Days older than the last 3 come from the local cache (much faster). "
                                 "Untick to re-fetch everything from the database.")
    auto_build = st.checkbox("Build workbook automatically after Generate", value=True)
    st.caption(
        f"Today's cutoff is {CUTOFF:%H:%M:%S}; earlier dates run to 23:59:59."
    )

start_dt = datetime.combine(start_date, time(0, 0, 0))
end_dt = datetime.combine(end_date, CUTOFF if end_date == today else time(23, 59, 59))
st.write(f"report_sent_on between **{start_dt:%Y-%m-%d %H:%M:%S}** and **{end_dt:%Y-%m-%d %H:%M:%S}**")

generate = st.button("Generate tracker", type="primary")

if generate:
    if start_dt > end_dt:
        st.error("'From' is after 'To'.")
    else:
        try:
            bar = st.progress(0.0, text="Starting...")
            df = load_data(start_dt, end_dt, workers, bar, use_cache)
            bar.empty()
            at = datetime.now()
            st.session_state["result"] = {"df": df, "at": at, "start": start_dt, "end": end_dt}
            st.session_state.pop("xlsx_zip", None)
            st.session_state.pop("csv_zip", None)
            if auto_build:
                label0 = at.strftime("%Y%m%d_%H%M%S")
                start_workbook_job(df, Path("output") / f"PS Sent & Not Sent_{label0}.xlsb", label0)
        except SQLAlchemyError as exc:
            st.error(f"Database error: {exc}")
        except Exception as exc:
            st.error(f"Failed: {exc}")

res = st.session_state.get("result")
if res:
    df = res["df"]
    stamp = res["at"].strftime("%Y%m%d_%H%M%S")
    st.success(f"{len(df):,} rows | generated {res['at']:%Y-%m-%d %H:%M:%S}")
    st.dataframe(df.head(200), use_container_width=True)
    st.caption("Preview of the first 200 rows.")

    d1, d2 = st.columns(2)
    if d1.button("Prepare CSV (zipped)"):
        with st.spinner("Zipping CSV..."):
            st.session_state["csv_zip"] = zip_bytes(f"report_sent_tracker_{stamp}.csv",
                                                    df.to_csv(index=False).encode("utf-8-sig"))
    if "csv_zip" in st.session_state:
        d1.download_button("Download CSV (zipped)", data=st.session_state["csv_zip"],
                           file_name=f"report_sent_tracker_{stamp}.zip", mime="application/zip")
    if len(df) >= 1_048_576:
        d2.warning("Too many rows for a single Excel sheet; use the CSV.")
    else:
        if d2.button("Prepare Excel (zipped)"):
            with st.spinner("Building Excel..."):
                st.session_state["xlsx_zip"] = zip_bytes(f"report_sent_tracker_{stamp}.xlsx", to_xlsx(df))
        if "xlsx_zip" in st.session_state:
            d2.download_button(
                "Download Excel (zipped)",
                data=st.session_state["xlsx_zip"],
                file_name=f"report_sent_tracker_{stamp}_xlsx.zip",
                mime="application/zip",
            )

    st.divider()
    st.subheader("PS Sent & Not Sent workbook")
    st.caption("Tabs: Sent Summary, Client Wise Check Sent, Cluster Wise Check Sent, Manual sentcount, "
               "Sent Cases (CTS) - Bridge. Pivot tabs are real Excel PivotTables, built by desktop Excel "
               "and saved as .xlsb. Manual sentcount has the cluster drop-down (B1) and live SUMIFS formulas.")
    if st.button("Build workbook (.xlsb with pivots)"):
        start_workbook_job(df, Path("output") / f"PS Sent & Not Sent_{stamp}.xlsb", stamp)

    @st.fragment(run_every=2)
    def workbook_status():
        job = job_state()
        status = job.get("status")
        if status == "idle":
            st.caption("No workbook built for this session yet.")
            return
        same = job.get("label") == stamp
        if not same:
            st.warning(f"The workbook below is for data generated at {job['label']}, not the data shown above "
                       f"({stamp}). Click 'Build workbook' to build it for the current data.")
        if status in ("queued", "running"):
            secs = int(time_mod.time() - job.get("started", time_mod.time()))
            st.progress(min(max(job.get("frac", 0.0), 0.0), 1.0),
                        text=f"{job.get('msg') or 'Working...'}  ({secs // 60}m {secs % 60:02d}s elapsed)")
            st.caption("Building in the background - you can keep using the app.")
        elif status == "error":
            st.error(f"Workbook failed: {job['error']}")
        elif status == "done" and job.get("path") and Path(job["path"]).exists():
            st.success(f"Saved: {Path(job['path']).resolve()}")
            with open(job["path"], "rb") as fh:
                st.download_button("Download workbook (.xlsb)", data=fh, file_name=Path(job["path"]).name,
                                   mime="application/vnd.ms-excel.sheet.binary.macroEnabled.12")

    workbook_status()
else:
    st.info("Click 'Generate tracker' to fetch data.")
