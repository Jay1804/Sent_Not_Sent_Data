"""
Streamlit app that generates the WIP, Insufficient and On-Hold trackers.

Reuses the exact DB-fetch / styling / TAT-formula logic from pipeline.py
and queries.py. Database credentials are read entirely from the backend
(.env, via pipeline.get_db_config) — there is no connection UI. Fetched
data is cached per tracker for CACHE_TTL_SECONDS to avoid re-hitting the
database on every rerun/click.
"""

import os
from datetime import datetime

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from sqlalchemy.exc import SQLAlchemyError

from pipeline import (
    TrackerBuilder,
    get_db_config,
    build_file_name,
    run_bridge_total_wip_final,
)
from queries import TRACKERS

load_dotenv()

st.set_page_config(page_title="Checkpoint Trackers", page_icon="\U0001F4CB", layout="wide")

CACHE_TTL_SECONDS = int(os.environ.get("CACHE_TTL_SECONDS", 900))


@st.cache_data(ttl=CACHE_TTL_SECONDS, show_spinner=False)
def load_tracker(tracker_name: str, data_source: str, holidays_override: tuple | None = None):
    """Fetch + build one tracker's report. Cached per (tracker_name,
    data_source, holidays_override) so repeated views within CACHE_TTL_SECONDS
    don't re-run the fetch, and a newly-uploaded holiday list correctly busts
    the cache instead of silently reusing a stale build. data_source is the
    user's explicit choice: 'live' always queries the DB and refreshes the
    local cache file for next time; 'cache' reads that local file instead
    (fails with a clear message if it doesn't exist yet — pick Live Database
    once first). holidays_override, if given, replaces the tracker's default
    holiday list (used by any NETWORKDAYS-based formula)."""
    cfg = TRACKERS[tracker_name]
    db_config = get_db_config()
    holidays = list(holidays_override) if holidays_override is not None else cfg["holidays"]
    builder = TrackerBuilder(
        db_config, cfg["query"], cfg["headers"], holidays,
        apply_tat_formulas=cfg.get("apply_tat_formulas", True),
        post_process=cfg.get("post_process"),
        cache_key=cfg.get("cache_key"),
        data_source=data_source,
        formula_columns=cfg.get("formula_columns"),
        formula_fn=cfg.get("formula_fn"),
        column_formats=cfg.get("column_formats"),
    )
    df, xlsx_buffer = builder.run()
    return df, xlsx_buffer.getvalue()


def dedupe_for_preview(df):
    """st.dataframe converts through PyArrow, which rejects duplicate column
    labels outright (e.g. Bridge_total_WIP_Final's two 'Bucket' columns, kept
    duplicated on purpose to match the requested layout) - so give the
    preview-only copy disambiguated labels rather than touching the real
    output columns."""
    if not df.columns.duplicated().any():
        return df
    seen = {}
    labels = []
    for c in df.columns:
        seen[c] = seen.get(c, 0) + 1
        labels.append(c if seen[c] == 1 else f"{c} ({seen[c]})")
    return df.set_axis(labels, axis=1)


def parse_holiday_file(uploaded_file):
    """First column of an uploaded .csv/.xlsx -> sorted list of 'YYYY-MM-DD'
    strings, matching the format the default HOLIDAYS list already uses."""
    if uploaded_file.name.lower().endswith(".csv"):
        raw = pd.read_csv(uploaded_file)
    else:
        raw = pd.read_excel(uploaded_file)
    if raw.empty:
        raise ValueError("The uploaded file has no rows.")
    dates = pd.to_datetime(raw.iloc[:, 0], errors="coerce").dropna()
    if dates.empty:
        raise ValueError("No valid dates found in the first column of the uploaded file.")
    return sorted(dates.dt.strftime("%Y-%m-%d").unique().tolist())


st.title("Checkpoint Trackers")
st.write("Generate the WIP, Insufficient, or On-Hold tracker directly from the database.")

with st.expander("Bridge_total_WIP_Final — generate directly", expanded=True):
    st.caption(
        "Runs the full pipeline in one click, exactly as `python pipeline.py \"Bridge_total_WIP_Final\"` "
        "does on the command line: fetches Bridge_total_WIP data, recalculates it in Excel, reshapes it "
        "into the Final column layout, recalculates that too, and adds the Summary pivot sheet. "
        "Requires Excel installed; can take several minutes on a large report."
    )

    direct_source_label = st.radio(
        "Data source",
        options=["Live Database", "Local Cache"],
        horizontal=True,
        key="direct_final_source",
        help=(
            "Live Database always queries the DB and refreshes the local cache for next time. "
            "Local Cache skips the DB entirely and reuses whatever Bridge_total_WIP data was saved "
            "on the last Live Database run."
        ),
    )
    direct_data_source = "live" if direct_source_label == "Live Database" else "cache"

    direct_holiday_file = st.file_uploader(
        "Holiday list (optional — .csv or .xlsx, dates in the first column)",
        type=["csv", "xlsx"],
        key="direct_final_holidays",
        help=(
            "Overrides Bridge_total_WIP's default holiday list for its NETWORKDAYS-based "
            "formulas, and keeps Bridge_total_WIP_Final's own ageing formulas in sync with the "
            "same list. Leave empty to use the built-in list."
        ),
    )
    direct_holidays = None
    if direct_holiday_file is not None:
        try:
            direct_holidays = parse_holiday_file(direct_holiday_file)
            st.caption(f"Using {len(direct_holidays)} uploaded holiday dates instead of the default list.")
        except Exception as exc:
            st.error(f"Couldn't read the holiday file: {exc}")
            direct_holidays = None

    if st.button("Generate Bridge_total_WIP_Final directly", type="primary", key="direct_final_generate"):
        try:
            with st.spinner("Running the full Bridge_total_WIP_Final pipeline... this can take several minutes"):
                direct_out_path = run_bridge_total_wip_final(
                    data_source=direct_data_source, holidays=direct_holidays
                )
            with open(direct_out_path, "rb") as f:
                direct_bytes = f.read()
            st.session_state["direct_final_result"] = {
                "xlsx_bytes": direct_bytes,
                "path": direct_out_path,
                "generated_at": datetime.now(),
            }
            st.success(f"Bridge_total_WIP_Final generated directly -> {direct_out_path}")
        except Exception as exc:
            st.error(f"Failed to build Bridge_total_WIP_Final: {exc}")

    direct_result = st.session_state.get("direct_final_result")
    if direct_result:
        st.caption(
            f"Generated at {direct_result['generated_at'].strftime('%Y-%m-%d %H:%M:%S')} "
            f"-> {direct_result['path']}"
        )
        st.download_button(
            label="Download Bridge_total_WIP_Final (.xlsx)",
            data=direct_result["xlsx_bytes"],
            file_name="Bridge_total_WIP_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="direct_final_download",
        )

st.divider()

tracker_name = st.radio(
    "Select tracker to generate",
    options=[name for name in TRACKERS if name != "Bridge_total_WIP"],
    horizontal=True,
)

tracker_cfg = TRACKERS[tracker_name]
st.caption(tracker_cfg["description"])

if "results" not in st.session_state:
    st.session_state["results"] = {}

supports_cache = bool(tracker_cfg.get("cache_key"))
if supports_cache:
    source_label = st.radio(
        "Data source",
        options=["Live Database", "Local Cache"],
        horizontal=True,
        help=(
            "Live Database always queries the DB and refreshes the local cache "
            "for next time. Local Cache skips the DB entirely and reuses whatever "
            "was saved on the last Live Database run."
        ),
    )
    data_source = "live" if source_label == "Live Database" else "cache"
else:
    data_source = "live"

uses_formulas = bool(tracker_cfg.get("apply_tat_formulas", True) or tracker_cfg.get("formula_fn"))
holidays_tuple = None
if uses_formulas:
    holiday_file = st.file_uploader(
        "Holiday list (optional — .csv or .xlsx, dates in the first column)",
        type=["csv", "xlsx"],
        help=(
            "Overrides this report's default holiday list for NETWORKDAYS-based "
            "formulas. Leave empty to use the built-in list."
        ),
    )
    if holiday_file is not None:
        try:
            holidays_tuple = tuple(parse_holiday_file(holiday_file))
            st.caption(f"Using {len(holidays_tuple)} uploaded holiday dates instead of the default list.")
        except Exception as exc:
            st.error(f"Couldn't read the holiday file: {exc}")
            holidays_tuple = None

col1, col2 = st.columns([1, 1])
generate = col1.button(f"Generate {tracker_name}", type="primary")
refresh = col2.button("Refresh (bypass app cache)")

if refresh:
    load_tracker.clear()
    generate = True

if generate:
    try:
        spinner_msg = (
            f"Fetching data and building {tracker_name}..."
            if data_source == "live"
            else f"Building {tracker_name} from local cache..."
        )
        with st.spinner(spinner_msg):
            df, xlsx_bytes = load_tracker(tracker_name, data_source, holidays_tuple)
        st.session_state["results"][tracker_name] = {
            "df": df,
            "xlsx_bytes": xlsx_bytes,
            "generated_at": datetime.now(),
            "data_source": data_source,
            "holidays": list(holidays_tuple) if holidays_tuple is not None else tracker_cfg["holidays"],
        }
        st.success(f"{tracker_name} generated from {source_label if supports_cache else 'the live database'}: {len(df)} rows.")
    except SQLAlchemyError as exc:
        st.error(f"Database error while fetching data: {exc}")
    except Exception as exc:
        st.error(f"Failed to generate {tracker_name}: {exc}")

result = st.session_state["results"].get(tracker_name)
if result:
    st.subheader("Preview")
    st.dataframe(dedupe_for_preview(result["df"].head(200)), use_container_width=True)
    st.caption(
        f"{len(result['df'])} total rows fetched from "
        f"{'the live database' if result['data_source'] == 'live' else 'the local cache'}. "
        f"Generated at {result['generated_at'].strftime('%Y-%m-%d %H:%M:%S')}. "
        f"Cached for {CACHE_TTL_SECONDS // 60} minutes. Showing first 200 rows above."
    )

    st.download_button(
        label=f"Download {tracker_name} (.xlsx)",
        data=result["xlsx_bytes"],
        file_name=build_file_name(tracker_cfg, result["generated_at"], "xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.download_button(
        label="Download raw data (.csv)",
        data=result["df"].to_csv(index=False).encode("utf-8"),
        file_name=build_file_name(tracker_cfg, result["generated_at"], "csv"),
        mime="text/csv",
    )
else:
    st.info("Click 'Generate' to fetch data and build the report.")
