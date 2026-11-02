"""Query + post-processing for the Report Sent tracker.

The original query ran ~30 correlated sub-selects per report row. Here each
group of them is collapsed into one pre-aggregated derived table (one scan per
table, restricted to the reports in the date window) and joined back by key.
Vertical mapping and Acc_TAT are done in pandas instead of giant CASE/IN lists.
"""

import numpy as np
import pandas as pd

from cluster_map import cluster_for

CLIENT_SPECIAL = 363362

# Check families -> output column label (case-level counts, status <> 9)
FAMILY_COLS = [
    (1, "Address Checks Count"),
    (2, "DB Checks Count"),
    (3, "Drug Checks Count"),
    (4, "Edu Checks Count"),
    (5, "Emp Checks Count"),
    (7, "PV Checks Count"),
    (8, "Documentation Checks Count"),
    (13, "Due Delegence Checks Count"),
    (16, "Financial Checks Count"),
    (6, "Fingerprint Checks Count"),
    (9, "Business Checks Count"),
    (12, "DL Checks Count"),
    (10, "NBGV Checks Count"),
    (15, "Overseas Checks Count"),
]

# Report-level counts (report check status = 8)
REPORT_FAMILY_COLS = [
    (1, "Address check count"),
    (4, "Education check count"),
    (5, "Employment check count"),
]

_VERTICALS = {
    "BFSI": [
        269238, 272029, 323849, 420657, 436963, 452055, 460685, 462256, 462586, 462941,
        468958, 469477, 471542, 472161, 475280, 480813, 482937, 484956, 486727,
    ],
    "Business Solution": [489999],
    "Business Solutions": [
        218983, 281228, 308074, 322549, 374928, 403149, 411152, 421629, 430943, 475549,
        476148, 479484, 484390, 486907,
    ],
    "Workforce": [
        78249, 78242, 75598, 75731, 75721, 78167, 78475, 78534, 75778, 78698, 78154, 76351,
        78619, 76376, 76235, 79002, 82100, 85909, 94712, 95110, 103727, 108942, 110059,
        110150, 113404, 115012, 115550, 116525, 117162, 117191, 119015, 119149, 122166,
        123403, 124921, 125382, 125861, 126912, 131479, 133454, 135927, 138191, 138889,
        140679, 140707, 142320, 142368, 142677, 142762, 142856, 145896, 147341, 147932,
        148852, 150787, 152319, 153798, 155331, 155428, 155829, 155946, 158686, 162395,
        162999, 164034, 164339, 164669, 165872, 168054, 169441, 170092, 170765, 171086,
        171621, 172727, 172989, 173339, 173478, 173963, 175389, 176765, 177013, 177051,
        177224, 178430, 178926, 179110, 181532, 183165, 183561, 183630, 184256, 185214,
        186138, 186198, 186706, 187479, 188843, 189053, 189436, 189676, 190146, 190463,
        190470, 191499, 193017, 193806, 194152, 198170, 200404, 201613, 201681, 201816,
        201817, 201872, 203295, 203457, 206002, 207495, 208560, 208596, 211905, 211974,
        212820, 213216, 213648, 214690, 214967, 215415, 215875, 216126, 216195, 216204,
        216542, 217378, 217549, 217893, 218166, 218684, 219325, 219707, 220823, 224285,
        225217, 225223, 225285, 225506, 225543, 227376, 2,
    ],
}

# First list wins on a duplicate id, same as the original CASE ordering.
VERTICAL_BY_COMPANY = {}
for _name, _ids in _VERTICALS.items():
    for _id in _ids:
        VERTICAL_BY_COMPANY.setdefault(_id, _name)


def _flex():
    return ",\n    ".join(f"ecff.case_flex_field{i}" for i in range(1, 21))


BASE_SQL = f"""
SELECT
    company_name AS Client_name,
    CONCAT_WS(' ', first_name, Middle_name, Last_name) AS Candidate_name,
    Process_name,
    office_name AS location,
    ecm.case_id,
    case_ars_no,
    received_date,
    ecm.created_date,
    requested_on,
    ecm.case_expected_closure_date,
    {_flex()},
    report_delivery_date,
    report_sent_on,
    CASE report_type WHEN 0 THEN 'Interim' WHEN 1 THEN 'Final' WHEN 2 THEN 'Additional' END AS `Report Type`,
    IF(ecr.client_id = {CLIENT_SPECIAL}, ecr.spoc_comments, '') AS `Report Type 1`,
    report_severity,
    CONCAT(eud.User_first_name, ' ', eud.User_last_name) AS PS,
    CONCAT(eud1.User_first_name, ' ', eud1.User_last_name) AS QC_PS,
    checkpoint_live.fn_report_status(report_status) AS `Report Status`,
    requested_on AS report_triggered_on,
    IF(DATE(report_sent_on) > report_delivery_date, 'OT', 'IT') AS `Report Sent TAT`,
    IF(DATE(requested_on) > report_delivery_date, 'OT', 'IT') AS `Report Triggered TAT`,
    report_sub_name,
    IF(ecm.client_id = {CLIENT_SPECIAL}, DATEDIFF(CURRENT_DATE, received_date), NULL) AS Acc_Ageing,
    company_id AS _company_id,
    ecr.case_report_id AS _rid,
    ecr.case_id AS _cid
FROM ec_case_reports ecr
LEFT JOIN ec_case_master ecm ON ecr.case_id = ecm.case_id
LEFT JOIN ec_case_candidates ec ON ec.candidate_id = ecm.candidate_id
LEFT JOIN ec_master_company emc ON ecr.CLIENT_ID = emc.COMPANY_ID
LEFT JOIN ec_user_details eud ON ecr.assigned_to = eud.user_id
LEFT JOIN ec_user_details eud1 ON ecr.qc_assigned_to = eud1.user_id
LEFT JOIN ec_client_process ecp ON ecm.process_id = ecp.process_id
LEFT JOIN ec_master_company_locations emcl ON ecm.client_office_id = emcl.office_id
LEFT JOIN ec_case_fields ecff ON ecm.case_id = ecff.case_id
WHERE ecr.report_sent_on >= :start AND ecr.report_sent_on <= :end
"""

_case_fam = ",\n    ".join(f"SUM(ecc.family_id = {fid}) AS `{label}`" for fid, label in FAMILY_COLS)
_rep_fam = ",\n    ".join(f"SUM(ecc.family_id = {fid}) AS `{label}`" for fid, label in REPORT_FAMILY_COLS)

# Aux queries: (key column, id kind, sql). Each takes an expanding :ids list.
AUX_SQL = {
    "cc": ("_cid", "case", f"""
SELECT ecc.case_id AS _cid,
    COUNT(*) AS `Case Check Count`,
    GROUP_CONCAT(ec2.check_name) AS `Check Names`,
    MAX(ecc.check_closure_date) AS `Last Check Closure Date`,
    {_case_fam}
FROM ec_case_checks ecc
LEFT JOIN ec_checks ec2 ON ecc.check_id = ec2.check_id
WHERE ecc.case_id IN :ids AND ecc.check_status <> 9
GROUP BY ecc.case_id"""),
    "rc": ("_rid", "report", f"""
SELECT ecrc.case_report_id AS _rid,
    COUNT(*) AS `Report Checks Count`,
    SUM(ecc.family_id NOT IN (1, 4, 5)) AS `Database check count`,
    {_rep_fam}
FROM ec_case_report_checks ecrc
LEFT JOIN ec_case_checks ecc ON ecrc.case_check_id = ecc.case_check_id
WHERE ecrc.case_report_id IN :ids AND ecrc.check_status = 8
GROUP BY ecrc.case_report_id"""),
    "rw": ("_rid", "report", """
SELECT case_report_id AS _rid, MAX(reworked_date) AS `Latest Reworked Date`
FROM ec_case_rework_details
WHERE case_report_id IN :ids
GROUP BY case_report_id"""),
    "ai": ("_rid", "report", """
SELECT p.case_report_id AS _rid,
       IF(p.response_status IS NULL AND p.attempt > 0, 'In Process', p.response_status) AS `Report AI Status`
FROM ec_check_ai_processing p
JOIN (SELECT case_report_id, MAX(id) AS mid
      FROM ec_check_ai_processing
      WHERE case_report_id IN :ids
      GROUP BY case_report_id) m ON p.id = m.mid"""),
    "hist": ("_cid", "case", """
SELECT case_id AS _cid, action_id, action_taken, action_comments, action_taken_on
FROM ec_case_history
WHERE case_id IN :ids AND check_id IS NULL
  AND action_taken IN ('Report writer\\'s comment',
                       'Report QC Status Marked as Accepted',
                       'Report QC Status Marked as Rejected')"""),
}

FINAL_ORDER_TAIL = (
    ["Latest Reworked Date", "Case Check Count", "Check Names"]
    + [l for _, l in FAMILY_COLS]
    + ["Last Check Closure Date", "Report Checks Count"]
    + [l for _, l in REPORT_FAMILY_COLS]
    + ["Database check count", "Report Triggered TAT", "Report writer", "Report writer Timing",
       "Report QC Status", "Report QC Comment", "report_sub_name", "Acc_Ageing", "Acc_TAT", "Vertical", "Report AI Status",
       "Date", "Hour", "Cluster", "Client Code", "Type", "LastDate"]
)


def history_frames(hist: pd.DataFrame):
    """Latest writer comment and earliest QC decision per case, by action_id."""
    if hist.empty:
        empty = pd.DataFrame({"_cid": []})
        return empty, empty
    w = hist[hist["action_taken"] == "Report writer's comment"]
    w = w.loc[w.groupby("_cid")["action_id"].idxmax()]
    w = w.rename(columns={"action_comments": "Report writer", "action_taken_on": "Report writer Timing"})
    w = w[["_cid", "Report writer", "Report writer Timing"]]
    q = hist[hist["action_taken"] != "Report writer's comment"]
    q = q.loc[q.groupby("_cid")["action_id"].idxmin()].copy()
    q["Report QC Status"] = q["action_taken"].map(
        lambda a: "Accepted" if a == "Report QC Status Marked as Accepted" else "Rejected")
    q = q.rename(columns={"action_comments": "Report QC Comment"})
    return w, q[["_cid", "Report QC Status", "Report QC Comment"]]


def assemble(base: pd.DataFrame, aux: dict) -> pd.DataFrame:
    """Merge the aux aggregates onto the base rows and finish derived columns."""
    df = base
    w, q = history_frames(aux["hist"])
    for frame, key in ((aux["cc"], "_cid"), (aux["rc"], "_rid"), (aux["rw"], "_rid"), (aux["ai"], "_rid"), (w, "_cid"), (q, "_cid")):
        if len(frame):
            df = df.merge(frame.drop_duplicates(key), on=key, how="left")
    for col in FINAL_ORDER_TAIL:
        if col not in df.columns:
            df[col] = None
    df = post_process(df)
    head = [c for c in base.columns if c not in ("_company_id", "_rid", "_cid", "report_sub_name", "Acc_Ageing", "Report Triggered TAT")]
    return df[head + FINAL_ORDER_TAIL]


def compute_acc_tat(df: pd.DataFrame) -> pd.DataFrame:
    age = pd.to_numeric(df["Acc_Ageing"], errors="coerce")
    sub = df["report_sub_name"]
    one_a, one_b, final = sub == "Report 1A", sub == "Report 1B", sub == "Final"
    conds = [
        one_a & age.between(0, 1), one_a & (age == 2), one_a & (age > 2),
        one_b & (age <= 10), one_b & (age == 11), one_b & (age > 11),
        final & (age <= 34), final & (age == 35), final & (age > 35),
    ]
    choices = ["IT", "AT", "OT"] * 3
    df["Acc_TAT"] = np.select(conds, choices, default=None)
    df["Acc_TAT"] = df["Acc_TAT"].where(df["Acc_TAT"].notna(), None)
    return df


def post_process(df: pd.DataFrame) -> pd.DataFrame:
    """Acc_TAT and Vertical, replacing the SQL session-variable / giant CASE logic."""
    df = compute_acc_tat(df)
    df["Vertical"] = (
        pd.to_numeric(df["_company_id"], errors="coerce")
        .map(VERTICAL_BY_COMPANY)
        .fillna("Not Found")
    )
    sent = pd.to_datetime(df["report_sent_on"], errors="coerce")
    df["Date"] = sent.dt.strftime("%d-%b-%y")
    df["Hour"] = sent.dt.strftime("%H")
    df["Cluster"] = cluster_for(df["Client_name"])
    df["Client Code"] = df["case_ars_no"].astype("string").str[:4]
    df["Type"] = df["PS"].isin(["Neetu Yadav2", "System User"]).map({True: "System User", False: "Manual"})
    # before 08:00 counts towards the previous day
    df["LastDate"] = sent.where(sent.dt.hour >= 8, sent - pd.Timedelta(days=1)).dt.strftime("%d-%b-%y")

    count_cols = [c for c in df.columns if c.lower().endswith("count") and not c.startswith("Last")]
    df[count_cols] = df[count_cols].apply(pd.to_numeric, errors="coerce").fillna(0).astype(int)
    return df.drop(columns=["_company_id"])


# ---------------------------------------------------------------- fetching
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, time, timedelta

from sqlalchemy import bindparam, text
from sqlalchemy.exc import SQLAlchemyError

BATCH = 1500


def day_chunks(start: datetime, end: datetime):
    """Split [start, end] into one window per calendar day."""
    chunks, cur = [], start
    while cur <= end:
        day_end = datetime.combine(cur.date(), time(23, 59, 59))
        chunks.append((cur, min(day_end, end)))
        cur = datetime.combine(cur.date() + timedelta(days=1), time(0, 0, 0))
    return chunks


def _run(engine, stmt, params, retries=1):
    for attempt in range(retries + 1):
        try:
            with engine.connect() as conn:
                return pd.read_sql(stmt, conn, params=params)
        except SQLAlchemyError:
            if attempt == retries:
                raise


def fetch_range(engine, start, end, workers=6, progress=None):
    """Stage 1: one cheap base query per day. Stage 2: the per-check/history
    aggregates in id batches, by index lookup. Every query runs on the pool."""
    progress = progress or (lambda frac, msg: None)
    base_stmt = text(BASE_SQL)

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futs = [pool.submit(_run, engine, base_stmt, {"start": s, "end": e}) for s, e in day_chunks(start, end)]
        frames = []
        for i, f in enumerate(as_completed(futs), 1):
            frames.append(f.result())
            progress(0.3 * i / len(futs), f"Base rows: {i}/{len(futs)} days")
        base = pd.concat(frames, ignore_index=True)
        if base.empty:
            return base

        case_ids = base["_cid"].dropna().astype(int).unique().tolist()
        rep_ids = base["_rid"].dropna().astype(int).unique().tolist()
        jobs = []
        for name, (_, kind, sql) in AUX_SQL.items():
            stmt = text(sql).bindparams(bindparam("ids", expanding=True))
            ids = case_ids if kind == "case" else rep_ids
            for i in range(0, len(ids), BATCH):
                jobs.append((name, pool.submit(_run, engine, stmt, {"ids": ids[i:i + BATCH]})))
        parts = {name: [] for name in AUX_SQL}
        fut_name = {f: n for n, f in jobs}
        for i, f in enumerate(as_completed(fut_name), 1):
            parts[fut_name[f]].append(f.result())
            progress(0.3 + 0.7 * i / len(jobs), f"Aggregates: {i}/{len(jobs)} batches")

    aux = {n: (pd.concat([f for f in p if not f.empty], ignore_index=True) if any(not f.empty for f in p) else pd.DataFrame())
           for n, p in parts.items()}
    base = base.sort_values(["report_sent_on", "_rid"], kind="stable", ignore_index=True)
    return assemble(base, aux)


# ------------------------------------------------------------ day cache
import hashlib
from pathlib import Path

CACHE_ROOT = Path(".cache") / "days"
FRESH_DAYS = 3  # today and the 2 days before are always re-fetched


def _cache_dir() -> Path:
    """Cache is keyed on the query/mapping code, so editing either invalidates it."""
    h = hashlib.md5()
    for f in ("sent_tracker_query.py", "cluster_map.py"):
        h.update(Path(__file__).with_name(f).read_bytes())
    d = CACHE_ROOT / h.hexdigest()[:10]
    d.mkdir(parents=True, exist_ok=True)
    return d


def refresh_ageing(df: pd.DataFrame) -> pd.DataFrame:
    """Acc_Ageing/Acc_TAT depend on today's date, so recompute them for cached rows."""
    mask = df["Acc_Ageing"].notna()
    if mask.any():
        recv = pd.to_datetime(df.loc[mask, "received_date"], errors="coerce")
        df["Acc_Ageing"] = pd.to_numeric(df["Acc_Ageing"], errors="coerce")
        df.loc[mask, "Acc_Ageing"] = (pd.Timestamp.today().normalize() - recv).dt.days
        df = compute_acc_tat(df)
    return df


def fetch_range_cached(engine, start, end, workers=6, progress=None, use_cache=True):
    """Past days come from the local cache; recent days are fetched fresh and
    (when complete) saved for next time."""
    progress = progress or (lambda frac, msg: None)
    if not use_cache:
        return fetch_range(engine, start, end, workers, progress)
    cdir = _cache_dir()
    today = pd.Timestamp.today().normalize()
    cutoff = today - pd.Timedelta(days=FRESH_DAYS - 1)
    cached, need = [], []
    for s_, e_ in day_chunks(start, end):
        day = pd.Timestamp(s_.date())
        f = cdir / f"{day:%Y-%m-%d}.pkl"
        full_day = s_.time() == time(0, 0, 0) and e_.time() == time(23, 59, 59)
        if full_day and day < cutoff and f.exists():
            cached.append(pd.read_pickle(f))
        else:
            need.append((s_, e_))
    progress(0.02, f"{len(cached)} day(s) from cache, {len(need)} to fetch")
    fresh = pd.DataFrame()
    if need:
        fresh = fetch_range(engine, need[0][0], need[-1][1], workers,
                            lambda fr, m: progress(0.02 + 0.98 * fr, m))
        if len(fresh):
            days = pd.to_datetime(fresh["report_sent_on"]).dt.normalize()
            for day, part in fresh.groupby(days):
                complete = day < today and any(s_.date() == day.date() and e_.time() == time(23, 59, 59) for s_, e_ in need)
                if complete:
                    part.to_pickle(cdir / f"{day:%Y-%m-%d}.pkl")
    frames = [f for f in cached + [fresh] if len(f)]
    if not frames:
        return fresh
    df = pd.concat(frames, ignore_index=True)
    df = refresh_ageing(df)
    return df.sort_values(["report_sent_on", "case_id"], kind="stable", ignore_index=True)
