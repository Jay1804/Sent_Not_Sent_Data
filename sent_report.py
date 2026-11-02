"""Builds the 'PS Sent & Not Sent' workbook: five tabs, values only (no formulas).

Sent Summary | Client Wise Check Sent | Cluster Wise Check Sent |
Manual sentcount | Sent Cases (CTS) - Bridge (the extracted dump)

The workbook is written with xlsxwriter in constant-memory mode, so every
sheet is written strictly row by row (summary tabs first, dump last).
"""

import subprocess
from pathlib import Path

import pandas as pd
import xlsxwriter

DUMP_SHEET = "Sent Cases (CTS) - Bridge"
DATE_FMT = "dd-mmm-yy"
MANUAL_ROWS = [  # label, dump column
    ("Address", "Address check count"),
    ("Database", "Database check count"),
    ("Education", "Education check count"),
    ("Employment", "Employment check count"),
]
DUMP_DATETIME_COLS = [
    "received_date", "created_date", "requested_on", "case_expected_closure_date",
    "report_delivery_date", "report_sent_on", "report_triggered_on",
    "Latest Reworked Date", "Last Check Closure Date", "Report writer Timing",
]


def _formats(wb):
    return {
        "title": wb.add_format({"bold": True, "font_size": 12}),
        "head": wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "center"}),
        "headl": wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1}),
        "cell": wb.add_format({"border": 1, "num_format": "#,##0"}),
        "text": wb.add_format({"border": 1}),
        "tot": wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1, "num_format": "#,##0"}),
        "totl": wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1}),
        "date": wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1,
                               "align": "center", "num_format": DATE_FMT}),
        "dump_date": wb.add_format({"num_format": "dd-mmm-yy hh:mm:ss"}),
    }


def _pivot_rows(df, cluster_col, status_col, sub_col):
    """Rows of a Cluster x Status x Sub pivot of row counts, with subtotals.
    Returns (header_rows, body_rows). Clusters sorted by grand total, desc."""
    statuses = [s for s in ["sent", "Not Sent", "New", "Report Generation Failed"]
                if s in set(df[status_col])] + sorted(set(df[status_col]) - {"sent", "Not Sent", "New", "Report Generation Failed"})
    cols = []  # (status, sub or None)
    for s in statuses:
        for sub in sorted(df.loc[df[status_col] == s, sub_col].dropna().unique()):
            cols.append((s, sub))
        cols.append((s, None))
    counts = df.groupby([cluster_col, status_col, sub_col]).size()
    by_status = df.groupby([cluster_col, status_col]).size()
    totals = df.groupby(cluster_col).size().sort_values(ascending=False)

    def row_for(c):
        out = []
        for s, sub in cols:
            out.append(int(by_status.get((c, s), 0)) if sub is None else int(counts.get((c, s, sub), 0)))
        return out + [int(totals[c])]

    head1 = [""]
    head2 = ["Cluster"]
    for s, sub in cols:
        head1.append(s if sub is not None and (not head1 or True) else f"{s} Total")
        head2.append(sub if sub is not None else "")
    head1.append("Grand Total")
    head2.append("")
    # label only the first column of each status group in the top header
    seen, top = set(), [""]
    for (s, sub), _ in zip(cols, range(len(cols))):
        if sub is None:
            top.append(f"{s} Total")
        elif s not in seen:
            seen.add(s)
            top.append(s)
        else:
            top.append("")
    top.append("Grand Total")

    body = [[c] + row_for(c) for c in totals.index]
    grand = ["Grand Total"]
    for j in range(len(cols) + 1):
        grand.append(sum(r[j + 1] for r in body))
    return [top, head2], body, grand


def _write_pivot(ws, r, title, df, status_col, sub_col, f):
    ws.write(r, 0, title, f["title"]); r += 1
    heads, body, grand = _pivot_rows(df, "Cluster", status_col, sub_col)
    for h in heads:
        ws.write_row(r, 0, h, f["head"]); r += 1
    for row in body:
        ws.write(r, 0, row[0], f["text"]); ws.write_row(r, 1, row[1:], f["cell"]); r += 1
    ws.write(r, 0, grand[0], f["totl"]); ws.write_row(r, 1, grand[1:], f["tot"]); r += 1
    return r


def _sheet_summary(wb, df, f):
    ws = wb.add_worksheet("Sent Summary")
    ws.set_column(0, 0, 22); ws.set_column(1, 20, 14)
    r = _write_pivot(ws, 0, "Cluster Wise Summary", df, "Report Status", "Type", f)
    _write_pivot(ws, r + 2, "ReportType Summary", df, "Report Status", "Report Type", f)


def _sheet_client(wb, df, f):
    ws = wb.add_worksheet("Client Wise Check Sent")
    ws.set_column(0, 0, 14); ws.set_column(1, 1, 60); ws.set_column(2, 2, 26)
    g = (df.groupby(["Client Code", "Client_name"], dropna=False)["Report Checks Count"].sum()
           .sort_values(ascending=False))
    ws.write(0, 0, "Bridge check Count Summary", f["title"])
    ws.write_row(1, 0, ["Client Code", "Client_name", "Sum of Report Checks Count"], f["headl"])
    r = 2
    for (code, name), n in g.items():
        ws.write(r, 0, "" if pd.isna(code) else code, f["text"])
        ws.write(r, 1, "" if pd.isna(name) else name, f["text"])
        ws.write(r, 2, int(n), f["cell"]); r += 1
    ws.write(r, 0, "Grand Total", f["totl"]); ws.write(r, 1, "", f["totl"]); ws.write(r, 2, int(g.sum()), f["tot"])


def _sheet_cluster(wb, df, f):
    ws = wb.add_worksheet("Cluster Wise Check Sent")
    ws.set_column(0, 0, 22); ws.set_column(1, 5, 26)
    cols = ["Report Checks Count", "Database check count", "Address check count",
            "Education check count", "Employment check count"]
    g = df.groupby("Cluster")[cols].sum().sort_values("Report Checks Count", ascending=False)
    ws.write(1, 0, "Cluster Wise", f["title"])
    ws.write_row(2, 0, ["Cluster"] + [f"Sum of {c}" for c in cols], f["headl"])
    r = 3
    for name, row in g.iterrows():
        ws.write(r, 0, name, f["text"]); ws.write_row(r, 1, [int(v) for v in row], f["cell"]); r += 1
    ws.write(r, 0, "Grand Total", f["totl"]); ws.write_row(r, 1, [int(v) for v in g.sum()], f["tot"])


def _sheet_manual(wb, df, f):
    """Per cluster: check counts of Manual-sent reports by LastDate."""
    ws = wb.add_worksheet("Manual sentcount")
    ws.set_column(0, 0, 3); ws.set_column(1, 1, 18); ws.set_column(2, 80, 11)
    man = df[df["Type"] == "Manual"].copy()
    man["_d"] = pd.to_datetime(man["LastDate"], format="%d-%b-%y", errors="coerce")
    all_days = pd.to_datetime(df["LastDate"], format="%d-%b-%y", errors="coerce").dropna()
    days = pd.date_range(all_days.min(), all_days.max()) if len(all_days) else []
    sums = man.groupby(["Cluster", "_d"])[[c for _, c in MANUAL_ROWS]].sum()
    order = df.groupby("Cluster").size().sort_values(ascending=False).index
    r = 0
    for cluster in order:
        ws.write(r, 1, cluster, f["title"]); ws.write(r, 3, "Manual Sent Count cluster wise", f["title"]); r += 1
        ws.write(r, 1, "Type", f["headl"])
        for j, d in enumerate(days):
            ws.write_datetime(r, 2 + j, d.to_pydatetime(), f["date"])
        ws.write(r, 2 + len(days), "Total", f["head"]); r += 1
        for label, col in MANUAL_ROWS:
            vals = [int(sums[col].get((cluster, d), 0)) if (cluster, d) in sums.index else 0 for d in days]
            ws.write(r, 1, label, f["text"]); ws.write_row(r, 2, vals, f["cell"])
            ws.write(r, 2 + len(days), sum(vals), f["tot"]); r += 1
        r += 1


def _sheet_dump(wb, df, f, progress=None):
    ws = wb.add_worksheet(DUMP_SHEET)
    cols = list(df.columns)
    ws.write_row(0, 0, cols)
    dt_idx = {i for i, c in enumerate(cols) if c in DUMP_DATETIME_COLS}
    data = df.astype(object).where(df.notna(), None)
    n = len(data)
    for r, row in enumerate(data.itertuples(index=False, name=None), start=1):
        for j, v in enumerate(row):
            if v is None:
                continue
            if j in dt_idx and hasattr(v, "to_pydatetime"):
                ws.write_datetime(r, j, v.to_pydatetime(), f["dump_date"])
            elif j in dt_idx and hasattr(v, "isoformat"):
                ws.write_datetime(r, j, pd.Timestamp(v).to_pydatetime(), f["dump_date"])
            else:
                ws.write(r, j, v)
        if progress and r % 20000 == 0:
            progress(r / n, f"Writing dump: {r:,}/{n:,} rows")
    ws.freeze_panes(1, 0)


def _placeholders(wb, f):
    """Empty tabs (titles only) that Excel fills with real PivotTables afterwards."""
    ws = wb.add_worksheet("Sent Summary")
    ws.write(0, 0, "Cluster Wise Summary", f["title"])
    ws.write(17, 0, "ReportType Summary", f["title"])
    ws = wb.add_worksheet("Client Wise Check Sent")
    ws.write(0, 0, "Bridge check Count Summary", f["title"])
    ws = wb.add_worksheet("Cluster Wise Check Sent")
    ws.write(1, 0, "Cluster Wise", f["title"])


def build_workbook(df: pd.DataFrame, path, progress=None, real_pivots=True) -> Path:
    """real_pivots=True leaves the three pivot tabs empty for add_pivots();
    False fills them with static value tables instead."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(path), {"constant_memory": True, "nan_inf_to_errors": True,
                                         "strings_to_urls": False})
    f = _formats(wb)
    if real_pivots:
        _placeholders(wb, f)
        _sheet_manual(wb, df, f)
    else:
        for step in (_sheet_summary, _sheet_client, _sheet_cluster, _sheet_manual):
            step(wb, df, f)
    _sheet_dump(wb, df, f, progress)
    wb.close()
    return path


_PIVOT_PS1 = Path(__file__).parent / "tools" / "make_pivots.ps1"


def add_pivots(xlsx_path, out_path, as_xlsb=True) -> Path:
    """Open the workbook in desktop Excel, create the real PivotTables on the dump
    and save as .xlsb (or .xlsx). Slow on a month of data; needs Excel installed."""
    xlsx_path, out_path = Path(xlsx_path).resolve(), Path(out_path).resolve()
    r = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(_PIVOT_PS1),
         str(xlsx_path), str(out_path), "50" if as_xlsb else "51"],
        capture_output=True, text=True)
    if r.returncode != 0 or not out_path.exists():
        raise RuntimeError(f"Excel pivot step failed: {(r.stderr or r.stdout).strip()[:800]}")
    return out_path


def convert_to_xlsb(xlsx_path, xlsb_path) -> Path:
    """Re-save as .xlsb through desktop Excel (COM). Slow on big files."""
    xlsx_path, xlsb_path = Path(xlsx_path).resolve(), Path(xlsb_path).resolve()
    script = (
        "$ErrorActionPreference='Stop';"
        "$x=New-Object -ComObject Excel.Application;$x.Visible=$false;$x.DisplayAlerts=$false;"
        "try{"
        f"$w=$x.Workbooks.Open('{xlsx_path}');"
        f"$w.SaveAs('{xlsb_path}',50);$w.Close($false)"
        "}finally{$x.Quit()}"
    )
    subprocess.run(["powershell", "-NoProfile", "-Command", script], check=True,
                   capture_output=True, text=True)
    return xlsb_path


# ------------------------------------------------------------ fast path
import csv
import json
import tempfile

IMPORT_BLOCKS = 8
_FAST_PS1 = Path(__file__).parent / "tools" / "make_workbook.ps1"


def _manual_meta(df):
    """Drop-down choices and date range for the Manual sentcount formulas."""
    days = pd.to_datetime(df["LastDate"], format="%d-%b-%y", errors="coerce").dropna()
    first = days.min().normalize() if len(days) else pd.Timestamp.today().normalize().replace(day=1)
    month_end = (first + pd.offsets.MonthEnd(0)).normalize()
    last = max(month_end, days.max().normalize()) if len(days) else month_end
    clusters = list(df.groupby("Cluster").size().sort_values(ascending=False).index)
    return {
        "clusters": clusters,
        "default_cluster": clusters[0] if clusters else "",
        "start_serial": (first - pd.Timestamp("1899-12-30")).days,
        "ndays": int((last - first).days) + 1,
    }


def build_xlsb_fast(df: pd.DataFrame, out_path, progress=None) -> Path:
    """CSV -> Excel's native importer -> formulas + pivots -> .xlsb. Avoids
    writing tens of millions of cells from Python one by one."""
    out_path = Path(out_path).resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)
    progress = progress or (lambda f, m: None)
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        progress(0.05, "Writing CSV for Excel...")
        out = df.copy()
        # real dates in Excel (so SUMIFS on LastDate works); shown as dd-mmm-yy
        out["LastDate"] = pd.to_datetime(out["LastDate"], format="%d-%b-%y", errors="coerce").dt.strftime("%Y-%m-%d")
        # strings stay text (IDs, comments starting with '=' etc.); numbers/dates auto-detect
        text_cols = [i + 1 for i, c in enumerate(out.columns)
                     if not (pd.api.types.is_numeric_dtype(out[c]) or pd.api.types.is_datetime64_any_dtype(out[c]))
                     and c not in DUMP_DATETIME_COLS and c != "LastDate"]
        for c in out.columns:  # tabs / line breaks would break the tab-delimited import
            if out[c].dtype == object or str(out[c].dtype) == "string":
                out[c] = out[c].astype("string").str.replace(r"[\t\r\n]+", " ", regex=True)
        # row blocks imported side by side by Excel; first block carries the header
        import numpy as np
        n_blocks = max(1, min(IMPORT_BLOCKS, len(out) // 20000 or 1))
        chunks, row = [], 1
        for i, idx in enumerate(np.array_split(np.arange(len(out)), n_blocks)):
            f = tmp / f"dump_{i}.txt"
            out.iloc[idx].to_csv(f, index=False, header=(i == 0), sep=chr(9), quoting=csv.QUOTE_NONE,
                                 encoding="utf-8-sig",  # BOM: Excel imports ~40x faster with it
                                 date_format="%Y-%m-%d %H:%M:%S")
            chunks.append({"file": str(f), "row": row})
            row += len(idx) + (1 if i == 0 else 0)
        meta = _manual_meta(df)
        meta.update({"chunks": chunks, "text_cols": text_cols, "date_cols": ["LastDate", "received_date", "case_expected_closure_date", "report_delivery_date"],
                     "datetime_cols": [c for c in DUMP_DATETIME_COLS if c not in ("received_date", "case_expected_closure_date", "report_delivery_date")], "ncols": len(out.columns)})
        (tmp / "meta.json").write_text(json.dumps(meta))
        progress(0.3, "Excel is importing the data, building formulas, pivots and saving .xlsb...")
        r = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(_FAST_PS1),
             "unused", str(tmp / "meta.json"), str(out_path), "50"],
            capture_output=True, text=True)
    if r.returncode != 0 or not out_path.exists():
        raise RuntimeError(f"Excel step failed: {(r.stderr or r.stdout).strip()[:800]}")
    progress(1.0, "Done")
    return out_path
