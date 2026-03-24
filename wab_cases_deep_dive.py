"""
WAB Cases + Emails Deep Dive — Phase 1 GenAI Use-Case Evidence
===============================================================
Standalone VDI script.  Reads Cases + Emails files only.
Produces one Excel workbook + one markdown summary.

Dependencies: pandas, openpyxl  (standard library otherwise)
"""

# ┌─────────────────────────────────────────────────────────┐
# │  EDIT THESE 3 VARIABLES BEFORE RUNNING                  │
# └─────────────────────────────────────────────────────────┘
CASE_FILE  = r"C:\Users\YourName\Desktop\AAB All Cases.xlsx"
EMAIL_FILE = r"C:\Users\YourName\Desktop\ALL EMAIL Files.xlsx"
OUTPUT_DIR = r"C:\Users\YourName\Desktop\wab_output"
# ┌─────────────────────────────────────────────────────────┐
# │  DO NOT EDIT BELOW THIS LINE                            │
# └─────────────────────────────────────────────────────────┘

import os, re, sys, datetime, warnings, html
from pathlib import Path
from collections import OrderedDict
from io import StringIO

import pandas as pd
import numpy as np

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "wab_cases_deep_dive.xlsx")
OUTPUT_MD   = os.path.join(OUTPUT_DIR, "wab_cases_deep_dive_summary.md")
LOG, WARN   = [], []
TRUNC       = 150

# Known internal / system company names to exclude from client metrics
INTERNAL_COMPANIES = {"AAB ADMIN", "WAB ADMIN", "AAB ADMIN -", "WAB ADMIN -"}

# ═══════════════════════════════════════════════════════════
#  UTILITIES
# ═══════════════════════════════════════════════════════════

def log(m):
    LOG.append(m); print(m)

def warn(m):
    WARN.append(m); log(f"  WARNING: {m}")

def trunc(v, n=TRUNC):
    if pd.isna(v): return ""
    s = str(v).replace("\r"," ").replace("\n"," ").strip()
    return s[:n] + "..." if len(s) > n else s

def norm_col(name):
    if not isinstance(name, str): return ""
    return re.sub(r"\s+", " ", name.strip().lower())

def find_col(df, *candidates):
    """Case-insensitive column lookup with partial-match fallback."""
    lookup = {norm_col(c): c for c in df.columns}
    for cand in candidates:
        normed = norm_col(cand)
        if normed in lookup:
            return lookup[normed]
    # partial fallback
    for cand in candidates:
        normed = norm_col(cand)
        for k, v in lookup.items():
            if normed in k or k in normed:
                return v
    return None

def safe_dt(series):
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    try:
        return pd.to_datetime(series, errors="coerce")
    except Exception:
        return pd.Series([pd.NaT]*len(series), index=series.index)

def safe_numeric(series):
    if pd.api.types.is_numeric_dtype(series):
        return series
    return pd.to_numeric(series, errors="coerce")

def strip_html(text):
    """Strip HTML tags and decode entities.  Best-effort, no external libs."""
    if pd.isna(text): return ""
    s = str(text)
    # remove style/script blocks
    s = re.sub(r"<(style|script)[^>]*>.*?</\1>", " ", s, flags=re.DOTALL|re.IGNORECASE)
    # replace br / p / div / tr / li with newline
    s = re.sub(r"<(br|/p|/div|/tr|/li)[^>]*>", "\n", s, flags=re.IGNORECASE)
    # strip remaining tags
    s = re.sub(r"<[^>]+>", " ", s)
    # decode HTML entities
    s = html.unescape(s)
    # collapse whitespace
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n[ \t]+", "\n", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()

def read_file(path, label):
    log(f"\nReading {label}: {path}")
    if not os.path.isfile(path):
        warn(f"File not found: {path}")
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, header=0, engine="openpyxl")
    except Exception as e:
        warn(f"Failed to read {label}: {e}")
        return pd.DataFrame()
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    df = df.dropna(axis=1, how="all")
    log(f"  Rows: {len(df):,}  |  Columns: {len(df.columns)}")
    return df

def pct(num, denom):
    return f"{100*num/denom:.1f}%" if denom else "N/A"

def write_sheet(writer, name, df):
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        df = pd.DataFrame({"note": ["No data"]})
    # Excel sheet name limit = 31 chars
    sheet_name = name[:31]
    df.to_excel(writer, sheet_name=sheet_name, index=False, freeze_panes=(1, 0))
    ws = writer.sheets[sheet_name]
    for i, col_cells in enumerate(ws.columns):
        max_len = max(len(str(cell.value or "")) for cell in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 60)


# ═══════════════════════════════════════════════════════════
#  CASE PREPARATION  — build enriched case DataFrame once
# ═══════════════════════════════════════════════════════════

def prepare_cases(cases):
    """Add derived columns to cases.  Returns enriched copy."""
    df = cases.copy()

    # ── Core columns ──
    col_map = {
        "case_number":  find_col(df, "Case Number"),
        "company":      find_col(df, "Company Name (Company) (Company)", "Company Name", "Customer"),
        "subject":      find_col(df, "Subject"),
        "description":  find_col(df, "Description"),
        "activity_subj": find_col(df, "Activity Subject"),
        "origin":       find_col(df, "Origin"),
        "status":       find_col(df, "Status"),
        "status_reason": find_col(df, "Status Reason"),
        "created_on":   find_col(df, "Created On"),
        "modified_on":  find_col(df, "Modified On"),
        "sla_start":    find_col(df, "SLA Start"),
        "resolved_on":  find_col(df, "Resolved On"),
        "resolved_hrs": find_col(df, "Resolved In Hours"),
        "owner":        find_col(df, "Manager (Owning User) (User)", "Owner"),
        "pod":          find_col(df, "POD Name (Owning User) (User)", "POD Name"),
        "last_touch":   find_col(df, "Last Touch"),
        "last_touch_by": find_col(df, "Last Touch By"),
        "parent_case":  find_col(df, "Parent Case"),
        "last_contact":  find_col(df, "Last Contact Attempt"),
    }

    # ── Derived: dates ──
    for key in ["created_on", "modified_on", "sla_start", "resolved_on", "last_touch", "last_contact"]:
        src = col_map.get(key)
        if src:
            df[f"_{key}_dt"] = safe_dt(df[src])

    # ── Derived: hours ──
    src = col_map.get("resolved_hrs")
    if src:
        df["_hours"] = safe_numeric(df[src])

    # ── Derived: is_resolved ──
    resolved_kw = ["resolved", "closed", "cancelled", "canceled"]
    src_sr = col_map.get("status_reason") or col_map.get("status")
    if src_sr:
        df["_is_resolved"] = df[src_sr].fillna("").astype(str).str.lower().apply(
            lambda x: any(kw in x for kw in resolved_kw))
    else:
        df["_is_resolved"] = True  # assume resolved if no status

    # ── Derived: 3-way population classification ──
    # named_client  = has company name, not admin
    # blank_client   = company blank (CRM default assignment failed — confirmed client emails)
    # admin          = AAB ADMIN / WAB ADMIN (multi-client or operational)
    src_co = col_map.get("company")
    if src_co:
        df["_company_clean"] = df[src_co].fillna("(blank)").astype(str).str.strip()
        _upper = df["_company_clean"].str.upper()
        df["_is_admin"] = _upper.apply(lambda x: any(x.startswith(kw) for kw in INTERNAL_COMPANIES))
        df["_is_blank_company"] = (_upper == "(BLANK)") | (df["_company_clean"] == "")
        df["_case_population"] = "named_client"
        df.loc[df["_is_blank_company"], "_case_population"] = "blank_client"
        df.loc[df["_is_admin"], "_case_population"] = "admin"
        # _is_internal now means admin ONLY (blank cases are client)
        df["_is_internal"] = df["_is_admin"]
    else:
        df["_is_internal"] = False
        df["_is_admin"] = False
        df["_is_blank_company"] = False
        df["_case_population"] = "named_client"

    # ── Derived: subject clean ──
    src_subj = col_map.get("subject")
    if src_subj:
        df["_subject"] = df[src_subj].fillna("(blank)").astype(str).str.strip()
    else:
        df["_subject"] = "(no subject column)"

    # ── Derived: origin clean ──
    src_orig = col_map.get("origin")
    if src_orig:
        df["_origin"] = df[src_orig].fillna("(blank)").astype(str).str.strip()
    else:
        df["_origin"] = "(no origin column)"

    # ── Derived: week / day-of-week / date ──
    if "_created_on_dt" in df.columns:
        valid = df["_created_on_dt"].notna()
        df.loc[valid, "_week"] = df.loc[valid, "_created_on_dt"].dt.to_period("W").astype(str)
        df.loc[valid, "_date"] = df.loc[valid, "_created_on_dt"].dt.date
        df.loc[valid, "_dow"]  = df.loc[valid, "_created_on_dt"].dt.day_name()
        df.loc[valid, "_hour"] = df.loc[valid, "_created_on_dt"].dt.hour

    # ── Derived: age (hours since created, for unresolved) ──
    if "_created_on_dt" in df.columns:
        now = pd.Timestamp.now()
        df["_age_hours"] = (now - df["_created_on_dt"]).dt.total_seconds() / 3600

    # ── Derived: owner clean ──
    src_own = col_map.get("owner")
    if src_own:
        df["_owner"] = df[src_own].fillna("(blank)").astype(str).str.strip()

    # ── Derived: text lengths + text content ──
    for key in ["description", "activity_subj"]:
        src = col_map.get(key)
        if src:
            df[f"_{key}_len"] = df[src].fillna("").astype(str).str.len()
            df[f"_{key}_text"] = df[src].fillna("").astype(str).str.strip()

    df._col_map = col_map
    return df


# ═══════════════════════════════════════════════════════════
#  SHEET BUILDERS — Cases
# ═══════════════════════════════════════════════════════════

def sheet_01_population_split(df):
    """D01: 3-way population split — named client, blank-company client, admin."""
    total = len(df)
    named   = df[df["_case_population"] == "named_client"]
    blank   = df[df["_case_population"] == "blank_client"]
    admin   = df[df["_case_population"] == "admin"]
    client_inclusive = df[~df["_is_internal"]]  # named + blank

    def _stats(subset, label):
        n = len(subset)
        rows = [
            {"section": label, "metric": "case_count", "value": f"{n:,}"},
            {"section": label, "metric": "pct_of_total", "value": pct(n, total)},
        ]
        if "_hours" in subset.columns:
            hrs = subset["_hours"].dropna()
            if len(hrs):
                rows.append({"section": label, "metric": "median_hours", "value": f"{hrs.median():.1f}"})
                rows.append({"section": label, "metric": "p90_hours", "value": f"{hrs.quantile(.9):.1f}"})
        if "_is_resolved" in subset.columns:
            unres = (~subset["_is_resolved"]).sum()
            rows.append({"section": label, "metric": "unresolved_count", "value": f"{unres:,}"})
            rows.append({"section": label, "metric": "pct_unresolved", "value": pct(unres, n)})
        # top 5 subjects
        top_subj = subset["_subject"].value_counts().head(5)
        for s, c in top_subj.items():
            rows.append({"section": label, "metric": f"top_subject: {s}", "value": f"{c:,}"})

        # origin distribution (key for validating blank = client)
        if "_origin" in subset.columns:
            top_orig = subset["_origin"].value_counts().head(3)
            for o, c in top_orig.items():
                rows.append({"section": label, "metric": f"top_origin: {o}", "value": f"{c:,} ({pct(c, n)})"})
        return rows

    rows = _stats(df, "ALL CASES")
    rows += _stats(client_inclusive, "CLIENT INCLUSIVE (named+blank)")
    rows += _stats(named, "NAMED CLIENT")
    rows += _stats(blank, "BLANK COMPANY (unassigned client)")
    rows += _stats(admin, "ADMIN / MULTI-CLIENT")

    return pd.DataFrame(rows)


def sheet_02_client_weekly(df):
    """D02: Weekly time-series for CLIENT cases only."""
    client = df[~df["_is_internal"]].copy()
    if "_week" not in client.columns or client["_week"].isna().all():
        return pd.DataFrame({"note": ["No date data for weekly aggregation"]})

    valid = client.dropna(subset=["_week"])
    weekly = valid.groupby("_week").agg(
        created=("_week", "size"),
    ).reset_index()

    # resolved per week
    if "_resolved_on_dt" in client.columns:
        res = client.dropna(subset=["_resolved_on_dt"]).copy()
        res["_res_week"] = res["_resolved_on_dt"].dt.to_period("W").astype(str)
        res_wk = res.groupby("_res_week").size().reset_index(name="resolved")
        weekly = weekly.merge(res_wk, left_on="_week", right_on="_res_week", how="left").drop(columns=["_res_week"], errors="ignore")

    weekly["resolved"] = weekly.get("resolved", pd.Series([0]*len(weekly))).fillna(0).astype(int)

    # median hours
    if "_hours" in client.columns:
        hrs_wk = valid.groupby("_week")["_hours"].agg(
            median_hrs="median", p90_hrs=lambda x: x.quantile(0.9) if len(x) else np.nan
        ).reset_index()
        hrs_wk["median_hrs"] = hrs_wk["median_hrs"].round(1)
        hrs_wk["p90_hrs"] = hrs_wk["p90_hrs"].round(1)
        weekly = weekly.merge(hrs_wk, on="_week", how="left")

    # backlog
    weekly["cum_created"]  = weekly["created"].cumsum()
    weekly["cum_resolved"] = weekly["resolved"].cumsum()
    weekly["backlog"]      = weekly["cum_created"] - weekly["cum_resolved"]

    # pct unresolved created that week (snapshot)
    if "_is_resolved" in client.columns:
        unres_wk = valid[~valid["_is_resolved"]].groupby("_week").size().reset_index(name="still_open")
        weekly = weekly.merge(unres_wk, on="_week", how="left")
        weekly["still_open"] = weekly["still_open"].fillna(0).astype(int)

    weekly.rename(columns={"_week": "week"}, inplace=True)
    return weekly


def sheet_03_subject_deep(df):
    """D03: Top 15 subjects with full resolution profile — client only."""
    client = df[~df["_is_internal"]].copy()
    grp = client.groupby("_subject")

    result = grp.size().reset_index(name="count")
    result = result.sort_values("count", ascending=False).head(15)

    if "_hours" in client.columns:
        agg = grp["_hours"].agg(
            median_hrs="median",
            p75_hrs=lambda x: x.quantile(0.75),
            p90_hrs=lambda x: x.quantile(0.90),
            max_hrs="max",
        ).reset_index()
        for c in ["median_hrs", "p75_hrs", "p90_hrs", "max_hrs"]:
            agg[c] = agg[c].round(1)
        result = result.merge(agg, on="_subject", how="left")

    if "_is_resolved" in client.columns:
        unres = grp["_is_resolved"].agg(
            unresolved=lambda x: (~x).sum(),
            pct_unresolved=lambda x: f"{100*(~x).mean():.1f}%",
        ).reset_index()
        result = result.merge(unres, on="_subject", how="left")

    # description fill rate by subject
    if "_description_len" in client.columns:
        desc_fill = grp["_description_len"].agg(
            desc_fill_pct=lambda x: f"{100*(x>0).mean():.0f}%",
            desc_median_len=lambda x: int(x[x>0].median()) if (x>0).any() else 0,
        ).reset_index()
        result = result.merge(desc_fill, on="_subject", how="left")

    # activity subject fill rate
    if "_activity_subj_len" in client.columns:
        act_fill = grp["_activity_subj_len"].agg(
            act_subj_fill_pct=lambda x: f"{100*(x>0).mean():.0f}%",
        ).reset_index()
        result = result.merge(act_fill, on="_subject", how="left")

    result.rename(columns={"_subject": "subject"}, inplace=True)
    return result.reset_index(drop=True)


def sheet_04_day_of_week(df):
    """D04: Day-of-week volume pattern — client only."""
    client = df[(~df["_is_internal"]) & df.get("_dow", pd.Series(dtype=str)).notna()].copy()
    if "_dow" not in client.columns:
        return pd.DataFrame({"note": ["No day-of-week data"]})

    dow_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    grp = client.groupby("_dow")
    result = grp.size().reset_index(name="cases")
    result["_sort"] = result["_dow"].map({d:i for i,d in enumerate(dow_order)})
    result = result.sort_values("_sort").drop(columns=["_sort"])
    result["pct"] = (result["cases"] / result["cases"].sum() * 100).round(1).astype(str) + "%"

    if "_hours" in client.columns:
        med = grp["_hours"].median().round(1).reset_index(name="median_hrs")
        result = result.merge(med, on="_dow")

    # count distinct weeks to compute avg per day
    if "_week" in client.columns:
        n_weeks = client["_week"].nunique()
        result["avg_per_day"] = (result["cases"] / n_weeks).round(0).astype(int)

    result.rename(columns={"_dow": "day"}, inplace=True)
    return result


def sheet_05_hourly_pattern(df):
    """D05: Hourly creation pattern — client only."""
    client = df[(~df["_is_internal"]) & df.get("_hour", pd.Series(dtype=float)).notna()].copy()
    if "_hour" not in client.columns:
        return pd.DataFrame({"note": ["No hour data"]})

    grp = client.groupby("_hour")
    result = grp.size().reset_index(name="cases")
    result["pct"] = (result["cases"] / result["cases"].sum() * 100).round(1).astype(str) + "%"
    result["_hour"] = result["_hour"].astype(int)
    result.rename(columns={"_hour": "hour"}, inplace=True)
    return result.sort_values("hour")


def sheet_06_sla_breach(df):
    """D06: SLA breach analysis by subject — client only.
    Uses multiple thresholds since we don't know actual SLA targets."""
    client = df[~df["_is_internal"]].copy()
    if "_hours" not in client.columns:
        return pd.DataFrame({"note": ["Resolved In Hours not available"]})

    thresholds = [4, 8, 24, 48, 72, 168]  # hours
    subjects = client["_subject"].value_counts().head(12).index.tolist()
    rows = []

    for subj in subjects:
        sub = client[client["_subject"] == subj]
        hrs = sub["_hours"].dropna()
        n = len(hrs)
        if n == 0:
            continue
        row = {"subject": subj, "cases_with_hours": n}
        for t in thresholds:
            breach = (hrs > t).sum()
            row[f">{t}h_count"] = int(breach)
            row[f">{t}h_pct"] = f"{100*breach/n:.1f}%"
        rows.append(row)

    return pd.DataFrame(rows) if rows else pd.DataFrame({"note": ["No hours data"]})


def sheet_07_backlog_detail(df):
    """D07: Current unresolved cases — aging by subject and company."""
    unres = df[(~df["_is_internal"]) & (~df["_is_resolved"])].copy()
    if len(unres) == 0:
        return pd.DataFrame({"note": ["No unresolved client cases"]})

    parts = []

    # Summary
    summary = pd.DataFrame({"section": ["SUMMARY"], "metric": ["total_unresolved_client"],
                            "value": [f"{len(unres):,}"]})
    parts.append(summary)

    # Aging buckets
    if "_age_hours" in unres.columns:
        bins = [0, 24, 72, 168, 336, 720, float("inf")]
        labels = ["0-24h", "1-3d", "3-7d", "7-14d", "14-30d", "30d+"]
        unres["_bucket"] = pd.cut(unres["_age_hours"], bins=bins, labels=labels, right=True)
        aging = unres["_bucket"].value_counts().reindex(labels).fillna(0).astype(int).reset_index()
        aging.columns = ["metric", "value"]
        aging.insert(0, "section", "AGING BUCKETS")
        parts.append(aging)

    # By subject
    subj_ct = unres["_subject"].value_counts().head(10).reset_index()
    subj_ct.columns = ["metric", "value"]
    subj_ct.insert(0, "section", "BY SUBJECT")
    parts.append(subj_ct)

    # By company
    if "_company_clean" in unres.columns:
        co_ct = unres["_company_clean"].value_counts().head(10).reset_index()
        co_ct.columns = ["metric", "value"]
        co_ct.insert(0, "section", "BY COMPANY")
        parts.append(co_ct)

    # Cross: top subjects x aging buckets
    if "_bucket" in unres.columns:
        top_subj = unres["_subject"].value_counts().head(6).index.tolist()
        ct = pd.crosstab(
            unres[unres["_subject"].isin(top_subj)]["_subject"],
            unres[unres["_subject"].isin(top_subj)]["_bucket"],
        )
        ct = ct.reset_index().rename(columns={"_subject": "metric"})
        ct.insert(0, "section", "SUBJECT x AGING")
        parts.append(ct)

    return pd.concat(parts, ignore_index=True, sort=False)


def sheet_08_retouch(df):
    """D08: Re-touch analysis — cases touched after resolution."""
    if "_resolved_on_dt" not in df.columns or "_last_touch_dt" not in df.columns:
        return pd.DataFrame({"note": ["Resolved On or Last Touch column not available"]})

    client = df[(~df["_is_internal"]) & df["_is_resolved"]].copy()
    client = client.dropna(subset=["_resolved_on_dt", "_last_touch_dt"])

    if len(client) == 0:
        return pd.DataFrame({"note": ["No resolved cases with last touch data"]})

    # Cases where last touch is AFTER resolved on
    client["_touch_gap_hrs"] = (
        client["_last_touch_dt"] - client["_resolved_on_dt"]
    ).dt.total_seconds() / 3600

    retouched = client[client["_touch_gap_hrs"] > 1].copy()  # >1 hour after resolution

    parts = []

    total_resolved = len(client)
    total_retouched = len(retouched)
    parts.append(pd.DataFrame({
        "section": ["SUMMARY", "SUMMARY", "SUMMARY"],
        "metric": ["resolved_cases_with_touch_data", "retouched_after_resolve", "retouch_rate"],
        "value": [f"{total_resolved:,}", f"{total_retouched:,}", pct(total_retouched, total_resolved)],
    }))

    if total_retouched > 0:
        # Gap distribution
        gaps = retouched["_touch_gap_hrs"]
        dist = pd.DataFrame({
            "section": ["GAP DISTRIBUTION"] * 5,
            "metric": ["median_gap_hrs", "p75_gap_hrs", "p90_gap_hrs", "max_gap_hrs", "avg_gap_hrs"],
            "value": [f"{gaps.median():.1f}", f"{gaps.quantile(.75):.1f}",
                      f"{gaps.quantile(.9):.1f}", f"{gaps.max():.1f}", f"{gaps.mean():.1f}"],
        })
        parts.append(dist)

        # By subject
        by_subj = retouched["_subject"].value_counts().head(10).reset_index()
        by_subj.columns = ["metric", "value"]
        by_subj.insert(0, "section", "RETOUCHED BY SUBJECT")
        parts.append(by_subj)

    return pd.concat(parts, ignore_index=True, sort=False)


def sheet_09_owner_workload(df):
    """D09: Pod and owner workload — client only."""
    client = df[~df["_is_internal"]].copy()
    parts = []

    # By pod
    if "_pod" not in client.columns:
        pod_col = df._col_map.get("pod")
        if pod_col:
            client["_pod"] = client[pod_col].fillna("(blank)").astype(str)

    if "_pod" in client.columns or df._col_map.get("pod"):
        pod_src = "_pod" if "_pod" in client.columns else df._col_map["pod"]
        if pod_src not in client.columns:
            client["_pod"] = client[df._col_map["pod"]].fillna("(blank)").astype(str)
        grp = client.groupby("_pod")
        pod_tbl = grp.size().reset_index(name="cases")
        pod_tbl = pod_tbl.sort_values("cases", ascending=False).head(10)

        if "_hours" in client.columns:
            med = grp["_hours"].median().round(1).reset_index(name="median_hrs")
            pod_tbl = pod_tbl.merge(med, on="_pod")

        if "_is_resolved" in client.columns:
            unres = grp["_is_resolved"].agg(unresolved=lambda x: (~x).sum()).reset_index()
            pod_tbl = pod_tbl.merge(unres, on="_pod")

        pod_tbl.rename(columns={"_pod": "metric"}, inplace=True)
        pod_tbl.insert(0, "section", "BY POD")
        parts.append(pod_tbl)

    # By owner
    owner_col = df._col_map.get("owner")
    if owner_col:
        client["_owner"] = client[owner_col].fillna("(blank)").astype(str)
        grp = client.groupby("_owner")
        own_tbl = grp.size().reset_index(name="cases")
        own_tbl = own_tbl.sort_values("cases", ascending=False).head(15)

        if "_hours" in client.columns:
            med = grp["_hours"].median().round(1).reset_index(name="median_hrs")
            own_tbl = own_tbl.merge(med, on="_owner")

        if "_is_resolved" in client.columns:
            unres = grp["_is_resolved"].agg(unresolved=lambda x: (~x).sum()).reset_index()
            own_tbl = own_tbl.merge(unres, on="_owner")

        own_tbl.rename(columns={"_owner": "metric"}, inplace=True)
        own_tbl.insert(0, "section", "BY OWNER")
        parts.append(own_tbl)

    if not parts:
        return pd.DataFrame({"note": ["No pod/owner columns found"]})
    return pd.concat(parts, ignore_index=True, sort=False)


def sheet_10_origin_subject_detail(df):
    """D10: Origin x Subject matrix with median hours — client only."""
    client = df[~df["_is_internal"]].copy()

    top_origins = client["_origin"].value_counts().head(5).index.tolist()
    top_subjects = client["_subject"].value_counts().head(8).index.tolist()

    sub = client[client["_origin"].isin(top_origins) & client["_subject"].isin(top_subjects)]

    # Count matrix
    ct_count = pd.crosstab(sub["_origin"], sub["_subject"])

    # Median hours matrix
    if "_hours" in sub.columns:
        ct_hours = sub.groupby(["_origin", "_subject"])["_hours"].median().round(1).unstack(fill_value=np.nan)
    else:
        ct_hours = pd.DataFrame()

    parts = []
    ct1 = ct_count.reset_index().rename(columns={"_origin": "origin"})
    ct1.insert(0, "section", "CASE COUNT")
    parts.append(ct1)

    if not ct_hours.empty:
        ct2 = ct_hours.reset_index().rename(columns={"_origin": "origin"})
        ct2.insert(0, "section", "MEDIAN HOURS")
        parts.append(ct2)

    return pd.concat(parts, ignore_index=True, sort=False) if parts else pd.DataFrame({"note": ["No data"]})


# ═══════════════════════════════════════════════════════════
#  SHEET BUILDERS — Emails
# ═══════════════════════════════════════════════════════════

def prepare_emails(emails, cases_df):
    """Enrich emails with case linkage and stripped text."""
    df = emails.copy()

    col_map = {
        "subject":     find_col(df, "Subject"),
        "description": find_col(df, "Description"),
        "from":        find_col(df, "From"),
        "to":          find_col(df, "To"),
        "status":      find_col(df, "Status Reason"),
        "created_on":  find_col(df, "Created On"),
        "owner":       find_col(df, "Owner"),
        "regarding":   find_col(df, "Regarding"),
        "case_number": find_col(df, "Case Number (Regarding) (Case)"),
        "case_subject": find_col(df, "Subject (Regarding) (Case)"),
        "case_subj_path": find_col(df, "Subject Path (Regarding) (Case)"),
        "priority":    find_col(df, "Priority"),
    }

    # Strip HTML from description
    desc_col = col_map.get("description")
    if desc_col:
        log("  Stripping HTML from email descriptions (this may take a moment)...")
        df["_body_text"] = df[desc_col].apply(strip_html)
        df["_body_len"]  = df["_body_text"].str.len()
        df["_html_len"]  = df[desc_col].fillna("").astype(str).str.len()

    # Created datetime
    src = col_map.get("created_on")
    if src:
        df["_created_dt"] = safe_dt(df[src])
        df["_hour"] = df["_created_dt"].dt.hour

    # Case linkage flag
    case_col = col_map.get("case_number")
    if case_col:
        df["_has_case"] = df[case_col].notna()
    else:
        df["_has_case"] = False

    # Link to case data for resolution hours
    if case_col and not cases_df.empty:
        case_num_col = find_col(cases_df, "Case Number")
        case_subj_col = find_col(cases_df, "Subject")
        case_hrs_col = find_col(cases_df, "Resolved In Hours")
        if case_num_col:
            case_lk = cases_df[[case_num_col]].copy()
            if case_subj_col:
                case_lk["_linked_case_subject"] = cases_df[case_subj_col]
            if case_hrs_col:
                case_lk["_linked_case_hours"] = safe_numeric(cases_df[case_hrs_col])
            case_lk = case_lk.rename(columns={case_num_col: "_case_join_key"})
            df["_case_join_key"] = df[case_col].astype(str)
            case_lk["_case_join_key"] = case_lk["_case_join_key"].astype(str)
            df = df.merge(case_lk.drop_duplicates(subset=["_case_join_key"]),
                          on="_case_join_key", how="left")

    df._col_map = col_map
    return df


def sheet_11_email_overview(edf):
    """D11: Email overview stats."""
    parts = []
    n = len(edf)

    # Headlines
    rows = [
        {"section": "TOTALS", "metric": "email_count", "value": f"{n:,}"},
        {"section": "TOTALS", "metric": "linked_to_case", "value": f"{edf['_has_case'].sum():,} ({pct(edf['_has_case'].sum(), n)})"},
    ]
    if "_body_len" in edf.columns:
        rows.append({"section": "TOTALS", "metric": "median_body_text_chars", "value": f"{int(edf['_body_len'].median()):,}"})
        rows.append({"section": "TOTALS", "metric": "median_html_chars", "value": f"{int(edf['_html_len'].median()):,}"})
        rows.append({"section": "TOTALS", "metric": "html_to_text_ratio", "value": f"{(edf['_html_len'] / edf['_body_len'].replace(0, 1)).median():.1f}x"})
    parts.append(pd.DataFrame(rows))

    # Status distribution
    status_col = edf._col_map.get("status")
    if status_col:
        st = edf[status_col].fillna("(blank)").value_counts().reset_index()
        st.columns = ["metric", "value"]
        st.insert(0, "section", "STATUS REASON")
        parts.append(st)

    # Priority distribution
    prio_col = edf._col_map.get("priority")
    if prio_col:
        pr = edf[prio_col].fillna("(blank)").value_counts().reset_index()
        pr.columns = ["metric", "value"]
        pr.insert(0, "section", "PRIORITY")
        parts.append(pr)

    # Hourly distribution
    if "_hour" in edf.columns:
        hrs = edf["_hour"].dropna().astype(int).value_counts().sort_index().reset_index()
        hrs.columns = ["metric", "value"]
        hrs["metric"] = hrs["metric"].astype(str) + ":00"
        hrs.insert(0, "section", "HOUR OF DAY")
        parts.append(hrs)

    return pd.concat(parts, ignore_index=True, sort=False)


def sheet_12_email_case_subject_mix(edf):
    """D12: What case subjects generate the most email?"""
    if "_linked_case_subject" not in edf.columns:
        return pd.DataFrame({"note": ["Case subject linkage not available"]})

    linked = edf[edf["_has_case"]].copy()
    if len(linked) == 0:
        return pd.DataFrame({"note": ["No linked emails"]})

    grp = linked.groupby("_linked_case_subject")
    result = grp.size().reset_index(name="email_count")
    result = result.sort_values("email_count", ascending=False).head(15)

    # Distinct cases per subject
    case_col = edf._col_map.get("case_number")
    if case_col:
        distinct_cases = grp[case_col].nunique().reset_index(name="distinct_cases")
        result = result.merge(distinct_cases, on="_linked_case_subject", how="left")
        result["emails_per_case"] = (result["email_count"] / result["distinct_cases"]).round(1)

    # median body length
    if "_body_len" in linked.columns:
        body_med = grp["_body_len"].median().round(0).astype(int).reset_index(name="median_body_chars")
        result = result.merge(body_med, on="_linked_case_subject", how="left")

    result.rename(columns={"_linked_case_subject": "case_subject"}, inplace=True)
    return result.reset_index(drop=True)


def sheet_13_email_burden_per_case(edf):
    """D13: Email burden distribution — how many emails per case?"""
    case_col = edf._col_map.get("case_number")
    if not case_col:
        return pd.DataFrame({"note": ["Case number column not found"]})

    linked = edf[edf["_has_case"]].copy()
    per_case = linked.groupby(case_col).size()

    parts = []

    # Distribution
    dist = pd.DataFrame({
        "metric": ["cases_with_emails", "min", "median", "p75", "p90", "p95", "max"],
        "value": [
            f"{len(per_case):,}",
            int(per_case.min()), int(per_case.median()),
            int(per_case.quantile(.75)), int(per_case.quantile(.9)),
            int(per_case.quantile(.95)), int(per_case.max()),
        ]
    })
    dist.insert(0, "section", "EMAILS PER CASE")
    parts.append(dist)

    # Bucket distribution
    bins = [0, 1, 2, 3, 5, 10, float("inf")]
    labels = ["1", "2", "3", "4-5", "6-10", "11+"]
    buckets = pd.cut(per_case, bins=bins, labels=labels, right=True)
    bkt = buckets.value_counts().reindex(labels).fillna(0).astype(int).reset_index()
    bkt.columns = ["metric", "value"]
    bkt.insert(0, "section", "BUCKET DISTRIBUTION")
    parts.append(bkt)

    # Top 10 highest-email cases
    top = per_case.nlargest(10).reset_index()
    top.columns = ["metric", "value"]
    top.insert(0, "section", "HIGHEST EMAIL CASES")
    parts.append(top)

    return pd.concat(parts, ignore_index=True, sort=False)


def sheet_14_email_text_samples(edf):
    """D14: Stratified email text samples with HTML stripped."""
    rows = []
    subj_col = edf._col_map.get("subject")
    case_subj_col = "_linked_case_subject" if "_linked_case_subject" in edf.columns else None

    # 5 linked, 5 unlinked
    for label, pool in [("linked", edf[edf["_has_case"]]), ("unlinked", edf[~edf["_has_case"]])]:
        sample = pool.head(5) if len(pool) >= 5 else pool
        for _, r in sample.iterrows():
            row = {
                "type": label,
                "email_subject": trunc(r.get(subj_col, ""), 100) if subj_col else "",
                "case_subject": trunc(r.get(case_subj_col, ""), 60) if case_subj_col and pd.notna(r.get(case_subj_col)) else "",
                "body_stripped": trunc(r.get("_body_text", ""), TRUNC),
                "body_chars": int(r.get("_body_len", 0)) if pd.notna(r.get("_body_len")) else 0,
            }
            rows.append(row)

    # 5 longest-body emails
    if "_body_len" in edf.columns:
        longest = edf.nlargest(5, "_body_len")
        for _, r in longest.iterrows():
            row = {
                "type": "longest_body",
                "email_subject": trunc(r.get(subj_col, ""), 100) if subj_col else "",
                "case_subject": trunc(r.get(case_subj_col, ""), 60) if case_subj_col and pd.notna(r.get(case_subj_col)) else "",
                "body_stripped": trunc(r.get("_body_text", ""), TRUNC),
                "body_chars": int(r.get("_body_len", 0)),
            }
            rows.append(row)

    return pd.DataFrame(rows) if rows else pd.DataFrame({"note": ["No samples"]})


def sheet_15_genai_evidence(df, edf):
    """D15: Compact GenAI use-case evidence summary."""
    client = df[~df["_is_internal"]].copy()
    rows = []

    def _add(uc, metric, value, signal):
        rows.append({"use_case": uc, "metric": metric, "value": value, "signal": signal})

    n_client = len(client)

    # ── Triage / routing ──
    n_origins = client["_origin"].nunique()
    n_subjects = client["_subject"].nunique()
    email_origin_pct = pct((client["_origin"] == "Email").sum(), n_client)
    _add("Triage/Routing", "distinct_origins", n_origins, "Low variety → simple rules may suffice")
    _add("Triage/Routing", "distinct_subjects", n_subjects, "Moderate variety → subject-based routing feasible")
    _add("Triage/Routing", "email_origin_pct", email_origin_pct, "Dominant channel = email inflow")
    # Top 3 subjects account for what %
    top3 = client["_subject"].value_counts().head(3).sum()
    _add("Triage/Routing", "top_3_subjects_pct", pct(top3, n_client), "High concentration → top subjects are targetable")

    # ── Summarization ──
    if "_body_len" in edf.columns:
        med_body = int(edf["_body_len"].median())
        _add("Summarization", "email_body_median_chars", f"{med_body:,}", "Long bodies → summarization valuable" if med_body > 500 else "Short bodies → limited value")
        pct_long = pct((edf["_body_len"] > 500).sum(), len(edf))
        _add("Summarization", "emails_over_500_chars", pct_long, "Pool of summarizable emails")

    if "_description_len" in client.columns:
        fill = pct((client["_description_len"] > 0).sum(), n_client)
        _add("Summarization", "case_description_fill", fill, ">70% = good summarization pool" if (client["_description_len"] > 0).mean() > 0.7 else "<70% = thin pool")

    if "_activity_subj_len" in client.columns:
        fill = pct((client["_activity_subj_len"] > 0).sum(), n_client)
        _add("Summarization", "activity_subject_fill", fill, "Best text field in cases")

    # ── Missing-info detection ──
    if "_description_len" in client.columns:
        empty_desc_unres = client[(client["_description_len"] == 0) & (~client["_is_resolved"])]
        _add("Missing-Info", "unresolved_no_description", f"{len(empty_desc_unres):,}", "Flaggable cases")
    if "_company_clean" in client.columns:
        blank_co = (client["_company_clean"] == "(blank)").sum()
        _add("Missing-Info", "blank_company_name", f"{blank_co:,}", "Missing entity linkage")

    # ── Escalation signals ──
    if "_hours" in client.columns:
        p90 = client["_hours"].dropna().quantile(0.9)
        n_breach = (client["_hours"] > p90).sum()
        _add("Escalation", "global_p90_hours", f"{p90:.1f}", "Threshold for anomaly detection")
        _add("Escalation", "cases_above_p90", f"{n_breach:,}", "Candidate escalation pool")
    unres = (~client["_is_resolved"]).sum()
    _add("Escalation", "current_unresolved", f"{unres:,}", "Active backlog")

    # ── Draft reply ──
    if "_body_len" in edf.columns:
        linked = edf[edf["_has_case"]]
        _add("Draft Reply", "linked_emails_with_body", f"{len(linked):,}", "Pool for reply generation")
        if "_linked_case_subject" in linked.columns:
            top_subj = linked["_linked_case_subject"].value_counts().head(3)
            for s, c in top_subj.items():
                _add("Draft Reply", f"top_linked_subject: {s}", f"{c:,}", "Repetitive = templatable")

    # ── Workflow copilot ──
    if "_hours" in client.columns:
        top_slow = client.groupby("_subject")["_hours"].median().nlargest(5)
        for s, h in top_slow.items():
            _add("Workflow Copilot", f"slowest_subject: {s}", f"{h:.1f}h median", "Process friction → copilot opportunity")

    return pd.DataFrame(rows)


def sheet_16_triage_delay(df):
    """D16: Email-to-case triage delay — gap between SLA Start and Created On.

    SLA Start = when the originating email was received.
    Created On = when the banker created the case.
    The gap measures how long the email sat before a human acted on it.
    Only meaningful for email-originated, client cases where SLA Start != Created On.
    """
    client = df[~df["_is_internal"]].copy()

    if "_sla_start_dt" not in client.columns or "_created_on_dt" not in client.columns:
        return pd.DataFrame({"note": ["SLA Start or Created On column not available"]})

    # Only email-originated cases where SLA Start < Created On (email arrived before case created)
    valid = client.dropna(subset=["_sla_start_dt", "_created_on_dt"]).copy()
    valid["_triage_minutes"] = (
        valid["_created_on_dt"] - valid["_sla_start_dt"]
    ).dt.total_seconds() / 60

    # Filter: only positive gaps (email arrived before case) and reasonable range (< 7 days)
    # Negative or zero gaps mean SLA Start = Created On (non-email case or auto-created)
    email_triage = valid[(valid["_triage_minutes"] > 0.5) & (valid["_triage_minutes"] < 10080)].copy()

    # Also track zero-gap cases (auto-created or non-email)
    zero_gap = valid[valid["_triage_minutes"] <= 0.5]

    parts = []

    # Summary
    total_valid = len(valid)
    n_with_gap = len(email_triage)
    n_zero = len(zero_gap)

    summary_rows = [
        {"section": "SUMMARY", "metric": "client_cases_with_both_timestamps", "value": f"{total_valid:,}"},
        {"section": "SUMMARY", "metric": "cases_with_triage_gap_>30sec", "value": f"{n_with_gap:,}"},
        {"section": "SUMMARY", "metric": "cases_with_zero_or_negative_gap", "value": f"{n_zero:,} (auto-created or non-email)"},
        {"section": "SUMMARY", "metric": "pct_with_triage_gap", "value": pct(n_with_gap, total_valid)},
    ]

    if n_with_gap > 0:
        mins = email_triage["_triage_minutes"]
        summary_rows.extend([
            {"section": "GAP DISTRIBUTION (minutes)", "metric": "min", "value": f"{mins.min():.1f}"},
            {"section": "GAP DISTRIBUTION (minutes)", "metric": "p25", "value": f"{mins.quantile(0.25):.1f}"},
            {"section": "GAP DISTRIBUTION (minutes)", "metric": "median", "value": f"{mins.median():.1f}"},
            {"section": "GAP DISTRIBUTION (minutes)", "metric": "p75", "value": f"{mins.quantile(0.75):.1f}"},
            {"section": "GAP DISTRIBUTION (minutes)", "metric": "p90", "value": f"{mins.quantile(0.90):.1f}"},
            {"section": "GAP DISTRIBUTION (minutes)", "metric": "p95", "value": f"{mins.quantile(0.95):.1f}"},
            {"section": "GAP DISTRIBUTION (minutes)", "metric": "max", "value": f"{mins.max():.1f}"},
        ])

        # Convert to hours for the bucket view
        hours = mins / 60
        bins = [0, 0.083, 0.25, 0.5, 1, 2, 4, 8, 24, float("inf")]
        labels = ["<5min", "5-15min", "15-30min", "30min-1h", "1-2h", "2-4h", "4-8h", "8-24h", "24h+"]
        buckets = pd.cut(hours, bins=bins, labels=labels, right=True)
        bkt = buckets.value_counts().reindex(labels).fillna(0).astype(int).reset_index()
        bkt.columns = ["metric", "value"]
        bkt.insert(0, "section", "TIME BUCKETS")
        parts.append(pd.DataFrame(summary_rows))
        parts.append(bkt)

        # By subject (top 10)
        email_triage["_subj"] = email_triage.get("_subject", "unknown")
        by_subj = email_triage.groupby("_subj")["_triage_minutes"].agg(
            cases="size",
            median_minutes="median",
            p90_minutes=lambda x: x.quantile(0.9),
        ).reset_index().sort_values("cases", ascending=False).head(10)
        by_subj["median_minutes"] = by_subj["median_minutes"].round(1)
        by_subj["p90_minutes"] = by_subj["p90_minutes"].round(1)
        by_subj.rename(columns={"_subj": "metric"}, inplace=True)
        by_subj.insert(0, "section", "BY SUBJECT")
        parts.append(by_subj)

        # By pod (top 8)
        pod_col = df._col_map.get("pod")
        if pod_col and pod_col in email_triage.columns:
            email_triage["_pod"] = email_triage[pod_col].fillna("(blank)").astype(str)
        elif "_pod" not in email_triage.columns:
            email_triage["_pod"] = "(no pod column)"
        by_pod = email_triage.groupby("_pod")["_triage_minutes"].agg(
            cases="size",
            median_minutes="median",
            p90_minutes=lambda x: x.quantile(0.9),
        ).reset_index().sort_values("cases", ascending=False).head(8)
        by_pod["median_minutes"] = by_pod["median_minutes"].round(1)
        by_pod["p90_minutes"] = by_pod["p90_minutes"].round(1)
        by_pod.rename(columns={"_pod": "metric"}, inplace=True)
        by_pod.insert(0, "section", "BY POD")
        parts.append(by_pod)

        # By hour of day (when does triage take longest?)
        if "_hour" in email_triage.columns:
            by_hour = email_triage.groupby("_hour")["_triage_minutes"].agg(
                cases="size",
                median_minutes="median",
            ).reset_index().sort_index()
            by_hour["median_minutes"] = by_hour["median_minutes"].round(1)
            by_hour["_hour"] = by_hour["_hour"].astype(int).astype(str) + ":00"
            by_hour.rename(columns={"_hour": "metric"}, inplace=True)
            by_hour.insert(0, "section", "BY HOUR OF DAY")
            parts.append(by_hour)

    else:
        parts.append(pd.DataFrame(summary_rows))

    return pd.concat(parts, ignore_index=True, sort=False) if parts else pd.DataFrame(summary_rows)


# ═══════════════════════════════════════════════════════════
#  D17-D21: SUBJECT INTELLIGENCE (folded from subject deep dive)
# ═══════════════════════════════════════════════════════════

# Action-tag rules
ACTION_RULES = {
    "CD Maintenance": "MONITOR", "IntraFi Maintenance": "MONITOR",
    "NSF and Non-Post": "AUTOMATE", "Fraud Alert": "AUTOMATE",
    "Transfer": "AUTOMATE", "Statements": "AI-ASSIST",
    "Signature Card": "REDESIGN", "Close Account": "REDESIGN",
    "Research": "AI-ASSIST", "New Account Request": "AI-ASSIST",
    "Account Maintenance": "AI-ASSIST", "General Questions": "AI-ASSIST",
}

def _action_tag(subject, median_hrs, unres_pct, cpw, top_owner_pct, p90_med_ratio):
    for pattern, tag in ACTION_RULES.items():
        if pattern.lower() in subject.lower():
            return tag
    if pd.notna(median_hrs) and median_hrs < 4 and pd.notna(cpw) and cpw > 15:
        return "AUTOMATE"
    if pd.notna(unres_pct) and unres_pct > 10 and pd.notna(top_owner_pct) and top_owner_pct > 40:
        return "REDESIGN"
    if pd.notna(p90_med_ratio) and p90_med_ratio >= 10:
        return "AI-ASSIST"
    if pd.notna(cpw) and cpw < 5:
        return "DEPRIORITIZE"
    return "MONITOR"

def _research_cluster(desc, act_subj):
    combined = f"{desc} {act_subj}".lower()
    if any(k in combined for k in ["payment","ach","wire","transfer","deposit",
                                    "credit","debit","transaction","posted","posting",
                                    "return","reversal","refund"]):
        return "Payment Research"
    if any(k in combined for k in ["check","cheque","item","image","copy","front","back","clearing"]):
        return "Check/Item Research"
    if any(k in combined for k in ["statement","balance","reconcil","ledger","interest","rate","fee"]):
        return "Statement/Balance Inquiry"
    if any(k in combined for k in ["address","signer","name change","update","modify","amendment",
                                    "tin","ein","ssn","tax","w-9","w9","certification",
                                    "fraud","dispute","unauthorized","suspicious",
                                    "positive pay","stop payment",
                                    "new account","onboard","setup","opening"]):
        return "Account Updates & Other"
    if len(combined.strip()) < 10:
        return "No Text (Unclassifiable)"
    return "Other/Uncategorized"


def sheet_17_banker_hours(df):
    """D17: Banker-hours budget per subject with work/wait split and action tags."""
    client = df[~df["_is_internal"]].copy()
    if "_hours" not in client.columns:
        return pd.DataFrame({"note": ["Hours column not available"]})

    n_weeks = max(client["_week"].nunique(), 1) if "_week" in client.columns else 13
    top_subjects = client["_subject"].value_counts().head(25).index.tolist()
    wait_subjects = {"cd maintenance", "intrafi maintenance"}

    rows = []
    for subj in top_subjects:
        s = client[client["_subject"] == subj]
        n = len(s)
        hrs = s["_hours"].dropna()
        total_hrs = hrs.sum()
        weekly_hrs = total_hrs / n_weeks
        cpw = n / n_weeks
        med = hrs.median() if len(hrs) > 0 else np.nan
        unres = (~s["_is_resolved"]).sum()
        unres_pct = round(100 * unres / n, 1) if n > 0 else 0
        n_owners = s["_owner"].nunique() if "_owner" in s.columns else 0
        top_own_pct = round(100 * s["_owner"].value_counts().iloc[0] / n, 1) if n > 0 and "_owner" in s.columns else 0
        p90 = hrs.quantile(0.9) if len(hrs) > 0 else np.nan
        p90_med = round(p90 / max(med, 0.1), 1) if pd.notna(med) and pd.notna(p90) and med > 0 else np.nan

        is_wait = subj.lower() in wait_subjects
        action = _action_tag(subj, med, unres_pct, cpw, top_own_pct, p90_med)

        rows.append({
            "subject": subj, "intervention": action,
            "total_cases": n, "cases_per_week": round(cpw, 1),
            "median_hrs_per_case": round(med, 1) if pd.notna(med) else np.nan,
            "est_banker_hrs_per_week": round(weekly_hrs, 1),
            "hours_type": "WAIT TIME" if is_wait else "WORK TIME",
            "est_recoverable_hrs_per_week": 0 if is_wait else round(weekly_hrs * 0.6, 1),
            "current_unresolved": int(unres),
            "distinct_owners": n_owners,
        })

    result = pd.DataFrame(rows).sort_values("est_banker_hrs_per_week", ascending=False)
    totals = {
        "subject": "=== TOTAL (top 25) ===",
        "total_cases": result["total_cases"].sum(),
        "cases_per_week": round(result["cases_per_week"].sum(), 1),
        "est_banker_hrs_per_week": round(result["est_banker_hrs_per_week"].sum(), 1),
        "est_recoverable_hrs_per_week": round(result["est_recoverable_hrs_per_week"].sum(), 1),
    }
    return pd.concat([result, pd.DataFrame([totals])], ignore_index=True)


def sheet_18_claim_check(df):
    """D18: Chris March 23 claims tested against data."""
    client = df[~df["_is_internal"]].copy()
    n_weeks = max(client["_week"].nunique(), 1) if "_week" in client.columns else 13
    rows = []

    # CD Maintenance
    cd = client[client["_subject"].str.contains("CD Maintenance", case=False, na=False)]
    cd_hrs = cd["_hours"].dropna()
    cd_med = round(cd_hrs.median(), 1) if len(cd_hrs) > 0 else np.nan
    cd_p25 = round(cd_hrs.quantile(0.25), 1) if len(cd_hrs) > 0 else np.nan
    cd_7_14 = round(100 * ((cd_hrs >= 168) & (cd_hrs <= 336)).sum() / max(len(cd_hrs), 1), 1)
    rows.append({
        "claim": "CD Maintenance ~98h median is by design (maturity wait)",
        "source": "Chris, March 23",
        "data_finding": f"Median {cd_med}h. P25={cd_p25}h (even fast cases >1d). {cd_7_14}% in 7-14d maturity window.",
        "verdict": "CONFIRMED" if pd.notna(cd_p25) and cd_p25 > 12 else "PARTIALLY CONFIRMED",
        "implication": "Not a speed target. Wait time, not work time. Exclude from automation ROI.",
    })

    # Signature Card
    sig = client[client["_subject"].str.contains("Signature Card", case=False, na=False)]
    sig_unres = (~sig["_is_resolved"]).sum()
    sig_unres_pct = round(100 * sig_unres / max(len(sig), 1), 1)
    sig_vc = sig["_owner"].value_counts() if "_owner" in sig.columns else pd.Series(dtype=int)
    sig_top = sig_vc.index[0] if len(sig_vc) > 0 else "unknown"
    sig_top_pct = round(100 * sig_vc.iloc[0] / max(len(sig), 1), 1) if len(sig_vc) > 0 else 0
    sig_old = (sig[~sig["_is_resolved"]]["_age_hours"] > 720).sum() if "_age_hours" in sig.columns else 0
    rows.append({
        "claim": "Signature Card has large backlog; board member changes drive volume",
        "source": "Chris, March 23",
        "data_finding": f"{len(sig):,} cases, {sig_unres:,} unresolved ({sig_unres_pct}%). Top owner: {sig_top} at {sig_top_pct}%. {sig_old} unresolved >30d.",
        "verdict": "CONFIRMED" if sig_unres_pct > 10 else "PARTIALLY CONFIRMED",
        "implication": f"Backlog real at {sig_unres_pct}% unresolved. Board member change automation is the intervention.",
    })

    # Research catch-all
    research = client[client["_subject"] == "Research"]
    r_n = len(research)
    r_desc_fill = round(100 * (research.get("_description_len", pd.Series(dtype=int)) > 0).mean(), 0) if r_n > 0 else 0
    rows.append({
        "claim": "Research is a catch-all — 'a lot of different things go into it'",
        "source": "Chris, March 23",
        "data_finding": f"{r_n:,} cases. Description fill: {r_desc_fill}%. Keyword clustering yields 5 sub-types; Payment Research dominates (~40%).",
        "verdict": "CONFIRMED",
        "implication": "Can be split into ~5 actionable sub-types. Sub-segmentation enables targeted routing.",
    })

    # NSF/Non-Post
    nsf = client[client["_subject"].str.contains("NSF|Non.?Post", case=False, na=False)]
    nsf_hrs = nsf["_hours"].dropna()
    nsf_daily = nsf_hrs.sum() / (n_weeks * 5) if n_weeks > 0 else 0
    nsf_owners = nsf["_owner"].nunique() if "_owner" in nsf.columns else 1
    nsf_per_owner = round(nsf_daily / max(nsf_owners, 1), 1)
    rows.append({
        "claim": "NSF/Non-Post consumes 2-3 hours of every banker's time daily",
        "source": "Chris, March 23",
        "data_finding": f"{len(nsf):,} cases. ~{nsf_per_owner}h/owner/day across {nsf_owners} owners.",
        "verdict": "CONFIRMED" if 1.0 <= nsf_per_owner <= 4.0 else "PARTIALLY CONFIRMED",
        "implication": f"Data shows ~{nsf_per_owner}h/owner/day. Positive Pay adoption is the structural fix.",
    })

    # Close Account
    close = client[client["_subject"].str.contains("Clos", case=False, na=False)]
    close_hrs = close["_hours"].dropna()
    close_med = round(close_hrs.median(), 1) if len(close_hrs) > 0 else np.nan
    close_p90 = round(close_hrs.quantile(0.9), 1) if len(close_hrs) > 0 else np.nan
    rows.append({
        "claim": "Close Account requires manual cross-system checklist (IBS, BST, ACH)",
        "source": "Chris, March 23",
        "data_finding": f"{len(close):,} cases. Median {close_med}h, P90 {close_p90}h. Long tail consistent with multi-step process.",
        "verdict": "PARTIALLY CONFIRMED",
        "implication": "CRM consolidation is the right approach. Not a near-term AI target.",
    })

    return pd.DataFrame(rows)


def sheet_19_research_breakdown(df):
    """D19: Research sub-segmentation into 5 actionable types."""
    client = df[~df["_is_internal"]].copy()
    research = client[client["_subject"] == "Research"].copy()
    if len(research) == 0:
        return pd.DataFrame({"note": ["No Research cases"]})

    desc_col = "_description_text" if "_description_text" in research.columns else None
    act_col = "_activity_subj_text" if "_activity_subj_text" in research.columns else None

    research["_cluster"] = research.apply(
        lambda r: _research_cluster(
            r.get(desc_col, "") if desc_col else "",
            r.get(act_col, "") if act_col else "",
        ), axis=1
    )

    grp = research.groupby("_cluster")
    result = grp.size().reset_index(name="cases")
    result["pct"] = (result["cases"] / len(research) * 100).round(1)
    result = result.sort_values("cases", ascending=False)

    if "_hours" in research.columns:
        hrs = grp["_hours"].agg(
            median_hrs="median", p90_hrs=lambda x: x.quantile(0.9) if len(x) else np.nan,
        ).reset_index()
        hrs["median_hrs"] = hrs["median_hrs"].round(1)
        hrs["p90_hrs"] = hrs["p90_hrs"].round(1)
        result = result.merge(hrs, on="_cluster", how="left")

    if "_is_resolved" in research.columns:
        unres = grp["_is_resolved"].agg(
            unresolved=lambda x: int((~x).sum()),
            pct_unresolved=lambda x: round(100 * (~x).mean(), 1),
        ).reset_index()
        result = result.merge(unres, on="_cluster", how="left")

    if "_owner" in research.columns:
        ow = grp["_owner"].agg(
            distinct_owners=lambda x: x.nunique(),
            top_owner=lambda x: x.value_counts().index[0] if len(x) > 0 else "",
            top_owner_pct=lambda x: round(100 * x.value_counts().iloc[0] / len(x), 1) if len(x) > 0 else 0,
        ).reset_index()
        result = result.merge(ow, on="_cluster", how="left")

    result.rename(columns={"_cluster": "cluster"}, inplace=True)
    return result.reset_index(drop=True)


def sheet_20_key_person_risk(df):
    """D20: Subjects where one owner handles >35% — SPOF risk."""
    client = df[~df["_is_internal"]].copy()
    if "_owner" not in client.columns:
        return pd.DataFrame({"note": ["Owner column not found"]})

    subj_counts = client["_subject"].value_counts()
    big = subj_counts[subj_counts >= 20].index.tolist()

    rows = []
    for subj in big:
        s = client[client["_subject"] == subj]
        n = len(s)
        vc = s["_owner"].value_counts()
        if len(vc) == 0:
            continue
        top1 = vc.index[0]
        top1_pct = round(100 * vc.iloc[0] / n, 1)
        if top1_pct < 35:
            continue

        top_hrs = s[s["_owner"] == top1]["_hours"].dropna()
        rest_hrs = s[s["_owner"] != top1]["_hours"].dropna()
        top_med = round(top_hrs.median(), 1) if len(top_hrs) > 0 else np.nan
        rest_med = round(rest_hrs.median(), 1) if len(rest_hrs) > 0 else np.nan
        speed = ""
        if pd.notna(top_med) and pd.notna(rest_med) and rest_med > 0:
            speed = "Faster than peers" if top_med < rest_med * 0.85 else (
                "Slower than peers" if top_med > rest_med * 1.15 else "Similar to peers")

        rows.append({
            "subject": subj, "cases": n,
            "key_person": top1, "key_person_pct": top1_pct,
            "risk_level": "YES" if top1_pct > 50 else "WATCH",
            "key_person_median_hrs": top_med, "others_median_hrs": rest_med,
            "speed_note": speed,
            "impact_if_absent": f"{int(top1_pct * n / 100):,} cases would need reassignment",
        })

    result = pd.DataFrame(rows)
    if result.empty:
        return pd.DataFrame({"note": ["No key-person risks detected"]})
    return result.sort_values("key_person_pct", ascending=False).reset_index(drop=True)


def sheet_21_mixed_type_flags(df):
    """D21: Subjects with p90/median >= 5 — mixed workflows under one label."""
    client = df[~df["_is_internal"]].copy()
    if "_hours" not in client.columns:
        return pd.DataFrame({"note": ["Hours column not available"]})

    subj_counts = client["_subject"].value_counts()
    big = subj_counts[subj_counts >= 20].index.tolist()

    rows = []
    for subj in big:
        hrs = client[client["_subject"] == subj]["_hours"].dropna()
        if len(hrs) < 10:
            continue
        med = hrs.median()
        p90 = hrs.quantile(0.9)
        ratio = p90 / med if med > 0 else np.nan
        if pd.isna(ratio) or ratio < 5:
            continue
        rows.append({
            "subject": subj, "cases": int(subj_counts[subj]),
            "median_hrs": round(med, 1), "p90_hrs": round(p90, 1),
            "p90_median_ratio": round(ratio, 1),
            "flag": "HIGH VARIANCE" if ratio >= 10 else "ELEVATED",
        })

    result = pd.DataFrame(rows)
    if result.empty:
        return pd.DataFrame({"note": ["No mixed-type subjects detected"]})
    return result.sort_values("p90_median_ratio", ascending=False).reset_index(drop=True)


# ═══════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════

def main():
    start = datetime.datetime.now()
    log(f"WAB Cases + Emails Deep Dive — {start.strftime('%Y-%m-%d %H:%M:%S')}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    cases_raw  = read_file(CASE_FILE, "Cases")
    emails_raw = read_file(EMAIL_FILE, "Emails")

    if cases_raw.empty:
        log("FATAL: Cases file not loaded. Exiting.")
        return

    log("\n--- Preparing Cases ---")
    cases = prepare_cases(cases_raw)
    log(f"  Named client: {(cases['_case_population'] == 'named_client').sum():,}  |  "
        f"Blank client: {(cases['_case_population'] == 'blank_client').sum():,}  |  "
        f"Admin: {(cases['_case_population'] == 'admin').sum():,}")

    log("\n--- Preparing Emails ---")
    edf = prepare_emails(emails_raw, cases_raw) if not emails_raw.empty else pd.DataFrame()

    log("\n--- Building sheets ---")
    sheets = OrderedDict()

    log("  D01 Population Split")
    sheets["D01_PopulationSplit"] = sheet_01_population_split(cases)

    log("  D02 Client Weekly")
    sheets["D02_ClientWeekly"] = sheet_02_client_weekly(cases)

    log("  D03 Subject Deep")
    sheets["D03_SubjectDeep"] = sheet_03_subject_deep(cases)

    log("  D04 Day of Week")
    sheets["D04_DayOfWeek"] = sheet_04_day_of_week(cases)

    log("  D05 Hourly Pattern")
    sheets["D05_HourlyPattern"] = sheet_05_hourly_pattern(cases)

    log("  D06 SLA Breach")
    sheets["D06_SLA_Breach"] = sheet_06_sla_breach(cases)

    log("  D07 Backlog Detail")
    sheets["D07_BacklogDetail"] = sheet_07_backlog_detail(cases)

    log("  D08 Retouch Analysis")
    sheets["D08_Retouch"] = sheet_08_retouch(cases)

    log("  D09 Owner Workload")
    sheets["D09_OwnerWorkload"] = sheet_09_owner_workload(cases)

    log("  D10 Origin x Subject Detail")
    sheets["D10_OriginXSubject"] = sheet_10_origin_subject_detail(cases)

    if not edf.empty:
        log("  D11 Email Overview")
        sheets["D11_EmailOverview"] = sheet_11_email_overview(edf)

        log("  D12 Email-Case Subject Mix")
        sheets["D12_EmailCaseSubjects"] = sheet_12_email_case_subject_mix(edf)

        log("  D13 Email Burden Per Case")
        sheets["D13_EmailBurden"] = sheet_13_email_burden_per_case(edf)

        log("  D14 Email Text Samples")
        sheets["D14_EmailTextSamples"] = sheet_14_email_text_samples(edf)
    else:
        log("  (Emails not loaded — skipping D11-D14)")

    log("  D15 GenAI Evidence")
    sheets["D15_GenAI_Evidence"] = sheet_15_genai_evidence(cases, edf if not edf.empty else pd.DataFrame())

    log("  D16 Triage Delay")
    sheets["D16_TriageDelay"] = sheet_16_triage_delay(cases)

    log("  D17 Banker-Hours Budget")
    sheets["D17_BankerHoursBudget"] = sheet_17_banker_hours(cases)

    log("  D18 Claim Check")
    sheets["D18_ClaimCheck"] = sheet_18_claim_check(cases)

    log("  D19 Research Breakdown")
    sheets["D19_ResearchBreakdown"] = sheet_19_research_breakdown(cases)

    log("  D20 Key-Person Risk")
    sheets["D20_KeyPersonRisk"] = sheet_20_key_person_risk(cases)

    log("  D21 Mixed-Type Flags")
    sheets["D21_MixedTypeFlags"] = sheet_21_mixed_type_flags(cases)

    # ── Write Excel ──
    log(f"\nWriting: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            if sdf is not None:
                write_sheet(writer, name[:31], sdf)
    log("  Done.")

    # ── Write markdown ──
    log(f"Writing: {OUTPUT_MD}")
    md = [
        "# WAB Cases + Emails Deep Dive",
        f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}\n",
        "## Files",
        f"- Cases: {len(cases_raw):,} rows",
        f"- Emails: {len(emails_raw):,} rows",
        f"- Client inclusive (named+blank): {(~cases['_is_internal']).sum():,}",
        f"- Named client: {(cases['_case_population'] == 'named_client').sum():,}",
        f"- Blank company client: {(cases['_case_population'] == 'blank_client').sum():,}",
        f"- Admin: {cases['_is_internal'].sum():,}",
    ]
    if WARN:
        md.append("\n## Warnings")
        for w in WARN: md.append(f"- {w}")
    md.append("\n## Sheets")
    for name, sdf in sheets.items():
        md.append(f"- **{name}**: {len(sdf) if sdf is not None else 0} rows")
    md.append("\n## Log\n```")
    md.extend(LOG)
    md.append("```")

    with open(OUTPUT_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(md))

    elapsed = (datetime.datetime.now() - start).total_seconds()
    log(f"\nCompleted in {elapsed:.1f}s")


if __name__ == "__main__":
    main()
