"""
WAB Internal Data Extract — Phase 1 GenAI Use-Case Discovery
=============================================================
Standalone script for VDI execution against 4 WAB Excel files.
Outputs one Excel workbook (16 sheets) + one markdown summary.

Dependencies: pandas, openpyxl  (standard library otherwise)
"""

# ┌─────────────────────────────────────────────────────────┐
# │  EDIT THESE 5 VARIABLES BEFORE RUNNING                  │
# └─────────────────────────────────────────────────────────┘
PMC_FILE   = r"C:\Users\YourName\Desktop\AAB - ALL PMCs.xlsx"
HOA_FILE   = r"C:\Users\YourName\Desktop\AAB - All HOAs.xlsx"
EMAIL_FILE = r"C:\Users\YourName\Desktop\ALL EMAIL Files.xlsx"
CASE_FILE  = r"C:\Users\YourName\Desktop\AAB All Cases.xlsx"
OUTPUT_DIR = r"C:\Users\YourName\Desktop\wab_output"
# ┌─────────────────────────────────────────────────────────┐
# │  DO NOT EDIT BELOW THIS LINE                            │
# └─────────────────────────────────────────────────────────┘

import os
import re
import sys
import datetime
import warnings
from pathlib import Path
from collections import OrderedDict

import pandas as pd
import numpy as np

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ── globals ──────────────────────────────────────────────
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "wab_internal_extract.xlsx")
OUTPUT_MD   = os.path.join(OUTPUT_DIR, "wab_internal_extract_summary.md")
LOG_LINES   = []          # accumulated markdown summary lines
WARNINGS    = []          # accumulated warning messages

TRUNC = 150              # max chars for text field display


# ═══════════════════════════════════════════════════════════
#  UTILITIES
# ═══════════════════════════════════════════════════════════

def log(msg):
    LOG_LINES.append(msg)
    print(msg)


def warn(msg):
    WARNINGS.append(msg)
    log(f"  WARNING: {msg}")


def trunc(val, n=TRUNC):
    """Truncate a value to n chars for display."""
    if pd.isna(val):
        return ""
    s = str(val).replace("\r", " ").replace("\n", " ").strip()
    return s[:n] + "..." if len(s) > n else s


def norm_key(val):
    """Normalize a join key: trim, upper, collapse spaces, strip punctuation."""
    if pd.isna(val):
        return np.nan
    s = str(val).strip().upper()
    s = re.sub(r"[^\w\s]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s if s else np.nan


def norm_col(name):
    """Normalize column header for matching: lower, strip, collapse spaces."""
    if not isinstance(name, str):
        return ""
    return re.sub(r"\s+", " ", name.strip().lower())


def find_col(df, *candidates):
    """Find first matching column by case-insensitive, whitespace-tolerant match.
    Returns actual column name or None."""
    lookup = {norm_col(c): c for c in df.columns}
    for cand in candidates:
        normed = norm_col(cand)
        if normed in lookup:
            return lookup[normed]
        # partial / contains fallback
        for k, v in lookup.items():
            if normed in k or k in normed:
                return v
    return None


def find_col_strict(df, *candidates):
    """Like find_col but only exact normalized match, no partial."""
    lookup = {norm_col(c): c for c in df.columns}
    for cand in candidates:
        normed = norm_col(cand)
        if normed in lookup:
            return lookup[normed]
    return None


def safe_dt(series):
    """Attempt to parse a series as datetime."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    try:
        return pd.to_datetime(series, errors="coerce")
    except Exception:
        return pd.Series([pd.NaT] * len(series), index=series.index)


def read_file(path, label):
    """Read an Excel file. Row 1 = header. Drop unnamed/empty columns."""
    log(f"\nReading {label}: {path}")
    if not os.path.isfile(path):
        warn(f"File not found: {path}")
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, header=0, engine="openpyxl")
    except Exception as e:
        warn(f"Failed to read {label}: {e}")
        return pd.DataFrame()

    orig_cols = len(df.columns)
    # Drop columns with no name or name like 'Unnamed: X'
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    # Drop columns that are entirely null
    df = df.dropna(axis=1, how="all")
    dropped = orig_cols - len(df.columns)
    log(f"  Rows: {len(df):,}  |  Columns kept: {len(df.columns)} (dropped {dropped} empty/unnamed)")
    return df


# ═══════════════════════════════════════════════════════════
#  SHEET BUILDERS
# ═══════════════════════════════════════════════════════════

def build_vitals(df, label):
    """Build a vitals table for one file: column, dtype, nulls, distinct, sample."""
    if df.empty:
        return pd.DataFrame({"note": [f"{label}: file could not be read"]})
    rows = []
    for c in df.columns:
        s = df[c]
        non_null = s.dropna()
        sample = trunc(non_null.iloc[0]) if len(non_null) > 0 else ""
        rows.append({
            "column":    c,
            "dtype":     str(s.dtype),
            "null_count": int(s.isna().sum()),
            "null_pct":  f"{100 * s.isna().mean():.1f}%",
            "distinct":  int(s.nunique(dropna=True)),
            "sample":    sample,
        })
    return pd.DataFrame(rows)


def build_date_coverage(frames):
    """Sheet 2_DateCoverage: scan all files for date-like columns."""
    rows = []
    for label, df in frames.items():
        if df.empty:
            continue
        for c in df.columns:
            if not (pd.api.types.is_datetime64_any_dtype(df[c])
                    or any(kw in norm_col(c) for kw in
                           ["date", "created", "modified", "resolved", "sla", "start", "on"])):
                continue
            dt = safe_dt(df[c])
            valid = dt.dropna()
            if len(valid) == 0:
                continue
            rows.append({
                "file":           label,
                "column":         c,
                "min_date":       str(valid.min().date()) if len(valid) else "",
                "max_date":       str(valid.max().date()) if len(valid) else "",
                "distinct_days":  int(valid.dt.date.nunique()),
                "distinct_weeks": int(valid.dt.isocalendar().week.nunique()) if len(valid) else 0,
                "distinct_months": int(valid.dt.to_period("M").nunique()) if len(valid) else 0,
            })
    if not rows:
        return pd.DataFrame({"note": ["No date columns detected"]})
    return pd.DataFrame(rows)


def build_key_candidates(frames):
    """Sheet 3_KeyCandidates: per file, identify likely PK / FK / date / text columns."""
    rows = []
    for label, df in frames.items():
        if df.empty:
            rows.append({"file": label, "role": "ERROR", "columns": "file not loaded", "notes": ""})
            continue

        # Likely primary keys: high cardinality, low nulls, contains 'id' or 'number'
        pks = []
        fks = []
        dates = []
        texts = []
        for c in df.columns:
            nc = norm_col(c)
            s = df[c]
            nuniq = s.nunique(dropna=True)
            null_pct = s.isna().mean()

            # Date
            if pd.api.types.is_datetime64_any_dtype(s) or any(
                kw in nc for kw in ["created on", "modified on", "resolved on", "sla start", "date"]
            ):
                dates.append(c)

            # Text: object dtype, median length > 20
            if s.dtype == "object":
                lengths = s.dropna().astype(str).str.len()
                if len(lengths) > 0 and lengths.median() > 20:
                    texts.append(c)

            # PK heuristics
            if any(kw in nc for kw in ["case number", "cis number", "pmc id", "ein"]):
                if null_pct < 0.05 and nuniq > 0.8 * len(df):
                    pks.append(c)
                elif null_pct < 0.3:
                    fks.append(c)

            # FK heuristics
            if any(kw in nc for kw in [
                "parent pmc", "regarding", "company name", "customer",
                "parent company", "pmc id"
            ]) and c not in pks:
                fks.append(c)

        rows.append({"file": label, "role": "Likely PK",     "columns": "; ".join(pks) or "(none detected)", "notes": ""})
        rows.append({"file": label, "role": "Likely FK",     "columns": "; ".join(fks) or "(none detected)", "notes": ""})
        rows.append({"file": label, "role": "Date fields",   "columns": "; ".join(dates) or "(none detected)", "notes": ""})
        rows.append({"file": label, "role": "Text fields",   "columns": "; ".join(texts) or "(none detected)", "notes": ""})

    return pd.DataFrame(rows)


def build_join_scorecard(pmc, hoa, cases, emails):
    """Sheet 4_JoinScorecard: test 4 join paths."""
    results = []

    def _score(left_df, left_label, left_key_cands, right_df, right_label, right_key_cands, join_name):
        lk = None
        for cand in left_key_cands:
            lk = find_col(left_df, cand)
            if lk:
                break
        rk = None
        for cand in right_key_cands:
            rk = find_col(right_df, cand)
            if rk:
                break

        if lk is None or rk is None:
            results.append({
                "join": join_name,
                "left_col": lk or f"NOT FOUND in {left_label}",
                "right_col": rk or f"NOT FOUND in {right_label}",
                "left_non_null": "",
                "raw_exact_matches": "",
                "raw_exact_pct": "",
                "norm_exact_matches": "",
                "norm_exact_pct": "",
                "top_unmatched_keys": "COLUMN MISSING — join not attempted",
            })
            return

        left_vals = left_df[lk].dropna()
        right_set_raw = set(right_df[rk].dropna().astype(str))
        right_set_norm = set(right_df[rk].dropna().apply(norm_key).dropna())

        raw_match = left_vals.astype(str).isin(right_set_raw).sum()
        norm_vals = left_vals.apply(norm_key).dropna()
        norm_match = norm_vals.isin(right_set_norm).sum()

        n = len(left_vals)
        raw_pct  = f"{100 * raw_match / n:.1f}%" if n else "N/A"
        norm_pct = f"{100 * norm_match / n:.1f}%" if n else "N/A"

        # top unmatched
        unmatched_mask = ~norm_vals.isin(right_set_norm)
        unmatched = norm_vals[unmatched_mask]
        if len(unmatched) > 0:
            top_un = unmatched.value_counts().head(5)
            top_str = "; ".join(f"{k} ({v})" for k, v in top_un.items())
        else:
            top_str = "(all matched)"

        results.append({
            "join": join_name,
            "left_col": f"{left_label}.{lk}",
            "right_col": f"{right_label}.{rk}",
            "left_non_null": f"{n:,}",
            "raw_exact_matches": f"{raw_match:,}",
            "raw_exact_pct": raw_pct,
            "norm_exact_matches": f"{norm_match:,}",
            "norm_exact_pct": norm_pct,
            "top_unmatched_keys": trunc(top_str),
        })

    if not hoa.empty and not pmc.empty:
        _score(hoa, "HOA",
               ["Parent PMC ID", "PMC ID (Parent Company) (Company)"],
               pmc, "PMC",
               ["PMC ID"],
               "HOA → PMC")

    if not emails.empty and not cases.empty:
        _score(emails, "Email",
               ["Case Number (Regarding) (Case)", "Case Number (Regarding)", "Regarding"],
               cases, "Cases",
               ["Case Number"],
               "Email → Case")

    if not cases.empty and not pmc.empty:
        _score(cases, "Cases",
               ["Company Name", "Company Name (Company) (Company)"],
               pmc, "PMC",
               ["Company Name"],
               "Case → PMC")

    if not cases.empty and not hoa.empty:
        _score(cases, "Cases",
               ["Company Name", "Company Name (Company) (Company)"],
               hoa, "HOA",
               ["Company Name"],
               "Case → HOA")

    if not results:
        return pd.DataFrame({"note": ["No joins could be tested"]})
    return pd.DataFrame(results)


def build_case_weekly(cases):
    """Sheet 5_CaseWeekly: weekly created, resolved, median hours, backlog proxy."""
    created_col = find_col(cases, "Created On")
    resolved_col = find_col(cases, "Resolved On")
    hours_col = find_col(cases, "Resolved In Hours")

    if not created_col:
        return pd.DataFrame({"note": ["Created On column not found in Cases"]})

    df = cases.copy()
    df["_created_dt"] = safe_dt(df[created_col])
    df["_week"] = df["_created_dt"].dt.to_period("W").astype(str)
    valid = df.dropna(subset=["_created_dt"]).copy()

    weekly = valid.groupby("_week").agg(
        created_count=("_week", "size"),
    ).reset_index()

    # resolved count per week
    if resolved_col:
        df["_resolved_dt"] = safe_dt(df[resolved_col])
        res = df.dropna(subset=["_resolved_dt"]).copy()
        res["_res_week"] = res["_resolved_dt"].dt.to_period("W").astype(str)
        res_wk = res.groupby("_res_week").size().reset_index(name="resolved_count")
        weekly = weekly.merge(res_wk, left_on="_week", right_on="_res_week", how="left")
        weekly.drop(columns=["_res_week"], errors="ignore", inplace=True)
    else:
        weekly["resolved_count"] = ""

    weekly["resolved_count"] = weekly["resolved_count"].fillna(0)

    # median hours
    if hours_col:
        hrs = valid.copy()
        hrs["_hours"] = pd.to_numeric(hrs[hours_col], errors="coerce")
        hrs_wk = hrs.groupby("_week")["_hours"].median().reset_index(name="median_resolved_hours")
        weekly = weekly.merge(hrs_wk, on="_week", how="left")
        weekly["median_resolved_hours"] = weekly["median_resolved_hours"].round(1)
    else:
        weekly["median_resolved_hours"] = ""

    # cumulative backlog
    weekly["resolved_count"] = pd.to_numeric(weekly["resolved_count"], errors="coerce").fillna(0).astype(int)
    weekly["cum_created"] = weekly["created_count"].cumsum()
    weekly["cum_resolved"] = weekly["resolved_count"].cumsum()
    weekly["backlog_proxy"] = weekly["cum_created"] - weekly["cum_resolved"]

    weekly.rename(columns={"_week": "week"}, inplace=True)
    return weekly[["week", "created_count", "resolved_count",
                    "median_resolved_hours", "backlog_proxy"]]


def build_case_subjects(cases):
    """Sheet 6_CaseSubjects: top 15 subjects by count + cycle time."""
    subj_col = find_col(cases, "Subject", "Subject Path", "Description")
    hours_col = find_col(cases, "Resolved In Hours")
    status_col = find_col(cases, "Status", "Status Reason")

    if not subj_col:
        return pd.DataFrame({"note": ["Subject column not found in Cases"]})

    df = cases.copy()
    df["_subj"] = df[subj_col].fillna("(blank)").astype(str)
    if hours_col:
        df["_hours"] = pd.to_numeric(df[hours_col], errors="coerce")

    grp = df.groupby("_subj")
    result = grp.size().reset_index(name="count")

    if hours_col:
        med = grp["_hours"].median().reset_index(name="median_hours")
        p90 = grp["_hours"].quantile(0.9).reset_index(name="p90_hours")
        result = result.merge(med, on="_subj").merge(p90, on="_subj")
        result["median_hours"] = result["median_hours"].round(1)
        result["p90_hours"] = result["p90_hours"].round(1)

    if status_col:
        resolved_kw = ["resolved", "closed", "cancelled", "canceled"]
        df["_is_unresolved"] = ~df[status_col].fillna("").astype(str).str.lower().apply(
            lambda x: any(kw in x for kw in resolved_kw)
        )
        unres = df.groupby("_subj")["_is_unresolved"].mean().reset_index(name="pct_unresolved")
        unres["pct_unresolved"] = (unres["pct_unresolved"] * 100).round(1).astype(str) + "%"
        result = result.merge(unres, on="_subj")

    result = result.sort_values("count", ascending=False).head(15).reset_index(drop=True)
    result.rename(columns={"_subj": "subject"}, inplace=True)
    return result


def build_case_origins(cases):
    """Sheet 7_CaseOrigins: origin distribution + origin x top subject cross-tab."""
    origin_col = find_col(cases, "Origin")
    subj_col = find_col(cases, "Subject", "Subject Path")

    if not origin_col:
        return pd.DataFrame({"note": ["Origin column not found in Cases"]})

    df = cases.copy()
    df["_origin"] = df[origin_col].fillna("(blank)").astype(str)

    # Distribution
    dist = df["_origin"].value_counts().head(10).reset_index()
    dist.columns = ["origin", "count"]
    dist["pct"] = (dist["count"] / len(df) * 100).round(1).astype(str) + "%"

    if not subj_col:
        return dist

    # Cross-tab: top 5 origins x top 5 subjects
    df["_subj"] = df[subj_col].fillna("(blank)").astype(str)
    top_origins = df["_origin"].value_counts().head(5).index.tolist()
    top_subjects = df["_subj"].value_counts().head(5).index.tolist()
    sub = df[df["_origin"].isin(top_origins) & df["_subj"].isin(top_subjects)]
    ct = pd.crosstab(sub["_origin"], sub["_subj"])

    # Append cross-tab below distribution with a separator
    sep = pd.DataFrame({"origin": ["", "--- ORIGIN x SUBJECT CROSS-TAB ---"], "count": ["", ""], "pct": ["", ""]})
    # Convert cross-tab to flat format
    ct_flat = ct.reset_index().rename(columns={"_origin": "origin"})
    # Combine — align columns loosely
    return pd.concat([dist, sep], ignore_index=True), ct_flat


def build_pmc_concentration(cases, pmc):
    """Sheet 8_PMC_Concentration: top 15 company names by case count."""
    co_col_case = find_col(cases, "Company Name", "Company Name (Company) (Company)", "Customer")
    if not co_col_case:
        return pd.DataFrame({"note": ["Company Name column not found in Cases"]})

    df = cases.copy()
    df["_co"] = df[co_col_case].fillna("(blank)").astype(str)
    top = df["_co"].value_counts().head(15).reset_index()
    top.columns = ["company_name", "case_count"]

    # Try to join deposits
    dep_col = find_col(pmc, "Est. Total Deposits", "Deposits Rollup", "Total Deposits")
    co_col_pmc = find_col(pmc, "Company Name")
    if co_col_pmc and dep_col and not pmc.empty:
        pmc_lk = pmc[[co_col_pmc, dep_col]].copy()
        pmc_lk["_co_norm"] = pmc_lk[co_col_pmc].apply(norm_key)
        pmc_lk = pmc_lk.drop_duplicates(subset=["_co_norm"])
        top["_co_norm"] = top["company_name"].apply(norm_key)
        top = top.merge(pmc_lk[["_co_norm", dep_col]], on="_co_norm", how="left")
        top["matched_pmc"] = top[dep_col].notna().map({True: "Yes", False: "No"})
        top.rename(columns={dep_col: "deposits"}, inplace=True)
        top.drop(columns=["_co_norm"], inplace=True)
    else:
        top["matched_pmc"] = ""
        top["deposits"] = ""

    return top


def build_naics_diagnostic(pmc, hoa):
    """Sheet 9_NAICS_Diagnostic: NAICS distributions and cross-tab."""
    parts = []

    for label, df in [("PMC", pmc), ("HOA", hoa)]:
        naics_col = find_col(df, "NAICS")
        if not naics_col or df.empty:
            parts.append(pd.DataFrame({"section": [f"{label} NAICS"], "note": ["NAICS column not found"]}))
            continue

        s = df[naics_col]
        total = len(s)
        null_n = int(s.isna().sum()) + int((s.astype(str).str.strip() == "").sum())
        null_pct = f"{100 * null_n / total:.1f}%" if total else "N/A"
        distinct = int(s.nunique(dropna=True))

        header = pd.DataFrame({
            "section": [f"--- {label} NAICS SUMMARY ---"],
            "value": [f"Total: {total:,} | Null/blank: {null_n:,} ({null_pct}) | Distinct: {distinct}"],
        })
        parts.append(header)

        top_n = 10 if label == "PMC" else 5
        top = s.fillna("(blank)").astype(str).str.strip().value_counts().head(top_n).reset_index()
        top.columns = ["naics", "count"]
        top["pct"] = (top["count"] / total * 100).round(1).astype(str) + "%"
        top.insert(0, "section", f"{label} Top {top_n}")
        parts.append(top)

        # PMC cross-tab: Company Type x NAICS
        if label == "PMC":
            ct_col = find_col(df, "Company Type")
            if ct_col:
                df2 = df.copy()
                df2["_naics"] = df2[naics_col].fillna("(blank)").astype(str).str.strip()
                df2["_ctype"] = df2[ct_col].fillna("(blank)").astype(str).str.strip()
                top_naics = df2["_naics"].value_counts().head(8).index.tolist()
                sub = df2[df2["_naics"].isin(top_naics)]
                ct = pd.crosstab(sub["_ctype"], sub["_naics"])
                ct_reset = ct.reset_index().rename(columns={"_ctype": "company_type"})
                ct_reset.insert(0, "section", "PMC CompanyType x NAICS")
                parts.append(ct_reset)

    # Combine all parts
    combined = pd.concat(parts, ignore_index=True, sort=False)
    return combined


def build_text_samples(cases, emails):
    """Sheet 10_TextSamples: stratified sample of text fields."""
    rows = []

    # ── Cases ──
    subj_col = find_col(cases, "Subject", "Subject Path")
    desc_col = find_col(cases, "Description")
    act_col  = find_col(cases, "Activity Subject")
    hours_col = find_col(cases, "Resolved In Hours")
    status_col = find_col(cases, "Status", "Status Reason")
    id_col = find_col(cases, "Case Number")

    if not cases.empty:
        df = cases.copy()
        if hours_col:
            df["_hours"] = pd.to_numeric(df[hours_col], errors="coerce")

        # Top-volume subjects: 5 rows
        if subj_col:
            top_subj = df[subj_col].value_counts().head(3).index.tolist()
            pool = df[df[subj_col].isin(top_subj)]
            sample = pool.head(5) if len(pool) >= 5 else pool
            for _, r in sample.iterrows():
                rows.append({
                    "source": "Case (top-volume)",
                    "id": trunc(r.get(id_col, ""), 40) if id_col else "",
                    "subject": trunc(r.get(subj_col, "")) if subj_col else "",
                    "description": trunc(r.get(desc_col, "")) if desc_col else "",
                    "activity_subject": trunc(r.get(act_col, "")) if act_col else "",
                })

        # Longest-cycle: 5 rows
        if hours_col:
            longest = df.nlargest(5, "_hours", keep="first")
            for _, r in longest.iterrows():
                rows.append({
                    "source": "Case (longest-cycle)",
                    "id": trunc(r.get(id_col, ""), 40) if id_col else "",
                    "subject": trunc(r.get(subj_col, "")) if subj_col else "",
                    "description": trunc(r.get(desc_col, "")) if desc_col else "",
                    "activity_subject": trunc(r.get(act_col, "")) if act_col else "",
                })

        # Unresolved: 5 rows
        if status_col:
            resolved_kw = ["resolved", "closed", "cancelled", "canceled"]
            unres = df[~df[status_col].fillna("").astype(str).str.lower().apply(
                lambda x: any(kw in x for kw in resolved_kw)
            )]
            sample = unres.head(5) if len(unres) >= 5 else unres
            for _, r in sample.iterrows():
                rows.append({
                    "source": "Case (unresolved)",
                    "id": trunc(r.get(id_col, ""), 40) if id_col else "",
                    "subject": trunc(r.get(subj_col, "")) if subj_col else "",
                    "description": trunc(r.get(desc_col, "")) if desc_col else "",
                    "activity_subject": trunc(r.get(act_col, "")) if act_col else "",
                })

    # ── Emails ──
    e_subj_col = find_col(emails, "Subject")
    e_desc_col = find_col(emails, "Description")
    e_from_col = find_col(emails, "From")
    e_case_col = find_col(emails, "Case Number (Regarding) (Case)", "Case Number (Regarding)", "Regarding")

    if not emails.empty:
        df = emails.copy()
        has_case = False
        if e_case_col:
            has_case = True
            linked = df[df[e_case_col].notna()]
            unlinked = df[df[e_case_col].isna()]
        else:
            linked = pd.DataFrame()
            unlinked = df

        for lbl, pool, n in [("Email (linked)", linked, 5), ("Email (unlinked)", unlinked, 5)]:
            sample = pool.head(n) if len(pool) >= n else pool
            for _, r in sample.iterrows():
                rows.append({
                    "source": lbl,
                    "id": trunc(r.get(e_case_col, ""), 40) if e_case_col else "",
                    "subject": trunc(r.get(e_subj_col, "")) if e_subj_col else "",
                    "description": trunc(r.get(e_desc_col, "")) if e_desc_col else "",
                    "activity_subject": trunc(r.get(e_from_col, ""), 80) if e_from_col else "",
                })

    if not rows:
        return pd.DataFrame({"note": ["No text samples could be extracted"]})
    return pd.DataFrame(rows)


def build_text_stats(cases, emails):
    """Sheet 11_TextFieldStats: fill rate and length stats for text fields."""
    rows = []
    specs = [
        ("Cases", cases, ["Subject", "Description", "Activity Subject"]),
        ("Emails", emails, ["Subject", "Description", "From", "To"]),
    ]
    for label, df, candidates in specs:
        if df.empty:
            continue
        for cand in candidates:
            col = find_col(df, cand)
            if not col:
                rows.append({"file": label, "field": cand, "non_null_pct": "NOT FOUND",
                             "median_chars": "", "p90_chars": ""})
                continue
            s = df[col]
            non_null = s.dropna()
            lengths = non_null.astype(str).str.len()
            rows.append({
                "file": label,
                "field": col,
                "non_null_pct": f"{100 * len(non_null) / len(df):.1f}%",
                "median_chars": int(lengths.median()) if len(lengths) > 0 else 0,
                "p90_chars": int(lengths.quantile(0.9)) if len(lengths) > 0 else 0,
            })
    if not rows:
        return pd.DataFrame({"note": ["No text fields found"]})
    return pd.DataFrame(rows)


def build_email_day_profile(emails):
    """Sheet 12_EmailDayProfile: one-day diagnostic snapshot."""
    if emails.empty:
        return pd.DataFrame({"note": ["Emails file not loaded"]})

    parts = []

    # Summary
    n = len(emails)
    case_col = find_col(emails, "Case Number (Regarding) (Case)", "Case Number (Regarding)", "Regarding")
    linkage = ""
    if case_col:
        linked = emails[case_col].notna().sum()
        linkage = f"{linked:,} / {n:,} ({100 * linked / n:.1f}%)"
    summary = pd.DataFrame({
        "metric": ["total_rows", "case_linkage"],
        "value": [f"{n:,}", linkage if linkage else "case ref column not found"],
    })
    summary.insert(0, "section", "SUMMARY")
    parts.append(summary)

    # Status Reason
    sr_col = find_col(emails, "Status Reason")
    if sr_col:
        dist = emails[sr_col].fillna("(blank)").value_counts().head(10).reset_index()
        dist.columns = ["metric", "value"]
        dist.insert(0, "section", "STATUS REASON")
        parts.append(dist)

    # Top subjects
    subj_col = find_col(emails, "Subject")
    if subj_col:
        top = emails[subj_col].fillna("(blank)").astype(str).apply(lambda x: trunc(x, 80)).value_counts().head(10).reset_index()
        top.columns = ["metric", "value"]
        top.insert(0, "section", "TOP SUBJECTS")
        parts.append(top)

    # Top Regarding
    reg_col = find_col(emails, "Regarding", "Subject (Regarding) (Case)", "Subject Path (Regarding) (Case)")
    if reg_col:
        top = emails[reg_col].fillna("(blank)").astype(str).apply(lambda x: trunc(x, 80)).value_counts().head(10).reset_index()
        top.columns = ["metric", "value"]
        top.insert(0, "section", "TOP REGARDING")
        parts.append(top)

    # Hour of day
    created_col = find_col(emails, "Created On")
    if created_col:
        dt = safe_dt(emails[created_col]).dropna()
        if len(dt) > 0:
            hrs = dt.dt.hour.value_counts().sort_index().reset_index()
            hrs.columns = ["metric", "value"]
            hrs["metric"] = hrs["metric"].astype(str) + ":00"
            hrs.insert(0, "section", "HOUR OF DAY")
            parts.append(hrs)

    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame({"note": ["No email metrics computed"]})


def build_entity_hierarchy(pmc, hoa):
    """Sheet 13_EntityHierarchy: HOA→PMC linkage and concentration."""
    if hoa.empty:
        return pd.DataFrame({"note": ["HOA file not loaded"]})

    parts = []
    parent_col = find_col(hoa, "Parent PMC ID", "PMC ID (Parent Company) (Company)")

    if not parent_col:
        return pd.DataFrame({"note": ["Parent PMC ID column not found in HOA file"]})

    total = len(hoa)
    filled = hoa[parent_col].notna().sum()
    fill_pct = f"{100 * filled / total:.1f}%"

    # Match rate
    match_n = ""
    if not pmc.empty:
        pmc_id_col = find_col(pmc, "PMC ID")
        if pmc_id_col:
            pmc_ids = set(pmc[pmc_id_col].dropna().astype(str))
            matched = hoa[parent_col].dropna().astype(str).isin(pmc_ids).sum()
            match_n = f"{matched:,} / {filled:,} ({100 * matched / filled:.1f}%)" if filled else "N/A"

    summary = pd.DataFrame({
        "metric": ["total_hoas", "parent_pmc_filled", "fill_pct", "matched_to_pmc"],
        "value": [f"{total:,}", f"{filled:,}", fill_pct, match_n if match_n else "PMC ID col not found"],
    })
    summary.insert(0, "section", "LINKAGE SUMMARY")
    parts.append(summary)

    # HOAs per PMC distribution
    hoas_per = hoa[parent_col].dropna().astype(str).value_counts()
    if len(hoas_per) > 0:
        dist = pd.DataFrame({
            "metric": ["min", "p25", "median", "p75", "p90", "max", "distinct_pmcs"],
            "value": [
                int(hoas_per.min()),
                int(hoas_per.quantile(0.25)),
                int(hoas_per.median()),
                int(hoas_per.quantile(0.75)),
                int(hoas_per.quantile(0.90)),
                int(hoas_per.max()),
                len(hoas_per),
            ],
        })
        dist.insert(0, "section", "HOAS PER PMC")
        parts.append(dist)

    # Top 10 PMCs by HOA count
    top_pmc = hoas_per.head(10).reset_index()
    top_pmc.columns = ["metric", "value"]
    top_pmc["metric"] = top_pmc["metric"].astype(str)
    top_pmc.insert(0, "section", "TOP 10 PMCs BY HOA COUNT")

    # Try to resolve PMC name
    if not pmc.empty:
        pmc_id_col = find_col(pmc, "PMC ID")
        pmc_name_col = find_col(pmc, "Company Name")
        if pmc_id_col and pmc_name_col:
            lk = pmc[[pmc_id_col, pmc_name_col]].drop_duplicates(subset=[pmc_id_col])
            lk[pmc_id_col] = lk[pmc_id_col].astype(str)
            top_pmc = top_pmc.merge(lk, left_on="metric", right_on=pmc_id_col, how="left")
            top_pmc["metric"] = top_pmc.apply(
                lambda r: f"{trunc(r.get(pmc_name_col, ''), 50)} ({r['metric']})"
                          if pd.notna(r.get(pmc_name_col)) else r["metric"],
                axis=1)
            top_pmc = top_pmc[["section", "metric", "value"]]

    parts.append(top_pmc)

    # HOA status distribution
    status_col = find_col(hoa, "Status")
    if status_col:
        st = hoa[status_col].fillna("(blank)").value_counts().head(8).reset_index()
        st.columns = ["metric", "value"]
        st.insert(0, "section", "HOA STATUS")
        parts.append(st)

    return pd.concat(parts, ignore_index=True)


def build_unresolved_aging(cases):
    """Sheet 14_UnresolvedAging: aging buckets for open cases."""
    status_col = find_col(cases, "Status", "Status Reason")
    created_col = find_col(cases, "Created On")
    subj_col = find_col(cases, "Subject", "Subject Path")
    co_col = find_col(cases, "Company Name", "Company Name (Company) (Company)")

    if cases.empty or not status_col or not created_col:
        return pd.DataFrame({"note": ["Cases file or required columns not available"]})

    df = cases.copy()
    resolved_kw = ["resolved", "closed", "cancelled", "canceled"]
    df["_is_unresolved"] = ~df[status_col].fillna("").astype(str).str.lower().apply(
        lambda x: any(kw in x for kw in resolved_kw)
    )
    unres = df[df["_is_unresolved"]].copy()

    if len(unres) == 0:
        return pd.DataFrame({"note": ["No unresolved cases found"]})

    parts = []

    # Summary
    summary = pd.DataFrame({
        "metric": ["total_unresolved", "pct_of_all_cases"],
        "value": [f"{len(unres):,}", f"{100 * len(unres) / len(df):.1f}%"],
    })
    summary.insert(0, "section", "SUMMARY")
    parts.append(summary)

    # Aging buckets
    unres["_created_dt"] = safe_dt(unres[created_col])
    now = pd.Timestamp.now()
    unres["_age_hours"] = (now - unres["_created_dt"]).dt.total_seconds() / 3600

    bins = [0, 24, 72, 168, 336, float("inf")]
    labels_b = ["0-24h", "1-3d", "3-7d", "7-14d", "14d+"]
    unres["_bucket"] = pd.cut(unres["_age_hours"], bins=bins, labels=labels_b, right=True)
    aging = unres["_bucket"].value_counts().reindex(labels_b).fillna(0).astype(int).reset_index()
    aging.columns = ["metric", "value"]
    aging.insert(0, "section", "AGING BUCKETS")
    parts.append(aging)

    # Top subjects among unresolved
    if subj_col:
        top_subj = unres[subj_col].fillna("(blank)").value_counts().head(10).reset_index()
        top_subj.columns = ["metric", "value"]
        top_subj.insert(0, "section", "TOP UNRESOLVED SUBJECTS")
        parts.append(top_subj)

    # Top companies among unresolved
    if co_col:
        top_co = unres[co_col].fillna("(blank)").astype(str).apply(
            lambda x: trunc(x, 60)
        ).value_counts().head(10).reset_index()
        top_co.columns = ["metric", "value"]
        top_co.insert(0, "section", "TOP UNRESOLVED COMPANIES")
        parts.append(top_co)

    return pd.concat(parts, ignore_index=True)


# ═══════════════════════════════════════════════════════════
#  EXCEL WRITER HELPERS
# ═══════════════════════════════════════════════════════════

def write_sheet(writer, name, df):
    """Write a DataFrame to a named sheet with frozen top row and auto-fit."""
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        df = pd.DataFrame({"note": ["No data"]})
    df.to_excel(writer, sheet_name=name, index=False, freeze_panes=(1, 0))
    ws = writer.sheets[name]
    for i, col_cells in enumerate(ws.columns):
        max_len = max(len(str(cell.value or "")) for cell in col_cells)
        max_len = min(max_len + 2, 60)  # cap width
        ws.column_dimensions[col_cells[0].column_letter].width = max_len


# ═══════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════

def main():
    start = datetime.datetime.now()
    log(f"WAB Internal Extract — started {start.strftime('%Y-%m-%d %H:%M:%S')}")
    log(f"Output directory: {OUTPUT_DIR}")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ── Read files ──
    pmc    = read_file(PMC_FILE,   "PMC")
    hoa    = read_file(HOA_FILE,   "HOA")
    emails = read_file(EMAIL_FILE, "Emails")
    cases  = read_file(CASE_FILE,  "Cases")

    frames = OrderedDict([("PMC", pmc), ("HOA", hoa), ("Cases", cases), ("Emails", emails)])

    # ── Build all sheets ──
    log("\n--- Building sheets ---")

    sheets = OrderedDict()

    log("  1A_PMC_Vitals")
    sheets["1A_PMC_Vitals"]  = build_vitals(pmc, "PMC")

    log("  1B_HOA_Vitals")
    sheets["1B_HOA_Vitals"]  = build_vitals(hoa, "HOA")

    log("  1C_Case_Vitals")
    sheets["1C_Case_Vitals"] = build_vitals(cases, "Cases")

    log("  1D_Email_Vitals")
    sheets["1D_Email_Vitals"] = build_vitals(emails, "Emails")

    log("  2_DateCoverage")
    sheets["2_DateCoverage"] = build_date_coverage(frames)

    log("  3_KeyCandidates")
    sheets["3_KeyCandidates"] = build_key_candidates(frames)

    log("  4_JoinScorecard")
    sheets["4_JoinScorecard"] = build_join_scorecard(pmc, hoa, cases, emails)

    log("  5_CaseWeekly")
    sheets["5_CaseWeekly"] = build_case_weekly(cases)

    log("  6_CaseSubjects")
    sheets["6_CaseSubjects"] = build_case_subjects(cases)

    log("  7_CaseOrigins")
    result = build_case_origins(cases)
    if isinstance(result, tuple):
        sheets["7_CaseOrigins"] = result[0]
        sheets["7B_OriginXSubj"] = result[1]
    else:
        sheets["7_CaseOrigins"] = result

    log("  8_PMC_Concentration")
    sheets["8_PMC_Concentration"] = build_pmc_concentration(cases, pmc)

    log("  9_NAICS_Diagnostic")
    sheets["9_NAICS_Diagnostic"] = build_naics_diagnostic(pmc, hoa)

    log("  10_TextSamples")
    sheets["10_TextSamples"] = build_text_samples(cases, emails)

    log("  11_TextFieldStats")
    sheets["11_TextFieldStats"] = build_text_stats(cases, emails)

    log("  12_EmailDayProfile")
    sheets["12_EmailDayProfile"] = build_email_day_profile(emails)

    log("  13_EntityHierarchy")
    sheets["13_EntityHierarchy"] = build_entity_hierarchy(pmc, hoa)

    log("  14_UnresolvedAging")
    sheets["14_UnresolvedAging"] = build_unresolved_aging(cases)

    # ── Write Excel ──
    log(f"\nWriting workbook: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for name, df in sheets.items():
            if df is not None:
                write_sheet(writer, name[:31], df)  # Excel 31-char sheet name limit
    log("  Workbook written.")

    # ── Write Markdown summary ──
    log(f"Writing summary: {OUTPUT_MD}")
    md = []
    md.append(f"# WAB Internal Extract Summary")
    md.append(f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}\n")

    md.append("## Files loaded")
    for label, df in frames.items():
        md.append(f"- **{label}**: {len(df):,} rows, {len(df.columns)} columns")

    if WARNINGS:
        md.append("\n## Warnings")
        for w in WARNINGS:
            md.append(f"- {w}")

    md.append("\n## Sheets produced")
    for name, df in sheets.items():
        if df is not None:
            md.append(f"- **{name}**: {len(df)} rows")

    md.append("\n## Console log")
    md.append("```")
    md.extend(LOG_LINES)
    md.append("```")

    with open(OUTPUT_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(md))

    elapsed = (datetime.datetime.now() - start).total_seconds()
    log(f"\nDone in {elapsed:.1f}s.  Output: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
