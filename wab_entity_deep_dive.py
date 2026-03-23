"""
WAB Entity & Relationship Deep Dive — Phase 1 GenAI Use-Case Discovery
========================================================================
Standalone VDI script.  Reads all 4 WAB files.
Focuses on PMC/HOA entity analytics enriched with Cases + Emails.
Produces one Excel workbook + one markdown summary.

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

import os, re, datetime, warnings
from collections import OrderedDict

import pandas as pd
import numpy as np

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "wab_entity_deep_dive.xlsx")
OUTPUT_MD   = os.path.join(OUTPUT_DIR, "wab_entity_deep_dive_summary.md")
LOG, WARN   = [], []
TRUNC       = 100

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
    lookup = {norm_col(c): c for c in df.columns}
    for cand in candidates:
        normed = norm_col(cand)
        if normed in lookup:
            return lookup[normed]
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

def safe_num(series):
    if pd.api.types.is_numeric_dtype(series):
        return series
    return pd.to_numeric(series, errors="coerce")

def norm_key(val):
    if pd.isna(val): return np.nan
    s = str(val).strip().upper()
    s = re.sub(r"[^\w\s]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s if s else np.nan

def pct(num, denom):
    return f"{100*num/denom:.1f}%" if denom else "N/A"

def fmt_usd(val):
    if pd.isna(val) or val == 0: return ""
    if abs(val) >= 1e9: return f"${val/1e9:.2f}B"
    if abs(val) >= 1e6: return f"${val/1e6:.1f}M"
    if abs(val) >= 1e3: return f"${val/1e3:.0f}K"
    return f"${val:.0f}"

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

def write_sheet(writer, name, df):
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        df = pd.DataFrame({"note": ["No data"]})
    sheet_name = name[:31]
    df.to_excel(writer, sheet_name=sheet_name, index=False, freeze_panes=(1, 0))
    ws = writer.sheets[sheet_name]
    for i, col_cells in enumerate(ws.columns):
        max_len = max(len(str(cell.value or "")) for cell in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 55)


# ═══════════════════════════════════════════════════════════
#  DATA PREPARATION
# ═══════════════════════════════════════════════════════════

# Module-level col_map storage — pandas loses custom attributes on
# DataFrame operations (filter, reset_index, copy, merge).  Storing
# col_maps here guarantees they survive any downstream transformation.
_COL_MAPS = {}

def prepare_pmc(pmc):
    df = pmc.copy()
    col_map = {
        "pmc_id":       find_col(df, "PMC ID"),
        "company_name": find_col(df, "Company Name"),
        "cis_number":   find_col(df, "CIS Number"),
        "ein":          find_col(df, "EIN"),
        "company_type": find_col(df, "Company Type"),
        "naics":        find_col(df, "NAICS"),
        "deposits":     find_col(df, "Est. Total Deposits"),
        "deposits_roll": find_col(df, "Deposits Rollup"),
        "state":        find_col(df, "Address 1: State/Province"),
        "city":         find_col(df, "Address 1: City"),
        "rm":           find_col(df, "Relationship Manager"),
        "orig_officer":  find_col(df, "Originating Officer"),
        "pod":          find_col(df, "POD Name (Originating Officer) (User)"),
        "rm_checkin":   find_col(df, "RM Last Check-in"),
        "acct_platform": find_col(df, "Accounting Platform"),
    }

    # Always prefer Deposits Rollup — it is the confirmed single source of truth
    # (consolidated weekly, includes HOA deposits + ICS/CDARS per stakeholder).
    # Est. Total Deposits may contain stale or aggregated values.
    dep_col = col_map["deposits"]
    dep_roll = col_map["deposits_roll"]
    if dep_col:
        df["_deposits"] = safe_num(df[dep_col])
    if dep_roll:
        df["_deposits_roll"] = safe_num(df[dep_roll])
        df["_deposits_best"] = df["_deposits_roll"]
        log("  Using Deposits Rollup as primary deposit field (stakeholder-confirmed source)")
    elif dep_col:
        df["_deposits_best"] = df["_deposits"]
        log("  WARNING: Deposits Rollup not found, falling back to Est. Total Deposits")
    else:
        df["_deposits_best"] = np.nan

    if col_map["state"]:
        df["_state"] = df[col_map["state"]].fillna("(blank)").astype(str).str.strip().str.upper()

    if col_map["company_name"]:
        df["_name"] = df[col_map["company_name"]].fillna("").astype(str).str.strip()
        df["_name_norm"] = df["_name"].apply(norm_key)

    if col_map["company_type"]:
        df["_type"] = df[col_map["company_type"]].fillna("(blank)").astype(str).str.strip()

    if col_map["rm_checkin"]:
        df["_rm_checkin_dt"] = safe_dt(df[col_map["rm_checkin"]])

    if col_map["pmc_id"]:
        df["_pmc_id"] = df[col_map["pmc_id"]].astype(str).str.strip()

    # Exclude blank-name PMC rows — these carry aggregated deposit figures
    # (e.g. HOA-level rollups) that distort concentration and top-PMC metrics.
    if "_name" in df.columns:
        blank = df["_name"].str.strip().eq("")
        if blank.any():
            log(f"  Excluding {blank.sum()} blank-name PMC row(s)")
            df = df[~blank].reset_index(drop=True)

    _COL_MAPS["pmc"] = col_map
    return df


def prepare_hoa(hoa):
    df = hoa.copy()
    col_map = {
        "cis_number":   find_col(df, "CIS Number"),
        "company_name": find_col(df, "Company Name"),
        "parent_pmc_id": find_col(df, "Parent PMC ID"),
        "parent_company": find_col(df, "Parent Company"),
        "pmc_id_parent": find_col(df, "PMC ID (Parent Company) (Company)"),
        "company_type": find_col(df, "Company Type"),
        "deposits_roll": find_col(df, "Deposits Rollup"),
        "state":        find_col(df, "Address 1: State/Province"),
        "status":       find_col(df, "Status"),
    }

    if col_map["parent_pmc_id"]:
        df["_parent_pmc_id"] = df[col_map["parent_pmc_id"]].astype(str).str.strip()

    if col_map["state"]:
        df["_state"] = df[col_map["state"]].fillna("(blank)").astype(str).str.strip().str.upper()

    if col_map["deposits_roll"]:
        df["_deposits_hoa"] = safe_num(df[col_map["deposits_roll"]])

    if col_map["status"]:
        df["_status"] = df[col_map["status"]].fillna("(blank)").astype(str).str.strip()

    _COL_MAPS["hoa"] = col_map
    return df


def prepare_cases_light(cases):
    """Lightweight case prep — just what's needed for entity joins."""
    df = cases.copy()
    col_map = {
        "case_number":  find_col(df, "Case Number"),
        "company":      find_col(df, "Company Name (Company) (Company)", "Company Name", "Customer"),
        "subject":      find_col(df, "Subject"),
        "created_on":   find_col(df, "Created On"),
        "resolved_on":  find_col(df, "Resolved On"),
        "resolved_hrs": find_col(df, "Resolved In Hours"),
        "status_reason": find_col(df, "Status Reason"),
        "origin":       find_col(df, "Origin"),
    }

    if col_map["company"]:
        df["_company"] = df[col_map["company"]].fillna("(blank)").astype(str).str.strip()
        df["_company_norm"] = df["_company"].apply(norm_key)
        df["_is_internal"] = df["_company"].str.upper().apply(
            lambda x: any(x.startswith(kw) for kw in INTERNAL_COMPANIES) or x == "(BLANK)"
        )
    else:
        df["_is_internal"] = False

    if col_map["resolved_hrs"]:
        df["_hours"] = safe_num(df[col_map["resolved_hrs"]])

    if col_map["subject"]:
        df["_subject"] = df[col_map["subject"]].fillna("(blank)").astype(str).str.strip()

    resolved_kw = ["resolved", "closed", "cancelled", "canceled"]
    sr = col_map.get("status_reason")
    if sr:
        df["_is_resolved"] = df[sr].fillna("").astype(str).str.lower().apply(
            lambda x: any(kw in x for kw in resolved_kw))
    else:
        df["_is_resolved"] = True

    _COL_MAPS["cases"] = col_map
    return df


def prepare_emails_light(emails):
    df = emails.copy()
    col_map = {
        "case_number": find_col(df, "Case Number (Regarding) (Case)"),
    }
    if col_map["case_number"]:
        df["_case_ref"] = df[col_map["case_number"]].astype(str).str.strip()
    _COL_MAPS["emails"] = col_map
    return df


# ═══════════════════════════════════════════════════════════
#  BUILD THE JOINED PMC TABLE  (central to most sheets)
# ═══════════════════════════════════════════════════════════

def build_pmc_master(pmc, hoa, cases, emails):
    """Create a PMC-level summary table joining HOA counts, case stats, email stats."""
    log("\n--- Building PMC master join ---")

    if pmc.empty or "_pmc_id" not in pmc.columns:
        warn("PMC file not usable for master join")
        return pd.DataFrame()

    master = pmc[["_pmc_id"]].copy()
    if "_name" in pmc.columns:
        master["company_name"] = pmc["_name"]
    if "_type" in pmc.columns:
        master["company_type"] = pmc["_type"]
    if "_state" in pmc.columns:
        master["state"] = pmc["_state"]
    if "_deposits_best" in pmc.columns:
        master["deposits"] = pmc["_deposits_best"]
    if "_rm_checkin_dt" in pmc.columns:
        master["rm_last_checkin"] = pmc["_rm_checkin_dt"]

    col_rm = _COL_MAPS["pmc"].get("rm")
    if col_rm:
        master["relationship_manager"] = pmc[col_rm].fillna("").astype(str)

    col_pod = _COL_MAPS["pmc"].get("pod")
    if col_pod:
        master["pod"] = pmc[col_pod].fillna("").astype(str)

    col_platform = _COL_MAPS["pmc"].get("acct_platform")
    if col_platform:
        master["accounting_platform"] = pmc[col_platform].fillna("").astype(str)

    # ── HOA counts ──
    if not hoa.empty and "_parent_pmc_id" in hoa.columns:
        hoa_counts = hoa.groupby("_parent_pmc_id").agg(
            hoa_count=("_parent_pmc_id", "size"),
        ).reset_index()

        if "_deposits_hoa" in hoa.columns:
            hoa_dep = hoa.groupby("_parent_pmc_id")["_deposits_hoa"].sum().reset_index(name="hoa_deposits_sum")
            hoa_counts = hoa_counts.merge(hoa_dep, on="_parent_pmc_id", how="left")

        if "_state" in hoa.columns:
            hoa_states = hoa.groupby("_parent_pmc_id")["_state"].nunique().reset_index(name="hoa_state_count")
            hoa_counts = hoa_counts.merge(hoa_states, on="_parent_pmc_id", how="left")

        master = master.merge(hoa_counts, left_on="_pmc_id", right_on="_parent_pmc_id", how="left")
        master.drop(columns=["_parent_pmc_id"], errors="ignore", inplace=True)
    else:
        master["hoa_count"] = np.nan

    # ── Case stats (join on normalized company name) ──
    if not cases.empty and "_company_norm" in cases.columns and "_name_norm" in pmc.columns:
        client_cases = cases[~cases["_is_internal"]].copy()
        # Build PMC name lookup
        name_to_pmc = pmc[["_pmc_id", "_name_norm"]].drop_duplicates(subset=["_name_norm"])

        client_cases = client_cases.merge(name_to_pmc, left_on="_company_norm", right_on="_name_norm", how="left")
        matched = client_cases[client_cases["_pmc_id"].notna()]

        if len(matched) > 0:
            case_stats = matched.groupby("_pmc_id").agg(
                case_count=("_pmc_id", "size"),
            ).reset_index()

            if "_hours" in matched.columns:
                hrs_stats = matched.groupby("_pmc_id")["_hours"].agg(
                    median_hrs="median",
                    p90_hrs=lambda x: x.quantile(0.9) if len(x) else np.nan,
                ).reset_index()
                hrs_stats["median_hrs"] = hrs_stats["median_hrs"].round(1)
                hrs_stats["p90_hrs"] = hrs_stats["p90_hrs"].round(1)
                case_stats = case_stats.merge(hrs_stats, on="_pmc_id", how="left")

            if "_is_resolved" in matched.columns:
                unres = matched.groupby("_pmc_id")["_is_resolved"].agg(
                    unresolved=lambda x: (~x).sum()
                ).reset_index()
                case_stats = case_stats.merge(unres, on="_pmc_id", how="left")

            if "_subject" in matched.columns:
                top_subj = matched.groupby("_pmc_id")["_subject"].agg(
                    top_subject=lambda x: x.value_counts().index[0] if len(x) else ""
                ).reset_index()
                case_stats = case_stats.merge(top_subj, on="_pmc_id", how="left")

            master = master.merge(case_stats, on="_pmc_id", how="left")
        else:
            master["case_count"] = 0

        # ── Email counts via case linkage ──
        if not emails.empty and "_case_ref" in emails.columns:
            case_num_col = _COL_MAPS["cases"].get("case_number")
            if case_num_col:
                case_pmc = matched[[case_num_col, "_pmc_id"]].copy()
                case_pmc[case_num_col] = case_pmc[case_num_col].astype(str)
                email_joined = emails.merge(case_pmc, left_on="_case_ref", right_on=case_num_col, how="inner")
                if len(email_joined) > 0:
                    email_stats = email_joined.groupby("_pmc_id").size().reset_index(name="email_count")
                    master = master.merge(email_stats, on="_pmc_id", how="left")

    # Fill NaN with 0 for count columns
    for c in ["hoa_count", "case_count", "email_count", "unresolved"]:
        if c in master.columns:
            master[c] = master[c].fillna(0).astype(int)

    master.drop(columns=["_pmc_id"], inplace=True)
    log(f"  PMC master: {len(master)} rows, {len(master.columns)} columns")
    return master


# ═══════════════════════════════════════════════════════════
#  SHEET BUILDERS
# ═══════════════════════════════════════════════════════════

def sheet_e01_deposit_concentration(pmc):
    """E01: Deposit distribution and concentration."""
    if "_deposits_best" not in pmc.columns:
        return pd.DataFrame({"note": ["No deposit column found"]})

    deps = pmc["_deposits_best"].dropna()
    deps = deps[deps > 0]

    parts = []

    # Distribution
    dist = pd.DataFrame({
        "metric": ["pmcs_with_deposits", "total_deposits", "min", "p25", "median",
                   "p75", "p90", "p95", "max"],
        "value": [
            f"{len(deps):,}",
            fmt_usd(deps.sum()),
            fmt_usd(deps.min()),
            fmt_usd(deps.quantile(0.25)),
            fmt_usd(deps.median()),
            fmt_usd(deps.quantile(0.75)),
            fmt_usd(deps.quantile(0.90)),
            fmt_usd(deps.quantile(0.95)),
            fmt_usd(deps.max()),
        ]
    })
    dist.insert(0, "section", "DISTRIBUTION")
    parts.append(dist)

    # Concentration: top N PMCs by deposit share
    pmc_sorted = pmc.dropna(subset=["_deposits_best"]).sort_values("_deposits_best", ascending=False)
    total = pmc_sorted["_deposits_best"].sum()

    for n in [5, 10, 20, 50]:
        top_n = pmc_sorted.head(n)
        top_sum = top_n["_deposits_best"].sum()
        parts.append(pd.DataFrame({
            "section": [f"TOP {n} CONCENTRATION"],
            "metric": [f"top_{n}_deposit_share"],
            "value": [f"{fmt_usd(top_sum)} ({pct(top_sum, total)})"],
        }))

    # Deposit buckets
    bins = [0, 1e6, 10e6, 50e6, 100e6, 500e6, float("inf")]
    labels = ["<$1M", "$1-10M", "$10-50M", "$50-100M", "$100-500M", "$500M+"]
    pmc_sorted["_dep_bucket"] = pd.cut(pmc_sorted["_deposits_best"], bins=bins, labels=labels, right=True)
    bkt = pmc_sorted["_dep_bucket"].value_counts().reindex(labels).fillna(0).astype(int).reset_index()
    bkt.columns = ["metric", "value"]
    bkt.insert(0, "section", "DEPOSIT BUCKETS")
    parts.append(bkt)

    return pd.concat(parts, ignore_index=True, sort=False)


def sheet_e02_top_pmcs(master):
    """E02: Top 25 PMCs by deposits with full profile."""
    if master.empty or "deposits" not in master.columns:
        return pd.DataFrame({"note": ["No master table with deposits"]})

    top = master.dropna(subset=["deposits"]).sort_values("deposits", ascending=False).head(25).copy()
    top["deposits_fmt"] = top["deposits"].apply(fmt_usd)

    cols = ["company_name", "deposits_fmt", "state"]
    for c in ["hoa_count", "case_count", "email_count", "median_hrs", "p90_hrs",
              "unresolved", "top_subject", "pod", "relationship_manager"]:
        if c in top.columns:
            cols.append(c)

    return top[[c for c in cols if c in top.columns]].reset_index(drop=True)


def sheet_e03_friction_value(master):
    """E03: Cases-per-deposit-dollar analysis — the friction-vs-value map."""
    if master.empty or "deposits" not in master.columns or "case_count" not in master.columns:
        return pd.DataFrame({"note": ["Cannot compute friction-value map"]})

    df = master[(master["deposits"] > 0) & (master["case_count"] > 0)].copy()
    if len(df) == 0:
        return pd.DataFrame({"note": ["No PMCs with both deposits and cases"]})

    df["cases_per_1M_deposits"] = (df["case_count"] / (df["deposits"] / 1e6)).round(1)
    df["deposits_fmt"] = df["deposits"].apply(fmt_usd)

    # Quadrant assignment
    med_dep = df["deposits"].median()
    med_cases = df["cases_per_1M_deposits"].median()
    def _quad(row):
        hi_val = row["deposits"] >= med_dep
        hi_fric = row["cases_per_1M_deposits"] >= med_cases
        if hi_val and not hi_fric: return "High Value / Low Friction"
        if hi_val and hi_fric:     return "High Value / High Friction"
        if not hi_val and hi_fric: return "Low Value / High Friction"
        return "Low Value / Low Friction"

    df["quadrant"] = df.apply(_quad, axis=1)

    cols = ["company_name", "deposits_fmt", "case_count", "cases_per_1M_deposits", "quadrant"]
    for c in ["state", "hoa_count", "median_hrs", "unresolved", "top_subject"]:
        if c in df.columns:
            cols.append(c)

    result = df[[c for c in cols if c in df.columns]].sort_values(
        "cases_per_1M_deposits", ascending=False).head(30).reset_index(drop=True)
    return result


def sheet_e04_state_profile(pmc, hoa, cases):
    """E04: State-level profile — PMCs, HOAs, deposits, cases."""
    parts_data = {}

    # PMC by state
    if "_state" in pmc.columns:
        pmc_st = pmc.groupby("_state").size().reset_index(name="pmc_count")
        pmc_st.rename(columns={"_state": "state"}, inplace=True)
        if "_deposits_best" in pmc.columns:
            dep_st = pmc.groupby("_state")["_deposits_best"].sum().reset_index(name="pmc_deposits")
            dep_st.rename(columns={"_state": "state"}, inplace=True)
            pmc_st = pmc_st.merge(dep_st, on="state", how="left")
            pmc_st["pmc_deposits"] = pmc_st["pmc_deposits"].apply(fmt_usd)
        parts_data["pmc"] = pmc_st

    # HOA by state
    if not hoa.empty and "_state" in hoa.columns:
        hoa_st = hoa.groupby("_state").size().reset_index(name="hoa_count")
        hoa_st.rename(columns={"_state": "state"}, inplace=True)
        parts_data["hoa"] = hoa_st

    # Cases by state (via company → PMC → state)
    if not cases.empty and "_company_norm" in cases.columns and "_name_norm" in pmc.columns and "_state" in pmc.columns:
        client = cases[~cases["_is_internal"]].copy()
        name_to_state = pmc[["_name_norm", "_state"]].drop_duplicates(subset=["_name_norm"])
        client = client.merge(name_to_state, left_on="_company_norm", right_on="_name_norm", how="left")
        matched = client[client["_state"].notna()]
        if len(matched) > 0:
            case_st = matched.groupby("_state").size().reset_index(name="case_count")
            case_st.rename(columns={"_state": "state"}, inplace=True)
            parts_data["cases"] = case_st

    if not parts_data:
        return pd.DataFrame({"note": ["No state data available"]})

    # Merge all
    result = parts_data.get("pmc", pd.DataFrame({"state": []}))
    for key in ["hoa", "cases"]:
        if key in parts_data:
            result = result.merge(parts_data[key], on="state", how="outer")

    for c in ["pmc_count", "hoa_count", "case_count"]:
        if c in result.columns:
            result[c] = result[c].fillna(0).astype(int)

    sort_col = "pmc_count" if "pmc_count" in result.columns else result.columns[0]
    result = result.sort_values(sort_col, ascending=False).head(25).reset_index(drop=True)
    return result


def sheet_e05_rm_coverage(pmc):
    """E05: Relationship Manager coverage and recency."""
    if "_rm_checkin_dt" not in pmc.columns:
        return pd.DataFrame({"note": ["RM Last Check-in column not found"]})

    now = pd.Timestamp.now()
    df = pmc.copy()
    df["_days_since_checkin"] = (now - df["_rm_checkin_dt"]).dt.days

    parts = []

    # Overall coverage
    total = len(df)
    has_checkin = df["_rm_checkin_dt"].notna().sum()
    parts.append(pd.DataFrame({
        "section": ["COVERAGE"] * 3,
        "metric": ["total_pmcs", "has_rm_checkin", "checkin_fill_rate"],
        "value": [f"{total:,}", f"{has_checkin:,}", pct(has_checkin, total)],
    }))

    # Recency buckets
    valid = df.dropna(subset=["_days_since_checkin"])
    bins = [0, 30, 90, 180, 365, float("inf")]
    labels = ["<30d", "30-90d", "90-180d", "180-365d", "365d+"]
    valid["_recency"] = pd.cut(valid["_days_since_checkin"], bins=bins, labels=labels, right=True)
    rec = valid["_recency"].value_counts().reindex(labels).fillna(0).astype(int).reset_index()
    rec.columns = ["metric", "value"]
    rec.insert(0, "section", "RECENCY BUCKETS")
    parts.append(rec)

    # High-deposit PMCs with stale check-in (>180 days)
    if "_deposits_best" in df.columns:
        stale = valid[(valid["_days_since_checkin"] > 180) & (valid["_deposits_best"] > 0)]
        stale = stale.sort_values("_deposits_best", ascending=False).head(10)
        if len(stale) > 0:
            stale_tbl = stale[["_name", "_deposits_best", "_days_since_checkin", "_state"]].copy()
            stale_tbl.columns = ["metric", "deposits", "days_since_checkin", "state"]
            stale_tbl["deposits"] = stale_tbl["deposits"].apply(fmt_usd)
            stale_tbl.insert(0, "section", "HIGH-DEPOSIT STALE (>180d)")
            # flatten to section/metric/value
            rows = []
            for _, r in stale_tbl.iterrows():
                rows.append({
                    "section": "HIGH-DEPOSIT STALE (>180d)",
                    "metric": f"{r['metric']} ({r['state']})",
                    "value": f"{r['deposits']} | {int(r['days_since_checkin'])}d ago",
                })
            parts.append(pd.DataFrame(rows))

    return pd.concat(parts, ignore_index=True, sort=False)


def sheet_e06_company_type(pmc, cases):
    """E06: Company Type analysis — PMC profile by type."""
    if "_type" not in pmc.columns:
        return pd.DataFrame({"note": ["Company Type not found in PMC"]})

    grp = pmc.groupby("_type")
    result = grp.size().reset_index(name="pmc_count")
    result = result.sort_values("pmc_count", ascending=False)

    if "_deposits_best" in pmc.columns:
        dep = grp["_deposits_best"].agg(
            total_deposits="sum",
            median_deposits="median",
        ).reset_index()
        dep["total_deposits"] = dep["total_deposits"].apply(fmt_usd)
        dep["median_deposits"] = dep["median_deposits"].apply(fmt_usd)
        result = result.merge(dep, on="_type")

    # Join case counts by company type
    if not cases.empty and "_company_norm" in cases.columns and "_name_norm" in pmc.columns:
        client = cases[~cases["_is_internal"]].copy()
        name_to_type = pmc[["_name_norm", "_type"]].drop_duplicates(subset=["_name_norm"])
        client = client.merge(name_to_type, left_on="_company_norm", right_on="_name_norm", how="inner")
        if len(client) > 0:
            case_by_type = client.groupby("_type").size().reset_index(name="case_count")
            result = result.merge(case_by_type, on="_type", how="left")
            result["case_count"] = result["case_count"].fillna(0).astype(int)

            if "_hours" in client.columns:
                hrs_by_type = client.groupby("_type")["_hours"].median().round(1).reset_index(name="median_hrs")
                result = result.merge(hrs_by_type, on="_type", how="left")

    result.rename(columns={"_type": "company_type"}, inplace=True)
    return result.reset_index(drop=True)


def sheet_e07_hierarchy_depth(master, hoa, pmc):
    """E07: HOA hierarchy — concentration, depth, multi-state PMCs."""
    if master.empty or "hoa_count" not in master.columns:
        return pd.DataFrame({"note": ["No HOA join data"]})

    parts = []
    has_hoa = master[master["hoa_count"] > 0]

    # Distribution
    hoa_counts = has_hoa["hoa_count"]
    dist = pd.DataFrame({
        "metric": ["pmcs_with_hoas", "total_hoas_linked", "min", "p25", "median",
                   "p75", "p90", "p95", "max"],
        "value": [
            f"{len(has_hoa):,}", f"{int(hoa_counts.sum()):,}",
            int(hoa_counts.min()), int(hoa_counts.quantile(0.25)),
            int(hoa_counts.median()), int(hoa_counts.quantile(0.75)),
            int(hoa_counts.quantile(0.90)), int(hoa_counts.quantile(0.95)),
            int(hoa_counts.max()),
        ]
    })
    dist.insert(0, "section", "HOAS PER PMC")
    parts.append(dist)

    # Top 15 PMCs by HOA count
    top = has_hoa.sort_values("hoa_count", ascending=False).head(15)
    cols = ["company_name", "hoa_count"]
    if "deposits" in top.columns:
        top_disp = top.copy()
        top_disp["deposits_fmt"] = top_disp["deposits"].apply(fmt_usd)
        cols.append("deposits_fmt")
    else:
        top_disp = top.copy()
    for c in ["case_count", "state", "unresolved"]:
        if c in top_disp.columns:
            cols.append(c)

    top_tbl = top_disp[[c for c in cols if c in top_disp.columns]].copy()
    top_tbl.insert(0, "section", "TOP 15 BY HOA COUNT")
    top_tbl = top_tbl.rename(columns={"company_name": "metric", "hoa_count": "value"})
    parts.append(top_tbl)

    # Multi-state PMCs
    if "hoa_state_count" in master.columns:
        multi = master[master.get("hoa_state_count", 0) > 1].sort_values(
            "hoa_state_count", ascending=False).head(10)
        if len(multi) > 0:
            ms_rows = []
            for _, r in multi.iterrows():
                ms_rows.append({
                    "section": "MULTI-STATE PMCs",
                    "metric": r.get("company_name", ""),
                    "value": f"{int(r['hoa_state_count'])} states, {int(r['hoa_count'])} HOAs",
                })
            parts.append(pd.DataFrame(ms_rows))

    return pd.concat(parts, ignore_index=True, sort=False)


def sheet_e08_platform_mix(pmc, cases):
    """E08: Accounting Platform analysis."""
    col = _COL_MAPS["pmc"].get("acct_platform")
    if not col:
        return pd.DataFrame({"note": ["Accounting Platform column not found"]})

    df = pmc.copy()
    df["_platform"] = df[col].fillna("(blank)").astype(str).str.strip()

    grp = df.groupby("_platform")
    result = grp.size().reset_index(name="pmc_count")
    result = result.sort_values("pmc_count", ascending=False).head(15)

    if "_deposits_best" in df.columns:
        dep = grp["_deposits_best"].agg(total_deposits="sum", median_deposits="median").reset_index()
        dep["total_deposits"] = dep["total_deposits"].apply(fmt_usd)
        dep["median_deposits"] = dep["median_deposits"].apply(fmt_usd)
        result = result.merge(dep, on="_platform")

    if not cases.empty and "_company_norm" in cases.columns and "_name_norm" in df.columns:
        client = cases[~cases["_is_internal"]].copy()
        name_to_plat = df[["_name_norm", "_platform"]].drop_duplicates(subset=["_name_norm"])
        client = client.merge(name_to_plat, left_on="_company_norm", right_on="_name_norm", how="inner")
        if len(client) > 0:
            case_by_plat = client.groupby("_platform").size().reset_index(name="case_count")
            result = result.merge(case_by_plat, on="_platform", how="left")
            result["case_count"] = result["case_count"].fillna(0).astype(int)

    result.rename(columns={"_platform": "platform"}, inplace=True)
    return result.reset_index(drop=True)


def sheet_e09_pod_geography(pmc):
    """E09: Pod-to-state mapping inferred from PMC addresses."""
    col_pod = _COL_MAPS["pmc"].get("pod")
    if not col_pod or "_state" not in pmc.columns:
        return pd.DataFrame({"note": ["Pod or state column not found"]})

    df = pmc.copy()
    df["_pod"] = df[col_pod].fillna("(blank)").astype(str).str.strip()

    ct = pd.crosstab(df["_pod"], df["_state"])
    # Keep only states with >5 PMCs total to reduce noise
    sig_states = ct.sum()[ct.sum() > 5].index.tolist()
    ct = ct[sig_states]
    # Keep top 10 pods by PMC count
    top_pods = df["_pod"].value_counts().head(10).index.tolist()
    ct = ct.loc[ct.index.isin(top_pods)]

    return ct.reset_index().rename(columns={"_pod": "pod"})


def sheet_e10_entity_completeness(master):
    """E10: Data completeness scorecard for each PMC."""
    if master.empty:
        return pd.DataFrame({"note": ["No master table"]})

    checks = {}
    if "company_name" in master.columns:
        checks["has_name"] = master["company_name"].notna() & (master["company_name"] != "")
    if "deposits" in master.columns:
        checks["has_deposits"] = master["deposits"].notna() & (master["deposits"] > 0)
    if "state" in master.columns:
        checks["has_state"] = master["state"].notna() & (master["state"] != "(BLANK)") & (master["state"] != "")
    if "hoa_count" in master.columns:
        checks["has_hoas"] = master["hoa_count"] > 0
    if "case_count" in master.columns:
        checks["has_cases"] = master["case_count"] > 0
    if "relationship_manager" in master.columns:
        checks["has_rm"] = master["relationship_manager"].notna() & (master["relationship_manager"] != "")
    if "rm_last_checkin" in master.columns:
        checks["has_recent_checkin"] = master["rm_last_checkin"].notna()
    if "pod" in master.columns:
        checks["has_pod"] = master["pod"].notna() & (master["pod"] != "")

    total = len(master)
    rows = []
    for field, mask in checks.items():
        n = mask.sum()
        rows.append({"field": field, "count": f"{n:,}", "pct": pct(n, total)})

    result = pd.DataFrame(rows)

    # Fully complete PMCs (all checks pass)
    if checks:
        all_complete = pd.concat(checks.values(), axis=1).all(axis=1).sum()
        result = pd.concat([
            result,
            pd.DataFrame({"field": ["ALL FIELDS COMPLETE"], "count": [f"{all_complete:,}"],
                          "pct": [pct(all_complete, total)]})
        ], ignore_index=True)

    return result


def sheet_e11_story_numbers(pmc, hoa, cases, emails, master):
    """E11: Key numbers for the unified HTML story."""
    rows = []
    def _add(cat, metric, value):
        rows.append({"category": cat, "metric": metric, "value": value})

    # PMC
    _add("PMC Universe", "total_pmcs", f"{len(pmc):,}")
    if "_deposits_best" in pmc.columns:
        deps = pmc["_deposits_best"].dropna()
        _add("PMC Universe", "total_deposits", fmt_usd(deps.sum()))
        _add("PMC Universe", "median_deposit", fmt_usd(deps.median()))
        _add("PMC Universe", "pmcs_with_deposits", f"{len(deps[deps>0]):,}")
    if "_state" in pmc.columns:
        _add("PMC Universe", "distinct_states", f"{pmc['_state'].nunique()}")
    if "_type" in pmc.columns:
        _add("PMC Universe", "pct_management_company", pct((pmc["_type"]=="Management Company").sum(), len(pmc)))

    # HOA
    _add("HOA Universe", "total_hoas", f"{len(hoa):,}")
    if "_parent_pmc_id" in hoa.columns:
        linked = hoa["_parent_pmc_id"].notna().sum()
        _add("HOA Universe", "linked_to_pmc", f"{linked:,} ({pct(linked, len(hoa))})")
    if "_deposits_hoa" in hoa.columns:
        _add("HOA Universe", "total_hoa_deposits", fmt_usd(hoa["_deposits_hoa"].dropna().sum()))

    # Cases
    if not cases.empty:
        client = cases[~cases["_is_internal"]]
        _add("Cases (3mo)", "total_cases", f"{len(cases):,}")
        _add("Cases (3mo)", "client_cases", f"{len(client):,}")
        _add("Cases (3mo)", "internal_cases", f"{cases['_is_internal'].sum():,}")
        if "_hours" in client.columns:
            _add("Cases (3mo)", "client_median_hours", f"{client['_hours'].median():.1f}")
        if "_is_resolved" in client.columns:
            unres = (~client["_is_resolved"]).sum()
            _add("Cases (3mo)", "client_unresolved", f"{unres:,} ({pct(unres, len(client))})")

    # Emails
    if not emails.empty:
        _add("Emails (1day)", "total_emails", f"{len(emails):,}")
        if "_case_ref" in emails.columns:
            linked = emails["_case_ref"].notna().sum()
            _add("Emails (1day)", "linked_to_case", f"{linked:,} ({pct(linked, len(emails))})")

    # Joins
    if not master.empty:
        with_cases = (master.get("case_count", 0) > 0).sum()
        _add("Joins", "pmcs_with_cases", f"{with_cases:,} of {len(master):,}")
        with_hoas = (master.get("hoa_count", 0) > 0).sum()
        _add("Joins", "pmcs_with_hoas", f"{with_hoas:,} of {len(master):,}")

    return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════

def main():
    start = datetime.datetime.now()
    log(f"WAB Entity Deep Dive — {start.strftime('%Y-%m-%d %H:%M:%S')}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    pmc_raw   = read_file(PMC_FILE, "PMC")
    hoa_raw   = read_file(HOA_FILE, "HOA")
    cases_raw = read_file(CASE_FILE, "Cases")
    email_raw = read_file(EMAIL_FILE, "Emails")

    log("\n--- Preparing data ---")
    pmc    = prepare_pmc(pmc_raw) if not pmc_raw.empty else pd.DataFrame()
    hoa    = prepare_hoa(hoa_raw) if not hoa_raw.empty else pd.DataFrame()
    cases  = prepare_cases_light(cases_raw) if not cases_raw.empty else pd.DataFrame()
    emails = prepare_emails_light(email_raw) if not email_raw.empty else pd.DataFrame()

    master = build_pmc_master(pmc, hoa, cases, emails)

    log("\n--- Building sheets ---")
    sheets = OrderedDict()

    log("  E01 Deposit Concentration")
    sheets["E01_DepositConcentr"] = sheet_e01_deposit_concentration(pmc)

    log("  E02 Top PMCs")
    sheets["E02_TopPMCs"] = sheet_e02_top_pmcs(master)

    log("  E03 Friction-Value Map")
    sheets["E03_FrictionValue"] = sheet_e03_friction_value(master)

    log("  E04 State Profile")
    sheets["E04_StateProfile"] = sheet_e04_state_profile(pmc, hoa, cases)

    log("  E05 RM Coverage")
    sheets["E05_RM_Coverage"] = sheet_e05_rm_coverage(pmc)

    log("  E06 Company Type")
    sheets["E06_CompanyType"] = sheet_e06_company_type(pmc, cases)

    log("  E07 Hierarchy Depth")
    sheets["E07_HierarchyDepth"] = sheet_e07_hierarchy_depth(master, hoa, pmc)

    log("  E08 Platform Mix")
    sheets["E08_PlatformMix"] = sheet_e08_platform_mix(pmc, cases)

    log("  E09 Pod Geography")
    sheets["E09_PodGeography"] = sheet_e09_pod_geography(pmc)

    log("  E10 Entity Completeness")
    sheets["E10_Completeness"] = sheet_e10_entity_completeness(master)

    log("  E11 Story Numbers")
    sheets["E11_StoryNumbers"] = sheet_e11_story_numbers(pmc, hoa, cases, emails, master)

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
        "# WAB Entity & Relationship Deep Dive",
        f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}\n",
        "## Files",
        f"- PMC: {len(pmc_raw):,} rows",
        f"- HOA: {len(hoa_raw):,} rows",
        f"- Cases: {len(cases_raw):,} rows",
        f"- Emails: {len(email_raw):,} rows",
        f"- PMC Master (joined): {len(master)} rows",
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
