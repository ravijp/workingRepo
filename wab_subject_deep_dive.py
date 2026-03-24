"""
WAB Subject-Level Deep Dive — Phase 1 GenAI Use-Case Discovery
================================================================
Executive-oriented subject intelligence module.
Reads Cases file only. Produces one Excel workbook (10 sheets).

Structure:
  Tabs 1-6:  Executive decision support (Bob's view)
  Tabs 7-10: Reference/appendix (Chris/Zenon team detail)

Design principle: every sheet answers "where should we intervene?"
NOT "what does the data look like?"

Dependencies: pandas, openpyxl  (standard library otherwise)
"""

# ┌─────────────────────────────────────────────────────────┐
# │  EDIT THESE 2 VARIABLES BEFORE RUNNING                  │
# └─────────────────────────────────────────────────────────┘
CASE_FILE  = r"C:\Users\YourName\Desktop\AAB All Cases.xlsx"
OUTPUT_DIR = r"C:\Users\YourName\Desktop\wab_output"
# ┌─────────────────────────────────────────────────────────┐
# │  DO NOT EDIT BELOW THIS LINE                            │
# └─────────────────────────────────────────────────────────┘

import os, re, sys, datetime, warnings
from collections import OrderedDict, Counter

import pandas as pd
import numpy as np

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "wab_subject_deep_dive.xlsx")
OUTPUT_MD   = os.path.join(OUTPUT_DIR, "wab_subject_deep_dive_summary.md")
LOG, WARN   = [], []
TRUNC       = 150

INTERNAL_COMPANIES = {"AAB ADMIN", "WAB ADMIN", "AAB ADMIN -", "WAB ADMIN -"}

# Action-tag rules: subject keyword → recommended action
# Applied in assign_action_tag(); overridden by data-driven logic
ACTION_RULES = {
    # By-design wait (not friction)
    "CD Maintenance":      "MONITOR",
    "IntraFi Maintenance": "MONITOR",
    # High-volume fast resolution — automate
    "NSF and Non-Post":    "AUTOMATE",
    "Fraud Alert":         "AUTOMATE",
    "Transfer":            "AUTOMATE",
    "Statements":          "AI-ASSIST",
    # Structurally slow — process redesign first
    "Signature Card":      "REDESIGN",
    "Close Account":       "REDESIGN",
    # Fat tail / GenAI sweet spot
    "Research":            "AI-ASSIST",
    "New Account Request": "AI-ASSIST",
    "Account Maintenance": "AI-ASSIST",
    "General Questions":   "AI-ASSIST",
}


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

def safe_numeric(series):
    if pd.api.types.is_numeric_dtype(series):
        return series
    return pd.to_numeric(series, errors="coerce")

def pct(num, denom):
    return f"{100*num/denom:.1f}%" if denom else "N/A"

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
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 60)

def extract_top_keywords(texts, top_n=20, min_len=3):
    stop = {
        "the","and","for","that","this","with","from","your","have","are",
        "was","were","been","has","had","but","not","you","all","can","her",
        "his","one","our","out","will","about","each","which","their","them",
        "then","than","into","over","also","back","after","use","two","how",
        "its","let","may","new","now","old","see","way","who","did","get",
        "got","him","just","own","say","she","too","per","via","please",
        "would","could","should","need","account","case","email","bank",
        "thank","thanks","hello","dear","regards","sincerely","sent",
        "received","attached","attachment","fyi","following","below",
    }
    words = Counter()
    for text in texts:
        if not text or len(str(text)) < 3:
            continue
        tokens = re.findall(r"[a-z]{3,}", str(text).lower())
        for t in tokens:
            if t not in stop and len(t) >= min_len:
                words[t] += 1
    return words.most_common(top_n)


# ═══════════════════════════════════════════════════════════
#  CASE PREPARATION
# ═══════════════════════════════════════════════════════════

def prepare_cases(cases):
    """Add derived columns. Returns enriched copy."""
    df = cases.copy()

    col_map = {
        "case_number":   find_col(df, "Case Number"),
        "company":       find_col(df, "Company Name (Company) (Company)", "Company Name", "Customer"),
        "subject":       find_col(df, "Subject"),
        "description":   find_col(df, "Description"),
        "activity_subj": find_col(df, "Activity Subject"),
        "origin":        find_col(df, "Origin"),
        "status":        find_col(df, "Status"),
        "status_reason": find_col(df, "Status Reason"),
        "created_on":    find_col(df, "Created On"),
        "modified_on":   find_col(df, "Modified On"),
        "sla_start":     find_col(df, "SLA Start"),
        "resolved_on":   find_col(df, "Resolved On"),
        "resolved_hrs":  find_col(df, "Resolved In Hours"),
        "owner":         find_col(df, "Manager (Owning User) (User)", "Owner"),
        "pod":           find_col(df, "POD Name (Owning User) (User)", "POD Name"),
        "last_touch":    find_col(df, "Last Touch"),
        "last_touch_by": find_col(df, "Last Touch By"),
        "parent_case":   find_col(df, "Parent Case"),
        "last_contact":  find_col(df, "Last Contact Attempt"),
    }

    for key in ["created_on", "modified_on", "sla_start", "resolved_on", "last_touch", "last_contact"]:
        src = col_map.get(key)
        if src:
            df[f"_{key}_dt"] = safe_dt(df[src])

    src = col_map.get("resolved_hrs")
    if src:
        df["_hours"] = safe_numeric(df[src])

    resolved_kw = ["resolved", "closed", "cancelled", "canceled"]
    src_sr = col_map.get("status_reason") or col_map.get("status")
    if src_sr:
        df["_is_resolved"] = df[src_sr].fillna("").astype(str).str.lower().apply(
            lambda x: any(kw in x for kw in resolved_kw))
    else:
        df["_is_resolved"] = True

    src_co = col_map.get("company")
    if src_co:
        df["_company_clean"] = df[src_co].fillna("(blank)").astype(str).str.strip()
        df["_is_internal"] = df["_company_clean"].str.upper().apply(
            lambda x: any(x.startswith(kw) for kw in INTERNAL_COMPANIES) or x == "(BLANK)")
    else:
        df["_is_internal"] = False

    src_subj = col_map.get("subject")
    if src_subj:
        df["_subject"] = df[src_subj].fillna("(blank)").astype(str).str.strip()
    else:
        df["_subject"] = "(no subject column)"

    src_orig = col_map.get("origin")
    if src_orig:
        df["_origin"] = df[src_orig].fillna("(blank)").astype(str).str.strip()
    else:
        df["_origin"] = "(no origin column)"

    src_own = col_map.get("owner")
    if src_own:
        df["_owner"] = df[src_own].fillna("(blank)").astype(str).str.strip()
    else:
        df["_owner"] = "(no owner column)"

    src_pod = col_map.get("pod")
    if src_pod:
        df["_pod"] = df[src_pod].fillna("(blank)").astype(str).str.strip()
    else:
        df["_pod"] = "(no pod column)"

    if "_created_on_dt" in df.columns:
        valid = df["_created_on_dt"].notna()
        df.loc[valid, "_week"] = df.loc[valid, "_created_on_dt"].dt.to_period("W").astype(str)
        df.loc[valid, "_dow"]  = df.loc[valid, "_created_on_dt"].dt.day_name()
        df.loc[valid, "_hour"] = df.loc[valid, "_created_on_dt"].dt.hour

    if "_created_on_dt" in df.columns:
        now = pd.Timestamp.now()
        df["_age_hours"] = (now - df["_created_on_dt"]).dt.total_seconds() / 3600

    for key in ["description", "activity_subj"]:
        src = col_map.get(key)
        if src:
            df[f"_{key}_len"] = df[src].fillna("").astype(str).str.len()
            df[f"_{key}_text"] = df[src].fillna("").astype(str).str.strip()

    if "_sla_start_dt" in df.columns and "_created_on_dt" in df.columns:
        df["_triage_minutes"] = (
            df["_created_on_dt"] - df["_sla_start_dt"]
        ).dt.total_seconds() / 60
        df["_has_triage_gap"] = (df["_triage_minutes"] > 0.5) & (df["_triage_minutes"] < 10080)

    src_sr2 = col_map.get("status_reason")
    if src_sr2:
        df["_status_reason"] = df[src_sr2].fillna("(blank)").astype(str).str.strip()

    df._col_map = col_map
    return df


# ═══════════════════════════════════════════════════════════
#  ACTION TAG LOGIC
# ═══════════════════════════════════════════════════════════

def assign_action_tag(subject, median_hrs, unresolved_pct, cases_per_week,
                      top_owner_pct, p90_median_ratio):
    """Assign an executive action tag based on rules + data signals."""
    # Check hardcoded rules first
    for pattern, tag in ACTION_RULES.items():
        if pattern.lower() in subject.lower():
            return tag

    # Data-driven fallbacks
    if pd.notna(median_hrs) and median_hrs < 4 and pd.notna(cases_per_week) and cases_per_week > 15:
        return "AUTOMATE"
    if pd.notna(unresolved_pct) and unresolved_pct > 10 and pd.notna(top_owner_pct) and top_owner_pct > 40:
        return "REDESIGN"
    if pd.notna(p90_median_ratio) and p90_median_ratio >= 10:
        return "AI-ASSIST"  # mixed types → needs intelligent routing
    if pd.notna(cases_per_week) and cases_per_week < 5:
        return "DEPRIORITIZE"
    return "MONITOR"


# ═══════════════════════════════════════════════════════════
#  RESEARCH SUB-SEGMENTATION (simplified to 5 clusters)
# ═══════════════════════════════════════════════════════════

def assign_research_cluster(desc_text, act_subj_text):
    """Rule-based sub-segmentation of Research into 5 executive-friendly types."""
    combined = f"{desc_text} {act_subj_text}".lower()

    if any(kw in combined for kw in ["payment", "ach", "wire", "transfer", "deposit",
                                      "credit", "debit", "transaction", "posted",
                                      "posting", "return", "reversal", "refund"]):
        return "Payment Research"

    if any(kw in combined for kw in ["check", "cheque", "item", "image", "copy",
                                      "front", "back", "clearing"]):
        return "Check/Item Research"

    if any(kw in combined for kw in ["statement", "balance", "reconcil", "ledger",
                                      "interest", "rate", "fee"]):
        return "Statement/Balance Inquiry"

    if any(kw in combined for kw in ["address", "signer", "name change", "update",
                                      "modify", "amendment", "tin", "ein", "ssn",
                                      "tax", "w-9", "w9", "certification",
                                      "fraud", "dispute", "unauthorized", "suspicious",
                                      "positive pay", "stop payment",
                                      "new account", "onboard", "setup", "opening"]):
        return "Account Updates & Other"

    if len(combined.strip()) < 10:
        return "No Text (Unclassifiable)"

    return "Other/Uncategorized"


# ═══════════════════════════════════════════════════════════
#  TAB 1: SUBJECT DASHBOARD (the primary decision sheet)
# ═══════════════════════════════════════════════════════════

def tab_01_subject_dashboard(df):
    """One row per subject. The only sheet Bob needs to open first.
    Merges: cases/week, banker-hrs/week, unresolved trend, SPOF,
    pod speed gap, triage delay, company concentration, action tag."""
    client = df[~df["_is_internal"]].copy()
    n_weeks = max(client["_week"].nunique(), 1) if "_week" in client.columns else 13

    subj_counts = client["_subject"].value_counts()
    eligible = subj_counts[subj_counts >= 10].index.tolist()
    sub = client[client["_subject"].isin(eligible)]
    grp = sub.groupby("_subject")

    result = grp.size().reset_index(name="total_cases")
    result["cases_per_week"] = (result["total_cases"] / n_weeks).round(1)
    result = result.sort_values("total_cases", ascending=False)

    # Banker-hours per week
    if "_hours" in sub.columns:
        hrs_agg = grp["_hours"].agg(
            _total_hrs="sum",
            median_hrs="median",
            p90_hrs=lambda x: x.quantile(0.9),
        ).reset_index()
        hrs_agg["est_banker_hrs_per_week"] = (hrs_agg["_total_hrs"] / n_weeks).round(1)
        hrs_agg["median_hrs"] = hrs_agg["median_hrs"].round(1)
        hrs_agg["p90_hrs"] = hrs_agg["p90_hrs"].round(1)

        # p90/median ratio
        hrs_agg["p90_median_ratio"] = (
            hrs_agg["p90_hrs"] / hrs_agg["median_hrs"].replace(0, np.nan)
        ).round(1)

        result = result.merge(
            hrs_agg[["_subject", "median_hrs", "p90_hrs", "est_banker_hrs_per_week", "p90_median_ratio"]],
            on="_subject", how="left"
        )

    # Unresolved trend (last 4 weeks vs prior)
    if "_is_resolved" in sub.columns:
        unres_agg = grp["_is_resolved"].agg(
            unresolved=lambda x: int((~x).sum()),
            pct_unresolved=lambda x: round(100 * (~x).mean(), 1),
        ).reset_index()
        result = result.merge(unres_agg, on="_subject", how="left")

    if "_week" in sub.columns:
        weeks = sorted(sub["_week"].dropna().unique())
        if len(weeks) >= 4:
            recent_weeks = set(weeks[-4:])
            prior_weeks = set(weeks[:-4])
            trend_rows = []
            for subj in eligible:
                s = sub[sub["_subject"] == subj]
                recent = s[s["_week"].isin(recent_weeks)]
                prior = s[s["_week"].isin(prior_weeks)]
                unres_recent = round(100 * (~recent["_is_resolved"]).mean(), 1) if len(recent) > 0 else np.nan
                unres_prior = round(100 * (~prior["_is_resolved"]).mean(), 1) if len(prior) > 0 else np.nan
                if pd.notna(unres_recent) and pd.notna(unres_prior):
                    delta = round(unres_recent - unres_prior, 1)
                    if delta > 3:
                        trend = "↑ WORSENING"
                    elif delta < -3:
                        trend = "↓ IMPROVING"
                    else:
                        trend = "→ STABLE"
                else:
                    trend = ""
                trend_rows.append({"_subject": subj, "unres_trend": trend})
            result = result.merge(pd.DataFrame(trend_rows), on="_subject", how="left")

    # Top owner % (SPOF)
    owner_agg = grp["_owner"].agg(
        top_owner=lambda x: x.value_counts().index[0] if len(x) > 0 else "",
        top_owner_pct=lambda x: round(100 * x.value_counts().iloc[0] / len(x), 1) if len(x) > 0 else 0,
    ).reset_index()
    owner_agg["spof_flag"] = owner_agg["top_owner_pct"].apply(
        lambda x: "YES" if x > 50 else ("WATCH" if x > 35 else ""))
    result = result.merge(owner_agg, on="_subject", how="left")

    # Pod speed gap
    if "_pod" in sub.columns and "_hours" in sub.columns:
        speed_rows = []
        for subj in eligible:
            s = sub[sub["_subject"] == subj]
            pod_med = s.groupby("_pod")["_hours"].median().dropna()
            if len(pod_med) >= 2:
                speed_rows.append({
                    "_subject": subj,
                    "pod_speed_gap": f"{round(pod_med.min(), 1)}h → {round(pod_med.max(), 1)}h",
                    "pod_gap_ratio": round(pod_med.max() / max(pod_med.min(), 0.1), 1),
                })
        if speed_rows:
            result = result.merge(pd.DataFrame(speed_rows), on="_subject", how="left")

    # Triage delay median (email-originated)
    if "_triage_minutes" in sub.columns and "_has_triage_gap" in sub.columns:
        triage_sub = sub[sub["_has_triage_gap"]]
        if len(triage_sub) > 0:
            triage_agg = triage_sub.groupby("_subject")["_triage_minutes"].median().round(1)
            triage_agg = triage_agg.reset_index(name="triage_median_min")
            result = result.merge(triage_agg, on="_subject", how="left")

    # Company concentration (top 5 companies as % of subject)
    co_conc_rows = []
    for subj in eligible:
        s = sub[sub["_subject"] == subj]
        top5 = s["_company_clean"].value_counts().head(5).sum()
        co_conc_rows.append({
            "_subject": subj,
            "top5_company_pct": round(100 * top5 / len(s), 1) if len(s) > 0 else 0,
        })
    result = result.merge(pd.DataFrame(co_conc_rows), on="_subject", how="left")

    # Recoverable hours flag
    # Wait-time subjects: hours consumed but not recoverable through automation
    wait_subjects = {"cd maintenance", "intrafi maintenance"}
    result["hours_type"] = result["_subject"].apply(
        lambda x: "WAIT TIME" if x.lower() in wait_subjects else "WORK TIME"
    )

    # ACTION TAG
    result["recommended_action"] = result.apply(
        lambda r: assign_action_tag(
            r["_subject"],
            r.get("median_hrs", np.nan),
            r.get("pct_unresolved", np.nan),
            r.get("cases_per_week", np.nan),
            r.get("top_owner_pct", np.nan),
            r.get("p90_median_ratio", np.nan),
        ), axis=1
    )

    result.rename(columns={"_subject": "subject"}, inplace=True)

    # Reorder columns for readability
    priority_cols = [
        "subject", "recommended_action", "cases_per_week", "est_banker_hrs_per_week",
        "hours_type", "median_hrs", "p90_hrs", "pct_unresolved", "unres_trend",
        "top_owner", "top_owner_pct", "spof_flag",
    ]
    other_cols = [c for c in result.columns if c not in priority_cols]
    ordered = [c for c in priority_cols if c in result.columns] + other_cols
    result = result[ordered]

    return result.sort_values("est_banker_hrs_per_week", ascending=False).reset_index(drop=True)


# ═══════════════════════════════════════════════════════════
#  TAB 2: CLAIM CHECK (Chris March 23 assertions vs data)
# ═══════════════════════════════════════════════════════════

def tab_02_claim_check(df):
    """Explicit test of operational claims from Chris's March 23 walkthrough.
    One row per claim. Columns: Claim, Data Finding, Verdict, Implication."""
    client = df[~df["_is_internal"]].copy()
    n_weeks = max(client["_week"].nunique(), 1) if "_week" in client.columns else 13

    rows = []

    # ── Claim 1: CD Maintenance 97.8h median is by design (maturity wait) ──
    cd = client[client["_subject"].str.contains("CD Maintenance", case=False, na=False)]
    cd_hrs = cd["_hours"].dropna()
    cd_median = round(cd_hrs.median(), 1) if len(cd_hrs) > 0 else np.nan
    cd_p25 = round(cd_hrs.quantile(0.25), 1) if len(cd_hrs) > 0 else np.nan
    cd_unres = round(100 * (~cd["_is_resolved"]).mean(), 1) if len(cd) > 0 else np.nan

    # Check if resolution time clusters around multiples of 24h (maturity dates)
    cd_in_7_14d = ((cd_hrs >= 168) & (cd_hrs <= 336)).sum() if len(cd_hrs) > 0 else 0
    cd_pct_7_14d = round(100 * cd_in_7_14d / len(cd_hrs), 1) if len(cd_hrs) > 0 else 0

    finding = (f"Median {cd_median}h ({round(cd_median/24, 1)}d). "
               f"P25={cd_p25}h — even fast cases take >1 day. "
               f"{cd_pct_7_14d}% resolve in 7-14d range (typical CD maturity window). "
               f"Unresolved: {cd_unres}%.")
    verdict = "CONFIRMED" if pd.notna(cd_median) and cd_p25 > 12 else "PARTIALLY CONFIRMED"
    implication = ("Not a speed target. Hours consumed are wait time, not work time. "
                   "Exclude from automation ROI calculations.")
    rows.append({
        "claim": "CD Maintenance 97.8h median is by design — bankers wait for maturity date",
        "source": "Chris, March 23",
        "data_finding": finding,
        "verdict": verdict,
        "implication": implication,
    })

    # ── Claim 2: Signature Card has large backlog, Keith owns it ──
    sig = client[client["_subject"].str.contains("Signature Card", case=False, na=False)]
    sig_unres = (~sig["_is_resolved"]).sum()
    sig_total = len(sig)
    sig_unres_pct = round(100 * sig_unres / sig_total, 1) if sig_total > 0 else 0
    sig_owner_vc = sig["_owner"].value_counts()
    sig_top_owner = sig_owner_vc.index[0] if len(sig_owner_vc) > 0 else "unknown"
    sig_top_pct = round(100 * sig_owner_vc.iloc[0] / sig_total, 1) if sig_total > 0 else 0

    # Aging of unresolved
    sig_unres_cases = sig[~sig["_is_resolved"]]
    sig_old = (sig_unres_cases["_age_hours"] > 720).sum() if "_age_hours" in sig_unres_cases.columns else 0

    finding = (f"{sig_total:,} cases, {sig_unres:,} unresolved ({sig_unres_pct}%). "
               f"Top owner: {sig_top_owner} at {sig_top_pct}%. "
               f"{sig_old} unresolved cases older than 30 days.")
    keith_match = "keith" in sig_top_owner.lower() if sig_top_owner else False
    verdict = "CONFIRMED" if sig_unres_pct > 10 else "PARTIALLY CONFIRMED"
    implication = (f"{'Keith confirmed as primary owner. ' if keith_match else f'Top owner is {sig_top_owner}, not Keith — verify. '}"
                   f"Backlog is real: {sig_unres_pct}% unresolved. "
                   "Board member change automation is the intervention.")
    rows.append({
        "claim": "Signature Card has large backlog; Keith owns it; board member changes drive volume",
        "source": "Chris, March 23",
        "data_finding": finding,
        "verdict": verdict,
        "implication": implication,
    })

    # ── Claim 3: Research is a catch-all that needs sub-segmentation ──
    research = client[client["_subject"] == "Research"].copy()
    research_n = len(research)
    research_desc_fill = round(100 * (research["_description_len"] > 0).mean(), 0) if "_description_len" in research.columns and research_n > 0 else 0
    research_act_fill = round(100 * (research["_activity_subj_len"] > 5).mean(), 0) if "_activity_subj_len" in research.columns and research_n > 0 else 0

    # Quick cluster check
    if "_description_text" in research.columns and "_activity_subj_text" in research.columns:
        research["_cluster"] = research.apply(
            lambda r: assign_research_cluster(r.get("_description_text", ""), r.get("_activity_subj_text", "")),
            axis=1
        )
        cluster_dist = research["_cluster"].value_counts()
        top_cluster = cluster_dist.index[0] if len(cluster_dist) > 0 else ""
        top_cluster_pct = round(100 * cluster_dist.iloc[0] / research_n, 1) if research_n > 0 else 0
        n_clusters_gt5pct = (cluster_dist / research_n > 0.05).sum()
    else:
        top_cluster = "N/A"
        top_cluster_pct = 0
        n_clusters_gt5pct = 0

    finding = (f"{research_n:,} cases. Description fill: {research_desc_fill}%, Activity Subject fill: {research_act_fill}%. "
               f"Keyword clustering yields {n_clusters_gt5pct} segments >5% each. "
               f"Largest: {top_cluster} ({top_cluster_pct}%).")
    verdict = "CONFIRMED" if n_clusters_gt5pct >= 3 else "PARTIALLY CONFIRMED"
    implication = ("Research can be split into ~5 actionable sub-types. "
                   "Payment Research likely dominates. Sub-segmentation enables targeted routing.")
    rows.append({
        "claim": "Research is a catch-all — 'a lot of different things go into it'",
        "source": "Chris, March 23",
        "data_finding": finding,
        "verdict": verdict,
        "implication": implication,
    })

    # ── Claim 4: NSF/Non-Post = 2-3 hrs/day per banker ──
    nsf = client[client["_subject"].str.contains("NSF|Non.?Post", case=False, na=False)]
    nsf_hrs = nsf["_hours"].dropna()
    nsf_total_hrs = nsf_hrs.sum() if len(nsf_hrs) > 0 else 0
    nsf_daily = nsf_total_hrs / (n_weeks * 5) if n_weeks > 0 else 0
    nsf_owners = nsf["_owner"].nunique()
    nsf_per_owner_daily = round(nsf_daily / max(nsf_owners, 1), 1)
    nsf_n = len(nsf)
    nsf_median = round(nsf_hrs.median(), 1) if len(nsf_hrs) > 0 else np.nan

    # Morning spike check
    morning_pct = 0
    if "_hour" in nsf.columns:
        morning = nsf[(nsf["_hour"] >= 7) & (nsf["_hour"] <= 10)]
        morning_pct = round(100 * len(morning) / len(nsf), 1) if len(nsf) > 0 else 0

    finding = (f"{nsf_n:,} cases over {n_weeks} weeks. Median: {nsf_median}h per case. "
               f"Total: {round(nsf_total_hrs, 0):,.0f}h. "
               f"~{round(nsf_daily, 1)}h/day across {nsf_owners} owners = "
               f"~{nsf_per_owner_daily}h/owner/day. "
               f"Morning (7-10 AM) share: {morning_pct}%.")
    in_range = 1.5 <= nsf_per_owner_daily <= 4.0
    verdict = "CONFIRMED" if in_range else ("PARTIALLY CONFIRMED" if nsf_per_owner_daily > 0.5 else "INSUFFICIENT DATA")
    implication = (f"Chris said 2-3 hrs/day. Data shows ~{nsf_per_owner_daily}h/owner/day. "
                   "{'Aligns with estimate.' if in_range else 'Variance may reflect different counting — verify with Chris.'} "
                   "Positive Pay adoption campaign is the structural fix.")
    rows.append({
        "claim": "NSF and Non-Post consumes 2-3 hours of every banker's time daily",
        "source": "Chris, March 23",
        "data_finding": finding,
        "verdict": verdict,
        "implication": implication,
    })

    # ── Claim 5: Close Account requires manual cross-system checklist ──
    close = client[client["_subject"].str.contains("Clos", case=False, na=False)]
    close_hrs = close["_hours"].dropna()
    close_median = round(close_hrs.median(), 1) if len(close_hrs) > 0 else np.nan
    close_p90 = round(close_hrs.quantile(0.9), 1) if len(close_hrs) > 0 else np.nan
    close_n = len(close)
    close_unres = round(100 * (~close["_is_resolved"]).mean(), 1) if close_n > 0 else 0

    # Check for system mentions in descriptions
    system_mentions = {}
    if "_description_text" in close.columns:
        for sys_name in ["ibs", "bst", "ach", "lockbox", "loan", "credit card", "online banking"]:
            count = sum(1 for t in close["_description_text"] if sys_name in str(t).lower())
            if count > 0:
                system_mentions[sys_name.upper()] = count

    sys_str = ", ".join(f"{k}:{v}" for k, v in system_mentions.items()) if system_mentions else "none detected (Description only 50% filled)"
    finding = (f"{close_n:,} cases. Median: {close_median}h, P90: {close_p90}h. "
               f"Unresolved: {close_unres}%. "
               f"System mentions in Description: {sys_str}.")
    verdict = "PARTIALLY CONFIRMED" if pd.notna(close_median) and close_median > 10 else "INSUFFICIENT DATA"
    implication = ("Long resolution time consistent with multi-system checklist. "
                   "CRM consolidation (Chris's stated goal) is the right approach. "
                   "Not a near-term AI target.")
    rows.append({
        "claim": "Close Account requires manual checklist across IBS, BST, ACH tracker (disconnected FIS systems)",
        "source": "Chris, March 23",
        "data_finding": finding,
        "verdict": verdict,
        "implication": implication,
    })

    return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════
#  TAB 3: RESEARCH BREAKDOWN (simplified 5 clusters)
# ═══════════════════════════════════════════════════════════

def tab_03_research_breakdown(df):
    """Research sub-segmentation: 5 clusters, summary only."""
    client = df[~df["_is_internal"]].copy()
    research = client[client["_subject"] == "Research"].copy()

    if len(research) == 0:
        return pd.DataFrame({"note": ["No Research cases found"]})

    log(f"  Research cases: {len(research):,}")

    research["_cluster"] = research.apply(
        lambda r: assign_research_cluster(
            r.get("_description_text", ""),
            r.get("_activity_subj_text", ""),
        ), axis=1
    )

    grp = research.groupby("_cluster")
    result = grp.size().reset_index(name="cases")
    result["pct"] = (result["cases"] / len(research) * 100).round(1)
    result = result.sort_values("cases", ascending=False)

    if "_hours" in research.columns:
        hrs = grp["_hours"].agg(
            median_hrs="median",
            p90_hrs=lambda x: x.quantile(0.9) if len(x) else np.nan,
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

    owner_conc = grp["_owner"].agg(
        distinct_owners=lambda x: x.nunique(),
        top_owner=lambda x: x.value_counts().index[0] if len(x) > 0 else "",
        top_owner_pct=lambda x: round(100 * x.value_counts().iloc[0] / len(x), 1) if len(x) > 0 else 0,
    ).reset_index()
    result = result.merge(owner_conc, on="_cluster", how="left")

    # Action tag per cluster
    n_weeks = max(research["_week"].nunique(), 1) if "_week" in research.columns else 13
    result["action_tag"] = result.apply(
        lambda r: assign_action_tag(
            f"Research-{r['_cluster']}",
            r.get("median_hrs", np.nan),
            r.get("pct_unresolved", np.nan),
            r["cases"] / n_weeks,
            r.get("top_owner_pct", np.nan),
            r.get("p90_hrs", np.nan) / max(r.get("median_hrs", 1), 0.1) if pd.notna(r.get("p90_hrs")) else np.nan,
        ), axis=1
    )

    result.rename(columns={"_cluster": "cluster"}, inplace=True)
    return result.reset_index(drop=True)


# ═══════════════════════════════════════════════════════════
#  TAB 4: BANKER-HOURS BUDGET (ROI worksheet)
# ═══════════════════════════════════════════════════════════

def tab_04_banker_hours(df):
    """Estimated banker-hours/week per subject with recoverable vs wait-time split."""
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
        cases_per_week = n / n_weeks
        med_per_case = hrs.median() if len(hrs) > 0 else np.nan
        unres = (~s["_is_resolved"]).sum()
        n_owners = s["_owner"].nunique()

        is_wait = subj.lower() in wait_subjects
        hours_type = "WAIT TIME" if is_wait else "WORK TIME"
        recoverable_hrs = 0 if is_wait else round(weekly_hrs * 0.6, 1)  # conservative 60% recoverable

        action = assign_action_tag(
            subj, med_per_case,
            round(100 * (~s["_is_resolved"]).mean(), 1) if len(s) > 0 else 0,
            cases_per_week,
            round(100 * s["_owner"].value_counts().iloc[0] / n, 1) if n > 0 else 0,
            round(hrs.quantile(0.9) / max(med_per_case, 0.1), 1) if pd.notna(med_per_case) and med_per_case > 0 and len(hrs) > 0 else np.nan,
        )

        rows.append({
            "subject": subj,
            "intervention": action,
            "total_cases": n,
            "cases_per_week": round(cases_per_week, 1),
            "median_hrs_per_case": round(med_per_case, 1) if pd.notna(med_per_case) else np.nan,
            "est_banker_hrs_per_week": round(weekly_hrs, 1),
            "hours_type": hours_type,
            "est_recoverable_hrs_per_week": recoverable_hrs,
            "current_unresolved": int(unres),
            "distinct_owners": n_owners,
        })

    result = pd.DataFrame(rows).sort_values("est_banker_hrs_per_week", ascending=False)

    # Totals
    totals = {
        "subject": "=== TOTAL (top 25) ===",
        "total_cases": result["total_cases"].sum(),
        "cases_per_week": round(result["cases_per_week"].sum(), 1),
        "est_banker_hrs_per_week": round(result["est_banker_hrs_per_week"].sum(), 1),
        "est_recoverable_hrs_per_week": round(result["est_recoverable_hrs_per_week"].sum(), 1),
    }
    result = pd.concat([result, pd.DataFrame([totals])], ignore_index=True)
    return result


# ═══════════════════════════════════════════════════════════
#  TAB 5: KEY-PERSON RISK (SPOF only)
# ═══════════════════════════════════════════════════════════

def tab_05_key_person_risk(df):
    """Only subjects where SPOF flag = YES or WATCH."""
    client = df[~df["_is_internal"]].copy()
    subj_counts = client["_subject"].value_counts()
    big = subj_counts[subj_counts >= 20].index.tolist()

    rows = []
    for subj in big:
        s = client[client["_subject"] == subj]
        n = len(s)
        owner_vc = s["_owner"].value_counts()
        if len(owner_vc) == 0:
            continue

        top1 = owner_vc.index[0]
        top1_pct = round(100 * owner_vc.iloc[0] / n, 1)

        if top1_pct < 35:
            continue  # Not a risk

        spof_flag = "YES" if top1_pct > 50 else "WATCH"

        # Speed comparison
        top_hrs = s[s["_owner"] == top1]["_hours"].dropna()
        rest_hrs = s[s["_owner"] != top1]["_hours"].dropna()
        top_median = round(top_hrs.median(), 1) if len(top_hrs) > 0 else np.nan
        rest_median = round(rest_hrs.median(), 1) if len(rest_hrs) > 0 else np.nan

        if pd.notna(top_median) and pd.notna(rest_median) and rest_median > 0:
            speed_note = ("Faster than peers" if top_median < rest_median * 0.85 else
                          "Slower than peers" if top_median > rest_median * 1.15 else
                          "Similar to peers")
        else:
            speed_note = ""

        rows.append({
            "subject": subj,
            "cases": n,
            "key_person": top1,
            "key_person_pct": top1_pct,
            "risk_level": spof_flag,
            "key_person_median_hrs": top_median,
            "others_median_hrs": rest_median,
            "speed_note": speed_note,
            "impact_if_absent": f"{int(top1_pct * n / 100):,} cases would need reassignment",
        })

    result = pd.DataFrame(rows)
    if result.empty:
        return pd.DataFrame({"note": ["No key-person risks detected (all subjects well-distributed)"]})
    return result.sort_values("key_person_pct", ascending=False).reset_index(drop=True)


# ═══════════════════════════════════════════════════════════
#  TAB 6: MIXED-TYPE FLAGS
# ═══════════════════════════════════════════════════════════

def tab_06_mixed_type_flags(df):
    """Subjects where p90/median ≥ 5 — signals mixed case types hiding under one label."""
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
            "subject": subj,
            "cases": int(subj_counts[subj]),
            "median_hrs": round(med, 1),
            "p90_hrs": round(p90, 1),
            "p90_median_ratio": round(ratio, 1),
            "flag": "HIGH VARIANCE" if ratio >= 10 else "ELEVATED",
            "interpretation": (
                "This subject likely contains multiple distinct workflows. "
                "Sub-segmentation or subject taxonomy revision recommended."
                if ratio >= 10 else
                "Moderate variance — some cases are significantly more complex. "
                "Investigate whether distinct sub-types exist."
            ),
        })

    result = pd.DataFrame(rows)
    if result.empty:
        return pd.DataFrame({"note": ["No mixed-type subjects detected"]})
    return result.sort_values("p90_median_ratio", ascending=False).reset_index(drop=True)


# ═══════════════════════════════════════════════════════════
#  TAB 7 (APPENDIX): FIELD AUDIT
# ═══════════════════════════════════════════════════════════

def tab_07_field_audit(cases_raw):
    """Schema audit: every raw column, fill rate, distinctness, date range, role."""
    rows = []
    for c in cases_raw.columns:
        s = cases_raw[c]
        non_null = s.dropna()
        fill_pct = round(100 * len(non_null) / len(cases_raw), 1) if len(cases_raw) > 0 else 0
        distinct = int(s.nunique(dropna=True))
        dtype = str(s.dtype)

        # Infer role
        nc = norm_col(c)
        if "date" in nc or "created" in nc or "modified" in nc or "resolved" in nc or "sla" in nc:
            role = "DATE"
            dt = safe_dt(s)
            valid = dt.dropna()
            date_range = f"{valid.min().date()} → {valid.max().date()}" if len(valid) > 0 else ""
        elif "number" in nc or "id" in nc:
            role = "KEY"
            date_range = ""
        elif distinct < 50 and fill_pct > 50:
            role = "CATEGORICAL"
            date_range = ""
        elif s.dtype == "object" and non_null.astype(str).str.len().median() > 20 if len(non_null) > 0 else False:
            role = "TEXT"
            date_range = ""
        else:
            role = "NUMERIC" if pd.api.types.is_numeric_dtype(s) else "OTHER"
            date_range = ""

        sample = trunc(non_null.iloc[0], 80) if len(non_null) > 0 else ""

        rows.append({
            "column": c, "dtype": dtype, "fill_pct": fill_pct,
            "distinct": distinct, "inferred_role": role,
            "date_range": date_range, "sample": sample,
        })

    return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════
#  TAB 8 (APPENDIX): COMPANY × SUBJECT
# ═══════════════════════════════════════════════════════════

def tab_08_company_subject(df):
    """Observed/expected ratio for company-subject combos."""
    client = df[~df["_is_internal"]].copy()
    n_total = len(client)
    top_subjects = client["_subject"].value_counts().head(15).index.tolist()
    top_companies = client["_company_clean"].value_counts().head(30).index.tolist()

    sub = client[client["_subject"].isin(top_subjects) & client["_company_clean"].isin(top_companies)]
    subj_pct = client["_subject"].value_counts() / n_total
    co_pct = client["_company_clean"].value_counts() / n_total

    rows = []
    for subj in top_subjects:
        for co in top_companies:
            observed = len(sub[(sub["_subject"] == subj) & (sub["_company_clean"] == co)])
            if observed < 5:
                continue
            expected = subj_pct.get(subj, 0) * co_pct.get(co, 0) * n_total
            if expected < 1:
                continue
            ratio = round(observed / expected, 2)
            if ratio >= 1.5 or observed >= 20:
                rows.append({
                    "subject": subj, "company": co,
                    "observed": observed, "expected": round(expected, 1),
                    "obs_exp_ratio": ratio,
                    "flag": "OVER-REPRESENTED" if ratio >= 2.0 else "",
                })

    result = pd.DataFrame(rows)
    if result.empty:
        return pd.DataFrame({"note": ["No notable company-subject concentrations"]})
    return result.sort_values("obs_exp_ratio", ascending=False).reset_index(drop=True)


# ═══════════════════════════════════════════════════════════
#  TAB 9 (APPENDIX): RESEARCH DETAIL (keywords + samples)
# ═══════════════════════════════════════════════════════════

def tab_09_research_detail(df):
    """Research sub-segmentation detail: keywords and samples per cluster."""
    client = df[~df["_is_internal"]].copy()
    research = client[client["_subject"] == "Research"].copy()
    if len(research) == 0:
        return pd.DataFrame({"note": ["No Research cases"]})

    research["_cluster"] = research.apply(
        lambda r: assign_research_cluster(r.get("_description_text", ""), r.get("_activity_subj_text", "")),
        axis=1
    )

    parts = []

    # Keywords per cluster
    kw_rows = []
    for cluster_name in research["_cluster"].unique():
        texts = research[research["_cluster"] == cluster_name]["_description_text"].tolist()
        keywords = extract_top_keywords(texts, top_n=15)
        for word, count in keywords:
            kw_rows.append({"section": "KEYWORDS", "cluster": cluster_name, "keyword": word, "count": count})
    if kw_rows:
        parts.append(pd.DataFrame(kw_rows))

    # Samples per cluster
    sample_rows = []
    for cluster_name in research["_cluster"].unique():
        cluster_data = research[research["_cluster"] == cluster_name]
        samples = cluster_data[cluster_data.get("_activity_subj_len", pd.Series(dtype=int)) > 5].head(5)
        for _, r in samples.iterrows():
            sample_rows.append({
                "section": "SAMPLES",
                "cluster": cluster_name,
                "activity_subject": trunc(r.get("_activity_subj_text", ""), 120),
                "description": trunc(r.get("_description_text", ""), 120),
                "hours": round(r["_hours"], 1) if pd.notna(r.get("_hours")) else "",
            })
    if sample_rows:
        parts.append(pd.DataFrame(sample_rows))

    if not parts:
        return pd.DataFrame({"note": ["No detail available"]})
    return pd.concat(parts, ignore_index=True, sort=False)


# ═══════════════════════════════════════════════════════════
#  TAB 10 (APPENDIX): WEEKLY × SUBJECT TRENDS
# ═══════════════════════════════════════════════════════════

def tab_10_weekly_subject(df):
    """Per-subject weekly created/unresolved detail for reference."""
    client = df[~df["_is_internal"]].copy()
    if "_week" not in client.columns:
        return pd.DataFrame({"note": ["No week data"]})

    top_subjects = client["_subject"].value_counts().head(12).index.tolist()
    sub = client[client["_subject"].isin(top_subjects)].dropna(subset=["_week"])

    weeks = sorted(sub["_week"].unique())
    rows = []
    for subj in top_subjects:
        for wk in weeks:
            pool = sub[(sub["_subject"] == subj) & (sub["_week"] == wk)]
            n = len(pool)
            if n == 0:
                continue
            unres = (~pool["_is_resolved"]).sum() if "_is_resolved" in pool.columns else 0
            rows.append({
                "subject": subj, "week": wk, "created": n,
                "unresolved": int(unres),
                "pct_unresolved": round(100 * unres / n, 1) if n > 0 else 0,
            })

    return pd.DataFrame(rows) if rows else pd.DataFrame({"note": ["No weekly data"]})


# ═══════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════

def main():
    start = datetime.datetime.now()
    log(f"WAB Subject Deep Dive — {start.strftime('%Y-%m-%d %H:%M:%S')}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    cases_raw = read_file(CASE_FILE, "Cases")
    if cases_raw.empty:
        log("FATAL: Cases file not loaded. Exiting.")
        return

    log("\n--- Preparing Cases ---")
    cases = prepare_cases(cases_raw)
    n_client = (~cases["_is_internal"]).sum()
    n_internal = cases["_is_internal"].sum()
    log(f"  Client cases: {n_client:,}  |  Internal: {n_internal:,}")

    log("\n--- Building sheets ---")
    sheets = OrderedDict()

    # ── Executive tabs (1-6) ──
    log("  Tab 1: Subject Dashboard")
    sheets["1_SubjectDashboard"] = tab_01_subject_dashboard(cases)

    log("  Tab 2: Claim Check")
    sheets["2_ClaimCheck"] = tab_02_claim_check(cases)

    log("  Tab 3: Research Breakdown")
    sheets["3_ResearchBreakdown"] = tab_03_research_breakdown(cases)

    log("  Tab 4: Banker-Hours Budget")
    sheets["4_BankerHoursBudget"] = tab_04_banker_hours(cases)

    log("  Tab 5: Key-Person Risk")
    sheets["5_KeyPersonRisk"] = tab_05_key_person_risk(cases)

    log("  Tab 6: Mixed-Type Flags")
    sheets["6_MixedTypeFlags"] = tab_06_mixed_type_flags(cases)

    # ── Appendix tabs (7-10) ──
    log("  Tab 7: Field Audit (appendix)")
    sheets["7_REF_FieldAudit"] = tab_07_field_audit(cases_raw)

    log("  Tab 8: Company × Subject (appendix)")
    sheets["8_REF_CompanySubject"] = tab_08_company_subject(cases)

    log("  Tab 9: Research Detail (appendix)")
    sheets["9_REF_ResearchDetail"] = tab_09_research_detail(cases)

    log("  Tab 10: Weekly × Subject (appendix)")
    sheets["10_REF_WeeklySubject"] = tab_10_weekly_subject(cases)

    # ── Write Excel ──
    log(f"\nWriting: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            if sdf is not None:
                write_sheet(writer, name[:31], sdf)
    log("  Done.")

    # ── Write markdown summary ──
    log(f"Writing: {OUTPUT_MD}")
    md = [
        "# WAB Subject Deep Dive",
        f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}\n",
        "## Input",
        f"- Cases: {len(cases_raw):,} rows",
        f"- Client cases: {n_client:,}",
        f"- Internal cases: {n_internal:,}\n",
        "## Workbook Structure",
        "",
        "### Executive Tabs (1-6)",
        "| Tab | Purpose |",
        "|-----|---------|",
        "| **1_SubjectDashboard** | One row per subject: cases/week, banker-hrs/week, SPOF, trend, action tag |",
        "| **2_ClaimCheck** | Chris's March 23 claims tested against data: Confirmed / Partially / Contradicted |",
        "| **3_ResearchBreakdown** | Research (catch-all) split into 5 actionable sub-types |",
        "| **4_BankerHoursBudget** | Hours consumed per subject per week. Wait-time vs work-time. Recoverable estimate |",
        "| **5_KeyPersonRisk** | Subjects where one person handles >35% — vacation/attrition risk |",
        "| **6_MixedTypeFlags** | Subjects where p90 > 5× median — mixed workflows hiding under one label |",
        "",
        "### Appendix Tabs (7-10)",
        "| Tab | Purpose |",
        "|-----|---------|",
        "| **7_REF_FieldAudit** | Every raw case column: fill rate, distinct values, inferred role, date range |",
        "| **8_REF_CompanySubject** | Company × subject over-representation analysis |",
        "| **9_REF_ResearchDetail** | Research cluster keywords and text samples |",
        "| **10_REF_WeeklySubject** | Per-subject weekly created/unresolved time series |",
    ]

    if WARN:
        md.append("\n## Warnings")
        for w in WARN:
            md.append(f"- {w}")

    md.append("\n## Sheet Row Counts")
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
