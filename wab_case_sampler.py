"""
WAB Case Sampler — Tier 1 & Tier 2 Subject Deep Dive
=====================================================
Standalone VDI script.  Reads Cases file only.
Produces one Excel workbook with flagged case lists for CRM review.

Each sheet = one anomaly type within a specific subject.
Case Number column lets you paste directly into CRM search.

Dependencies: pandas, openpyxl  (standard library otherwise)

RUN
---
# Full deep dive — all 15 subjects (default output: wab_case_sampler.xlsx):
python wab_case_sampler.py

# Only the 5 newly added subjects
# (New Account Child Case, CD Maintenance, Close Account,
#  IntraFi Maintenance, Online Banking)
# Output: wab_case_sampler_new_subjects.xlsx
python wab_case_sampler.py --only-new

# Custom subset — pass a comma-separated list of Subject names
# Output: wab_case_sampler_subset.xlsx
python wab_case_sampler.py --subjects "Online Banking,CD Maintenance"
"""

# ┌─────────────────────────────────────────────────────────┐
# │  EDIT THESE 2 VARIABLES BEFORE RUNNING                  │
# └─────────────────────────────────────────────────────────┘
CASE_FILE  = r"C:\Users\YourName\Desktop\AAB All Cases.xlsx"
OUTPUT_DIR = r"C:\Users\YourName\Desktop\wab_output"
# ┌─────────────────────────────────────────────────────────┐
# │  DO NOT EDIT BELOW THIS LINE                            │
# └─────────────────────────────────────────────────────────┘

import os, re, html, datetime, warnings
import pandas as pd
import numpy as np

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "wab_case_sampler.xlsx")
LOG = []

INTERNAL_COMPANIES = {"AAB ADMIN", "WAB ADMIN", "AAB ADMIN -", "WAB ADMIN -"}

# Subjects added in the latest revision — selectable via --only-new
NEW_SUBJECTS = [
    "New Account Child Case",
    "CD Maintenance",
    "Close Account",
    "IntraFi Maintenance",
    "Online Banking",
]

# When None, run all subjects. When a set/list of subject names,
# subject_subset() returns empty for any subject not in this list,
# and main() will skip empty sheets / summary rows at write time.
SELECTED_SUBJECTS = None

# Max rows per sample sheet — enough to review in CRM without being overwhelming
SAMPLE_SLOW      = 40
SAMPLE_FAST      = 30
SAMPLE_UNRESOLVED = 60   # show all if fewer


# ═══════════════════════════════════════════════════════════
#  UTILITIES  (mirrored from wab_cases_deep_dive.py)
# ═══════════════════════════════════════════════════════════

def log(m):
    LOG.append(m); print(m)

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
        return pd.Series([pd.NaT] * len(series), index=series.index)

def safe_numeric(series):
    if pd.api.types.is_numeric_dtype(series):
        return series
    return pd.to_numeric(series, errors="coerce")

def trunc(v, n=200):
    if pd.isna(v): return ""
    s = str(v).replace("\r", " ").replace("\n", " ").strip()
    return s[:n] + "…" if len(s) > n else s

def fmt_hrs(h):
    """Human-readable hours: '2.3h' or '4.5 days'."""
    if pd.isna(h): return ""
    if h < 48:
        return f"{h:.1f}h"
    return f"{h/24:.1f}d"


# ═══════════════════════════════════════════════════════════
#  CASE PREPARATION
# ═══════════════════════════════════════════════════════════

def prepare_cases(raw):
    df = raw.copy()

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
        "resolved_on":   find_col(df, "Resolved On"),
        "resolved_hrs":  find_col(df, "Resolved In Hours"),
        "owner":         find_col(df, "Manager (Owning User) (User)", "Owner"),
        "pod":           find_col(df, "POD Name (Owning User) (User)", "POD Name"),
        "sla_start":     find_col(df, "SLA Start"),
    }

    # dates
    for key in ["created_on", "modified_on", "resolved_on", "sla_start"]:
        src = col_map.get(key)
        if src:
            df[f"_{key}_dt"] = safe_dt(df[src])

    # resolution hours
    src = col_map.get("resolved_hrs")
    if src:
        df["_hours"] = safe_numeric(df[src])
    else:
        df["_hours"] = np.nan

    # resolved flag
    resolved_kw = ["resolved", "closed", "cancelled", "canceled"]
    src_sr = col_map.get("status_reason") or col_map.get("status")
    if src_sr:
        df["_is_resolved"] = df[src_sr].fillna("").astype(str).str.lower().apply(
            lambda x: any(kw in x for kw in resolved_kw))
    else:
        df["_is_resolved"] = True

    # internal flag
    src_co = col_map.get("company")
    if src_co:
        df["_company_clean"] = df[src_co].fillna("(blank)").astype(str).str.strip()
        df["_is_internal"] = df["_company_clean"].str.upper().apply(
            lambda x: any(x.startswith(kw) for kw in INTERNAL_COMPANIES) or x == "(BLANK)"
        )
    else:
        df["_company_clean"] = "(blank)"
        df["_is_internal"] = False

    # clean subject / origin
    src_subj = col_map.get("subject")
    df["_subject"] = df[src_subj].fillna("(blank)").astype(str).str.strip() if src_subj else "(blank)"

    src_orig = col_map.get("origin")
    df["_origin"] = df[src_orig].fillna("(blank)").astype(str).str.strip() if src_orig else "(blank)"

    # age for unresolved
    now = pd.Timestamp.now()
    if "_created_on_dt" in df.columns:
        df["_age_days"] = ((now - df["_created_on_dt"]).dt.total_seconds() / 86400).round(1)
    else:
        df["_age_days"] = np.nan

    if "_modified_on_dt" in df.columns:
        df["_days_since_modified"] = ((now - df["_modified_on_dt"]).dt.total_seconds() / 86400).round(1)
    else:
        df["_days_since_modified"] = np.nan

    df._col_map = col_map
    return df, col_map


# ═══════════════════════════════════════════════════════════
#  OUTPUT COLUMNS  — what goes on every sheet
# ═══════════════════════════════════════════════════════════

def build_output_row(df, col_map, extra_cols=None):
    """
    Return a display DataFrame with the standard columns every sheet needs.
    extra_cols: list of (output_label, source_series) pairs to append.
    """
    out = pd.DataFrame()

    def _add(label, src_key, fallback_series=None):
        src = col_map.get(src_key)
        if src and src in df.columns:
            out[label] = df[src].fillna("").astype(str).apply(lambda v: trunc(v, 200))
        elif fallback_series is not None:
            out[label] = fallback_series
        else:
            out[label] = ""

    _add("Case Number",     "case_number")
    _add("Company",         "company",       df["_company_clean"])
    _add("Subject",         "subject",       df["_subject"])
    _add("Activity Subject","activity_subj")
    _add("Description",     "description")
    _add("Origin",          "origin",        df["_origin"])
    _add("Pod",             "pod")
    _add("Owner",           "owner")
    _add("Created On",      "created_on")
    _add("Modified On",     "modified_on")

    # resolution hours in human-readable form
    out["Resolution"] = df["_hours"].apply(fmt_hrs)

    if extra_cols:
        for label, series in extra_cols:
            out[label] = series.values

    return out.reset_index(drop=True)


# ═══════════════════════════════════════════════════════════
#  SHEET WRITERS
# ═══════════════════════════════════════════════════════════

def write_sheet(writer, name, df, note=None):
    if df is None or df.empty:
        df = pd.DataFrame({"note": [note or "No cases matched this filter."]})
    sheet_name = name[:31]
    df.to_excel(writer, sheet_name=sheet_name, index=False, freeze_panes=(1, 0))
    ws = writer.sheets[sheet_name]
    for col_cells in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col_cells), default=8)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 60)


# ═══════════════════════════════════════════════════════════
#  SAMPLING FUNCTIONS — one per subject / anomaly type
# ═══════════════════════════════════════════════════════════

def subject_subset(client, subject_name):
    """
    Case-insensitive subject filter.

    Honors the module-level SELECTED_SUBJECTS allow-list. When a filter is
    active and this subject is not in it, returns an empty DataFrame so the
    downstream sample/summary calls naturally yield nothing — main() then
    drops the empty sheets and summary rows before writing the workbook.
    """
    if SELECTED_SUBJECTS is not None and subject_name not in SELECTED_SUBJECTS:
        return client.iloc[0:0].copy()
    return client[client["_subject"].str.lower() == subject_name.lower()].copy()


def p90_threshold(subset):
    hrs = subset["_hours"].dropna()
    return hrs.quantile(0.9) if len(hrs) >= 10 else None


def slow_sample(subset, threshold, n=SAMPLE_SLOW):
    """Cases at or above the p90 threshold, sorted slowest first."""
    if threshold is None:
        return pd.DataFrame()
    return subset[subset["_hours"] >= threshold].sort_values("_hours", ascending=False).head(n)


def fast_sample(subset, cutoff_hrs, n=SAMPLE_FAST):
    """Cases resolved at or below cutoff_hrs, sorted fastest first."""
    return subset[
        subset["_hours"].notna() & (subset["_hours"] <= cutoff_hrs)
    ].sort_values("_hours").head(n)


def unresolved_sample(subset, n=SAMPLE_UNRESOLVED):
    """Unresolved cases sorted by age descending (oldest first)."""
    unres = subset[~subset["_is_resolved"]].copy()
    if "_age_days" in unres.columns:
        unres = unres.sort_values("_age_days", ascending=False)
    return unres.head(n)


def origin_filter(subset, origin_value, n=SAMPLE_FAST):
    """Cases matching a specific origin, sorted by creation date."""
    filtered = subset[subset["_origin"].str.lower() == origin_value.lower()].copy()
    if "_created_on_dt" in filtered.columns:
        filtered = filtered.sort_values("_created_on_dt", ascending=False)
    return filtered.head(n)


def keyword_sample(subset, keywords, n=30):
    """
    Cases where Activity Subject contains any of the keywords (case-insensitive).
    Returns DataFrame with an extra 'Matched Keyword' column.
    """
    act_col = subset.columns[subset.columns.str.lower() == "activity subject"]
    # fall back to raw column if display col not present
    mask = pd.Series(False, index=subset.index)
    matched = pd.Series("", index=subset.index)

    act_src = None
    for c in subset.columns:
        if "activity" in c.lower() and "subj" in c.lower():
            act_src = c
            break

    if act_src is None:
        return pd.DataFrame(), []

    for kw in keywords:
        hit = subset[act_src].fillna("").astype(str).str.lower().str.contains(kw.lower(), regex=False)
        matched[hit & (matched == "")] = kw
        mask |= hit

    result = subset[mask].copy()
    result["_matched_kw"] = matched[mask]
    result = result.sort_values(["_matched_kw", "_hours"], ascending=[True, False])
    return result.head(n), keywords


# ═══════════════════════════════════════════════════════════
#  SUMMARY ROW BUILDER
# ═══════════════════════════════════════════════════════════

SUMMARY_ROWS = []

def record_summary(sheet_name, subject, anomaly_type, n_total_subject,
                   n_flagged, p90_val, fast_cutoff, question):
    SUMMARY_ROWS.append({
        "Sheet":          sheet_name,
        "Subject":        subject,
        "Anomaly Type":   anomaly_type,
        "Subject Total":  n_total_subject,
        "Cases Flagged":  n_flagged,
        "P90 Threshold":  fmt_hrs(p90_val) if p90_val else "—",
        "Fast Cutoff":    fmt_hrs(fast_cutoff) if fast_cutoff else "—",
        "Question to Answer in CRM": question,
    })


# ═══════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ── Load ──
    log(f"Reading cases: {CASE_FILE}")
    if not os.path.isfile(CASE_FILE):
        print(f"ERROR: File not found: {CASE_FILE}")
        return

    raw = pd.read_excel(CASE_FILE, header=0, engine="openpyxl")
    raw = raw.loc[:, ~raw.columns.astype(str).str.match(r"^Unnamed")]
    raw = raw.dropna(axis=1, how="all")
    log(f"  Loaded {len(raw):,} rows, {len(raw.columns)} columns")

    cases, col_map = prepare_cases(raw)
    client = cases[~cases["_is_internal"]].copy()
    log(f"  Client cases: {len(client):,}  |  Internal excluded: {cases['_is_internal'].sum():,}")

    sheets = {}

    # ───────────────────────────────────────────────────────
    #  TIER 1 SUBJECTS
    # ───────────────────────────────────────────────────────

    # ── Fraud Alert ──
    log("\n[Tier 1] Fraud Alert")
    fa = subject_subset(client, "Fraud Alert")
    fa_p90 = p90_threshold(fa)
    log(f"  Total: {len(fa):,}  |  P90: {fmt_hrs(fa_p90)}")

    slo = slow_sample(fa, fa_p90)
    out = build_output_row(slo, col_map)
    sheets["T1_FraudAlert_Slow"] = out
    record_summary("T1_FraudAlert_Slow", "Fraud Alert", "Slow (≥p90)",
                   len(fa), len(slo), fa_p90, None,
                   "Fraud Alerts take median 2.6h but p90=24h. What makes 10% take a full day? "
                   "Is this a different type of alert (genuine investigation vs. routine notification)?")

    fst = fast_sample(fa, 0.5)  # ≤30 min
    out = build_output_row(fst, col_map)
    sheets["T1_FraudAlert_Fast"] = out
    record_summary("T1_FraudAlert_Fast", "Fraud Alert", "Fast (≤30 min)",
                   len(fa), len(fst), None, 0.5,
                   "What does a routine Fraud Alert look like? Compare Activity Subject with slow cases "
                   "to check if slow=complex and fast=notification noise.")

    # ── Transfer ──
    log("\n[Tier 1] Transfer")
    tr = subject_subset(client, "Transfer")
    tr_p90 = p90_threshold(tr)
    log(f"  Total: {len(tr):,}  |  P90: {fmt_hrs(tr_p90)}")

    slo = slow_sample(tr, tr_p90)
    out = build_output_row(slo, col_map)
    sheets["T1_Transfer_Slow"] = out
    record_summary("T1_Transfer_Slow", "Transfer", "Slow (≥p90)",
                   len(tr), len(slo), tr_p90, None,
                   "Transfers median 1.3h but p90=23.7h (18x jump). A mechanical transfer shouldn't "
                   "take a day. Are these misrouted wires, disputes, or wrong-account transfers?")

    fst = fast_sample(tr, 1.0)  # ≤1 hour
    out = build_output_row(fst, col_map)
    sheets["T1_Transfer_Fast"] = out
    record_summary("T1_Transfer_Fast", "Transfer", "Fast (≤1h)",
                   len(tr), len(fst), None, 1.0,
                   "What does a clean Transfer look like? Activity Subject should confirm these are "
                   "simple intra-account moves vs. the slow cases.")

    # ── Statements ──
    log("\n[Tier 1] Statements")
    st = subject_subset(client, "Statements")
    st_p90 = p90_threshold(st)
    log(f"  Total: {len(st):,}  |  P90: {fmt_hrs(st_p90)}")

    slo = slow_sample(st, st_p90)
    out = build_output_row(slo, col_map)
    sheets["T1_Statements_Slow"] = out
    record_summary("T1_Statements_Slow", "Statements", "Slow (≥p90)",
                   len(st), len(slo), st_p90, None,
                   "66.7% of Statement emails are Proofpoint receipt noise. Are these slow cases "
                   "the real statement requests that survived noise, or data artifacts?")

    fst = fast_sample(st, 0.5)  # ≤30 min
    out = build_output_row(fst, col_map)
    sheets["T1_Statements_Fast"] = out
    record_summary("T1_Statements_Fast", "Statements", "Fast (≤30 min)",
                   len(st), len(fst), None, 0.5,
                   "Are the fast Statements all Proofpoint receipts auto-closed? "
                   "Check Activity Subject — if it says 'Read:' these confirm the noise pattern.")

    # ── NSF and Non-Post ──
    log("\n[Tier 1] NSF and Non-Post")
    nsf = subject_subset(client, "NSF and Non-Post")
    log(f"  Total: {len(nsf):,}")

    unres = unresolved_sample(nsf, n=SAMPLE_UNRESOLVED)
    extra = [("Age (days)", unres["_age_days"])] if "_age_days" in unres.columns else []
    out = build_output_row(unres, col_map, extra_cols=extra)
    sheets["T1_NSF_Unresolved"] = out
    record_summary("T1_NSF_Unresolved", "NSF and Non-Post", "Unresolved",
                   len(nsf), len(unres), None, None,
                   "NSF resolves 99.8% of the time. The 0.2% unresolved are anomalies — "
                   "data error, genuine dispute, or a case opened and abandoned?")

    email_nsf = origin_filter(nsf, "Email", n=SAMPLE_FAST)
    out = build_output_row(email_nsf, col_map)
    sheets["T1_NSF_EmailOrigin"] = out
    record_summary("T1_NSF_EmailOrigin", "NSF and Non-Post", "Email origin (not Report)",
                   len(nsf), len(email_nsf), None, None,
                   "Most NSF cases come from automated Report. Email-origin NSF cases are "
                   "human-initiated — what do clients write in? Different workflow from Report NSF?")

    # ───────────────────────────────────────────────────────
    #  TIER 2 SUBJECTS
    # ───────────────────────────────────────────────────────

    # ── Research ──
    log("\n[Tier 2] Research")
    res = subject_subset(client, "Research")
    res_p90 = p90_threshold(res)
    log(f"  Total: {len(res):,}  |  P90: {fmt_hrs(res_p90)}")

    slo = slow_sample(res, res_p90)
    out = build_output_row(slo, col_map)
    sheets["T2_Research_Slow"] = out
    record_summary("T2_Research_Slow", "Research", "Slow (≥p90)",
                   len(res), len(slo), res_p90, None,
                   "Research p90=106.5h (4.4 days). Chris said it's mostly payment research. "
                   "What's in Activity Subject for 4-day+ cases — trace requests waiting on FedWire? "
                   "Disputes? Reconciliation? Each sub-type is a different use case.")

    fst = fast_sample(res, 2.0)  # ≤2 hours
    out = build_output_row(fst, col_map)
    sheets["T2_Research_Fast"] = out
    record_summary("T2_Research_Fast", "Research", "Fast (≤2h)",
                   len(res), len(fst), None, 2.0,
                   "Research cases resolved in <2h — what kind of 'research' resolves that quickly? "
                   "If Activity Subject shows a pattern, these may be a distinct sub-type "
                   "that could be templated.")

    unres = unresolved_sample(res, n=SAMPLE_UNRESOLVED)
    extra = [("Age (days)", unres["_age_days"])] if "_age_days" in unres.columns else []
    out = build_output_row(unres, col_map, extra_cols=extra)
    sheets["T2_Research_Unresolved"] = out
    record_summary("T2_Research_Unresolved", "Research", "Unresolved",
                   len(res), len(unres), None, None,
                   "3% of Research cases are open. Are these genuine ongoing investigations "
                   "or cases that got created and forgotten? Days since modified will tell you.")

    # Research keyword clusters
    log("  Research keyword clusters")
    research_keywords = ["payment", "trace", "return", "ACH", "wire", "dispute", "reconcil",
                         "fraud", "NSF", "deposit", "fee", "interest"]
    kw_hits, _ = keyword_sample(res, research_keywords, n=80)
    if not kw_hits.empty:
        extra = [("Matched Keyword", kw_hits["_matched_kw"])]
        out = build_output_row(kw_hits, col_map, extra_cols=extra)
        sheets["T2_Research_Keywords"] = out
        record_summary("T2_Research_Keywords", "Research", "Keyword clusters in Activity Subject",
                       len(res), len(kw_hits), None, None,
                       "Sub-segment Research by keyword in Activity Subject. Each keyword cluster "
                       "is a candidate sub-workflow. Compare resolution hours across clusters — "
                       "if 'trace' takes 5x longer than 'fee', they are different use cases.")

    # ── New Account Request ──
    log("\n[Tier 2] New Account Request")
    nar = subject_subset(client, "New Account Request")
    nar_p90 = p90_threshold(nar)
    log(f"  Total: {len(nar):,}  |  P90: {fmt_hrs(nar_p90)}")

    slo = slow_sample(nar, nar_p90)
    out = build_output_row(slo, col_map)
    sheets["T2_NewAccount_Slow"] = out
    record_summary("T2_NewAccount_Slow", "New Account Request", "Slow (≥p90)",
                   len(nar), len(slo), nar_p90, None,
                   "New Account median=25.2h (already long), p90=190.2h (8 days). "
                   "Missing TIN/EIN/signature card is the hypothesis. Does Description "
                   "or Activity Subject mention missing documents in these slow cases?")

    fst = fast_sample(nar, 4.0)  # ≤4 hours
    out = build_output_row(fst, col_map)
    sheets["T2_NewAccount_Fast"] = out
    record_summary("T2_NewAccount_Fast", "New Account Request", "Fast (≤4h)",
                   len(nar), len(fst), None, 4.0,
                   "New accounts resolved in <4h — what made these fast? "
                   "Were all documents present upfront? If so, this is the baseline "
                   "for what AI-assisted missing-info detection should aim to replicate.")

    unres = unresolved_sample(nar, n=SAMPLE_UNRESOLVED)
    extra = [("Age (days)", unres["_age_days"])] if "_age_days" in unres.columns else []
    out = build_output_row(unres, col_map, extra_cols=extra)
    sheets["T2_NewAccount_Unresolved"] = out
    record_summary("T2_NewAccount_Unresolved", "New Account Request", "Unresolved",
                   len(nar), len(unres), None, None,
                   "3.9% (~135 cases) of new account requests are open. "
                   "These are blocking revenue. What's the bottleneck — "
                   "client hasn't sent docs, or WAB internal processing delay?")

    # ── Signature Card ──
    log("\n[Tier 2] Signature Card")
    sc = subject_subset(client, "Signature Card")
    sc_p90 = p90_threshold(sc)
    log(f"  Total: {len(sc):,}  |  P90: {fmt_hrs(sc_p90)}")

    slo = slow_sample(sc, sc_p90)
    out = build_output_row(slo, col_map)
    sheets["T2_SignatureCard_Slow"] = out
    record_summary("T2_SignatureCard_Slow", "Signature Card", "Slow (≥p90)",
                   len(sc), len(slo), sc_p90, None,
                   "Signature Card work should be document-driven. Slow outliers likely indicate "
                   "missing signatures, stale paperwork, or back-and-forth clarification. "
                   "Check Description and Activity Subject for the blocking pattern.")

    fst = fast_sample(sc, 4.0)  # ≤4 hours
    out = build_output_row(fst, col_map)
    sheets["T2_SignatureCard_Fast"] = out
    record_summary("T2_SignatureCard_Fast", "Signature Card", "Fast (≤4h)",
                   len(sc), len(fst), None, 4.0,
                   "Fast Signature Card cases show the clean baseline: complete documentation, "
                   "clear request type, and no follow-up loop. Compare these against slow cases.")

    unres = unresolved_sample(sc, n=SAMPLE_UNRESOLVED)
    extra = [("Age (days)", unres["_age_days"])] if "_age_days" in unres.columns else []
    out = build_output_row(unres, col_map, extra_cols=extra)
    sheets["T2_SignatureCard_Unresolved"] = out
    record_summary("T2_SignatureCard_Unresolved", "Signature Card", "Unresolved",
                   len(sc), len(unres), None, None,
                   "Open Signature Card cases are likely waiting on client action or stuck in review. "
                   "Use age and last activity to separate expected pending work from neglected cases.")

    # ── Account Maintenance ──
    log("\n[Tier 2] Account Maintenance")
    am = subject_subset(client, "Account Maintenance")
    am_p90 = p90_threshold(am)
    log(f"  Total: {len(am):,}  |  P90: {fmt_hrs(am_p90)}")

    slo = slow_sample(am, am_p90)
    out = build_output_row(slo, col_map)
    sheets["T2_AcctMaint_Slow"] = out
    record_summary("T2_AcctMaint_Slow", "Account Maintenance", "Slow (≥p90)",
                   len(am), len(slo), am_p90, None,
                   "Account Maintenance median=3.6h but p90=316.6h (13+ days) — "
                   "the widest variance of any subject (88x). These slow cases are almost certainly "
                   "a different workflow. What does Activity Subject say — CD rollover? "
                   "Beneficiary update? Address change? Each is a different product.")

    fst = fast_sample(am, 1.0)  # ≤1 hour
    out = build_output_row(fst, col_map)
    sheets["T2_AcctMaint_Fast"] = out
    record_summary("T2_AcctMaint_Fast", "Account Maintenance", "Fast (≤1h)",
                   len(am), len(fst), None, 1.0,
                   "What maintenance resolves in under an hour? "
                   "These are likely simple updates (address, phone, contact name). "
                   "Confirms the split: routine maintenance vs. complex product-level changes.")

    unres = unresolved_sample(am, n=SAMPLE_UNRESOLVED)
    extra = [("Age (days)", unres["_age_days"])] if "_age_days" in unres.columns else []
    out = build_output_row(unres, col_map, extra_cols=extra)
    sheets["T2_AcctMaint_Unresolved"] = out
    record_summary("T2_AcctMaint_Unresolved", "Account Maintenance", "Unresolved",
                   len(am), len(unres), None, None,
                   "6.2% (~187 cases) of Account Maintenance is open — highest Tier 2 unresolved rate. "
                   "Days since modified: are these being actively worked or sitting untouched? "
                   "If untouched >14 days, that is the escalation/missing-info use case.")

    # Account Maintenance keyword clusters
    log("  Account Maintenance keyword clusters")
    am_keywords = ["CD", "certificate", "beneficiary", "address", "contact", "phone",
                   "signer", "authorized", "rate", "interest", "ICS", "IntraFi",
                   "wire", "ACH", "online", "statement"]
    kw_hits, _ = keyword_sample(am, am_keywords, n=80)
    if not kw_hits.empty:
        extra = [("Matched Keyword", kw_hits["_matched_kw"])]
        out = build_output_row(kw_hits, col_map, extra_cols=extra)
        sheets["T2_AcctMaint_Keywords"] = out
        record_summary("T2_AcctMaint_Keywords", "Account Maintenance", "Keyword clusters in Activity Subject",
                       len(am), len(kw_hits), None, None,
                       "Account Maintenance lumps very different workflows together. "
                       "Keyword clusters in Activity Subject will confirm whether "
                       "CD/beneficiary/signer changes have different resolution profiles "
                       "and should be separate use cases.")

    # ── QC Finding ──
    log("\n[Tier 2] QC Finding")
    qc = subject_subset(client, "QC Finding")
    qc_p90 = p90_threshold(qc)
    log(f"  Total: {len(qc):,}  |  P90: {fmt_hrs(qc_p90)}")

    slo = slow_sample(qc, qc_p90)
    out = build_output_row(slo, col_map)
    sheets["T2_QCFinding_Slow"] = out
    record_summary("T2_QCFinding_Slow", "QC Finding", "Slow (≥p90)",
                   len(qc), len(slo), qc_p90, None,
                   "Slow QC Finding cases likely need research, correction, or cross-team follow-up. "
                   "Review the activity trail to see which finding types create the longest loops.")

    fst = fast_sample(qc, 2.0)  # ≤2 hours
    out = build_output_row(fst, col_map)
    sheets["T2_QCFinding_Fast"] = out
    record_summary("T2_QCFinding_Fast", "QC Finding", "Fast (≤2h)",
                   len(qc), len(fst), None, 2.0,
                   "Fast QC Finding cases should represent straightforward corrections or false alarms. "
                   "These define the simplest remediation pattern.")

    unres = unresolved_sample(qc, n=SAMPLE_UNRESOLVED)
    extra = [("Age (days)", unres["_age_days"])] if "_age_days" in unres.columns else []
    out = build_output_row(unres, col_map, extra_cols=extra)
    sheets["T2_QCFinding_Unresolved"] = out
    record_summary("T2_QCFinding_Unresolved", "QC Finding", "Unresolved",
                   len(qc), len(unres), None, None,
                   "Open QC Finding cases may indicate unresolved control issues or weak follow-through. "
                   "Check whether they are actively progressing or simply aging in queue.")

    # ── General Questions ──
    log("\n[Tier 2] General Questions")
    gq = subject_subset(client, "General Questions")
    gq_p90 = p90_threshold(gq)
    log(f"  Total: {len(gq):,}  |  P90: {fmt_hrs(gq_p90)}")

    slo = slow_sample(gq, gq_p90)
    out = build_output_row(slo, col_map)
    sheets["T2_GenQuestions_Slow"] = out
    record_summary("T2_GenQuestions_Slow", "General Questions", "Slow (≥p90)",
                   len(gq), len(slo), gq_p90, None,
                   "General Questions median=1.7h but p90=88.4h (3.7 days). "
                   "A question that takes 3+ days is not a general question — "
                   "it is a miscategorized case or a genuinely complex issue. "
                   "Activity Subject will tell you which.")

    fst = fast_sample(gq, 0.5)  # ≤30 min
    out = build_output_row(fst, col_map)
    sheets["T2_GenQuestions_Fast"] = out
    record_summary("T2_GenQuestions_Fast", "General Questions", "Fast (≤30 min)",
                   len(gq), len(fst), None, 0.5,
                   "What does a genuinely simple General Question look like? "
                   "Compare Activity Subject with the slow cases. "
                   "If fast=short questions and slow=escalations, the subject is doing "
                   "double duty and should be split.")

    # ── New Account Child Case ──
    log("\n[Tier 2] New Account Child Case")
    nacc = subject_subset(client, "New Account Child Case")
    nacc_p90 = p90_threshold(nacc)
    log(f"  Total: {len(nacc):,}  |  P90: {fmt_hrs(nacc_p90)}")

    slo = slow_sample(nacc, nacc_p90)
    out = build_output_row(slo, col_map)
    sheets["T2_NewAcctChild_Slow"] = out
    record_summary("T2_NewAcctChild_Slow", "New Account Child Case", "Slow (≥p90)",
                   len(nacc), len(slo), nacc_p90, None,
                   "Child cases spawn from parent New Account Requests. Slow children likely mean "
                   "the parent account is stuck waiting on a downstream setup step "
                   "(online banking enrollment, debit card, ACH origination). "
                   "Activity Subject should reveal which downstream step is the bottleneck.")

    fst = fast_sample(nacc, 2.0)  # ≤2 hours
    out = build_output_row(fst, col_map)
    sheets["T2_NewAcctChild_Fast"] = out
    record_summary("T2_NewAcctChild_Fast", "New Account Child Case", "Fast (≤2h)",
                   len(nacc), len(fst), None, 2.0,
                   "Fast child cases represent the clean downstream setup pattern — "
                   "all parent docs done, child task executed without rework. "
                   "Use these to baseline what a healthy account-onboarding tail looks like.")

    unres = unresolved_sample(nacc, n=SAMPLE_UNRESOLVED)
    extra = [("Age (days)", unres["_age_days"])] if "_age_days" in unres.columns else []
    out = build_output_row(unres, col_map, extra_cols=extra)
    sheets["T2_NewAcctChild_Unresolved"] = out
    record_summary("T2_NewAcctChild_Unresolved", "New Account Child Case", "Unresolved",
                   len(nacc), len(unres), None, None,
                   "Open child cases are setup steps that never closed. "
                   "If the parent account is live but the child is open, that is "
                   "missed fulfillment — and likely the biggest silent NPS risk.")

    # ── CD Maintenance ──
    log("\n[Tier 2] CD Maintenance")
    cd = subject_subset(client, "CD Maintenance")
    cd_p90 = p90_threshold(cd)
    log(f"  Total: {len(cd):,}  |  P90: {fmt_hrs(cd_p90)}")

    slo = slow_sample(cd, cd_p90)
    out = build_output_row(slo, col_map)
    sheets["T2_CDMaint_Slow"] = out
    record_summary("T2_CDMaint_Slow", "CD Maintenance", "Slow (≥p90)",
                   len(cd), len(slo), cd_p90, None,
                   "CD Maintenance is deadline-driven (rollover dates, rate resets). "
                   "Slow cases risk missing the maturity window — check whether these "
                   "cluster around specific rollover events or rate-change instructions.")

    fst = fast_sample(cd, 1.0)  # ≤1 hour
    out = build_output_row(fst, col_map)
    sheets["T2_CDMaint_Fast"] = out
    record_summary("T2_CDMaint_Fast", "CD Maintenance", "Fast (≤1h)",
                   len(cd), len(fst), None, 1.0,
                   "Fast CD Maintenance likely reflects automated rollover confirmations "
                   "and simple rate lookups. Confirms the split between mechanical and "
                   "judgment-driven CD work.")

    unres = unresolved_sample(cd, n=SAMPLE_UNRESOLVED)
    extra = [("Age (days)", unres["_age_days"])] if "_age_days" in unres.columns else []
    out = build_output_row(unres, col_map, extra_cols=extra)
    sheets["T2_CDMaint_Unresolved"] = out
    record_summary("T2_CDMaint_Unresolved", "CD Maintenance", "Unresolved",
                   len(cd), len(unres), None, None,
                   "Unresolved CD cases are deadline risks. Aging beyond the maturity "
                   "window means client either auto-rolled unwillingly or lost interest — "
                   "either way a service failure.")

    # ── Close Account ──
    log("\n[Tier 2] Close Account")
    ca = subject_subset(client, "Close Account")
    ca_p90 = p90_threshold(ca)
    log(f"  Total: {len(ca):,}  |  P90: {fmt_hrs(ca_p90)}")

    slo = slow_sample(ca, ca_p90)
    out = build_output_row(slo, col_map)
    sheets["T2_CloseAccount_Slow"] = out
    record_summary("T2_CloseAccount_Slow", "Close Account", "Slow (≥p90)",
                   len(ca), len(slo), ca_p90, None,
                   "Slow closures usually mean residual balances, pending items, or "
                   "retention outreach. These are the attrition cases most worth reviewing "
                   "— which drop-offs were preventable vs. clean exits?")

    fst = fast_sample(ca, 1.0)  # ≤1 hour
    out = build_output_row(fst, col_map)
    sheets["T2_CloseAccount_Fast"] = out
    record_summary("T2_CloseAccount_Fast", "Close Account", "Fast (≤1h)",
                   len(ca), len(fst), None, 1.0,
                   "Fast closures are typically zero-balance, no-activity shutdowns. "
                   "The baseline for how a clean exit should look.")

    unres = unresolved_sample(ca, n=SAMPLE_UNRESOLVED)
    extra = [("Age (days)", unres["_age_days"])] if "_age_days" in unres.columns else []
    out = build_output_row(unres, col_map, extra_cols=extra)
    sheets["T2_CloseAccount_Unresolved"] = out
    record_summary("T2_CloseAccount_Unresolved", "Close Account", "Unresolved",
                   len(ca), len(unres), None, None,
                   "Open Close Account cases are accounts in limbo — client wants out "
                   "but something is blocking. Highest-risk queue for compliance and NPS.")

    # ── IntraFi Maintenance ──
    log("\n[Tier 2] IntraFi Maintenance")
    intf = subject_subset(client, "IntraFi Maintenance")
    intf_p90 = p90_threshold(intf)
    log(f"  Total: {len(intf):,}  |  P90: {fmt_hrs(intf_p90)}")

    slo = slow_sample(intf, intf_p90)
    out = build_output_row(slo, col_map)
    sheets["T2_IntraFi_Slow"] = out
    record_summary("T2_IntraFi_Slow", "IntraFi Maintenance", "Slow (≥p90)",
                   len(intf), len(slo), intf_p90, None,
                   "IntraFi (ICS/CDARS) work depends on external network confirmations. "
                   "Slow cases likely reflect allocation rebalances or program-bank changes. "
                   "Check Activity Subject for which IntraFi product and step is blocking.")

    fst = fast_sample(intf, 2.0)  # ≤2 hours
    out = build_output_row(fst, col_map)
    sheets["T2_IntraFi_Fast"] = out
    record_summary("T2_IntraFi_Fast", "IntraFi Maintenance", "Fast (≤2h)",
                   len(intf), len(fst), None, 2.0,
                   "Fast IntraFi cases are likely routine deposit placement confirmations. "
                   "Confirms the mechanical baseline vs. multi-day allocation work.")

    unres = unresolved_sample(intf, n=SAMPLE_UNRESOLVED)
    extra = [("Age (days)", unres["_age_days"])] if "_age_days" in unres.columns else []
    out = build_output_row(unres, col_map, extra_cols=extra)
    sheets["T2_IntraFi_Unresolved"] = out
    record_summary("T2_IntraFi_Unresolved", "IntraFi Maintenance", "Unresolved",
                   len(intf), len(unres), None, None,
                   "Open IntraFi cases may be waiting on external network responses. "
                   "Separate truly pending (external) from stalled (internal) using age.")

    # ── Online Banking ──
    log("\n[Tier 2] Online Banking")
    ob = subject_subset(client, "Online Banking")
    ob_p90 = p90_threshold(ob)
    log(f"  Total: {len(ob):,}  |  P90: {fmt_hrs(ob_p90)}")

    slo = slow_sample(ob, ob_p90)
    out = build_output_row(slo, col_map)
    sheets["T2_OnlineBanking_Slow"] = out
    record_summary("T2_OnlineBanking_Slow", "Online Banking", "Slow (≥p90)",
                   len(ob), len(slo), ob_p90, None,
                   "Slow Online Banking cases are usually entitlement changes, token resets, "
                   "or multi-user setup. What should be a 5-minute flip becomes a day when "
                   "signer authority or admin permissions are unclear.")

    fst = fast_sample(ob, 0.5)  # ≤30 min
    out = build_output_row(fst, col_map)
    sheets["T2_OnlineBanking_Fast"] = out
    record_summary("T2_OnlineBanking_Fast", "Online Banking", "Fast (≤30 min)",
                   len(ob), len(fst), None, 0.5,
                   "Fast Online Banking cases are password resets and simple unlocks — "
                   "the canonical self-service candidates. Quantify this volume to "
                   "justify a self-service deflection use case.")

    unres = unresolved_sample(ob, n=SAMPLE_UNRESOLVED)
    extra = [("Age (days)", unres["_age_days"])] if "_age_days" in unres.columns else []
    out = build_output_row(unres, col_map, extra_cols=extra)
    sheets["T2_OnlineBanking_Unresolved"] = out
    record_summary("T2_OnlineBanking_Unresolved", "Online Banking", "Unresolved",
                   len(ob), len(unres), None, None,
                   "Open Online Banking cases typically wait on client action "
                   "(enrollment acceptance, 2FA registration). Age tells you whether "
                   "we abandoned the client or the client abandoned us.")

    # ───────────────────────────────────────────────────────
    #  CROSS-SUBJECT: potential mislabels
    #  Cases where Subject = General Questions or Research
    #  but Activity Subject strongly implies a different subject
    # ───────────────────────────────────────────────────────
    # Skip when a subject filter is active and neither source subject is selected.
    run_xsubject = (
        SELECTED_SUBJECTS is None
        or "General Questions" in SELECTED_SUBJECTS
        or "Research" in SELECTED_SUBJECTS
    )
    log("\n[Cross-subject] Potential mislabels"
        + ("" if run_xsubject else " — skipped (filter active)"))
    catchall = (
        client[client["_subject"].isin(["General Questions", "Research"])].copy()
        if run_xsubject else client.iloc[0:0].copy()
    )

    mislabel_kw = {
        "signature card": "→ Signature Card?",
        "qc finding":    "→ QC Finding?",
        "qc":            "→ QC Finding?",
        "close account":  "→ Close Account?",
        "new account":    "→ New Account Request?",
        "CD":             "→ CD Maintenance?",
        "IntraFi":        "→ IntraFi Maintenance?",
        "fraud":          "→ Fraud Alert?",
        "transfer":       "→ Transfer?",
        "statement":      "→ Statements?",
        "NSF":            "→ NSF and Non-Post?",
        "non-post":       "→ NSF and Non-Post?",
    }

    act_src = None
    for c in catchall.columns:
        if "activity" in c.lower() and "subj" in c.lower():
            act_src = c
            break

    if act_src:
        mask = pd.Series(False, index=catchall.index)
        suggestion = pd.Series("", index=catchall.index)
        for kw, label in mislabel_kw.items():
            hit = catchall[act_src].fillna("").astype(str).str.lower().str.contains(
                kw.lower(), regex=False)
            suggestion[hit & (suggestion == "")] = label
            mask |= hit

        mislabels = catchall[mask].copy()
        mislabels["_suggestion"] = suggestion[mask]
        mislabels = mislabels.sort_values(["_subject", "_suggestion"])
        extra = [("Suggested Subject", mislabels["_suggestion"]),
                 ("Resolution", mislabels["_hours"].apply(fmt_hrs))]
        out = build_output_row(mislabels.head(60), col_map,
                               extra_cols=[(l, s.head(60)) for l, s in extra])
        sheets["X_PotentialMislabels"] = out
        record_summary("X_PotentialMislabels", "General Questions + Research", "Potential mislabels",
                       len(catchall), len(mislabels), None, None,
                       "Cases filed as General Questions or Research but whose Activity Subject "
                       "implies a more specific subject. If significant, triage routing has a "
                       "labelling problem that inflates GQ and Research counts.")

    # ───────────────────────────────────────────────────────
    #  FILTER OUT EMPTY SHEETS / SUMMARY ROWS WHEN A SUBJECT
    #  FILTER IS ACTIVE — keeps the workbook tight.
    # ───────────────────────────────────────────────────────
    if SELECTED_SUBJECTS is not None:
        kept = {k: v for k, v in sheets.items()
                if v is not None and not (hasattr(v, "empty") and v.empty)}
        dropped = set(sheets) - set(kept)
        if dropped:
            log(f"\nFilter active: dropping {len(dropped)} empty sheet(s)")
        sheets = kept
        SUMMARY_ROWS[:] = [r for r in SUMMARY_ROWS
                           if r["Sheet"] in sheets and r["Cases Flagged"] > 0]

    # ───────────────────────────────────────────────────────
    #  SUMMARY SHEET
    # ───────────────────────────────────────────────────────
    log("\nBuilding summary sheet")
    summary_df = pd.DataFrame(SUMMARY_ROWS)
    # add row counts from sheets dict
    if not summary_df.empty:
        summary_df["Rows in Sheet"] = summary_df["Sheet"].map(
            lambda s: len(sheets[s]) if s in sheets and sheets[s] is not None else 0
        )
    sheets["Summary"] = summary_df

    # ───────────────────────────────────────────────────────
    #  WRITE EXCEL — output filename reflects the active filter
    # ───────────────────────────────────────────────────────
    output_xlsx = _output_path()
    log(f"\nWriting: {output_xlsx}")
    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        # Summary first
        write_sheet(writer, "Summary", sheets.pop("Summary"))
        for name, df in sheets.items():
            write_sheet(writer, name, df)
    log("Done.")

    log("\n── Sheet inventory ──")
    for name, df in {**{"Summary": summary_df}, **sheets}.items():
        log(f"  {name}: {len(df) if df is not None else 0} rows")


def _output_path():
    """
    Workbook filename varies by filter so a filtered run does not
    overwrite the full deep-dive workbook.
    """
    if SELECTED_SUBJECTS is None:
        return OUTPUT_XLSX
    if set(SELECTED_SUBJECTS) == set(NEW_SUBJECTS):
        suffix = "_new_subjects"
    else:
        suffix = "_subset"
    return os.path.join(OUTPUT_DIR, f"wab_case_sampler{suffix}.xlsx")


def _parse_args():
    import argparse
    p = argparse.ArgumentParser(
        description="WAB Case Sampler — Tier 1 & Tier 2 Subject Deep Dive",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python wab_case_sampler.py\n"
            "  python wab_case_sampler.py --only-new\n"
            "  python wab_case_sampler.py --subjects \"Online Banking,CD Maintenance\"\n"
        ),
    )
    p.add_argument(
        "--only-new", action="store_true",
        help=("Generate sheets only for the 5 newly added subjects: "
              + ", ".join(NEW_SUBJECTS)),
    )
    p.add_argument(
        "--subjects", default=None,
        help=("Comma-separated list of Subject names to include "
              "(case-sensitive, must match the Subject column exactly). "
              "Overrides --only-new when both are passed."),
    )
    return p.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    if args.subjects:
        SELECTED_SUBJECTS = {s.strip() for s in args.subjects.split(",") if s.strip()}
        print(f"[filter] Selected subjects: {sorted(SELECTED_SUBJECTS)}")
    elif args.only_new:
        SELECTED_SUBJECTS = set(NEW_SUBJECTS)
        print(f"[filter] --only-new → {sorted(SELECTED_SUBJECTS)}")
    main()
