"""
WAB Blank-Company Case Validation
===================================
Quick diagnostic script to profile the 2,214 blank-company cases
and determine whether they behave like client cases or system artifacts.

Run on VDI alongside wab_subject_deep_dive.py.
Share the output Excel with the team to finalize population logic.

Output: one Excel workbook (7 sheets) + console summary.

Dependencies: pandas, openpyxl
"""

# ┌─────────────────────────────────────────────────────────┐
# │  EDIT THESE 2 VARIABLES BEFORE RUNNING                  │
# └─────────────────────────────────────────────────────────┘
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

OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "wab_blank_case_validation.xlsx")

ADMIN_PREFIXES = {"AAB ADMIN", "WAB ADMIN", "AAB ADMIN -", "WAB ADMIN -"}


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

def write_sheet(writer, name, df):
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        df = pd.DataFrame({"note": ["No data"]})
    sheet_name = name[:31]
    df.to_excel(writer, sheet_name=sheet_name, index=False, freeze_panes=(1, 0))
    ws = writer.sheets[sheet_name]
    for col_cells in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 60)


def main():
    start = datetime.datetime.now()
    print(f"=== Blank-Company Case Validation — {start.strftime('%Y-%m-%d %H:%M')} ===\n")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ── Read ──
    print(f"Reading: {CASE_FILE}")
    df = pd.read_excel(CASE_FILE, header=0, engine="openpyxl")
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    df = df.dropna(axis=1, how="all")
    print(f"  Rows: {len(df):,}  Columns: {len(df.columns)}\n")

    # ── Classify into 3 populations ──
    co_col = find_col(df, "Company Name (Company) (Company)", "Company Name", "Customer")
    if not co_col:
        print("FATAL: Cannot find Company Name column.")
        return

    df["_company"] = df[co_col].fillna("").astype(str).str.strip()
    df["_company_upper"] = df["_company"].str.upper()

    def classify(row):
        c = row["_company_upper"]
        if c == "" or c == "(BLANK)":
            return "blank_company"
        if any(c.startswith(p) for p in ADMIN_PREFIXES):
            return "admin"
        return "named_client"

    df["_population"] = df.apply(classify, axis=1)

    # ── Core field mapping ──
    subj_col = find_col(df, "Subject")
    origin_col = find_col(df, "Origin")
    owner_col = find_col(df, "Manager (Owning User) (User)", "Owner")
    pod_col = find_col(df, "POD Name (Owning User) (User)", "POD Name")
    hrs_col = find_col(df, "Resolved In Hours")
    created_col = find_col(df, "Created On")
    sla_col = find_col(df, "SLA Start")
    status_col = find_col(df, "Status Reason", "Status")
    desc_col = find_col(df, "Description")
    act_subj_col = find_col(df, "Activity Subject")

    if subj_col: df["_subject"] = df[subj_col].fillna("(blank)").astype(str).str.strip()
    if origin_col: df["_origin"] = df[origin_col].fillna("(blank)").astype(str).str.strip()
    if owner_col: df["_owner"] = df[owner_col].fillna("(blank)").astype(str).str.strip()
    if pod_col: df["_pod"] = df[pod_col].fillna("(blank)").astype(str).str.strip()
    if hrs_col: df["_hours"] = safe_numeric(df[hrs_col])
    if created_col: df["_created_dt"] = safe_dt(df[created_col])
    if sla_col: df["_sla_dt"] = safe_dt(df[sla_col])
    if desc_col: df["_desc_len"] = df[desc_col].fillna("").astype(str).str.len()
    if act_subj_col: df["_act_subj_len"] = df[act_subj_col].fillna("").astype(str).str.len()

    resolved_kw = ["resolved", "closed", "cancelled", "canceled"]
    if status_col:
        df["_is_resolved"] = df[status_col].fillna("").astype(str).str.lower().apply(
            lambda x: any(kw in x for kw in resolved_kw))
    else:
        df["_is_resolved"] = True

    if "_created_dt" in df.columns and "_sla_dt" in df.columns:
        df["_triage_min"] = (df["_created_dt"] - df["_sla_dt"]).dt.total_seconds() / 60

    # ── Splits ──
    named = df[df["_population"] == "named_client"]
    blank = df[df["_population"] == "blank_company"]
    admin = df[df["_population"] == "admin"]

    print(f"Population split:")
    print(f"  Named client:    {len(named):>7,}  ({pct(len(named), len(df))})")
    print(f"  Blank company:   {len(blank):>7,}  ({pct(len(blank), len(df))})")
    print(f"  Admin:           {len(admin):>7,}  ({pct(len(admin), len(df))})")
    print(f"  Total:           {len(df):>7,}\n")

    sheets = OrderedDict()

    # ════════════════════════════════════════════════════════
    #  V1: Side-by-side population comparison
    # ════════════════════════════════════════════════════════

    def profile(subset, label):
        n = len(subset)
        rows = [{"metric": "case_count", "value": f"{n:,}"}]

        if "_hours" in subset.columns:
            hrs = subset["_hours"].dropna()
            if len(hrs) > 0:
                rows.append({"metric": "median_hours", "value": f"{hrs.median():.1f}"})
                rows.append({"metric": "mean_hours", "value": f"{hrs.mean():.1f}"})
                rows.append({"metric": "p90_hours", "value": f"{hrs.quantile(0.9):.1f}"})

        if "_is_resolved" in subset.columns:
            unres = (~subset["_is_resolved"]).sum()
            rows.append({"metric": "unresolved", "value": f"{unres:,}"})
            rows.append({"metric": "pct_unresolved", "value": pct(unres, n)})

        if "_subject" in subset.columns:
            top3 = subset["_subject"].value_counts().head(3)
            for s, c in top3.items():
                rows.append({"metric": f"top_subject: {s}", "value": f"{c:,} ({pct(c, n)})"})

        if "_origin" in subset.columns:
            top3 = subset["_origin"].value_counts().head(3)
            for o, c in top3.items():
                rows.append({"metric": f"top_origin: {o}", "value": f"{c:,} ({pct(c, n)})"})

        if "_owner" in subset.columns:
            n_owners = subset["_owner"].nunique()
            rows.append({"metric": "distinct_owners", "value": str(n_owners)})
            top1 = subset["_owner"].value_counts().head(1)
            if len(top1) > 0:
                rows.append({"metric": f"top_owner: {top1.index[0]}", "value": f"{top1.iloc[0]:,} ({pct(top1.iloc[0], n)})"})

        if "_pod" in subset.columns:
            n_pods = subset["_pod"].nunique()
            rows.append({"metric": "distinct_pods", "value": str(n_pods)})

        if "_desc_len" in subset.columns:
            fill = (subset["_desc_len"] > 0).sum()
            rows.append({"metric": "description_fill_pct", "value": pct(fill, n)})

        if "_act_subj_len" in subset.columns:
            fill = (subset["_act_subj_len"] > 5).sum()
            rows.append({"metric": "activity_subject_fill_pct", "value": pct(fill, n)})

        if "_triage_min" in subset.columns:
            valid = subset["_triage_min"][(subset["_triage_min"] > 0.5) & (subset["_triage_min"] < 10080)]
            if len(valid) > 0:
                rows.append({"metric": "triage_cases_with_gap", "value": f"{len(valid):,} ({pct(len(valid), n)})"})
                rows.append({"metric": "triage_median_min", "value": f"{valid.median():.1f}"})

        for r in rows:
            r["population"] = label
        return rows

    comparison_rows = profile(named, "Named Client")
    comparison_rows += profile(blank, "Blank Company")
    comparison_rows += profile(admin, "Admin")
    sheets["V1_PopulationComparison"] = pd.DataFrame(comparison_rows)[["population", "metric", "value"]]

    print("V1: Population comparison built")

    # ════════════════════════════════════════════════════════
    #  V2: Blank cases — subject distribution vs named client
    # ════════════════════════════════════════════════════════

    if "_subject" in df.columns:
        named_subj = named["_subject"].value_counts().reset_index()
        named_subj.columns = ["subject", "named_count"]
        named_subj["named_pct"] = (named_subj["named_count"] / len(named) * 100).round(1)

        blank_subj = blank["_subject"].value_counts().reset_index()
        blank_subj.columns = ["subject", "blank_count"]
        blank_subj["blank_pct"] = (blank_subj["blank_count"] / len(blank) * 100).round(1)

        admin_subj = admin["_subject"].value_counts().reset_index()
        admin_subj.columns = ["subject", "admin_count"]
        admin_subj["admin_pct"] = (admin_subj["admin_count"] / len(admin) * 100).round(1)

        merged = named_subj.merge(blank_subj, on="subject", how="outer").merge(admin_subj, on="subject", how="outer")
        merged = merged.fillna(0).sort_values("named_count", ascending=False)

        # Similarity signal: if blank distribution looks like named, they're client cases
        sheets["V2_SubjectDistribution"] = merged
        print("V2: Subject distribution built")

    # ════════════════════════════════════════════════════════
    #  V3: Blank cases — origin distribution (Email = likely client)
    # ════════════════════════════════════════════════════════

    if "_origin" in df.columns:
        origin_rows = []
        for label, subset in [("Named Client", named), ("Blank Company", blank), ("Admin", admin)]:
            dist = subset["_origin"].value_counts()
            for origin, count in dist.items():
                origin_rows.append({
                    "population": label, "origin": origin,
                    "count": count, "pct": round(100 * count / len(subset), 1),
                })
        sheets["V3_OriginDistribution"] = pd.DataFrame(origin_rows)
        print("V3: Origin distribution built")

    # ════════════════════════════════════════════════════════
    #  V4: Blank cases — owner/pod overlap with named clients
    # ════════════════════════════════════════════════════════

    if "_owner" in df.columns:
        named_owners = set(named["_owner"].unique())
        blank_owners = set(blank["_owner"].unique())
        admin_owners = set(admin["_owner"].unique())

        overlap_rows = [
            {"metric": "Named client distinct owners", "value": len(named_owners)},
            {"metric": "Blank company distinct owners", "value": len(blank_owners)},
            {"metric": "Admin distinct owners", "value": len(admin_owners)},
            {"metric": "Blank owners also in Named", "value": len(blank_owners & named_owners)},
            {"metric": "Blank owners ONLY in Blank", "value": len(blank_owners - named_owners - admin_owners)},
            {"metric": "Overlap pct (blank owners in named)", "value": f"{pct(len(blank_owners & named_owners), len(blank_owners))}"},
        ]

        if "_pod" in df.columns:
            named_pods = set(named["_pod"].unique())
            blank_pods = set(blank["_pod"].unique())
            overlap_rows.extend([
                {"metric": "Named client distinct pods", "value": len(named_pods)},
                {"metric": "Blank company distinct pods", "value": len(blank_pods)},
                {"metric": "Blank pods also in Named", "value": len(blank_pods & named_pods)},
                {"metric": "Overlap pct (blank pods in named)", "value": f"{pct(len(blank_pods & named_pods), len(blank_pods))}"},
            ])

        sheets["V4_OwnerPodOverlap"] = pd.DataFrame(overlap_rows)
        print("V4: Owner/pod overlap built")

    # ════════════════════════════════════════════════════════
    #  V5: Blank cases — resolution time comparison
    # ════════════════════════════════════════════════════════

    if "_hours" in df.columns:
        time_rows = []
        for label, subset in [("Named Client", named), ("Blank Company", blank), ("Admin", admin)]:
            hrs = subset["_hours"].dropna()
            if len(hrs) == 0:
                continue
            time_rows.append({
                "population": label,
                "cases_with_hours": len(hrs),
                "p10": round(hrs.quantile(0.1), 1),
                "p25": round(hrs.quantile(0.25), 1),
                "median": round(hrs.median(), 1),
                "p75": round(hrs.quantile(0.75), 1),
                "p90": round(hrs.quantile(0.9), 1),
                "max": round(hrs.max(), 1),
                "mean": round(hrs.mean(), 1),
            })
        sheets["V5_ResolutionComparison"] = pd.DataFrame(time_rows)
        print("V5: Resolution time comparison built")

    # ════════════════════════════════════════════════════════
    #  V6: Blank cases — hourly/DOW pattern (system vs human)
    # ════════════════════════════════════════════════════════

    if "_created_dt" in df.columns:
        pattern_rows = []
        for label, subset in [("Named Client", named), ("Blank Company", blank), ("Admin", admin)]:
            valid = subset.dropna(subset=["_created_dt"])
            if len(valid) == 0:
                continue
            hours = valid["_created_dt"].dt.hour
            # Business hours (8-17) vs off-hours
            biz = ((hours >= 8) & (hours < 17)).sum()
            off = len(hours) - biz
            # Weekend
            dow = valid["_created_dt"].dt.dayofweek
            weekend = (dow >= 5).sum()
            weekday = len(dow) - weekend

            pattern_rows.append({
                "population": label,
                "total_with_date": len(valid),
                "business_hours_pct": round(100 * biz / len(valid), 1),
                "off_hours_pct": round(100 * off / len(valid), 1),
                "weekday_pct": round(100 * weekday / len(valid), 1),
                "weekend_pct": round(100 * weekend / len(valid), 1),
                "peak_hour": int(hours.mode().iloc[0]) if len(hours.mode()) > 0 else "",
            })

            # Hourly breakdown for this population
            hourly = hours.value_counts().sort_index()
            for h, c in hourly.items():
                pattern_rows.append({
                    "population": f"{label} (hourly)",
                    "total_with_date": f"{int(h):02d}:00",
                    "business_hours_pct": c,
                    "off_hours_pct": round(100 * c / len(valid), 1),
                })

        sheets["V6_CreationPatterns"] = pd.DataFrame(pattern_rows)
        print("V6: Creation patterns built")

    # ════════════════════════════════════════════════════════
    #  V7: Decision matrix — the answer sheet
    # ════════════════════════════════════════════════════════

    decision_rows = [
        {
            "test": "1. Subject distribution similarity",
            "question": "Do blank cases have the same subject mix as named clients?",
            "what_to_look_for": "Compare V2. If top subjects match (Research, NSF, New Account, etc) → client-like. If dominated by one unusual subject → system artifact.",
            "finding": "[FILL AFTER RUNNING]",
            "verdict": "",
        },
        {
            "test": "2. Origin channel",
            "question": "Are blank cases coming from Email (client) or Report/System?",
            "what_to_look_for": "V3. If >70% Email origin → strongly suggests client cases. If Report-dominated → system-generated.",
            "finding": "[FILL AFTER RUNNING]",
            "verdict": "",
        },
        {
            "test": "3. Owner/pod overlap",
            "question": "Are the same bankers handling blank cases as named client cases?",
            "what_to_look_for": "V4. If >80% of blank-case owners also handle named clients → same workforce, same workflow.",
            "finding": "[FILL AFTER RUNNING]",
            "verdict": "",
        },
        {
            "test": "4. Resolution time profile",
            "question": "Do blank cases resolve at similar speed to named clients?",
            "what_to_look_for": "V5. If median within 2x of named client median → similar complexity. If 10x faster → auto-resolved system cases.",
            "finding": "[FILL AFTER RUNNING]",
            "verdict": "",
        },
        {
            "test": "5. Creation time pattern",
            "question": "Are blank cases created during business hours by humans?",
            "what_to_look_for": "V6. If business-hours pattern matches named clients → human-created. If flat 24/7 or midnight spikes → automated.",
            "finding": "[FILL AFTER RUNNING]",
            "verdict": "",
        },
        {
            "test": "6. Triage delay presence",
            "question": "Do blank cases have SLA Start ≠ Created On (email triage gap)?",
            "what_to_look_for": "V1. If triage gap % is similar to named clients → email-originated client cases.",
            "finding": "[FILL AFTER RUNNING]",
            "verdict": "",
        },
        {
            "test": "=== OVERALL VERDICT ===",
            "question": "Should blank-company cases be included in client-facing analysis?",
            "what_to_look_for": "If 4+ of 6 tests point to 'client-like' → YES, include. If mixed → include but flag. If 4+ point to 'system' → keep excluded.",
            "finding": "[FILL AFTER REVIEWING V1-V6]",
            "verdict": "",
        },
    ]
    sheets["V7_DecisionMatrix"] = pd.DataFrame(decision_rows)
    print("V7: Decision matrix built\n")

    # ── Write ──
    print(f"Writing: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            write_sheet(writer, name[:31], sdf)

    print(f"\nDone in {(datetime.datetime.now() - start).total_seconds():.1f}s")
    print(f"\n{'='*60}")
    print("NEXT STEPS:")
    print("  1. Open wab_blank_case_validation.xlsx")
    print("  2. Review V1-V6 sheets")
    print("  3. Fill in V7_DecisionMatrix findings and verdicts")
    print("  4. Share with team to finalize population logic")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
