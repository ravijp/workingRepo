"""
WAB Subject Sub-Segmentation View — Weekly Call & Leadership Artifact
======================================================================
Produces a clean Excel workbook with:
  Sheet 1: Master subject table (top 15, ranked by volume)
  Sheet 2: Research breakdown (5 clusters + keyword rationale)
  Sheet 3: Account Maintenance breakdown (if warranted by data)
  Sheet 4: Keyword rationale (full list per cluster — for Chris validation)
  Sheet 5: Methodology note (for the email to Chris)

Run on VDI. Share the workbook + email draft with Chris for validation
before presenting to Bob.

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
from collections import OrderedDict, Counter

import pandas as pd
import numpy as np

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "wab_subject_subseg_view.xlsx")
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

def safe_dt(s):
    if pd.api.types.is_datetime64_any_dtype(s): return s
    try: return pd.to_datetime(s, errors="coerce")
    except: return pd.Series([pd.NaT]*len(s), index=s.index)

def safe_num(s):
    if pd.api.types.is_numeric_dtype(s): return s
    return pd.to_numeric(s, errors="coerce")

def write_sheet(writer, name, df):
    if df is None or df.empty:
        df = pd.DataFrame({"note": ["No data"]})
    sn = name[:31]
    df.to_excel(writer, sheet_name=sn, index=False, freeze_panes=(1, 0))
    ws = writer.sheets[sn]
    for col_cells in ws.columns:
        mx = max(len(str(cell.value or "")) for cell in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(mx + 2, 55)


# ═══════════════════════════════════════════════════════════
#  KEYWORD DEFINITIONS (the core assumption to validate)
# ═══════════════════════════════════════════════════════════

# Each cluster has: keywords, rationale (why these keywords), source
RESEARCH_CLUSTERS = OrderedDict([
    ("Payment Research", {
        "keywords": ["payment", "ach", "wire", "transfer", "deposit",
                     "credit", "debit", "transaction", "posted", "posting",
                     "return", "reversal", "refund"],
        "rationale": "Chris stated Research is 'typically payment research.' These keywords "
                     "capture the core payment lifecycle: inflows (deposit, credit, ach), "
                     "outflows (wire, transfer, debit), and corrections (return, reversal, refund).",
        "source": "Chris March 23: 'Typically payment research; a lot different things may go into it.'",
    }),
    ("Check/Item Research", {
        "keywords": ["check", "cheque", "item", "image", "copy",
                     "front", "back", "clearing"],
        "rationale": "Check image requests and item research are a distinct sub-workflow. "
                     "Bankers pull check images from FIS/IBS. 'Front' and 'back' refer to "
                     "check image sides. 'Clearing' indicates item clearing research.",
        "source": "Inferred from Activity Subject patterns (e.g. 'check copy request', 'item research').",
    }),
    ("Statement/Balance Inquiry", {
        "keywords": ["statement", "balance", "reconcil", "ledger",
                     "interest", "rate", "fee"],
        "rationale": "Statement requests, balance inquiries, and reconciliation support are a "
                     "recurring client need. 'Interest' and 'rate' questions often accompany "
                     "balance inquiries. 'Fee' research is a common follow-up.",
        "source": "Inferred from Activity Subject patterns (e.g. 'statement request', 'balance inquiry').",
    }),
    ("Account Updates & Other", {
        "keywords": ["address", "signer", "name change", "update", "modify",
                     "amendment", "tin", "ein", "ssn", "tax", "w-9", "w9",
                     "certification", "fraud", "dispute", "unauthorized",
                     "suspicious", "positive pay", "stop payment",
                     "new account", "onboard", "setup", "opening"],
        "rationale": "Catch-all for non-payment research: account maintenance items (address, signer, "
                     "TIN), fraud/dispute cases, and onboarding tasks that were miscategorized as Research "
                     "instead of their proper subject.",
        "source": "Residual category. These cases may reflect mis-categorization at case creation.",
    }),
    ("No Text (Unclassifiable)", {
        "keywords": [],
        "rationale": "Cases where both Description and Activity Subject are empty or under 10 characters. "
                     "Cannot be classified without text. This is ~1% of Research cases.",
        "source": "Data-driven: insufficient text for any keyword match.",
    }),
    ("Other/Uncategorized", {
        "keywords": [],
        "rationale": "Cases that have text (Description or Activity Subject) but no keywords matched. "
                     "This is ~36% of Research — the largest gap. These cases likely contain domain-specific "
                     "language not captured by our keyword lists. Top keywords from this bucket should be "
                     "reviewed with Chris to expand the classification.",
        "source": "Residual. Requires SME review to improve classification coverage.",
    }),
])

# Account Maintenance sub-segmentation (lighter — to check if warranted)
ACCT_MAINT_CLUSTERS = OrderedDict([
    ("Address/Signer Updates", {
        "keywords": ["address", "signer", "authorized", "officer", "board member",
                     "name change", "title", "beneficiary"],
    }),
    ("TIN/Tax/Certification", {
        "keywords": ["tin", "ein", "ssn", "tax", "w-9", "w9", "certification",
                     "irs", "backup withholding"],
    }),
    ("CD/Rate Related", {
        "keywords": ["cd", "certificate", "maturity", "rate", "renewal", "interest"],
    }),
    ("Fee/Adjustment", {
        "keywords": ["fee", "waive", "refund", "adjustment", "credit", "reversal", "nsf"],
    }),
    ("Other Maintenance", {
        "keywords": [],
    }),
])


def classify_research(desc, act_subj):
    combined = f"{desc} {act_subj}".lower()
    for cluster_name, cfg in RESEARCH_CLUSTERS.items():
        if not cfg["keywords"]:
            continue
        if any(kw in combined for kw in cfg["keywords"]):
            return cluster_name
    if len(combined.strip()) < 10:
        return "No Text (Unclassifiable)"
    return "Other/Uncategorized"


def classify_acct_maint(desc, act_subj):
    combined = f"{desc} {act_subj}".lower()
    for cluster_name, cfg in ACCT_MAINT_CLUSTERS.items():
        if not cfg["keywords"]:
            continue
        if any(kw in combined for kw in cfg["keywords"]):
            return cluster_name
    if len(combined.strip()) < 10:
        return "No Text"
    return "Other Maintenance"


def top_keywords_from_texts(texts, top_n=20):
    stop = {"the","and","for","that","this","with","from","your","have","are",
            "was","were","been","has","had","but","not","you","all","can",
            "will","about","which","their","them","into","also","our","out",
            "would","could","should","need","account","case","email","bank",
            "thank","thanks","hello","dear","regards","please","sent","received",
            "attached","fyi","following","below","western","alliance"}
    words = Counter()
    for t in texts:
        if not t or len(str(t)) < 3: continue
        for w in re.findall(r"[a-z]{3,}", str(t).lower()):
            if w not in stop:
                words[w] += 1
    return words.most_common(top_n)


def main():
    start = datetime.datetime.now()
    print(f"=== Subject Sub-Segmentation View — {start.strftime('%Y-%m-%d %H:%M')} ===\n")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    df = pd.read_excel(CASE_FILE, header=0, engine="openpyxl")
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    df = df.dropna(axis=1, how="all")
    print(f"  Cases: {len(df):,}")

    # Classify
    co_col = find_col(df, "Company Name (Company) (Company)", "Company Name", "Customer")
    subj_col = find_col(df, "Subject")
    hrs_col = find_col(df, "Resolved In Hours")
    desc_col = find_col(df, "Description")
    act_col = find_col(df, "Activity Subject")
    status_col = find_col(df, "Status Reason", "Status")

    df["_company"] = df[co_col].fillna("").astype(str).str.strip() if co_col else ""
    _upper = df["_company"].str.upper()
    df["_is_admin"] = _upper.apply(lambda x: any(x.startswith(p) for p in ADMIN_PREFIXES))
    df["_is_blank"] = (_upper == "") | (_upper == "(BLANK)")
    df["_is_internal"] = df["_is_admin"]

    df["_subject"] = df[subj_col].fillna("(blank)").astype(str).str.strip() if subj_col else ""
    df["_hours"] = safe_num(df[hrs_col]) if hrs_col else np.nan
    df["_desc"] = df[desc_col].fillna("").astype(str).str.strip() if desc_col else ""
    df["_act_subj"] = df[act_col].fillna("").astype(str).str.strip() if act_col else ""

    resolved_kw = ["resolved", "closed", "cancelled", "canceled"]
    if status_col:
        df["_is_resolved"] = df[status_col].fillna("").astype(str).str.lower().apply(
            lambda x: any(k in x for k in resolved_kw))
    else:
        df["_is_resolved"] = True

    client = df[~df["_is_internal"]].copy()
    print(f"  Client inclusive: {len(client):,}\n")

    sheets = OrderedDict()

    # ═══════════════════════════════════════════════════════
    #  Sheet 1: Master Subject Table
    # ═══════════════════════════════════════════════════════
    subj_vc = client["_subject"].value_counts()
    top_subjects = subj_vc.head(15).index.tolist()
    n_total = len(client)

    master_rows = []
    for subj in top_subjects:
        s = client[client["_subject"] == subj]
        n = len(s)
        hrs = s["_hours"].dropna()
        unres = (~s["_is_resolved"]).sum()
        desc_fill = round(100 * (s["_desc"].str.len() > 0).mean(), 0)
        act_fill = round(100 * (s["_act_subj"].str.len() > 5).mean(), 0)

        # Decide if sub-segmentation is warranted
        has_subseg = ""
        if subj == "Research":
            has_subseg = "YES — see Sheet 2"
        elif subj == "Account Maintenance":
            has_subseg = "EXPLORATORY — see Sheet 3"

        master_rows.append({
            "subject": subj,
            "cases": n,
            "pct_of_total": f"{round(100 * n / n_total, 1)}%",
            "median_hrs": round(hrs.median(), 1) if len(hrs) > 0 else "",
            "p75_hrs": round(hrs.quantile(0.75), 1) if len(hrs) > 0 else "",
            "p90_hrs": round(hrs.quantile(0.9), 1) if len(hrs) > 0 else "",
            "unresolved": int(unres),
            "pct_unresolved": f"{round(100 * unres / n, 1)}%",
            "desc_fill": f"{desc_fill}%",
            "act_subj_fill": f"{act_fill}%",
            "chris_context": "",
            "sub_segmentation": has_subseg,
            "genai_opportunity": "",
        })

    # Add Chris context from March 23 notes
    chris_notes = {
        "Research": "Typically payment research; a lot different things go into it. Break into smaller pieces.",
        "CD Maintenance": "Banker sits on case till maturity date. Task itself not time-taking. By design.",
        "Signature Card": "Keith owns. Large backlog. Board member changes drive volume. Want to automate.",
        "NSF and Non-Post": "2-3 hrs/day per banker. Daily report, decision items one at a time in IBS.",
        "Close Account": "Manual checklist across IBS, BST, ACH tracker. Systems don't talk. Working to consolidate into CRM.",
        "IntraFi Maintenance": "Similar to CD maintenance. Check with James.",
        "QC Finding": "Eduardo Jacobo. Reports of prior day changes. Goal: build QC into workflow.",
    }
    for row in master_rows:
        row["chris_context"] = chris_notes.get(row["subject"], "")

    # GenAI opportunity assessment
    genai_notes = {
        "Research": "HIGH — sub-segmentation enables targeted routing. AI can classify email intent.",
        "New Account Request": "HIGH — 65% desc fill, 86% act_subj. Draft reply + missing-info detection.",
        "Account Maintenance": "HIGH — 56% desc fill, 98% act_subj. Fat tail (P90 312h). AI triage.",
        "NSF and Non-Post": "MEDIUM — already fast (1.5h). Automation of IBS decisioning, not AI.",
        "New Account Child Case": "MEDIUM — linked to parent. AI value in document review (IDP).",
        "CD Maintenance": "LOW — by design wait. Not a speed target.",
        "Close Account": "LOW — process redesign needed first. CRM consolidation.",
        "General Questions": "MEDIUM — fast median but fat tail. AI routing could help.",
        "Signature Card": "MEDIUM — backlog problem. IDP for document processing.",
        "Fraud Alert": "LOW — already fast (2.6h). Rules-based automation.",
        "IntraFi Maintenance": "LOW — by design wait like CD.",
        "Transfer": "LOW — already fast (1.3h). Straight-through automation.",
        "Statements": "LOW — fast. Proofpoint noise is 67% of email volume.",
        "Online Banking": "MEDIUM — check if self-service can absorb.",
        "QC Finding": "LOW — internal process, not client-facing.",
    }
    for row in master_rows:
        row["genai_opportunity"] = genai_notes.get(row["subject"], "")

    # Totals row
    total_cases = sum(r["cases"] for r in master_rows)
    master_rows.append({
        "subject": f"=== TOP 15 TOTAL ===",
        "cases": total_cases,
        "pct_of_total": f"{round(100 * total_cases / n_total, 1)}%",
    })

    sheets["1_MasterSubjectView"] = pd.DataFrame(master_rows)
    print("  Sheet 1: Master Subject View built")

    # ═══════════════════════════════════════════════════════
    #  Sheet 2: Research Breakdown
    # ═══════════════════════════════════════════════════════
    research = client[client["_subject"] == "Research"].copy()
    if len(research) > 0:
        research["_cluster"] = research.apply(
            lambda r: classify_research(r["_desc"], r["_act_subj"]), axis=1)

        grp = research.groupby("_cluster")
        res_rows = []
        for cluster in RESEARCH_CLUSTERS.keys():
            if cluster not in grp.groups:
                continue
            s = grp.get_group(cluster)
            n = len(s)
            hrs = s["_hours"].dropna()
            unres = (~s["_is_resolved"]).sum()
            desc_fill = round(100 * (s["_desc"].str.len() > 0).mean(), 0)

            res_rows.append({
                "cluster": cluster,
                "cases": n,
                "pct_of_research": f"{round(100 * n / len(research), 1)}%",
                "median_hrs": round(hrs.median(), 1) if len(hrs) > 0 else "",
                "p90_hrs": round(hrs.quantile(0.9), 1) if len(hrs) > 0 else "",
                "unresolved": int(unres),
                "pct_unresolved": f"{round(100 * unres / n, 1)}%",
                "desc_fill": f"{desc_fill}%",
                "classification_basis": ", ".join(RESEARCH_CLUSTERS[cluster]["keywords"][:5]) + "..." if RESEARCH_CLUSTERS[cluster]["keywords"] else "(no keywords — residual)",
            })

        # Top keywords from Other/Uncategorized (for Chris to review)
        other = research[research["_cluster"] == "Other/Uncategorized"]
        if len(other) > 0:
            other_kw = top_keywords_from_texts(other["_desc"].tolist() + other["_act_subj"].tolist(), top_n=25)
            kw_str = ", ".join(f"{w}({c})" for w, c in other_kw[:15])
            res_rows.append({
                "cluster": ">>> TOP KEYWORDS IN OTHER/UNCATEGORIZED <<<",
                "cases": "",
                "pct_of_research": "",
                "classification_basis": kw_str,
            })

        sheets["2_ResearchBreakdown"] = pd.DataFrame(res_rows)
        print(f"  Sheet 2: Research Breakdown built ({len(research):,} cases)")

    # ═══════════════════════════════════════════════════════
    #  Sheet 3: Account Maintenance Breakdown (exploratory)
    # ═══════════════════════════════════════════════════════
    acct = client[client["_subject"] == "Account Maintenance"].copy()
    if len(acct) > 0:
        acct["_cluster"] = acct.apply(
            lambda r: classify_acct_maint(r["_desc"], r["_act_subj"]), axis=1)

        grp = acct.groupby("_cluster")
        acct_rows = []
        for cluster in list(ACCT_MAINT_CLUSTERS.keys()) + ["No Text"]:
            if cluster not in grp.groups:
                continue
            s = grp.get_group(cluster)
            n = len(s)
            hrs = s["_hours"].dropna()
            unres = (~s["_is_resolved"]).sum()

            acct_rows.append({
                "cluster": cluster,
                "cases": n,
                "pct_of_acct_maint": f"{round(100 * n / len(acct), 1)}%",
                "median_hrs": round(hrs.median(), 1) if len(hrs) > 0 else "",
                "p90_hrs": round(hrs.quantile(0.9), 1) if len(hrs) > 0 else "",
                "unresolved": int(unres),
                "pct_unresolved": f"{round(100 * unres / n, 1)}%",
            })

        sheets["3_AcctMaintBreakdown"] = pd.DataFrame(acct_rows)
        print(f"  Sheet 3: Account Maintenance Breakdown built ({len(acct):,} cases)")

    # ═══════════════════════════════════════════════════════
    #  Sheet 4: Keyword Rationale (for Chris validation)
    # ═══════════════════════════════════════════════════════
    kw_rows = []
    for cluster_name, cfg in RESEARCH_CLUSTERS.items():
        kw_rows.append({
            "subject": "Research",
            "cluster": cluster_name,
            "keywords": ", ".join(cfg["keywords"]) if cfg["keywords"] else "(residual — no keywords)",
            "rationale": cfg["rationale"],
            "source": cfg["source"],
            "validation_status": "PENDING — needs Chris confirmation",
        })
    sheets["4_KeywordRationale"] = pd.DataFrame(kw_rows)
    print("  Sheet 4: Keyword Rationale built")

    # ═══════════════════════════════════════════════════════
    #  Sheet 5: Methodology + Email Draft
    # ═══════════════════════════════════════════════════════
    email_rows = [
        {"section": "METHODOLOGY", "content":
         "Sub-segmentation approach: rule-based keyword matching on Description and Activity Subject fields. "
         "For each Research case, we concatenate Description + Activity Subject into one text string, "
         "then check keyword lists in priority order. First match wins. Cases with no text (<10 chars) "
         "are marked 'No Text.' Cases with text but no keyword match become 'Other/Uncategorized.'"},
        {"section": "METHODOLOGY", "content":
         "Key limitation: Description is only 36% filled for Research cases. Activity Subject (98% filled) "
         "carries most of the classification signal. The 36% 'Other/Uncategorized' rate is largely driven "
         "by cases where both fields lack classifiable keywords."},
        {"section": "METHODOLOGY", "content":
         "For a production classifier, we would use email body text (not CRM fields) and likely an "
         "LLM-based or TF-IDF approach. This keyword method is transparent and auditable — it's designed "
         "to validate the concept with stakeholders, not to serve as the final routing logic."},
        {"section": "DRAFT EMAIL TO CHRIS", "content": "---"},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "Subject: Research Case Sub-Segmentation — Validation Request"},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "Hi Chris,"},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "Following up on our March 23 discussion where you mentioned Research cases are 'typically "
         "payment research' but 'a lot of different things go into it.' We've done an initial "
         "sub-segmentation of the 4,407 Research cases using the Description and Activity Subject "
         "fields in CRM."},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "Our approach: we classified each Research case using keyword matching on the text fields. "
         "For example, cases mentioning 'payment,' 'ACH,' 'wire,' 'transfer,' 'deposit,' etc. are "
         "classified as Payment Research. Cases mentioning 'check,' 'image,' 'copy' go to Check/Item "
         "Research, and so on. The full keyword list and rationale are in the attached workbook (Sheet 4)."},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "Results:\n"
         "  - Payment Research: ~39% (1,724 cases, 5.5h median)\n"
         "  - Other/Uncategorized: ~36% (1,603 cases — text exists but no keywords matched)\n"
         "  - Check/Item Research: ~13% (568 cases, 4.4h median)\n"
         "  - Statement/Balance Inquiry: ~6% (270 cases, 1.6h median)\n"
         "  - Account Updates & Other: ~5% (201 cases, 4.0h median)\n"
         "  - No Text: ~1% (41 cases — Description and Activity Subject both empty)"},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "A few questions for you:\n"
         "  1. Do these categories match how your team thinks about Research cases?\n"
         "  2. The 'Other/Uncategorized' bucket is 36% — what types of research cases "
         "might we be missing? I've included the top keywords from this bucket in Sheet 2.\n"
         "  3. Are there any categories that should be split further or merged?\n"
         "  4. For the weekly sync and leadership presentation, we'd like to show this breakdown. "
         "Any concerns with the framing?"},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "The workbook is attached. Sheet 1 has the master subject view, Sheet 2 has the Research "
         "breakdown with sample keywords, and Sheet 4 has the full keyword rationale.\n\n"
         "Happy to walk through this on a quick call if that's easier.\n\nBest,\nRavi"},
    ]
    sheets["5_Methodology_EmailDraft"] = pd.DataFrame(email_rows)
    print("  Sheet 5: Methodology + Email Draft built")

    # ── Write ──
    print(f"\nWriting: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            write_sheet(writer, name[:31], sdf)

    print(f"Done in {(datetime.datetime.now() - start).total_seconds():.1f}s")


if __name__ == "__main__":
    main()
