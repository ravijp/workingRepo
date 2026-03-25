"""
WAB Subject Sub-Segmentation View — Weekly Call & Leadership Artifact
======================================================================
Produces a clean Excel workbook with:
  Sheet 1: Master subject table (top 15, ranked by volume)
  Sheet 2: Research breakdown (data-driven clusters using WAB-specific terms)
  Sheet 3: Account Maintenance breakdown
  Sheet 4: New Account Request breakdown
  Sheet 5: General Questions breakdown
  Sheet 6: Close Account breakdown
  Sheet 7: Keyword rationale (full list per cluster — for Chris validation)
  Sheet 8: Top keywords per subject (data diagnostic for refining clusters)
  Sheet 9: Methodology + Email Draft

Run on VDI. Share the workbook + email draft with Chris for validation.

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


# ═══════════════════════════════════════════════════════════
#  UTILITIES
# ═══════════════════════════════════════════════════════════

def norm_col(name):
    if not isinstance(name, str): return ""
    return re.sub(r"\s+", " ", name.strip().lower())

def find_col(df, *candidates):
    lookup = {norm_col(c): c for c in df.columns}
    for cand in candidates:
        normed = norm_col(cand)
        if normed in lookup: return lookup[normed]
    for cand in candidates:
        normed = norm_col(cand)
        for k, v in lookup.items():
            if normed in k or k in normed: return v
    return None

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
        ws.column_dimensions[col_cells[0].column_letter].width = min(mx + 3, 60)

def top_keywords(texts, top_n=30):
    """Extract top keywords from a list of texts, excluding WAB noise words."""
    stop = {
        "the","and","for","that","this","with","from","your","have","are",
        "was","were","been","has","had","but","not","you","all","can",
        "will","about","which","their","them","into","also","our","out",
        "would","could","should","need","may","just","get","got","per",
        "via","please","thank","thanks","hello","dear","regards","sincerely",
        "sent","received","fyi","following","below","above","let","its",
        "who","how","when","where","what","why","any","each","more","some",
        "very","been","being","other","only","same","than","then","there",
        "these","those","such","both","does","doing","done","did","make",
        "made","take","took","give","gave","like","know","see","way",
        # WAB noise — security banners, email formatting
        "external","message","caution","originated","outside","organization",
        "click","links","open","attachments","unless","recognize","sender",
        "safe","content","secure","proofpoint","encrypted",
        # Generic banking terms too broad to classify
        "bank","account","case","email","western","alliance","banker",
        "client","customer","company","number","date","time","call","office",
    }
    words = Counter()
    for t in texts:
        if not t or len(str(t)) < 3: continue
        for w in re.findall(r"[a-z]{3,}", str(t).lower()):
            if w not in stop:
                words[w] += 1
    return words.most_common(top_n)


# ═══════════════════════════════════════════════════════════
#  CLUSTER DEFINITIONS — WAB-specific terminology
# ═══════════════════════════════════════════════════════════

# Research sub-segments
RESEARCH_CLUSTERS = OrderedDict([
    ("Payment / ACH / Wire Research", {
        "keywords": ["payment", "ach", "wire", "transfer", "deposit", "credit",
                     "debit", "transaction", "posted", "posting", "return",
                     "reversal", "refund", "incoming", "outgoing", "funds",
                     "remittance", "payee", "originator", "batch", "aq2"],
        "rationale": "Chris stated Research is 'typically payment research.' These keywords "
                     "capture the core payment lifecycle including WAB-specific terms (aq2 = "
                     "ACH processing system). Covers inflows, outflows, and corrections.",
        "source": "Chris March 23: 'Typically payment research.'",
    }),
    ("Check / Item Research", {
        "keywords": ["check", "cheque", "item", "image", "copy", "front", "back",
                     "clearing", "cashier", "teller", "draft", "stop payment",
                     "stale", "void", "endorsement"],
        "rationale": "Check image requests and item research. Bankers pull images from "
                     "FIS/IBS. Includes stop payments, stale-dated checks, endorsement issues.",
        "source": "Inferred from Activity Subject patterns + banking operations.",
    }),
    ("Lockbox / HOA Deposit Research", {
        "keywords": ["lockbox", "hoa", "homeowner", "homeowners", "association",
                     "dues", "assessment", "coupon", "remit", "bulk",
                     "pmc", "cmc", "management company"],
        "rationale": "HOA-specific payment research: lockbox processing, association dues, "
                     "PMC/CMC deposit inquiries. These are the core HOA banking product and "
                     "a significant share of Research cases.",
        "source": "Top keywords in Other/Uncategorized: association(154), hoa(86), "
                  "lockbox(72), homeowners(62), pmc(50).",
    }),
    ("CD / IntraFi / Maturity Research", {
        "keywords": ["maturity", "cdars", "ics", "intrafi", "certificate",
                     "renewal", "auto-renew", "auto renew", "rollover", "term",
                     "cd ", "rate change", "rate inquiry"],
        "rationale": "CD and IntraFi product research that landed in Research instead of "
                     "CD Maintenance or IntraFi Maintenance. Maturity inquiries, rate "
                     "questions, CDARS/ICS product issues.",
        "source": "Top keywords in Other/Uncategorized: maturity(67), cdars(45). "
                  "Chris March 23: CD and IntraFi have their own subjects but some leak into Research.",
    }),
    ("Statement / Balance / Reconciliation", {
        "keywords": ["statement", "balance", "reconcil", "ledger", "interest",
                     "fee", "analysis", "monthly", "quarterly", "annual",
                     "audit", "confirmation", "letter", "verify"],
        "rationale": "Statement requests, balance confirmations, reconciliation support, "
                     "and audit/verification letters. Often triggered by HOA board reviews.",
        "source": "Inferred from Activity Subject patterns.",
    }),
    ("Notice / Correspondence Research", {
        "keywords": ["notice", "notification", "letter", "correspondence",
                     "mail", "returned mail", "undeliverable", "address",
                     "contact", "update address"],
        "rationale": "Cases involving notices sent to clients (e.g. rate change notices, "
                     "returned mail, address corrections). Distinct from account maintenance "
                     "because the trigger is a communication event, not a client request.",
        "source": "Top keywords in Other/Uncategorized: notice(68).",
    }),
    ("Account Setup / Misc Research", {
        "keywords": ["signer", "name change", "update", "modify", "amendment",
                     "tin", "ein", "ssn", "tax", "w-9", "w9", "certification",
                     "fraud", "dispute", "unauthorized", "suspicious",
                     "positive pay", "new account", "onboard", "setup", "opening",
                     "park", "hold", "restrict", "freeze", "dormant"],
        "rationale": "Catch-all for non-payment, non-check research: account setup items, "
                     "TIN/tax issues, fraud disputes, and account restrictions. Many of these "
                     "are likely mis-categorized and should have been filed under other subjects. "
                     "'Park' refers to parked/held accounts.",
        "source": "Residual category + top keyword: park(51).",
    }),
])

# Account Maintenance sub-segments
ACCT_MAINT_CLUSTERS = OrderedDict([
    ("Address / Signer / Officer Updates", {
        "keywords": ["address", "signer", "authorized", "officer", "board member",
                     "name change", "title", "beneficiary", "contact", "phone",
                     "power of attorney", "poa", "trustee", "resolution"],
        "rationale": "Account ownership and contact changes. Board member changes for HOAs "
                     "are a significant driver (Chris confirmed for Signature Card — same "
                     "pattern applies to Account Maintenance).",
    }),
    ("TIN / Tax / Certification / W-9", {
        "keywords": ["tin", "ein", "ssn", "tax", "w-9", "w9", "certification",
                     "irs", "backup withholding", "recertif", "1099", "1042"],
        "rationale": "Tax ID management. TIN is the #1 document term in emails (258 mentions). "
                     "Onboarding nudge for missing TIN is the clear first automation target.",
    }),
    ("CD / Rate / Interest Maintenance", {
        "keywords": ["cd", "certificate", "maturity", "rate", "renewal", "interest",
                     "cdars", "ics", "intrafi", "rollover", "term", "yield",
                     "apy", "apr"],
        "rationale": "CD product maintenance that landed under Account Maintenance instead of "
                     "CD Maintenance. Includes rate inquiries and IntraFi product changes.",
    }),
    ("Fee / Adjustment / Refund", {
        "keywords": ["fee", "waive", "refund", "adjustment", "credit", "reversal",
                     "nsf", "overdraft", "service charge", "analysis fee",
                     "monthly fee", "maintenance fee"],
        "rationale": "Fee-related maintenance: waivers, refunds, adjustments. NSF fees that "
                     "appear here are fee disputes, not the NSF decisioning workflow.",
    }),
    ("Lockbox / HOA Product Maintenance", {
        "keywords": ["lockbox", "hoa", "homeowner", "association", "dues",
                     "assessment", "coupon", "pmc", "cmc", "management company",
                     "bulk", "sweep", "zba"],
        "rationale": "HOA-specific product maintenance: lockbox setup/changes, association "
                     "account configurations, sweep/ZBA setups for PMC structures.",
    }),
    ("Online Banking / Access / Portal", {
        "keywords": ["online", "portal", "login", "password", "access", "user",
                     "enroll", "token", "authentication", "mobile", "app",
                     "digital", "bst", "connectlive"],
        "rationale": "Digital channel maintenance: online banking access, portal enrollment, "
                     "BST (online banking system) issues, ConnectLive platform questions.",
    }),
    ("Account Open / Close / Status", {
        "keywords": ["open", "close", "closing", "new account", "dormant",
                     "inactive", "restrict", "freeze", "hold", "status",
                     "reactivat", "convert"],
        "rationale": "Account lifecycle changes: opening, closing, status changes. Cases "
                     "that are too small for their own subject category.",
    }),
])

# New Account Request sub-segments
NEW_ACCT_CLUSTERS = OrderedDict([
    ("HOA / Association New Account", {
        "keywords": ["hoa", "homeowner", "association", "community", "condo",
                     "townhome", "pmc", "cmc", "management company", "board",
                     "governing", "declaration"],
        "rationale": "New HOA accounts — the core product. PMC onboarding a new "
                     "community association.",
    }),
    ("Document Collection / Missing Docs", {
        "keywords": ["document", "missing", "need", "require", "submit", "pending",
                     "outstanding", "provide", "signed", "signature", "form",
                     "certification", "tin", "ein", "w-9", "w9", "resolution",
                     "articles", "bylaws", "operating agreement"],
        "rationale": "Document-gathering phase of new account setup. Missing-info detection "
                     "is a prime GenAI target here (Chris confirmed onboarding nudge is acceptable).",
    }),
    ("CD / ICS / CDARS / IntraFi New Account", {
        "keywords": ["cd", "certificate", "cdars", "ics", "intrafi", "deposit",
                     "placement", "term", "maturity", "rate"],
        "rationale": "New CD, ICS, or CDARS account setup — product-specific onboarding.",
    }),
    ("Lockbox / Payment Setup", {
        "keywords": ["lockbox", "coupon", "remittance", "payment", "ach",
                     "wire", "sweep", "zba", "setup", "configure"],
        "rationale": "Setting up payment infrastructure for new accounts: lockbox, "
                     "ACH origination, wire templates, sweep/ZBA configurations.",
    }),
    ("ConnectLive / Online Banking Setup", {
        "keywords": ["connectlive", "online", "portal", "digital", "access",
                     "enroll", "user", "login", "bst", "mobile"],
        "rationale": "Digital channel onboarding: ConnectLive platform setup, online "
                     "banking enrollment for new accounts.",
    }),
])

# General Questions sub-segments
GEN_Q_CLUSTERS = OrderedDict([
    ("Balance / Transaction Inquiry", {
        "keywords": ["balance", "transaction", "payment", "deposit", "credit",
                     "debit", "posted", "pending", "available", "ledger",
                     "statement", "history"],
        "rationale": "Quick balance checks and transaction inquiries — the most common "
                     "general question type.",
    }),
    ("Rate / Product Inquiry", {
        "keywords": ["rate", "interest", "cd", "yield", "apy", "product",
                     "option", "offer", "pricing", "term", "maturity"],
        "rationale": "Product and rate inquiries — clients comparing options or asking "
                     "about current rates.",
    }),
    ("Fee / Charge Inquiry", {
        "keywords": ["fee", "charge", "cost", "price", "waive", "service",
                     "monthly", "analysis", "overdraft", "nsf"],
        "rationale": "Fee-related questions: what was this charge, can you waive it.",
    }),
    ("Access / Technical / Portal", {
        "keywords": ["login", "password", "access", "online", "portal",
                     "connectlive", "mobile", "app", "token", "reset",
                     "locked", "error"],
        "rationale": "Technical access questions — portal login issues, password resets.",
    }),
    ("HOA / Association Inquiry", {
        "keywords": ["hoa", "homeowner", "association", "community", "board",
                     "dues", "assessment", "pmc", "cmc", "lockbox"],
        "rationale": "General questions about HOA account structure, dues processing, "
                     "association-specific inquiries.",
    }),
])

# Close Account sub-segments
CLOSE_ACCT_CLUSTERS = OrderedDict([
    ("Standard Account Closure", {
        "keywords": ["close", "closing", "closure", "terminate", "final",
                     "remaining balance", "disburs", "last statement"],
        "rationale": "Standard account closure requests — the main workflow.",
    }),
    ("Lockbox / Product Closure", {
        "keywords": ["lockbox", "ach", "wire", "sweep", "zba", "positive pay",
                     "online banking", "bst", "connectlive", "deactivat"],
        "rationale": "Closing ancillary products/services tied to the account. Chris noted "
                     "this requires manual checklist across IBS, BST, ACH tracker — each "
                     "product must be closed separately.",
    }),
    ("Loan / Credit Payoff", {
        "keywords": ["loan", "credit", "payoff", "paid", "lien", "collateral",
                     "mortgage", "line of credit", "credit card"],
        "rationale": "Loan/credit product payoff verification before account can be closed. "
                     "Chris: 'Any loans? check paid off. Credit card → closed.'",
    }),
    ("HOA / PMC Relationship Closure", {
        "keywords": ["hoa", "homeowner", "association", "pmc", "cmc",
                     "management company", "transition", "transfer",
                     "new bank", "moving", "departing"],
        "rationale": "Full HOA relationship departure — PMC moving to another bank. "
                     "Involves multiple accounts, lockboxes, and digital access.",
    }),
])


# ═══════════════════════════════════════════════════════════
#  CLASSIFICATION ENGINE
# ═══════════════════════════════════════════════════════════

def classify(desc, act_subj, cluster_defs):
    """Classify a case into a cluster using keyword matching.
    Returns the first matching cluster name, or residual categories."""
    combined = f"{desc} {act_subj}".lower()
    # Strip external email security banners
    combined = re.sub(
        r"caution.*?originated outside.*?organization[.\s]*",
        " ", combined, flags=re.IGNORECASE | re.DOTALL
    )
    combined = re.sub(r"https?://\S+", " ", combined)

    for cluster_name, cfg in cluster_defs.items():
        if not cfg.get("keywords"):
            continue
        if any(kw in combined for kw in cfg["keywords"]):
            return cluster_name
    if len(combined.strip()) < 10:
        return "(No Text)"
    return "(Other/Uncategorized)"


def build_breakdown(client_df, subject_name, cluster_defs, parent_subject_filter=None):
    """Build a breakdown table for a subject using the given cluster definitions."""
    if parent_subject_filter:
        subset = client_df[client_df["_subject"].str.contains(parent_subject_filter, case=False, na=False)].copy()
    else:
        subset = client_df[client_df["_subject"] == subject_name].copy()

    if len(subset) == 0:
        return pd.DataFrame({"note": [f"No {subject_name} cases found"]})

    subset["_cluster"] = subset.apply(
        lambda r: classify(r["_desc"], r["_act_subj"], cluster_defs), axis=1)

    grp = subset.groupby("_cluster")
    rows = []

    # Ordered by cluster definition, then residuals
    all_clusters = list(cluster_defs.keys()) + ["(No Text)", "(Other/Uncategorized)"]
    for cluster in all_clusters:
        if cluster not in grp.groups:
            continue
        s = grp.get_group(cluster)
        n = len(s)
        hrs = s["_hours"].dropna()
        unres = (~s["_is_resolved"]).sum()
        desc_fill = round(100 * (s["_desc"].str.len() > 0).mean(), 0)

        rows.append({
            "cluster": cluster,
            "cases": n,
            "pct": f"{round(100 * n / len(subset), 1)}%",
            "median_hrs": round(hrs.median(), 1) if len(hrs) > 0 else "",
            "p90_hrs": round(hrs.quantile(0.9), 1) if len(hrs) > 0 else "",
            "unresolved": int(unres),
            "pct_unresolved": f"{round(100 * unres / n, 1)}%",
            "desc_fill": f"{desc_fill}%",
            "classification_basis": ", ".join(cluster_defs.get(cluster, {}).get("keywords", [])[:6]) + ("..." if len(cluster_defs.get(cluster, {}).get("keywords", [])) > 6 else "") if cluster in cluster_defs else "(residual)",
        })

    # Top keywords from Other/Uncategorized
    other = subset[subset["_cluster"] == "(Other/Uncategorized)"]
    if len(other) > 0:
        other_kw = top_keywords(other["_desc"].tolist() + other["_act_subj"].tolist(), top_n=20)
        kw_str = ", ".join(f"{w}({c})" for w, c in other_kw)
        rows.append({
            "cluster": ">>> TOP KEYWORDS IN OTHER <<<",
            "classification_basis": kw_str,
        })

    return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════

def main():
    start = datetime.datetime.now()
    print(f"=== Subject Sub-Segmentation View — {start.strftime('%Y-%m-%d %H:%M')} ===\n")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    df = pd.read_excel(CASE_FILE, header=0, engine="openpyxl")
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    df = df.dropna(axis=1, how="all")
    print(f"  Cases: {len(df):,}")

    # Column mapping
    co_col = find_col(df, "Company Name (Company) (Company)", "Company Name", "Customer")
    subj_col = find_col(df, "Subject")
    hrs_col = find_col(df, "Resolved In Hours")
    desc_col = find_col(df, "Description")
    act_col = find_col(df, "Activity Subject")
    status_col = find_col(df, "Status Reason", "Status")

    df["_company"] = df[co_col].fillna("").astype(str).str.strip() if co_col else ""
    _upper = df["_company"].str.upper()
    df["_is_internal"] = _upper.apply(lambda x: any(x.startswith(p) for p in ADMIN_PREFIXES))
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
    n_total = len(client)
    print(f"  Client inclusive: {n_total:,}\n")

    sheets = OrderedDict()

    # ═══════════════════════════════════════════════════════
    #  Sheet 1: Master Subject Table
    # ═══════════════════════════════════════════════════════
    subj_vc = client["_subject"].value_counts()
    top_subjects = subj_vc.head(15).index.tolist()

    chris_notes = {
        "Research": "Typically payment research; a lot different things go into it. Break into smaller pieces.",
        "CD Maintenance": "Banker sits on case till maturity date. Task itself not time-taking. By design.",
        "Signature Card": "Keith owns. Large backlog. Board member changes drive volume. Want to automate.",
        "NSF and Non-Post": "2-3 hrs/day per banker. Daily report, decision items one at a time in IBS.",
        "Close Account": "Manual checklist across IBS, BST, ACH tracker. Systems don't talk. Consolidating into CRM.",
        "IntraFi Maintenance": "Similar to CD maintenance. Check with James.",
        "QC Finding": "Eduardo Jacobo. Reports of prior day changes. Goal: build QC into workflow.",
    }

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

    subseg_subjects = {
        "Research": "YES — see Sheet 2",
        "Account Maintenance": "YES — see Sheet 3",
        "New Account Request": "YES — see Sheet 4",
        "General Questions": "YES — see Sheet 5",
        "Close Account": "YES — see Sheet 6",
    }

    master_rows = []
    for subj in top_subjects:
        s = client[client["_subject"] == subj]
        n = len(s)
        hrs = s["_hours"].dropna()
        unres = (~s["_is_resolved"]).sum()
        desc_fill = round(100 * (s["_desc"].str.len() > 0).mean(), 0)
        act_fill = round(100 * (s["_act_subj"].str.len() > 5).mean(), 0)

        master_rows.append({
            "subject": subj, "cases": n,
            "pct_of_total": f"{round(100 * n / n_total, 1)}%",
            "median_hrs": round(hrs.median(), 1) if len(hrs) > 0 else "",
            "p75_hrs": round(hrs.quantile(0.75), 1) if len(hrs) > 0 else "",
            "p90_hrs": round(hrs.quantile(0.9), 1) if len(hrs) > 0 else "",
            "unresolved": int(unres),
            "pct_unresolved": f"{round(100 * unres / n, 1)}%",
            "desc_fill": f"{desc_fill}%",
            "act_subj_fill": f"{act_fill}%",
            "chris_context": chris_notes.get(subj, ""),
            "genai_opportunity": genai_notes.get(subj, ""),
            "sub_segmentation": subseg_subjects.get(subj, ""),
        })

    total_cases = sum(r["cases"] for r in master_rows)
    master_rows.append({
        "subject": "=== TOP 15 TOTAL ===",
        "cases": total_cases,
        "pct_of_total": f"{round(100 * total_cases / n_total, 1)}%",
    })

    sheets["1_MasterSubjectView"] = pd.DataFrame(master_rows)
    print("  Sheet 1: Master Subject View")

    # ═══════════════════════════════════════════════════════
    #  Sheets 2-6: Subject Breakdowns
    # ═══════════════════════════════════════════════════════
    breakdowns = [
        ("2_ResearchBreakdown", "Research", RESEARCH_CLUSTERS, None),
        ("3_AcctMaintBreakdown", "Account Maintenance", ACCT_MAINT_CLUSTERS, None),
        ("4_NewAcctReqBreakdown", "New Account Request", NEW_ACCT_CLUSTERS, None),
        ("5_GenQuestBreakdown", "General Questions", GEN_Q_CLUSTERS, None),
        ("6_CloseAcctBreakdown", "Close Account", CLOSE_ACCT_CLUSTERS, "Clos"),
    ]

    for sheet_name, subject, clusters, filter_pattern in breakdowns:
        bd = build_breakdown(client, subject, clusters, filter_pattern)
        sheets[sheet_name] = bd
        n_cases = client[client["_subject"] == subject].shape[0] if not filter_pattern else client[client["_subject"].str.contains(filter_pattern, case=False, na=False)].shape[0]
        print(f"  {sheet_name}: {subject} ({n_cases:,} cases)")

    # ═══════════════════════════════════════════════════════
    #  Sheet 7: Keyword Rationale (for Chris)
    # ═══════════════════════════════════════════════════════
    kw_rows = []
    all_cluster_sets = [
        ("Research", RESEARCH_CLUSTERS),
        ("Account Maintenance", ACCT_MAINT_CLUSTERS),
        ("New Account Request", NEW_ACCT_CLUSTERS),
        ("General Questions", GEN_Q_CLUSTERS),
        ("Close Account", CLOSE_ACCT_CLUSTERS),
    ]
    for subj_name, cluster_defs in all_cluster_sets:
        for cluster_name, cfg in cluster_defs.items():
            kw_rows.append({
                "subject": subj_name,
                "cluster": cluster_name,
                "keywords": ", ".join(cfg.get("keywords", [])) or "(residual)",
                "rationale": cfg.get("rationale", ""),
                "source": cfg.get("source", ""),
                "validation": "PENDING — needs Chris confirmation",
            })
    sheets["7_KeywordRationale"] = pd.DataFrame(kw_rows)
    print("  Sheet 7: Keyword Rationale")

    # ═══════════════════════════════════════════════════════
    #  Sheet 8: Top Keywords Per Subject (diagnostic)
    # ═══════════════════════════════════════════════════════
    diag_rows = []
    for subj in top_subjects:
        s = client[client["_subject"] == subj]
        texts = s["_desc"].tolist() + s["_act_subj"].tolist()
        kws = top_keywords(texts, top_n=25)
        for word, count in kws:
            diag_rows.append({
                "subject": subj, "keyword": word, "count": count,
                "pct_of_subject_texts": f"{round(100 * count / max(len(s), 1), 1)}%",
            })
    sheets["8_TopKeywordsPerSubject"] = pd.DataFrame(diag_rows)
    print("  Sheet 8: Top Keywords Per Subject (diagnostic)")

    # ═══════════════════════════════════════════════════════
    #  Sheet 9: Methodology + Email Draft
    # ═══════════════════════════════════════════════════════
    email_rows = [
        {"section": "METHODOLOGY", "content":
         "Sub-segmentation approach: rule-based keyword matching on Description + Activity Subject fields. "
         "For each case, we concatenate the two text fields, strip external email security banners, "
         "then check keyword lists in priority order (first match wins). Keywords were selected based on: "
         "(1) Chris's March 23 characterizations, (2) top keywords extracted from the data itself, "
         "(3) WAB-specific product terminology (lockbox, CDARS, ICS, ConnectLive, BST, aq2, PMC, CMC, HOA)."},
        {"section": "METHODOLOGY", "content":
         "Five subjects are sub-segmented: Research (7 clusters), Account Maintenance (7 clusters), "
         "New Account Request (5 clusters), General Questions (5 clusters), Close Account (4 clusters). "
         "Together these cover ~17,000 cases or 44% of the client portfolio."},
        {"section": "METHODOLOGY", "content":
         "Key limitation: Description is only 36-65% filled depending on subject. Activity Subject "
         "(86-99% filled) carries most of the signal. The '(Other/Uncategorized)' rate per subject "
         "indicates keyword coverage gaps — top keywords from that bucket are shown for SME review."},
        {"section": "---", "content": ""},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "Subject: Case Subject Sub-Segmentation — Validation Before Leadership Presentation"},
        {"section": "DRAFT EMAIL TO CHRIS", "content": "Hi Chris,"},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "Following our March 23 discussion, we've sub-segmented five case subjects using keyword "
         "matching on the Description and Activity Subject fields in CRM. Before we present this to "
         "Bob, I'd like your confirmation that the categories and keyword logic make sense."},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "The approach: for each case, we look for specific keywords in the text. For example, "
         "Research cases mentioning 'payment,' 'ACH,' 'wire,' 'deposit' are classified as "
         "'Payment / ACH / Wire Research.' Cases mentioning 'lockbox,' 'HOA,' 'association,' 'PMC' "
         "become 'Lockbox / HOA Deposit Research.' We added WAB-specific terms like aq2, CDARS, ICS, "
         "ConnectLive, BST based on what we see in the data."},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "Subjects sub-segmented:\n"
         "  1. Research (4,407 cases → 7 sub-types)\n"
         "  2. Account Maintenance (3,242 cases → 7 sub-types)\n"
         "  3. New Account Request (3,665 cases → 5 sub-types)\n"
         "  4. General Questions (1,783 cases → 5 sub-types)\n"
         "  5. Close Account (1,859 cases → 4 sub-types)"},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "The attached workbook has:\n"
         "  - Sheet 1: Master view of all top 15 subjects with your March 23 context\n"
         "  - Sheets 2-6: Breakdown per subject with case counts, resolution times, and keywords used\n"
         "  - Sheet 7: Full keyword rationale — exactly which words drive each classification\n"
         "  - Sheet 8: Raw top keywords per subject (what the data actually contains)"},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "Questions for you:\n"
         "  1. Do the sub-types make sense for how your team thinks about this work?\n"
         "  2. Are there WAB-specific terms we're missing? (e.g., internal system names, "
         "product codes, workflow terms)\n"
         "  3. The 'Other/Uncategorized' buckets show what we couldn't classify — the top "
         "keywords from those buckets are listed at the bottom of each breakdown sheet. "
         "Do any of those suggest categories we should add?\n"
         "  4. For the leadership presentation, is there a different way you'd group these?\n"
         "  5. Any subjects we should NOT sub-segment (too sensitive, too noisy)?"},
        {"section": "DRAFT EMAIL TO CHRIS", "content":
         "Happy to walk through this on a quick call if that's easier.\n\nBest,\nRavi"},
    ]
    sheets["9_Methodology_EmailDraft"] = pd.DataFrame(email_rows)
    print("  Sheet 9: Methodology + Email Draft")

    # ── Write ──
    print(f"\nWriting: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            write_sheet(writer, name[:31], sdf)

    print(f"Done in {(datetime.datetime.now() - start).total_seconds():.1f}s")
    print(f"\n{'='*60}")
    print("Subjects sub-segmented: Research, Account Maintenance,")
    print("  New Account Request, General Questions, Close Account")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
