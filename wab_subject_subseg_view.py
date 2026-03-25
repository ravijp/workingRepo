"""
WAB Subject Sub-Segmentation View — Executive + Validation Artifact
====================================================================
Two-layer workbook:
  Tabs 1-5:   Executive view (Bob-ready, 3 subject breakdowns)
  Tabs 6-11:  Appendix (Chris validation, full detail, keyword rationale)

Sub-segmentation approach: keyword matching on Description + Activity Subject.
Keywords selected from actual WAB case data (keyword diagnostic, March 25).
Clusters collapsed to 3-4 per subject for exec presentation.

Dependencies: pandas, openpyxl
"""

# ┌─────────────────────────────────────────────────────────┐
CASE_FILE  = r"C:\Users\YourName\Desktop\AAB All Cases.xlsx"
OUTPUT_DIR = r"C:\Users\YourName\Desktop\wab_output"
# └─────────────────────────────────────────────────────────┘

import os, re, datetime, warnings
from collections import OrderedDict, Counter
import pandas as pd, numpy as np

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "wab_subject_subseg_view.xlsx")
ADMIN_PREFIXES = {"AAB ADMIN", "WAB ADMIN", "AAB ADMIN -", "WAB ADMIN -"}

def find_col(df, *c):
    lookup = {re.sub(r"\s+"," ",str(x).strip().lower()): x for x in df.columns}
    for cand in c:
        n = re.sub(r"\s+"," ",cand.strip().lower())
        if n in lookup: return lookup[n]
    for cand in c:
        n = re.sub(r"\s+"," ",cand.strip().lower())
        for k,v in lookup.items():
            if n in k or k in n: return v
    return None

def safe_num(s):
    if pd.api.types.is_numeric_dtype(s): return s
    return pd.to_numeric(s, errors="coerce")

def write_sheet(writer, name, df):
    if df is None or df.empty:
        df = pd.DataFrame({"note": ["No data"]})
    sn = name[:31]
    df.to_excel(writer, sheet_name=sn, index=False, freeze_panes=(1,0))
    ws = writer.sheets[sn]
    for col_cells in ws.columns:
        mx = max(len(str(cell.value or "")) for cell in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(mx+3, 60)

def top_keywords(texts, top_n=15):
    stop = {"the","and","for","that","this","with","from","your","have","are",
            "was","were","been","has","had","but","not","you","all","can",
            "will","about","which","their","them","into","also","our","out",
            "would","could","should","need","may","just","get","got","per",
            "via","please","thank","thanks","hello","dear","regards","sincerely",
            "sent","received","fyi","following","below","above","let","its",
            "who","how","when","where","what","why","any","each","more","some",
            "very","being","other","only","same","than","then","there",
            "these","those","such","both","does","doing","done","did","make",
            "made","take","took","give","gave","like","know","see","way",
            "external","message","caution","originated","outside","organization",
            "click","links","attachments","unless","recognize","sender",
            "safe","content","secure","proofpoint","encrypted",
            "bank","account","case","email","western","alliance","banker",
            "client","customer","company","number","date","time","call","office"}
    words = Counter()
    for t in texts:
        if not t or len(str(t)) < 3: continue
        for w in re.findall(r"[a-z][a-z0-9]{2,}", str(t).lower()):
            if w not in stop: words[w] += 1
    return words.most_common(top_n)


# ═══════════════════════════════════════════════════════════
#  EXEC CLUSTERS — collapsed to 3-4 per subject
# ═══════════════════════════════════════════════════════════

# Research: 4 buckets (from 7 detailed → 4 exec)
RESEARCH_EXEC = OrderedDict([
    ("Payment / ACH Investigation", {
        "keywords": ["payment research", "homeowner payment", "missing payment",
                     "returned payment", "payment failed", "payment return",
                     "ach", "wire alert", "wire", "misc debit",
                     "deposit", "return", "reversal", "refund",
                     "incoming", "outgoing", "remittance", "posted", "posting",
                     "transaction", "credit card", "debit card", "card"],
        "rationale": "Chris: 'typically payment research.' Data confirms: 'payment research'(103x), "
                     "'homeowner payment'(45x), 'missing payment'(29x). Absorbs card research (46 cases).",
    }),
    ("Check / Image Investigation", {
        "keywords": ["check research", "check copy", "check image", "missing check",
                     "check error", "checks", "check", "image request",
                     "cashier", "stop payment", "stale", "endorsement"],
        "rationale": "Data: 'check research'(76x), 'check copy'(29x), 'stop payment'(52x). "
                     "Distinct workflow — bankers pull images from FIS/IBS.",
    }),
    ("Rate / CD / Maturity Inquiry", {
        "keywords": ["rate sheet", "current rate", "cd rate", "weekly cd",
                     "interest rate sheet", "maturity notice", "cdars maturity",
                     "maturity", "cdars", "ics", "intrafi",
                     "certificate", "renewal", "rollover", "rate"],
        "rationale": "Data: 'rate sheet'(59x), 'maturity notice'(57x), 'cdars maturity'(41x). "
                     "TAXONOMY LEAKAGE: these are CD/IntraFi inquiries that landed in Research.",
    }),
    ("HOA Entity / Lockbox Research", {
        "keywords": ["homeowners association", "condominium association", "community association",
                     "owners association", "association inc", "inquiry pmc",
                     "research pmc", "research homeowner",
                     "village", "park", "ridge", "manor", "estates",
                     "hoa", "pmc", "cmc", "management company",
                     "lockbox", "lockbox file", "lockbox ach",
                     "coupon", "remit", "bulk deposit", "dues", "assessment",
                     "notice available", "notice", "correspondence",
                     "returned mail", "undeliverable", "letter"],
        "rationale": "Data: 'homeowners association'(104x), 'association inc'(88x), 'lockbox'(197x unigram). "
                     "Combines HOA entity lookups, lockbox product research, and correspondence. "
                     "All HOA-operation-specific research.",
    }),
])

# Account Maintenance: 3 buckets
ACCT_MAINT_EXEC = OrderedDict([
    ("CD / IntraFi Maturity Processing", {
        "keywords": ["maturity notice", "cdars maturity", "maturity",
                     "cdars", "ics", "intrafi", "inc month",
                     "certificate", "renewal", "rollover", "cd "],
        "rationale": "Data: 'maturity notice'(330x) is the #1 bigram — by far. "
                     "TAXONOMY LEAKAGE: 580 cases are CD/IntraFi maturity events, not maintenance. "
                     "This is the headline finding for Account Maintenance.",
    }),
    ("Lockbox File / Configuration", {
        "keywords": ["lockbox file", "lockbox", "close lockbox", "lockbox ach",
                     "file properties", "file", "coupon",
                     "validation file", "management validation", "validation",
                     "first choice", "choice property",
                     "action required", "required"],
        "rationale": "Data: 'lockbox file'(165x), 'file properties'(76x), 'validation file'(91x). "
                     "Combines lockbox configuration and document validation — both are file/product workflows.",
    }),
    ("Stop Payment Processing", {
        "keywords": ["stop payment"],
        "rationale": "Data: 'stop payment'(98x bigram). Small but operationally distinct — "
                     "single-action workflow that bankers would immediately recognize.",
    }),
])

# Close Account: 3 buckets
CLOSE_ACCT_EXEC = OrderedDict([
    ("CD / IntraFi Maturity Closure", {
        "keywords": ["maturity notice", "cdars maturity", "matured",
                     "maturity", "cdars", "ics", "intrafi",
                     "cd maturity", "certificate"],
        "rationale": "Data: 'maturity notice'(218x) is the #1 Close Account bigram. "
                     "TAXONOMY LEAKAGE: ~355 cases are CDs reaching maturity, not relationship closures. "
                     "1 in 8 Close Account cases is actually a maturity event.",
    }),
    ("HOA / Association Account Closure", {
        "keywords": ["homeowners association", "association inc",
                     "condominium association", "community association",
                     "owners association", "property owners",
                     "hoa", "pmc", "village", "park", "ridge"],
        "rationale": "Data: 'homeowners association'(202x), 'association inc'(172x). "
                     "Full HOA relationship departures — PMC leaving the bank.",
    }),
    ("Standard / Product Closure", {
        "keywords": ["account closure", "close account", "close accounts",
                     "closure request", "request close", "account closures",
                     "close acct", "closing accounts", "close bank",
                     "acct closure", "offboarding",
                     "reserve", "petty cash", "funds", "bank accounts",
                     "bank account", "cash check", "request redeem",
                     "lockbox", "ach", "wire", "sweep",
                     "online banking", "positive pay", "connectlive", "bst"],
        "rationale": "Data: 'account closure'(217x), 'close accounts'(134x). "
                     "Standard closure + reserve/sub-account closures + product closures. "
                     "Chris: manual checklist across IBS, BST, ACH tracker.",
    }),
])


# ═══════════════════════════════════════════════════════════
#  FULL DETAIL CLUSTERS (appendix — for Chris validation)
# ═══════════════════════════════════════════════════════════

RESEARCH_FULL = OrderedDict([
    ("Payment / ACH Research", {
        "keywords": ["payment research", "homeowner payment", "missing payment",
                     "returned payment", "payment failed", "payment return",
                     "ach", "wire alert", "wire", "misc debit",
                     "deposit", "return", "reversal", "refund",
                     "incoming", "outgoing", "remittance", "posted",
                     "posting", "transaction"],
    }),
    ("Check / Image Research", {
        "keywords": ["check research", "check copy", "check image", "missing check",
                     "check error", "checks", "check", "image request",
                     "cashier", "stop payment", "stale", "endorsement"],
    }),
    ("Rate Sheet / CD / Maturity Inquiry", {
        "keywords": ["rate sheet", "current rate", "cd rate", "weekly cd",
                     "interest rate sheet", "maturity notice", "cdars maturity",
                     "maturity", "cdars", "ics", "intrafi",
                     "certificate", "renewal", "rollover", "rate"],
    }),
    ("Lockbox / HOA Deposit Research", {
        "keywords": ["lockbox", "lockbox file", "lockbox ach",
                     "coupon", "remit", "bulk deposit", "dues", "assessment"],
    }),
    ("HOA / PMC Entity Research", {
        "keywords": ["homeowners association", "condominium association", "community association",
                     "owners association", "association inc", "inquiry pmc",
                     "research pmc", "research homeowner",
                     "village", "park", "ridge", "manor", "estates",
                     "hoa", "pmc", "cmc", "management company"],
    }),
    ("Notice / Correspondence", {
        "keywords": ["notice available", "notice", "correspondence",
                     "returned mail", "undeliverable", "letter"],
    }),
    ("Card Research (Credit/Debit)", {
        "keywords": ["credit card", "debit card", "card"],
    }),
])

ACCT_MAINT_FULL = OrderedDict([
    ("CD / IntraFi Maturity Processing", {
        "keywords": ["maturity notice", "cdars maturity", "maturity",
                     "cdars", "ics", "intrafi", "inc month",
                     "certificate", "renewal", "rollover", "cd "],
    }),
    ("Lockbox File / Configuration", {
        "keywords": ["lockbox file", "lockbox", "close lockbox", "lockbox ach",
                     "file properties", "file", "coupon"],
    }),
    ("Validation / Document Processing", {
        "keywords": ["validation file", "management validation", "validation",
                     "first choice", "choice property", "action required", "required"],
    }),
    ("Stop Payment Processing", {
        "keywords": ["stop payment"],
    }),
    ("Association / HOA Maintenance", {
        "keywords": ["association inc", "homeowners association", "condominium association",
                     "community association", "owners association",
                     "association month", "association account", "hoa", "pmc", "cmc"],
    }),
    ("Card Maintenance (Debit/Credit)", {
        "keywords": ["debit card", "credit card", "card"],
    }),
    ("Fee Waiver / Adjustment", {
        "keywords": ["waive return", "waive", "fee", "refund", "adjustment",
                     "service charge", "overdraft"],
    }),
    ("Current Rates / Rate Sheet", {
        "keywords": ["current rates", "rate sheet", "rates", "interest rate"],
    }),
])

NEW_ACCT_FULL = OrderedDict([
    ("HOA / Association New Account", {
        "keywords": ["homeowners association", "condominium association",
                     "community association", "owners association",
                     "association inc", "association alliance",
                     "hoa", "pmc", "cmc", "village", "park",
                     "ridge", "manor", "estates", "condominium"],
    }),
    ("Reserve / Petty Cash / Sub-Account", {
        "keywords": ["reserve account", "petty cash", "reserve",
                     "child case", "parent case", "sub account",
                     "operating", "money market"],
    }),
    ("Document / Validation Required", {
        "keywords": ["validation file", "response required", "maturity notice",
                     "document", "missing", "pending", "submit",
                     "tin", "ein", "w-9", "w9", "signed"],
    }),
    ("Bank Account / Product Setup", {
        "keywords": ["bank account", "bank accounts", "new bank",
                     "open", "acct", "account request",
                     "inc alliance", "alliance aab"],
    }),
    ("Lockbox / ACH / Wire Setup", {
        "keywords": ["lockbox", "ach", "wire", "sweep", "zba",
                     "positive pay", "online banking", "connectlive"],
    }),
])

GEN_Q_FULL = OrderedDict([
    ("Rate Sheet / Interest Rate Inquiry", {
        "keywords": ["rate sheet", "rate sheets", "interest rate",
                     "interest rates", "current rate", "rates",
                     "money market", "cd rate", "yield"],
    }),
    ("Online Banking / Digital Access", {
        "keywords": ["online banking", "portal", "login", "password",
                     "access", "connectlive", "mobile", "app", "token", "reset", "locked"],
    }),
    ("Card Inquiry (Debit/Credit)", {
        "keywords": ["debit card", "credit card", "card", "visa", "mastercard", "atm", "pin"],
    }),
    ("HOA / Association Question", {
        "keywords": ["hoa", "homeowners", "association", "community",
                     "condominium", "owners association",
                     "hoa payment", "pmc", "cmc", "lockbox"],
    }),
    ("Product / Account Inquiry", {
        "keywords": ["ics account", "cdars", "reserve account",
                     "petty cash", "money market", "savings",
                     "checking", "deposit", "balance", "statement", "account ending"],
    }),
    ("Process Support / Internal", {
        "keywords": ["process support", "quick question", "question", "help", "assist", "information"],
    }),
])

CLOSE_ACCT_FULL = OrderedDict([
    ("CD / IntraFi Maturity Closure", {
        "keywords": ["maturity notice", "cdars maturity", "matured",
                     "maturity", "cdars", "ics", "intrafi", "cd maturity", "certificate"],
    }),
    ("HOA / Association Account Closure", {
        "keywords": ["homeowners association", "association inc",
                     "condominium association", "community association",
                     "owners association", "property owners", "hoa", "pmc", "village", "park", "ridge"],
    }),
    ("Standard Account Closure Request", {
        "keywords": ["account closure", "close account", "close accounts",
                     "closure request", "request close", "account closures",
                     "close acct", "closing accounts", "close bank", "acct closure", "offboarding"],
    }),
    ("Reserve / Sub-Account Closure", {
        "keywords": ["reserve", "petty cash", "funds", "bank accounts",
                     "bank account", "cash check", "request redeem"],
    }),
    ("Lockbox / Product Closure", {
        "keywords": ["lockbox", "ach", "wire", "sweep",
                     "online banking", "positive pay", "connectlive", "bst"],
    }),
])


# ═══════════════════════════════════════════════════════════
#  CLASSIFICATION ENGINE
# ═══════════════════════════════════════════════════════════

def clean_text(desc, act_subj):
    combined = f"{desc} {act_subj}".lower()
    combined = re.sub(r"caution.*?originated outside.*?organization[.\s]*", " ",
                      combined, flags=re.IGNORECASE | re.DOTALL)
    combined = re.sub(r"https?://\S+", " ", combined)
    combined = re.sub(r"\*{2,}", " ", combined)
    return combined

def classify(desc, act_subj, cluster_defs):
    combined = clean_text(desc, act_subj)
    for cluster_name, cfg in cluster_defs.items():
        kws = cfg.get("keywords", [])
        if kws and any(kw in combined for kw in kws):
            return cluster_name
    if len(combined.strip()) < 10:
        return "(No Text)"
    return "(Unclassified)"


def build_breakdown(client_df, subject_name, cluster_defs, subject_filter=None,
                    unclassified_label="(Unclassified — requires email body text)"):
    """Build a cluster breakdown table. Returns DataFrame."""
    if subject_filter:
        subset = client_df[client_df["_subject"].str.contains(subject_filter, case=False, na=False)].copy()
    else:
        subset = client_df[client_df["_subject"] == subject_name].copy()

    if len(subset) == 0:
        return pd.DataFrame({"note": [f"No {subject_name} cases"]})

    subset["_cluster"] = subset.apply(
        lambda r: classify(r["_desc"], r["_act_subj"], cluster_defs), axis=1)

    # Merge No Text into Unclassified for exec view
    subset["_cluster"] = subset["_cluster"].replace("(No Text)", unclassified_label)
    subset["_cluster"] = subset["_cluster"].replace("(Unclassified)", unclassified_label)

    grp = subset.groupby("_cluster")
    rows = []
    # Show defined clusters first, then residual
    ordered = [c for c in cluster_defs.keys() if c in grp.groups]
    ordered += [c for c in grp.groups if c not in ordered]

    for cluster in ordered:
        s = grp.get_group(cluster)
        n = len(s)
        hrs = s["_hours"].dropna()
        unres = (~s["_is_resolved"]).sum()

        rows.append({
            "cluster": cluster,
            "cases": n,
            "pct": f"{round(100 * n / len(subset), 1)}%",
            "median_hrs": round(hrs.median(), 1) if len(hrs) > 0 else "",
            "p90_hrs": round(hrs.quantile(0.9), 1) if len(hrs) > 0 else "",
            "unresolved": int(unres),
            "pct_unresolved": f"{round(100 * unres / n, 1)}%",
        })

    return pd.DataFrame(rows)


def build_full_breakdown(client_df, subject_name, cluster_defs, subject_filter=None):
    """Full breakdown with keywords shown — for appendix/Chris validation."""
    if subject_filter:
        subset = client_df[client_df["_subject"].str.contains(subject_filter, case=False, na=False)].copy()
    else:
        subset = client_df[client_df["_subject"] == subject_name].copy()

    if len(subset) == 0:
        return pd.DataFrame({"note": [f"No {subject_name} cases"]})

    subset["_cluster"] = subset.apply(
        lambda r: classify(r["_desc"], r["_act_subj"], cluster_defs), axis=1)

    grp = subset.groupby("_cluster")
    rows = []
    all_clusters = list(cluster_defs.keys()) + ["(No Text)", "(Unclassified)"]
    for cluster in all_clusters:
        if cluster not in grp.groups: continue
        s = grp.get_group(cluster)
        n = len(s)
        hrs = s["_hours"].dropna()
        unres = (~s["_is_resolved"]).sum()
        desc_fill = round(100 * (s["_desc"].str.len() > 0).mean(), 0)

        rows.append({
            "cluster": cluster, "cases": n,
            "pct": f"{round(100 * n / len(subset), 1)}%",
            "median_hrs": round(hrs.median(), 1) if len(hrs) > 0 else "",
            "p90_hrs": round(hrs.quantile(0.9), 1) if len(hrs) > 0 else "",
            "unresolved": int(unres),
            "pct_unresolved": f"{round(100 * unres / n, 1)}%",
            "desc_fill": f"{desc_fill}%",
            "key_terms": ", ".join(cluster_defs.get(cluster, {}).get("keywords", [])[:6]) + ("..." if len(cluster_defs.get(cluster, {}).get("keywords", [])) > 6 else "") if cluster in cluster_defs else "(residual)",
        })

    # Top keywords from Unclassified
    other = subset[subset["_cluster"] == "(Unclassified)"]
    if len(other) > 0:
        kw = top_keywords(other["_desc"].tolist() + other["_act_subj"].tolist(), 15)
        rows.append({"cluster": ">>> TOP KEYWORDS IN UNCLASSIFIED <<<",
                      "key_terms": ", ".join(f"{w}({c})" for w, c in kw)})

    return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════

def main():
    start = datetime.datetime.now()
    print(f"=== Subject Sub-Segmentation View — {start.strftime('%Y-%m-%d %H:%M')} ===\n")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    df = pd.read_excel(CASE_FILE, header=0, engine="openpyxl")
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")].dropna(axis=1, how="all")

    co_col = find_col(df, "Company Name (Company) (Company)", "Company Name")
    subj_col = find_col(df, "Subject")
    hrs_col = find_col(df, "Resolved In Hours")
    desc_col = find_col(df, "Description")
    act_col = find_col(df, "Activity Subject")
    status_col = find_col(df, "Status Reason", "Status")

    df["_co"] = df[co_col].fillna("").astype(str).str.strip() if co_col else ""
    df["_is_internal"] = df["_co"].str.upper().apply(lambda x: any(x.startswith(p) for p in ADMIN_PREFIXES))
    df["_subject"] = df[subj_col].fillna("").astype(str).str.strip() if subj_col else ""
    df["_hours"] = safe_num(df[hrs_col]) if hrs_col else np.nan
    df["_desc"] = df[desc_col].fillna("").astype(str).str.strip() if desc_col else ""
    df["_act_subj"] = df[act_col].fillna("").astype(str).str.strip() if act_col else ""
    resolved_kw = ["resolved", "closed", "cancelled", "canceled"]
    df["_is_resolved"] = df[status_col].fillna("").astype(str).str.lower().apply(
        lambda x: any(k in x for k in resolved_kw)) if status_col else True

    client = df[~df["_is_internal"]].copy()
    n_total = len(client)
    print(f"  Client: {n_total:,}\n")

    sheets = OrderedDict()

    # ═══════════════════════════════════════════════════════
    #  EXEC TAB 1: Master Subject View
    # ═══════════════════════════════════════════════════════
    action_lanes = {
        "Research":            "AI-ASSIST — sub-segment into 4 types, route by email intent",
        "New Account Request": "AI-ASSIST — draft reply + missing-info detection",
        "Account Maintenance": "AI-ASSIST — fat tail + CD maturity leakage finding",
        "NSF and Non-Post":    "AUTOMATE — IBS decisioning automation, not AI",
        "CD Maintenance":      "MONITOR — by design wait time, not friction",
        "Close Account":       "REDESIGN — CRM consolidation (IBS/BST/ACH)",
        "General Questions":   "AI-ASSIST — rate sheet queries deflectable to self-service",
        "Signature Card":      "REDESIGN — backlog + IDP for document processing",
        "Fraud Alert":         "AUTOMATE — already fast (2.6h), rules-based",
        "IntraFi Maintenance": "MONITOR — by design wait like CD",
        "Transfer":            "AUTOMATE — already fast (1.3h), straight-through",
        "Statements":          "AUTOMATE — fast + 67% Proofpoint noise",
        "Online Banking":      "AI-ASSIST — check if self-service can absorb",
        "QC Finding":          "MONITOR — internal process, Eduardo Jacobo",
        "New Account Child Case": "AI-ASSIST — linked to parent, IDP for docs",
    }
    chris = {
        "Research": "Typically payment research; a lot different things go into it.",
        "CD Maintenance": "Banker sits on case till maturity date. By design.",
        "Signature Card": "Keith owns. Large backlog. Board member changes.",
        "NSF and Non-Post": "2-3 hrs/day per banker. Decision items one at a time.",
        "Close Account": "Manual checklist: IBS, BST, ACH tracker.",
        "IntraFi Maintenance": "Similar to CD maintenance.",
        "QC Finding": "Eduardo Jacobo. Prior day change reports.",
    }
    subseg_ref = {
        "Research":            "Tab 2 (4 types + unclassified)",
        "Account Maintenance": "Tab 3 (3 types + residual)",
        "Close Account":       "Tab 4 (3 types + residual)",
    }

    # Show 10 subjects chosen by actionability
    show_subjects = [
        "Research", "New Account Request", "Account Maintenance",
        "NSF and Non-Post", "New Account Child Case", "CD Maintenance",
        "Close Account", "General Questions", "Signature Card", "Fraud Alert",
    ]
    # Fallback: if a subject isn't in the data, skip it
    show_subjects = [s for s in show_subjects if s in client["_subject"].values]

    master_rows = []
    for subj in show_subjects:
        s = client[client["_subject"] == subj]
        n = len(s)
        hrs = s["_hours"].dropna()
        unres = (~s["_is_resolved"]).sum()
        master_rows.append({
            "subject": subj, "cases": n,
            "pct": f"{round(100*n/n_total,1)}%",
            "median_hrs": round(hrs.median(),1) if len(hrs) > 0 else "",
            "p90_hrs": round(hrs.quantile(0.9),1) if len(hrs) > 0 else "",
            "unresolved": int(unres),
            "pct_unresolved": f"{round(100*unres/n,1)}%",
            "action_lane": action_lanes.get(subj, ""),
            "chris_context": chris.get(subj, ""),
            "breakdown": subseg_ref.get(subj, ""),
        })

    sheets["1_SubjectView"] = pd.DataFrame(master_rows)
    print("  Tab 1: Master Subject View")

    # ═══════════════════════════════════════════════════════
    #  EXEC TAB 2: Research Breakdown (4 buckets)
    # ═══════════════════════════════════════════════════════
    sheets["2_Research"] = build_breakdown(client, "Research", RESEARCH_EXEC)
    print("  Tab 2: Research (exec)")

    # ═══════════════════════════════════════════════════════
    #  EXEC TAB 3: Account Maintenance Breakdown (3 buckets)
    # ═══════════════════════════════════════════════════════
    sheets["3_AcctMaint"] = build_breakdown(client, "Account Maintenance", ACCT_MAINT_EXEC)
    print("  Tab 3: Account Maintenance (exec)")

    # ═══════════════════════════════════════════════════════
    #  EXEC TAB 4: Close Account Breakdown (3 buckets)
    # ═══════════════════════════════════════════════════════
    sheets["4_CloseAcct"] = build_breakdown(client, "Close Account", CLOSE_ACCT_EXEC, "Clos")
    print("  Tab 4: Close Account (exec)")

    # ═══════════════════════════════════════════════════════
    #  EXEC TAB 5: Cross-Subject Findings
    # ═══════════════════════════════════════════════════════
    findings = [
        {"finding": "CD/IntraFi Maturity Taxonomy Leakage (~935 cases)",
         "detail": "CD/IntraFi maturity events appear in Account Maintenance (580 cases, 18%), "
                   "Close Account (355, 12%), and Research (271, 6%). These ~935 cases are maturity "
                   "processing that landed under the wrong subject. The pick list doesn't have a clean "
                   "path for 'CD reached maturity — action needed.'",
         "management_implication": "Either fix the subject pick list to add a maturity workflow path, or "
                                   "build an AI pre-classifier that detects maturity cases at creation time "
                                   "and routes them correctly. This alone would clean up ~935 cases/quarter.",
         "source": "Keyword matching on 'maturity notice,' 'cdars maturity,' 'matured' across 3 subjects."},
        {"finding": "Rate Sheet Queries Are Deflectable (~270+ cases)",
         "detail": "Rate sheet / interest rate inquiries are the #1 General Questions sub-type (174 cases, 10%) "
                   "and appear in Research (271 cases, 6%). These are lookups, not complex requests.",
         "management_implication": "A rate sheet self-service page or bot could deflect these entirely. "
                                   "No AI needed — just information availability. Quick win.",
         "source": "Bigrams: 'rate sheet'(98x in General Questions, 59x in Research)."},
        {"finding": "Lockbox Is a Cross-Subject Workflow",
         "detail": "Lockbox appears in Research (129 cases), Account Maintenance (482 cases), "
                   "New Account Request (setup), and Close Account (81 cases, 19.8% unresolved). "
                   "It's a product that spans the full case lifecycle.",
         "management_implication": "Lockbox cases may benefit from a dedicated subject or sub-tag. "
                                   "AI routing could identify lockbox-related cases regardless of subject.",
         "source": "Unigram 'lockbox' across all 5 sub-segmented subjects."},
        {"finding": "Unclassified Rate Is 38-41% Across All Subjects",
         "detail": "Keyword matching on CRM text fields (Description + Activity Subject) leaves ~40% "
                   "of cases unclassifiable. Description fill is only 36-66% depending on subject.",
         "management_implication": "Production-grade classification will require email body text, not just "
                                   "CRM fields. The 1-day email sample we have confirms email bodies are rich "
                                   "enough (82% > 500 chars). Longer email extract is the critical next data request.",
         "source": "Consistent across all 5 subjects in this analysis."},
    ]
    sheets["5_CrossSubjFindings"] = pd.DataFrame(findings)
    print("  Tab 5: Cross-Subject Findings")

    # ═══════════════════════════════════════════════════════
    #  APPENDIX TABS 6-10: Full Detail (Chris validation)
    # ═══════════════════════════════════════════════════════
    appendix = [
        ("6_APP_Research", "Research", RESEARCH_FULL, None),
        ("7_APP_AcctMaint", "Account Maintenance", ACCT_MAINT_FULL, None),
        ("8_APP_NewAcctReq", "New Account Request", NEW_ACCT_FULL, None),
        ("9_APP_GeneralQ", "General Questions", GEN_Q_FULL, None),
        ("10_APP_CloseAcct", "Close Account", CLOSE_ACCT_FULL, "Clos"),
    ]
    for sn, subj, clust, filt in appendix:
        sheets[sn] = build_full_breakdown(client, subj, clust, filt)
        print(f"  {sn}: {subj} (appendix)")

    # ═══════════════════════════════════════════════════════
    #  APPENDIX TAB 11: Email Draft
    # ═══════════════════════════════════════════════════════
    email = [
        {"section": "DRAFT EMAIL", "content":
         "Subject: Case Subject Sub-Segmentation — Validation Before Leadership Presentation"},
        {"section": "DRAFT EMAIL", "content": "Hi Chris,"},
        {"section": "DRAFT EMAIL", "content":
         "Following our March 23 discussion, we sub-segmented three case subjects using keyword "
         "matching on the Description and Activity Subject fields. Keywords were selected from "
         "what's actually in the data — including WAB-specific terms like lockbox, CDARS, ICS, "
         "ConnectLive, PMC. Before presenting to Bob, I'd like your confirmation."},
        {"section": "DRAFT EMAIL", "content":
         "1. Research (4,407 cases → 4 sub-types): Payment/ACH Investigation is 27%, confirming your "
         "'typically payment research' characterization. Check/Image is 13%. Rate/CD/Maturity inquiries "
         "are 6% — these are CD/IntraFi questions that landed in Research instead of CD Maintenance."},
        {"section": "DRAFT EMAIL", "content":
         "2. Account Maintenance (3,242 → 3 sub-types): The headline finding — 580 cases (18%) are "
         "CD/IntraFi maturity processing, not account maintenance. 'Maturity notice' is the #1 term "
         "in Account Maintenance text. Lockbox file/configuration is another 15%."},
        {"section": "DRAFT EMAIL", "content":
         "3. Close Account (1,859 → 3 sub-types): 355 cases (12%) are CD maturity closures, not "
         "relationship departures. HOA account closures are 27%."},
        {"section": "DRAFT EMAIL", "content":
         "Cross-cutting finding: ~935 cases across these 3 subjects are actually CD/IntraFi maturity "
         "events that landed under the wrong subject. This suggests the subject pick list doesn't have "
         "a clean path for maturity processing."},
        {"section": "DRAFT EMAIL", "content":
         "Questions:\n"
         "  1. Do the 3-4 sub-types per subject match how your team thinks about the work?\n"
         "  2. The maturity leakage — is this known? Would it be useful to surface for leadership?\n"
         "  3. Any WAB-specific terms we're missing? (full keyword lists in appendix tabs)\n"
         "  4. Should we show all 3 breakdowns to Bob, or focus on Research only?"},
        {"section": "DRAFT EMAIL", "content":
         "The workbook is attached — Tabs 1-5 are the exec view, Tabs 6-10 have the full detail.\n\n"
         "Happy to walk through on a quick call.\n\nBest,\nRavi"},
    ]
    sheets["11_APP_EmailDraft"] = pd.DataFrame(email)
    print("  Tab 11: Email Draft")

    # Write
    print(f"\nWriting: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            write_sheet(writer, name[:31], sdf)
    print(f"Done in {(datetime.datetime.now() - start).total_seconds():.1f}s")


if __name__ == "__main__":
    main()
