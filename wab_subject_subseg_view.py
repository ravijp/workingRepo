"""
WAB Subject Sub-Segmentation View — Weekly Call & Leadership Artifact
======================================================================
Data-grounded clusters based on keyword diagnostic output (March 25, 2026).
Keywords selected from actual unigrams, bigrams, and Activity Subject
phrases observed in WAB case data.

Output: 9-sheet Excel workbook for Chris validation and Bob presentation.

Dependencies: pandas, openpyxl
"""

# ┌─────────────────────────────────────────────────────────┐
CASE_FILE  = r"C:\Users\YourName\Desktop\AAB All Cases.xlsx"
OUTPUT_DIR = r"C:\Users\YourName\Desktop\wab_output"
# └─────────────────────────────────────────────────────────┘

import os, re, datetime, warnings
from collections import OrderedDict, Counter

import pandas as pd
import numpy as np

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

def top_keywords(texts, top_n=20):
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
#  CLUSTER DEFINITIONS — grounded in actual WAB keyword data
#  Source: wab_keyword_diagnostic.xlsx bigrams + unigrams
# ═══════════════════════════════════════════════════════════

# Priority order matters: first match wins. Most specific patterns first.

RESEARCH_CLUSTERS = OrderedDict([
    ("Payment / ACH Research", {
        "keywords": ["payment research", "homeowner payment", "missing payment",
                     "returned payment", "payment failed", "payment return",
                     "ach", "wire alert", "wire", "misc debit",
                     "deposit", "return", "reversal", "refund",
                     "incoming", "outgoing", "remittance", "posted",
                     "posting", "transaction"],
        "rationale": "Chris: 'typically payment research.' Confirmed by data: 'payment research'(103), "
                     "'homeowner payment'(45), 'missing payment'(29), 'returned payment'(29). "
                     "ACH (262 unigram), wire alert(27) are distinct payment sub-types.",
    }),
    ("Check / Image Research", {
        "keywords": ["check research", "check copy", "check image", "missing check",
                     "check error", "checks", "check", "image request",
                     "cashier", "stop payment", "stale", "endorsement"],
        "rationale": "Data: 'check research'(76), 'check copy'(29), 'missing check'(29), "
                     "'check image'(25), 'stop payment'(52). Check-related research is ~15% of Research.",
    }),
    ("Rate Sheet / CD / Maturity Inquiry", {
        "keywords": ["rate sheet", "current rate", "cd rate", "weekly cd",
                     "interest rate sheet", "maturity notice", "cdars maturity",
                     "maturity", "cdars", "ics", "intrafi",
                     "certificate", "renewal", "rollover", "rate"],
        "rationale": "Data: 'rate sheet'(59), 'maturity notice'(57), 'cdars maturity'(41). "
                     "These are CD/IntraFi inquiries that landed in Research instead of "
                     "CD Maintenance. Taxonomy leakage — a finding in itself.",
    }),
    ("Lockbox / HOA Deposit Research", {
        "keywords": ["lockbox", "lockbox file", "lockbox ach",
                     "coupon", "remit", "bulk deposit",
                     "dues", "assessment"],
        "rationale": "Data: 'lockbox'(197 unigram), bigrams 'close lockbox'(47), 'lockbox ach'(29). "
                     "Lockbox is a core HOA banking product — research on lockbox processing "
                     "is a distinct workflow.",
    }),
    ("HOA / PMC Entity Research", {
        "keywords": ["homeowners association", "condominium association", "community association",
                     "owners association", "association inc", "inquiry pmc",
                     "research pmc", "research homeowner",
                     "village", "park", "ridge", "manor", "estates",
                     "hoa", "pmc", "cmc", "management company"],
        "rationale": "Data: 'homeowners association'(104), 'association inc'(88), 'condominium association'(29). "
                     "HOA/PMC entity-level research: looking up client info, verifying relationships. "
                     "Community names (village, park, ridge) are entity identifiers.",
    }),
    ("Notice / Correspondence", {
        "keywords": ["notice available", "notice", "correspondence",
                     "returned mail", "undeliverable", "letter"],
        "rationale": "Data: 'notice'(132 unigram), 'notice available'(29 bigram). "
                     "Cases triggered by notices or correspondence events.",
    }),
    ("Card Research (Credit/Debit)", {
        "keywords": ["credit card", "debit card", "card"],
        "rationale": "Data: 'credit card'(29), 'debit'(unigram). Card-related research "
                     "distinct from payment or check research.",
    }),
])

ACCT_MAINT_CLUSTERS = OrderedDict([
    ("CD / IntraFi Maturity Processing", {
        "keywords": ["maturity notice", "cdars maturity", "maturity",
                     "cdars", "ics", "intrafi", "inc month",
                     "certificate", "renewal", "rollover", "cd "],
        "rationale": "Data: 'maturity notice'(330!) is the #1 bigram. CD/IntraFi maturity "
                     "processing is the single largest Account Maintenance sub-type. "
                     "Major taxonomy leakage finding — these should arguably be CD Maintenance.",
    }),
    ("Lockbox File / Configuration", {
        "keywords": ["lockbox file", "lockbox", "close lockbox", "lockbox ach",
                     "file properties", "file", "coupon"],
        "rationale": "Data: 'lockbox file'(165), 'file properties'(76). Lockbox setup, "
                     "configuration changes, and file management.",
    }),
    ("Validation / Document Processing", {
        "keywords": ["validation file", "management validation", "validation",
                     "first choice", "choice property",
                     "action required", "required"],
        "rationale": "Data: 'validation file'(91), 'management validation'(29), "
                     "'first choice'(30), 'action required'(31). Document/file validation workflow.",
    }),
    ("Stop Payment Processing", {
        "keywords": ["stop payment"],
        "rationale": "Data: 'stop payment'(98 bigram). Distinct workflow — banker places "
                     "a stop on a check. Clear, single-action maintenance.",
    }),
    ("Association / HOA Maintenance", {
        "keywords": ["association inc", "homeowners association", "condominium association",
                     "community association", "owners association",
                     "association month", "association account",
                     "hoa", "pmc", "cmc"],
        "rationale": "Data: 'association inc'(164), 'homeowners association'(57). "
                     "HOA entity-level account maintenance.",
    }),
    ("Card Maintenance (Debit/Credit)", {
        "keywords": ["debit card", "credit card", "card"],
        "rationale": "Data: 'debit card'(41). Card-related maintenance — replacements, "
                     "limits, activations.",
    }),
    ("Fee Waiver / Adjustment", {
        "keywords": ["waive return", "waive", "fee", "refund", "adjustment",
                     "service charge", "overdraft"],
        "rationale": "Data: 'waive return'(33). Fee-related maintenance and adjustments.",
    }),
    ("Current Rates / Rate Sheet", {
        "keywords": ["current rates", "rate sheet", "rates", "interest rate"],
        "rationale": "Data: bigram 'current rates'(29). Rate inquiries routed to Account Maintenance.",
    }),
])

NEW_ACCT_CLUSTERS = OrderedDict([
    ("HOA / Association New Account", {
        "keywords": ["homeowners association", "condominium association",
                     "community association", "owners association",
                     "association inc", "association alliance",
                     "hoa", "pmc", "cmc", "village", "park",
                     "ridge", "manor", "estates", "condominium"],
        "rationale": "Data: 'association inc'(550!), 'homeowners association'(460), "
                     "'condominium association'(159). The overwhelming majority of new "
                     "account requests are HOA onboarding.",
    }),
    ("Reserve / Petty Cash / Sub-Account", {
        "keywords": ["reserve account", "petty cash", "reserve",
                     "child case", "parent case", "sub account",
                     "operating", "money market"],
        "rationale": "Data: 'reserve account'(98), 'petty cash'(42), 'child case'(91), "
                     "'parent case'(68). HOAs often have multiple sub-accounts: operating, "
                     "reserve, petty cash.",
    }),
    ("Document / Validation Required", {
        "keywords": ["validation file", "response required", "maturity notice",
                     "document", "missing", "pending", "submit",
                     "tin", "ein", "w-9", "w9", "signed"],
        "rationale": "Data: 'validation file'(113), 'response required'(64). Document "
                     "collection and validation during onboarding.",
    }),
    ("Bank Account / Product Setup", {
        "keywords": ["bank account", "bank accounts", "new bank",
                     "open", "acct", "account request",
                     "inc alliance", "alliance aab"],
        "rationale": "Data: 'bank accounts'(72), 'account request'(164 bigram). "
                     "Generic new account setup not specific to HOA.",
    }),
    ("Lockbox / ACH / Wire Setup", {
        "keywords": ["lockbox", "ach", "wire", "sweep", "zba",
                     "positive pay", "online banking", "connectlive"],
        "rationale": "Setting up payment infrastructure for new accounts.",
    }),
])

GEN_Q_CLUSTERS = OrderedDict([
    ("Rate Sheet / Interest Rate Inquiry", {
        "keywords": ["rate sheet", "rate sheets", "interest rate",
                     "interest rates", "current rate", "rates",
                     "money market", "cd rate", "yield"],
        "rationale": "Data: 'rate sheet'(98!), 'interest rates'(82), 'current rate'(72). "
                     "Rate inquiries are the dominant General Questions sub-type.",
    }),
    ("Online Banking / Digital Access", {
        "keywords": ["online banking", "portal", "login", "password",
                     "access", "connectlive", "mobile", "app",
                     "token", "reset", "locked"],
        "rationale": "Data: 'online banking'(153 unigram), 'debit card'(123). "
                     "Digital channel access questions.",
    }),
    ("Card Inquiry (Debit/Credit)", {
        "keywords": ["debit card", "credit card", "card", "visa",
                     "mastercard", "atm", "pin"],
        "rationale": "Data: 'debit card'(123 unigram), 'credit card'(113). "
                     "Card-related questions — limits, activation, replacements.",
    }),
    ("HOA / Association Question", {
        "keywords": ["hoa", "homeowners", "association", "community",
                     "condominium", "owners association",
                     "hoa payment", "pmc", "cmc", "lockbox"],
        "rationale": "Data: 'hoa'(524 unigram), 'association'(487), "
                     "'community association'(81). HOA-specific questions.",
    }),
    ("Product / Account Inquiry", {
        "keywords": ["ics account", "cdars", "reserve account",
                     "petty cash", "money market", "savings",
                     "checking", "deposit", "balance", "statement",
                     "account ending"],
        "rationale": "Data: 'ics'(155 unigram), 'cdars'(190), 'reserve account'(53). "
                     "Product-specific balance/feature inquiries.",
    }),
    ("Process Support / Internal", {
        "keywords": ["process support", "quick question", "question",
                     "help", "assist", "information"],
        "rationale": "Data: 'process support'(72), 'quick question'(53 bigram). "
                     "Internal process questions or generic help requests.",
    }),
])

CLOSE_ACCT_CLUSTERS = OrderedDict([
    ("CD / IntraFi Maturity Closure", {
        "keywords": ["maturity notice", "cdars maturity", "matured",
                     "maturity", "cdars", "ics", "intrafi",
                     "cd maturity", "certificate"],
        "rationale": "Data: 'maturity notice'(218!) is the #1 Close Account bigram. "
                     "CDs reaching maturity and closing rather than renewing. "
                     "Major finding: 218 out of ~1,859 Close Account cases are actually "
                     "CD maturity events, not full relationship closures.",
    }),
    ("HOA / Association Account Closure", {
        "keywords": ["homeowners association", "association inc",
                     "condominium association", "community association",
                     "owners association", "property owners",
                     "hoa", "pmc", "village", "park", "ridge"],
        "rationale": "Data: 'homeowners association'(202), 'association inc'(172). "
                     "Full HOA account closures — PMC departing the bank.",
    }),
    ("Standard Account Closure Request", {
        "keywords": ["account closure", "close account", "close accounts",
                     "closure request", "request close", "account closures",
                     "close acct", "closing accounts",
                     "close bank", "acct closure", "offboarding"],
        "rationale": "Data: 'account closure'(217), 'close accounts'(134), "
                     "'closure request'(94). Standard closure workflow — the checklist "
                     "Chris described (IBS, BST, ACH tracker).",
    }),
    ("Reserve / Sub-Account Closure", {
        "keywords": ["reserve", "petty cash", "funds",
                     "bank accounts", "bank account",
                     "cash check", "request redeem"],
        "rationale": "Data: 'reserve'(75 unigram), 'bank accounts'(66). "
                     "Closing sub-accounts within a relationship (reserve, petty cash) "
                     "rather than the full relationship.",
    }),
    ("Lockbox / Product Closure", {
        "keywords": ["lockbox", "ach", "wire", "sweep",
                     "online banking", "positive pay",
                     "connectlive", "bst"],
        "rationale": "Closing ancillary products. Chris: 'Lockbox has to be close — "
                     "need to keep it open for 10 days for all txns to go through.'",
    }),
])


# ═══════════════════════════════════════════════════════════
#  CLASSIFICATION ENGINE
# ═══════════════════════════════════════════════════════════

def classify(desc, act_subj, cluster_defs):
    combined = f"{desc} {act_subj}".lower()
    combined = re.sub(r"caution.*?originated outside.*?organization[.\s]*", " ",
                      combined, flags=re.IGNORECASE | re.DOTALL)
    combined = re.sub(r"https?://\S+", " ", combined)
    combined = re.sub(r"\*{2,}", " ", combined)  # strip *** markers

    for cluster_name, cfg in cluster_defs.items():
        kws = cfg.get("keywords", [])
        if not kws: continue
        if any(kw in combined for kw in kws):
            return cluster_name
    if len(combined.strip()) < 10:
        return "(No Text)"
    return "(Other/Uncategorized)"


def build_breakdown(client_df, subject_name, cluster_defs, subject_filter=None):
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
    all_clusters = list(cluster_defs.keys()) + ["(No Text)", "(Other/Uncategorized)"]
    for cluster in all_clusters:
        if cluster not in grp.groups: continue
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
            "key_terms": ", ".join(cluster_defs.get(cluster, {}).get("keywords", [])[:5]) + ("..." if len(cluster_defs.get(cluster, {}).get("keywords", [])) > 5 else "") if cluster in cluster_defs else "(residual)",
        })

    # Top keywords from Other
    other = subset[subset["_cluster"] == "(Other/Uncategorized)"]
    if len(other) > 0:
        kw = top_keywords(other["_desc"].tolist() + other["_act_subj"].tolist(), 15)
        rows.append({
            "cluster": ">>> TOP KEYWORDS IN OTHER <<<",
            "key_terms": ", ".join(f"{w}({c})" for w, c in kw),
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

    # ── Sheet 1: Master Subject View ──
    subj_vc = client["_subject"].value_counts()
    top15 = subj_vc.head(15).index.tolist()

    chris = {
        "Research": "Typically payment research; a lot different things go into it.",
        "CD Maintenance": "Banker sits on case till maturity date. By design.",
        "Signature Card": "Keith owns. Large backlog. Board member changes.",
        "NSF and Non-Post": "2-3 hrs/day per banker. Decision items one at a time in IBS.",
        "Close Account": "Manual checklist: IBS, BST, ACH tracker. Consolidating into CRM.",
        "IntraFi Maintenance": "Similar to CD maintenance. Check with James.",
        "QC Finding": "Eduardo Jacobo. Prior day change reports.",
    }
    genai = {
        "Research": "HIGH — sub-segment into 7 types, route by email intent",
        "New Account Request": "HIGH — draft reply + missing-info detection",
        "Account Maintenance": "HIGH — fat tail, CD maturity leakage finding",
        "NSF and Non-Post": "MEDIUM — IBS automation, not AI",
        "CD Maintenance": "LOW — by design wait time",
        "Close Account": "LOW — process redesign first (CRM consolidation)",
        "General Questions": "MEDIUM — rate sheet queries could be self-service",
        "Signature Card": "MEDIUM — IDP for document processing",
        "Fraud Alert": "LOW — already fast, rules-based",
        "IntraFi Maintenance": "LOW — by design wait like CD",
        "Transfer": "LOW — already fast, straight-through",
    }
    subseg = {
        "Research": "YES — Sheet 2 (7 sub-types)",
        "Account Maintenance": "YES — Sheet 3 (8 sub-types)",
        "New Account Request": "YES — Sheet 4 (5 sub-types)",
        "General Questions": "YES — Sheet 5 (6 sub-types)",
        "Close Account": "YES — Sheet 6 (5 sub-types)",
    }

    master_rows = []
    for subj in top15:
        s = client[client["_subject"] == subj]
        n = len(s)
        hrs = s["_hours"].dropna()
        unres = (~s["_is_resolved"]).sum()
        master_rows.append({
            "subject": subj, "cases": n,
            "pct": f"{round(100*n/n_total,1)}%",
            "median_hrs": round(hrs.median(),1) if len(hrs) > 0 else "",
            "p75_hrs": round(hrs.quantile(0.75),1) if len(hrs) > 0 else "",
            "p90_hrs": round(hrs.quantile(0.9),1) if len(hrs) > 0 else "",
            "unresolved": int(unres),
            "pct_unresolved": f"{round(100*unres/n,1)}%",
            "chris_context": chris.get(subj, ""),
            "genai_opportunity": genai.get(subj, ""),
            "sub_segmentation": subseg.get(subj, ""),
        })
    total = sum(r["cases"] for r in master_rows)
    master_rows.append({"subject": "=== TOP 15 ===", "cases": total, "pct": f"{round(100*total/n_total,1)}%"})
    sheets["1_MasterSubjectView"] = pd.DataFrame(master_rows)
    print("  Sheet 1: Master")

    # ── Sheets 2-6: Breakdowns ──
    for sn, subj, clust, filt in [
        ("2_Research", "Research", RESEARCH_CLUSTERS, None),
        ("3_AcctMaint", "Account Maintenance", ACCT_MAINT_CLUSTERS, None),
        ("4_NewAcctReq", "New Account Request", NEW_ACCT_CLUSTERS, None),
        ("5_GeneralQ", "General Questions", GEN_Q_CLUSTERS, None),
        ("6_CloseAcct", "Close Account", CLOSE_ACCT_CLUSTERS, "Clos"),
    ]:
        sheets[sn] = build_breakdown(client, subj, clust, filt)
        nc = len(client[client["_subject"]==subj]) if not filt else len(client[client["_subject"].str.contains(filt,case=False,na=False)])
        print(f"  {sn}: {subj} ({nc:,})")

    # ── Sheet 7: Keyword Rationale ──
    kw_rows = []
    for subj_name, cluster_defs in [
        ("Research", RESEARCH_CLUSTERS), ("Account Maintenance", ACCT_MAINT_CLUSTERS),
        ("New Account Request", NEW_ACCT_CLUSTERS), ("General Questions", GEN_Q_CLUSTERS),
        ("Close Account", CLOSE_ACCT_CLUSTERS),
    ]:
        for cn, cfg in cluster_defs.items():
            kw_rows.append({
                "subject": subj_name, "cluster": cn,
                "keywords": ", ".join(cfg.get("keywords",[])) or "(residual)",
                "rationale": cfg.get("rationale",""),
                "validation": "PENDING",
            })
    sheets["7_KeywordRationale"] = pd.DataFrame(kw_rows)
    print("  Sheet 7: Rationale")

    # ── Sheet 8: Cross-Subject Findings ──
    findings = [
        {"finding": "CD/IntraFi Maturity Leakage",
         "detail": "Maturity notice appears as #1 bigram in Account Maintenance (330), Close Account (218), and Research (57). "
                   "These are CD/IntraFi maturity events landing in the wrong subject. ~600 cases are mis-categorized.",
         "implication": "Subject taxonomy needs a maturity workflow path. Or — these subjects should have a "
                        "sub-tag for maturity processing. AI could auto-tag at case creation."},
        {"finding": "HOA Entity Names Dominate All Subjects",
         "detail": "'homeowners association,' 'association inc,' 'condominium association' are top bigrams "
                   "in Research, Account Maintenance, New Account Request, AND Close Account. "
                   "These are HOA entity identifiers, not workflow descriptors.",
         "implication": "The Subject field captures WHAT the case is about, but the text fields are full of "
                        "WHO it's about (the HOA name). AI classification needs to look past entity names to find intent."},
        {"finding": "Rate Sheet Queries Could Be Self-Service",
         "detail": "Rate sheet/interest rate inquiries are the #1 General Questions sub-type (98 bigrams). "
                   "These are lookups, not complex requests.",
         "implication": "A rate sheet bot or self-service portal page could deflect these entirely. "
                        "No AI needed — just information availability."},
        {"finding": "Lockbox Is a Cross-Subject Workflow",
         "detail": "Lockbox appears in Research (197 unigram), Account Maintenance (165 bigram 'lockbox file'), "
                   "New Account (setup), and Close Account (closure). It's a product that spans the case lifecycle.",
         "implication": "Lockbox cases may benefit from a dedicated subject or at minimum a sub-tag. "
                        "AI routing could identify lockbox-related cases regardless of subject."},
        {"finding": "Validation File Workflow in Account Maintenance",
         "detail": "'validation file'(91), 'file properties'(76), 'management validation'(29) form a "
                   "distinct document processing workflow within Account Maintenance.",
         "implication": "This is potentially addressable by the IDP pilot Chris mentioned. "
                        "If IDP works for <20 page docs, validation files are a candidate."},
    ]
    sheets["8_CrossSubjectFindings"] = pd.DataFrame(findings)
    print("  Sheet 8: Cross-Subject Findings")

    # ── Sheet 9: Email Draft ──
    email = [
        {"section": "DRAFT EMAIL", "content": "Subject: Case Subject Sub-Segmentation — Validation Before Leadership Presentation"},
        {"section": "DRAFT EMAIL", "content": "Hi Chris,"},
        {"section": "DRAFT EMAIL", "content":
         "Following our March 23 discussion, we've sub-segmented five case subjects using keyword "
         "matching on Description and Activity Subject fields. We used terms from the actual data "
         "(not assumed vocabulary) — including WAB-specific terms like lockbox, CDARS, ICS, "
         "ConnectLive, PMC, aq2 that we found in the text."},
        {"section": "DRAFT EMAIL", "content":
         "Subjects broken down:\n"
         "  1. Research (4,407 → 7 sub-types: Payment/ACH, Check/Image, Rate/CD/Maturity, Lockbox/HOA, Entity, Notice, Card)\n"
         "  2. Account Maintenance (3,242 → 8 sub-types: CD/Maturity is #1 at ~330 cases — taxonomy leakage finding)\n"
         "  3. New Account Request (3,665 → 5 sub-types: HOA onboarding dominates)\n"
         "  4. General Questions (1,783 → 6 sub-types: Rate sheet queries are #1 — self-service candidate)\n"
         "  5. Close Account (1,859 → 5 sub-types: CD maturity closures are ~218 of these)"},
        {"section": "DRAFT EMAIL", "content":
         "Key finding: CD/IntraFi maturity events appear in Account Maintenance, Close Account, AND "
         "Research — about 600 cases total that are really maturity processing but landed under other "
         "subjects. This is a taxonomy finding we should discuss."},
        {"section": "DRAFT EMAIL", "content":
         "Questions:\n"
         "  1. Do these sub-types match how your team thinks about the work?\n"
         "  2. The keyword lists are in Sheet 7 — any terms we're missing?\n"
         "  3. The maturity leakage finding — is this known? Would it be useful to surface for Bob?\n"
         "  4. Should we show all 5 breakdowns to leadership, or focus on Research + one other?\n"
         "  5. Any subjects we should NOT break down (too sensitive)?"},
        {"section": "DRAFT EMAIL", "content": "Happy to walk through on a quick call.\n\nBest,\nRavi"},
    ]
    sheets["9_EmailDraft"] = pd.DataFrame(email)
    print("  Sheet 9: Email Draft")

    # Write
    print(f"\nWriting: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            write_sheet(writer, name[:31], sdf)
    print(f"Done in {(datetime.datetime.now() - start).total_seconds():.1f}s")


if __name__ == "__main__":
    main()
