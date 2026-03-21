"""
WAB Email Deep Insights — Phase 1 Evidence Expansion (v2)
==========================================================
Standalone VDI script. Reads Cases + Emails files.
Focuses on 3 questions that matter for the GenAI business case:

  1. What are bankers actually reading? (intent taxonomy)
  2. How much of the work is repetitive? (template coverage)
  3. Where is the text evidence for proposed use cases?

Produces one Excel workbook + one markdown summary.

Required:  pandas, openpyxl, bs4, scikit-learn
Optional:  rapidfuzz (improves template detection)
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

import os, re, html, datetime, warnings
from collections import OrderedDict

import numpy as np
import pandas as pd
from bs4 import BeautifulSoup
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.decomposition import NMF
from sklearn.metrics.pairwise import cosine_similarity

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

try:
    from rapidfuzz import fuzz  # noqa: F401
    HAS_FUZZ = True
except ImportError:
    HAS_FUZZ = False

OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "wab_email_deep_insights.xlsx")
OUTPUT_MD   = os.path.join(OUTPUT_DIR, "wab_email_deep_insights_summary.md")
LOG, WARN   = [], []
TRUNC       = 160

INTERNAL_COMPANIES = {"AAB ADMIN", "WAB ADMIN", "AAB ADMIN -", "WAB ADMIN -"}

# ── Domain knowledge from Chris call + data analysis ──
# Documents commonly required in HOA banking workflows
DOC_TERMS = [
    "signature card", "voided check", "void check", "management agreement",
    "ein", "w9", "w-9", "tin", "board resolution", "articles of incorporation",
    "articles", "driver license", "identification", "operating agreement",
    "reserve study", "budget", "insurance certificate", "fidelity bond",
    "certificate of deposit", "cd maturity", "ics agreement", "cdars agreement",
]

# Cues that an email is requesting missing information
MISSING_CUES = [
    "missing", "not received", "still need", "need the", "need your",
    "please provide", "please send", "awaiting", "required", "incomplete",
    "once we receive", "can you send", "have not received", "pending receipt",
    "not yet received", "still waiting for", "in order to proceed",
    "unable to process", "cannot proceed without",
]

# Follow-up / chase language
FOLLOWUP_CUES = [
    "following up", "follow up", "just following up", "checking in",
    "circling back", "still waiting", "any update", "please advise",
    "status update", "when can we expect", "have you had a chance",
    "wanted to check", "reaching out again", "second request",
]

# Urgency language
URGENCY_CUES = [
    "urgent", "asap", "immediately", "today", "end of day", "eod",
    "rush", "critical", "time sensitive", "deadline", "priority",
    "expedite", "right away", "as soon as possible",
]

# Stopwords for token-based analysis
STOPWORDS = {
    "a", "an", "and", "are", "as", "at", "be", "been", "but", "by", "can",
    "could", "do", "for", "from", "had", "has", "have", "if", "in", "into",
    "is", "it", "its", "just", "me", "my", "not", "of", "on", "or", "our",
    "please", "re", "so", "that", "the", "their", "them", "there", "this",
    "to", "up", "us", "was", "we", "were", "will", "with", "you", "your",
    "fw", "fwd", "re", "subject", "thanks", "thank", "hello", "hi",
    "regards", "regarding", "email", "num", "date",
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
    s = str(v).replace("\r", " ").replace("\n", " ").strip()
    return s[:n] + "..." if len(s) > n else s

def pct(num, denom):
    return f"{100 * num / denom:.1f}%" if denom else "N/A"

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

def safe_num(series):
    if pd.api.types.is_numeric_dtype(series):
        return series
    return pd.to_numeric(series, errors="coerce")

def write_sheet(writer, name, df):
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        df = pd.DataFrame({"note": ["No data"]})
    sheet_name = name[:31]
    df.to_excel(writer, sheet_name=sheet_name, index=False, freeze_panes=(1, 0))
    ws = writer.sheets[sheet_name]
    for col_cells in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 55)

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
    log(f"  Rows: {len(df):,} | Columns: {len(df.columns)}")
    return df


# ═══════════════════════════════════════════════════════════
#  TEXT PREPROCESSING
# ═══════════════════════════════════════════════════════════

def strip_html(text):
    """Convert HTML email to plain text using BeautifulSoup."""
    if pd.isna(text): return ""
    s = str(text)
    # Remove style/script blocks before parsing
    s = re.sub(r"<(style|script)[^>]*>.*?</\1>", " ", s, flags=re.I | re.S)
    soup = BeautifulSoup(s, "html.parser")
    text = soup.get_text("\n")
    text = html.unescape(text)
    text = re.sub(r"\r", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def remove_noise(text):
    """Remove security banners, forwarded headers, and signature blocks."""
    if not text: return ""
    s = text

    # Security banners (external email warnings)
    s = re.sub(r"(?i)attention:\s*this email originated from outside.*?(?=\n\n|\Z)", " ", s, flags=re.S)
    s = re.sub(r"(?i)external email warning.*?(?=\n\n|\Z)", " ", s, flags=re.S)
    s = re.sub(r"(?i)caution:\s*external.*?(?=\n\n|\Z)", " ", s, flags=re.S)
    # Do NOT click links warnings
    s = re.sub(r"(?i)do not click links.*?(?=\n\n|\Z)", " ", s, flags=re.S)

    # Forwarded email headers (From: ... Sent: ... To: ... Subject: ...)
    s = re.sub(r"(?i)from:\s.*?sent:\s.*?to:\s.*?subject:\s[^\n]*", " ", s, flags=re.S)
    s = re.sub(r"(?i)on .{10,60} wrote:\s*", " ", s)

    # Horizontal rules / separator lines
    s = re.sub(r"_{5,}", " ", s)
    s = re.sub(r"-{5,}", " ", s)

    # Signature blocks — only strip from last 30% of text to avoid cutting
    # "Thank you for the update, here is the status..." mid-sentence
    cutpoint = max(len(s) * 7 // 10, 200)
    head = s[:cutpoint]
    tail = s[cutpoint:]
    tail = re.sub(r"(?i)\b(?:best regards|regards|sincerely|thanks|thank you)\s*,?\s*\n.*", "", tail, flags=re.S)
    s = head + tail

    s = re.sub(r"\s+", " ", s).strip()
    return s


def extract_new_content(raw_text):
    """Estimate how much of the email is new vs quoted/forwarded content.
    Returns (new_text, quoted_text, new_ratio)."""
    if not raw_text:
        return "", "", 0.0

    lines = raw_text.split("\n")
    new_lines = []
    quoted_lines = []
    in_quoted = False

    for line in lines:
        stripped = line.strip()
        # Detect start of quoted section
        if re.match(r"^>", stripped) or re.match(r"(?i)^(from|sent|to|subject|on .* wrote):", stripped):
            in_quoted = True
        # Detect separator lines
        if re.match(r"^[-_=]{4,}$", stripped):
            in_quoted = True

        if in_quoted:
            quoted_lines.append(line)
        else:
            new_lines.append(line)

    new_text = "\n".join(new_lines).strip()
    quoted_text = "\n".join(quoted_lines).strip()
    total = len(new_text) + len(quoted_text)
    new_ratio = len(new_text) / total if total > 0 else 0.0
    return new_text, quoted_text, new_ratio


def canonical(text):
    """Normalize text for NLP: lowercase, replace emails/dates/numbers, strip punctuation."""
    if not text: return ""
    s = text.lower()
    s = re.sub(r"\b[\w\.-]+@[\w\.-]+\.\w+\b", " _EMAIL_ ", s)
    s = re.sub(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", " _DATE_ ", s)
    s = re.sub(r"\$[\d,]+\.?\d*", " _AMOUNT_ ", s)
    s = re.sub(r"\b[A-Z]{2,3}-\d{5,}-\w+\b", " _CASEID_ ", s, flags=re.I)
    s = re.sub(r"\b\d{4,}\b", " _NUM_ ", s)
    s = re.sub(r"[^a-z0-9_\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def tokens(text):
    """Extract meaningful tokens from canonical text."""
    if not text: return []
    return [t for t in canonical(text).split() if len(t) > 2 and t not in STOPWORDS and not t.startswith("_")]


# ═══════════════════════════════════════════════════════════
#  DATA PREPARATION
# ═══════════════════════════════════════════════════════════

def prepare_cases(cases):
    df = cases.copy()
    col_map = {
        "case_number": find_col(df, "Case Number"),
        "company":     find_col(df, "Company Name (Company) (Company)", "Company Name"),
        "subject":     find_col(df, "Subject"),
        "activity_subj": find_col(df, "Activity Subject"),
        "status_reason": find_col(df, "Status Reason"),
        "status":      find_col(df, "Status"),
        "created_on":  find_col(df, "Created On"),
        "sla_start":   find_col(df, "SLA Start"),
        "resolved_hrs": find_col(df, "Resolved In Hours"),
        "owner":       find_col(df, "Manager (Owning User) (User)", "Owner"),
        "pod":         find_col(df, "POD Name (Owning User) (User)", "POD Name"),
    }

    if col_map["created_on"]:
        df["_created_dt"] = safe_dt(df[col_map["created_on"]])
    if col_map["sla_start"]:
        df["_sla_start_dt"] = safe_dt(df[col_map["sla_start"]])
    if col_map["resolved_hrs"]:
        df["_hours"] = safe_num(df[col_map["resolved_hrs"]])

    if col_map["company"]:
        df["_company"] = df[col_map["company"]].fillna("(blank)").astype(str).str.strip()
        df["_is_internal"] = df["_company"].str.upper().apply(
            lambda x: any(x.startswith(k) for k in INTERNAL_COMPANIES) or x == "(BLANK)")
    else:
        df["_company"] = "(blank)"
        df["_is_internal"] = False

    df["_case_number"] = df[col_map["case_number"]].astype(str) if col_map["case_number"] else ""
    df["_subject"] = df[col_map["subject"]].fillna("(blank)").astype(str).str.strip() if col_map["subject"] else "(blank)"
    df["_activity_subj"] = df[col_map["activity_subj"]].fillna("").astype(str).str.strip() if col_map["activity_subj"] else ""
    df["_owner"] = df[col_map["owner"]].fillna("(blank)").astype(str).str.strip() if col_map["owner"] else "(blank)"
    df["_pod"] = df[col_map["pod"]].fillna("(blank)").astype(str).str.strip() if col_map["pod"] else "(blank)"

    sr = col_map["status_reason"] or col_map["status"]
    if sr:
        df["_is_resolved"] = df[sr].fillna("").astype(str).str.lower().apply(
            lambda x: any(k in x for k in ("resolved", "closed", "cancelled", "canceled")))
    else:
        df["_is_resolved"] = True

    # Triage delay (from D16)
    if "_sla_start_dt" in df.columns and "_created_dt" in df.columns:
        df["_triage_minutes"] = (df["_created_dt"] - df["_sla_start_dt"]).dt.total_seconds() / 60
        df.loc[df["_triage_minutes"] < 0.5, "_triage_minutes"] = np.nan

    df._col_map = col_map
    return df


def prepare_emails(emails, cases):
    df = emails.copy()
    col_map = {
        "subject":     find_col(df, "Subject"),
        "description": find_col(df, "Description"),
        "from":        find_col(df, "From"),
        "to":          find_col(df, "To"),
        "status_reason": find_col(df, "Status Reason"),
        "created_on":  find_col(df, "Created On"),
        "owner":       find_col(df, "Owner"),
        "priority":    find_col(df, "Priority"),
        "case_number": find_col(df, "Case Number (Regarding) (Case)", "Case Number (Regarding)"),
        "case_subject": find_col(df, "Subject (Regarding) (Case)", "Subject Path (Regarding) (Case)"),
    }

    df["_email_id"] = np.arange(1, len(df) + 1)

    if col_map["created_on"]:
        df["_created_dt"] = safe_dt(df[col_map["created_on"]])
        df["_hour"] = df["_created_dt"].dt.hour
    else:
        df["_created_dt"] = pd.NaT
        df["_hour"] = np.nan

    df["_subject_raw"] = df[col_map["subject"]].fillna("").astype(str) if col_map["subject"] else ""
    df["_body_html"] = df[col_map["description"]].fillna("").astype(str) if col_map["description"] else ""
    df["_from"] = df[col_map["from"]].fillna("").astype(str) if col_map["from"] else ""
    df["_to"] = df[col_map["to"]].fillna("").astype(str) if col_map["to"] else ""
    df["_status_reason"] = df[col_map["status_reason"]].fillna("").astype(str) if col_map["status_reason"] else ""
    df["_case_ref"] = df[col_map["case_number"]].fillna("").astype(str) if col_map["case_number"] else ""
    df["_case_subject_raw"] = df[col_map["case_subject"]].fillna("(blank)").astype(str) if col_map["case_subject"] else "(blank)"
    df["_owner_email"] = df[col_map["owner"]].fillna("(blank)").astype(str) if col_map["owner"] else "(blank)"
    df["_has_case"] = df["_case_ref"].str.strip().ne("")

    # Direction
    dir_s = df["_status_reason"].str.lower()
    df["_direction"] = np.where(
        dir_s.str.contains("sent|completed|outgoing|outbound"), "Outbound",
        np.where(dir_s.str.contains("received|incoming|inbound"), "Inbound", "Other"))

    # Text preprocessing
    log("  Stripping HTML from email bodies...")
    df["_body_raw"] = df["_body_html"].apply(strip_html)
    df["_body_clean"] = df["_body_raw"].apply(remove_noise)
    df["_subject_clean"] = df["_subject_raw"].apply(remove_noise)
    df["_body_len"] = df["_body_clean"].str.len()

    # New vs quoted content analysis
    log("  Analyzing new vs quoted content...")
    content_results = df["_body_raw"].apply(extract_new_content)
    df["_new_text"] = content_results.apply(lambda x: x[0])
    df["_quoted_text"] = content_results.apply(lambda x: x[1])
    df["_new_ratio"] = content_results.apply(lambda x: x[2])
    df["_new_len"] = df["_new_text"].str.len()
    df["_quoted_len"] = df["_quoted_text"].str.len()

    # Canonical text for NLP (subject + first 1200 chars of clean body)
    df["_text_for_nlp"] = (df["_subject_clean"] + " " + df["_body_clean"].str[:1200]).str.strip()
    df["_text_canonical"] = df["_text_for_nlp"].apply(canonical)

    # Signal detection
    log("  Detecting missing-info, follow-up, and urgency signals...")
    df["_has_missing_cue"] = df["_text_for_nlp"].str.lower().apply(
        lambda t: any(cue in t for cue in MISSING_CUES))
    df["_has_followup_cue"] = df["_text_for_nlp"].str.lower().apply(
        lambda t: any(cue in t for cue in FOLLOWUP_CUES))
    df["_has_urgency_cue"] = df["_text_for_nlp"].str.lower().apply(
        lambda t: any(cue in t for cue in URGENCY_CUES))

    # Document terms mentioned
    df["_doc_terms_found"] = df["_text_for_nlp"].str.lower().apply(
        lambda t: ", ".join(sorted(set(term for term in DOC_TERMS if term in t)))
    )

    # Join case data
    case_cols = ["_case_number", "_subject", "_activity_subj", "_owner", "_pod",
                 "_company", "_hours", "_is_resolved", "_is_internal"]
    case_cols = [c for c in case_cols if c in cases.columns]
    case_lk = cases[case_cols].drop_duplicates(subset=["_case_number"])
    df = df.merge(case_lk, left_on="_case_ref", right_on="_case_number", how="left", suffixes=("", "_case"))

    df["_case_subject"] = df.get("_subject", pd.Series("(blank)")).fillna(df["_case_subject_raw"]).fillna("(blank)")
    df["_company"] = df.get("_company", pd.Series("(blank)")).fillna("(blank)").astype(str)
    df["_is_internal"] = df.get("_is_internal", pd.Series(False)).fillna(False).astype(bool)
    df["_is_resolved"] = df.get("_is_resolved", pd.Series(True)).fillna(True).astype(bool)
    df["_is_client"] = df["_has_case"] & (~df["_is_internal"])
    df["_is_inbound"] = df["_direction"] == "Inbound"
    df["_is_outbound"] = df["_direction"] == "Outbound"

    # Triage delay from case
    if "_triage_minutes" in cases.columns:
        triage_lk = cases[["_case_number", "_triage_minutes"]].drop_duplicates(subset=["_case_number"])
        df = df.merge(triage_lk, left_on="_case_ref", right_on="_case_number",
                      how="left", suffixes=("", "_triage"))

    df._col_map = col_map
    return df


# ═══════════════════════════════════════════════════════════
#  SHEET BUILDERS
# ═══════════════════════════════════════════════════════════

def sheet_01_scope(edf):
    """I01: What are we working with?"""
    n = len(edf)
    inbound = edf["_is_inbound"].sum()
    outbound = edf["_is_outbound"].sum()
    client = edf["_is_client"].sum()

    rows = [
        {"section": "POPULATION", "metric": "total_emails", "value": f"{n:,}"},
        {"section": "POPULATION", "metric": "linked_to_case", "value": f"{edf['_has_case'].sum():,} ({pct(edf['_has_case'].sum(), n)})"},
        {"section": "POPULATION", "metric": "client_case_emails", "value": f"{client:,} ({pct(client, n)})"},
        {"section": "POPULATION", "metric": "inbound", "value": f"{inbound:,} ({pct(inbound, n)})"},
        {"section": "POPULATION", "metric": "outbound", "value": f"{outbound:,} ({pct(outbound, n)})"},
        {"section": "TEXT QUALITY", "metric": "median_body_chars_clean", "value": f"{int(edf['_body_len'].median()):,}"},
        {"section": "TEXT QUALITY", "metric": "median_new_content_chars", "value": f"{int(edf['_new_len'].median()):,}"},
        {"section": "TEXT QUALITY", "metric": "median_new_content_ratio", "value": f"{edf['_new_ratio'].median():.0%}"},
        {"section": "TEXT QUALITY", "metric": "emails_with_over_500_new_chars", "value": f"{(edf['_new_len'] > 500).sum():,} ({pct((edf['_new_len'] > 500).sum(), n)})"},
        {"section": "SIGNAL DETECTION", "metric": "emails_with_missing_info_cue", "value": f"{edf['_has_missing_cue'].sum():,} ({pct(edf['_has_missing_cue'].sum(), n)})"},
        {"section": "SIGNAL DETECTION", "metric": "emails_with_followup_cue", "value": f"{edf['_has_followup_cue'].sum():,} ({pct(edf['_has_followup_cue'].sum(), n)})"},
        {"section": "SIGNAL DETECTION", "metric": "emails_with_urgency_cue", "value": f"{edf['_has_urgency_cue'].sum():,} ({pct(edf['_has_urgency_cue'].sum(), n)})"},
    ]
    return pd.DataFrame(rows)


def sheet_02_content_structure(edf):
    """I02: How much of each email is new content vs quoted/forwarded?
    This directly measures context-reconstruction burden."""
    parts = []

    # Overall distribution
    dist_rows = []
    for label, col in [("new_content_chars", "_new_len"), ("quoted_content_chars", "_quoted_len"),
                        ("new_content_ratio", "_new_ratio")]:
        s = edf[col].dropna()
        if col == "_new_ratio":
            dist_rows.append({"metric": label, "p25": f"{s.quantile(.25):.0%}",
                             "median": f"{s.median():.0%}", "p75": f"{s.quantile(.75):.0%}",
                             "mean": f"{s.mean():.0%}"})
        else:
            dist_rows.append({"metric": label, "p25": f"{int(s.quantile(.25)):,}",
                             "median": f"{int(s.median()):,}", "p75": f"{int(s.quantile(.75)):,}",
                             "mean": f"{int(s.mean()):,}"})
    dist_rows.insert(0, {"metric": "--- OVERALL ---", "p25": "", "median": "", "p75": "", "mean": ""})
    parts.append(pd.DataFrame(dist_rows))

    # By direction
    for direction in ["Inbound", "Outbound"]:
        sub = edf[edf["_direction"] == direction]
        if sub.empty: continue
        dir_rows = [{
            "metric": f"--- {direction.upper()} ---",
            "p25": "", "median": "", "p75": "", "mean": ""
        }]
        for label, col in [("new_content_chars", "_new_len"), ("new_content_ratio", "_new_ratio")]:
            s = sub[col].dropna()
            if col == "_new_ratio":
                dir_rows.append({"metric": label, "p25": f"{s.quantile(.25):.0%}",
                                "median": f"{s.median():.0%}", "p75": f"{s.quantile(.75):.0%}",
                                "mean": f"{s.mean():.0%}"})
            else:
                dir_rows.append({"metric": label, "p25": f"{int(s.quantile(.25)):,}",
                                "median": f"{int(s.median()):,}", "p75": f"{int(s.quantile(.75)):,}",
                                "mean": f"{int(s.mean()):,}"})
        parts.append(pd.DataFrame(dir_rows))

    # By case subject (client inbound only)
    client_in = edf[edf["_is_client"] & edf["_is_inbound"]]
    if not client_in.empty:
        by_subj = client_in.groupby("_case_subject").agg(
            emails=("_email_id", "size"),
            median_new_chars=("_new_len", "median"),
            median_new_ratio=("_new_ratio", "median"),
            median_body_chars=("_body_len", "median"),
        ).reset_index().sort_values("emails", ascending=False).head(12)
        by_subj["median_new_chars"] = by_subj["median_new_chars"].round(0).astype(int)
        by_subj["median_new_ratio"] = (by_subj["median_new_ratio"] * 100).round(0).astype(int).astype(str) + "%"
        by_subj["median_body_chars"] = by_subj["median_body_chars"].round(0).astype(int)
        by_subj.rename(columns={"_case_subject": "subject"}, inplace=True)
        sep = pd.DataFrame({"metric": ["--- BY SUBJECT (inbound client) ---"]})
        parts.append(sep)
        parts.append(by_subj)

    return pd.concat(parts, ignore_index=True, sort=False)


def sheet_03_topic_discovery(edf):
    """I03: NMF topic modeling to discover hidden intent families within broad subjects.
    This is the 'hidden intent taxonomy' the blueprint asks for."""
    client_in = edf[edf["_is_client"] & edf["_is_inbound"]].copy()
    if len(client_in) < 20:
        return pd.DataFrame({"note": ["Not enough inbound client emails for topic modeling"]})

    # Focus on broad subjects with enough volume
    broad_subjects = client_in["_case_subject"].value_counts()
    broad_subjects = broad_subjects[broad_subjects >= 10].index.tolist()[:8]

    all_results = []

    for subj in broad_subjects:
        sub = client_in[client_in["_case_subject"] == subj]
        texts = sub["_text_canonical"].tolist()
        texts = [t for t in texts if len(t) >= 30]
        if len(texts) < 10:
            continue

        # TF-IDF + NMF
        n_topics = min(5, max(2, len(texts) // 15))
        try:
            tfidf = TfidfVectorizer(max_features=1500, ngram_range=(1, 2),
                                     min_df=2, max_df=0.85, stop_words="english")
            mat = tfidf.fit_transform(texts)
            nmf = NMF(n_components=n_topics, random_state=42, max_iter=300)
            W = nmf.fit_transform(mat)
            feature_names = tfidf.get_feature_names_out()

            for topic_idx in range(n_topics):
                top_words = [feature_names[i] for i in nmf.components_[topic_idx].argsort()[::-1][:8]]
                # Filter out placeholder tokens
                top_words = [w for w in top_words if not w.startswith("_")][:6]
                # Count emails assigned to this topic
                assigned = (W.argmax(axis=1) == topic_idx).sum()
                all_results.append({
                    "case_subject": subj,
                    "topic_id": topic_idx + 1,
                    "top_terms": " | ".join(top_words),
                    "emails_assigned": int(assigned),
                    "pct_of_subject": pct(assigned, len(texts)),
                    "subject_total": len(texts),
                })
        except Exception as e:
            warn(f"Topic modeling failed for {subj}: {e}")

    if not all_results:
        return pd.DataFrame({"note": ["Topic modeling produced no results"]})

    return pd.DataFrame(all_results).sort_values(["case_subject", "emails_assigned"], ascending=[True, False])


def sheet_04_missing_info(edf):
    """I04: Where are inbound emails requesting or flagging missing information?
    Connects to the missing-info detection use case."""
    client_in = edf[edf["_is_client"] & edf["_is_inbound"]].copy()
    if client_in.empty:
        return pd.DataFrame({"note": ["No inbound client emails"]})

    parts = []

    # Summary by subject
    by_subj = client_in.groupby("_case_subject").agg(
        emails=("_email_id", "size"),
        missing_cue_count=("_has_missing_cue", "sum"),
        doc_terms_mentioned=("_doc_terms_found", lambda x: (x.str.len() > 0).sum()),
        median_case_hours=("_hours", "median"),
    ).reset_index()
    by_subj["missing_cue_pct"] = (100 * by_subj["missing_cue_count"] / by_subj["emails"]).round(1)
    by_subj["doc_terms_pct"] = (100 * by_subj["doc_terms_mentioned"] / by_subj["emails"]).round(1)
    by_subj["median_case_hours"] = by_subj["median_case_hours"].round(1)
    by_subj = by_subj.sort_values("missing_cue_count", ascending=False).head(12)
    by_subj.rename(columns={"_case_subject": "subject"}, inplace=True)
    by_subj.insert(0, "section", "BY SUBJECT")
    parts.append(by_subj)

    # Most common document terms mentioned across all inbound
    all_terms = client_in["_doc_terms_found"].str.split(", ").explode().str.strip()
    all_terms = all_terms[all_terms.str.len() > 0]
    if len(all_terms) > 0:
        term_counts = all_terms.value_counts().head(15).reset_index()
        term_counts.columns = ["subject", "emails"]
        term_counts["missing_cue_count"] = ""
        term_counts["doc_terms_mentioned"] = ""
        term_counts["missing_cue_pct"] = ""
        term_counts["doc_terms_pct"] = ""
        term_counts["median_case_hours"] = ""
        term_counts.insert(0, "section", "TOP DOC TERMS")
        parts.append(term_counts)

    # Sample emails with missing-info signals
    samples = client_in[client_in["_has_missing_cue"]].head(10)
    if not samples.empty:
        sample_df = samples[["_case_subject", "_subject_clean", "_doc_terms_found", "_new_text"]].copy()
        sample_df["_new_text"] = sample_df["_new_text"].apply(lambda x: trunc(x, 200))
        sample_df.rename(columns={
            "_case_subject": "subject", "_subject_clean": "email_subject",
            "_doc_terms_found": "doc_terms", "_new_text": "sample_new_content"
        }, inplace=True)
        sample_df.insert(0, "section", "EXAMPLES")
        parts.append(sample_df)

    return pd.concat(parts, ignore_index=True, sort=False)


def sheet_05_outbound_templates(edf):
    """I05: How repetitive are outbound banker responses?
    This is the draft-assist ROI measurement."""
    client_out = edf[edf["_is_client"] & edf["_is_outbound"]].copy()
    if client_out.empty:
        return pd.DataFrame({"note": ["No outbound client emails"]})

    parts = []

    # For each subject, compute outbound template coverage
    top_subjects = client_out["_case_subject"].value_counts().head(8).index.tolist()

    for subj in top_subjects:
        sub = client_out[client_out["_case_subject"] == subj]
        texts = sub["_body_clean"].tolist()
        texts = [t for t in texts if len(t) >= 50]
        if len(texts) < 5:
            continue

        # Use TF-IDF + cosine similarity to find clusters
        try:
            tfidf = TfidfVectorizer(max_features=1000, ngram_range=(1, 2),
                                     min_df=2, max_df=0.9, stop_words="english")
            mat = tfidf.fit_transform(texts)
            sim_matrix = cosine_similarity(mat)

            # Greedy clustering: group emails with >0.7 similarity
            n = len(texts)
            assigned = [False] * n
            clusters = []
            for i in range(n):
                if assigned[i]: continue
                group = [i]
                for j in range(i + 1, n):
                    if assigned[j]: continue
                    if sim_matrix[i][j] >= 0.65:
                        group.append(j)
                        assigned[j] = True
                assigned[i] = True
                if len(group) >= 2:
                    clusters.append(group)

            clusters.sort(key=len, reverse=True)

            total_clustered = sum(len(c) for c in clusters)
            coverage_pct = 100 * total_clustered / len(texts) if texts else 0

            # Report top clusters
            for ci, cluster in enumerate(clusters[:5]):
                sample_text = trunc(texts[cluster[0]], 150)
                parts.append(pd.DataFrame([{
                    "subject": subj,
                    "cluster_id": ci + 1,
                    "cluster_size": len(cluster),
                    "total_outbound": len(texts),
                    "template_coverage_pct": f"{coverage_pct:.1f}%",
                    "sample_text": sample_text,
                }]))

            if not clusters:
                parts.append(pd.DataFrame([{
                    "subject": subj,
                    "cluster_id": 0,
                    "cluster_size": 0,
                    "total_outbound": len(texts),
                    "template_coverage_pct": "0.0%",
                    "sample_text": "(no repeated templates found)",
                }]))

        except Exception as e:
            warn(f"Template clustering failed for {subj}: {e}")

    if not parts:
        return pd.DataFrame({"note": ["No template clusters found"]})

    return pd.concat(parts, ignore_index=True, sort=False)


def sheet_06_conversation_threads(edf):
    """I06: Thread-level analysis — how many emails per case, which are heavy?"""
    client = edf[edf["_is_client"] & edf["_has_case"]].copy()
    if client.empty:
        return pd.DataFrame({"note": ["No client emails with case linkage"]})

    client = client.sort_values(["_case_ref", "_created_dt"])

    thread_stats = client.groupby("_case_ref").agg(
        email_count=("_email_id", "size"),
        inbound=("_is_inbound", "sum"),
        outbound=("_is_outbound", "sum"),
        total_new_chars=("_new_len", "sum"),
        total_quoted_chars=("_quoted_len", "sum"),
        median_new_ratio=("_new_ratio", "median"),
        actors=("_from", "nunique"),
        case_subject=("_case_subject", "first"),
        company=("_company", "first"),
        case_hours=("_hours", "first"),
        is_resolved=("_is_resolved", "first"),
        has_missing_cue=("_has_missing_cue", "any"),
        has_followup_cue=("_has_followup_cue", "any"),
    ).reset_index()

    thread_stats["new_content_pct"] = (
        100 * thread_stats["total_new_chars"] /
        (thread_stats["total_new_chars"] + thread_stats["total_quoted_chars"]).replace(0, 1)
    ).round(0).astype(int)

    parts = []

    # Distribution
    dist = thread_stats["email_count"]
    dist_rows = [
        {"section": "THREAD DISTRIBUTION", "metric": "cases_with_emails", "value": f"{len(thread_stats):,}"},
        {"section": "THREAD DISTRIBUTION", "metric": "median_emails", "value": f"{dist.median():.1f}"},
        {"section": "THREAD DISTRIBUTION", "metric": "p90_emails", "value": f"{dist.quantile(.9):.1f}"},
        {"section": "THREAD DISTRIBUTION", "metric": "max_emails", "value": f"{int(dist.max())}"},
        {"section": "THREAD DISTRIBUTION", "metric": "heavy_threads_5plus", "value": f"{(dist >= 5).sum():,} ({pct((dist >= 5).sum(), len(dist))})"},
    ]
    parts.append(pd.DataFrame(dist_rows))

    # By subject
    by_subj = thread_stats.groupby("case_subject").agg(
        cases=("case_subject", "size"),
        median_emails=("email_count", "median"),
        p90_emails=("email_count", lambda s: s.quantile(.9)),
        pct_with_missing_cue=("has_missing_cue", "mean"),
        pct_with_followup=("has_followup_cue", "mean"),
        median_case_hours=("case_hours", "median"),
    ).reset_index().sort_values("cases", ascending=False).head(12)
    by_subj["median_emails"] = by_subj["median_emails"].round(1)
    by_subj["p90_emails"] = by_subj["p90_emails"].round(1)
    by_subj["pct_with_missing_cue"] = (by_subj["pct_with_missing_cue"] * 100).round(1).astype(str) + "%"
    by_subj["pct_with_followup"] = (by_subj["pct_with_followup"] * 100).round(1).astype(str) + "%"
    by_subj["median_case_hours"] = by_subj["median_case_hours"].round(1)
    by_subj.insert(0, "section", "BY SUBJECT")
    parts.append(by_subj)

    # Heaviest threads
    heavy = thread_stats.nlargest(15, "email_count")[
        ["_case_ref", "case_subject", "email_count", "inbound", "outbound",
         "actors", "new_content_pct", "case_hours", "has_missing_cue", "has_followup_cue"]
    ].copy()
    heavy.insert(0, "section", "HEAVIEST THREADS")
    heavy.rename(columns={"_case_ref": "case_number"}, inplace=True)
    parts.append(heavy)

    return pd.concat(parts, ignore_index=True, sort=False)


def sheet_07_signal_by_subject(edf):
    """I07: Cross-tabulation of all signals by subject.
    Shows where missing-info, follow-up, and urgency concentrate."""
    client_in = edf[edf["_is_client"] & edf["_is_inbound"]].copy()
    if client_in.empty:
        return pd.DataFrame({"note": ["No inbound client emails"]})

    by_subj = client_in.groupby("_case_subject").agg(
        total_inbound=("_email_id", "size"),
        missing_info=("_has_missing_cue", "sum"),
        followup=("_has_followup_cue", "sum"),
        urgency=("_has_urgency_cue", "sum"),
        median_new_chars=("_new_len", "median"),
        median_case_hours=("_hours", "median"),
        pct_unresolved=("_is_resolved", lambda s: f"{100 * (~s.astype(bool)).mean():.1f}%"),
    ).reset_index().sort_values("total_inbound", ascending=False).head(12)

    by_subj["missing_info_pct"] = (100 * by_subj["missing_info"] / by_subj["total_inbound"]).round(1).astype(str) + "%"
    by_subj["followup_pct"] = (100 * by_subj["followup"] / by_subj["total_inbound"]).round(1).astype(str) + "%"
    by_subj["urgency_pct"] = (100 * by_subj["urgency"] / by_subj["total_inbound"]).round(1).astype(str) + "%"
    by_subj["median_new_chars"] = by_subj["median_new_chars"].round(0).astype(int)
    by_subj["median_case_hours"] = by_subj["median_case_hours"].round(1)

    by_subj.rename(columns={"_case_subject": "subject"}, inplace=True)
    return by_subj


def sheet_08_triage_by_intent(edf):
    """I08: Triage delay broken out by email content signals.
    Connects D16 triage delay to text-derived features."""
    if "_triage_minutes" not in edf.columns:
        return pd.DataFrame({"note": ["Triage delay not available — SLA Start or Created On missing"]})

    client = edf[edf["_is_client"]].copy()
    valid = client[client["_triage_minutes"].notna() & (client["_triage_minutes"] > 0)].copy()
    if valid.empty:
        return pd.DataFrame({"note": ["No cases with measurable triage delay"]})

    parts = []

    # Triage delay by signal presence
    for label, col in [("has_missing_cue", "_has_missing_cue"),
                        ("has_followup_cue", "_has_followup_cue"),
                        ("has_urgency_cue", "_has_urgency_cue")]:
        flagged = valid[valid[col]]
        unflagged = valid[~valid[col]]
        parts.append(pd.DataFrame([{
            "section": "TRIAGE BY SIGNAL",
            "signal": label,
            "flagged_count": len(flagged),
            "flagged_median_triage_min": f"{flagged['_triage_minutes'].median():.1f}" if len(flagged) else "N/A",
            "unflagged_count": len(unflagged),
            "unflagged_median_triage_min": f"{unflagged['_triage_minutes'].median():.1f}" if len(unflagged) else "N/A",
        }]))

    # Triage delay by new-content volume
    valid["_new_content_bucket"] = pd.cut(
        valid["_new_len"], bins=[0, 200, 500, 1500, float("inf")],
        labels=["<200 chars", "200-500", "500-1500", "1500+"], right=True)
    by_content = valid.groupby("_new_content_bucket", observed=True).agg(
        emails=("_email_id", "size"),
        median_triage_min=("_triage_minutes", "median"),
    ).reset_index()
    by_content["median_triage_min"] = by_content["median_triage_min"].round(1)
    by_content.insert(0, "section", "TRIAGE BY CONTENT LENGTH")
    by_content.rename(columns={"_new_content_bucket": "signal"}, inplace=True)
    parts.append(by_content)

    return pd.concat(parts, ignore_index=True, sort=False)


def sheet_09_evidence_scorecard(edf):
    """I09: Final scorecard — does the email data support each GenAI use case?
    Tests against the blueprint's decision thresholds."""
    client_in = edf[edf["_is_client"] & edf["_is_inbound"]]
    client_out = edf[edf["_is_client"] & edf["_is_outbound"]]
    n_in = len(client_in)
    n_out = len(client_out)

    rows = []

    # 1. Triage/Routing — is there enough text to classify?
    text_coverage = (client_in["_new_len"] > 100).sum()
    rows.append({
        "use_case": "Triage / Routing",
        "metric": f"{pct(text_coverage, n_in)} of inbound emails have >100 chars new content",
        "threshold": "Need >60% with classifiable text",
        "verdict": "PASS" if text_coverage / max(n_in, 1) > 0.6 else "FAIL",
        "implication": "Emails carry enough text for classification beyond the structured Subject field",
    })

    # 2. Summarization — is there enough content worth summarizing?
    long_emails = (client_in["_new_len"] > 500).sum()
    rows.append({
        "use_case": "Summarization",
        "metric": f"{pct(long_emails, n_in)} of inbound emails have >500 chars new content",
        "threshold": "Need >40% with substantial new content",
        "verdict": "PASS" if long_emails / max(n_in, 1) > 0.4 else "MARGINAL",
        "implication": "Many emails have enough new prose to justify summarization, though some are short",
    })

    # 3. Missing-info detection
    missing_count = client_in["_has_missing_cue"].sum()
    rows.append({
        "use_case": "Missing-Info Detection",
        "metric": f"{missing_count:,} inbound emails ({pct(missing_count, n_in)}) contain missing-info language",
        "threshold": "Need >10% in at least 2 subject families",
        "verdict": "CHECK TABLE" if missing_count > 0 else "FAIL",
        "implication": "Keyword-based detection finds signals; LLM-based would find more nuanced patterns",
    })

    # 4. Draft reply — are outbound responses repetitive?
    rows.append({
        "use_case": "Draft Reply Assistance",
        "metric": f"{n_out:,} outbound client emails available for template analysis",
        "threshold": "Need >25% of outbound in a subject covered by top-5 templates",
        "verdict": "CHECK TEMPLATE SHEET",
        "implication": "Template coverage determines draft-assist ROI by subject",
    })

    # 5. Escalation — do signals correlate with slow cases?
    if "_hours" in client_in.columns:
        flagged = client_in[client_in["_has_followup_cue"] | client_in["_has_urgency_cue"]]
        unflagged = client_in[~(client_in["_has_followup_cue"] | client_in["_has_urgency_cue"])]
        flag_med = flagged["_hours"].median() if len(flagged) else np.nan
        unflag_med = unflagged["_hours"].median() if len(unflagged) else np.nan
        rows.append({
            "use_case": "Escalation Detection",
            "metric": f"Flagged email cases: {flag_med:.1f}h median vs unflagged: {unflag_med:.1f}h",
            "threshold": "Flagged cases should be materially slower",
            "verdict": "SIGNAL" if pd.notna(flag_med) and pd.notna(unflag_med) and flag_med > unflag_med * 1.3 else "WEAK",
            "implication": "Follow-up/urgency language in emails may predict slower cases",
        })

    # 6. Context reconstruction — do heavy threads have more cognitive load?
    rows.append({
        "use_case": "Thread Summarization",
        "metric": f"See content structure and thread analysis sheets",
        "threshold": "Heavy threads should have low new-content ratio (lots of quoted history)",
        "verdict": "CHECK THREAD SHEET",
        "implication": "If bankers spend time re-reading quoted content, summarization has clear value",
    })

    return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════

def main():
    start = datetime.datetime.now()
    log(f"WAB Email Deep Insights v2 — {start.strftime('%Y-%m-%d %H:%M:%S')}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    cases_raw = read_file(CASE_FILE, "Cases")
    emails_raw = read_file(EMAIL_FILE, "Emails")
    if cases_raw.empty or emails_raw.empty:
        log("FATAL: required files not loaded.")
        return

    log("\n--- Preparing data ---")
    cases = prepare_cases(cases_raw)
    edf = prepare_emails(emails_raw, cases)
    log(f"  Client emails: {edf['_is_client'].sum():,}")
    log(f"  Inbound: {edf['_is_inbound'].sum():,} | Outbound: {edf['_is_outbound'].sum():,}")

    log("\n--- Building sheets ---")
    sheets = OrderedDict()

    log("  I01 Data Scope")
    sheets["I01_DataScope"] = sheet_01_scope(edf)

    log("  I02 Content Structure")
    sheets["I02_ContentStructure"] = sheet_02_content_structure(edf)

    log("  I03 Topic Discovery")
    sheets["I03_TopicDiscovery"] = sheet_03_topic_discovery(edf)

    log("  I04 Missing Info")
    sheets["I04_MissingInfo"] = sheet_04_missing_info(edf)

    log("  I05 Outbound Templates")
    sheets["I05_OutboundTemplates"] = sheet_05_outbound_templates(edf)

    log("  I06 Conversation Threads")
    sheets["I06_ConversationThreads"] = sheet_06_conversation_threads(edf)

    log("  I07 Signal by Subject")
    sheets["I07_SignalBySubject"] = sheet_07_signal_by_subject(edf)

    log("  I08 Triage by Intent")
    sheets["I08_TriageByIntent"] = sheet_08_triage_by_intent(edf)

    log("  I09 Evidence Scorecard")
    sheets["I09_EvidenceScorecard"] = sheet_09_evidence_scorecard(edf)

    # Write
    log(f"\nWriting: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            write_sheet(writer, name[:31], sdf)
    log("  Done.")

    log(f"Writing: {OUTPUT_MD}")
    md = [
        "# WAB Email Deep Insights v2",
        f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
        "## Three Questions This Module Answers",
        "1. What are bankers actually reading? (I02 Content Structure, I03 Topic Discovery)",
        "2. How much of the work is repetitive? (I05 Outbound Templates, I06 Threads)",
        "3. Where is the text evidence for GenAI use cases? (I04 Missing Info, I07 Signals, I08 Triage, I09 Scorecard)",
        "",
        "## Sheets",
    ]
    for name, sdf in sheets.items():
        md.append(f"- **{name}**: {len(sdf) if sdf is not None else 0} rows")
    if WARN:
        md.extend(["", "## Warnings"])
        for w in WARN: md.append(f"- {w}")
    md.extend(["", "## Log", "```"] + LOG + ["```"])
    with open(OUTPUT_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(md))

    elapsed = (datetime.datetime.now() - start).total_seconds()
    log(f"\nCompleted in {elapsed:.1f}s")


if __name__ == "__main__":
    main()
