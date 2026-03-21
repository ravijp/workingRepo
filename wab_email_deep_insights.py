"""
WAB Email Deep Insights — Phase 1 Evidence Expansion
====================================================
Standalone VDI script for extracting deeper email insights that go beyond
the existing HTML story and first-round profiling modules.

Reads Cases + Emails files and produces:
- one Excel workbook with decision-grade email insight sheets
- one markdown summary with headline observations and warnings

Recommended packages:
    pip install scikit-learn rapidfuzz regex

The script degrades gracefully if optional packages are unavailable.
"""

# EDIT THESE 3 VARIABLES BEFORE RUNNING
CASE_FILE = r"C:\Users\YourName\Desktop\AAB All Cases.xlsx"
EMAIL_FILE = r"C:\Users\YourName\Desktop\ALL EMAIL Files.xlsx"
OUTPUT_DIR = r"C:\Users\YourName\Desktop\wab_output"

import os
import re
import html
import datetime
import warnings
from collections import Counter, OrderedDict

import numpy as np
import pandas as pd
from bs4 import BeautifulSoup

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

try:
    import regex as regex_re
except Exception:
    regex_re = re

try:
    from rapidfuzz import fuzz
except Exception:
    fuzz = None

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import linear_kernel
except Exception:
    TfidfVectorizer = None
    linear_kernel = None


OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "wab_email_deep_insights.xlsx")
OUTPUT_MD = os.path.join(OUTPUT_DIR, "wab_email_deep_insights_summary.md")
LOG = []
WARN = []
TRUNC = 160

INTERNAL_COMPANIES = {"AAB ADMIN", "WAB ADMIN", "AAB ADMIN -", "WAB ADMIN -"}
STOPWORDS = {
    "a", "an", "and", "are", "as", "at", "be", "been", "but", "by", "can", "could",
    "do", "for", "from", "had", "has", "have", "if", "in", "into", "is", "it", "its",
    "just", "me", "my", "not", "of", "on", "or", "our", "please", "re", "regards",
    "regarding", "so", "that", "the", "their", "them", "there", "this", "to", "up",
    "us", "was", "we", "were", "will", "with", "you", "your",
    "account", "request", "questions", "question", "research", "maintenance",
    "fw", "fwd", "re", "subject", "thanks", "thank", "hello", "hi"
}
DOC_TERMS = {
    "signature card", "voided check", "void check", "management agreement",
    "ein", "w9", "w-9", "tin", "minutes", "resolution", "board resolution",
    "articles", "driver license", "id", "identification", "operating account",
    "reserve account", "cdars", "ics"
}
MISSING_CUES = {
    "missing", "not received", "still need", "need the", "need your", "please provide",
    "please send", "awaiting", "required", "incomplete", "once we receive",
    "can you send", "have not received", "pending receipt"
}
FOLLOWUP_CUES = {
    "following up", "follow up", "just following up", "checking in", "circling back",
    "still waiting", "any update", "please advise", "status?", "status update",
    "when can", "have you had a chance"
}
URGENCY_CUES = {
    "urgent", "asap", "immediately", "today", "end of day", "eod", "rush",
    "critical", "time sensitive", "deadline", "priority", "expedite"
}


def log(msg):
    LOG.append(msg)
    print(msg)


def warn(msg):
    WARN.append(msg)
    log(f"  WARNING: {msg}")


def trunc(val, n=TRUNC):
    if pd.isna(val):
        return ""
    s = str(val).replace("\r", " ").replace("\n", " ").strip()
    return s[:n] + "..." if len(s) > n else s


def norm_col(name):
    if not isinstance(name, str):
        return ""
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


def pct(num, denom):
    return f"{100 * num / denom:.1f}%" if denom else "N/A"


def write_sheet(writer, name, df):
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        df = pd.DataFrame({"note": ["No data"]})
    sheet_name = name[:31]
    df.to_excel(writer, sheet_name=sheet_name, index=False, freeze_panes=(1, 0))
    ws = writer.sheets[sheet_name]
    for col_cells in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 60)


def read_file(path, label):
    log(f"\nReading {label}: {path}")
    if not os.path.isfile(path):
        warn(f"File not found: {path}")
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, header=0, engine="openpyxl")
    except Exception as exc:
        warn(f"Failed to read {label}: {exc}")
        return pd.DataFrame()
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    df = df.dropna(axis=1, how="all")
    log(f"  Rows: {len(df):,} | Columns: {len(df.columns)}")
    return df


def strip_html(text):
    if pd.isna(text):
        return ""
    s = str(text)
    s = re.sub(r"<(style|script)[^>]*>.*?</\1>", " ", s, flags=re.I | re.S)
    soup = BeautifulSoup(s, "html.parser")
    text = soup.get_text("\n")
    text = html.unescape(text)
    text = re.sub(r"\r", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def remove_email_noise(text):
    if not text:
        return ""
    s = text
    banner_patterns = [
        r"attention:\s*this email originated from outside.*",
        r"external email warning.*",
        r"caution:\s*external.*",
    ]
    for pat in banner_patterns:
        s = regex_re.sub(pat, " ", s, flags=regex_re.I | regex_re.S)
    s = regex_re.sub(r"from:\s.*?sent:\s.*?to:\s.*?subject:\s.*", " ", s, flags=regex_re.I | regex_re.S)
    s = regex_re.sub(r"on .* wrote:\s*", " ", s, flags=regex_re.I)
    s = regex_re.sub(r"_{4,}.*", " ", s, flags=regex_re.S)
    s = regex_re.sub(r"-{4,}.*", " ", s, flags=regex_re.S)
    s = regex_re.sub(r"\b(?:best regards|regards|thanks|thank you),?\b.*", " ", s, flags=regex_re.I | regex_re.S)
    s = regex_re.sub(r"\s+", " ", s)
    return s.strip()


def canonical_text(text):
    if not text:
        return ""
    s = text.lower()
    s = regex_re.sub(r"\b[\w\.-]+@[\w\.-]+\.\w+\b", " EMAIL ", s)
    s = regex_re.sub(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", " DATE ", s)
    s = regex_re.sub(r"\b\d+\b", " NUM ", s)
    s = regex_re.sub(r"[^a-z0-9\s]", " ", s)
    s = regex_re.sub(r"\s+", " ", s).strip()
    return s


def tokens(text):
    if not text:
        return []
    return [t for t in canonical_text(text).split() if len(t) > 2 and t not in STOPWORDS]


def fingerprint(text, top_n=14):
    toks = tokens(text)
    if not toks:
        return ""
    c = Counter(toks)
    return " ".join(tok for tok, _ in c.most_common(top_n))


def make_ngrams(tokens_list, n=2):
    if len(tokens_list) < n:
        return []
    return [" ".join(tokens_list[i:i + n]) for i in range(len(tokens_list) - n + 1)]


def top_phrases(series, top_n=12):
    grams = Counter()
    for text in series.dropna():
        toks = tokens(text)
        grams.update(make_ngrams(toks, 2))
        grams.update(make_ngrams(toks, 3))
    rows = []
    for phrase, count in grams.most_common(top_n):
        if any(tok in STOPWORDS for tok in phrase.split()):
            continue
        rows.append({"phrase": phrase, "count": count})
    return pd.DataFrame(rows)


def prepare_cases(cases):
    df = cases.copy()
    col_map = {
        "case_number": find_col(df, "Case Number"),
        "company": find_col(df, "Company Name (Company) (Company)", "Company Name"),
        "subject": find_col(df, "Subject"),
        "activity_subject": find_col(df, "Activity Subject"),
        "status": find_col(df, "Status"),
        "status_reason": find_col(df, "Status Reason"),
        "created_on": find_col(df, "Created On"),
        "resolved_hrs": find_col(df, "Resolved In Hours"),
        "owner": find_col(df, "Manager (Owning User) (User)", "Owner"),
        "pod": find_col(df, "POD Name (Owning User) (User)", "POD Name"),
    }

    if col_map["created_on"]:
        df["_created_dt"] = safe_dt(df[col_map["created_on"]])
    if col_map["resolved_hrs"]:
        df["_hours"] = safe_numeric(df[col_map["resolved_hrs"]])
    if col_map["company"]:
        df["_company_clean"] = df[col_map["company"]].fillna("(blank)").astype(str).str.strip()
        df["_is_internal"] = df["_company_clean"].str.upper().apply(
            lambda x: any(x.startswith(k) for k in INTERNAL_COMPANIES) or x == "(BLANK)"
        )
    else:
        df["_company_clean"] = "(blank)"
        df["_is_internal"] = False

    df["_case_number"] = df[col_map["case_number"]].astype(str) if col_map["case_number"] else ""
    df["_subject"] = df[col_map["subject"]].fillna("(blank)").astype(str).str.strip() if col_map["subject"] else "(blank)"
    df["_activity_subject"] = df[col_map["activity_subject"]].fillna("").astype(str).str.strip() if col_map["activity_subject"] else ""
    df["_owner"] = df[col_map["owner"]].fillna("(blank)").astype(str).str.strip() if col_map["owner"] else "(blank)"
    df["_pod"] = df[col_map["pod"]].fillna("(blank)").astype(str).str.strip() if col_map["pod"] else "(blank)"

    src_sr = col_map["status_reason"] or col_map["status"]
    if src_sr:
        df["_status_clean"] = df[src_sr].fillna("").astype(str).str.lower()
        df["_is_resolved"] = df["_status_clean"].apply(
            lambda x: any(k in x for k in ("resolved", "closed", "cancelled", "canceled", "complete"))
        )
    else:
        df["_status_clean"] = ""
        df["_is_resolved"] = True

    df._col_map = col_map
    return df


def prepare_emails(emails, cases):
    df = emails.copy()
    col_map = {
        "subject": find_col(df, "Subject"),
        "description": find_col(df, "Description"),
        "from": find_col(df, "From"),
        "to": find_col(df, "To"),
        "status_reason": find_col(df, "Status Reason"),
        "created_on": find_col(df, "Created On"),
        "owner": find_col(df, "Owner"),
        "priority": find_col(df, "Priority"),
        "case_number": find_col(df, "Case Number (Regarding) (Case)", "Case Number (Regarding)", "Regarding"),
        "case_subject": find_col(df, "Subject (Regarding) (Case)", "Subject Path (Regarding) (Case)"),
    }

    df["_email_id"] = np.arange(1, len(df) + 1)
    if col_map["created_on"]:
        df["_created_dt"] = safe_dt(df[col_map["created_on"]])
        df["_date"] = df["_created_dt"].dt.date
        df["_hour"] = df["_created_dt"].dt.hour
    else:
        df["_created_dt"] = pd.NaT
        df["_date"] = pd.NaT
        df["_hour"] = np.nan

    df["_subject_raw"] = df[col_map["subject"]].fillna("").astype(str) if col_map["subject"] else ""
    df["_body_html"] = df[col_map["description"]].fillna("").astype(str) if col_map["description"] else ""
    df["_from"] = df[col_map["from"]].fillna("").astype(str) if col_map["from"] else ""
    df["_to"] = df[col_map["to"]].fillna("").astype(str) if col_map["to"] else ""
    df["_status_reason"] = df[col_map["status_reason"]].fillna("").astype(str) if col_map["status_reason"] else ""
    df["_priority"] = df[col_map["priority"]].fillna("(blank)").astype(str) if col_map["priority"] else "(blank)"
    df["_case_number"] = df[col_map["case_number"]].fillna("").astype(str) if col_map["case_number"] else ""
    df["_linked_case_subject"] = df[col_map["case_subject"]].fillna("(blank)").astype(str) if col_map["case_subject"] else "(blank)"
    df["_owner"] = df[col_map["owner"]].fillna("(blank)").astype(str) if col_map["owner"] else "(blank)"
    df["_has_case"] = df["_case_number"].str.strip().ne("")

    log("  Stripping and normalizing email text...")
    df["_body_text_raw"] = df["_body_html"].apply(strip_html)
    df["_body_text"] = df["_body_text_raw"].apply(remove_email_noise)
    df["_subject_clean"] = df["_subject_raw"].apply(remove_email_noise)
    df["_text_for_nlp"] = (df["_subject_clean"] + " " + df["_body_text"].str[:1200]).str.strip()
    df["_text_canonical"] = df["_text_for_nlp"].apply(canonical_text)
    df["_body_len"] = df["_body_text"].str.len()
    df["_token_count"] = df["_body_text"].apply(lambda x: len(tokens(x)))
    df["_quoted_depth_proxy"] = df["_body_text_raw"].fillna("").astype(str).str.count(r"From:|Sent:|To:|Subject:")
    df["_participant_signature"] = (df["_from"].fillna("") + " -> " + df["_to"].fillna("")).str[:200]

    dir_series = df["_status_reason"].str.lower()
    df["_direction"] = np.where(
        dir_series.str.contains("sent|completed|outgoing|outbound"),
        "Outbound",
        np.where(dir_series.str.contains("received|incoming|inbound"), "Inbound", "Unknown")
    )

    case_cols = ["_case_number", "_subject", "_activity_subject", "_owner", "_pod", "_company_clean", "_hours", "_is_resolved"]
    case_lookup = cases[[c for c in case_cols if c in cases.columns]].drop_duplicates(subset=["_case_number"])
    df = df.merge(case_lookup, on="_case_number", how="left", suffixes=("", "_case"))
    df["_subject_final"] = df["_subject"].fillna(df["_linked_case_subject"]).fillna("(blank)").astype(str)
    df["_company_clean"] = df["_company_clean"].fillna("(blank)").astype(str)
    df["_is_internal"] = df["_company_clean"].str.upper().apply(
        lambda x: any(x.startswith(k) for k in INTERNAL_COMPANIES) or x == "(BLANK)"
    )
    df["_is_client_case"] = df["_has_case"] & (~df["_is_internal"])
    df["_is_inbound"] = df["_direction"].eq("Inbound")
    df["_is_outbound"] = df["_direction"].eq("Outbound")

    df._col_map = col_map
    return df


def sheet_01_data_scope(edf):
    rows = [
        {"section": "TOTALS", "metric": "email_rows", "value": f"{len(edf):,}"},
        {"section": "TOTALS", "metric": "linked_to_case", "value": f"{edf['_has_case'].sum():,} ({pct(edf['_has_case'].sum(), len(edf))})"},
        {"section": "TOTALS", "metric": "client_case_emails", "value": f"{edf['_is_client_case'].sum():,} ({pct(edf['_is_client_case'].sum(), len(edf))})"},
        {"section": "TOTALS", "metric": "inbound", "value": f"{edf['_is_inbound'].sum():,} ({pct(edf['_is_inbound'].sum(), len(edf))})"},
        {"section": "TOTALS", "metric": "outbound", "value": f"{edf['_is_outbound'].sum():,} ({pct(edf['_is_outbound'].sum(), len(edf))})"},
        {"section": "TOTALS", "metric": "median_body_chars", "value": f"{int(edf['_body_len'].median()):,}"},
        {"section": "TOTALS", "metric": "median_tokens", "value": f"{int(edf['_token_count'].median()):,}"},
    ]
    if "_created_dt" in edf.columns and edf["_created_dt"].notna().any():
        rows.append({"section": "DATES", "metric": "min_created", "value": str(edf["_created_dt"].min())})
        rows.append({"section": "DATES", "metric": "max_created", "value": str(edf["_created_dt"].max())})
        rows.append({"section": "DATES", "metric": "distinct_days", "value": f"{edf['_date'].nunique():,}"})
    return pd.DataFrame(rows)


def infer_intent_label(text, fallback_subject="(blank)"):
    text_l = canonical_text(text)
    for term in (
        "signature card", "wire transfer", "transfer request", "close account",
        "new account", "cd maintenance", "cdars", "ics", "research request",
        "statement request", "fraud alert", "positive pay", "account maintenance"
    ):
        if term in text_l:
            return term
    toks = tokens(text)
    if len(toks) >= 2:
        return " ".join(toks[:2])
    if toks:
        return toks[0]
    return fallback_subject.lower()[:40]


def sheet_02_hidden_intents(edf):
    subset = edf[(edf["_is_client_case"]) & (edf["_is_inbound"])].copy()
    if subset.empty:
        return pd.DataFrame({"note": ["No inbound client-case emails available"]})

    subset["_intent_label"] = subset.apply(
        lambda r: infer_intent_label(r["_text_for_nlp"], r["_subject_final"]), axis=1
    )
    summary = (
        subset.groupby(["_subject_final", "_intent_label"])
        .agg(
            email_count=("_email_id", "size"),
            distinct_cases=("_case_number", pd.Series.nunique),
            median_body_chars=("_body_len", "median"),
            median_case_hours=("_hours", "median"),
        )
        .reset_index()
        .sort_values(["_subject_final", "email_count"], ascending=[True, False])
    )
    summary["median_body_chars"] = summary["median_body_chars"].round(0)
    summary["median_case_hours"] = summary["median_case_hours"].round(1)

    top_subjects = subset["_subject_final"].value_counts().head(8).index.tolist()
    summary = summary[summary["_subject_final"].isin(top_subjects)].copy()

    phrase_rows = []
    for subj in top_subjects:
        phrases = top_phrases(subset.loc[subset["_subject_final"] == subj, "_text_for_nlp"], top_n=8)
        for _, r in phrases.iterrows():
            phrase_rows.append({
                "_subject_final": subj,
                "_intent_label": "(top phrase)",
                "email_count": r["count"],
                "distinct_cases": "",
                "median_body_chars": "",
                "median_case_hours": "",
                "phrase": r["phrase"],
            })

    summary["phrase"] = ""
    if phrase_rows:
        summary = pd.concat([summary, pd.DataFrame(phrase_rows)], ignore_index=True, sort=False)
    summary.rename(columns={"_subject_final": "subject", "_intent_label": "intent_label"}, inplace=True)
    return summary


def build_thread_features(edf):
    subset = edf[(edf["_is_client_case"]) & (edf["_has_case"])].copy()
    subset = subset.sort_values(["_case_number", "_created_dt", "_email_id"])
    if subset.empty:
        return pd.DataFrame()

    parts = []
    for case_num, grp in subset.groupby("_case_number"):
        grp = grp.sort_values(["_created_dt", "_email_id"])
        inbound = int(grp["_is_inbound"].sum())
        outbound = int(grp["_is_outbound"].sum())
        directions = grp["_direction"].fillna("Unknown").tolist()
        alternations = sum(1 for i in range(1, len(directions)) if directions[i] != directions[i - 1])
        actors = set()
        for _, row in grp.iterrows():
            actors.add(str(row["_from"]).strip().lower())
            actors.add(str(row["_to"]).strip().lower())
        actors.discard("")
        total_chars = int(grp["_body_len"].sum())
        quoted_depth = int(grp["_quoted_depth_proxy"].sum())
        msg_count = len(grp)
        complexity = (
            min(msg_count, 10) * 1.5
            + min(len(actors), 8) * 1.0
            + min(alternations, 10) * 1.2
            + min(total_chars / 2000.0, 12) * 1.0
            + min(quoted_depth, 10) * 0.7
        )
        parts.append({
            "_case_number": case_num,
            "subject": grp["_subject_final"].iloc[0],
            "company": grp["_company_clean"].iloc[0],
            "owner": grp["_owner"].iloc[0],
            "pod": grp["_pod"].iloc[0],
            "email_count": msg_count,
            "inbound_count": inbound,
            "outbound_count": outbound,
            "actor_count": len(actors),
            "direction_switches": alternations,
            "quoted_depth_proxy": quoted_depth,
            "total_body_chars": total_chars,
            "max_body_chars": int(grp["_body_len"].max()) if len(grp) else 0,
            "thread_complexity_score": round(complexity, 1),
            "case_hours": round(float(grp["_hours"].dropna().median()), 1) if grp["_hours"].notna().any() else np.nan,
            "is_resolved": bool(grp["_is_resolved"].dropna().iloc[0]) if grp["_is_resolved"].notna().any() else False,
        })
    return pd.DataFrame(parts).sort_values("thread_complexity_score", ascending=False)


def sheet_03_conversation_loops(thread_df):
    if thread_df.empty:
        return pd.DataFrame({"note": ["No thread features available"]})
    rows = []
    rows.append({"section": "SUMMARY", "metric": "cases_with_email_threads", "value": f"{len(thread_df):,}"})
    rows.append({"section": "SUMMARY", "metric": "median_emails_per_case", "value": f"{thread_df['email_count'].median():.1f}"})
    rows.append({"section": "SUMMARY", "metric": "p90_emails_per_case", "value": f"{thread_df['email_count'].quantile(.9):.1f}"})
    rows.append({"section": "SUMMARY", "metric": "median_complexity_score", "value": f"{thread_df['thread_complexity_score'].median():.1f}"})
    high_loop = (thread_df["email_count"] >= 5).sum()
    rows.append({"section": "SUMMARY", "metric": "high_loop_cases_5plus_emails", "value": f"{high_loop:,} ({pct(high_loop, len(thread_df))})"})

    top_subject = (
        thread_df.groupby("subject")
        .agg(
            cases=("subject", "size"),
            median_emails=("email_count", "median"),
            p90_emails=("email_count", lambda s: s.quantile(.9)),
            median_case_hours=("case_hours", "median"),
        )
        .reset_index()
        .sort_values(["median_emails", "cases"], ascending=[False, False])
        .head(12)
    )
    top_subject.insert(0, "section", "BY SUBJECT")
    top_subject.rename(columns={"subject": "metric"}, inplace=True)
    for c in ["median_emails", "p90_emails", "median_case_hours"]:
        top_subject[c] = top_subject[c].round(1)
    return pd.concat([pd.DataFrame(rows), top_subject], ignore_index=True, sort=False)


def template_signature(text):
    s = canonical_text(text)
    s = regex_re.sub(r"\b(?:dear|hello|hi)\b\s+\w+", " GREETING ", s)
    s = regex_re.sub(r"\b(?:thank|thanks)\b.*", " CLOSING ", s)
    toks = [t for t in s.split() if t not in STOPWORDS]
    return " ".join(toks[:50])


def cluster_repetitive_messages(series):
    clean = [template_signature(x) for x in series if x]
    clean = [x for x in clean if len(x) >= 30]
    if not clean:
        return []

    exact = Counter(clean)
    clusters = [{"signature": sig, "count": count, "method": "exact"} for sig, count in exact.most_common(12) if count >= 2]
    if fuzz is None or len(clean) > 800:
        return clusters

    seen = set()
    for i in range(len(clean)):
        if i in seen:
            continue
        group = [i]
        for j in range(i + 1, len(clean)):
            if j in seen or abs(len(clean[i]) - len(clean[j])) > 80:
                continue
            if fuzz.token_set_ratio(clean[i], clean[j]) >= 92:
                group.append(j)
        if len(group) >= 3:
            seen.update(group)
            clusters.append({"signature": clean[i][:120], "count": len(group), "method": "fuzzy"})
    return sorted(clusters, key=lambda x: x["count"], reverse=True)[:15]


def sheet_04_repetitive_outbound(edf):
    subset = edf[(edf["_is_client_case"]) & (edf["_is_outbound"])].copy()
    if subset.empty:
        return pd.DataFrame({"note": ["No outbound client-case emails available"]})

    rows = []
    for subj in subset["_subject_final"].value_counts().head(8).index.tolist():
        sub = subset[subset["_subject_final"] == subj]
        clusters = cluster_repetitive_messages(sub["_body_text"])
        if not clusters:
            rows.append({
                "subject": subj,
                "method": "none",
                "signature": "(no repeated outbound template detected)",
                "count": 0,
                "subject_email_total": len(sub),
            })
            continue
        for cl in clusters:
            rows.append({
                "subject": subj,
                "method": cl["method"],
                "signature": cl["signature"],
                "count": cl["count"],
                "subject_email_total": len(sub),
            })
    return pd.DataFrame(rows).sort_values(["subject", "count"], ascending=[True, False])


def detect_doc_terms(text):
    t = canonical_text(text)
    hits = [term for term in DOC_TERMS if term in t]
    return ", ".join(sorted(hits[:6]))


def detect_missing_signal(text):
    t = canonical_text(text)
    return any(cue in t for cue in MISSING_CUES)


def sheet_05_missing_info(edf):
    subset = edf[(edf["_is_client_case"]) & (edf["_is_inbound"])].copy()
    if subset.empty:
        return pd.DataFrame({"note": ["No inbound client-case emails available"]})

    subset["missing_signal"] = subset["_text_for_nlp"].apply(detect_missing_signal)
    subset["doc_terms"] = subset["_text_for_nlp"].apply(detect_doc_terms)
    subset["intent_label"] = subset.apply(
        lambda r: infer_intent_label(r["_text_for_nlp"], r["_subject_final"]), axis=1
    )

    summary = (
        subset.groupby(["_subject_final", "intent_label"])
        .agg(
            email_count=("_email_id", "size"),
            missing_signal_count=("missing_signal", "sum"),
            cases=("_case_number", pd.Series.nunique),
            median_case_hours=("_hours", "median"),
        )
        .reset_index()
    )
    summary["missing_signal_pct"] = (100 * summary["missing_signal_count"] / summary["email_count"]).round(1)
    summary["median_case_hours"] = summary["median_case_hours"].round(1)
    summary = summary.sort_values(["missing_signal_pct", "email_count"], ascending=[False, False]).head(25)
    summary.rename(columns={"_subject_final": "subject"}, inplace=True)

    examples = subset[subset["missing_signal"]].copy().head(12)
    if not examples.empty:
        examples = examples[["_subject_final", "intent_label", "doc_terms", "_subject_clean", "_body_text"]].copy()
        examples.rename(columns={
            "_subject_final": "subject",
            "_subject_clean": "email_subject",
            "_body_text": "example_text"
        }, inplace=True)
        examples["email_count"] = ""
        examples["missing_signal_count"] = ""
        examples["cases"] = ""
        examples["median_case_hours"] = ""
        examples["missing_signal_pct"] = ""
        summary["doc_terms"] = ""
        summary["email_subject"] = ""
        summary["example_text"] = ""
        summary = pd.concat([summary, examples], ignore_index=True, sort=False)
    return summary


def sheet_06_context_burden(thread_df):
    if thread_df.empty:
        return pd.DataFrame({"note": ["No thread features available"]})
    df = thread_df.copy()
    slow_cut = df["case_hours"].dropna().quantile(.75) if df["case_hours"].notna().any() else np.inf
    df["slow_case_flag"] = df["case_hours"].fillna(0).ge(slow_cut)
    df["high_context_flag"] = df["thread_complexity_score"].ge(df["thread_complexity_score"].quantile(.75))

    rows = [
        {"section": "SUMMARY", "metric": "high_context_cases", "value": f"{df['high_context_flag'].sum():,} ({pct(df['high_context_flag'].sum(), len(df))})"},
        {"section": "SUMMARY", "metric": "high_context_and_slow", "value": f"{(df['high_context_flag'] & df['slow_case_flag']).sum():,}"},
        {"section": "SUMMARY", "metric": "median_case_hours_high_context", "value": f"{df.loc[df['high_context_flag'], 'case_hours'].median():.1f}" if df.loc[df['high_context_flag'], 'case_hours'].notna().any() else "N/A"},
        {"section": "SUMMARY", "metric": "median_case_hours_low_context", "value": f"{df.loc[~df['high_context_flag'], 'case_hours'].median():.1f}" if df.loc[~df['high_context_flag'], 'case_hours'].notna().any() else "N/A"},
    ]
    top = df.sort_values("thread_complexity_score", ascending=False).head(20).copy()
    top.insert(0, "section", "TOP CASES")
    top.rename(columns={"_case_number": "metric"}, inplace=True)
    return pd.concat([pd.DataFrame(rows), top], ignore_index=True, sort=False)


def simple_similarity(a, b):
    ta = set(tokens(a))
    tb = set(tokens(b))
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / len(ta | tb)


def compute_similarity_pairs(texts):
    n = len(texts)
    if n == 0:
        return []
    if TfidfVectorizer is not None and linear_kernel is not None:
        vec = TfidfVectorizer(max_features=3000, ngram_range=(1, 2), min_df=2, stop_words="english")
        try:
            mat = vec.fit_transform(texts)
            sims = linear_kernel(mat, mat)
            pairs = []
            for i in range(n):
                top_idx = np.argsort(sims[i])[::-1][1:6]
                for j in top_idx:
                    score = float(sims[i][j])
                    if score >= 0.45:
                        pairs.append((i, int(j), round(score, 3), "tfidf"))
            return pairs
        except Exception as exc:
            warn(f"TF-IDF similarity failed; falling back to token overlap: {exc}")
    pairs = []
    for i in range(n):
        for j in range(i + 1, min(i + 25, n)):
            score = simple_similarity(texts[i], texts[j])
            if score >= 0.5:
                pairs.append((i, j, round(score, 3), "jaccard"))
    return pairs


def sheet_07_routing_ambiguity(edf):
    subset = edf[(edf["_is_client_case"]) & (edf["_is_inbound"])].copy()
    subset = subset[subset["_text_for_nlp"].str.len() >= 40].copy()
    if len(subset) < 15:
        return pd.DataFrame({"note": ["Not enough inbound emails for routing ambiguity analysis"]})

    subset = subset.head(600).reset_index(drop=True)
    pairs = compute_similarity_pairs(subset["_text_for_nlp"].tolist())
    rows = []
    seen = set()
    for i, j, score, method in pairs:
        a = subset.iloc[i]
        b = subset.iloc[j]
        if a["_subject_final"] == b["_subject_final"]:
            continue
        key = tuple(sorted((int(a["_email_id"]), int(b["_email_id"]))))
        if key in seen:
            continue
        seen.add(key)
        rows.append({
            "similarity_method": method,
            "similarity_score": score,
            "subject_a": a["_subject_final"],
            "subject_b": b["_subject_final"],
            "owner_a": a["_owner"],
            "owner_b": b["_owner"],
            "pod_a": a["_pod"],
            "pod_b": b["_pod"],
            "email_subject_a": trunc(a["_subject_clean"]),
            "email_subject_b": trunc(b["_subject_clean"]),
            "sample_text_a": trunc(a["_body_text"], 220),
            "sample_text_b": trunc(b["_body_text"], 220),
        })
    if not rows:
        return pd.DataFrame({"note": ["No high-similarity cross-subject pairs detected"]})
    return pd.DataFrame(rows).sort_values("similarity_score", ascending=False).head(50)


def detect_followup(text):
    t = canonical_text(text)
    return any(cue in t for cue in FOLLOWUP_CUES)


def detect_urgency(text):
    t = canonical_text(text)
    return any(cue in t for cue in URGENCY_CUES)


def sheet_08_urgency_signals(edf, thread_df):
    subset = edf[(edf["_is_client_case"]) & (edf["_is_inbound"])].copy()
    if subset.empty:
        return pd.DataFrame({"note": ["No inbound client-case emails available"]})

    subset["followup_flag"] = subset["_text_for_nlp"].apply(detect_followup)
    subset["urgency_flag"] = subset["_text_for_nlp"].apply(detect_urgency)
    merged = subset.merge(thread_df[["_case_number", "email_count", "thread_complexity_score", "case_hours"]], on="_case_number", how="left")

    rows = []
    for label, flag_col in [("followup", "followup_flag"), ("urgency", "urgency_flag")]:
        flagged = merged[merged[flag_col]]
        rows.append({"section": "SUMMARY", "metric": f"{label}_emails", "value": f"{len(flagged):,} ({pct(len(flagged), len(merged))})"})
        rows.append({"section": "SUMMARY", "metric": f"{label}_median_case_hours", "value": f"{flagged['case_hours'].median():.1f}" if flagged["case_hours"].notna().any() else "N/A"})
        rows.append({"section": "SUMMARY", "metric": f"{label}_median_emails_per_case", "value": f"{flagged['email_count'].median():.1f}" if flagged["email_count"].notna().any() else "N/A"})

    top = merged[(merged["followup_flag"]) | (merged["urgency_flag"])].copy().head(25)
    if not top.empty:
        top.insert(0, "section", "EXAMPLES")
        top.rename(columns={"_subject_final": "subject", "_subject_clean": "email_subject", "_body_text": "sample_text"}, inplace=True)
        top = top[["section", "subject", "followup_flag", "urgency_flag", "email_count", "thread_complexity_score", "case_hours", "email_subject", "sample_text"]]
        top["sample_text"] = top["sample_text"].apply(lambda x: trunc(x, 220))
        return pd.concat([pd.DataFrame(rows), top], ignore_index=True, sort=False)
    return pd.DataFrame(rows)


def sheet_09_pmc_stress(thread_df):
    if thread_df.empty:
        return pd.DataFrame({"note": ["No thread features available"]})
    grp = (
        thread_df.groupby("company")
        .agg(
            cases=("company", "size"),
            total_emails=("email_count", "sum"),
            median_emails_per_case=("email_count", "median"),
            median_complexity=("thread_complexity_score", "median"),
            unresolved_cases=("is_resolved", lambda s: int((~s).sum())),
            median_case_hours=("case_hours", "median"),
        )
        .reset_index()
        .sort_values(["total_emails", "median_complexity"], ascending=[False, False])
    )
    grp["emails_per_case"] = (grp["total_emails"] / grp["cases"]).round(1)
    grp["unresolved_pct"] = (100 * grp["unresolved_cases"] / grp["cases"]).round(1)
    return grp.head(25)


def sheet_10_temporal_demand(edf):
    subset = edf[(edf["_is_client_case"]) & (edf["_is_inbound"])].copy()
    if subset.empty or subset["_created_dt"].isna().all():
        return pd.DataFrame({"note": ["No timestamped inbound client-case emails available"]})
    grp = (
        subset.groupby(["_hour", "_subject_final"])
        .size()
        .reset_index(name="email_count")
        .sort_values(["email_count"], ascending=False)
        .head(40)
    )
    grp["_hour"] = grp["_hour"].fillna(-1).astype(int).astype(str) + ":00"
    grp.rename(columns={"_hour": "hour", "_subject_final": "subject"}, inplace=True)
    return grp


def sheet_11_insight_scorecard(edf, thread_df, routing_df, outbound_df, missing_df, urgency_df):
    inbound = edf[(edf["_is_client_case"]) & (edf["_is_inbound"])]
    hidden_taxonomy_signal = inbound["_subject_final"].value_counts().head(5).sum()
    repetitive_templates = 0
    if isinstance(outbound_df, pd.DataFrame) and not outbound_df.empty and "count" in outbound_df.columns:
        repetitive_templates = int(outbound_df["count"].fillna(0).sum())
    missing_signal_pct = "N/A"
    if isinstance(missing_df, pd.DataFrame) and not missing_df.empty and "missing_signal_pct" in missing_df.columns:
        vals = pd.to_numeric(missing_df["missing_signal_pct"], errors="coerce").dropna()
        if len(vals):
            missing_signal_pct = f"{vals.max():.1f}%"
    routing_pairs = 0 if routing_df is None or routing_df.empty or "similarity_score" not in routing_df.columns else len(routing_df)
    high_context_cases = 0 if thread_df.empty else int((thread_df["thread_complexity_score"] >= thread_df["thread_complexity_score"].quantile(.75)).sum())
    followup_examples = 0 if urgency_df is None or urgency_df.empty else int(urgency_df.astype(str).apply(lambda s: s.str.contains("True")).any(axis=1).sum())
    rows = [
        {"insight": "Hidden intent taxonomy", "evidence_metric": f"Top 5 broad subjects cover {pct(hidden_taxonomy_signal, len(inbound))} of inbound client-case emails", "why_it_matters": "Tests whether broad case labels hide repeated text-work families", "ai_genai_signal": "Strong if repeated subtypes are concentrated"},
        {"insight": "Conversation-loop intensity", "evidence_metric": f"{high_context_cases:,} high-context cases in current extract", "why_it_matters": "Tests whether reading and reconstructing threads is a real burden", "ai_genai_signal": "Strong for summarization if high-context cases also run slow"},
        {"insight": "Repetitive outbound language", "evidence_metric": f"{repetitive_templates:,} repeated outbound-template hits detected", "why_it_matters": "Tests whether banker drafting is repetitive enough for GenAI support", "ai_genai_signal": "Strong only if repetition is material"},
        {"insight": "Missing-information burden", "evidence_metric": f"Highest observed missing-signal share in summary table = {missing_signal_pct}", "why_it_matters": "Tests whether delay is driven by incomplete inbound requests", "ai_genai_signal": "Strong if missing patterns recur in specific request families"},
        {"insight": "Routing ambiguity", "evidence_metric": f"{routing_pairs:,} high-similarity cross-subject email pairs", "why_it_matters": "Tests whether similar text gets routed inconsistently", "ai_genai_signal": "Strong for classification if semantic ambiguity is real"},
        {"insight": "Urgency and follow-up cues", "evidence_metric": f"{followup_examples:,} rows in urgency and follow-up output", "why_it_matters": "Tests whether service-risk signals live in language rather than fields", "ai_genai_signal": "Medium until linked to outcomes"},
        {"insight": "PMC communication stress", "evidence_metric": "Exploratory only with one-day extract", "why_it_matters": "Useful for relationship risk once longer history exists", "ai_genai_signal": "Not a Phase 1 proof point from one day"},
    ]
    return pd.DataFrame(rows)


def write_markdown(sheets):
    md = [
        "# WAB Email Deep Insights",
        f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
        "## Purpose",
        "This run expands beyond first-round email profiling to look for deeper signals that can justify AI/GenAI involvement.",
        "",
        "## Outputs",
    ]
    for name, sdf in sheets.items():
        md.append(f"- **{name}**: {len(sdf) if sdf is not None else 0} rows")
    if WARN:
        md.extend(["", "## Warnings"])
        for w in WARN:
            md.append(f"- {w}")
    md.extend(["", "## Log", "```"])
    md.extend(LOG)
    md.append("```")
    with open(OUTPUT_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(md))


def main():
    start = datetime.datetime.now()
    log(f"WAB Email Deep Insights — {start.strftime('%Y-%m-%d %H:%M:%S')}")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    cases_raw = read_file(CASE_FILE, "Cases")
    emails_raw = read_file(EMAIL_FILE, "Emails")
    if cases_raw.empty or emails_raw.empty:
        log("FATAL: required files not loaded. Exiting.")
        return

    log("\n--- Preparing frames ---")
    cases = prepare_cases(cases_raw)
    edf = prepare_emails(emails_raw, cases)
    log(f"  Client-case emails: {edf['_is_client_case'].sum():,}")

    log("\n--- Building thread features ---")
    thread_df = build_thread_features(edf)
    log(f"  Thread rows: {len(thread_df):,}")

    log("\n--- Building insight sheets ---")
    sheets = OrderedDict()
    sheets["I01_DataScope"] = sheet_01_data_scope(edf)
    sheets["I02_HiddenIntents"] = sheet_02_hidden_intents(edf)
    sheets["I03_ConversationLoops"] = sheet_03_conversation_loops(thread_df)
    sheets["I04_RepetitiveOutbound"] = sheet_04_repetitive_outbound(edf)
    sheets["I05_MissingInfo"] = sheet_05_missing_info(edf)
    sheets["I06_ContextBurden"] = sheet_06_context_burden(thread_df)
    routing_df = sheet_07_routing_ambiguity(edf)
    sheets["I07_RoutingAmbiguity"] = routing_df
    urgency_df = sheet_08_urgency_signals(edf, thread_df)
    sheets["I08_UrgencySignals"] = urgency_df
    sheets["I09_PMCStress"] = sheet_09_pmc_stress(thread_df)
    sheets["I10_TemporalDemand"] = sheet_10_temporal_demand(edf)
    sheets["I11_InsightScorecard"] = sheet_11_insight_scorecard(edf, thread_df, routing_df, sheets["I04_RepetitiveOutbound"], sheets["I05_MissingInfo"], urgency_df)

    log(f"\nWriting: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            write_sheet(writer, name, sdf)
    log("  Excel done.")

    log(f"Writing: {OUTPUT_MD}")
    write_markdown(sheets)
    elapsed = (datetime.datetime.now() - start).total_seconds()
    log(f"\nCompleted in {elapsed:.1f}s")


if __name__ == "__main__":
    main()
