"""
WAB Email Deep Insights — Assumption Validator
================================================
Standalone script. Run AFTER wab_email_deep_insights.py.
Reads the same 2 input files and produces one small Excel workbook
with spot-checks for assumptions A19, A21, A22, A24, A25.

Each sheet is designed to fit in 1-2 screenshots.
"""

CASE_FILE  = r"C:\Users\YourName\Desktop\AAB All Cases.xlsx"
EMAIL_FILE = r"C:\Users\YourName\Desktop\ALL EMAIL Files.xlsx"
OUTPUT_DIR = r"C:\Users\YourName\Desktop\wab_output"

import os, re, html, warnings
import pandas as pd
import numpy as np
from bs4 import BeautifulSoup

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "wab_email_validation.xlsx")


# ── Reuse the same preprocessing functions from the main module ──

def strip_html(text):
    if pd.isna(text): return ""
    s = str(text)
    s = re.sub(r"<(style|script)[^>]*>.*?</\1>", " ", s, flags=re.I | re.S)
    soup = BeautifulSoup(s, "html.parser")
    text = soup.get_text("\n")
    text = html.unescape(text)
    text = re.sub(r"\r", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

def remove_noise(text):
    if pd.isna(text) or not text: return ""
    s = str(text)

    # FIX A19: WAB security banner is a 3-line block:
    #   ATTENTION: This email originated from outside of the WAB Network.
    #   DO NOT click on any links or download attachments from unknown
    #   senders!!!
    # Followed by a horizontal rule. Match the full block as one unit.
    s = re.sub(
        r"(?i)ATTENTION:\s*This email originated from outside of the WAB Network\.?\s*"
        r"DO NOT click on any links or download attachments from unknown\s*"
        r"senders\s*!{0,5}",
        " ", s
    )
    # Also catch slight variations (different wording, missing parts)
    s = re.sub(r"(?i)ATTENTION:\s*This email originated from outside[^\n]*\n?[^\n]*DO NOT click[^\n]*\n?[^\n]*senders[^\n]*", " ", s)
    # Catch standalone fragments if the banner was partially stripped by HTML parsing
    s = re.sub(r"(?i)DO NOT click on any links or download attachments from unknown\s*senders\s*!{0,5}", " ", s)
    # Other security banner variants
    s = re.sub(r"(?i)external email warning[^\n]*", " ", s)
    s = re.sub(r"(?i)caution:\s*external[^\n]*", " ", s)

    # Forwarded headers — match ONE header block per occurrence, not across multiple blocks.
    # Use non-DOTALL so .*? stops at newlines within each field.
    s = re.sub(r"(?i)^from:\s+[^\n]+\n\s*sent:\s+[^\n]+\n\s*to:\s+[^\n]+\n\s*(?:cc:\s+[^\n]+\n\s*)?subject:\s+[^\n]*", " ", s, flags=re.MULTILINE)
    s = re.sub(r"(?i)on .{10,60} wrote:\s*", " ", s)

    # Separator lines
    s = re.sub(r"_{5,}", " ", s)
    s = re.sub(r"-{5,}", " ", s)

    # Signature blocks — only from last 30% of text
    cutpoint = max(len(s) * 7 // 10, 200)
    head = s[:cutpoint]
    tail = s[cutpoint:]
    tail = re.sub(r"(?i)\b(?:best regards|regards|sincerely|thanks|thank you)\s*,?\s*\n.*", "", tail, flags=re.S)
    s = head + tail

    s = re.sub(r"\s+", " ", s).strip()
    return s

def extract_new_content(raw_text):
    if not raw_text: return "", "", 0.0
    lines = raw_text.split("\n")
    new_lines, quoted_lines = [], []
    in_quoted = False
    for line in lines:
        stripped = line.strip()
        if re.match(r"^>", stripped) or re.match(r"(?i)^(from|sent|to|subject|on .* wrote):", stripped):
            in_quoted = True
        if re.match(r"^[-_=]{4,}$", stripped):
            in_quoted = True
        (quoted_lines if in_quoted else new_lines).append(line)
    new_text = "\n".join(new_lines).strip()
    quoted_text = "\n".join(quoted_lines).strip()
    total = len(new_text) + len(quoted_text)
    return new_text, quoted_text, len(new_text) / total if total > 0 else 0.0

def canonical(text):
    if not text: return ""
    s = text.lower()
    s = re.sub(r"\b[\w\.-]+@[\w\.-]+\.\w+\b", " _EMAIL_ ", s)
    s = re.sub(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", " _DATE_ ", s)
    s = re.sub(r"\$[\d,]+\.?\d*", " _AMOUNT_ ", s)
    s = re.sub(r"\b\d{4,}\b", " _NUM_ ", s)
    s = re.sub(r"[^a-z0-9_\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def find_col(df, *candidates):
    lookup = {re.sub(r"\s+", " ", c.strip().lower()): c for c in df.columns}
    for cand in candidates:
        n = re.sub(r"\s+", " ", cand.strip().lower())
        if n in lookup: return lookup[n]
    for cand in candidates:
        n = re.sub(r"\s+", " ", cand.strip().lower())
        for k, v in lookup.items():
            if n in k or k in n: return v
    return None

MISSING_CUES = [
    # FIX A22: Removed "required" (triggers on "Action Required" / "Response Required" headers)
    # FIX A22: Removed "need the" (too generic — triggers on normal conversational English)
    "missing", "not received", "still need", "need your",
    "please provide", "please send", "awaiting", "incomplete",
    "once we receive", "can you send", "have not received", "pending receipt",
    "not yet received", "still waiting for", "in order to proceed",
    "unable to process", "cannot proceed without",
]
FOLLOWUP_CUES = [
    "following up", "follow up", "just following up", "checking in",
    "circling back", "still waiting", "any update", "please advise",
    "status update", "when can we expect", "have you had a chance",
    "wanted to check", "reaching out again", "second request",
]
URGENCY_CUES = [
    # FIX A22: Removed "today" (triggers on date references like "today's transactions")
    "urgent", "asap", "immediately", "end of day", "eod",
    "rush", "critical", "time sensitive", "deadline", "priority",
    "expedite", "right away", "as soon as possible",
]


def remove_noise_OLD(text):
    """Original version with the aggressive banner regex — kept for V1 comparison only."""
    if pd.isna(text) or not text: return ""
    s = str(text)
    s = re.sub(r"(?i)attention:\s*this email originated from outside.*?(?=\n\n|\Z)", " ", s, flags=re.S)
    s = re.sub(r"(?i)external email warning.*?(?=\n\n|\Z)", " ", s, flags=re.S)
    s = re.sub(r"(?i)caution:\s*external.*?(?=\n\n|\Z)", " ", s, flags=re.S)
    s = re.sub(r"(?i)do not click links.*?(?=\n\n|\Z)", " ", s, flags=re.S)
    s = re.sub(r"(?i)from:\s.*?sent:\s.*?to:\s.*?subject:\s[^\n]*", " ", s, flags=re.S)
    s = re.sub(r"(?i)on .{10,60} wrote:\s*", " ", s)
    s = re.sub(r"_{5,}", " ", s)
    s = re.sub(r"-{5,}", " ", s)
    cutpoint = max(len(s) * 7 // 10, 200)
    head = s[:cutpoint]
    tail = s[cutpoint:]
    tail = re.sub(r"(?i)\b(?:best regards|regards|sincerely|thanks|thank you)\s*,?\s*\n.*", "", tail, flags=re.S)
    s = head + tail
    s = re.sub(r"\s+", " ", s).strip()
    return s


def write_sheet(writer, name, df):
    if df is None or df.empty:
        df = pd.DataFrame({"note": ["No data"]})
    df.to_excel(writer, sheet_name=name[:31], index=False, freeze_panes=(1, 0))
    ws = writer.sheets[name[:31]]
    for col_cells in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 55)


def trunc(v, n=200):
    if pd.isna(v): return ""
    s = str(v).replace("\r", " ").replace("\n", " ").strip()
    return s[:n] + "..." if len(s) > n else s


def main():
    print("Reading files...")
    cases = pd.read_excel(CASE_FILE, header=0, engine="openpyxl")
    emails = pd.read_excel(EMAIL_FILE, header=0, engine="openpyxl")
    cases = cases.loc[:, ~cases.columns.astype(str).str.match(r"^Unnamed")].dropna(axis=1, how="all")
    emails = emails.loc[:, ~emails.columns.astype(str).str.match(r"^Unnamed")].dropna(axis=1, how="all")

    desc_col = find_col(emails, "Description")
    subj_col = find_col(emails, "Subject")
    status_col = find_col(emails, "Status Reason")
    case_subj_col = find_col(emails, "Subject (Regarding) (Case)", "Subject Path (Regarding) (Case)")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ════════════════════════════════════════════════════════
    # V1: A19 — HTML stripping quality
    # Show 10 emails: raw HTML (first 300 chars), stripped (first 300),
    # cleaned (first 300). Human can eyeball whether meaning is preserved.
    # ════════════════════════════════════════════════════════
    print("V1: HTML stripping spot-check...")
    sample_idx = emails.sample(min(10, len(emails)), random_state=42).index
    v1_rows = []
    for idx in sample_idx:
        raw_val = emails.loc[idx, desc_col] if desc_col else ""
        raw_html = "" if pd.isna(raw_val) else str(raw_val)
        stripped = strip_html(raw_html)
        cleaned_OLD = remove_noise_OLD(stripped)
        cleaned_NEW = remove_noise(stripped)
        subj_val = emails.loc[idx, subj_col] if subj_col else ""
        v1_rows.append({
            "email_row": int(idx) + 2,
            "subject": trunc(subj_val, 80) if not pd.isna(subj_val) else "",
            "stripped_chars": len(stripped),
            "stripped_first200": trunc(stripped, 200),
            "OLD_cleaned_chars": len(cleaned_OLD),
            "OLD_cleaned_first200": trunc(cleaned_OLD, 200),
            "NEW_cleaned_chars": len(cleaned_NEW),
            "NEW_cleaned_first200": trunc(cleaned_NEW, 200),
            "chars_recovered": len(cleaned_NEW) - len(cleaned_OLD),
        })
    v1 = pd.DataFrame(v1_rows)

    # ════════════════════════════════════════════════════════
    # V2: A21 — New-vs-quoted separation accuracy
    # Show 10 emails: full stripped text, what was classified as "new",
    # what was classified as "quoted", and the ratio.
    # Human can judge: did the split make sense?
    # ════════════════════════════════════════════════════════
    print("V2: New vs quoted separation spot-check...")
    v2_rows = []
    for idx in sample_idx:
        raw_val2 = emails.loc[idx, desc_col] if desc_col else ""
        raw_html2 = "" if pd.isna(raw_val2) else str(raw_val2)
        stripped2 = strip_html(raw_html2)
        new_text, quoted_text, new_ratio = extract_new_content(stripped2)
        subj_val2 = emails.loc[idx, subj_col] if subj_col else ""
        v2_rows.append({
            "email_row": int(idx) + 2,
            "subject": trunc(subj_val2, 80) if not pd.isna(subj_val2) else "",
            "total_chars": len(stripped),
            "new_chars": len(new_text),
            "quoted_chars": len(quoted_text),
            "new_ratio": f"{new_ratio:.0%}",
            "new_text_first300": trunc(new_text, 300),
            "quoted_text_first300": trunc(quoted_text, 300),
        })
    v2 = pd.DataFrame(v2_rows)

    # ════════════════════════════════════════════════════════
    # V3: A22 — Signal detection keyword accuracy
    # For each signal type, show 5 TRUE and 5 FALSE examples with
    # the actual text, so human can judge if the keyword matched correctly.
    # ════════════════════════════════════════════════════════
    print("V3: Signal detection spot-check...")

    # Preprocess all emails for signal detection
    emails["_stripped"] = emails[desc_col].fillna("").apply(strip_html) if desc_col else ""
    emails["_cleaned"] = emails["_stripped"].fillna("").apply(remove_noise)
    emails["_subj_clean"] = emails[subj_col].fillna("").apply(remove_noise) if subj_col else ""
    emails["_text"] = (emails["_subj_clean"] + " " + emails["_cleaned"].str[:1200]).str.strip()
    emails["_text_lower"] = emails["_text"].str.lower()

    v3_rows = []
    for signal_name, cue_list in [("missing_info", MISSING_CUES),
                                    ("followup", FOLLOWUP_CUES),
                                    ("urgency", URGENCY_CUES)]:
        emails[f"_{signal_name}"] = emails["_text_lower"].apply(
            lambda t: any(cue in t for cue in cue_list) if isinstance(t, str) else False)

        # Find which specific cue matched
        def find_matched_cue(text):
            if not isinstance(text, str): return ""
            for cue in cue_list:
                if cue in text: return cue
            return ""

        flagged = emails[emails[f"_{signal_name}"]]
        unflagged = emails[~emails[f"_{signal_name}"]]

        for _, row in flagged.head(5).iterrows():
            matched_cue = find_matched_cue(row["_text_lower"])
            v3_rows.append({
                "signal": signal_name,
                "detected": "TRUE",
                "matched_cue": matched_cue,
                "email_subject": trunc(row.get(subj_col, ""), 80),
                "case_subject": trunc(row.get(case_subj_col, ""), 40) if case_subj_col else "",
                "text_around_cue": trunc(row["_text"], 250),
            })
        for _, row in unflagged.head(3).iterrows():
            v3_rows.append({
                "signal": signal_name,
                "detected": "FALSE",
                "matched_cue": "",
                "email_subject": trunc(row.get(subj_col, ""), 80),
                "case_subject": trunc(row.get(case_subj_col, ""), 40) if case_subj_col else "",
                "text_around_cue": trunc(row["_text"], 250),
            })
    v3 = pd.DataFrame(v3_rows)

    # ════════════════════════════════════════════════════════
    # V4: A24 — Template clustering sensitivity
    # Pick the highest-volume outbound subject. Run clustering at
    # 3 thresholds (0.5, 0.65, 0.8) and compare coverage.
    # ════════════════════════════════════════════════════════
    print("V4: Template clustering sensitivity...")
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity

    dir_col = status_col
    if dir_col:
        emails["_direction"] = np.where(
            emails[dir_col].fillna("").str.lower().str.contains("sent|completed|outgoing"),
            "Outbound",
            np.where(emails[dir_col].fillna("").str.lower().str.contains("received|incoming"),
                     "Inbound", "Other"))
    else:
        emails["_direction"] = "Other"

    case_subj = find_col(emails, "Subject (Regarding) (Case)")
    outbound = emails[emails["_direction"] == "Outbound"].copy()
    if case_subj:
        top_subj = outbound[case_subj].value_counts().head(1).index[0] if len(outbound) > 0 else None
    else:
        top_subj = None

    v4_rows = []
    if top_subj and case_subj:
        sub = outbound[outbound[case_subj] == top_subj]
        texts = sub["_cleaned"].tolist()
        texts = [t for t in texts if isinstance(t, str) and len(t) >= 50]

        if len(texts) >= 10:
            tfidf = TfidfVectorizer(max_features=1000, ngram_range=(1, 2),
                                     min_df=2, max_df=0.9, stop_words="english")
            mat = tfidf.fit_transform(texts)
            sim_matrix = cosine_similarity(mat)

            for threshold in [0.50, 0.65, 0.80]:
                n = len(texts)
                assigned = [False] * n
                clusters = []
                for i in range(n):
                    if assigned[i]: continue
                    group = [i]
                    for j in range(i + 1, n):
                        if assigned[j]: continue
                        if sim_matrix[i][j] >= threshold:
                            group.append(j)
                            assigned[j] = True
                    assigned[i] = True
                    if len(group) >= 2:
                        clusters.append(group)

                total_clustered = sum(len(c) for c in clusters)
                coverage = 100 * total_clustered / len(texts) if texts else 0
                v4_rows.append({
                    "subject": top_subj,
                    "threshold": threshold,
                    "total_outbound": len(texts),
                    "clusters_found": len(clusters),
                    "largest_cluster": max(len(c) for c in clusters) if clusters else 0,
                    "total_clustered": total_clustered,
                    "coverage_pct": f"{coverage:.1f}%",
                })
        else:
            v4_rows.append({"subject": top_subj, "threshold": "N/A",
                           "total_outbound": len(texts), "note": "Too few texts for clustering"})
    else:
        v4_rows.append({"note": "No outbound subject found for sensitivity test"})
    v4 = pd.DataFrame(v4_rows)

    # ════════════════════════════════════════════════════════
    # V5: A25 — Direction classification accuracy
    # Show all distinct Status Reason values, our direction mapping,
    # and the count for each. Human can verify the mapping.
    # ════════════════════════════════════════════════════════
    print("V5: Direction classification check...")
    v5_rows = []
    if status_col:
        for val, count in emails[status_col].fillna("(blank)").value_counts().items():
            val_lower = str(val).lower()
            if any(k in val_lower for k in ("sent", "completed", "outgoing", "outbound")):
                mapped = "Outbound"
            elif any(k in val_lower for k in ("received", "incoming", "inbound")):
                mapped = "Inbound"
            else:
                mapped = "Other"
            v5_rows.append({
                "status_reason_value": val,
                "count": count,
                "mapped_direction": mapped,
                "review_correct": "",  # Human fills this in
            })
    else:
        v5_rows.append({"note": "Status Reason column not found"})
    v5 = pd.DataFrame(v5_rows)

    # ════════════════════════════════════════════════════════
    # Write
    # ════════════════════════════════════════════════════════
    print(f"Writing: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        write_sheet(writer, "V1_HTMLStrip", v1)
        write_sheet(writer, "V2_NewVsQuoted", v2)
        write_sheet(writer, "V3_SignalDetection", v3)
        write_sheet(writer, "V4_TemplateSensitivity", v4)
        write_sheet(writer, "V5_DirectionMapping", v5)

    print(f"Done. Open {OUTPUT_XLSX} and screenshot each sheet.")
    print()
    print("What to look for in each sheet:")
    print("  V1: Does stripped text preserve the meaning of the HTML? Is anything important lost?")
    print("  V2: Is the new/quoted split reasonable? Does 'new' contain actual new content?")
    print("  V3: Are TRUE signals real? Are FALSE signals correctly excluded? Check the matched_cue column.")
    print("  V4: Does coverage change dramatically at different thresholds? (0.5 vs 0.65 vs 0.8)")
    print("  V5: Is each Status Reason correctly mapped to Inbound/Outbound/Other?")


if __name__ == "__main__":
    main()
