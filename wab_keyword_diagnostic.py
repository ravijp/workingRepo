"""
WAB Keyword Diagnostic — What's Actually in the Text Fields
=============================================================
Run FIRST before defining sub-segmentation clusters.
Dumps top keywords from Description + Activity Subject for each major subject.
Also shows bigrams (2-word phrases) which are more informative than single words.

Output: one Excel workbook with one sheet per subject.
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
from collections import Counter, OrderedDict

import pandas as pd
import numpy as np

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "wab_keyword_diagnostic.xlsx")
ADMIN_PREFIXES = {"AAB ADMIN", "WAB ADMIN", "AAB ADMIN -", "WAB ADMIN -"}

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


# Noise words to exclude — broader than usual to surface domain-specific terms
STOP = {
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
    "will","one","two","new","now","use","used","using","set",
    # Email/security banner noise
    "external","message","caution","originated","outside","organization",
    "click","links","open","attachments","unless","recognize","sender",
    "safe","content","secure","proofpoint","encrypted","https","http","www",
    "com","org","net",
}


def extract_unigrams(texts, top_n=50):
    """Top single words, excluding stop words."""
    words = Counter()
    for t in texts:
        if not t or len(str(t)) < 3: continue
        for w in re.findall(r"[a-z][a-z0-9]{2,}", str(t).lower()):
            if w not in STOP:
                words[w] += 1
    return words.most_common(top_n)


def extract_bigrams(texts, top_n=40):
    """Top 2-word phrases — much more informative than single words."""
    bigrams = Counter()
    for t in texts:
        if not t or len(str(t)) < 5: continue
        words = re.findall(r"[a-z][a-z0-9]{2,}", str(t).lower())
        words = [w for w in words if w not in STOP]
        for i in range(len(words) - 1):
            bg = f"{words[i]} {words[i+1]}"
            bigrams[bg] += 1
    return bigrams.most_common(top_n)


def extract_from_act_subj_only(texts, top_n=40):
    """Top phrases from Activity Subject only (shorter, more structured)."""
    phrases = Counter()
    for t in texts:
        if not t or len(str(t)) < 5: continue
        # Normalize and count the full activity subject as a phrase
        clean = re.sub(r"\s+", " ", str(t).strip().lower())
        # Remove RE: FW: prefixes
        clean = re.sub(r"^(re|fw|fwd)\s*:\s*", "", clean).strip()
        if len(clean) > 3:
            phrases[clean] += 1
    return phrases.most_common(top_n)


def main():
    start = datetime.datetime.now()
    print(f"=== Keyword Diagnostic — {start.strftime('%Y-%m-%d %H:%M')} ===\n")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    df = pd.read_excel(CASE_FILE, header=0, engine="openpyxl")
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    df = df.dropna(axis=1, how="all")
    print(f"  Cases: {len(df):,}")

    co_col = find_col(df, "Company Name (Company) (Company)", "Company Name", "Customer")
    subj_col = find_col(df, "Subject")
    desc_col = find_col(df, "Description")
    act_col = find_col(df, "Activity Subject")
    hrs_col = find_col(df, "Resolved In Hours")

    df["_company"] = df[co_col].fillna("").astype(str).str.strip() if co_col else ""
    _upper = df["_company"].str.upper()
    df["_is_internal"] = _upper.apply(lambda x: any(x.startswith(p) for p in ADMIN_PREFIXES))
    df["_subject"] = df[subj_col].fillna("(blank)").astype(str).str.strip() if subj_col else ""
    df["_desc"] = df[desc_col].fillna("").astype(str).str.strip() if desc_col else ""
    df["_act_subj"] = df[act_col].fillna("").astype(str).str.strip() if act_col else ""
    df["_hours"] = safe_num(df[hrs_col]) if hrs_col else np.nan

    client = df[~df["_is_internal"]].copy()
    print(f"  Client inclusive: {len(client):,}\n")

    # Target subjects for sub-segmentation
    targets = ["Research", "Account Maintenance", "New Account Request",
               "General Questions", "Close Account"]

    sheets = OrderedDict()

    for subj in targets:
        s = client[client["_subject"] == subj]
        n = len(s)
        if n == 0:
            continue

        desc_texts = s[s["_desc"].str.len() > 0]["_desc"].tolist()
        act_texts = s[s["_act_subj"].str.len() > 5]["_act_subj"].tolist()
        all_texts = desc_texts + act_texts

        desc_fill = round(100 * len(desc_texts) / n, 1)
        act_fill = round(100 * len(act_texts) / n, 1)

        print(f"  {subj}: {n:,} cases (desc fill {desc_fill}%, act_subj fill {act_fill}%)")

        rows = []

        # Header info
        rows.append({"type": "INFO", "rank": 0, "term": f"Total cases: {n:,}", "count": "", "pct": ""})
        rows.append({"type": "INFO", "rank": 0, "term": f"Description fill: {desc_fill}%", "count": "", "pct": ""})
        rows.append({"type": "INFO", "rank": 0, "term": f"Activity Subject fill: {act_fill}%", "count": "", "pct": ""})
        rows.append({"type": "---", "rank": 0, "term": "", "count": "", "pct": ""})

        # Unigrams from all text
        uni = extract_unigrams(all_texts, top_n=50)
        for i, (word, count) in enumerate(uni, 1):
            rows.append({
                "type": "UNIGRAM (desc+act_subj)",
                "rank": i, "term": word,
                "count": count,
                "pct": f"{round(100 * count / n, 1)}%",
            })

        rows.append({"type": "---", "rank": 0, "term": "", "count": "", "pct": ""})

        # Bigrams from all text
        bi = extract_bigrams(all_texts, top_n=40)
        for i, (phrase, count) in enumerate(bi, 1):
            rows.append({
                "type": "BIGRAM (desc+act_subj)",
                "rank": i, "term": phrase,
                "count": count,
                "pct": f"{round(100 * count / n, 1)}%",
            })

        rows.append({"type": "---", "rank": 0, "term": "", "count": "", "pct": ""})

        # Top Activity Subject values (full phrases)
        act_phrases = extract_from_act_subj_only(act_texts, top_n=40)
        for i, (phrase, count) in enumerate(act_phrases, 1):
            rows.append({
                "type": "ACTIVITY SUBJECT (full value)",
                "rank": i, "term": phrase[:100],
                "count": count,
                "pct": f"{round(100 * count / n, 1)}%",
            })

        rows.append({"type": "---", "rank": 0, "term": "", "count": "", "pct": ""})

        # Unigrams from Description ONLY
        desc_uni = extract_unigrams(desc_texts, top_n=30)
        for i, (word, count) in enumerate(desc_uni, 1):
            rows.append({
                "type": "UNIGRAM (desc only)",
                "rank": i, "term": word,
                "count": count,
                "pct": f"{round(100 * count / len(desc_texts), 1)}%" if desc_texts else "",
            })

        sheet_name = subj.replace(" ", "")[:25]
        sheets[sheet_name] = pd.DataFrame(rows)

    # Write
    print(f"\nWriting: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        for name, sdf in sheets.items():
            write_sheet(writer, name[:31], sdf)

    print(f"Done in {(datetime.datetime.now() - start).total_seconds():.1f}s")
    print(f"\n{'='*60}")
    print("Review the output, then share screenshots.")
    print("We'll build clusters from what the data actually contains.")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
