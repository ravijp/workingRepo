"""
WAB Keyword Diagnostic — Compact View
=======================================
Produces ONE sheet with all 5 subjects side by side.
3 screenshots max to share the full picture.

Dependencies: pandas, openpyxl
"""

# ┌─────────────────────────────────────────────────────────┐
CASE_FILE  = r"C:\Users\YourName\Desktop\AAB All Cases.xlsx"
OUTPUT_DIR = r"C:\Users\YourName\Desktop\wab_output"
# └─────────────────────────────────────────────────────────┘

import os, re, datetime, warnings
from collections import Counter, OrderedDict
import pandas as pd, numpy as np

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "wab_keyword_diagnostic.xlsx")
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

STOP = {
    "the","and","for","that","this","with","from","your","have","are",
    "was","were","been","has","had","but","not","you","all","can",
    "will","about","which","their","them","into","also","our","out",
    "would","could","should","need","may","just","get","got","per",
    "via","please","thank","thanks","hello","dear","regards","sincerely",
    "sent","received","fyi","following","below","above","let","its",
    "who","how","when","where","what","why","any","each","more","some",
    "very","being","other","only","same","than","then","there",
    "these","those","such","both","does","doing","done","did","make",
    "made","take","took","give","gave","like","know","see","way",
    "one","two","new","now","use","used","using","set","said",
    "external","message","caution","originated","outside","organization",
    "click","links","open","attachments","unless","recognize","sender",
    "safe","content","secure","proofpoint","encrypted","https","http","www",
    "com","org","net","subject","mailto",
}

def bigrams(texts, top_n=25):
    bg = Counter()
    for t in texts:
        if not t or len(str(t)) < 5: continue
        words = [w for w in re.findall(r"[a-z][a-z0-9]{2,}", str(t).lower()) if w not in STOP]
        for i in range(len(words)-1):
            bg[f"{words[i]} {words[i+1]}"] += 1
    return bg.most_common(top_n)

def unigrams(texts, top_n=25):
    ug = Counter()
    for t in texts:
        if not t or len(str(t)) < 3: continue
        for w in re.findall(r"[a-z][a-z0-9]{2,}", str(t).lower()):
            if w not in STOP: ug[w] += 1
    return ug.most_common(top_n)

def act_subj_phrases(texts, top_n=20):
    ph = Counter()
    for t in texts:
        if not t or len(str(t)) < 5: continue
        c = re.sub(r"^(re|fw|fwd)\s*:\s*","", str(t).strip().lower(), flags=re.IGNORECASE).strip()
        if len(c) > 3: ph[c[:80]] += 1
    return ph.most_common(top_n)

def write_sheet(writer, name, df):
    sn = name[:31]
    df.to_excel(writer, sheet_name=sn, index=False, freeze_panes=(1,0))
    ws = writer.sheets[sn]
    for col_cells in ws.columns:
        mx = max(len(str(cell.value or "")) for cell in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(mx+2, 45)

def main():
    print(f"=== Keyword Diagnostic ===\n")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    df = pd.read_excel(CASE_FILE, header=0, engine="openpyxl")
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")].dropna(axis=1, how="all")

    co_col = find_col(df, "Company Name (Company) (Company)", "Company Name")
    subj_col = find_col(df, "Subject")
    desc_col = find_col(df, "Description")
    act_col = find_col(df, "Activity Subject")

    df["_co"] = df[co_col].fillna("").astype(str).str.strip() if co_col else ""
    df["_is_internal"] = df["_co"].str.upper().apply(lambda x: any(x.startswith(p) for p in ADMIN_PREFIXES))
    df["_subj"] = df[subj_col].fillna("").astype(str).str.strip() if subj_col else ""
    df["_desc"] = df[desc_col].fillna("").astype(str).str.strip() if desc_col else ""
    df["_act"] = df[act_col].fillna("").astype(str).str.strip() if act_col else ""

    client = df[~df["_is_internal"]]
    targets = ["Research", "Account Maintenance", "New Account Request",
               "General Questions", "Close Account"]

    # ── Sheet 1: Bigrams side by side (most informative) ──
    max_rows = 25
    bg_data = {}
    for subj in targets:
        s = client[client["_subj"] == subj]
        texts = s[s["_desc"].str.len() > 0]["_desc"].tolist() + s[s["_act"].str.len() > 5]["_act"].tolist()
        bg = bigrams(texts, max_rows)
        bg_data[subj] = bg
        print(f"  {subj}: {len(s):,} cases, {len(texts):,} texts")

    rows = []
    for i in range(max_rows):
        row = {"rank": i+1}
        for subj in targets:
            short = subj.replace(" ","")[:10]
            if i < len(bg_data[subj]):
                term, count = bg_data[subj][i]
                row[f"{short}_bigram"] = term
                row[f"{short}_count"] = count
            else:
                row[f"{short}_bigram"] = ""
                row[f"{short}_count"] = ""
        rows.append(row)

    sheet1 = pd.DataFrame(rows)

    # ── Sheet 2: Unigrams side by side ──
    ug_data = {}
    for subj in targets:
        s = client[client["_subj"] == subj]
        texts = s[s["_desc"].str.len() > 0]["_desc"].tolist() + s[s["_act"].str.len() > 5]["_act"].tolist()
        ug_data[subj] = unigrams(texts, max_rows)

    rows2 = []
    for i in range(max_rows):
        row = {"rank": i+1}
        for subj in targets:
            short = subj.replace(" ","")[:10]
            if i < len(ug_data[subj]):
                term, count = ug_data[subj][i]
                row[f"{short}_word"] = term
                row[f"{short}_count"] = count
            else:
                row[f"{short}_word"] = ""
                row[f"{short}_count"] = ""
        rows2.append(row)

    sheet2 = pd.DataFrame(rows2)

    # ── Sheet 3: Top Activity Subject phrases (full values) side by side ──
    act_data = {}
    for subj in targets:
        s = client[client["_subj"] == subj]
        act_texts = s[s["_act"].str.len() > 5]["_act"].tolist()
        act_data[subj] = act_subj_phrases(act_texts, 20)

    rows3 = []
    for i in range(20):
        row = {"rank": i+1}
        for subj in targets:
            short = subj.replace(" ","")[:10]
            if i < len(act_data[subj]):
                phrase, count = act_data[subj][i]
                row[f"{short}_phrase"] = phrase
                row[f"{short}_count"] = count
            else:
                row[f"{short}_phrase"] = ""
                row[f"{short}_count"] = ""
        rows3.append(row)

    sheet3 = pd.DataFrame(rows3)

    # Write
    print(f"\nWriting: {OUTPUT_XLSX}")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        write_sheet(writer, "Bigrams", sheet1)
        write_sheet(writer, "Unigrams", sheet2)
        write_sheet(writer, "ActivitySubjectPhrases", sheet3)

    print("Done. 3 sheets — screenshot each one.")

if __name__ == "__main__":
    main()
