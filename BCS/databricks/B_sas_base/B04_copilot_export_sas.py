# Databricks notebook source
# MAGIC %md
# MAGIC # B04. The masked-excerpt export (assistant discovery batches), SAS gate
# MAGIC
# MAGIC =====================================================================
# MAGIC GATE CLEARED FOR DISCOVERY (2026-07-13): transcripts may be used in the
# MAGIC bank-tenant assistant UI for discovery; governance questions apply to
# MAGIC production deployment. The owner's explicit go is on record.
# MAGIC Masking: kept as the DEFAULT. The digit mask costs discovery nothing
# MAGIC (digits carry no signal for phrase and category proposals) and removes
# MAGIC accidental account-number spill from clipboard files. To unmask (owner's
# MAGIC call, one edit): change the marked line in the turns CTE from
# MAGIC   regexp_replace(t.content, '[0-9]{3,}', '###')   to   t.content
# MAGIC Output is row-level text: download the CSV to the batch folder but NEVER
# MAGIC import it into the story JSON; keep excerpt content out of screenshots
# MAGIC (excerpts travel only inside the batch files).
# MAGIC =====================================================================
# MAGIC
# MAGIC Spark port of the copilot-batch-kit masking-export, driven off the 04s
# MAGIC table (SAS-gate strata per the approved Phase-2 plan). Strata (priority
# MAGIC order, first match wins; quotas A 9 / B 6 / C 9 / D 6):
# MAGIC   B. deceased-adjacent: the deceased lexicon fires on the episode
# MAGIC   D. high-CO12 language: still DQ1 in M2 (export DLNQT_CD_M2 = 1),
# MAGIC      plan/settlement or hardship group
# MAGIC   A. leaked-intent work list: payment-language episode on a leaked_sas
# MAGIC      account, non-deceased
# MAGIC   C. silent still-DQ1: DLNQT_CD_M2 = 1, has a transcript, no
# MAGIC      payment-related language, account NOT captured_sas
# MAGIC The AWS bridge/caller class rides along as a carried LABEL only.
# MAGIC Deterministic pick: xxhash64(contactid) order within stratum (no
# MAGIC random()); outer LIMIT 100 guard; batches round-robin 1..3.

# COMMAND ----------

# =====================================================================
# SETUP - keep in sync with B00_setup.py (the canonical copy).
# =====================================================================
import datetime as _dt
import time as _time

# --- parameters: plain constants; widgets override them when available ---
CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
ANCHOR_YM = "202501"
SAS_CSV_PATH = "/Volumes/cda_model_shared/ecm_cld_model/ecm_cld/collections_zenon/WATERFALL_COLL_CALL_V2_202501.csv"
FMT_CATALOG = "634153504162_glue_connection_catalog"      # fmt_acct_c lives here
CC_CATALOG = "062108867742_glue_connectivity_catalog"     # call + transcript live here

try:
    dbutils.widgets.text("CATALOG", CATALOG);           CATALOG = dbutils.widgets.get("CATALOG")
    dbutils.widgets.text("SCHEMA", SCHEMA);             SCHEMA = dbutils.widgets.get("SCHEMA")
    dbutils.widgets.text("ANCHOR_YM", ANCHOR_YM);       ANCHOR_YM = dbutils.widgets.get("ANCHOR_YM")
    dbutils.widgets.text("SAS_CSV_PATH", SAS_CSV_PATH); SAS_CSV_PATH = dbutils.widgets.get("SAS_CSV_PATH")
    dbutils.widgets.text("FMT_CATALOG", FMT_CATALOG);   FMT_CATALOG = dbutils.widgets.get("FMT_CATALOG")
    dbutils.widgets.text("CC_CATALOG", CC_CATALOG);     CC_CATALOG = dbutils.widgets.get("CC_CATALOG")
except NameError:
    pass  # dbutils absent (plain Spark); the constants above apply

DB = f"{CATALOG}.{SCHEMA}"
FMT = f"`{FMT_CATALOG}`.fmt_acct_dba.fmt_acct_c"
CALL = f"`{CC_CATALOG}`.contactcenter_bdp_db.`call`"
TX = f"`{CC_CATALOG}`.contactcenter_bdp_db.transcript"

# --- derived windows (all from ANCHOR_YM; 202501 literals asserted below) ---
_a0 = _dt.date(int(ANCHOR_YM[:4]), int(ANCHOR_YM[4:6]), 1)
_mm = lambda d, k: _dt.date(d.year + (d.month - 1 + k) // 12, (d.month - 1 + k) % 12 + 1, 1)  # month move (date plumbing)

PRV_YM = _mm(_a0, -1).strftime("%Y%m")                                   # 202412
FEB_YM = _mm(_a0, 1).strftime("%Y%m")                                    # 202502
MAR_YM = _mm(_a0, 2).strftime("%Y%m")                                    # 202503
MONTH_WIN_START = _mm(_a0, -1).strftime("%Y%m%d")                        # 00 scan start
MONTH_WIN_END = _mm(_a0, 3).strftime("%Y%m%d")                           # 00 scan end (excl)
CALL_WIN_START = _a0.isoformat()                                         # 2025-01-01
CALL_WIN_END = _mm(_a0, 1).isoformat()                                   # 2025-02-01
EFFDT_CAP_START = _a0.isoformat()                                        # 2025-01-01
EFFDT_CAP_END = (_mm(_a0, 1) + _dt.timedelta(days=1)).isoformat()        # 2025-02-02
CLEANUP_DATE = _a0.isoformat()                                           # 2025-01-01
ANCHOR_EOM = (_mm(_a0, 1) - _dt.timedelta(days=1)).isoformat()           # 2025-01-31
FEB_START = _mm(_a0, 1).isoformat()                                      # 2025-02-01
MAR_START = _mm(_a0, 2).isoformat()                                      # 2025-03-01
APR_START = _mm(_a0, 3).isoformat()                                      # 2025-04-01
CO8_END = (_mm(_a0, 9) - _dt.timedelta(days=1)).isoformat()              # 2025-09-30
CO10_END = (_mm(_a0, 11) - _dt.timedelta(days=1)).isoformat()            # 2025-11-30
CO12_END = (_mm(_a0, 13) - _dt.timedelta(days=1)).isoformat()            # 2026-01-31
FWD_CO_START = _a0.strftime("%Y%m%d")                                    # 20250101
FWD_CO_END = _mm(_a0, 12).strftime("%Y%m%d")                             # 20260101
SNAP_DAILY_START = _mm(_a0, -7).strftime("%Y%m%d")                       # 20240601
SNAP_DAILY_END = _mm(_a0, 1).strftime("%Y%m%d")                          # 20250201
EFFDT_SCAN_START = _mm(_a0, -1).isoformat()                              # 2024-12-01
# NOT vintage-derived: the bounded-scan guard on the call table's live
# loading edge (round-11 lesson), and the frozen predicate behind the
# 75,883 evidence tie. Do not move it.
EFFDT_HARD_END = "2026-07-10"

if ANCHOR_YM == "202501":
    assert (MONTH_WIN_START, MONTH_WIN_END) == ("20241201", "20250401"), "derived 00 window drifted"
    assert (CALL_WIN_START, CALL_WIN_END, EFFDT_CAP_END) == ("2025-01-01", "2025-02-01", "2025-02-02"), "derived call window drifted"
    assert (ANCHOR_EOM, CO8_END, CO10_END, CO12_END) == ("2025-01-31", "2025-09-30", "2025-11-30", "2026-01-31"), "derived CO windows drifted"
    assert (FWD_CO_START, FWD_CO_END, SNAP_DAILY_START) == ("20250101", "20260101", "20240601"), "derived forward/daily windows drifted"

# --- EXPECTED: per-vintage expectations. None = measure mode (two-phase lock).
EXPECTED = {
    "202501": {
        # population anchors - assert EXACT (the key change must not move them)
        "ledger all": 204323,
        "ledger AA row": 15177,
        "ledger exaa": 189146,
        "ledger exaa balance": 457943987,   # +/- 5 tolerance (recorded platform rounding)
        "touched b1": 724848,
        "touched a. cured": 464023,
        "touched b. bucket 1": 186714,
        "touched c. rolled past": 69513,
        "touched d. jan chargeoff": 4598,
        # SAS-native waterfall - assert EXACT (replication facts about the export)
        "csv rows": 610183,
        "csv distinct accounts": 610183,
        "wf 02 dq1": 202479,
        "wf 03 +cpc": 186848,
        "wf 04 sas ledger": 186013,
        # call-table evidence ties - assert EXACT (immutable under the effdt bound)
        "jan acctid null rows": 481838,
        "jan acctid digits-only rows": 1267227,
        "jan acctid other-shape rows": 0,
        "jan key mismatches (id-carrying inbound)": 75883,
        # historical string-key references - context only, NEVER asserted
        "hist ledger callers (string key)": 9389,
        "hist ledger episodes (string key)": 11262,
        "hist callday b1 stream": 29114,
        # ---- measured then locked (None = measure mode) ----
        # B02 (key fix)
        "ledger callers (numeric key)": None,
        "ledger episodes (numeric key)": None,
        "addressable episodes (callday b1 stream)": None,
        "addressable work list episodes": None,
        "addressable work list accounts": None,
        "language partition": None,          # dict {group: episodes} at lock
        "caller classes (aws gate)": None,   # dict {class: accounts} at lock
        "W strict leaked accounts": None,
        "W deceased routed": None,
        "W accounts": None,
        "W balance": None,
        "gained callers": None,
        "gap1942 recovered": None,
        "gained outside 1942": None,
        "flagged overlap (202501 recon)": None,
        # B01 (spine)
        "inb_native 01 total": None,
        "inb_native 02 dq1": None,
        "inb_native 03 +cpc": None,
        "inb_native 04 ledger": None,
        "captured_sas all": None,
        "captured_sas ledger": None,
        "ledger eop_bal_m1 sum": None,
        "ledger ecl_m1 sum": None,
        # B02b (outcomes)
        "04s episodes": None,
        "04s callers": None,
        "04s captured_sas accounts": None,
        "04s leaked_sas accounts": None,
        "04s W_s accounts": None,
        # B03 (insights)
        "funnel called (inb_native, ledger)": None,
        "funnel callers with episodes": None,
        "funnel intent accounts": None,
        "funnel leaked accounts": None,
    },
}
assert ANCHOR_YM in EXPECTED, f"no expectations recorded for vintage {ANCHOR_YM} - STOP (add a reviewed EXPECTED block first)"
E = EXPECTED[ANCHOR_YM]

# --- analysis helpers: chk() and NUM_KEY only -----------------------------
NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"   # THE numeric key rule

# --- output plumbing (per the output-ergonomics instruction, 2026-07-16):
# lossless grids, transposed wide pulls, record block, on-platform metrics,
# failure context, section timing. Results cross as screenshots; output must
# be lossless and transcription-friendly.
RESULTS = []          # (name, actual, expected, status)
GRIDS = {}            # name -> (columns, rows); reprinted in the record block
_T0 = _time.time()
_TSEC = [_time.time()]


def sec(title):
    now = _time.time()
    print(f"\n===== {title} =====  (prev section {int(now - _TSEC[0])}s, elapsed {int(now - _T0)}s)")
    _TSEC[0] = now


def fmt(v):
    if v is None:
        return "NULL"
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return f"{v:,}"
    if isinstance(v, float):
        return f"{v:,.0f}"          # thousands separators, no scientific notation
    return str(v)


def chk(name, actual, expected, tol=0, ctx=None):
    """Raising anchor check; a miss STOPS the run. expected=None -> measure mode.
    ctx: optional DataFrame printed BEFORE raising, so the failure screenshot
    carries its own diagnosis."""
    if expected is None:
        RESULTS.append((name, actual, None, "MEASURED"))
        print(f"MEASURED  {name} = {fmt(actual)}")
        return
    ok = (abs(actual - expected) <= tol) if tol else (actual == expected)
    if not ok:
        if ctx is not None:
            print(f"CONTEXT for the failing check '{name}':")
            ctx.show(500, truncate=False)
        raise AssertionError(f"ANCHOR MISS {name}: got {fmt(actual)}, expected {fmt(expected)}"
                             + (f" (tol {tol})" if tol else ""))
    RESULTS.append((name, actual, expected, "PASS"))
    print(f"PASS  {name} = {fmt(actual)}")


def grid(name, df):
    """Lossless, transcription-friendly grid: markdown text, running index,
    40-row chunk banners. The SQL behind df must carry its own ORDER BY."""
    rows = df.collect()
    cols = df.columns
    print(f"{name}: {len(rows)} rows, all shown")
    header = "| idx | " + " | ".join(cols) + " |"
    rule = "|" + "---|" * (len(cols) + 1)
    for i, r in enumerate(rows):
        if i % 40 == 0:
            print(f"-- {name}: rows {i + 1}..{min(i + 40, len(rows))} of {len(rows)} --")
            print(header)
            print(rule)
        print("| " + str(i + 1) + " | " + " | ".join(fmt(r[c]) for c in cols) + " |")
    GRIDS[name] = (cols, rows)


def kv(name, df):
    """Transpose a wide one-row pull: one metric per line, never a wide table."""
    r = df.first()
    print(f"{name}: (one row, transposed)")
    for c in df.columns:
        print(f"  {c} = {fmt(r[c])}")
    GRIDS[name] = (["metric", "value"], [(c, r[c]) for c in df.columns])


def record_block(notebook):
    """The block that gets screenshotted and transcribed verbatim."""
    print("=" * 78)
    print(f"RECORD BLOCK  {notebook}  vintage {ANCHOR_YM}  platform Databricks")
    print("=" * 78)
    for name, actual, expected, status in RESULTS:
        tail = "" if expected is None else f"  (expected {fmt(expected)})"
        print(f"{status:8}  {name} = {fmt(actual)}{tail}")
    for gname, (cols, rows) in GRIDS.items():
        print(f"\n### {gname} ({len(rows)} rows)")
        print("| " + " | ".join(map(str, cols)) + " |")
        print("|" + "---|" * len(cols))
        for r in rows:
            vals = r if isinstance(r, tuple) else [r[c] for c in cols]
            print("| " + " | ".join(fmt(v) for v in vals) + " |")
    print("=" * 78)


def flush_metrics(notebook):
    """On-platform provenance: every chk and measured value survives across
    sittings in a small Delta table. Screenshots stay the transfer channel."""
    stage = [(notebook, n,
              float(a) if isinstance(a, (int, float)) and not isinstance(a, bool) else None,
              fmt(a),
              float(e) if isinstance(e, (int, float)) and not isinstance(e, bool) else None,
              s, ANCHOR_YM) for n, a, e, s in RESULTS]
    spark.createDataFrame(
        stage,
        "notebook string, name string, value double, value_str string, expected double, status string, vintage string"
    ).createOrReplaceTempView("_metrics_stage")
    spark.sql(f"""CREATE TABLE IF NOT EXISTS {DB}.uc2_run_metrics
                  (run_ts timestamp, notebook string, name string, value double,
                   value_str string, expected double, status string, vintage string)""")
    spark.sql(f"INSERT INTO {DB}.uc2_run_metrics SELECT current_timestamp(), * FROM _metrics_stage")
    print(f"metrics appended to {DB}.uc2_run_metrics ({len(stage)} rows)")


print(f"SETUP OK: vintage {ANCHOR_YM}; layers -> {DB}")
print(f"  fmt   = {FMT}")
print(f"  call  = {CALL}")
print(f"  tx    = {TX}")
# =====================================================================
# end of SETUP
# =====================================================================

# COMMAND ----------

# MAGIC %md
# MAGIC ## X1. Preconditions

# COMMAND ----------

sec("X1 preconditions")

OUT = f"{DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}"
assert spark.catalog.tableExists(OUT), f"PRECONDITION MISS: {OUT} missing - run B02b first"
print(f"PASS  table exists: {OUT}")
assert spark.catalog.tableExists(f"{CC_CATALOG}.contactcenter_bdp_db.transcript"), \
    "PRECONDITION MISS: transcript table not reachable"
print("PASS  transcript table reachable")

# COMMAND ----------

# MAGIC %md
# MAGIC ## X2. Stratum assignment (SAS-gate definitions; priority order B, D, A, C)

# COMMAND ----------

sec("X2 stratum assignment")

spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW cb_pool AS
    SELECT *,
           CASE
             WHEN deceased_f = 1
               THEN 'B. deceased-adjacent'
             WHEN try_cast(dlnqt_cd_m2 AS int) = 1
                  AND language_group IN ('d. plan or settlement talk', 'e. hardship talk')
               THEN 'D. high-CO12 language'
             WHEN pay_f > 0 AND leaked_sas AND deceased_acct = 0
               THEN 'A. leaked-intent work list'
             WHEN try_cast(dlnqt_cd_m2 AS int) = 1 AND has_tx = 1
                  AND language_group = 'g. no payment-related language'
                  AND NOT captured_sas
               THEN 'C. silent still-DQ1'
           END AS stratum
    FROM {OUT}
""")

grid("stratum pool sizes (episodes; sanity, not a record number)", spark.sql("""
    SELECT stratum, count(*) AS episodes, count(DISTINCT acct_key) AS accounts
    FROM cb_pool
    WHERE stratum IS NOT NULL
    GROUP BY 1 ORDER BY 1
"""))

# COMMAND ----------

# MAGIC %md
# MAGIC ## X3. Deterministic pick and batch assignment (quotas A 9 / B 6 / C 9 / D 6)

# COMMAND ----------

sec("X3 pick and batch")

spark.sql("""
    CREATE OR REPLACE TEMP VIEW cb_sampled AS
    SELECT *,
           ((row_number() OVER (ORDER BY stratum, pick_rn) - 1) % 3) + 1 AS batch_nbr
    FROM (
        SELECT *,
               -- xxhash64 on the string directly (the Trino original hashed a
               -- varbinary cast); pick ORDER differs from the old kit by
               -- design - this is a fresh deterministic sample, not a replay
               row_number() OVER (PARTITION BY stratum ORDER BY xxhash64(contactid)) AS pick_rn
        FROM cb_pool
        WHERE stratum IS NOT NULL
    )
    WHERE (stratum LIKE 'A.%' AND pick_rn <= 9)
       OR (stratum LIKE 'B.%' AND pick_rn <= 6)
       OR (stratum LIKE 'C.%' AND pick_rn <= 9)
       OR (stratum LIKE 'D.%' AND pick_rn <= 6)
""")

_pick_df = spark.sql("SELECT stratum, count(*) AS picked FROM cb_sampled GROUP BY 1 ORDER BY 1")
grid("picked per stratum", _pick_df)
for _row in _pick_df.collect():
    _quota = {"A": 9, "B": 6, "C": 9, "D": 6}[_row["stratum"][0]]
    assert _row["picked"] <= _quota, f"quota exceeded for {_row['stratum']}: {_row['picked']} > {_quota}"
_r = spark.sql("SELECT count(*) AS n FROM cb_sampled").first()
assert _r["n"] <= 30, f"sampled {_r['n']} exceeds the 30-excerpt design (quotas 9/6/9/6)"
chk("sampled excerpts", _r["n"], None)

# COMMAND ----------

# MAGIC %md
# MAGIC ## X4. Fetch the sampled turns (masked) and assemble the export
# MAGIC
# MAGIC ROW-LEVEL TEXT: use the display's download to save the CSV into the batch
# MAGIC folder. NEVER screenshot the excerpt content; never import it into the
# MAGIC story JSON. This is the package's only row-level-text query, and it
# MAGIC touches the transcript table only for the ~30 sampled contactids.

# COMMAND ----------

sec("X4 masked export")

export_df = spark.sql(f"""
    WITH turns_rows AS (
        SELECT t.contactid, t.beginmillis,
               concat(t.participantid, ': ',
                      -- UNMASK EDIT POINT (owner-gated): replace the
                      -- regexp_replace(...) below with t.content
                      regexp_replace(t.content, '[0-9]{{3,}}', '###')
               ) AS line
        FROM {TX} t
        JOIN (SELECT DISTINCT contactid FROM cb_sampled) s ON s.contactid = t.contactid
        WHERE t.content IS NOT NULL
          AND t.effdt >= '{EFFDT_CAP_START}' AND t.effdt < '{EFFDT_CAP_END}'
    ),
    turns AS (
        SELECT contactid,
               array_join(transform(array_sort(collect_list(struct(beginmillis, line))),
                                    x -> x.line), char(10)) AS convo
        FROM turns_rows
        GROUP BY 1
    )
    SELECT s.batch_nbr AS cb_batch,
           concat(substr(s.stratum, 1, 1), cast(s.pick_rn AS string)) AS cb_excerpt_id,
           s.stratum AS cb_stratum,
           s.aws_caller_class AS cb_aws_class_label,   -- carried label only, AWS day-grain classes
           s.contactid AS cb_contactid,
           s.call_dt AS cb_call_dt,
           CASE WHEN s.captured_sas
                THEN 'account captured (payment in call month or next, export month grain)'
                ELSE 'account leaked (no payment in call month or next, export month grain)'
           END AS cb_outcome,
           substr(v.convo, 1, 8000) AS cb_transcript_masked
    FROM cb_sampled s
    LEFT JOIN turns v ON v.contactid = s.contactid
    ORDER BY 1, 3, s.pick_rn
    LIMIT 100
""")
_n = export_df.count()
chk("export rows", _n, None)
print("Download the grid below as CSV into the batch folder. DO NOT screenshot")
print("the excerpt content; excerpts travel only inside the batch files.")
display(export_df)

# COMMAND ----------

sec("X5 verdict")

record_block("B04_copilot_export_sas")
flush_metrics("B04_copilot_export_sas")
