# Databricks notebook source
# MAGIC %md
# MAGIC # B06. Assistant-response ingestion: responses become DATA
# MAGIC
# MAGIC The scale campaign's persistence layer. The assistant UI is clipboard-only,
# MAGIC so raw responses come back as pasted text; this notebook parses the
# MAGIC format-locked response formats into a Delta table so every downstream
# MAGIC number (blocker rates, agreement stats, dollar weights) comes from SQL,
# MAGIC never from a reading session.
# MAGIC
# MAGIC Builds TWO tables:
# MAGIC   uc2_copilot_excerpt_map  - which excerpt id is which contactid/account
# MAGIC                              (wave 1 re-derived deterministically from the
# MAGIC                              B04 sampler; asserted against the batch files
# MAGIC                              of record before insert)
# MAGIC   uc2_copilot_labels       - one row per parsed response line
# MAGIC                              (wave, batch, prompt type, excerpt id, field,
# MAGIC                              value, provenance, parse status)
# MAGIC
# MAGIC RULES: the parser never edits or repairs response content - unparsed or
# MAGIC out-of-format lines land as FLAGGED rows with the raw line kept verbatim,
# MAGIC and a human adjudicates them. Raw response FILES stay the source of
# MAGIC record (saved verbatim per the runbook); this table is the analysis copy.
# MAGIC No excerpt transcript text enters this notebook or its outputs.
# MAGIC
# MAGIC Supported prompt types: 'discovery' (Round A), 'classify' (Round B),
# MAGIC 'second-read' (the 20% re-read; same format as classify).
# MAGIC 'contrast' and 'rubric' ingest here AFTER their wording freezes (wave 2).
# MAGIC
# MAGIC Per response file: fill the L3 paste cell, run L3 -> L4 -> L5. Repeat.
# MAGIC L6 prints the current campaign stats whenever data exists.

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
        # B02 (key fix) - LOCKED 2026-07-16 from the verified first run
        # (Phase2SS_0 screenshots; every arithmetic tie re-added by hand:
        # 9,389 + 1,958 = 11,347; 1,942 + 16 = 1,958; overlap 9,194 + 1,942
        # = 11,136; classes re-add to 189,146 and $457,943,985; language
        # re-adds to 13,788)
        "ledger callers (numeric key)": 11347,
        "ledger episodes (numeric key)": 13788,
        "addressable episodes (callday b1 stream)": 36594,
        "addressable work list episodes": 2709,
        "addressable work list accounts": 2543,
        "language partition": {
            "a. deceased or estate": 534,
            "b. future-dated promise": 1725,
            "c. payment talk, no promise": 6185,
            "d. plan or settlement talk": 511,
            "e. hardship talk": 99,
            "f. dispute or fraud talk": 306,
            "g. no payment-related language": 4428,
        },
        "caller classes (aws gate)": {
            "a. non-caller": 177799,
            "b. captured (>= 1 paid-30d episode)": 7101,
            "c. leaked-intent (intent, no payment 30d)": 2451,
            "d. other-caller": 1795,
        },
        "W strict leaked accounts": 2451,
        "W deceased routed": 172,
        "W accounts": 2279,
        "W balance": 9277926,
        "gained callers": 1958,
        "gap1942 recovered": 1942,
        "gained outside 1942": 16,
        "flagged overlap (202501 recon)": 11136,
        # B01 (spine) - LOCKED 2026-07-16 from the verified first run.
        # The one-time tie-out was PERFECTLY DIAGONAL: inb_native equals the
        # CSV flag on all 610,183 accounts (575,949 + 34,234; ledger
        # 174,877 + 11,136), so the CSV flag is retired and the native
        # ladder happens to equal the old CSV ladder exactly.
        "inb_native 01 total": 34234,
        "inb_native 02 dq1": 12615,
        "inb_native 03 +cpc": 11289,
        "inb_native 04 ledger": 11136,
        "captured_sas all": 278885,
        "captured_sas ledger": 125275,
        "ledger eop_bal_m1 sum": 452444591,
        "ledger ecl_m1 sum": 93543576,
        # B02b (outcomes) - LOCKED 2026-07-16 (8,037 + 1,801 + 1,298 other
        # = 11,136; W_s = 1,801 - 155 routed = 1,646)
        "04s episodes": 13486,
        "04s callers": 11136,
        "04s captured_sas accounts": 8037,
        "04s leaked_sas accounts": 1801,
        "04s W_s accounts": 1646,
        # B03 (insights) - LOCKED 2026-07-16 (funnel steps 2 and 3 are both
        # 11,136: a MEASURED equality - every native caller in the SAS
        # ledger has a standard episode - not an a-priori identity)
        "funnel called (inb_native, ledger)": 11136,
        "funnel callers with episodes": 11136,
        "funnel intent accounts": 7459,
        "funnel leaked accounts": 1801,
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
# MAGIC ## L1. Preconditions and the two campaign tables

# COMMAND ----------

sec("L1 preconditions + DDL")

OUT = f"{DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}"
assert spark.catalog.tableExists(OUT), f"PRECONDITION MISS: {OUT} missing - run B02b first"
print(f"PASS  table exists: {OUT}")

spark.sql(f"""
    CREATE TABLE IF NOT EXISTS {DB}.uc2_copilot_excerpt_map (
        created_ts timestamp, vintage string, wave int, batch int,
        excerpt_id string, stratum string, contactid string, acct_key string,
        call_dt date, captured_sas boolean, pick_rn int)
""")
spark.sql(f"""
    CREATE TABLE IF NOT EXISTS {DB}.uc2_copilot_labels (
        run_ts timestamp, vintage string, wave int, batch int,
        prompt_type string, excerpt_id string, field string, value string,
        response_file string, operator string, parse_status string,
        raw_line string)
""")
print(f"tables ready: {DB}.uc2_copilot_excerpt_map, {DB}.uc2_copilot_labels")

# COMMAND ----------

# MAGIC %md
# MAGIC ## L2. The wave-1 excerpt map (deterministic replay of the B04 sampler)
# MAGIC
# MAGIC The B04 pick is deterministic (xxhash64 order within stratum, no
# MAGIC random()), so re-deriving it here reproduces the batch of record
# MAGIC EXACTLY - verified by asserting the derived batch-to-id map against the
# MAGIC batch files of record (30 excerpts, round-12 addendum 5b) BEFORE any row
# MAGIC is inserted. A mismatch means the substrate moved: STOP.
# MAGIC
# MAGIC The CTE chain below is VERBATIM from B04 X2/X3 (keep in sync). Skipped
# MAGIC automatically when the wave-1 map already holds its 30 rows. Wave 2+
# MAGIC maps are written by the wave-2 sampler itself, not by replay.

# COMMAND ----------

sec("L2 wave-1 excerpt map")

# the batch files of record: cb_batch30 202501 (round-12 addendum 5b)
EXPECTED_WAVE1_IDS = {
    1: {"A1", "A4", "A7", "B1", "B4", "C1", "C4", "C7", "D1", "D4"},
    2: {"A2", "A5", "A8", "B2", "B5", "C2", "C5", "C8", "D2", "D5"},
    3: {"A3", "A6", "A9", "B3", "B6", "C3", "C6", "C9", "D3", "D6"},
}

_n = spark.sql(f"""
    SELECT count(*) AS n FROM {DB}.uc2_copilot_excerpt_map
    WHERE wave = 1 AND vintage = '{ANCHOR_YM}'
""").first()["n"]
if _n == 30:
    print("SKIP: wave-1 map already built (30 rows present)")
    chk("wave-1 map rows (already present)", _n, 30)
else:
    if _n > 0:
        print(f"wave-1 map incomplete ({_n} rows) - rebuilding")
        spark.sql(f"DELETE FROM {DB}.uc2_copilot_excerpt_map WHERE wave = 1 AND vintage = '{ANCHOR_YM}'")

    # ---- VERBATIM from B04 X2/X3 (keep in sync) ----
    spark.sql(f"""
        CREATE OR REPLACE TEMP VIEW c_with_tx AS
        SELECT DISTINCT t.contactid
        FROM {TX} t
        JOIN (
            SELECT DISTINCT contactid
            FROM {OUT}
            WHERE try_cast(dlnqt_cd_m2 AS int) = 1
              AND language_group = 'g. no payment-related language'
              AND NOT captured_sas
        ) c ON c.contactid = t.contactid
        WHERE t.content IS NOT NULL
          AND t.effdt >= '{EFFDT_CAP_START}' AND t.effdt < '{EFFDT_CAP_END}'
    """)
    spark.sql(f"""
        CREATE OR REPLACE TEMP VIEW cb_pool AS
        SELECT o.*,
               CASE
                 WHEN o.deceased_f = 1
                   THEN 'B. deceased-adjacent'
                 WHEN try_cast(o.dlnqt_cd_m2 AS int) = 1
                      AND o.language_group IN ('d. plan or settlement talk', 'e. hardship talk')
                   THEN 'D. high-CO12 language'
                 WHEN o.pay_f > 0 AND o.leaked_sas AND o.deceased_acct = 0
                   THEN 'A. leaked-intent work list'
                 WHEN try_cast(o.dlnqt_cd_m2 AS int) = 1
                      AND w.contactid IS NOT NULL
                      AND o.language_group = 'g. no payment-related language'
                      AND NOT o.captured_sas
                   THEN 'C. silent still-DQ1'
               END AS stratum
        FROM {OUT} o
        LEFT JOIN c_with_tx w ON w.contactid = o.contactid
    """)
    spark.sql("""
        CREATE OR REPLACE TEMP VIEW cb_sampled AS
        SELECT *,
               ((row_number() OVER (ORDER BY stratum, pick_rn) - 1) % 3) + 1 AS batch_nbr
        FROM (
            SELECT *,
                   row_number() OVER (PARTITION BY stratum ORDER BY xxhash64(contactid)) AS pick_rn
            FROM cb_pool
            WHERE stratum IS NOT NULL
        )
        WHERE (stratum LIKE 'A.%' AND pick_rn <= 9)
           OR (stratum LIKE 'B.%' AND pick_rn <= 6)
           OR (stratum LIKE 'C.%' AND pick_rn <= 9)
           OR (stratum LIKE 'D.%' AND pick_rn <= 6)
    """)
    # ---- end of the verbatim B04 chain ----

    _rows = spark.sql("""
        SELECT batch_nbr,
               concat(substr(stratum, 1, 1), cast(pick_rn AS string)) AS excerpt_id,
               stratum, contactid, acct_key, call_dt, captured_sas, pick_rn
        FROM cb_sampled
        ORDER BY batch_nbr, stratum, pick_rn
    """).collect()
    chk("wave-1 replay sampled excerpts", len(_rows), 30)
    _per_stratum = {}
    _per_batch = {}
    for _r in _rows:
        _per_stratum[_r["stratum"][0]] = _per_stratum.get(_r["stratum"][0], 0) + 1
        _per_batch.setdefault(_r["batch_nbr"], set()).add(_r["excerpt_id"])
    for _s, _q in {"A": 9, "B": 6, "C": 9, "D": 6}.items():
        chk(f"wave-1 replay stratum {_s} picked", _per_stratum.get(_s, 0), _q)
    for _b in (1, 2, 3):
        assert _per_batch.get(_b, set()) == EXPECTED_WAVE1_IDS[_b], (
            f"ANCHOR MISS: batch {_b} ids {sorted(_per_batch.get(_b, set()))} "
            f"!= batch file of record {sorted(EXPECTED_WAVE1_IDS[_b])} - the substrate moved, STOP")
        print(f"PASS  batch {_b} ids match the batch file of record")

    spark.createDataFrame(
        [(ANCHOR_YM, 1, r["batch_nbr"], r["excerpt_id"], r["stratum"], r["contactid"],
          r["acct_key"], r["call_dt"], r["captured_sas"], r["pick_rn"]) for r in _rows],
        "vintage string, wave int, batch int, excerpt_id string, stratum string, "
        "contactid string, acct_key string, call_dt date, captured_sas boolean, pick_rn int"
    ).createOrReplaceTempView("_map_stage")
    spark.sql(f"""
        INSERT INTO {DB}.uc2_copilot_excerpt_map
        SELECT current_timestamp(), * FROM _map_stage
    """)
    print("wave-1 map inserted (30 rows)")

grid("excerpt map state (no transcript text here)", spark.sql(f"""
    SELECT vintage, wave, batch, count(*) AS excerpts,
           count(DISTINCT acct_key) AS accounts
    FROM {DB}.uc2_copilot_excerpt_map
    GROUP BY 1, 2, 3 ORDER BY 1, 2, 3
"""))

# COMMAND ----------

# MAGIC %md
# MAGIC ## L3. PASTE CELL - one response file per run of L3 -> L4 -> L5
# MAGIC
# MAGIC Fill the five constants and paste the WHOLE raw response file between the
# MAGIC triple quotes, verbatim (do not edit, trim, or fix anything). Then run
# MAGIC L4 and L5. Repeat per response file. Re-ingesting the same RESPONSE_FILE
# MAGIC replaces its rows (idempotent).

# COMMAND ----------

sec("L3 paste cell")

WAVE = 1
BATCH = 1                                  # the batch number the response answers
PROMPT_TYPE = "discovery"                  # discovery | classify | second-read
RESPONSE_FILE = "copilot-response-batch1-roundA-2026-07-16.txt"   # the saved file's name, verbatim
OPERATOR = ""                              # who ran the UI session (initials)

RESPONSE_TEXT = r"""
(PASTE THE RAW RESPONSE HERE - the whole saved file, verbatim)
"""

_HAVE_RESPONSE = "(PASTE THE RAW RESPONSE HERE" not in RESPONSE_TEXT and RESPONSE_TEXT.strip() != ""
if _HAVE_RESPONSE:
    print(f"response loaded: wave {WAVE} batch {BATCH} type {PROMPT_TYPE} "
          f"file {RESPONSE_FILE} ({len(RESPONSE_TEXT):,} chars)")
else:
    print("NO RESPONSE PASTED YET - L4/L5 will no-op. Fill this cell and rerun L3 -> L5.")

# COMMAND ----------

# MAGIC %md
# MAGIC ## L4. Parse (format-locked; never repairs content)
# MAGIC
# MAGIC classify / second-read expect the Round-B format per excerpt:
# MAGIC   EXCERPT <id>:  then lines  - intent: ...  - agent_attempt: ...
# MAGIC   - blocker: <taxonomy token>  - convert_action: ...  - phrases: ...
# MAGIC discovery expects the Round-A sections: PHRASES: / CATEGORIES: / ASR:
# MAGIC with one bullet per proposal.
# MAGIC
# MAGIC Anything else lands as a FLAGGED '_unparsed' row (raw line verbatim) for
# MAGIC human adjudication. Blockers outside the frozen taxonomy, intent or
# MAGIC agent_attempt lines that do not start yes/no, duplicate fields, and ids
# MAGIC outside the wave's map are FLAGGED, never dropped, never corrected.

# COMMAND ----------

sec("L4 parse")

import re

BLOCKER_TAXONOMY = [
    "agent-did-not-ask", "no-payment-method-available",
    "authentication-or-authority-block", "offer-not-eligible-yet",
    "customer-deferred", "dispute", "none-visible",
]
CLASSIFY_FIELDS = ["intent", "agent_attempt", "blocker", "convert_action", "phrases"]
SUPPORTED_TYPES = ("discovery", "classify", "second-read")

_excerpt_re = re.compile(r'^[\s>*#]*EXCERPT\s+([A-Z]\d{1,2})\b', re.IGNORECASE)
_field_re = re.compile(
    r'^[\s>*\-]*(intent|agent[\s_]?attempt|blocker|convert[\s_]?action|phrases)'
    r'[\s*]*[:\-]\s*(.*)$', re.IGNORECASE)
_yesno_re = re.compile(r'^\W*(yes|no)\b', re.IGNORECASE)
_section_re = re.compile(r'^[\s>*#]*(PHRASES|CATEGORIES|ASR)\b[\s:*\-]*$', re.IGNORECASE)
_bullet_re = re.compile(r'^[\s>]*[-*]\s+(.*)$')


def _norm_blocker(v):
    """Normalize a blocker value and match it against the frozen taxonomy.
    Tokens match on WORD BOUNDARIES of the normalized value (so 'customer
    disputes the late fee' does NOT auto-resolve to the dispute token; it
    gets flagged for a human read instead - flag, never fix). Returns
    (token, 'ok') on exactly one match, (raw, 'flagged') otherwise."""
    norm = '-' + re.sub(r'[^a-z0-9]+', '-', v.lower()).strip('-') + '-'
    hits = sorted({t for t in BLOCKER_TAXONOMY if f'-{t}-' in norm})
    if len(hits) == 1:
        return hits[0], "ok"
    return v.strip(), "flagged"


def parse_classify(text):
    """Round-B format -> per-excerpt field rows. Returns (rows, unparsed)."""
    rows, unparsed = [], []
    cur_id = None
    open_row = None
    for raw in text.splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        m = _excerpt_re.match(line)
        if m:
            cur_id = m.group(1).upper()
            open_row = None
            continue
        m = _field_re.match(line)
        if m and cur_id:
            field = re.sub(r'[\s]+', '_', m.group(1).strip().lower())
            open_row = {"excerpt_id": cur_id, "field": field,
                        "value": m.group(2).strip(), "status": "ok", "raw_line": line}
            rows.append(open_row)
            continue
        if open_row is not None:
            open_row["value"] = (open_row["value"] + " " + line.strip()).strip()
            continue
        unparsed.append(line)
    # validation (flag, never fix)
    seen = set()
    for r in rows:
        key = (r["excerpt_id"], r["field"])
        if key in seen:
            r["status"] = "flagged"     # duplicate field for the excerpt
        seen.add(key)
        if r["field"] == "blocker":
            r["value"], s = _norm_blocker(r["value"])
            r["status"] = s if r["status"] == "ok" else r["status"]
        elif r["field"] in ("intent", "agent_attempt"):
            if not _yesno_re.match(r["value"]):
                r["status"] = "flagged"
        if not r["value"]:
            r["status"] = "flagged"
    return rows, unparsed


def parse_discovery(text):
    """Round-A format (PHRASES/CATEGORIES/ASR sections) -> proposal rows."""
    rows, unparsed = [], []
    secmap = {"PHRASES": "phrase", "CATEGORIES": "category", "ASR": "asr"}
    section = None
    open_row = None
    for raw in text.splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue
        m = _section_re.match(line)
        if m:
            section = secmap[m.group(1).upper()]
            open_row = None
            continue
        m = _bullet_re.match(line)
        if m and section:
            open_row = {"excerpt_id": None, "field": section,
                        "value": m.group(1).strip(), "status": "ok", "raw_line": line}
            rows.append(open_row)
            continue
        if open_row is not None:
            open_row["value"] = (open_row["value"] + " " + line.strip()).strip()
            continue
        unparsed.append(line)
    return rows, unparsed


parsed_rows, unparsed_lines = [], []
if not _HAVE_RESPONSE:
    print("no response loaded - nothing to parse")
else:
    assert PROMPT_TYPE in SUPPORTED_TYPES, \
        f"unsupported PROMPT_TYPE '{PROMPT_TYPE}' (contrast/rubric ingest lands with wave 2)"
    if PROMPT_TYPE == "discovery":
        parsed_rows, unparsed_lines = parse_discovery(RESPONSE_TEXT)
    else:
        parsed_rows, unparsed_lines = parse_classify(RESPONSE_TEXT)
        assert parsed_rows, "PARSE MISS: zero excerpt fields parsed from a classify response - " \
                            "check the paste and the response format before anything is written"
        # coverage against the wave's map (ids outside the map are flagged)
        _map_ids = {r["excerpt_id"] for r in spark.sql(f"""
            SELECT excerpt_id FROM {DB}.uc2_copilot_excerpt_map
            WHERE wave = {WAVE} AND batch = {BATCH} AND vintage = '{ANCHOR_YM}'
        """).collect()}
        _got_ids = {r["excerpt_id"] for r in parsed_rows}
        for r in parsed_rows:
            if _map_ids and r["excerpt_id"] not in _map_ids:
                r["status"] = "flagged"
        if _map_ids:
            print(f"batch coverage: {len(_got_ids & _map_ids)} of {len(_map_ids)} mapped ids answered; "
                  f"missing: {sorted(_map_ids - _got_ids) or 'none'}; "
                  f"outside the map: {sorted(_got_ids - _map_ids) or 'none'}")
        else:
            print(f"NOTE: no map rows for wave {WAVE} batch {BATCH} - id coverage not checkable")
    chk("parsed rows", len(parsed_rows), None)
    chk("unparsed lines (flagged for adjudication)", len(unparsed_lines), None)
    chk("flagged rows", sum(1 for r in parsed_rows if r["status"] == "flagged"), None)
    if unparsed_lines:
        print("UNPARSED LINES (verbatim, for the judge):")
        for _l in unparsed_lines:
            print(f"  ! {_l}")

# COMMAND ----------

# MAGIC %md
# MAGIC ## L5. Write (idempotent per response file)

# COMMAND ----------

sec("L5 write")

if not _HAVE_RESPONSE:
    print("no response loaded - nothing to write")
else:
    assert "'" not in RESPONSE_FILE, "RESPONSE_FILE must not contain quotes"
    _stage = (
        [(WAVE, BATCH, PROMPT_TYPE, r["excerpt_id"], r["field"], r["value"],
          RESPONSE_FILE, OPERATOR, r["status"], r["raw_line"]) for r in parsed_rows]
        + [(WAVE, BATCH, PROMPT_TYPE, None, "_unparsed", _l,
            RESPONSE_FILE, OPERATOR, "flagged", _l) for _l in unparsed_lines]
    )
    spark.createDataFrame(
        _stage,
        "wave int, batch int, prompt_type string, excerpt_id string, field string, "
        "value string, response_file string, operator string, parse_status string, "
        "raw_line string"
    ).createOrReplaceTempView("_labels_stage")
    spark.sql(f"""
        DELETE FROM {DB}.uc2_copilot_labels
        WHERE response_file = '{RESPONSE_FILE}' AND vintage = '{ANCHOR_YM}'
    """)
    spark.sql(f"""
        INSERT INTO {DB}.uc2_copilot_labels
        SELECT current_timestamp(), '{ANCHOR_YM}', * FROM _labels_stage
    """)
    chk("label rows written for this response file", len(_stage), None)

    if PROMPT_TYPE in ("classify", "second-read"):
        grid("per-excerpt field coverage (this response file)", spark.sql(f"""
            SELECT excerpt_id,
                   count(*) AS fields,
                   count_if(parse_status = 'flagged') AS flagged,
                   max(CASE WHEN field = 'blocker' AND parse_status = 'ok'
                            THEN value END) AS blocker
            FROM {DB}.uc2_copilot_labels
            WHERE response_file = '{RESPONSE_FILE}' AND vintage = '{ANCHOR_YM}'
              AND field <> '_unparsed'
            GROUP BY 1 ORDER BY 1
        """))
    else:
        grid("proposal counts by section (this response file)", spark.sql(f"""
            SELECT field, count(*) AS proposals,
                   count_if(parse_status = 'flagged') AS flagged
            FROM {DB}.uc2_copilot_labels
            WHERE response_file = '{RESPONSE_FILE}' AND vintage = '{ANCHOR_YM}'
            GROUP BY 1 ORDER BY 1
        """))

# COMMAND ----------

# MAGIC %md
# MAGIC ## L6. Campaign state (reads the labels table; safe to run any time)
# MAGIC
# MAGIC Rates come with a 95% Wilson interval next to the n - the credibility
# MAGIC floor: no rate is ever quoted without its n, interval, and stratum.
# MAGIC These are WORKING views; record numbers still come from a verified
# MAGIC screenshot of this section plus the round record.

# COMMAND ----------

sec("L6 campaign state")

_n = spark.sql(f"SELECT count(*) AS n FROM {DB}.uc2_copilot_labels").first()["n"]
chk("labels table rows (all files)", _n, None)
if _n == 0:
    print("labels table empty - come back after the first ingestion")
else:
    grid("ingested response files", spark.sql(f"""
        SELECT response_file, prompt_type, wave, batch, count(*) AS rows,
               count_if(parse_status = 'flagged') AS flagged
        FROM {DB}.uc2_copilot_labels
        GROUP BY 1, 2, 3, 4 ORDER BY 3, 4, 2, 1
    """))
    _nb = spark.sql(f"""
        SELECT count(*) AS n FROM {DB}.uc2_copilot_labels
        WHERE prompt_type = 'classify' AND field = 'blocker' AND parse_status = 'ok'
    """).first()["n"]
    if _nb == 0:
        print("no clean classify blocker rows yet - the blocker split appears after Round B lands")
    else:
        grid("blocker split by stratum (95% Wilson interval; SAMPLE rates, not population)",
             spark.sql(f"""
            WITH b AS (
                SELECT m.stratum, l.value AS blocker
                FROM {DB}.uc2_copilot_labels l
                JOIN {DB}.uc2_copilot_excerpt_map m
                  ON m.wave = l.wave AND m.excerpt_id = l.excerpt_id
                 AND m.vintage = l.vintage
                WHERE l.prompt_type = 'classify' AND l.field = 'blocker'
                  AND l.parse_status = 'ok'
            ),
            n AS (SELECT stratum, count(*) AS n FROM b GROUP BY 1)
            SELECT b.stratum, b.blocker, count(*) AS k, max(n.n) AS n,
                   round(count(*) / max(n.n), 3) AS share,
                   round(((count(*) / max(n.n)) + 1.9208 / max(n.n)
                          - 1.96 * sqrt((count(*) / max(n.n)) * (1 - count(*) / max(n.n)) / max(n.n)
                                        + 0.9604 / (max(n.n) * max(n.n))))
                         / (1 + 3.8416 / max(n.n)), 3) AS wilson_lo,
                   round(((count(*) / max(n.n)) + 1.9208 / max(n.n)
                          + 1.96 * sqrt((count(*) / max(n.n)) * (1 - count(*) / max(n.n)) / max(n.n)
                                        + 0.9604 / (max(n.n) * max(n.n))))
                         / (1 + 3.8416 / max(n.n)), 3) AS wilson_hi
            FROM b JOIN n ON n.stratum = b.stratum
            GROUP BY b.stratum, b.blocker
            ORDER BY 1, 3 DESC, 2
        """))
    _agree = spark.sql(f"""
        SELECT count(*) AS paired,
               count_if(a.value = b.value) AS agree
        FROM (SELECT vintage, wave, excerpt_id, value FROM {DB}.uc2_copilot_labels
              WHERE prompt_type = 'classify' AND field = 'blocker' AND parse_status = 'ok') a
        JOIN (SELECT vintage, wave, excerpt_id, value FROM {DB}.uc2_copilot_labels
              WHERE prompt_type = 'second-read' AND field = 'blocker' AND parse_status = 'ok') b
          ON a.vintage = b.vintage AND a.wave = b.wave AND a.excerpt_id = b.excerpt_id
    """).first()
    if _agree["paired"] == 0:
        print("no first-read/second-read pairs yet - agreement appears after the 20% re-read")
    else:
        chk("second-read pairs", _agree["paired"], None)
        chk("second-read agreements", _agree["agree"], None)
        print(f"agreement rate = {_agree['agree'] / _agree['paired']:.3f} "
              f"(report NEXT TO every rate estimate)")

# COMMAND ----------

sec("L7 verdict")

record_block("B06_copilot_labels_ingest")
flush_metrics("B06_copilot_labels_ingest")
