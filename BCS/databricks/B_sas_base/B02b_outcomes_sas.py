# Databricks notebook source
# MAGIC %md
# MAGIC # B02b. Outcomes on the SAS spine: `uc2_t16_04s_outcomes_<vintage>`
# MAGIC
# MAGIC Grain: one row per standard January episode (fixed numeric key) on an
# MAGIC account inside the SAS ledger (in_sas_ledger, 186,013 for 202501).
# MAGIC
# MAGIC The headline gate is captured_sas: ACCOUNT grain, month grain (payment in
# MAGIC the call month or the next, CQ-7 sign convention). There is NO
# MAGIC episode-grain "captured" under this gate; classes below are account-level
# MAGIC by construction. The AWS day-grain gate rides along as aws_ diagnostic
# MAGIC columns only, never a denominator.
# MAGIC
# MAGIC Account-level classes on captured_sas:
# MAGIC   caller_class_sas: b. captured / c. leaked-intent (payment language on
# MAGIC   >= 1 episode, no account payment in M1/M2) / d. other-caller
# MAGIC   ('a. non-caller' lives on the 01s side).
# MAGIC   leaked_sas = NOT captured_sas AND >= 1 payment-language episode.
# MAGIC   W_s = leaked_sas AND non-deceased (in_sas_ledger by construction here).

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
# MAGIC ## O1. Preconditions

# COMMAND ----------

sec("O1 preconditions")

for _t in ["uc2_t16_02n_episodes", "uc2_t16_04n_outcomes", f"uc2_t16_01s_populations_{ANCHOR_YM}"]:
    assert spark.catalog.tableExists(f"{DB}.{_t}"), \
        f"PRECONDITION MISS: {DB}.{_t} missing - run B02 then B01 first"
    print(f"PASS  table exists: {DB}.{_t}")

_r = spark.sql(f"SELECT count_if(in_sas_ledger) AS n FROM {DB}.uc2_t16_01s_populations_{ANCHOR_YM}").first()
chk("wf 04 sas ledger", _r["n"], E["wf 04 sas ledger"])

# COMMAND ----------

# MAGIC %md
# MAGIC ## O2. Build `uc2_t16_04s_outcomes_<vintage>`

# COMMAND ----------

sec("O2 build 04s")

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_04s_outcomes_{ANCHOR_YM} AS
WITH s AS (
    SELECT acct_key, captured_sas, dlnqt_cd_m2, dlnqt_cd_m3, stg_cd_m1, stg_cd_m2,
           cpc_flag_nw, eop_bal_m1, eop_bal_m2, ecl_m1, ecl_m2, ecl_liftm_m1,
           gross_loss_12m_amt, chrgoff_12m_amt, gross_loss_amt_lftm, chrgoff_amt_lftm
    FROM {DB}.uc2_t16_01s_populations_{ANCHOR_YM}
    WHERE in_sas_ledger
),
j AS (
    SELECT e.acct_key, e.contactid, e.call_dt,
           e.language_group, e.pay_f, e.deceased_f, e.exec_f, e.has_tx,
           e.callday_bucket, e.is_addressable,
           e.captured    AS aws_captured_episode,   -- diagnostic (AWS day-grain), never a denominator
           e.any_captured = 1 AS aws_captured,      -- diagnostic (AWS day-grain), never a denominator
           e.caller_class AS aws_caller_class,      -- diagnostic label (AWS day-grain classes)
           s.captured_sas,
           s.dlnqt_cd_m2, s.dlnqt_cd_m3, s.stg_cd_m1, s.stg_cd_m2, s.cpc_flag_nw,
           s.eop_bal_m1, s.eop_bal_m2, s.ecl_m1, s.ecl_m2, s.ecl_liftm_m1,
           s.gross_loss_12m_amt, s.chrgoff_12m_amt, s.gross_loss_amt_lftm, s.chrgoff_amt_lftm
    FROM {DB}.uc2_t16_04n_outcomes e
    JOIN s ON s.acct_key = e.acct_key
),
acct AS (
    SELECT acct_key,
           max(pay_f) AS any_pay,
           max(deceased_f) AS deceased_acct
    FROM j GROUP BY 1
)
SELECT j.*,
       a.deceased_acct,
       CASE
         WHEN j.captured_sas THEN 'b. captured (account payment in call month or next)'
         WHEN a.any_pay > 0 THEN 'c. leaked-intent (payment language, no account payment M1/M2)'
         ELSE 'd. other-caller'
       END AS caller_class_sas,
       (NOT j.captured_sas AND a.any_pay > 0) AS leaked_sas,
       (NOT j.captured_sas AND a.any_pay > 0 AND a.deceased_acct = 0) AS w_s_flag
FROM j
JOIN acct a ON a.acct_key = j.acct_key
""")
print(f"built {DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}")

# COMMAND ----------

# MAGIC %md
# MAGIC ## O3. Measured summary + the provable-direction consistency check
# MAGIC
# MAGIC 04s callers ⊇ (new callers ∩ CSV-flagged): flagged is a subset of
# MAGIC in_sas_ledger, so the containment is provable and asserted. EQUALITY is
# MAGIC NOT provable (our episode rows are not provably a subset of the export's
# MAGIC row universe), so the residual (04s callers the CSV flag missed) is
# MAGIC MEASURED with named per-account causes; STOP only on unexplained rows.

# COMMAND ----------

sec("O3 measured summary")

_r = spark.sql(f"""
    SELECT count(*) AS episodes,
           count(DISTINCT acct_key) AS callers,
           count(DISTINCT CASE WHEN captured_sas THEN acct_key END) AS cap_accts,
           count(DISTINCT CASE WHEN leaked_sas THEN acct_key END) AS leaked_accts,
           count(DISTINCT CASE WHEN w_s_flag THEN acct_key END) AS ws_accts
    FROM {DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}
""").first()
chk("04s episodes", _r["episodes"], E["04s episodes"])
chk("04s callers", _r["callers"], E["04s callers"])
chk("04s captured_sas accounts", _r["cap_accts"], E["04s captured_sas accounts"])
chk("04s leaked_sas accounts", _r["leaked_accts"], E["04s leaked_sas accounts"])
chk("04s W_s accounts", _r["ws_accts"], E["04s W_s accounts"])

if ANCHOR_YM == "202501" and spark.catalog.tableExists(f"{DB}.uc2_sasflag_202501"):
    # containment (provable direction): every CSV-flagged account that is a
    # numeric-keyed caller must appear among the 04s callers
    _missing_df = spark.sql(f"""
        WITH new_callers AS (SELECT DISTINCT acct_key FROM {DB}.uc2_t16_04n_outcomes WHERE in_ledger_exaa),
        overlap AS (SELECT f.acct_key FROM {DB}.uc2_sasflag_202501 f JOIN new_callers n ON n.acct_key = f.acct_key),
        callers_04s AS (SELECT DISTINCT acct_key FROM {DB}.uc2_t16_04s_outcomes_{ANCHOR_YM})
        SELECT o.acct_key FROM overlap o LEFT ANTI JOIN callers_04s c ON c.acct_key = o.acct_key
    """)
    chk("04s containment: flagged-overlap callers missing from 04s (must be zero)",
        _missing_df.count(), 0, ctx=_missing_df)

    # the residual: 04s callers the CSV flag missed, classified by what the
    # CSV knows about them (import window/universe and flag spelling)
    grid("04s callers not CSV-flagged: causes", spark.sql(f"""
        WITH callers_04s AS (SELECT DISTINCT acct_key FROM {DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}),
        residual AS (
            SELECT c.acct_key FROM callers_04s c
            LEFT ANTI JOIN {DB}.uc2_sasflag_202501 f ON f.acct_key = c.acct_key
        )
        SELECT CASE
                 WHEN w.call_type_INB IS NULL OR trim(w.call_type_INB) = ''
                   THEN 'a. CSV INB flag blank (call outside the SAS import window/universe)'
                 ELSE concat('b. CSV INB flag reads: ', w.call_type_INB)
               END AS cause,
               count(*) AS accounts
        FROM residual r
        JOIN {DB}.uc2_sas_wf_202501 w ON w.acct_key = r.acct_key
        GROUP BY 1 ORDER BY 1
    """))
    print("READ RULE: cause a. is expected (the import's own window/universe);")
    print("any other cause is investigated before this table is used downstream.")
else:
    print("consistency checks skipped (not 202501, or notebook A's tables absent)")

# the re-baseline size: account-grain cross-tab of the two gates (disclosed;
# their ONLY other meeting point is notebook A's delta table)
grid("account-grain cross-tab: captured_sas x aws_captured (04s callers)", spark.sql(f"""
    SELECT captured_sas, aws_captured, count(DISTINCT acct_key) AS accounts
    FROM {DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}
    GROUP BY 1, 2 ORDER BY 1, 2
"""))

# COMMAND ----------

# MAGIC %md
# MAGIC ## O4. Verdict and record block

# COMMAND ----------

sec("O4 verdict")

record_block("B02b_outcomes_sas")
flush_metrics("B02b_outcomes_sas")
