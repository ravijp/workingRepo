# Databricks notebook source
# MAGIC %md
# MAGIC # B02. The key fix: rebuild the AWS layers on the numeric key (runs FIRST)
# MAGIC
# MAGIC WHAT CHANGES: one line - the account key derivation moves from
# MAGIC `trim(cast(id AS string))` to `cast(try_cast(id AS bigint) AS string)`
# MAGIC in every layer. Everything else is the verified tier-16 logic verbatim.
# MAGIC
# MAGIC WHAT MUST NOT CHANGE: the population anchors (204,323 / 189,146 /
# MAGIC $457,943,987 +/- $5 / 724,848 and the class split). The fmt-side probe
# MAGIC (K2) proves the key change is a no-op on the population side; any anchor
# MAGIC drift = STOP, no downstream number is read.
# MAGIC
# MAGIC WHAT IS MEASURED, NEVER ASSERTED AGAINST OLD VALUES: every caller-side
# MAGIC number. The old 9,389 / 11,262 / 29,114 are historical references from the
# MAGIC string-keyed build. Pre-registered implication checks replace them:
# MAGIC dropped_old = 0 exactly; episodes and the call-day stream can only grow.
# MAGIC
# MAGIC OUTPUT TABLES (new names; the round-10 uc2_t16_00..04 tables stay frozen
# MAGIC as the dated record's substrate): uc2_t16_00n_acct_monthly,
# MAGIC uc2_t16_01n_populations, uc2_t16_02n_episodes, uc2_t16_03n_signals,
# MAGIC uc2_t16_04n_outcomes.
# MAGIC
# MAGIC DISCLOSED DEVIATION (D8): the 02n call scan adds an effdt bound
# MAGIC [EFFDT_SCAN_START, 2026-07-10) plus REFRESH TABLE (the round-11 live
# MAGIC loading edge lesson). Standard episodes require the stricter in-column
# MAGIC cap, so no anchor can move; only never-quoted diagnostic rows differ.

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
# MAGIC ## K1. Preconditions

# COMMAND ----------

sec("K1 preconditions")

# sources reachable
for _t in [f"{FMT_CATALOG}.fmt_acct_dba.fmt_acct_c",
           f"{CC_CATALOG}.contactcenter_bdp_db.call",
           f"{CC_CATALOG}.contactcenter_bdp_db.transcript"]:
    assert spark.catalog.tableExists(_t), f"PRECONDITION MISS: source {_t} not reachable - fix the catalog widget"
    print(f"PASS  source reachable: {_t}")

# the round-10 string-keyed tables (needed for the old-caller set)
for _t in ["uc2_t16_01_populations", "uc2_t16_02_episodes"]:
    assert spark.catalog.tableExists(f"{DB}.{_t}"), f"PRECONDITION MISS: round-10 table {DB}.{_t} missing"
    print(f"PASS  round-10 table exists: {DB}.{_t}")

# notebook A's persisted reconciliation tables (202501 sitting only)
if ANCHOR_YM == "202501":
    for _t in ["uc2_gap1942_202501", "uc2_sasflag_202501"]:
        assert spark.catalog.tableExists(f"{DB}.{_t}"), \
            f"PRECONDITION MISS: {DB}.{_t} missing - run A_recon_lock_202501 first"
        print(f"PASS  recon table exists: {DB}.{_t}")

# COMMAND ----------

# MAGIC %md
# MAGIC ## K2. fmt-side key probe (BEFORE any key changes)
# MAGIC
# MAGIC Expectation (pre-registered): ZERO mismatches - the account snapshot
# MAGIC stores unpadded ids, so the numeric key is a provable no-op on the
# MAGIC population side. That proof is why the population anchors stay
# MAGIC assert-exact through the rebuild. Any miss here = STOP.

# COMMAND ----------

sec("K2 fmt key probe")

_r = spark.sql(f"""
    SELECT count_if(extnl_acct_id IS NOT NULL AND try_cast(extnl_acct_id AS bigint) IS NULL) AS non_castable,
           count_if(extnl_acct_id IS NOT NULL AND try_cast(extnl_acct_id AS bigint) IS NOT NULL
                    AND trim(cast(extnl_acct_id AS string)) <> {NUM_KEY.format(c="extnl_acct_id")}) AS pad_mismatch,
           count_if(extnl_acct_id IS NULL) AS null_ids
    FROM {FMT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '{MONTH_WIN_START}' AND eff_dt < '{MONTH_WIN_END}'
""").first()
chk("fmt id non-castable (00 window)", _r["non_castable"], 0)
chk("fmt id pad-mismatch (00 window)", _r["pad_mismatch"], 0)
chk("fmt id null (00 window)", _r["null_ids"], None)   # context, expected tiny

# COMMAND ----------

# MAGIC %md
# MAGIC ## K3. Build `uc2_t16_00n_acct_monthly` (the expensive scan)
# MAGIC
# MAGIC Tier-16 layer 00 verbatim (Spark translations per the kit README section
# MAGIC 4); the ONLY logic change is the acct_key derivation (the numeric rule).
# MAGIC try_to_date is the Spark stand-in for the Trino try(date_parse(...))
# MAGIC payment-date fallback; if the runtime lacks it, use to_date with
# MAGIC spark.sql.ansi.enabled=false (the round-10 pattern) - one-line edit.

# COMMAND ----------

sec("K3 build 00n")

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_00n_acct_monthly AS
WITH snap AS (
    SELECT extnl_acct_id,
           substr(eff_dt, 1, 6) AS ym,
           eff_dt,
           CASE
             WHEN past_due_271_up_amt  > 0 THEN 10
             WHEN past_due_241_270_amt > 0 THEN 9
             WHEN past_due_211_240_amt > 0 THEN 8
             WHEN past_due_181_210_amt > 0 THEN 7
             WHEN past_due_151_180_amt > 0 THEN 6
             WHEN past_due_121_150_amt > 0 THEN 5
             WHEN past_due_91_120_amt  > 0 THEN 4
             WHEN past_due_61_90_amt   > 0 THEN 3
             WHEN past_due_31_60_amt   > 0 THEN 2
             WHEN past_due_1_30_amt    > 0 THEN 1
             ELSE 0
           END AS bucket,
           try_cast(acct_bal_amt AS double) AS bal,
           try_cast(chrgoff_dt AS date) AS co_dt,
           try_cast(chrgoff_amt AS double) AS co_amt,
           clnt_prdct_cd,
           try_cast(cr_lmt_origl_amt AS double) AS cr_lmt_origl_amt,
           coalesce(try_cast(paymt_last_dt AS date),
                    try_to_date(cast(paymt_last_dt AS string), 'ddMMMyyyy')) AS pay_dt,
           coalesce(try_cast(atmtc_paymt_last_dt AS date),
                    try_to_date(cast(atmtc_paymt_last_dt AS string), 'ddMMMyyyy')) AS auto_dt,
           coalesce(try_cast(nsf_last_paymt_dt AS date),
                    try_to_date(cast(nsf_last_paymt_dt AS string), 'ddMMMyyyy')) AS nsf_dt
    FROM {FMT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '{MONTH_WIN_START}' AND eff_dt < '{MONTH_WIN_END}'
)
SELECT {NUM_KEY.format(c="extnl_acct_id")} AS acct_key,   -- THE KEY CHANGE (D2); grain is the numeric key
       ym,
       max(bucket) AS max_bucket,
       max_by(bucket, eff_dt) AS eom_bucket,
       max_by(bal, eff_dt) AS eom_bal,
       min(co_dt) AS mth_co_dt,
       min_by(co_amt, co_dt) AS mth_co_amt,
       min(CASE WHEN bucket >= 1 THEN eff_dt END) AS first_dq_dt,
       min(CASE WHEN bucket = 1 THEN eff_dt END) AS first_b1_dt,
       max_by(clnt_prdct_cd, eff_dt) AS eom_cpc,
       max_by(cr_lmt_origl_amt, eff_dt) AS eom_cr_lmt_origl_amt,
       max(pay_dt) AS pay_dt,
       max(auto_dt) AS auto_dt,
       max(nsf_dt) AS nsf_dt
FROM snap
GROUP BY 1, 2
""")
print(f"built {DB}.uc2_t16_00n_acct_monthly")

# COMMAND ----------

# MAGIC %md
# MAGIC ## K4. Build `uc2_t16_01n_populations` + THE POPULATION ANCHOR SWEEP
# MAGIC
# MAGIC Tier-16 layer 01 verbatim on 00n. Every anchor below asserts EXACTLY;
# MAGIC any drift = STOP, nothing downstream is read.

# COMMAND ----------

sec("K4 build 01n")

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_01n_populations AS
WITH jan AS (SELECT * FROM {DB}.uc2_t16_00n_acct_monthly WHERE ym = '{ANCHOR_YM}'),
prv AS (SELECT * FROM {DB}.uc2_t16_00n_acct_monthly WHERE ym = '{PRV_YM}'),
feb AS (SELECT * FROM {DB}.uc2_t16_00n_acct_monthly WHERE ym = '{FEB_YM}'),
mar AS (SELECT * FROM {DB}.uc2_t16_00n_acct_monthly WHERE ym = '{MAR_YM}'),
future_co AS (
    SELECT {NUM_KEY.format(c="extnl_acct_id")} AS acct_key,
           min(try_cast(chrgoff_dt AS date)) AS co_dt_future,
           min_by(try_cast(chrgoff_amt AS double), try_cast(chrgoff_dt AS date)) AS co_amt
    FROM {FMT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '{FWD_CO_START}' AND eff_dt < '{FWD_CO_END}'
      AND chrgoff_dt IS NOT NULL
      AND {NUM_KEY.format(c="extnl_acct_id")} IN (SELECT acct_key FROM jan WHERE max_bucket >= 1)
    GROUP BY 1
),
pop_base AS (
    SELECT j.acct_key,
           j.max_bucket, j.eom_bucket, j.eom_bal, j.mth_co_dt AS jan_co_dt,
           j.first_dq_dt, j.first_b1_dt, j.eom_cpc, j.eom_cr_lmt_origl_amt,
           p.max_bucket AS prev_max_bucket,
           p.eom_bucket AS prev_eom_bucket,
           (j.mth_co_dt IS NULL OR j.mth_co_dt >= DATE '{CLEANUP_DATE}') AS cleaned,
           (j.eom_cpc IS NULL OR trim(j.eom_cpc) = ''
            OR j.eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                                  'AA3','AC3','AM3','AA4','AC4','AM4',
                                  'BGC','BGM','CGM','GMR',
                                  'FBS','IBS','U1C','U2C','U3C')) AS is_exaa,
           CASE
             WHEN j.eom_cpc IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                                 'AA3','AC3','AM3','AA4','AC4','AM4')     THEN 'AA'
             WHEN j.eom_cpc IN ('BGC','BGM','CGM','GMR')                 THEN 'GM'
             WHEN j.eom_cpc IN ('FBS','IBS','U1C','U2C','U3C')           THEN 'Bronco'
             WHEN j.eom_cpc IN ('BHA','BJT','BJC','BFR','BWY','BBB')     THEN 'Biz'
             WHEN j.eom_cpc IN ('GAP','GP2','ONV','ON2','BRP','BR2','ATH','AT2',
                                 'GPC','G2C','ONC','O2C','BRC','B2C','ATC','A2C')
                                                                          THEN 'CoBrand'
             WHEN j.eom_cpc IN ('8GP','8ON','8BR','8AT','9GP','9ON','9BR','9AT')
                                                                          THEN 'PLCC'
             ELSE 'OTHER'
           END AS cpc_class,
           CASE
             WHEN coalesce(p.eom_bucket, 0) >= 1
               THEN 'd. carried-in (past due at Dec-31 EOM)'
             WHEN j.first_dq_dt IS NULL THEN NULL
             WHEN cast(substr(j.first_dq_dt, 7, 2) AS int) <= 10
               THEN 'a. runway >= 21 days (entry day 1-10)'
             WHEN cast(substr(j.first_dq_dt, 7, 2) AS int) <= 20
               THEN 'b. runway 11-20 days (entry day 11-20)'
             ELSE 'c. runway <= 10 days (entry day 21-31)'
           END AS runway_band,
           CASE
             WHEN f.mth_co_dt >= DATE '{FEB_START}' AND f.mth_co_dt < DATE '{MAR_START}'
               THEN 'e. charged off in Feb'
             WHEN f.acct_key IS NULL THEN 'f. no Feb row'
             WHEN f.eom_bucket = 0 THEN 'a. Feb EOM bucket 0 (cured)'
             WHEN f.eom_bucket = 1 THEN 'b. Feb EOM bucket 1 (stayed)'
             WHEN f.eom_bucket = 2 THEN 'c. Feb EOM bucket 2 (rolled)'
             ELSE 'd. Feb EOM bucket 3+ (rolled deeper)'
           END AS feb_position_b14,
           CASE
             WHEN f.mth_co_dt >= DATE '{FEB_START}' AND f.mth_co_dt < DATE '{MAR_START}' THEN 'co'
             WHEN f.acct_key IS NULL THEN 'gone'
             ELSE cast(f.eom_bucket AS string)
           END AS feb_pos,
           CASE
             WHEN m.mth_co_dt >= DATE '{FEB_START}' AND m.mth_co_dt < DATE '{APR_START}' THEN 'co'
             WHEN m.acct_key IS NULL THEN 'gone'
             ELSE cast(m.eom_bucket AS string)
           END AS mar_pos,
           fc.co_dt_future,
           fc.co_amt,
           (fc.co_dt_future >= DATE '{ANCHOR_EOM}' AND fc.co_dt_future < DATE '{CO8_END}')  AS co_8m,
           (fc.co_dt_future >= DATE '{ANCHOR_EOM}' AND fc.co_dt_future < DATE '{CO10_END}') AS co_10m,
           (fc.co_dt_future >= DATE '{ANCHOR_EOM}' AND fc.co_dt_future < DATE '{CO12_END}') AS co_12m
    FROM jan j
    LEFT JOIN prv p ON p.acct_key = j.acct_key
    LEFT JOIN feb f ON f.acct_key = j.acct_key
    LEFT JOIN mar m ON m.acct_key = j.acct_key
    LEFT JOIN future_co fc ON fc.acct_key = j.acct_key
)
SELECT *,
       (eom_bucket = 1 AND cleaned)             AS in_ledger_all,
       (eom_bucket = 1 AND cleaned AND is_exaa) AS in_ledger_exaa,
       (first_b1_dt IS NOT NULL AND cleaned AND is_exaa) AS touched_b1,
       CASE
         WHEN NOT (first_b1_dt IS NOT NULL AND cleaned AND is_exaa) THEN NULL
         WHEN jan_co_dt >= DATE '{CLEANUP_DATE}' AND jan_co_dt < DATE '{FEB_START}'
           THEN 'd. charged off in January'
         WHEN eom_bucket = 0
           THEN 'a. current at 31 Jan (cured in month)'
         WHEN eom_bucket = 1
           THEN 'b. bucket 1 at 31 Jan'
         WHEN eom_bucket >= 2
           THEN 'c. bucket 2+ at 31 Jan (rolled past DQ1 within January)'
       END AS touched_b1_class
FROM pop_base
""")
print(f"built {DB}.uc2_t16_01n_populations")

# THE POPULATION ANCHOR SWEEP - all assert-exact
_r = spark.sql(f"""
    SELECT count(*) AS rows, count(DISTINCT acct_key) AS accts,
           count_if(in_ledger_all)                                        AS ledger_all,
           count_if(in_ledger_all AND cpc_class = 'AA')                   AS ledger_aa,
           count_if(in_ledger_exaa)                                       AS ledger_exaa,
           round(sum(CASE WHEN in_ledger_exaa THEN eom_bal END), 0)       AS exaa_bal,
           count_if(touched_b1)                                           AS touched_b1,
           count_if(touched_b1_class LIKE 'a.%')                          AS t_a,
           count_if(touched_b1_class LIKE 'b.%')                          AS t_b,
           count_if(touched_b1_class LIKE 'c.%')                          AS t_c,
           count_if(touched_b1_class LIKE 'd.%')                          AS t_d
    FROM {DB}.uc2_t16_01n_populations
""").first()
chk("01n grain (rows = distinct accounts)", _r["rows"], _r["accts"])
chk("ledger all", _r["ledger_all"], E["ledger all"])
chk("ledger AA row", _r["ledger_aa"], E["ledger AA row"])
chk("ledger exaa", _r["ledger_exaa"], E["ledger exaa"])
chk("ledger exaa balance", int(_r["exaa_bal"] or 0), E["ledger exaa balance"], tol=5)
chk("touched b1", _r["touched_b1"], E["touched b1"])
chk("touched a. cured", _r["t_a"], E["touched a. cured"])
chk("touched b. bucket 1", _r["t_b"], E["touched b. bucket 1"])
chk("touched c. rolled past", _r["t_c"], E["touched c. rolled past"])
chk("touched d. jan chargeoff", _r["t_d"], E["touched d. jan chargeoff"])

# COMMAND ----------

# MAGIC %md
# MAGIC ## K5. Call-table key probe (BEFORE the 02n build)
# MAGIC
# MAGIC REFRESH first (live loading edge). The shape census and the mismatch
# MAGIC count reproduce the phase-1d evidence EXACTLY: effdt is a load date, so
# MAGIC rows below the 2026-07-10 bound are immutable.

# COMMAND ----------

sec("K5 call key probe")

spark.sql(f"REFRESH TABLE {CALL}")

_r = spark.sql(f"""
    SELECT count_if(acctid IS NULL) AS null_rows,
           count_if(acctid IS NOT NULL AND cast(acctid AS string) rlike '^[0-9]+$') AS digits_rows,
           count_if(acctid IS NOT NULL AND NOT cast(acctid AS string) rlike '^[0-9]+$') AS other_rows,
           count_if(initiationmethod = 'INBOUND' AND acctid IS NOT NULL
                    AND (try_cast(acctid AS bigint) IS NULL
                         OR trim(cast(acctid AS string)) <> {NUM_KEY.format(c="acctid")})) AS mismatches
    FROM {CALL}
    WHERE `date` >= DATE '{CALL_WIN_START}' AND `date` < DATE '{CALL_WIN_END}'
      AND effdt >= '{EFFDT_SCAN_START}' AND effdt < '{EFFDT_HARD_END}'
""").first()
chk("jan acctid null rows", _r["null_rows"], E["jan acctid null rows"])
chk("jan acctid digits-only rows", _r["digits_rows"], E["jan acctid digits-only rows"])
chk("jan acctid other-shape rows", _r["other_rows"], E["jan acctid other-shape rows"])
chk("jan key mismatches (id-carrying inbound)", _r["mismatches"], E["jan key mismatches (id-carrying inbound)"])

# COMMAND ----------

# MAGIC %md
# MAGIC ## K6. Build `uc2_t16_02n_episodes` (the call layer on the numeric key)
# MAGIC
# MAGIC Tier-16 layer 02 verbatim, three deltas: the numeric acct_key (the fix),
# MAGIC a had_zero_pad diagnostic column, and the bounded effdt scan (D8,
# MAGIC disclosed; cannot move any anchor - standard episodes require the
# MAGIC stricter in-column cap). Filter-then-dedup order unchanged.

# COMMAND ----------

sec("K6 build 02n")

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_02n_episodes AS
WITH calls_flagged AS (
    SELECT {NUM_KEY.format(c="acctid")} AS acct_key,      -- THE KEY CHANGE (D2)
           contactid,
           `date` AS call_dt,
           cast(date_trunc('month', `date`) AS date) AS call_month,
           initiationtimestamp,
           CASE WHEN coalesce(cast(producttype AS string), '') = 'BUSINESS_CARD'
                THEN 1 ELSE 0 END AS is_biz,
           CASE WHEN effdt >= '{EFFDT_CAP_START}' AND effdt < '{EFFDT_CAP_END}'
                THEN 1 ELSE 0 END AS within_effdt_cap,
           CASE WHEN try_cast(acctid AS bigint) IS NOT NULL
                 AND trim(cast(acctid AS string)) <> {NUM_KEY.format(c="acctid")}
                THEN 1 ELSE 0 END AS had_zero_pad          -- diagnostic: the rows the string key lost
    FROM {CALL}
    WHERE initiationmethod = 'INBOUND'
      AND `date` >= DATE '{CALL_WIN_START}' AND `date` < DATE '{CALL_WIN_END}'
      AND acctid IS NOT NULL
      AND effdt >= '{EFFDT_SCAN_START}' AND effdt < '{EFFDT_HARD_END}'   -- D8 bounded scan
),
episodes_std AS (
    SELECT contactid
    FROM (
        SELECT contactid,
               row_number() OVER (PARTITION BY acct_key, call_dt
                                  ORDER BY initiationtimestamp) AS rn
        FROM calls_flagged
        WHERE acct_key IS NOT NULL AND acct_key <> ''
          AND is_biz = 0
          AND within_effdt_cap = 1
    )
    WHERE rn = 1
)
SELECT c.acct_key, c.contactid, c.call_dt, c.call_month,
       c.is_biz, c.within_effdt_cap, c.had_zero_pad,
       CASE WHEN e.contactid IS NOT NULL THEN 1 ELSE 0 END AS is_episode_std
FROM calls_flagged c
LEFT JOIN episodes_std e ON e.contactid = c.contactid
""")
print(f"built {DB}.uc2_t16_02n_episodes")

kv("02n build summary", spark.sql(f"""
    SELECT count(*) AS rows,
           count_if(is_episode_std = 1) AS standard_episodes,
           count_if(had_zero_pad = 1) AS zero_pad_rows,
           count_if(acct_key IS NULL) AS non_castable_id_rows
    FROM {DB}.uc2_t16_02n_episodes
"""))

# COMMAND ----------

# MAGIC %md
# MAGIC ## K7. Build `uc2_t16_03n_signals` (the ONE transcript pass)
# MAGIC
# MAGIC Tier-16 layer 03 verbatim on 01n/02n. The round-9 OOM lessons carried:
# MAGIC the transcript table is referenced exactly once in the whole chain;
# MAGIC participantid = CUSTOMER plus the coarse union pre-filter in the WHERE;
# MAGIC boolean presence per contactid, not raw counts.

# COMMAND ----------

sec("K7 build 03n")

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_03n_signals AS
WITH drivers AS (
    SELECT DISTINCT e.contactid
    FROM {DB}.uc2_t16_02n_episodes e
    LEFT JOIN {DB}.uc2_t16_01n_populations p ON p.acct_key = e.acct_key
    WHERE e.is_episode_std = 1
      AND (p.acct_key IS NULL OR p.is_exaa)
),
tx AS (
    SELECT t.contactid,
           max(CASE WHEN regexp_like(lower(t.content),
                     'passed away|death certificate|executor|deceased|calling on behalf') THEN 1 ELSE 0 END) AS deceased_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'pay|paid|payment|settle|payment plan|arrangement|work something out') THEN 1 ELSE 0 END) AS pay_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'settle|payment plan|arrangement|work something out') THEN 1 ELSE 0 END) AS plan_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford') THEN 1 ELSE 0 END) AS hard_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'dispute|not my charge|didn.t authorize|did not authorize|unauthorized|fraud|identity theft') THEN 1 ELSE 0 END) AS dispute_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'i.ll pay|i will pay|going to pay|gonna pay|pay (on|by|this|next)|when i get paid|payday|after my paycheck') THEN 1 ELSE 0 END) AS promise_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'bank routing|routing number|check number|checkbook|a check for|that check|on the check') THEN 1 ELSE 0 END) AS exec_f
    FROM {TX} t
    JOIN drivers d ON t.contactid = d.contactid
    WHERE t.effdt >= '{EFFDT_CAP_START}' AND t.effdt < '{EFFDT_CAP_END}'
      AND t.content IS NOT NULL
      AND t.participantid = 'CUSTOMER'
      AND regexp_like(lower(t.content),
            'pay|paid|payment|settle|arrangement|work something out|passed away|death certificate|executor|deceased|calling on behalf|hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford|dispute|not my charge|didn.t authorize|did not authorize|unauthorized|fraud|identity theft|i.ll pay|i will pay|going to pay|gonna pay|when i get paid|payday|after my paycheck|bank routing|routing number|check number|checkbook|a check for|that check|on the check')
    GROUP BY 1
)
SELECT contactid,
       deceased_f, promise_f, pay_f, plan_f, hard_f, dispute_f, exec_f,
       CASE
         WHEN deceased_f > 0 THEN 'a. deceased or estate'
         WHEN promise_f  > 0 THEN 'b. future-dated promise'
         WHEN pay_f > 0 AND plan_f = 0 THEN 'c. payment talk, no promise'
         WHEN plan_f     > 0 THEN 'd. plan or settlement talk'
         WHEN hard_f     > 0 THEN 'e. hardship talk'
         WHEN dispute_f  > 0 THEN 'f. dispute or fraud talk'
         ELSE 'g. no payment-related language'
       END AS language_group
FROM tx
""")
print(f"built {DB}.uc2_t16_03n_signals")

_r = spark.sql(f"""
    SELECT count(*) AS rows, count(DISTINCT contactid) AS cids
    FROM {DB}.uc2_t16_03n_signals
""").first()
chk("03n grain (rows = distinct contactids)", _r["rows"], _r["cids"])

# COMMAND ----------

# MAGIC %md
# MAGIC ## K8. Build `uc2_t16_04n_outcomes`
# MAGIC
# MAGIC Tier-16 layer 04 verbatim on the n-layers. One addition: a has_tx column
# MAGIC (transcript signal row exists), a faithful port of the copilot export's
# MAGIC has_tx, needed by B04's stratum C. The capture gate here is the AWS
# MAGIC day-grain 30-day gate: from Phase 2 onward it is a DIAGNOSTIC, never a
# MAGIC denominator (the headline gate is captured_sas on the SAS spine, B01).

# COMMAND ----------

sec("K8 build 04n")

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_04n_outcomes AS
WITH episodes_exaa AS (
    SELECT c.acct_key, c.contactid, c.call_dt, c.call_month
    FROM {DB}.uc2_t16_02n_episodes c
    LEFT JOIN {DB}.uc2_t16_01n_populations p ON p.acct_key = c.acct_key
    WHERE c.is_episode_std = 1
      AND (p.acct_key IS NULL OR p.is_exaa)
),
pay_lead AS (
    SELECT acct_key,
           to_date(concat(ym, '01'), 'yyyyMMdd') AS m,
           pay_dt, auto_dt, nsf_dt,
           lead(pay_dt)  OVER (PARTITION BY acct_key ORDER BY ym) AS next_pay_dt,
           lead(auto_dt) OVER (PARTITION BY acct_key ORDER BY ym) AS next_auto_dt,
           lead(nsf_dt)  OVER (PARTITION BY acct_key ORDER BY ym) AS next_nsf_dt
    FROM {DB}.uc2_t16_00n_acct_monthly
    WHERE acct_key IN (SELECT acct_key FROM episodes_exaa)
),
ep AS (
    SELECT e.acct_key, e.contactid, e.call_dt, e.call_month,
           CASE WHEN
                  (s.pay_dt IS NOT NULL
                   AND s.pay_dt >= e.call_dt
                   AND s.pay_dt <= date_add(e.call_dt, 30)
                   AND (s.auto_dt IS NULL OR s.auto_dt <> s.pay_dt)
                   AND (s.nsf_dt IS NULL OR s.nsf_dt <> s.pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.pay_dt))
                OR
                  (s.next_pay_dt IS NOT NULL
                   AND s.next_pay_dt >= e.call_dt
                   AND s.next_pay_dt <= date_add(e.call_dt, 30)
                   AND (s.next_auto_dt IS NULL OR s.next_auto_dt <> s.next_pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.next_pay_dt))
                THEN 1 ELSE 0 END AS captured
    FROM episodes_exaa e
    LEFT JOIN pay_lead s
      ON s.acct_key = e.acct_key
     AND s.m = e.call_month
),
snap_daily AS (
    SELECT {NUM_KEY.format(c="extnl_acct_id")} AS acct_key,
           eff_dt,
           to_date(eff_dt, 'yyyyMMdd') AS snap_dt,
           CASE
             WHEN past_due_271_up_amt  > 0 THEN 10
             WHEN past_due_241_270_amt > 0 THEN 9
             WHEN past_due_211_240_amt > 0 THEN 8
             WHEN past_due_181_210_amt > 0 THEN 7
             WHEN past_due_151_180_amt > 0 THEN 6
             WHEN past_due_121_150_amt > 0 THEN 5
             WHEN past_due_91_120_amt  > 0 THEN 4
             WHEN past_due_61_90_amt   > 0 THEN 3
             WHEN past_due_31_60_amt   > 0 THEN 2
             WHEN past_due_1_30_amt    > 0 THEN 1
             ELSE 0
           END AS bucket,
           try_cast(chrgoff_dt AS date) AS co_dt
    FROM {FMT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '{SNAP_DAILY_START}' AND eff_dt < '{SNAP_DAILY_END}'
      AND {NUM_KEY.format(c="extnl_acct_id")} IN (SELECT acct_key FROM episodes_exaa)
),
callday AS (
    SELECT e.acct_key, e.call_dt,
           max_by(s.bucket, s.eff_dt) AS callday_bucket,
           max_by(s.co_dt, s.eff_dt) AS callday_co_dt
    FROM (SELECT DISTINCT acct_key, call_dt FROM episodes_exaa) e
    JOIN snap_daily s
      ON s.acct_key = e.acct_key
     AND s.snap_dt <= e.call_dt
    GROUP BY 1, 2
),
esig AS (
    SELECT e.acct_key, e.contactid, e.call_dt, e.captured,
           coalesce(x.language_group, 'g. no payment-related language') AS language_group,
           coalesce(x.pay_f, 0)      AS pay_f,
           coalesce(x.deceased_f, 0) AS deceased_f,
           coalesce(x.exec_f, 0)     AS exec_f,
           CASE WHEN x.contactid IS NOT NULL THEN 1 ELSE 0 END AS has_tx,
           cd.callday_bucket,
           (cd.callday_bucket = 1
            AND (cd.callday_co_dt IS NULL OR cd.callday_co_dt >= DATE '{CLEANUP_DATE}'))
               AS is_addressable
    FROM ep e
    LEFT JOIN {DB}.uc2_t16_03n_signals x ON x.contactid = e.contactid
    LEFT JOIN callday cd ON cd.acct_key = e.acct_key AND cd.call_dt = e.call_dt
),
esig_acct AS (
    SELECT *,
           max(captured)   OVER (PARTITION BY acct_key) AS any_captured,
           max(CASE WHEN captured = 0 AND pay_f > 0 THEN 1 ELSE 0 END)
                           OVER (PARTITION BY acct_key) AS any_leaked_intent,
           max(deceased_f) OVER (PARTITION BY acct_key) AS deceased_acct
    FROM esig
)
SELECT a.acct_key, a.contactid, a.call_dt,
       a.captured, a.language_group, a.pay_f, a.deceased_f, a.exec_f, a.has_tx,
       a.callday_bucket, a.is_addressable,
       a.any_captured, a.any_leaked_intent, a.deceased_acct,
       CASE
         WHEN a.any_captured = 1 THEN 'b. captured (>= 1 paid-30d episode)'
         WHEN a.any_leaked_intent = 1 THEN 'c. leaked-intent (intent, no payment 30d)'
         ELSE 'd. other-caller'
       END AS caller_class,
       (a.any_captured = 0 AND a.any_leaked_intent = 1) AS leaked_acct,
       (a.any_captured = 0 AND a.any_leaked_intent = 1
        AND coalesce(p.in_ledger_exaa, false) AND a.deceased_acct = 0) AS w_flag,
       coalesce(p.in_ledger_all, false)  AS in_ledger_all,
       coalesce(p.in_ledger_exaa, false) AS in_ledger_exaa,
       coalesce(p.touched_b1, false)     AS touched_b1,
       p.touched_b1_class,
       p.eom_bal AS jan_eom_bal, p.cpc_class, p.runway_band,
       p.feb_position_b14, p.feb_pos, p.mar_pos,
       p.co_dt_future, p.co_amt, p.co_8m, p.co_10m, p.co_12m
FROM esig_acct a
LEFT JOIN {DB}.uc2_t16_01n_populations p ON p.acct_key = a.acct_key
""")
print(f"built {DB}.uc2_t16_04n_outcomes")

# COMMAND ----------

# MAGIC %md
# MAGIC ## K9. THE RE-ANCHOR
# MAGIC
# MAGIC Population side: already asserted exact (K4). Caller side: implication
# MAGIC checks assert (dropped_old = 0; episodes and the call-day stream can only
# MAGIC grow), everything else is MEASURED and locked only after verification.
# MAGIC Definition sentence for every number printed with it.

# COMMAND ----------

sec("K9 re-anchor: implication checks")

spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW _old_callers AS
    SELECT DISTINCT c.acct_key
    FROM {DB}.uc2_t16_02_episodes c
    JOIN {DB}.uc2_t16_01_populations p ON p.acct_key = c.acct_key
    WHERE c.is_episode_std = 1 AND p.in_ledger_exaa
""")
spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW _new_callers AS
    SELECT DISTINCT acct_key
    FROM {DB}.uc2_t16_04n_outcomes
    WHERE in_ledger_exaa
""")

# implication 1: no old caller may disappear (numeric normalization cannot
# unmatch a previously matched digits-only key; K2 proved no fmt key moved)
_dropped_df = spark.sql("SELECT o.acct_key FROM _old_callers o LEFT ANTI JOIN _new_callers n ON n.acct_key = o.acct_key")
chk("dropped old callers (must be zero)", _dropped_df.count(), 0, ctx=_dropped_df)

# measured: the numeric-keyed ledger callers and episodes
# definition: distinct ex-AA-ledger accounts with a standard January episode,
# numeric key; episodes = standard episodes on those accounts. Platform
# Databricks, vintage window per SETUP.
_r = spark.sql(f"""
    SELECT count_if(in_ledger_exaa) AS episodes,
           count(DISTINCT CASE WHEN in_ledger_exaa THEN acct_key END) AS callers
    FROM {DB}.uc2_t16_04n_outcomes
""").first()
if _r["episodes"] < E["hist ledger episodes (string key)"]:
    raise AssertionError(f"IMPLICATION MISS: episodes {fmt(_r['episodes'])} < historical {fmt(E['hist ledger episodes (string key)'])} - recovered rows can only ADD episodes")
chk("ledger episodes (numeric key)", _r["episodes"], E["ledger episodes (numeric key)"])
chk("ledger callers (numeric key)", _r["callers"], E["ledger callers (numeric key)"])
print(f"historical references (string key, never asserted): callers {fmt(E['hist ledger callers (string key)'])}, episodes {fmt(E['hist ledger episodes (string key)'])}")

# measured: the call-day bucket-1 stream (the 29,114 construct re-measured)
# definition: standard ex-AA episodes whose account was bucket 1 on the call
# day with no pre-anchor charge-off (is_addressable).
_r = spark.sql(f"""
    SELECT count_if(is_addressable) AS addr_episodes,
           count_if(is_addressable AND pay_f > 0 AND captured = 0) AS wl_episodes,
           count(DISTINCT CASE WHEN is_addressable AND pay_f > 0 AND captured = 0 THEN acct_key END) AS wl_accounts
    FROM {DB}.uc2_t16_04n_outcomes
""").first()
if _r["addr_episodes"] < E["hist callday b1 stream"]:
    raise AssertionError(f"IMPLICATION MISS: call-day stream {fmt(_r['addr_episodes'])} < historical {fmt(E['hist callday b1 stream'])} - recovered rows can only ADD episodes")
chk("addressable episodes (callday b1 stream)", _r["addr_episodes"], E["addressable episodes (callday b1 stream)"])
# definition: addressable episodes with payment language and no captured
# payment under the AWS DAY-GRAIN gate (diagnostic construct; the SAS-gate
# addressable split lives in B03 block 5)
chk("addressable work list episodes", _r["wl_episodes"], E["addressable work list episodes"])
chk("addressable work list accounts", _r["wl_accounts"], E["addressable work list accounts"])

# COMMAND ----------

sec("K9 re-anchor: language partition and caller classes (measured)")

# language partition over ledger episodes (numeric key); the partition must
# sum to the measured episode count exactly
_lang_df = spark.sql(f"""
    SELECT language_group, count(*) AS episodes, count(DISTINCT acct_key) AS accounts
    FROM {DB}.uc2_t16_04n_outcomes
    WHERE in_ledger_exaa
    GROUP BY 1 ORDER BY 1
""")
grid("language partition over ledger episodes (numeric key, aws layers)", _lang_df)
_exp_lang = E["language partition"] or {}
_lang_total = 0
for _row in _lang_df.collect():
    chk(f"lang: {_row['language_group']}", _row["episodes"], _exp_lang.get(_row["language_group"]))
    _lang_total += _row["episodes"]
_r = spark.sql(f"SELECT count_if(in_ledger_exaa) AS n FROM {DB}.uc2_t16_04n_outcomes").first()
chk("language partition re-adds to ledger episodes", _lang_total, _r["n"])

# caller classes with the AWS day-grain gate (diagnostic labels; the SAS-gate
# classes live in B02b). 'a. non-caller' lives on the 01n side.
_cls_df = spark.sql(f"""
    WITH callers AS (
        SELECT acct_key, max_by(caller_class, contactid) AS caller_class
        FROM {DB}.uc2_t16_04n_outcomes
        WHERE in_ledger_exaa
        GROUP BY 1
    )
    SELECT coalesce(k.caller_class, 'a. non-caller') AS caller_class,
           count(*) AS accounts,
           round(sum(p.eom_bal), 0) AS jan_eom_balance
    FROM {DB}.uc2_t16_01n_populations p
    LEFT JOIN callers k ON k.acct_key = p.acct_key
    WHERE p.in_ledger_exaa
    GROUP BY 1 ORDER BY 1
""")
grid("caller classes (aws day-grain gate, numeric key)", _cls_df)
_exp_cls = E["caller classes (aws gate)"] or {}
_cls_total = 0
for _row in _cls_df.collect():
    chk(f"class: {_row['caller_class']}", _row["accounts"], _exp_cls.get(_row["caller_class"]))
    _cls_total += _row["accounts"]
chk("caller classes re-add to the ex-AA ledger", _cls_total, E["ledger exaa"])

# COMMAND ----------

sec("K9 re-anchor: W steps (measured)")

# W definition: strict leaked-intent accounts (never captured under the AWS
# day-grain gate, >= 1 uncaptured payment-language episode), in the ex-AA
# ledger, deceased-language accounts routed out. Balance = 31-Jan EOM, one
# row per account.
_r = spark.sql(f"""
    SELECT count(DISTINCT CASE WHEN leaked_acct AND in_ledger_exaa THEN acct_key END) AS leaked,
           count(DISTINCT CASE WHEN leaked_acct AND in_ledger_exaa AND deceased_acct = 1 THEN acct_key END) AS routed,
           count(DISTINCT CASE WHEN w_flag THEN acct_key END) AS w_accts
    FROM {DB}.uc2_t16_04n_outcomes
""").first()
chk("W strict leaked accounts", _r["leaked"], E["W strict leaked accounts"])
chk("W deceased routed", _r["routed"], E["W deceased routed"])
chk("W accounts", _r["w_accts"], E["W accounts"])
_r = spark.sql(f"""
    SELECT round(sum(jan_eom_bal), 0) AS bal
    FROM (SELECT DISTINCT acct_key, jan_eom_bal FROM {DB}.uc2_t16_04n_outcomes WHERE w_flag)
""").first()
chk("W balance", int(_r["bal"] or 0), E["W balance"])

# COMMAND ----------

# MAGIC %md
# MAGIC ## K9r. The 202501 reconciliation cells (guarded; read A's persisted tables, never the CSV)
# MAGIC
# MAGIC gained = new callers minus old; its overlap with the persisted 1,942 list
# MAGIC measures the recovery. Shortfall accounts are classified per account by
# MAGIC cause (business-card-only / out-of-cap-only / mixed); an UNEXPLAINED
# MAGIC shortfall = STOP. The flagged-overlap arithmetic must tie to the unit:
# MAGIC overlap = 9,194 + recovered.

# COMMAND ----------

sec("K9r 202501 reconciliation")

if ANCHOR_YM != "202501":
    print("skipped: these reconciliation cells are 202501-only by design")
else:
    _gained = spark.sql("SELECT count(*) AS n FROM _new_callers n LEFT ANTI JOIN _old_callers o ON o.acct_key = n.acct_key").first()["n"]
    chk("gained callers", _gained, E["gained callers"])

    _recovered = spark.sql(f"""
        SELECT count(*) AS n FROM {DB}.uc2_gap1942_202501 g
        JOIN _new_callers n ON n.acct_key = g.acct_key
    """).first()["n"]
    chk("gap1942 recovered", _recovered, E["gap1942 recovered"])
    _shortfall = 1942 - _recovered
    print(f"gap1942 shortfall = {fmt(_shortfall)} (recovered accounts whose only rows fail the standard filters; classified below)")

    _cause_df = spark.sql(f"""
        WITH short AS (
            SELECT g.acct_key FROM {DB}.uc2_gap1942_202501 g
            LEFT ANTI JOIN _new_callers n ON n.acct_key = g.acct_key
        ),
        r AS (
            SELECT s.acct_key, e.is_biz, e.within_effdt_cap
            FROM short s
            LEFT JOIN {DB}.uc2_t16_02n_episodes e ON e.acct_key = s.acct_key
        ),
        classed AS (
            SELECT acct_key,
                   CASE
                     WHEN max(CASE WHEN is_biz = 0 AND within_effdt_cap = 1 THEN 1 ELSE 0 END) = 1
                       THEN 'z. eligible row exists yet not a caller (unexplained - STOP)'
                     WHEN max(is_biz) IS NULL
                       THEN 'z. no 02n rows at all (unexplained - STOP)'
                     WHEN min(is_biz) = 1
                       THEN 'a. only business-card rows'
                     WHEN max(is_biz) = 0
                       THEN 'b. only out-of-effdt-cap rows'
                     ELSE 'c. mixed business-card / out-of-cap rows'
                   END AS cause
            FROM r GROUP BY acct_key
        )
        SELECT cause, count(*) AS accounts FROM classed GROUP BY 1 ORDER BY 1
    """)
    grid("gap1942 shortfall causes (per account)", _cause_df)
    _unexplained = sum(r["accounts"] for r in _cause_df.collect() if r["cause"].startswith("z."))
    chk("gap1942 shortfall unexplained (must be zero)", _unexplained, 0, ctx=_cause_df)

    _outside = spark.sql(f"""
        SELECT count(*) AS n
        FROM _new_callers n
        LEFT ANTI JOIN _old_callers o ON o.acct_key = n.acct_key
        LEFT ANTI JOIN {DB}.uc2_gap1942_202501 g ON g.acct_key = n.acct_key
    """).first()["n"]
    # definition: recovered callers OUTSIDE the flagged 1,942 - real January
    # callers whose accounts fail the SAS slice filters (the 195-analog in
    # reverse). Measured, labeled, never merged with the flagged set.
    chk("gained outside 1942", _outside, E["gained outside 1942"])

    _overlap = spark.sql(f"""
        SELECT count(*) AS n FROM {DB}.uc2_sasflag_202501 f
        JOIN _new_callers n ON n.acct_key = f.acct_key
    """).first()["n"]
    chk("flagged overlap arithmetic tie (= 9,194 + recovered)", _overlap, 9194 + _recovered)
    chk("flagged overlap (202501 recon)", _overlap, E["flagged overlap (202501 recon)"])

# COMMAND ----------

# MAGIC %md
# MAGIC ## K10. Verdict and record block

# COMMAND ----------

sec("K10 verdict")

print("TWO-PHASE LOCK: after this run is verified from screenshots, the")
print("MEASURED values above are written into EXPECTED['202501'] in")
print("B00_setup.py and re-pasted into every B file; the second run asserts.")

record_block("B02_keyfix_aws_layers")
flush_metrics("B02_keyfix_aws_layers")
