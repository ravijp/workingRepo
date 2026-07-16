# Databricks notebook source
# MAGIC %md
# MAGIC # B01. The SAS spine: `uc2_t16_01s_populations_<vintage>` (runs AFTER B02)
# MAGIC
# MAGIC Grain: one row per export account (610,183 for 202501). Population,
# MAGIC delinquency, and dollars come from the client's SAS 003-program export;
# MAGIC this table is the denominator spine for every Phase-2 insight.
# MAGIC
# MAGIC GO-FORWARD RULE (the loop cut): the CSV's call_type_* columns are NOT
# MAGIC read here or anywhere downstream - they are AWS-origin (an Athena export
# MAGIC imported into SAS). The caller flag is rebuilt NATIVELY: inb_native =
# MAGIC the account has any January INBOUND id-resolved call row under the
# MAGIC numeric key (any 02n row; the SAS import had no effdt cap and no
# MAGIC business-card exclusion, so "any 02n row" is the faithful analog).
# MAGIC One-time tie-out against the CSV flag below; after that the CSV flag
# MAGIC retires to the frozen notebook A.
# MAGIC
# MAGIC captured_sas (the headline gate, month grain, ACCOUNT grain): a negative
# MAGIC PAYMT_AMT in M1 or M2 (the CQ-7 convention, confirmed by A11). There is
# MAGIC no episode-grain "captured" under this gate. Deltas from the AWS
# MAGIC day-grain gate, disclosed wherever captured_sas is quoted: month grain
# MAGIC not 30-days-from-call; no autopay/NSF exclusions; export window
# MAGIC semantics. aws_ columns below are diagnostics, never denominators.
# MAGIC
# MAGIC The typed-column extents (M0-M4 for the ECL families, M1-M3 for the ASP
# MAGIC families) are pinned by notebook A's schema census; a missing column
# MAGIC fails loud in the build cell and the fix is a one-line edit against the
# MAGIC census.

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
# MAGIC ## S1. Preconditions: the fixed-key n-layers exist and pass the sweep

# COMMAND ----------

sec("S1 preconditions")

for _t in ["uc2_t16_00n_acct_monthly", "uc2_t16_01n_populations",
           "uc2_t16_02n_episodes", "uc2_t16_04n_outcomes"]:
    assert spark.catalog.tableExists(f"{DB}.{_t}"), \
        f"PRECONDITION MISS: {DB}.{_t} missing - run B02_keyfix_aws_layers first"
    print(f"PASS  n-layer exists: {DB}.{_t}")

_r = spark.sql(f"""
    SELECT count_if(in_ledger_all) AS ledger_all,
           count_if(in_ledger_exaa) AS ledger_exaa,
           count_if(touched_b1) AS touched_b1
    FROM {DB}.uc2_t16_01n_populations
""").first()
chk("ledger all", _r["ledger_all"], E["ledger all"])
chk("ledger exaa", _r["ledger_exaa"], E["ledger exaa"])
chk("touched b1", _r["touched_b1"], E["touched b1"])

# COMMAND ----------

# MAGIC %md
# MAGIC ## S2. The CSV load (all-string, FAILFAST, grain proof)

# COMMAND ----------

sec("S2 CSV load")

csv_df = (spark.read.format("csv")
          .option("header", True)
          .option("inferSchema", False)     # every column arrives as string
          .option("mode", "FAILFAST")       # malformed row = STOP
          .load(SAS_CSV_PATH))
csv_df.createOrReplaceTempView("_sas_csv")
# At A-freeze the explicit all-string StructType from the A5 census replaces
# the header read here too (the locked explicit-schema convention).

_r = spark.sql("SELECT count(*) AS rows, count(DISTINCT EXTNL_ACCT_ID) AS accts FROM _sas_csv").first()
chk("csv rows", _r["rows"], E["csv rows"])
chk("csv distinct accounts", _r["accts"], E["csv distinct accounts"])

# COMMAND ----------

# MAGIC %md
# MAGIC ## S3. PAYMT sign re-probe (per vintage; the gate is defined only after this)
# MAGIC
# MAGIC Pre-registered (CQ-7): a true payment is a NEGATIVE PAYMT_AMT in M1/M2.
# MAGIC The coarse tripwire below STOPS the run if this vintage's file
# MAGIC contradicts the convention; the full probe is notebook A cell A11.

# COMMAND ----------

sec("S3 sign re-probe")

_sign_df = spark.sql("""
    SELECT count_if(try_cast(PAYMT_AMT_M1 AS double) < 0) AS neg_m1,
           count_if(try_cast(PAYMT_AMT_M1 AS double) > 0) AS pos_m1,
           count_if(try_cast(PAYMT_AMT_M2 AS double) < 0) AS neg_m2,
           count_if(try_cast(PAYMT_AMT_M2 AS double) > 0) AS pos_m2
    FROM _sas_csv
""")
kv("PAYMT sign tripwire", _sign_df)
_r = _sign_df.first()
assert _r["neg_m1"] > _r["pos_m1"] and _r["neg_m2"] > _r["pos_m2"], (
    "SIGN CONVENTION CONTRADICTED: negatives do not dominate PAYMT_AMT M1/M2. "
    "STOP AND INVESTIGATE (CQ-7 pre-registration) before any captured_sas number is read.")
print("PASS  sign convention holds: payments are NEGATIVE PAYMT_AMT values (CQ-7)")

# COMMAND ----------

# MAGIC %md
# MAGIC ## S4. Build `uc2_t16_01s_populations_<vintage>`
# MAGIC
# MAGIC Explicit column-by-column SELECT. NO call_type_* column is read (the
# MAGIC loop cut). aws_ columns are joined BY NUMERIC KEY from the FIXED n-build
# MAGIC and are diagnostics, never denominators.

# COMMAND ----------

sec("S4 build 01s")

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_01s_populations_{ANCHOR_YM} AS
WITH k AS (
    SELECT r.*, try_cast(EXTNL_ACCT_ID AS bigint) AS acct_num
    FROM _sas_csv r
),
f AS (
    SELECT k.*,
           coalesce(try_cast(DLNQT_CD_M1 AS int) = 1, false) AS wf_dq1,
           coalesce(upper(trim(CPC_FLAG_NW)) IN ('OTHER', 'OTHERS', 'COBRAND', 'PLCC'), false) AS wf_cpc,
           (CHRGOFF_RSN_M1 IS NULL OR trim(CHRGOFF_RSN_M1) = ''
            OR upper(trim(CHRGOFF_RSN_M1)) IN ('PLY', 'BLANK')) AS wf_non_co
    FROM k
),
inb AS (SELECT DISTINCT acct_key FROM {DB}.uc2_t16_02n_episodes),
awsp AS (SELECT acct_key, in_ledger_all, in_ledger_exaa, touched_b1, eom_bal
         FROM {DB}.uc2_t16_01n_populations),
awsc AS (SELECT acct_key, max(is_episode_std) AS has_std FROM {DB}.uc2_t16_02n_episodes GROUP BY 1),
awscap AS (SELECT acct_key, max(any_captured) AS any_captured FROM {DB}.uc2_t16_04n_outcomes GROUP BY 1)
SELECT cast(f.acct_num AS string) AS acct_key,
       f.acct_num,
       -- SAS-native waterfall flags
       f.wf_dq1, f.wf_cpc, f.wf_non_co,
       (f.wf_dq1 AND f.wf_cpc AND f.wf_non_co) AS in_sas_ledger,
       -- delinquency codes (strings)
       f.DLNQT_CD_M1 AS dlnqt_cd_m1, f.DLNQT_CD_M2 AS dlnqt_cd_m2, f.DLNQT_CD_M3 AS dlnqt_cd_m3,
       f.DLNQT_BKT_M1 AS dlnqt_bkt_m1, f.DLNQT_BKT_M2 AS dlnqt_bkt_m2, f.DLNQT_BKT_M3 AS dlnqt_bkt_m3,
       -- payments (typed; sign convention CQ-7: payments are negative)
       try_cast(f.PAYMT_AMT_M1 AS double) AS paymt_amt_m1,
       try_cast(f.PAYMT_AMT_M2 AS double) AS paymt_amt_m2,
       try_cast(f.PAYMT_AMT_M3 AS double) AS paymt_amt_m3,
       -- balances and limits (typed)
       try_cast(f.EOP_BAL_M1 AS double) AS eop_bal_m1,
       try_cast(f.EOP_BAL_M2 AS double) AS eop_bal_m2,
       try_cast(f.EOP_BAL_M3 AS double) AS eop_bal_m3,
       try_cast(f.CR_LMT_M1 AS double) AS cr_lmt_m1,
       try_cast(f.CR_LMT_M2 AS double) AS cr_lmt_m2,
       try_cast(f.CR_LMT_M3 AS double) AS cr_lmt_m3,
       -- monthly losses (typed)
       try_cast(f.GROSS_LOSS_M1 AS double) AS gross_loss_m1,
       try_cast(f.GROSS_LOSS_M2 AS double) AS gross_loss_m2,
       try_cast(f.GROSS_LOSS_M3 AS double) AS gross_loss_m3,
       try_cast(f.CHRGOFF_AMT_M1 AS double) AS chrgoff_amt_m1,
       try_cast(f.CHRGOFF_AMT_M2 AS double) AS chrgoff_amt_m2,
       try_cast(f.CHRGOFF_AMT_M3 AS double) AS chrgoff_amt_m3,
       try_cast(f.CHRGOFF_RVRSL_M1 AS double) AS chrgoff_rvrsl_m1,
       try_cast(f.CHRGOFF_RVRSL_M2 AS double) AS chrgoff_rvrsl_m2,
       try_cast(f.CHRGOFF_RVRSL_M3 AS double) AS chrgoff_rvrsl_m3,
       try_cast(f.PLCY_LOSS_M1 AS double) AS plcy_loss_m1,
       try_cast(f.PLCY_LOSS_M2 AS double) AS plcy_loss_m2,
       try_cast(f.PLCY_LOSS_M3 AS double) AS plcy_loss_m3,
       -- charge-off windows (typed; the export's own window semantics)
       try_cast(f.GROSS_LOSS_8M_AMT AS double) AS gross_loss_8m_amt,
       try_cast(f.GROSS_LOSS_10M_AMT AS double) AS gross_loss_10m_amt,
       try_cast(f.GROSS_LOSS_12M_AMT AS double) AS gross_loss_12m_amt,
       try_cast(f.CHRGOFF_8M_AMT AS double) AS chrgoff_8m_amt,
       try_cast(f.CHRGOFF_10M_AMT AS double) AS chrgoff_10m_amt,
       try_cast(f.CHRGOFF_12M_AMT AS double) AS chrgoff_12m_amt,
       try_cast(f.PLCY_LOSS_8M_AMT AS double) AS plcy_loss_8m_amt,
       try_cast(f.PLCY_LOSS_10M_AMT AS double) AS plcy_loss_10m_amt,
       try_cast(f.PLCY_LOSS_12M_AMT AS double) AS plcy_loss_12m_amt,
       try_cast(f.CHRGOFF_AMT_LFTM AS double) AS chrgoff_amt_lftm,
       try_cast(f.GROSS_LOSS_AMT_LFTM AS double) AS gross_loss_amt_lftm,
       f.CHRGOFF_VAL_FLAG AS chrgoff_val_flag,
       -- impairment / ECL families (typed; M0-M4 per the round-10 record,
       -- pinned by the A5 census)
       try_cast(f.ECL_M0 AS double) AS ecl_m0,
       try_cast(f.ECL_M1 AS double) AS ecl_m1,
       try_cast(f.ECL_M2 AS double) AS ecl_m2,
       try_cast(f.ECL_M3 AS double) AS ecl_m3,
       try_cast(f.ECL_M4 AS double) AS ecl_m4,
       try_cast(f.ECL_12MO_M0 AS double) AS ecl_12mo_m0,
       try_cast(f.ECL_12MO_M1 AS double) AS ecl_12mo_m1,
       try_cast(f.ECL_12MO_M2 AS double) AS ecl_12mo_m2,
       try_cast(f.ECL_12MO_M3 AS double) AS ecl_12mo_m3,
       try_cast(f.ECL_12MO_M4 AS double) AS ecl_12mo_m4,
       try_cast(f.ECL_LIFTM_M0 AS double) AS ecl_liftm_m0,
       try_cast(f.ECL_LIFTM_M1 AS double) AS ecl_liftm_m1,
       try_cast(f.ECL_LIFTM_M2 AS double) AS ecl_liftm_m2,
       try_cast(f.ECL_LIFTM_M3 AS double) AS ecl_liftm_m3,
       try_cast(f.ECL_LIFTM_M4 AS double) AS ecl_liftm_m4,
       f.STG_CD_M0 AS stg_cd_m0, f.STG_CD_M1 AS stg_cd_m1, f.STG_CD_M2 AS stg_cd_m2,
       f.STG_CD_M3 AS stg_cd_m3, f.STG_CD_M4 AS stg_cd_m4,
       try_cast(f.WRITE_OFF_M0 AS double) AS write_off_m0,
       try_cast(f.WRITE_OFF_M1 AS double) AS write_off_m1,
       try_cast(f.WRITE_OFF_M2 AS double) AS write_off_m2,
       try_cast(f.WRITE_OFF_M3 AS double) AS write_off_m3,
       try_cast(f.WRITE_OFF_M4 AS double) AS write_off_m4,
       -- codes and flags kept as strings
       f.CHRGOFF_RSN_M1 AS chrgoff_rsn_m1, f.CHRGOFF_RSN_M2 AS chrgoff_rsn_m2, f.CHRGOFF_RSN_M3 AS chrgoff_rsn_m3,
       f.CPC_FLAG_NW AS cpc_flag_nw,
       f.CO_CURRENT_FLAG AS co_current_flag, f.CO_8M_FLAG AS co_8m_flag,
       f.CO_10M_FLAG AS co_10m_flag, f.CO_12M_FLAG AS co_12m_flag,
       f.REAGE_EVER_FLAG AS reage_ever_flag,
       f.NEW_ROLL_FLAG AS new_roll_flag, f.NO_PRIOR_RECORD_FLAG AS no_prior_record_flag,
       f.hram_flag_refit_M1 AS hram_flag_refit_m1, f.hram_flag_refit_M2 AS hram_flag_refit_m2,
       f.hram_flag_refit_M3 AS hram_flag_refit_m3,
       f.hram_flag_apollo_M1 AS hram_flag_apollo_m1, f.hram_flag_apollo_M2 AS hram_flag_apollo_m2,
       f.hram_flag_apollo_M3 AS hram_flag_apollo_m3,
       -- the headline gate (month grain, ACCOUNT grain; CQ-7 convention)
       (coalesce(try_cast(f.PAYMT_AMT_M1 AS double), 0) < 0
        OR coalesce(try_cast(f.PAYMT_AMT_M2 AS double), 0) < 0) AS captured_sas,
       -- the NATIVE caller flag (the loop cut): any January INBOUND
       -- id-resolved call row under the numeric key (any 02n row)
       (i.acct_key IS NOT NULL) AS inb_native,
       -- aws_ diagnostics (numeric-key joins to the FIXED build); never a denominator
       coalesce(p.in_ledger_all, false)  AS aws_in_ledger_all,
       coalesce(p.in_ledger_exaa, false) AS aws_in_ledger_exaa,
       coalesce(p.touched_b1, false)     AS aws_touched_b1,
       p.eom_bal                          AS aws_jan_eom_bal,
       coalesce(c.has_std, 0) = 1         AS aws_caller,
       coalesce(cap.any_captured, 0) = 1  AS aws_captured
FROM f
LEFT JOIN inb i     ON i.acct_key = cast(f.acct_num AS string)
LEFT JOIN awsp p    ON p.acct_key = cast(f.acct_num AS string)
LEFT JOIN awsc c    ON c.acct_key = cast(f.acct_num AS string)
LEFT JOIN awscap cap ON cap.acct_key = cast(f.acct_num AS string)
""")
print(f"built {DB}.uc2_t16_01s_populations_{ANCHOR_YM}")

# COMMAND ----------

# MAGIC %md
# MAGIC ## S5. Asserts, the native-flag ladder, and the ONE-TIME CSV-flag tie-out

# COMMAND ----------

sec("S5 asserts and tie-outs")

_r = spark.sql(f"""
    SELECT count(*) AS s1,
           count_if(wf_dq1) AS s2,
           count_if(wf_dq1 AND wf_cpc) AS s3,
           count_if(in_sas_ledger) AS s4,
           count_if(inb_native) AS n1,
           count_if(inb_native AND wf_dq1) AS n2,
           count_if(inb_native AND wf_dq1 AND wf_cpc) AS n3,
           count_if(inb_native AND in_sas_ledger) AS n4,
           count_if(captured_sas) AS cap_all,
           count_if(captured_sas AND in_sas_ledger) AS cap_ledger
    FROM {DB}.uc2_t16_01s_populations_{ANCHOR_YM}
""").first()
chk("wf 01 total (rows)", _r["s1"], E["csv rows"])
chk("wf 02 dq1", _r["s2"], E["wf 02 dq1"])
chk("wf 03 +cpc", _r["s3"], E["wf 03 +cpc"])
chk("wf 04 sas ledger", _r["s4"], E["wf 04 sas ledger"])
# the native-flag ladder (definition: account has any January INBOUND
# id-resolved call row, numeric key; SAS-import analog, no standard filters)
chk("inb_native 01 total", _r["n1"], E["inb_native 01 total"])
chk("inb_native 02 dq1", _r["n2"], E["inb_native 02 dq1"])
chk("inb_native 03 +cpc", _r["n3"], E["inb_native 03 +cpc"])
chk("inb_native 04 ledger", _r["n4"], E["inb_native 04 ledger"])
chk("captured_sas all", _r["cap_all"], E["captured_sas all"])
chk("captured_sas ledger", _r["cap_ledger"], E["captured_sas ledger"])

# ONE-TIME tie-out vs the CSV flag (202501 only; after this the CSV flag
# retires everywhere but the frozen notebook A). Expect near-exact: the
# residual is the SAS import's own window/universe and flag spelling.
if ANCHOR_YM == "202501" and spark.catalog.tableExists(f"{DB}.uc2_sas_wf_202501"):
    grid("one-time tie-out: inb_native x csv_inb (all 610,183)", spark.sql(f"""
        SELECT s.inb_native, w.csv_inb, count(*) AS accounts
        FROM {DB}.uc2_t16_01s_populations_{ANCHOR_YM} s
        JOIN {DB}.uc2_sas_wf_202501 w ON w.acct_num = s.acct_num
        GROUP BY 1, 2 ORDER BY 1, 2
    """))
    grid("one-time tie-out: inb_native x csv_inb (SAS ledger only)", spark.sql(f"""
        SELECT s.inb_native, w.csv_inb, count(*) AS accounts
        FROM {DB}.uc2_t16_01s_populations_{ANCHOR_YM} s
        JOIN {DB}.uc2_sas_wf_202501 w ON w.acct_num = s.acct_num
        WHERE s.in_sas_ledger
        GROUP BY 1, 2 ORDER BY 1, 2
    """))
else:
    print("tie-out skipped (not 202501, or notebook A's wf table absent)")

# dollars, side by side with the SAS-recorded slice (NEVER asserted equal:
# 186,013 here is the export replication; the client's recorded slice is
# 186,412 / $454.2M / ECL $93.5M - different constructions)
_r = spark.sql(f"""
    SELECT round(sum(CASE WHEN in_sas_ledger THEN eop_bal_m1 END), 0) AS eop_bal,
           round(sum(CASE WHEN in_sas_ledger THEN ecl_m1 END), 0) AS ecl
    FROM {DB}.uc2_t16_01s_populations_{ANCHOR_YM}
""").first()
chk("ledger eop_bal_m1 sum", int(_r["eop_bal"] or 0), E["ledger eop_bal_m1 sum"])
chk("ledger ecl_m1 sum", int(_r["ecl"] or 0), E["ledger ecl_m1 sum"])
print("SIDE-BY-SIDE (disclosure, never asserted equal): SAS-recorded slice =")
print("186,412 accounts / $454.2M EOP balance / $93.5M ECL (client-side pivot record).")

# COMMAND ----------

# MAGIC %md
# MAGIC ## S6. Verdict and record block

# COMMAND ----------

sec("S6 verdict")

record_block("B01_sas_spine")
flush_metrics("B01_sas_spine")
