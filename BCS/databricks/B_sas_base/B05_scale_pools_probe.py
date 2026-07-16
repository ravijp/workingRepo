# Databricks notebook source
# MAGIC %md
# MAGIC # B05. Scale-round pool measures + the contact-center summary/category probe
# MAGIC
# MAGIC READ-ONLY PROBE: builds no tables (temp views only). Runs any time after
# MAGIC B02b. Purpose (the assistant-at-scale campaign, wave-2 design inputs):
# MAGIC
# MAGIC 1. Re-assert the B04 sampling substrate (the four verified stratum pools)
# MAGIC    so wave-2 quotas are cut on proven ground.
# MAGIC 2. MEASURE the new candidate strata pools:
# MAGIC      A-exec     leaked-intent episodes with payment-mechanics talk (exec_f)
# MAGIC      A-promise  leaked-intent episodes with promise language (promise_f, 03n)
# MAGIC      relaxed-G  no-payment-language episodes, account not captured_sas,
# MAGIC                 transcript exists (the strict C definition MINUS the
# MAGIC                 DLNQT_CD_M2 gate - the strict pool is 14 episodes, too
# MAGIC                 small to mine)
# MAGIC      K-contrast captured_sas episodes with payment language (contrastive
# MAGIC                 pairs need both sides)
# MAGIC      R          transcript coverage of all ledger episodes (random stratum
# MAGIC                 feasibility)
# MAGIC    New pools are RAW-PREDICATE pools (no priority order): they size the
# MAGIC    lanes. The wave-2 sampler assigns priority later so no excerpt is
# MAGIC    sampled twice. Accounts can sit in more than one pool; money compares
# MAGIC    WITHIN a row only and is never added down a column.
# MAGIC 3. PROBE the contact-center summary/category tables (coverage and
# MAGIC    vocabulary only): if they already hold a summary/category per call,
# MAGIC    they are the full-population layer the manual UI cannot be.
# MAGIC    PRINTS NO SUMMARY TEXT: the summary column's de-identification state
# MAGIC    is untested. Aggregates, schemas, and category labels only.

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
# MAGIC ## P1. Preconditions

# COMMAND ----------

sec("P1 preconditions")

OUT = f"{DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}"
SIG = f"{DB}.uc2_t16_03n_signals"
EPI = f"{DB}.uc2_t16_02n_episodes"
SUMMARY_T = f"`{CC_CATALOG}`.contactcenter_bdp_db.summary"
CATEGORY_T = f"`{CC_CATALOG}`.contactcenter_bdp_db.category"

for _t in [OUT, SIG, EPI]:
    assert spark.catalog.tableExists(_t), f"PRECONDITION MISS: {_t} missing - run B02 -> B01 -> B02b first"
    print(f"PASS  table exists: {_t}")
assert spark.catalog.tableExists(f"{CC_CATALOG}.contactcenter_bdp_db.transcript"), \
    "PRECONDITION MISS: transcript table not reachable"
print("PASS  transcript table reachable")

# summary/category reachability is RECORDED, not fatal: the pool measures
# (P2-P4) must land even if the contact-center side tables are absent.
HAS_SUMMARY = spark.catalog.tableExists(f"{CC_CATALOG}.contactcenter_bdp_db.summary")
HAS_CATEGORY = spark.catalog.tableExists(f"{CC_CATALOG}.contactcenter_bdp_db.category")
chk("summary table reachable (1 = yes)", 1 if HAS_SUMMARY else 0, None)
chk("category table reachable (1 = yes)", 1 if HAS_CATEGORY else 0, None)

# COMMAND ----------

# MAGIC %md
# MAGIC ## P2. The B04 substrate guard (priority-ordered stratum pools, verbatim)
# MAGIC
# MAGIC The cb_pool CASE below is VERBATIM from B04 X2 (keep in sync). The four
# MAGIC pool sizes are locked record values (bridge-round-12 addendum 5b,
# MAGIC verified 2026-07-16): A 1,857 / 1,646; B 528 / 497; C 14 / 12;
# MAGIC D 108 / 104. A miss means the sampling substrate moved under the
# MAGIC campaign - STOP and find out why before any wave-2 cut.

# COMMAND ----------

sec("P2 substrate guard")

_r = spark.sql(f"""
    SELECT count(*) AS episodes, count(DISTINCT acct_key) AS callers,
           count(DISTINCT CASE WHEN w_s_flag THEN acct_key END) AS ws_accts
    FROM {OUT}
""").first()
chk("04s episodes", _r["episodes"], E["04s episodes"])
chk("04s callers", _r["callers"], E["04s callers"])
chk("04s W_s accounts", _r["ws_accts"], E["04s W_s accounts"])

# transcript-exists lookup for ALL 04s contactids, ONE bounded pass (the
# same shape as B04's c_with_tx, widened to the full episode set so P3/P4
# reuse it; membership per contactid is identical)
spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW tx_exists AS
    SELECT DISTINCT t.contactid
    FROM {TX} t
    JOIN (SELECT DISTINCT contactid FROM {OUT}) c ON c.contactid = t.contactid
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
    LEFT JOIN tx_exists w ON w.contactid = o.contactid
""")

_pool_df = spark.sql("""
    SELECT stratum, count(*) AS episodes, count(DISTINCT acct_key) AS accounts
    FROM cb_pool
    WHERE stratum IS NOT NULL
    GROUP BY 1 ORDER BY 1
""")
grid("B04 stratum pools (priority-ordered, must match the record)", _pool_df)
_pool = {r["stratum"][0]: (r["episodes"], r["accounts"]) for r in _pool_df.collect()}
# record values: bridge-round-12 addendum 5b (verified 2026-07-16)
for _s, (_ep, _ac) in {"A": (1857, 1646), "B": (528, 497), "C": (14, 12), "D": (108, 104)}.items():
    assert _s in _pool, f"ANCHOR MISS: stratum {_s} pool is EMPTY (record expects {_ep}/{_ac})"
    chk(f"stratum {_s} pool episodes (record 5b)", _pool[_s][0], _ep)
    chk(f"stratum {_s} pool accounts (record 5b)", _pool[_s][1], _ac)

# COMMAND ----------

# MAGIC %md
# MAGIC ## P3. Transcript coverage of the ledger episodes (R-stratum feasibility)

# COMMAND ----------

sec("P3 transcript coverage")

_r = spark.sql(f"""
    SELECT count(*) AS episodes,
           count(w.contactid) AS episodes_with_tx,
           count(DISTINCT CASE WHEN w.contactid IS NOT NULL THEN o.acct_key END) AS accounts_with_tx
    FROM {OUT} o
    LEFT JOIN tx_exists w ON w.contactid = o.contactid
""").first()
chk("04s episodes with transcript", _r["episodes_with_tx"], None)
chk("04s accounts with >= 1 transcript episode", _r["accounts_with_tx"], None)
print(f"note: the R (random) stratum samples from the {fmt(_r['episodes_with_tx'])} "
      f"transcript-backed episodes of {fmt(_r['episodes'])} total")

# COMMAND ----------

# MAGIC %md
# MAGIC ## P4. The new candidate pools (raw predicates, measure mode)
# MAGIC
# MAGIC promise_f / plan_f / hard_f / dispute_f live on 03n (contactid grain);
# MAGIC 04s carries pay_f / deceased_f / exec_f already. Missing 03n row =
# MAGIC no lexicon match = all flags 0.
# MAGIC
# MAGIC Money rule: dollars are ACCOUNT-grain export columns, collapsed to one
# MAGIC row per account BEFORE summing. Accounts can sit in more than one pool:
# MAGIC money compares within a row only, never added down the column.

# COMMAND ----------

sec("P4 new pools")

spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW p04 AS
    SELECT o.*,
           coalesce(x.promise_f, 0) AS promise_f,
           coalesce(x.plan_f, 0)    AS plan_f,
           coalesce(x.hard_f, 0)    AS hard_f,
           coalesce(x.dispute_f, 0) AS dispute_f,
           CASE WHEN w.contactid IS NOT NULL THEN 1 ELSE 0 END AS tx_f,
           CASE WHEN o.pay_f > 0 AND o.leaked_sas AND o.deceased_acct = 0
                THEN 1 ELSE 0 END AS a_raw_f
    FROM {OUT} o
    LEFT JOIN {SIG} x ON x.contactid = o.contactid
    LEFT JOIN tx_exists w ON w.contactid = o.contactid
""")

# consistency ties (provable on the layer definitions; a miss = STOP):
# 1. raw A-pool accounts = W_s accounts (every W_s account has a pay episode)
_r = spark.sql("SELECT count(DISTINCT CASE WHEN a_raw_f = 1 THEN acct_key END) AS n FROM p04").first()
chk("raw A-pool accounts = W_s accounts", _r["n"], E["04s W_s accounts"])
# 2. inside the A pool (deceased_acct = 0 so no episode is deceased), promise
#    episodes are exactly the language-group-b episodes
_r = spark.sql("""
    SELECT count_if(a_raw_f = 1 AND promise_f = 1) AS via_flag,
           count_if(a_raw_f = 1 AND language_group = 'b. future-dated promise') AS via_group
    FROM p04
""").first()
chk("A-promise flag = language group b inside the A pool", _r["via_flag"], _r["via_group"])
# 3. relaxed-G restricted back to the DLNQT_CD_M2 gate = the strict C pool
_r = spark.sql("""
    SELECT count(*) AS episodes, count(DISTINCT acct_key) AS accounts
    FROM p04
    WHERE language_group = 'g. no payment-related language'
      AND NOT captured_sas AND tx_f = 1
      AND try_cast(dlnqt_cd_m2 AS int) = 1
""").first()
chk("relaxed-G with the M2 gate back on = strict C episodes (record 5b)", _r["episodes"], 14)
chk("relaxed-G with the M2 gate back on = strict C accounts (record 5b)", _r["accounts"], 12)

_pool_sql = """
    SELECT {ord} AS ord, '{label}' AS pool,
           (SELECT count(*) FROM p04 WHERE {pred}) AS episodes,
           (SELECT count(DISTINCT acct_key) FROM p04 WHERE {pred}) AS accounts,
           d.eop_bal_m1, d.ecl_liftm_m1, d.gross_loss_12m, d.chrgoff_12m
    FROM (
        SELECT sum(eop_bal_m1) AS eop_bal_m1,
               sum(ecl_liftm_m1) AS ecl_liftm_m1,
               sum(gross_loss_12m_amt) AS gross_loss_12m,
               sum(chrgoff_12m_amt) AS chrgoff_12m
        FROM (SELECT DISTINCT acct_key, eop_bal_m1, ecl_liftm_m1,
                     gross_loss_12m_amt, chrgoff_12m_amt
              FROM p04 WHERE {pred})
    ) d
"""
_pools = [
    ("A. leaked-intent (raw, ties W_s)",       "a_raw_f = 1"),
    ("A-exec. leaked + mechanics talk",        "a_raw_f = 1 AND exec_f = 1"),
    ("A-promise. leaked + promise language",   "a_raw_f = 1 AND promise_f = 1"),
    ("A-exec AND A-promise (overlap)",         "a_raw_f = 1 AND exec_f = 1 AND promise_f = 1"),
    ("G-relaxed. silent, not captured, tx",    "language_group = 'g. no payment-related language' AND NOT captured_sas AND tx_f = 1"),
    ("G-strict (= stratum C, reference)",      "language_group = 'g. no payment-related language' AND NOT captured_sas AND tx_f = 1 AND try_cast(dlnqt_cd_m2 AS int) = 1"),
    ("K. captured contrast, payment talk",     "captured_sas AND pay_f > 0"),
    ("K-exec. captured + mechanics talk",      "captured_sas AND exec_f = 1"),
    ("K-promise. captured + promise language", "captured_sas AND promise_f = 1"),
    ("R. all ledger episodes (reference)",     "1 = 1"),
    ("R-tx. ledger episodes with transcript",  "tx_f = 1"),
]
grid("scale-round pools with export dollars (money within-row ONLY)", spark.sql(
    "SELECT pool, episodes, accounts, eop_bal_m1, ecl_liftm_m1, gross_loss_12m, chrgoff_12m FROM (\n"
    + "\nUNION ALL\n".join(_pool_sql.format(ord=_i, label=_l, pred=_p)
                           for _i, (_l, _p) in enumerate(_pools))
    + "\n) ORDER BY ord"
))

for _label, _pred, _name in [
    ("A-exec", "a_raw_f = 1 AND exec_f = 1", "A-exec pool"),
    ("A-promise", "a_raw_f = 1 AND promise_f = 1", "A-promise pool"),
    ("G-relaxed", "language_group = 'g. no payment-related language' AND NOT captured_sas AND tx_f = 1", "relaxed-G pool"),
    ("K", "captured_sas AND pay_f > 0", "captured contrast pool"),
]:
    _r = spark.sql(f"SELECT count(*) AS e, count(DISTINCT acct_key) AS a FROM p04 WHERE {_pred}").first()
    chk(f"{_name} episodes", _r["e"], None)
    chk(f"{_name} accounts", _r["a"], None)

# the relax factor, visible: what dropping the M2 gate adds, by M2 state
grid("relaxed-G by DLNQT_CD_M2 (the dropped gate, made visible)", spark.sql("""
    SELECT coalesce(cast(try_cast(dlnqt_cd_m2 AS int) AS string), 'null') AS dlnqt_cd_m2,
           count(*) AS episodes, count(DISTINCT acct_key) AS accounts
    FROM p04
    WHERE language_group = 'g. no payment-related language'
      AND NOT captured_sas AND tx_f = 1
    GROUP BY 1 ORDER BY 1
"""))

# COMMAND ----------

# MAGIC %md
# MAGIC ## P5. The contact-center summary/category probe (coverage + vocabulary)
# MAGIC
# MAGIC If these tables carry a usable summary/category per call, they are the
# MAGIC full-population layer: the assistant UI then validates a SAMPLE of the
# MAGIC labels instead of reading raw transcripts.
# MAGIC
# MAGIC RULES: NO summary text is printed or displayed anywhere in this section
# MAGIC (its de-identification state is untested). Aggregates, schemas, lengths,
# MAGIC and category LABELS only. Known schemas from the 2026-07-13 DESCRIBE:
# MAGIC category(contactid, beginmillis, endmillis, category, participantid,
# MAGIC sentiment, calldate, effdt); summary(contactid, summary, year, month,
# MAGIC day). Column guards below re-check on this platform and skip loudly on
# MAGIC a mismatch instead of failing the sitting.

# COMMAND ----------

sec("P5a summary probe")

if not HAS_SUMMARY:
    print("SKIP: summary table not reachable on this platform (recorded in P1)")
else:
    grid("summary schema (DESCRIBE)", spark.sql(f"DESCRIBE TABLE {SUMMARY_T}"))
    _cols = [c.lower() for c in spark.sql(f"SELECT * FROM {SUMMARY_T} LIMIT 0").columns]
    _need = {"contactid", "summary", "year", "month"}
    if not _need.issubset(set(_cols)):
        print(f"SKIP: summary columns {sorted(_need - set(_cols))} missing; observed {_cols}")
    else:
        # partition-friendly month predicate; IN-list covers zero-padded and
        # plain month spellings, and coerces if the column is typed int
        _janw = f"year IN ('{ANCHOR_YM[:4]}') AND month IN ('{int(ANCHOR_YM[4:6])}', '{ANCHOR_YM[4:6]}')"
        spark.sql(f"""
            CREATE OR REPLACE TEMP VIEW sum_jan AS
            SELECT contactid, length(summary) AS summary_len
            FROM {SUMMARY_T}
            WHERE {_janw} AND summary IS NOT NULL
        """)
        _r = spark.sql("SELECT count(*) AS rows, count(DISTINCT contactid) AS cids FROM sum_jan").first()
        chk("summary rows (anchor month)", _r["rows"], None)
        chk("summary distinct contactids (anchor month)", _r["cids"], None)
        kv("summary length stats (chars; aggregates only)", spark.sql("""
            SELECT min(summary_len) AS min_len,
                   cast(avg(summary_len) AS int) AS avg_len,
                   cast(approx_percentile(summary_len, 0.5) AS int) AS p50_len,
                   cast(approx_percentile(summary_len, 0.9) AS int) AS p90_len,
                   max(summary_len) AS max_len
            FROM sum_jan
        """))
        _r = spark.sql(f"""
            SELECT count(*) AS episodes, count(s.contactid) AS with_summary
            FROM {OUT} o
            LEFT JOIN (SELECT DISTINCT contactid FROM sum_jan) s ON s.contactid = o.contactid
        """).first()
        chk("04s episodes with a summary row", _r["with_summary"], None)
        print(f"note: coverage {fmt(_r['with_summary'])} of {fmt(_r['episodes'])} ledger episodes")
        grid("04s summary coverage by caller class (episodes)", spark.sql(f"""
            SELECT o.caller_class_sas, count(*) AS episodes, count(s.contactid) AS with_summary
            FROM {OUT} o
            LEFT JOIN (SELECT DISTINCT contactid FROM sum_jan) s ON s.contactid = o.contactid
            GROUP BY 1 ORDER BY 1
        """))
        _r = spark.sql(f"""
            SELECT count(*) AS std_episodes, count(s.contactid) AS with_summary
            FROM (SELECT DISTINCT contactid FROM {EPI} WHERE is_episode_std = 1) e
            LEFT JOIN (SELECT DISTINCT contactid FROM sum_jan) s ON s.contactid = e.contactid
        """).first()
        chk("02n standard episodes with a summary row", _r["with_summary"], None)
        print(f"note: broader universe {fmt(_r['with_summary'])} of {fmt(_r['std_episodes'])} standard January episodes")

# COMMAND ----------

sec("P5b category probe")

if not HAS_CATEGORY:
    print("SKIP: category table not reachable on this platform (recorded in P1)")
else:
    grid("category schema (DESCRIBE)", spark.sql(f"DESCRIBE TABLE {CATEGORY_T}"))
    _cols = [c.lower() for c in spark.sql(f"SELECT * FROM {CATEGORY_T} LIMIT 0").columns]
    _need = {"contactid", "category", "effdt"}
    if not _need.issubset(set(_cols)):
        print(f"SKIP: category columns {sorted(_need - set(_cols))} missing; observed {_cols}")
    else:
        spark.sql(f"""
            CREATE OR REPLACE TEMP VIEW cat_jan AS
            SELECT contactid, category{', sentiment' if 'sentiment' in _cols else ''}
            FROM {CATEGORY_T}
            WHERE effdt >= '{EFFDT_CAP_START}' AND effdt < '{EFFDT_CAP_END}'
              AND category IS NOT NULL
        """)
        _r = spark.sql("""
            SELECT count(*) AS rows, count(DISTINCT contactid) AS cids,
                   count(DISTINCT category) AS labels
            FROM cat_jan
        """).first()
        chk("category rows (effdt window)", _r["rows"], None)
        chk("category distinct contactids (effdt window)", _r["cids"], None)
        chk("category distinct labels (effdt window)", _r["labels"], None)
        _r = spark.sql(f"""
            SELECT count(*) AS episodes, count(c.contactid) AS with_category
            FROM {OUT} o
            LEFT JOIN (SELECT DISTINCT contactid FROM cat_jan) c ON c.contactid = o.contactid
        """).first()
        chk("04s episodes with >= 1 category row", _r["with_category"], None)
        print(f"note: coverage {fmt(_r['with_category'])} of {fmt(_r['episodes'])} ledger episodes")
        grid("category vocabulary, top 40 by rows (labels only)", spark.sql("""
            SELECT category, count(*) AS rows, count(DISTINCT contactid) AS contactids
            FROM cat_jan
            GROUP BY 1 ORDER BY 2 DESC, 1 LIMIT 40
        """))
        if "sentiment" in _cols:
            grid("sentiment distribution (effdt window)", spark.sql("""
                SELECT coalesce(cast(sentiment AS string), 'null') AS sentiment,
                       count(*) AS rows, count(DISTINCT contactid) AS contactids
                FROM cat_jan
                GROUP BY 1 ORDER BY 2 DESC, 1
            """))
        grid("top 25 categories on ledger episodes, by capture outcome", spark.sql(f"""
            SELECT c.category,
                   count(DISTINCT o.contactid) AS episodes,
                   count(DISTINCT CASE WHEN o.captured_sas THEN o.contactid END) AS captured_eps,
                   count(DISTINCT CASE WHEN o.leaked_sas THEN o.contactid END) AS leaked_eps
            FROM cat_jan c
            JOIN {OUT} o ON o.contactid = c.contactid
            GROUP BY 1 ORDER BY 2 DESC, 1 LIMIT 25
        """))

# COMMAND ----------

sec("P6 verdict")

record_block("B05_scale_pools_probe")
flush_metrics("B05_scale_pools_probe")
