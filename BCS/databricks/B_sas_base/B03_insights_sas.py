# Databricks notebook source
# MAGIC %md
# MAGIC # B03. Insights, SAS-denominated (numbered blocks, each with its tie-out)
# MAGIC
# MAGIC Every denominator is the SAS spine (01s) or the 04s episode table; every
# MAGIC dollar is one of the export's own columns. All values here are MEASURED
# MAGIC on the first run; they enter records only after verification.
# MAGIC
# MAGIC STANDING RULES (kit README, verbatim):
# MAGIC * Balance / CO-dollar sums at episode grain double-count accounts with
# MAGIC   several episodes: first collapse to one row per (group, acct_key).
# MAGIC * An account can sit in two language groups: never add per-group balances
# MAGIC   down to a ledger total.
# MAGIC Plus: the 202,479 / 186,013 / 186,412 triplet is three constructions of
# MAGIC "the SAS population" - each use below names which one it is. 186,013
# MAGIC (export replication) and 186,412 (SAS-recorded slice) are disclosed side
# MAGIC by side, NEVER asserted equal. Nothing here references the 19,789
# MAGIC statement-window population or its recorded metrics (a separate caller
# MAGIC construct, never printed next to these).

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

sec("I0 preconditions")

POP = f"{DB}.uc2_t16_01s_populations_{ANCHOR_YM}"
OUT = f"{DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}"
for _t in [POP, OUT]:
    assert spark.catalog.tableExists(_t), f"PRECONDITION MISS: {_t} missing - run B01/B02b first"
    print(f"PASS  table exists: {_t}")
_r = spark.sql(f"SELECT count_if(in_sas_ledger) AS n FROM {POP}").first()
chk("wf 04 sas ledger", _r["n"], E["wf 04 sas ledger"])

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 1. The population walk with money (waterfall stages x EOP_BAL x ECL)
# MAGIC
# MAGIC 186,013 below is the EXPORT REPLICATION of the SAS slice. The SAS-RECORDED
# MAGIC slice is 186,412 accounts / $454.2M EOP balance / $93.5M ECL (client-side
# MAGIC pivot record). Different constructions, quoted side by side, never
# MAGIC asserted equal.

# COMMAND ----------

sec("Block 1 population walk")

grid("population walk with money (export replication)", spark.sql(f"""
    SELECT '01. export total (610,183 construction)' AS stage,
           count(*) AS accounts,
           round(sum(eop_bal_m1), 0) AS eop_bal_m1,
           round(sum(ecl_m1), 0) AS ecl_m1
    FROM {POP}
    UNION ALL
    SELECT '02. DQ1 (DLNQT_CD_M1 = 1; the 202,479 construction)',
           count(*), round(sum(eop_bal_m1), 0), round(sum(ecl_m1), 0)
    FROM {POP} WHERE wf_dq1
    UNION ALL
    SELECT '03. + CPC eligible',
           count(*), round(sum(eop_bal_m1), 0), round(sum(ecl_m1), 0)
    FROM {POP} WHERE wf_dq1 AND wf_cpc
    UNION ALL
    SELECT '04. + non-chargeoff = the SAS ledger (the 186,013 construction)',
           count(*), round(sum(eop_bal_m1), 0), round(sum(ecl_m1), 0)
    FROM {POP} WHERE in_sas_ledger
    ORDER BY stage
"""))
print("SIDE-BY-SIDE (disclosure): SAS-recorded slice = 186,412 / $454.2M / ECL $93.5M.")

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 2. The funnel, SAS-denominated
# MAGIC
# MAGIC Caller flag = inb_native (any January INBOUND id-resolved call, numeric
# MAGIC key). Episodes and language from the fixed 04s layer. Leaked money =
# MAGIC export columns at account grain.

# COMMAND ----------

sec("Block 2 funnel")

_r = spark.sql(f"""
    SELECT count_if(in_sas_ledger) AS ledger,
           count_if(in_sas_ledger AND inb_native) AS called
    FROM {POP}
""").first()
_s = spark.sql(f"""
    SELECT count(DISTINCT acct_key) AS ep_callers,
           count(*) AS episodes,
           count(DISTINCT CASE WHEN pay_f > 0 THEN acct_key END) AS intent_accts,
           count(DISTINCT CASE WHEN leaked_sas THEN acct_key END) AS leaked_accts
    FROM {OUT}
""").first()
grid("the funnel (SAS ledger, numeric key, captured_sas gate)", spark.sql(f"""
    SELECT '1. the SAS ledger (186,013 construction)' AS step, {_r["ledger"]} AS value
    UNION ALL SELECT '2. called in January (inb_native)', {_r["called"]}
    UNION ALL SELECT '3. accounts with standard episodes (04s)', {_s["ep_callers"]}
    UNION ALL SELECT '4. standard episodes (04s)', {_s["episodes"]}
    UNION ALL SELECT '5. accounts with payment language on >= 1 episode', {_s["intent_accts"]}
    UNION ALL SELECT '6. leaked_sas accounts (no account payment M1/M2)', {_s["leaked_accts"]}
    ORDER BY step
"""))
chk("funnel called (inb_native, ledger)", _r["called"], E["funnel called (inb_native, ledger)"])
chk("funnel callers with episodes", _s["ep_callers"], E["funnel callers with episodes"])
chk("funnel intent accounts", _s["intent_accts"], E["funnel intent accounts"])
chk("funnel leaked accounts", _s["leaked_accts"], E["funnel leaked accounts"])

kv("leaked_sas money (export columns, one row per account)", spark.sql(f"""
    SELECT count(*) AS leaked_accounts,
           round(sum(eop_bal_m1), 0) AS eop_bal_m1,
           round(sum(gross_loss_12m_amt), 0) AS gross_loss_12m,
           round(sum(chrgoff_12m_amt), 0) AS chrgoff_12m
    FROM (SELECT DISTINCT acct_key, eop_bal_m1, gross_loss_12m_amt, chrgoff_12m_amt
          FROM {OUT} WHERE leaked_sas)
"""))

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 3. Language groups with export dollars
# MAGIC
# MAGIC Episode counts are the clean split (they partition). Account and money
# MAGIC columns compare WITHIN a row only: an account can sit in two groups, so
# MAGIC per-group money is NEVER added down a column. Money is collapsed to one
# MAGIC row per (group, account) before summing (the kit's standing dedup rules,
# MAGIC verbatim).

# COMMAND ----------

sec("Block 3 language groups")

_lang_df = spark.sql(f"""
    WITH acct_grp AS (
        SELECT DISTINCT language_group, acct_key, eop_bal_m1, ecl_liftm_m1, gross_loss_12m_amt
        FROM {OUT}
    ),
    ep AS (
        SELECT language_group, count(*) AS episodes, count(DISTINCT acct_key) AS accounts
        FROM {OUT} GROUP BY 1
    ),
    money AS (
        SELECT language_group,
               round(sum(eop_bal_m1), 0) AS eop_bal_m1,
               round(sum(ecl_liftm_m1), 0) AS ecl_liftm_m1,
               round(sum(gross_loss_12m_amt), 0) AS gross_loss_12m
        FROM acct_grp GROUP BY 1
    )
    SELECT e.language_group, e.episodes, e.accounts,
           m.eop_bal_m1, m.ecl_liftm_m1, m.gross_loss_12m
    FROM ep e JOIN money m ON m.language_group = e.language_group
    ORDER BY e.language_group
""")
grid("language groups with export dollars (04s; within-row money only)", _lang_df)
_ep_sum = sum(r["episodes"] for r in _lang_df.collect())
_r = spark.sql(f"SELECT count(*) AS n FROM {OUT}").first()
chk("language episodes re-add to 04s episodes", _ep_sum, _r["n"])

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 4. W_s valued in the client's own columns
# MAGIC
# MAGIC W_s = leaked_sas accounts (no account payment M1/M2, >= 1 payment-language
# MAGIC episode), deceased-language accounts routed out; in_sas_ledger by
# MAGIC construction. Money at account grain, export columns.

# COMMAND ----------

sec("Block 4 W_s")

grid("W_s build steps with export dollars (one row per account)", spark.sql(f"""
    WITH acct AS (
        SELECT DISTINCT acct_key, leaked_sas, w_s_flag, deceased_acct,
               eop_bal_m1, ecl_liftm_m1, gross_loss_12m_amt, chrgoff_12m_amt
        FROM {OUT}
    )
    SELECT '1. leaked_sas accounts' AS step, count(*) AS accounts,
           round(sum(eop_bal_m1), 0) AS eop_bal_m1,
           round(sum(ecl_liftm_m1), 0) AS ecl_liftm_m1,
           round(sum(gross_loss_12m_amt), 0) AS gross_loss_12m,
           round(sum(chrgoff_12m_amt), 0) AS chrgoff_12m
    FROM acct WHERE leaked_sas
    UNION ALL
    SELECT '2. deceased or estate, routed out', count(*),
           round(sum(eop_bal_m1), 0), round(sum(ecl_liftm_m1), 0),
           round(sum(gross_loss_12m_amt), 0), round(sum(chrgoff_12m_amt), 0)
    FROM acct WHERE leaked_sas AND deceased_acct = 1
    UNION ALL
    SELECT '3. W_s, the work list', count(*),
           round(sum(eop_bal_m1), 0), round(sum(ecl_liftm_m1), 0),
           round(sum(gross_loss_12m_amt), 0), round(sum(chrgoff_12m_amt), 0)
    FROM acct WHERE w_s_flag
    ORDER BY step
"""))

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 5. The addressable moment, re-denominated
# MAGIC
# MAGIC The call-day walk-down on the fixed 04s episodes. DISCLOSED CONSTRUCTION
# MAGIC CHANGE: the capture split is the ACCOUNT-grain month-grain captured_sas,
# MAGIC not the old episode-grain 30-day gate; the old steps 3/4 (episode-grain)
# MAGIC do not exist under this gate. Money deduped at account grain.

# COMMAND ----------

sec("Block 5 addressable")

grid("the addressable walk-down (04s, captured_sas split)", spark.sql(f"""
    WITH addr AS (SELECT * FROM {OUT} WHERE is_addressable),
    intent_acct AS (
        SELECT DISTINCT acct_key, captured_sas, eop_bal_m1, gross_loss_12m_amt
        FROM addr WHERE pay_f > 0
    )
    SELECT '1. bucket 1 on the call day (episodes)' AS step,
           (SELECT count(*) FROM addr) AS value
    UNION ALL SELECT '2. from accounts', (SELECT count(DISTINCT acct_key) FROM addr)
    UNION ALL SELECT '3. payment-language episodes', (SELECT count(*) FROM addr WHERE pay_f > 0)
    UNION ALL SELECT '4. payment-language accounts', (SELECT count(*) FROM intent_acct)
    UNION ALL SELECT '5. of 4: account captured_sas (payment in call month or next)',
           (SELECT count(*) FROM intent_acct WHERE captured_sas)
    UNION ALL SELECT '6. of 4: account NOT captured_sas (the addressable moment)',
           (SELECT count(*) FROM intent_acct WHERE NOT captured_sas)
    ORDER BY step
"""))
kv("the addressable moment, money (accounts NOT captured_sas, deduped)", spark.sql(f"""
    WITH addr AS (SELECT * FROM {OUT} WHERE is_addressable),
    intent_acct AS (
        SELECT DISTINCT acct_key, captured_sas, eop_bal_m1, gross_loss_12m_amt
        FROM addr WHERE pay_f > 0
    )
    SELECT count(*) AS accounts,
           round(sum(eop_bal_m1), 0) AS eop_bal_m1,
           round(sum(gross_loss_12m_amt), 0) AS gross_loss_12m
    FROM intent_acct WHERE NOT captured_sas
"""))

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 6. The ECL step M1 -> M2 by caller class (captured_sas classes)
# MAGIC
# MAGIC Population = the SAS ledger (186,013 construction); ECL = the export's
# MAGIC own columns; classes = the captured_sas account classes ('a. non-caller'
# MAGIC = ledger accounts with no 04s episodes). Accounts with a null ECL in
# MAGIC either month are counted separately and excluded from the step sum.

# COMMAND ----------

sec("Block 6 ECL step by class")

grid("ECL step M1 -> M2 by caller class (SAS ledger)", spark.sql(f"""
    WITH cls AS (
        SELECT acct_key, max(caller_class_sas) AS caller_class
        FROM {OUT} GROUP BY 1
    )
    SELECT coalesce(k.caller_class, 'a. non-caller') AS caller_class,
           count(*) AS accounts,
           count_if(p.ecl_m1 IS NOT NULL AND p.ecl_m2 IS NOT NULL) AS accounts_with_both_ecl,
           round(sum(CASE WHEN p.ecl_m1 IS NOT NULL AND p.ecl_m2 IS NOT NULL
                          THEN p.ecl_m2 - p.ecl_m1 END), 0) AS ecl_step_m1_to_m2,
           round(sum(p.ecl_m1), 0) AS ecl_m1,
           round(sum(p.ecl_m2), 0) AS ecl_m2
    FROM {POP} p
    LEFT JOIN cls k ON k.acct_key = p.acct_key
    WHERE p.in_sas_ledger
    GROUP BY 1 ORDER BY 1
"""))

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 7. The ONE SAS x AWS continuity bridge
# MAGIC
# MAGIC The single table where the two populations meet: the SAS ledger
# MAGIC (186,013 construction) against the AWS ex-AA ledger (189,146), account
# MAGIC counts and gate rates over the export universe, plus the one count that
# MAGIC lives outside it (AWS-ledger accounts with no export row). Nothing else
# MAGIC ever mixes the two populations in one table.

# COMMAND ----------

sec("Block 7 continuity bridge")

grid("SAS x AWS continuity bridge (export universe)", spark.sql(f"""
    SELECT in_sas_ledger, aws_in_ledger_exaa,
           count(*) AS accounts,
           count_if(captured_sas) AS captured_sas_accts,
           count_if(aws_captured) AS aws_captured_accts
    FROM {POP}
    GROUP BY 1, 2 ORDER BY 1, 2
"""))
_r = spark.sql(f"""
    SELECT count(*) AS n
    FROM {DB}.uc2_t16_01n_populations p
    LEFT ANTI JOIN {POP} s ON s.acct_key = p.acct_key
    WHERE p.in_ledger_exaa
""").first()
chk("AWS ex-AA ledger accounts with NO export row", _r["n"], None)

# COMMAND ----------

sec("B03 verdict")

record_block("B03_insights_sas")
flush_metrics("B03_insights_sas")
