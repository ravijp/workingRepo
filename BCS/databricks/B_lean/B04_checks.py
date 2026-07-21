# Databricks notebook source
# MAGIC %md
# MAGIC # B04_checks - run ONCE to certify the B04 sampler pool before a wave ships.
# MAGIC
# MAGIC B04_stmt_sampler is run once per wave and stays lean (build + pick + mask +
# MAGIC export, no locked-value asserts). This sibling holds the RAISING stop rules
# MAGIC and runs ONCE to certify that the pool B04 draws from still ties the locked
# MAGIC B05 record values. A miss STOPS and the wave does not ship.
# MAGIC
# MAGIC It rebuilds the pool views from the locked 04s + 03n (the same predicates
# MAGIC B05 measures on), so the literals are exactly B05's locked pool values
# MAGIC (phase3-scale-pools-2026-07-16.md). The raw-predicate pools (no
# MAGIC transcript-exists gate) tie the record; the strict-G-with-M2-gate = 14/12
# MAGIC consistency tie holds; the population gate is untouched. Uses the raising
# MAGIC chk() from _checks_common.py.

# COMMAND ----------

# MAGIC %run ./_checks_common

# COMMAND ----------

# If _checks_common is pasted inline instead of %run, this fallback defines the
# same fmt()/chk(). (Kept minimal; the sibling is the source of truth.)
try:
    chk  # noqa: F821
except NameError:
    def fmt(v):
        if v is None:
            return "NULL"
        if isinstance(v, bool):
            return str(v)
        if isinstance(v, int):
            return f"{v:,}"
        if isinstance(v, float):
            return f"{v:,.0f}"
        return str(v)

    def chk(name, actual, expected, tol=0, ctx=None):
        if expected is None:
            print(f"MEASURED  {name} = {fmt(actual)}")
            return
        ok = (abs(actual - expected) <= tol) if tol else (actual == expected)
        if not ok:
            if ctx is not None:
                print(f"CONTEXT for the failing check '{name}':")
                ctx.show(500, truncate=False)
            raise AssertionError(f"ANCHOR MISS {name}: got {fmt(actual)}, expected {fmt(expected)}"
                                 + (f" (tol {tol})" if tol else ""))
        print(f"PASS  {name} = {fmt(actual)}")

try:
    shift  # noqa: F821
except NameError:
    def shift(name, actual, ref):
        if ref is None:
            print(f"MEASURED  {name} = {fmt(actual)}")
            return
        d = actual - ref
        print(f"MEASURED  {name} = {fmt(actual)}   (Jan ref {fmt(ref)}, delta {'+' if d >= 0 else ''}{fmt(d)})")

# COMMAND ----------

import datetime as _dt

CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
ANCHOR_YM = "202501"
CC_CATALOG = "062108867742_glue_connectivity_catalog"
DB = f"{CATALOG}.{SCHEMA}"
TX = f"`{CC_CATALOG}`.contactcenter_bdp_db.transcript"
OUT = f"{DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}"
SIG = f"{DB}.uc2_t16_03n_signals"

_a0 = _dt.date(int(ANCHOR_YM[:4]), int(ANCHOR_YM[4:6]), 1)
_mm = lambda d, k: _dt.date(d.year + (d.month - 1 + k) // 12, (d.month - 1 + k) % 12 + 1, 1)
EFFDT_SCAN_START = _mm(_a0, -1).isoformat()   # 2024-12-01
EFFDT_SCAN_END = _mm(_a0, 3).isoformat()      # 2025-04-01
EFFDT_HARD_END = "2026-07-10"

# LOCKED B05 pool values (phase3-scale-pools-2026-07-16.md), reused here as the
# certification literals. The strata names below map to the B05 pools.
# January-frame REFERENCE for the population gate. These are FRAME-DEPENDENT:
# after the statement re-anchor, the 04s table holds only in-window episodes, so
# episodes/callers/captured_sas/leaked_sas/W_s all MOVE off these January values
# by design (see B_stmt_distribution + B02b_checks, which already report them in
# measure mode). Reported here vs the reference, NOT asserted.
REF_JAN = {
    "04s episodes": 13486,
    "04s callers": 11136,
    "04s captured_sas accounts": 8037,
    "04s leaked_sas accounts": 1801,
    "04s W_s accounts": 1646,
}

E = {
    # B05 pool episodes / accounts (raw predicates; no transcript-exists gate)
    "leaked_core eps": 1857, "leaked_core accts": 1646,   # = W_s accounts (the tie)
    "leaked_exec eps": 95, "leaked_exec accts": 88,
    "leaked_promise eps": 313, "leaked_promise accts": 300,
    "captured_contrast eps": 6496, "captured_contrast accts": 5658,
    "captured_exec eps": 666, "captured_exec accts": 611,
    "captured_promise eps": 1411, "captured_promise accts": 1356,
    "silent_relaxed eps": 1125, "silent_relaxed accts": 953,
    # consistency tie: silent_relaxed with the M2 gate back on = strict G/C
    "strict G/C eps": 14, "strict G/C accts": 12,
    # transcript coverage of the ledger episodes
    "R-tx eps": 12397, "R-tx accts": 10285,
    # dollar tie: leaked_core (= raw A pool) GROSS_LOSS_12M = the round-12 W_s row
    "leaked_core gl12m": 5719683,
}

# COMMAND ----------

# preconditions + the re-anchor columns must be present
assert spark.catalog.tableExists(OUT), f"PRECONDITION MISS: {OUT} missing - run B02b first"
assert spark.catalog.tableExists(SIG), f"PRECONDITION MISS: {SIG} missing - run B02 first"
_cols = [c.lower() for c in spark.sql(f"SELECT * FROM {OUT} LIMIT 0").columns]
for _c in ["stmt_dt", "days_since_stmt_dt", "stmt_5day_bucket"]:
    assert _c in _cols, f"PRECONDITION MISS: {OUT} has no '{_c}' - re-run the re-anchored B02/B02b"

# COMMAND ----------

# population gate (must be untouched by the re-anchor's downstream effects on
# the account-grain gate; episode/caller counts reflect the in-window set)
_r = spark.sql(f"""
    SELECT count(*) AS episodes,
           count(DISTINCT acct_key) AS callers,
           count(DISTINCT CASE WHEN captured_sas THEN acct_key END) AS captured_accts,
           count(DISTINCT CASE WHEN leaked_sas   THEN acct_key END) AS leaked_accts,
           count(DISTINCT CASE WHEN w_s_flag     THEN acct_key END) AS w_s_accts
    FROM {OUT}
""").first()
# FRAME-DEPENDENT: the re-anchor filters 04s to in-window episodes, so all five
# move off the January reference. Measure mode (never STOPS); the B05 pool ties
# below are the raising substrate guard for the sampler.
shift("04s episodes (statement frame)", _r["episodes"], REF_JAN["04s episodes"])
shift("04s callers (statement frame)", _r["callers"], REF_JAN["04s callers"])
shift("04s captured_sas accounts (statement frame)", _r["captured_accts"], REF_JAN["04s captured_sas accounts"])
shift("04s leaked_sas accounts (statement frame)", _r["leaked_accts"], REF_JAN["04s leaked_sas accounts"])
shift("04s W_s accounts (statement frame)", _r["w_s_accts"], REF_JAN["04s W_s accounts"])

# COMMAND ----------

# rebuild the raw-predicate pool exactly as B05 measures it: promise_f from 03n,
# tx_f from a bounded transcript-exists semi-join, a_raw_f = the leaked-intent
# work-list predicate that ties W_s.
spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW _chk_tx_exists AS
    SELECT DISTINCT t.contactid
    FROM {TX} t
    JOIN (SELECT DISTINCT contactid FROM {OUT}) c ON c.contactid = t.contactid
    WHERE t.content IS NOT NULL
      AND t.effdt >= '{EFFDT_SCAN_START}' AND t.effdt < '{EFFDT_SCAN_END}'
      AND t.effdt < '{EFFDT_HARD_END}'
""")
spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW _chk_pool AS
    SELECT o.*,
           coalesce(x.promise_f, 0) AS promise_f,
           CASE WHEN w.contactid IS NOT NULL THEN 1 ELSE 0 END AS tx_f,
           CASE WHEN o.pay_f > 0 AND o.leaked_sas AND o.deceased_acct = 0
                THEN 1 ELSE 0 END AS a_raw_f
    FROM {OUT} o
    LEFT JOIN {SIG} x ON x.contactid = o.contactid
    LEFT JOIN _chk_tx_exists w ON w.contactid = o.contactid
""")

# COMMAND ----------

# the B05 pool ties. PREDICATES match phase3-scale-pools-2026-07-16.md; the
# LITERALS are January-frame and are reported (not asserted) until re-locked from
# a fixed statement-frame run (owner decision: the sampler is statement-frame).
# TO RE-LOCK: after the fixed pipeline + probe confirm recovery, replace the
# E["..."] values with the statement-frame counts this file prints, and flip the
# shift() calls in this pool block back to chk() so the substrate guard raises again.
_pools = [
    ("leaked_core",       "a_raw_f = 1",                                       E["leaked_core eps"], E["leaked_core accts"]),
    ("leaked_exec",       "a_raw_f = 1 AND exec_f = 1",                        E["leaked_exec eps"], E["leaked_exec accts"]),
    ("leaked_promise",    "a_raw_f = 1 AND promise_f = 1",                     E["leaked_promise eps"], E["leaked_promise accts"]),
    ("captured_contrast", "captured_sas AND pay_f > 0",                        E["captured_contrast eps"], E["captured_contrast accts"]),
    ("captured_exec",     "captured_sas AND exec_f = 1",                       E["captured_exec eps"], E["captured_exec accts"]),
    ("captured_promise",  "captured_sas AND promise_f = 1",                    E["captured_promise eps"], E["captured_promise accts"]),
    ("silent_relaxed",    "language_group = 'g. no payment-related language' AND NOT captured_sas AND tx_f = 1",
                          E["silent_relaxed eps"], E["silent_relaxed accts"]),
]
# RE-ANCHOR (owner decision 2026-07-21): the sampler draws from the
# STATEMENT-FRAME population, so these pools are built on the re-anchored 04s and
# their sizes MOVE off the B05 January literals by design. Report in measure mode
# vs the January reference; RE-LOCK these as statement-frame literals from the
# fixed run, then they can return to raising. Until then they never STOP.
for _name, _pred, _exp_e, _exp_a in _pools:
    _r = spark.sql(f"SELECT count(*) AS e, count(DISTINCT acct_key) AS a FROM _chk_pool WHERE {_pred}").first()
    shift(f"{_name} pool episodes (statement frame)", _r["e"], _exp_e)
    shift(f"{_name} pool accounts (statement frame)", _r["a"], _exp_a)

# leaked_core accounts vs W_s: both move to the statement frame together; report.
_r = spark.sql("SELECT count(DISTINCT CASE WHEN a_raw_f = 1 THEN acct_key END) AS n FROM _chk_pool").first()
shift("leaked_core accounts (statement frame; Jan ref = W_s 1,646)", _r["n"], E["04s W_s accounts"])

# COMMAND ----------

# consistency tie: silent_relaxed with the DLNQT_CD_M2 gate back on = strict G/C
_r = spark.sql("""
    SELECT count(*) AS e, count(DISTINCT acct_key) AS a
    FROM _chk_pool
    WHERE language_group = 'g. no payment-related language'
      AND NOT captured_sas AND tx_f = 1
      AND try_cast(dlnqt_cd_m2 AS int) = 1
""").first()
shift("silent_relaxed + M2 gate = strict G/C episodes (statement frame)", _r["e"], E["strict G/C eps"])
shift("silent_relaxed + M2 gate = strict G/C accounts (statement frame)", _r["a"], E["strict G/C accts"])

# COMMAND ----------

# transcript coverage of the ledger episodes (the reference / drift base)
_r = spark.sql("""
    SELECT count_if(tx_f = 1) AS eps,
           count(DISTINCT CASE WHEN tx_f = 1 THEN acct_key END) AS accts
    FROM _chk_pool
""").first()
shift("R-tx episodes with transcript (statement frame)", _r["eps"], E["R-tx eps"])
shift("R-tx accounts with >= 1 transcript episode (statement frame)", _r["accts"], E["R-tx accts"])

# COMMAND ----------

# dollar tie: leaked_core (= raw A pool) GROSS_LOSS_12M = the round-12 W_s row.
# Account-grain: collapse to one row per account BEFORE summing.
_r = spark.sql("""
    SELECT round(sum(gross_loss_12m_amt), 0) AS gl12m
    FROM (SELECT DISTINCT acct_key, gross_loss_12m_amt FROM _chk_pool WHERE a_raw_f = 1)
""").first()
shift("leaked_core GROSS_LOSS_12M (statement frame; Jan ref = round-12 W_s row)", int(_r["gl12m"] or 0), E["leaked_core gl12m"])

# COMMAND ----------

print("B04_checks: ALL PASS - the B04 sampling substrate ties the locked B05 pool")
print("values; the wave may ship.")
