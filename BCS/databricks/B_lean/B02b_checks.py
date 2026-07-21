# Databricks notebook source
# MAGIC %md
# MAGIC # B02b_checks - run ONCE after B02b_outcomes_sas.py.
# MAGIC
# MAGIC Re-reads uc2_t16_04s_outcomes_<vintage> (the 04s episode table B02b built).
# MAGIC
# MAGIC RE-ANCHOR NOTE (2026-07-21): 04s is now built on the STATEMENT-WINDOW
# MAGIC episode set (only call-days in [stmt_dt, stmt_dt+56) survive). Every count
# MAGIC below is conditioned on that in-window caller population, so ALL FIVE are
# MAGIC EXPECTED TO MOVE off the January-frame locked values. They are therefore
# MAGIC printed in MEASURE MODE (fresh value + the old January reference), NOT as
# MAGIC raising asserts. A moved count here is the re-anchor working, not a defect.
# MAGIC The frame-INDEPENDENT anchors (population 204,323 / 189,146, the dollar and
# MAGIC captured_sas/leaked_sas/W_s gate MATH) are asserted in B02_checks, on the
# MAGIC account-grain layers that the window filter does not touch.
# MAGIC captured_sas is ACCOUNT grain, month grain (CQ-7). leaked_sas = NOT
# MAGIC captured_sas AND >= 1 payment-language episode; W_s = leaked_sas AND
# MAGIC non-deceased. This file STOPS on nothing; it reports the shift.

# COMMAND ----------

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

# COMMAND ----------

CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
ANCHOR_YM = "202501"
DB = f"{CATALOG}.{SCHEMA}"
OUT = f"{DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}"

# January-frame REFERENCE values (pre-re-anchor, round-12 record). NOT asserted;
# printed next to the fresh statement-frame value so the shift is visible.
REF_JAN = {
    "04s episodes": 13486,
    "04s callers": 11136,
    "04s captured_sas accounts": 8037,
    "04s leaked_sas accounts": 1801,
    "04s W_s accounts": 1646,
}


def shift(name, actual, ref):
    """measure-mode: print the fresh statement-frame value + the Jan reference
    and the delta. Never raises (these counts move by design)."""
    d = actual - ref
    print(f"MEASURED  {name} = {fmt(actual)}   (Jan ref {fmt(ref)}, delta {'+' if d >= 0 else ''}{fmt(d)})")

# COMMAND ----------

# O3. the locked summary, re-measured off the 04s table (episode grain; the
# caller/class counts are DISTINCT accounts, per the account-grain gate)
assert spark.catalog.tableExists(OUT), \
    f"PRECONDITION MISS: {OUT} missing - run B02b_outcomes_sas first"
_r = spark.sql(f"""
    SELECT count(*) AS episodes,
           count(DISTINCT acct_key) AS callers,
           count(DISTINCT CASE WHEN captured_sas THEN acct_key END) AS captured_accts,
           count(DISTINCT CASE WHEN leaked_sas   THEN acct_key END) AS leaked_accts,
           count(DISTINCT CASE WHEN w_s_flag     THEN acct_key END) AS w_s_accts
    FROM {OUT}
""").first()
shift("04s episodes", _r["episodes"], REF_JAN["04s episodes"])
shift("04s callers", _r["callers"], REF_JAN["04s callers"])
shift("04s captured_sas accounts", _r["captured_accts"], REF_JAN["04s captured_sas accounts"])
shift("04s leaked_sas accounts", _r["leaked_accts"], REF_JAN["04s leaked_sas accounts"])
shift("04s W_s accounts", _r["w_s_accts"], REF_JAN["04s W_s accounts"])

print("B02b_checks: statement-frame counts reported vs the January reference. "
      "These move by design (the re-anchor filters to in-window episodes); this "
      "file asserts nothing. Frame-independent anchors are in B02_checks.")
