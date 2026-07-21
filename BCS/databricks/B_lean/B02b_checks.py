# Databricks notebook source
# MAGIC %md
# MAGIC # B02b_checks - run ONCE after B02b_outcomes_sas.py to certify.
# MAGIC
# MAGIC Re-reads uc2_t16_04s_outcomes_<vintage> (the 04s episode table B02b built)
# MAGIC and asserts the locked O3 summary: standard January episodes, the callers
# MAGIC behind them, and the account-grain captured_sas / leaked_sas / W_s counts.
# MAGIC captured_sas is ACCOUNT grain, month grain (CQ-7); classes are account
# MAGIC level. leaked_sas = NOT captured_sas AND >= 1 payment-language episode; W_s
# MAGIC = leaked_sas AND non-deceased. A miss STOPS. Rebuilds no logic.

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

E = {
    # B02b (outcomes) - LOCKED 2026-07-16 (8,037 + 1,801 + 1,298 other = 11,136;
    # W_s = 1,801 - 155 routed = 1,646)
    "04s episodes": 13486,
    "04s callers": 11136,
    "04s captured_sas accounts": 8037,
    "04s leaked_sas accounts": 1801,
    "04s W_s accounts": 1646,
}

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
chk("04s episodes", _r["episodes"], E["04s episodes"])
chk("04s callers", _r["callers"], E["04s callers"])
chk("04s captured_sas accounts", _r["captured_accts"], E["04s captured_sas accounts"])
chk("04s leaked_sas accounts", _r["leaked_accts"], E["04s leaked_sas accounts"])
chk("04s W_s accounts", _r["w_s_accts"], E["04s W_s accounts"])

print("B02b_checks: ALL PASS - the lean B02b build is certified equivalent to the locked original.")
