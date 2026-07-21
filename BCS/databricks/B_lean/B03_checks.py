# Databricks notebook source
# MAGIC %md
# MAGIC # B03_checks - run ONCE after B03_insights_sas.py to certify.
# MAGIC
# MAGIC Re-reads the SAS spine (01s) and the 04s episode table and asserts the
# MAGIC locked B03 funnel: accounts called in January (inb_native, in the SAS
# MAGIC ledger), accounts with standard episodes, accounts with payment language,
# MAGIC and leaked_sas accounts. Funnel steps 2 and 3 are both 11,136: a MEASURED
# MAGIC equality (every native caller in the SAS ledger has a standard episode),
# MAGIC not an a-priori identity. A miss STOPS. Rebuilds no logic.

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
POP = f"{DB}.uc2_t16_01s_populations_{ANCHOR_YM}"
OUT = f"{DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}"

E = {
    # B03 (insights) - LOCKED 2026-07-16 (funnel steps 2 and 3 are both 11,136:
    # a MEASURED equality - every native caller in the SAS ledger has a standard
    # episode - not an a-priori identity)
    "funnel called (inb_native, ledger)": 11136,
    "funnel callers with episodes": 11136,
    "funnel intent accounts": 7459,
    "funnel leaked accounts": 1801,
}

# COMMAND ----------

# the funnel, re-measured off the 01s spine and the 04s table (same queries the
# B03 core runs in Block 2)
for _t in [POP, OUT]:
    assert spark.catalog.tableExists(_t), f"PRECONDITION MISS: {_t} missing - run B01/B02b/B03 first"
_r = spark.sql(f"""
    SELECT count_if(in_sas_ledger AND inb_native) AS called
    FROM {POP}
""").first()
_s = spark.sql(f"""
    SELECT count(DISTINCT acct_key) AS ep_callers,
           count(DISTINCT CASE WHEN pay_f > 0 THEN acct_key END) AS intent_accts,
           count(DISTINCT CASE WHEN leaked_sas THEN acct_key END) AS leaked_accts
    FROM {OUT}
""").first()
chk("funnel called (inb_native, ledger)", _r["called"], E["funnel called (inb_native, ledger)"])
chk("funnel callers with episodes", _s["ep_callers"], E["funnel callers with episodes"])
chk("funnel intent accounts", _s["intent_accts"], E["funnel intent accounts"])
chk("funnel leaked accounts", _s["leaked_accts"], E["funnel leaked accounts"])

print("B03_checks: ALL PASS - the lean B03 build is certified equivalent to the locked original.")
