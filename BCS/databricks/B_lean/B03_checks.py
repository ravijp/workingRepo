# Databricks notebook source
# MAGIC %md
# MAGIC # B03_checks - run ONCE after B03_insights_sas.py.
# MAGIC
# MAGIC RE-ANCHOR NOTE (2026-07-21): SPLIT like B02.
# MAGIC   RAISING (a miss STOPS): funnel step 1 "called (inb_native, ledger)" reads
# MAGIC   the 01s spine (the native caller flag), which is PRE-re-anchor and
# MAGIC   frame-independent - it stays 11,136.
# MAGIC   MEASURE MODE (never STOPS): steps 2-4 (callers-with-episodes, intent,
# MAGIC   leaked) read the re-anchored 04s, so they MOVE. In January steps 1 and 2
# MAGIC   were equal (11,136 = 11,136); after the re-anchor step 2 DROPS below step
# MAGIC   1 (some native callers had only out-of-window episodes) - that divergence
# MAGIC   is informative, not a defect. Reported vs the January reference.

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
    """RAISING - frame-independent anchor. A miss STOPS."""
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


def shift(name, actual, ref):
    """MEASURE MODE - frame-dependent count that moves under the re-anchor."""
    d = actual - ref
    print(f"MEASURED  {name} = {fmt(actual)}   (Jan ref {fmt(ref)}, delta {'+' if d >= 0 else ''}{fmt(d)})")

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
# CORRECTION 2026-07-21: step 1 is NOT frame-independent. inb_native counts
# accounts with an inbound episode in the re-anchored 02n, so it GREW with the
# statement-window widening (Feb/Mar calls admitted). It moves like steps 2-4 ->
# measure-mode, not raising. (The genuinely frame-independent B03 tie is the
# language-partition sum-to-total, asserted inside B03_insights itself.)
shift("funnel called (inb_native, ledger) (statement frame)", _r["called"], E["funnel called (inb_native, ledger)"])
# steps 2-4: read the re-anchored 04s -> MEASURE vs the Jan reference
shift("funnel callers with episodes (statement frame)", _s["ep_callers"], E["funnel callers with episodes"])
shift("funnel intent accounts (statement frame)", _s["intent_accts"], E["funnel intent accounts"])
shift("funnel leaked accounts (statement frame)", _s["leaked_accts"], E["funnel leaked accounts"])

print("B03_checks: DONE. All four funnel steps are frame-DEPENDENT (inb_native and "
      "the 04s counts all grow under the statement re-anchor) and are reported in "
      "measure mode vs the January reference; this file asserts nothing. The "
      "frame-independent tie (language partition re-adds to the total) is asserted "
      "inside B03_insights_sas.")
