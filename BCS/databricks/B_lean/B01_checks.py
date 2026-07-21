# Databricks notebook source
# MAGIC %md
# MAGIC # B01_checks - run ONCE after B01_sas_spine.py to certify.
# MAGIC
# MAGIC Re-reads uc2_t16_01s_populations_202501 (and notebook A's wf table for the
# MAGIC one-time tie-out). Runs the CQ-7 sign tripwire, the waterfall + native
# MAGIC ladder + captured_sas asserts, the perfectly-diagonal CSV-flag tie-out, and
# MAGIC the ledger dollar sums. A miss STOPS. Rebuilds no logic.

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

E = {
    "csv rows": 610183,
    "wf 02 dq1": 202479,
    "wf 03 +cpc": 186848,
    "wf 04 sas ledger": 186013,
    "inb_native 01 total": 34234,
    "inb_native 02 dq1": 12615,
    "inb_native 03 +cpc": 11289,
    "inb_native 04 ledger": 11136,
    "captured_sas all": 278885,
    "captured_sas ledger": 125275,
    "ledger eop_bal_m1 sum": 452444591,
    "ledger ecl_m1 sum": 93543576,
    # the perfectly-diagonal one-time tie-out cells (inb_native x csv_inb)
    "tie all F/F": 575949,
    "tie all T/T": 34234,
    "tie ledger F/F": 174877,
    "tie ledger T/T": 11136,
}

# COMMAND ----------

# S3. PAYMT sign tripwire (CQ-7): negatives must dominate M1/M2, or STOP.
_r = spark.sql(f"""
    SELECT count_if(paymt_amt_m1 < 0) AS neg_m1, count_if(paymt_amt_m1 > 0) AS pos_m1,
           count_if(paymt_amt_m2 < 0) AS neg_m2, count_if(paymt_amt_m2 > 0) AS pos_m2
    FROM {POP}
""").first()
assert _r["neg_m1"] > _r["pos_m1"] and _r["neg_m2"] > _r["pos_m2"], (
    "SIGN CONVENTION CONTRADICTED: negatives do not dominate PAYMT_AMT M1/M2. "
    "STOP AND INVESTIGATE (CQ-7 pre-registration).")
print("PASS  sign convention holds: payments are NEGATIVE PAYMT_AMT values (CQ-7)")

# COMMAND ----------

# S5. waterfall + native ladder + captured_sas
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
    FROM {POP}
""").first()
chk("wf 01 total (rows)", _r["s1"], E["csv rows"])
chk("wf 02 dq1", _r["s2"], E["wf 02 dq1"])
chk("wf 03 +cpc", _r["s3"], E["wf 03 +cpc"])
chk("wf 04 sas ledger", _r["s4"], E["wf 04 sas ledger"])
chk("inb_native 01 total", _r["n1"], E["inb_native 01 total"])
chk("inb_native 02 dq1", _r["n2"], E["inb_native 02 dq1"])
chk("inb_native 03 +cpc", _r["n3"], E["inb_native 03 +cpc"])
chk("inb_native 04 ledger", _r["n4"], E["inb_native 04 ledger"])
chk("captured_sas all", _r["cap_all"], E["captured_sas all"])
chk("captured_sas ledger", _r["cap_ledger"], E["captured_sas ledger"])

# COMMAND ----------

# the ONE-TIME tie-out vs the CSV flag: must be PERFECTLY DIAGONAL (zero
# off-diagonal). Asserts the four diagonal cells and that off-diagonal = 0.
if spark.catalog.tableExists(f"{DB}.uc2_sas_wf_202501"):
    def _tie(where):
        return {(r["inb_native"], r["csv_inb"]): r["accounts"] for r in spark.sql(f"""
            SELECT s.inb_native, w.csv_inb, count(*) AS accounts
            FROM {POP} s JOIN {DB}.uc2_sas_wf_202501 w ON w.acct_num = s.acct_num
            {where}
            GROUP BY 1, 2
        """).collect()}

    _all = _tie("")
    chk("tie-out all: F/F", _all.get((False, False), 0), E["tie all F/F"])
    chk("tie-out all: T/T", _all.get((True, True), 0), E["tie all T/T"])
    chk("tie-out all: off-diagonal (must be zero)",
        _all.get((False, True), 0) + _all.get((True, False), 0), 0)
    _led = _tie("WHERE s.in_sas_ledger")
    chk("tie-out ledger: F/F", _led.get((False, False), 0), E["tie ledger F/F"])
    chk("tie-out ledger: T/T", _led.get((True, True), 0), E["tie ledger T/T"])
    chk("tie-out ledger: off-diagonal (must be zero)",
        _led.get((False, True), 0) + _led.get((True, False), 0), 0)
else:
    print("tie-out skipped (notebook A's uc2_sas_wf_202501 absent)")

# COMMAND ----------

# ledger dollar sums (export replication; never asserted equal to the recorded slice)
_r = spark.sql(f"""
    SELECT round(sum(CASE WHEN in_sas_ledger THEN eop_bal_m1 END), 0) AS eop_bal,
           round(sum(CASE WHEN in_sas_ledger THEN ecl_m1 END), 0) AS ecl
    FROM {POP}
""").first()
chk("ledger eop_bal_m1 sum", int(_r["eop_bal"] or 0), E["ledger eop_bal_m1 sum"])
chk("ledger ecl_m1 sum", int(_r["ecl"] or 0), E["ledger ecl_m1 sum"])

print("B01_checks: ALL PASS - the lean B01 build is certified equivalent to the locked original.")
