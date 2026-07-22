# Databricks notebook source
# =====================================================================
# 03_roll_impairment.py
# The roll + impairment headline (the client funding number).
#
#   Reads the folded ledger uc2_t16_01s_populations_<vintage> (funnel +
#   accts_called_* on one table) and writes uc2_t16_03r_roll_<vintage>.
#
#   uc2_t16_03r_roll_<vintage> is a flat, pivot-ready table: one row per
#   post-due-called ledger account, every delinquency / stage / ECL / charge-off /
#   gross-loss (8m/10m/12m + lifetime) / balance / credit-limit / payment-signal
#   column as its own field, plus the rolled_dq1_dq2 and post_due_base flags and
#   the cpc_class / bucket dimensions. It is the CSV/Excel-pivot source for step 3.
#
# THE ROLL
#   DQ1 -> DQ2 roll = dlnqt_cd_m1 = 1 AND dlnqt_cd_m2 = 2, measured on the
#   post-due-called ledger cohort (the accounts flagged accts_called_31_f).
#   The roll count feeds the funding ask, so it is printed raw and never
#   hardcoded.
#
# THE IMPAIRMENT (computed on two bases)
#   Row A = the roll cohort (DQ1->DQ2 within the post-due-called ledger)
#   Row B = the full post-due-called ledger base
#   Measures: sum(ECL_M2), sum(ECL_M3), CO_8M/10M/12M flag counts,
#             sum(CHRGOFF_8M/10M/12M_AMT) = gross charge-off NCL.
#
# THE STAGE BREAKDOWN
#   STG_CD_M2 (IFRS9 stage at Feb): blank / S1 / S2 / S3 on the roll cohort.
#   The stage-2 accounts include HRAM (high-risk-account-management) cases.
#   The HRAM flag (hram_flag_apollo_M1) is a real CSV column carried through, so
#   the STG_CD_M2 breakdown is computed both ALL and EXCLUDING-HRAM (an account
#   is HRAM when the apollo M1 flag is set: '1'/'Y'/'TRUE').
#
#   Every print/display below is prefixed with this file name for screenshots.
# =====================================================================

# COMMAND ----------

# ---------------------------------------------------------------------
# SETUP - catalog/schema, table handles.
# ---------------------------------------------------------------------
import datetime as _dt

CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
ANCHOR_YM = "202501"
FMT_CATALOG = "634153504162_glue_connection_catalog"
CC_CATALOG = "062108867742_glue_connectivity_catalog"

try:
    dbutils.widgets.text("CATALOG", CATALOG);       CATALOG = dbutils.widgets.get("CATALOG")
    dbutils.widgets.text("SCHEMA", SCHEMA);         SCHEMA = dbutils.widgets.get("SCHEMA")
    dbutils.widgets.text("ANCHOR_YM", ANCHOR_YM);   ANCHOR_YM = dbutils.widgets.get("ANCHOR_YM")
    dbutils.widgets.text("FMT_CATALOG", FMT_CATALOG); FMT_CATALOG = dbutils.widgets.get("FMT_CATALOG")
    dbutils.widgets.text("CC_CATALOG", CC_CATALOG);   CC_CATALOG = dbutils.widgets.get("CC_CATALOG")
except NameError:
    pass

DB = f"{CATALOG}.{SCHEMA}"
NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"

T_01S = f"{DB}.uc2_t16_01s_populations_{ANCHOR_YM}"
T_03R = f"{DB}.uc2_t16_03r_roll_{ANCHOR_YM}"

print(f"[03_roll_impairment.py] SETUP OK: vintage {ANCHOR_YM}; layers -> {DB}")
# ---------------------------------------------------------------------
# end of SETUP
# ---------------------------------------------------------------------

# COMMAND ----------

# ---------------------------------------------------------------------
# 03r. The roll cohort table = post-due-called ledger accounts that rolled
# DQ1 -> DQ2. One row per account on the post-due-called ledger base, carrying
# every delinquency / stage / ECL / charge-off / balance / limit column the roll
# and impairment cuts pivot on. This is a flat, pivot-ready table (one row per
# account, every dimension/measure its own column, no nested structures) and is
# the CSV/Excel-pivot source for step 3. rolled_dq1_dq2 = 1 marks the roll cohort;
# post_due_base = 1 marks the full post-due-called base (the whole table here).
# ---------------------------------------------------------------------
spark.sql(f"""
CREATE OR REPLACE TABLE {T_03R} AS
SELECT acct_key, acct_num,
       -- delinquency codes + buckets, M1/M2/M3
       dlnqt_cd_m1, dlnqt_cd_m2, dlnqt_cd_m3,
       dlnqt_bkt_m1, dlnqt_bkt_m2, dlnqt_bkt_m3,
       -- IFRS9 stage, M1/M2/M3 (STG_CD_M2 drives the stage breakdown)
       stg_cd_m1, stg_cd_m2, stg_cd_m3,
       -- ECL by measurement month + lifetime ECL
       ecl_m1, ecl_m2, ecl_m3, ecl_liftm_m1,
       -- forward charge-off flags (8m/10m/12m)
       co_8m_flag, co_10m_flag, co_12m_flag,
       -- forward charge-off amounts (8m/10m/12m) + lifetime
       chrgoff_8m_amt, chrgoff_10m_amt, chrgoff_12m_amt, chrgoff_amt_lftm,
       -- gross-loss amounts (8m/10m/12m) + lifetime
       gross_loss_8m_amt, gross_loss_10m_amt, gross_loss_12m_amt, gross_loss_amt_lftm,
       -- balances + credit limit (Avg. CL pivots on cr_lmt_m1) + utilization
       eop_bal_m1, eop_bal_m2, eop_bal_m3, cr_lmt_m1, utilization_m1,
       -- payment signals
       paymt_amt_m1, paymt_amt_m2, min_due_amt, paymt_last_amt, last_pay_vs_min_due,
       -- account dimensions
       cpc_class, cpc_flag_nw, max_bucket, eom_bucket,
       -- M0 delinquency bucket (Dec 2024), derived from 00n; carried through
       dlnqt_bkt_m0,
       -- new-roll flag (real CSV column) + HRAM apollo M1 (real CSV column; the
       -- HRAM cut below excludes accounts where the apollo HRAM flag is set)
       new_roll_flag, hram_flag_apollo_m1,
       -- past-due date carried two ways: past_due_last_dt (real observed) and
       -- due_dt_derived (= stmt_dt + 25; the cycle marker, built in 02_windows.py)
       past_due_last_dt,
       -- called-window flags carried from the folded ledger
       accts_called_25_f, accts_called_31_f, accts_called_overall_f,
       -- post-due(31) counts, two grains + indicator (carried from the folded ledger)
       inbound_call_stmnt_dt_25_plus_31_calls,
       inbound_call_stmnt_dt_25_plus_31_episodes,
       inbound_call_stmnt_dt_25_plus_31_ind,
       TRUE AS post_due_base,           -- the whole table is the post-due-called ledger base
       (try_cast(dlnqt_cd_m1 AS int) = 1 AND try_cast(dlnqt_cd_m2 AS int) = 2) AS rolled_dq1_dq2
FROM {T_01S}
WHERE in_sas_ledger
  AND accts_called_31_f = 1     -- 31 = post-due (day 25-55); the post-due-called ledger base
""")
print(f"[03_roll_impairment.py] built {T_03R}")
print(f"[03_roll_impairment.py] {T_03R} is the flat, pivot-ready CSV/Excel source "
      "(one row per post-due-called account; rolled_dq1_dq2 marks the roll cohort; "
      "every roll/impairment dim and measure its own column).")

# COMMAND ----------

# ---------------------------------------------------------------------
# The roll count on the post-due-called ledger. Printed raw; not hardcoded.
# ---------------------------------------------------------------------
_r = spark.sql(f"""
SELECT count(1) AS post_due_base,
       count_if(rolled_dq1_dq2) AS roll_dq1_dq2
FROM {T_03R}
""").first()
print("[03_roll_impairment.py] DQ1->DQ2 roll on the post-due-called ledger:")
print(f"  post-due-called base        : {_r['post_due_base']:>8,}")
print(f"  rolled DQ1->DQ2 (raw count) : {_r['roll_dq1_dq2']:>8,}")
print(f"  [VERIFY: roll-cohort definition] Pin which roll cohort is canonical")
print(f"  before quoting the dollar - the funding ask rests on this number.")

# COMMAND ----------

# ---------------------------------------------------------------------
# 03i. Impairment on the roll cohort AND the post-due base. ECL sums, forward
# charge-off counts, and gross charge-off NCL.
# ---------------------------------------------------------------------
print("[03_roll_impairment.py] Impairment: roll cohort vs post-due base:")
display(spark.sql(f"""
SELECT 1 AS row_order, 'A. roll cohort (DQ1->DQ2)' AS base,
       count(1)                       AS accts,
       round(avg(cr_lmt_m1), 0)       AS avg_cl,
       round(sum(ecl_m2), 0)          AS ecl_m2_sum,
       round(sum(ecl_m3), 0)          AS ecl_m3_sum,
       count_if(try_cast(co_8m_flag  AS int) = 1) AS co_8m_cnt,
       count_if(try_cast(co_10m_flag AS int) = 1) AS co_10m_cnt,
       count_if(try_cast(co_12m_flag AS int) = 1) AS co_12m_cnt,
       round(sum(chrgoff_8m_amt), 0)  AS gross_ncl_8m,
       round(sum(chrgoff_10m_amt), 0) AS gross_ncl_10m,
       round(sum(chrgoff_12m_amt), 0) AS gross_ncl_12m,
       round(sum(chrgoff_amt_lftm), 0) AS gross_ncl_lftm
FROM {T_03R} WHERE rolled_dq1_dq2
UNION ALL
SELECT 2, 'B. post-due-called base',
       count(1),
       round(avg(cr_lmt_m1), 0),
       round(sum(ecl_m2), 0), round(sum(ecl_m3), 0),
       count_if(try_cast(co_8m_flag  AS int) = 1),
       count_if(try_cast(co_10m_flag AS int) = 1),
       count_if(try_cast(co_12m_flag AS int) = 1),
       round(sum(chrgoff_8m_amt), 0), round(sum(chrgoff_10m_amt), 0), round(sum(chrgoff_12m_amt), 0),
       round(sum(chrgoff_amt_lftm), 0)
FROM {T_03R}
ORDER BY row_order
"""))

_i = spark.sql(f"""
SELECT round(sum(ecl_m2), 0) AS ecl_m2, round(sum(ecl_m3), 0) AS ecl_m3,
       count_if(try_cast(co_12m_flag AS int) = 1) AS co_12m,
       round(sum(chrgoff_12m_amt), 0) AS gross_ncl_12m
FROM {T_03R} WHERE rolled_dq1_dq2
""").first()
print("[03_roll_impairment.py] roll-cohort headline:")
print(f"  ECL_M2 sum       : {_i['ecl_m2']:>14,.0f}")
print(f"  ECL_M3 sum       : {_i['ecl_m3']:>14,.0f}")
print(f"  CO_12M count     : {_i['co_12m']:>14,}")
print(f"  gross 12M NCL    : {_i['gross_ncl_12m']:>14,.0f}   (gross; net of recovery/reversal pending)")

# COMMAND ----------

# ---------------------------------------------------------------------
# 03s. STG_CD_M2 (IFRS9 stage at Feb) breakdown on the roll cohort.
# blank / S1 / S2 / S3.
# ---------------------------------------------------------------------
print("[03_roll_impairment.py] STG_CD_M2 stage breakdown on the DQ1->DQ2 roll cohort:")
display(spark.sql(f"""
SELECT CASE WHEN stg_cd_m2 IS NULL OR trim(stg_cd_m2) = '' THEN '(blank)'
            ELSE upper(trim(stg_cd_m2)) END AS stg_cd_m2,
       count(1)              AS accts,
       round(avg(cr_lmt_m1), 0) AS avg_cl,
       round(sum(ecl_m2), 0) AS ecl_m2_sum,
       round(sum(ecl_m3), 0) AS ecl_m3_sum,
       count_if(try_cast(co_8m_flag  AS int) = 1) AS co_8m_cnt,
       count_if(try_cast(co_10m_flag AS int) = 1) AS co_10m_cnt,
       count_if(try_cast(co_12m_flag AS int) = 1) AS co_12m_cnt,
       round(sum(chrgoff_8m_amt), 0)   AS gross_ncl_8m,
       round(sum(chrgoff_10m_amt), 0)  AS gross_ncl_10m,
       round(sum(chrgoff_12m_amt), 0)  AS gross_ncl_12m,
       round(sum(chrgoff_amt_lftm), 0) AS gross_ncl_lftm
FROM {T_03R} WHERE rolled_dq1_dq2
GROUP BY 1 ORDER BY 1
"""))

# ---------------------------------------------------------------------
# 03h. STG_CD_M2 breakdown EXCLUDING HRAM on the roll cohort. An account is HRAM
# when hram_flag_apollo_m1 is set ('1' / 'Y' / 'TRUE'); those accounts are
# dropped, leaving the addressable stage cut. The HRAM flag is a real carried
# CSV column (hram_flag_apollo_M1), read as a string (inferSchema=False).
# ---------------------------------------------------------------------
HRAM_SET = ("upper(trim(cast(hram_flag_apollo_m1 AS string))) IN ('1','Y','TRUE','T')")

print("[03_roll_impairment.py] STG_CD_M2 breakdown on the DQ1->DQ2 roll cohort, EXCLUDING HRAM:")
display(spark.sql(f"""
SELECT CASE WHEN stg_cd_m2 IS NULL OR trim(stg_cd_m2) = '' THEN '(blank)'
            ELSE upper(trim(stg_cd_m2)) END AS stg_cd_m2,
       count(1)              AS accts,
       round(avg(cr_lmt_m1), 0) AS avg_cl,
       round(sum(ecl_m2), 0) AS ecl_m2_sum,
       round(sum(ecl_m3), 0) AS ecl_m3_sum,
       count_if(try_cast(co_8m_flag  AS int) = 1) AS co_8m_cnt,
       count_if(try_cast(co_10m_flag AS int) = 1) AS co_10m_cnt,
       count_if(try_cast(co_12m_flag AS int) = 1) AS co_12m_cnt,
       round(sum(chrgoff_8m_amt), 0)   AS gross_ncl_8m,
       round(sum(chrgoff_10m_amt), 0)  AS gross_ncl_10m,
       round(sum(chrgoff_12m_amt), 0)  AS gross_ncl_12m,
       round(sum(chrgoff_amt_lftm), 0) AS gross_ncl_lftm
FROM {T_03R} WHERE rolled_dq1_dq2 AND NOT ({HRAM_SET})
GROUP BY 1 ORDER BY 1
"""))

_h = spark.sql(f"""
SELECT count(1)                       AS roll_accts,
       count_if({HRAM_SET})           AS hram_accts,
       count_if(NOT ({HRAM_SET}))     AS non_hram_accts,
       count_if(upper(trim(stg_cd_m2)) = 'S2')                       AS s2_all,
       count_if(upper(trim(stg_cd_m2)) = 'S2' AND NOT ({HRAM_SET}))  AS s2_excl_hram
FROM {T_03R} WHERE rolled_dq1_dq2
""").first()
print("[03_roll_impairment.py] HRAM cut on the roll cohort:")
print(f"  roll accounts          : {_h['roll_accts']:>8,}")
print(f"  HRAM (apollo M1 set)   : {_h['hram_accts']:>8,}")
print(f"  non-HRAM               : {_h['non_hram_accts']:>8,}")
print(f"  STG_CD_M2 = S2 (all)   : {_h['s2_all']:>8,}")
print(f"  STG_CD_M2 = S2 (ex-HRAM): {_h['s2_excl_hram']:>7,}")

print("[03_roll_impairment.py] 03_roll_impairment complete: uc2_t16_03r_roll")
