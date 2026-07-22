# Databricks notebook source
# =====================================================================
# 03_roll_impairment.py
# The roll + impairment headline (the client funding number).
#
#   Reads the folded ledger uc2_t16_01s_populations_<vintage> (funnel +
#   accts_called_* on one table) and writes uc2_t16_03r_roll_<vintage>.
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
#   The stage-2 accounts include HRAM (high-risk-account-management) cases;
#   HRAM is run as an Excel-only daily process and has no column in the SAS
#   export carried here, so the HRAM-exclusion cut is left as a marked open item
#   rather than fabricated.
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
# DQ1 -> DQ2. Carries the impairment / stage columns forward for the cuts.
# ---------------------------------------------------------------------
spark.sql(f"""
CREATE OR REPLACE TABLE {T_03R} AS
SELECT acct_key, acct_num,
       dlnqt_cd_m1, dlnqt_cd_m2, dlnqt_cd_m3,
       ecl_m1, ecl_m2, ecl_m3,
       stg_cd_m1, stg_cd_m2, stg_cd_m3,
       co_8m_flag, co_10m_flag, co_12m_flag,
       chrgoff_8m_amt, chrgoff_10m_amt, chrgoff_12m_amt,
       gross_loss_8m_amt, gross_loss_10m_amt, gross_loss_12m_amt,
       cpc_class,
       accts_called_31_f,               -- 31 = post-due (day 25-55)
       (try_cast(dlnqt_cd_m1 AS int) = 1 AND try_cast(dlnqt_cd_m2 AS int) = 2) AS rolled_dq1_dq2
FROM {T_01S}
WHERE in_sas_ledger
  AND accts_called_31_f = 1     -- the post-due-called ledger base
""")
print(f"[03_roll_impairment.py] built {T_03R}")

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
       round(sum(ecl_m2), 0)          AS ecl_m2_sum,
       round(sum(ecl_m3), 0)          AS ecl_m3_sum,
       count_if(try_cast(co_8m_flag  AS int) = 1) AS co_8m_cnt,
       count_if(try_cast(co_10m_flag AS int) = 1) AS co_10m_cnt,
       count_if(try_cast(co_12m_flag AS int) = 1) AS co_12m_cnt,
       round(sum(chrgoff_8m_amt), 0)  AS gross_ncl_8m,
       round(sum(chrgoff_10m_amt), 0) AS gross_ncl_10m,
       round(sum(chrgoff_12m_amt), 0) AS gross_ncl_12m
FROM {T_03R} WHERE rolled_dq1_dq2
UNION ALL
SELECT 2, 'B. post-due-called base',
       count(1),
       round(sum(ecl_m2), 0), round(sum(ecl_m3), 0),
       count_if(try_cast(co_8m_flag  AS int) = 1),
       count_if(try_cast(co_10m_flag AS int) = 1),
       count_if(try_cast(co_12m_flag AS int) = 1),
       round(sum(chrgoff_8m_amt), 0), round(sum(chrgoff_10m_amt), 0), round(sum(chrgoff_12m_amt), 0)
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
       round(sum(ecl_m2), 0) AS ecl_m2_sum,
       round(sum(ecl_m3), 0) AS ecl_m3_sum,
       count_if(try_cast(co_12m_flag AS int) = 1) AS co_12m_cnt
FROM {T_03R} WHERE rolled_dq1_dq2
GROUP BY 1 ORDER BY 1
"""))

# ---------------------------------------------------------------------
# [OPEN: HRAM flag source] The stage-2 roll accounts include HRAM
# (high-risk-account-management) cases that should be excluded from the
# addressable cut. HRAM is an Excel-only daily process; there is no HRAM flag
# in the SAS export columns carried here, so the stage-2-excluding-HRAM cut is
# not computed. To complete it, source the HRAM flag (the hram_flag_refit_M2 /
# hram_flag_apollo_M2 columns exist in the wider SAS export) and add a
# stage-2-excluding-HRAM cut. Left as a marked open item; not fabricated.
# ---------------------------------------------------------------------
print("[03_roll_impairment.py] [OPEN: HRAM flag source] stage-2-excluding-HRAM "
      "split not computed - HRAM flag not present in the carried SAS columns; "
      "source hram_flag_refit_M2 / hram_flag_apollo_M2 to complete it. Not fabricated.")

print("[03_roll_impairment.py] 03_roll_impairment complete: uc2_t16_03r_roll")
