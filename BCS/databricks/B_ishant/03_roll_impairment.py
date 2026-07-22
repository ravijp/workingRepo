# Databricks notebook source
# =====================================================================
# B_ishant / 03_roll_impairment.py
# The roll + impairment headline (the client funding number).
#
# THE ROLL
#   DQ1 -> DQ2 roll = dlnqt_cd_m1 = 1 AND dlnqt_cd_m2 = 2, measured on the
#   post-due-called ledger cohort (the 19,025 accounts from 02_windows.py).
#   Meeting cohort ~2,879 (Namit, 21-Jul). Ishant's Excel pivot says ~3,306.
#   [VERIFY: 2,879 vs 3,306 is UNRECONCILED] - we PRINT the raw count and both
#   references; we hardcode neither.
#
# THE IMPAIRMENT (computed on TWO bases, like Ishant's pivot Table 2)
#   Row A = the roll cohort (DQ1->DQ2 within the post-due-called ledger)
#   Row B = the full post-due-called ledger base (~19,025)
#   Measures: sum(ECL_M2), sum(ECL_M3), CO_8M/10M/12M flag counts,
#             sum(CHRGOFF_8M/10M/12M_AMT) = gross charge-off NCL.
#   Meeting references (roll cohort): ECL_M2 ~$4.96M, ECL_M3 ~$7.17M,
#   CO_12M count 1,760-1,980, gross 12M NCL ~$7.599M.
#
# THE STAGE BREAKDOWN
#   STG_CD_M2 (IFRS9 stage at Feb): blank / S1 / S2 / S3 on the roll cohort.
#   Meeting: of ~2,879, ~1,744 already stage-2 at Feb 1 (primarily HRAM);
#   removing HRAM drops 1,744 -> 478; stage-1 ~655.
#   [OPEN: HRAM flag source not in code] - HRAM is an Excel-only daily process,
#   not a column in Ishant's SQL. We give the STG_CD_M2 split and leave a clearly
#   marked placeholder for the HRAM-exclusion cut rather than fabricating it.
#
#   Every print/display below is prefixed with this file name for screenshots.
# =====================================================================

# COMMAND ----------

# ---------------------------------------------------------------------
# SETUP - copied verbatim from B_lean (B00 is the canonical copy). Keep in sync.
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

print(f"[B_ishant/03_roll_impairment.py] SETUP OK: vintage {ANCHOR_YM}; layers -> {DB}")
# ---------------------------------------------------------------------
# end of SETUP
# ---------------------------------------------------------------------

# COMMAND ----------

# ---------------------------------------------------------------------
# Preconditions: 02_windows.py built the fixed population with window flags.
# ---------------------------------------------------------------------
if not spark.catalog.tableExists(f"{DB}.uc2_ish_02s_pop"):
    raise AssertionError(f"[B_ishant/03_roll_impairment.py] {DB}.uc2_ish_02s_pop missing - run 02_windows.py first")
print("[B_ishant/03_roll_impairment.py] preconditions OK: uc2_ish_02s_pop present")

# COMMAND ----------

# ---------------------------------------------------------------------
# 03r. The roll cohort table = post-due-called ledger accounts that rolled
# DQ1 -> DQ2. Carries the impairment / stage columns forward for the cuts.
# ---------------------------------------------------------------------
spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_ish_03r_roll AS
SELECT acct_key, acct_num,
       dlnqt_cd_m1, dlnqt_cd_m2, dlnqt_cd_m3,
       ecl_m1, ecl_m2, ecl_m3,
       stg_cd_m1, stg_cd_m2, stg_cd_m3,
       co_8m_flag, co_10m_flag, co_12m_flag,
       chrgoff_8m_amt, chrgoff_10m_amt, chrgoff_12m_amt,
       gross_loss_8m_amt, gross_loss_10m_amt, gross_loss_12m_amt,
       cpc_class,
       accts_called_post_f,
       (try_cast(dlnqt_cd_m1 AS int) = 1 AND try_cast(dlnqt_cd_m2 AS int) = 2) AS rolled_dq1_dq2
FROM {DB}.uc2_ish_02s_pop
WHERE in_sas_ledger
  AND accts_called_post_f = 1     -- the post-due-called ledger base (~19,025)
""")
print(f"[B_ishant/03_roll_impairment.py] built {DB}.uc2_ish_03r_roll")

# COMMAND ----------

# ---------------------------------------------------------------------
# The roll count. Print raw + BOTH unreconciled references; hardcode neither.
# ---------------------------------------------------------------------
_r = spark.sql(f"""
SELECT count(1) AS post_due_base,
       count_if(rolled_dq1_dq2) AS roll_dq1_dq2
FROM {DB}.uc2_ish_03r_roll
""").first()
print("[B_ishant/03_roll_impairment.py] DQ1->DQ2 roll on the post-due-called ledger:")
print(f"  post-due-called base        : {_r['post_due_base']:>8,}   expected ~19,025")
print(f"  rolled DQ1->DQ2 (raw count) : {_r['roll_dq1_dq2']:>8,}")
print(f"  [VERIFY: UNRECONCILED] Namit (21-Jul) quoted ~2,879; Ishant's pivot ~3,306.")
print(f"  Pin which roll cohort is canonical before quoting the dollar - the funding")
print(f"  ask rests on this number. Neither is hardcoded here.")

# COMMAND ----------

# ---------------------------------------------------------------------
# 03i. Impairment on the roll cohort AND the post-due base (Ishant pivot Table 2).
# ECL sums, forward charge-off counts, and gross charge-off NCL.
# ---------------------------------------------------------------------
print("[B_ishant/03_roll_impairment.py] Impairment: roll cohort vs post-due base:")
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
FROM {DB}.uc2_ish_03r_roll WHERE rolled_dq1_dq2
UNION ALL
SELECT 2, 'B. post-due-called base (~19,025)',
       count(1),
       round(sum(ecl_m2), 0), round(sum(ecl_m3), 0),
       count_if(try_cast(co_8m_flag  AS int) = 1),
       count_if(try_cast(co_10m_flag AS int) = 1),
       count_if(try_cast(co_12m_flag AS int) = 1),
       round(sum(chrgoff_8m_amt), 0), round(sum(chrgoff_10m_amt), 0), round(sum(chrgoff_12m_amt), 0)
FROM {DB}.uc2_ish_03r_roll
ORDER BY row_order
"""))

_i = spark.sql(f"""
SELECT round(sum(ecl_m2), 0) AS ecl_m2, round(sum(ecl_m3), 0) AS ecl_m3,
       count_if(try_cast(co_12m_flag AS int) = 1) AS co_12m,
       round(sum(chrgoff_12m_amt), 0) AS gross_ncl_12m
FROM {DB}.uc2_ish_03r_roll WHERE rolled_dq1_dq2
""").first()
print("[B_ishant/03_roll_impairment.py] roll-cohort headline - actual vs meeting reference:")
print(f"  ECL_M2 sum       : {_i['ecl_m2']:>14,.0f}   ref ~$4.96M")
print(f"  ECL_M3 sum       : {_i['ecl_m3']:>14,.0f}   ref ~$7.17M")
print(f"  CO_12M count     : {_i['co_12m']:>14,}   ref 1,760-1,980")
print(f"  gross 12M NCL    : {_i['gross_ncl_12m']:>14,.0f}   ref ~$7.599M (gross; net of recovery/reversal pending)")

# COMMAND ----------

# ---------------------------------------------------------------------
# 03s. STG_CD_M2 (IFRS9 stage at Feb) breakdown on the roll cohort.
# blank / S1 / S2 / S3. Meeting: ~1,744 already stage-2 at Feb 1.
# ---------------------------------------------------------------------
print("[B_ishant/03_roll_impairment.py] STG_CD_M2 stage breakdown on the DQ1->DQ2 roll cohort:")
display(spark.sql(f"""
SELECT CASE WHEN stg_cd_m2 IS NULL OR trim(stg_cd_m2) = '' THEN '(blank)'
            ELSE upper(trim(stg_cd_m2)) END AS stg_cd_m2,
       count(1)              AS accts,
       round(sum(ecl_m2), 0) AS ecl_m2_sum,
       round(sum(ecl_m3), 0) AS ecl_m3_sum,
       count_if(try_cast(co_12m_flag AS int) = 1) AS co_12m_cnt
FROM {DB}.uc2_ish_03r_roll WHERE rolled_dq1_dq2
GROUP BY 1 ORDER BY 1
"""))

# ---------------------------------------------------------------------
# [OPEN: HRAM flag source not in code]
# The meeting split the ~1,744 stage-2 roll accounts into HRAM vs non-HRAM
# (1,744 -> 478 excluding HRAM). HRAM is an Excel-only daily process; there is
# NO HRAM flag in Ishant's reconstructed SQL nor in the SAS export columns we
# carry. We do NOT fabricate the HRAM-exclusion split. To complete it, source
# the HRAM flag (refit/apollo hram columns exist in the wider SAS export - see
# B_lean/B01 lines 189-192 hram_flag_refit_M2 / hram_flag_apollo_M2) and add a
# stage-2-excluding-HRAM cut here. Left as a marked placeholder.
# ---------------------------------------------------------------------
print("[B_ishant/03_roll_impairment.py] [OPEN: HRAM flag source not in code] "
      "stage-2-excluding-HRAM split (meeting: 1,744 -> 478) NOT computed - "
      "HRAM flag not present in Ishant's SQL; see the hram_flag_refit_M2 / "
      "hram_flag_apollo_M2 SAS columns to complete this. Not fabricated.")

print("[B_ishant/03_roll_impairment.py] 03_roll_impairment complete: uc2_ish_03r_roll")
