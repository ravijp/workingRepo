# Databricks notebook source
# MAGIC %md
# MAGIC # B02b. Outcomes on the SAS spine: `uc2_t16_04s_outcomes_<vintage>`
# MAGIC
# MAGIC Grain: one row per standard January episode (numeric key) on an account in
# MAGIC the SAS ledger (in_sas_ledger, 186,013 for 202501).
# MAGIC Headline gate captured_sas: ACCOUNT grain, month grain (CQ-7). No
# MAGIC episode-grain "captured" under this gate; classes are account-level.
# MAGIC leaked_sas = NOT captured_sas AND >= 1 payment-language episode.
# MAGIC W_s = leaked_sas AND non-deceased. AWS day-grain gate rides as aws_
# MAGIC diagnostics only. Run B02b_checks.py once after this to certify.

# COMMAND ----------

# =====================================================================
# SETUP - keep in sync across B00/B01/B02/B02b/B03 (B00 is the canonical copy).
# B02b reads only tables, so the derived-window block is not needed here.
# =====================================================================
CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
ANCHOR_YM = "202501"
SAS_CSV_PATH = "/Volumes/cda_model_shared/ecm_cld_model/ecm_cld/collections_zenon/WATERFALL_COLL_CALL_V2_202501.csv"
FMT_CATALOG = "634153504162_glue_connection_catalog"
CC_CATALOG = "062108867742_glue_connectivity_catalog"

try:
    dbutils.widgets.text("CATALOG", CATALOG);           CATALOG = dbutils.widgets.get("CATALOG")
    dbutils.widgets.text("SCHEMA", SCHEMA);             SCHEMA = dbutils.widgets.get("SCHEMA")
    dbutils.widgets.text("ANCHOR_YM", ANCHOR_YM);       ANCHOR_YM = dbutils.widgets.get("ANCHOR_YM")
    dbutils.widgets.text("SAS_CSV_PATH", SAS_CSV_PATH); SAS_CSV_PATH = dbutils.widgets.get("SAS_CSV_PATH")
    dbutils.widgets.text("FMT_CATALOG", FMT_CATALOG);   FMT_CATALOG = dbutils.widgets.get("FMT_CATALOG")
    dbutils.widgets.text("CC_CATALOG", CC_CATALOG);     CC_CATALOG = dbutils.widgets.get("CC_CATALOG")
except NameError:
    pass

DB = f"{CATALOG}.{SCHEMA}"
FMT = f"`{FMT_CATALOG}`.fmt_acct_dba.fmt_acct_c"
CALL = f"`{CC_CATALOG}`.contactcenter_bdp_db.`call`"
TX = f"`{CC_CATALOG}`.contactcenter_bdp_db.transcript"

# B02b's 04s build reads only tables (no date literals of its own), so the
# derived-window variables of the canonical B00 SETUP are omitted here.
NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"

print(f"SETUP OK: vintage {ANCHOR_YM}; layers -> {DB}")
# =====================================================================
# end of SETUP
# =====================================================================

# COMMAND ----------

# O1. preconditions
for _t in ["uc2_t16_02n_episodes", "uc2_t16_04n_outcomes", f"uc2_t16_01s_populations_{ANCHOR_YM}"]:
    assert spark.catalog.tableExists(f"{DB}.{_t}"), \
        f"PRECONDITION MISS: {DB}.{_t} missing - run B02 then B01 first"

# COMMAND ----------

# MAGIC %md
# MAGIC ## O2. Build `uc2_t16_04s_outcomes_<vintage>`

# COMMAND ----------

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_04s_outcomes_{ANCHOR_YM} AS
WITH s AS (
    SELECT acct_key, captured_sas, dlnqt_cd_m2, dlnqt_cd_m3, stg_cd_m1, stg_cd_m2,
           cpc_flag_nw, eop_bal_m1, eop_bal_m2, ecl_m1, ecl_m2, ecl_liftm_m1,
           gross_loss_12m_amt, chrgoff_12m_amt, gross_loss_amt_lftm, chrgoff_amt_lftm
    FROM {DB}.uc2_t16_01s_populations_{ANCHOR_YM}
    WHERE in_sas_ledger
),
j AS (
    SELECT e.acct_key, e.contactid, e.call_dt,
           e.language_group, e.pay_f, e.deceased_f, e.exec_f, e.has_tx,
           e.callday_bucket, e.is_addressable,
           e.captured    AS aws_captured_episode,
           e.any_captured = 1 AS aws_captured,
           e.caller_class AS aws_caller_class,
           s.captured_sas,
           s.dlnqt_cd_m2, s.dlnqt_cd_m3, s.stg_cd_m1, s.stg_cd_m2, s.cpc_flag_nw,
           s.eop_bal_m1, s.eop_bal_m2, s.ecl_m1, s.ecl_m2, s.ecl_liftm_m1,
           s.gross_loss_12m_amt, s.chrgoff_12m_amt, s.gross_loss_amt_lftm, s.chrgoff_amt_lftm
    FROM {DB}.uc2_t16_04n_outcomes e
    JOIN s ON s.acct_key = e.acct_key
),
acct AS (
    SELECT acct_key, max(pay_f) AS any_pay, max(deceased_f) AS deceased_acct
    FROM j GROUP BY 1
)
SELECT j.*,
       a.deceased_acct,
       CASE
         WHEN j.captured_sas THEN 'b. captured (account payment in call month or next)'
         WHEN a.any_pay > 0 THEN 'c. leaked-intent (payment language, no account payment M1/M2)'
         ELSE 'd. other-caller'
       END AS caller_class_sas,
       (NOT j.captured_sas AND a.any_pay > 0) AS leaked_sas,
       (NOT j.captured_sas AND a.any_pay > 0 AND a.deceased_acct = 0) AS w_s_flag
FROM j
JOIN acct a ON a.acct_key = j.acct_key
""")
print(f"built {DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}")

# COMMAND ----------

print("B02b_outcomes_sas build complete: uc2_t16_04s_outcomes_" + ANCHOR_YM
      + ". Run B02b_checks.py once to certify the episode/caller/captured_sas/"
        "leaked_sas/W_s anchors and the flagged-overlap containment stop-rule.")
