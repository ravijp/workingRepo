# Databricks notebook source
# MAGIC %md
# MAGIC # B01. The SAS spine: `uc2_t16_01s_populations_<vintage>` (runs AFTER B02)
# MAGIC
# MAGIC Grain: one row per export account (610,183 for 202501). Population,
# MAGIC delinquency, and dollars from the client's SAS 003 export.
# MAGIC LOOP CUT: no call_type_* column is read; inb_native is rebuilt natively
# MAGIC (any January INBOUND id-resolved 02n row under the numeric key).
# MAGIC captured_sas: a negative PAYMT_AMT in M1 or M2 (CQ-7, confirmed by A11);
# MAGIC account grain, month grain. aws_ columns are diagnostics, never denominators.
# MAGIC Run B01_checks.py once after this to certify the ladders, tie-out, dollars.

# COMMAND ----------

# =====================================================================
# SETUP - keep in sync across B00/B01/B02/B02b/B03 (B00 is the canonical copy).
# =====================================================================
import datetime as _dt

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

_a0 = _dt.date(int(ANCHOR_YM[:4]), int(ANCHOR_YM[4:6]), 1)
_mm = lambda d, k: _dt.date(d.year + (d.month - 1 + k) // 12, (d.month - 1 + k) % 12 + 1, 1)

PRV_YM = _mm(_a0, -1).strftime("%Y%m")
FEB_YM = _mm(_a0, 1).strftime("%Y%m")
MAR_YM = _mm(_a0, 2).strftime("%Y%m")
MONTH_WIN_START = _mm(_a0, -1).strftime("%Y%m%d")
MONTH_WIN_END = _mm(_a0, 3).strftime("%Y%m%d")
CALL_WIN_START = _a0.isoformat()
CALL_WIN_END = _mm(_a0, 1).isoformat()
EFFDT_CAP_START = _a0.isoformat()
EFFDT_CAP_END = (_mm(_a0, 1) + _dt.timedelta(days=1)).isoformat()
CLEANUP_DATE = _a0.isoformat()
ANCHOR_EOM = (_mm(_a0, 1) - _dt.timedelta(days=1)).isoformat()
FEB_START = _mm(_a0, 1).isoformat()
MAR_START = _mm(_a0, 2).isoformat()
APR_START = _mm(_a0, 3).isoformat()
CO8_END = (_mm(_a0, 9) - _dt.timedelta(days=1)).isoformat()
CO10_END = (_mm(_a0, 11) - _dt.timedelta(days=1)).isoformat()
CO12_END = (_mm(_a0, 13) - _dt.timedelta(days=1)).isoformat()
FWD_CO_START = _a0.strftime("%Y%m%d")
FWD_CO_END = _mm(_a0, 12).strftime("%Y%m%d")
SNAP_DAILY_START = _mm(_a0, -7).strftime("%Y%m%d")
SNAP_DAILY_END = _mm(_a0, 1).strftime("%Y%m%d")
EFFDT_SCAN_START = _mm(_a0, -1).isoformat()
EFFDT_HARD_END = "2026-07-10"   # not vintage-derived: the live-loading-edge guard

NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"

print(f"SETUP OK: vintage {ANCHOR_YM}; layers -> {DB}")
# =====================================================================
# end of SETUP
# =====================================================================

# COMMAND ----------

# S1. preconditions: the fixed-key n-layers exist
for _t in ["uc2_t16_00n_acct_monthly", "uc2_t16_01n_populations",
           "uc2_t16_02n_episodes", "uc2_t16_04n_outcomes"]:
    assert spark.catalog.tableExists(f"{DB}.{_t}"), \
        f"PRECONDITION MISS: {DB}.{_t} missing - run B02_keyfix_aws_layers first"

# COMMAND ----------

# S2. CSV load (all-string, FAILFAST)
csv_df = (spark.read.format("csv")
          .option("header", True)
          .option("inferSchema", False)
          .option("mode", "FAILFAST")
          .load(SAS_CSV_PATH))
csv_df.createOrReplaceTempView("_sas_csv")

# COMMAND ----------

# MAGIC %md
# MAGIC ## S4. Build `uc2_t16_01s_populations_<vintage>`
# MAGIC Explicit column-by-column SELECT. NO call_type_* read. aws_ columns joined
# MAGIC BY NUMERIC KEY from the FIXED n-build; diagnostics, never denominators.

# COMMAND ----------

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_01s_populations_{ANCHOR_YM} AS
WITH k AS (
    SELECT r.*, try_cast(EXTNL_ACCT_ID AS bigint) AS acct_num
    FROM _sas_csv r
),
f AS (
    SELECT k.*,
           coalesce(try_cast(DLNQT_CD_M1 AS int) = 1, false) AS wf_dq1,
           coalesce(upper(trim(CPC_FLAG_NW)) IN ('OTHER', 'OTHERS', 'COBRAND', 'PLCC'), false) AS wf_cpc,
           (CHRGOFF_RSN_M1 IS NULL OR trim(CHRGOFF_RSN_M1) = ''
            OR upper(trim(CHRGOFF_RSN_M1)) IN ('PLY', 'BLANK')) AS wf_non_co
    FROM k
),
inb AS (SELECT DISTINCT acct_key FROM {DB}.uc2_t16_02n_episodes),
awsp AS (SELECT acct_key, in_ledger_all, in_ledger_exaa, touched_b1, eom_bal
         FROM {DB}.uc2_t16_01n_populations),
awsc AS (SELECT acct_key, max(is_episode_std) AS has_std FROM {DB}.uc2_t16_02n_episodes GROUP BY 1),
awscap AS (SELECT acct_key, max(any_captured) AS any_captured FROM {DB}.uc2_t16_04n_outcomes GROUP BY 1)
SELECT cast(f.acct_num AS string) AS acct_key,
       f.acct_num,
       f.wf_dq1, f.wf_cpc, f.wf_non_co,
       (f.wf_dq1 AND f.wf_cpc AND f.wf_non_co) AS in_sas_ledger,
       f.DLNQT_CD_M1 AS dlnqt_cd_m1, f.DLNQT_CD_M2 AS dlnqt_cd_m2, f.DLNQT_CD_M3 AS dlnqt_cd_m3,
       f.DLNQT_BKT_M1 AS dlnqt_bkt_m1, f.DLNQT_BKT_M2 AS dlnqt_bkt_m2, f.DLNQT_BKT_M3 AS dlnqt_bkt_m3,
       try_cast(f.PAYMT_AMT_M1 AS double) AS paymt_amt_m1,
       try_cast(f.PAYMT_AMT_M2 AS double) AS paymt_amt_m2,
       try_cast(f.PAYMT_AMT_M3 AS double) AS paymt_amt_m3,
       try_cast(f.EOP_BAL_M1 AS double) AS eop_bal_m1,
       try_cast(f.EOP_BAL_M2 AS double) AS eop_bal_m2,
       try_cast(f.EOP_BAL_M3 AS double) AS eop_bal_m3,
       try_cast(f.CR_LMT_M1 AS double) AS cr_lmt_m1,
       try_cast(f.CR_LMT_M2 AS double) AS cr_lmt_m2,
       try_cast(f.CR_LMT_M3 AS double) AS cr_lmt_m3,
       try_cast(f.GROSS_LOSS_M1 AS double) AS gross_loss_m1,
       try_cast(f.GROSS_LOSS_M2 AS double) AS gross_loss_m2,
       try_cast(f.GROSS_LOSS_M3 AS double) AS gross_loss_m3,
       try_cast(f.CHRGOFF_AMT_M1 AS double) AS chrgoff_amt_m1,
       try_cast(f.CHRGOFF_AMT_M2 AS double) AS chrgoff_amt_m2,
       try_cast(f.CHRGOFF_AMT_M3 AS double) AS chrgoff_amt_m3,
       try_cast(f.CHRGOFF_RVRSL_M1 AS double) AS chrgoff_rvrsl_m1,
       try_cast(f.CHRGOFF_RVRSL_M2 AS double) AS chrgoff_rvrsl_m2,
       try_cast(f.CHRGOFF_RVRSL_M3 AS double) AS chrgoff_rvrsl_m3,
       try_cast(f.PLCY_LOSS_M1 AS double) AS plcy_loss_m1,
       try_cast(f.PLCY_LOSS_M2 AS double) AS plcy_loss_m2,
       try_cast(f.PLCY_LOSS_M3 AS double) AS plcy_loss_m3,
       try_cast(f.GROSS_LOSS_8M_AMT AS double) AS gross_loss_8m_amt,
       try_cast(f.GROSS_LOSS_10M_AMT AS double) AS gross_loss_10m_amt,
       try_cast(f.GROSS_LOSS_12M_AMT AS double) AS gross_loss_12m_amt,
       try_cast(f.CHRGOFF_8M_AMT AS double) AS chrgoff_8m_amt,
       try_cast(f.CHRGOFF_10M_AMT AS double) AS chrgoff_10m_amt,
       try_cast(f.CHRGOFF_12M_AMT AS double) AS chrgoff_12m_amt,
       try_cast(f.PLCY_LOSS_8M_AMT AS double) AS plcy_loss_8m_amt,
       try_cast(f.PLCY_LOSS_10M_AMT AS double) AS plcy_loss_10m_amt,
       try_cast(f.PLCY_LOSS_12M_AMT AS double) AS plcy_loss_12m_amt,
       try_cast(f.CHRGOFF_AMT_LFTM AS double) AS chrgoff_amt_lftm,
       try_cast(f.GROSS_LOSS_AMT_LFTM AS double) AS gross_loss_amt_lftm,
       f.CHRGOFF_VAL_FLAG AS chrgoff_val_flag,
       try_cast(f.ECL_M0 AS double) AS ecl_m0,
       try_cast(f.ECL_M1 AS double) AS ecl_m1,
       try_cast(f.ECL_M2 AS double) AS ecl_m2,
       try_cast(f.ECL_M3 AS double) AS ecl_m3,
       try_cast(f.ECL_M4 AS double) AS ecl_m4,
       try_cast(f.ECL_12MO_M0 AS double) AS ecl_12mo_m0,
       try_cast(f.ECL_12MO_M1 AS double) AS ecl_12mo_m1,
       try_cast(f.ECL_12MO_M2 AS double) AS ecl_12mo_m2,
       try_cast(f.ECL_12MO_M3 AS double) AS ecl_12mo_m3,
       try_cast(f.ECL_12MO_M4 AS double) AS ecl_12mo_m4,
       try_cast(f.ECL_LIFTM_M0 AS double) AS ecl_liftm_m0,
       try_cast(f.ECL_LIFTM_M1 AS double) AS ecl_liftm_m1,
       try_cast(f.ECL_LIFTM_M2 AS double) AS ecl_liftm_m2,
       try_cast(f.ECL_LIFTM_M3 AS double) AS ecl_liftm_m3,
       try_cast(f.ECL_LIFTM_M4 AS double) AS ecl_liftm_m4,
       f.STG_CD_M0 AS stg_cd_m0, f.STG_CD_M1 AS stg_cd_m1, f.STG_CD_M2 AS stg_cd_m2,
       f.STG_CD_M3 AS stg_cd_m3, f.STG_CD_M4 AS stg_cd_m4,
       try_cast(f.WRITE_OFF_M0 AS double) AS write_off_m0,
       try_cast(f.WRITE_OFF_M1 AS double) AS write_off_m1,
       try_cast(f.WRITE_OFF_M2 AS double) AS write_off_m2,
       try_cast(f.WRITE_OFF_M3 AS double) AS write_off_m3,
       try_cast(f.WRITE_OFF_M4 AS double) AS write_off_m4,
       f.CHRGOFF_RSN_M1 AS chrgoff_rsn_m1, f.CHRGOFF_RSN_M2 AS chrgoff_rsn_m2, f.CHRGOFF_RSN_M3 AS chrgoff_rsn_m3,
       f.CPC_FLAG_NW AS cpc_flag_nw,
       f.CO_CURRENT_FLAG AS co_current_flag, f.CO_8M_FLAG AS co_8m_flag,
       f.CO_10M_FLAG AS co_10m_flag, f.CO_12M_FLAG AS co_12m_flag,
       f.REAGE_EVER_FLAG AS reage_ever_flag,
       f.NEW_ROLL_FLAG AS new_roll_flag, f.NO_PRIOR_RECORD_FLAG AS no_prior_record_flag,
       f.hram_flag_refit_M1 AS hram_flag_refit_m1, f.hram_flag_refit_M2 AS hram_flag_refit_m2,
       f.hram_flag_refit_M3 AS hram_flag_refit_m3,
       f.hram_flag_apollo_M1 AS hram_flag_apollo_m1, f.hram_flag_apollo_M2 AS hram_flag_apollo_m2,
       f.hram_flag_apollo_M3 AS hram_flag_apollo_m3,
       (coalesce(try_cast(f.PAYMT_AMT_M1 AS double), 0) < 0
        OR coalesce(try_cast(f.PAYMT_AMT_M2 AS double), 0) < 0) AS captured_sas,
       (i.acct_key IS NOT NULL) AS inb_native,
       coalesce(p.in_ledger_all, false)  AS aws_in_ledger_all,
       coalesce(p.in_ledger_exaa, false) AS aws_in_ledger_exaa,
       coalesce(p.touched_b1, false)     AS aws_touched_b1,
       p.eom_bal                          AS aws_jan_eom_bal,
       coalesce(c.has_std, 0) = 1         AS aws_caller,
       coalesce(cap.any_captured, 0) = 1  AS aws_captured
FROM f
LEFT JOIN inb i     ON i.acct_key = cast(f.acct_num AS string)
LEFT JOIN awsp p    ON p.acct_key = cast(f.acct_num AS string)
LEFT JOIN awsc c    ON c.acct_key = cast(f.acct_num AS string)
LEFT JOIN awscap cap ON cap.acct_key = cast(f.acct_num AS string)
""")
print(f"built {DB}.uc2_t16_01s_populations_{ANCHOR_YM}")

# COMMAND ----------

print("B01_sas_spine build complete: uc2_t16_01s_populations_" + ANCHOR_YM
      + ". Run B01_checks.py once to certify the waterfall, native ladder, "
        "CSV-flag tie-out, captured_sas, and dollar sums.")
