# Databricks notebook source
# =====================================================================
# B_ishant / 01_accounts.py
# The account layer + the SAS funnel (Ishant's client-blessed methodology).
#
# WHAT THIS BUILDS
#   1. uc2_ish_00n_acct_monthly  - monthly FMT delinquency/balance layer
#                                  (reused verbatim from B_lean/B02 K3).
#   2. uc2_ish_01s_ledger        - the SAS-csv funnel:
#                                  610,183 -> 202,479 DQ1 -> 186,848 +CPC
#                                  -> 186,013 non-chargeoff = in_sas_ledger,
#                                  plus every M1/M2/M3 delinquency / ECL / stage /
#                                  charge-off column the roll + impairment step needs,
#                                  and captured_sas (negative PAYMT in M1 or M2).
#
# METHODOLOGY NOTE (why this differs from our old B_lean chain)
#   The client-blessed approach (Ishant, endorsed by Anupam 21-Jul) keeps the
#   population FIXED at all in_sas_ledger accounts and lets calls only CLASSIFY
#   them (see 02_windows.py). There is NO per-call as-of re-anchor here. This
#   module is the frame-INDEPENDENT foundation both approaches share and it ties
#   to the digit on the four funnel numbers.
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

# --- FMT monthly scan window (the four snapshot months Dec24..Mar25) ----------
MONTH_WIN_START = _mm(_a0, -1).strftime("%Y%m%d")   # 20241201
MONTH_WIN_END = _mm(_a0, 3).strftime("%Y%m%d")      # 20250401 (exclusive)

# --- the numeric account key rule (repo convention: strip zero-pad via bigint) --
NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"

print(f"[B_ishant/01_accounts.py] SETUP OK: vintage {ANCHOR_YM}; layers -> {DB}")
# ---------------------------------------------------------------------
# end of SETUP
# ---------------------------------------------------------------------

# COMMAND ----------

# ---------------------------------------------------------------------
# Preconditions: the three source tables must be reachable.
# ---------------------------------------------------------------------
for _t in [f"{FMT_CATALOG}.fmt_acct_dba.fmt_acct_c",
           f"{CC_CATALOG}.contactcenter_bdp_db.call",
           f"{CC_CATALOG}.contactcenter_bdp_db.transcript"]:
    if not spark.catalog.tableExists(_t):
        raise AssertionError(f"[B_ishant/01_accounts.py] source not reachable: {_t}")
print("[B_ishant/01_accounts.py] preconditions OK: FMT / call / transcript reachable")

# COMMAND ----------

# ---------------------------------------------------------------------
# 00n. Monthly FMT account layer (reused verbatim from B_lean/B02 K3).
# One row per (acct_key, ym). Delinquency bucket 0-10 from the past-due
# amount bands; eom_* = end-of-month snapshot; payment/charge-off dates.
# ---------------------------------------------------------------------
spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_ish_00n_acct_monthly AS
WITH snap AS (
    SELECT extnl_acct_id,
           substr(eff_dt, 1, 6) AS ym,
           eff_dt,
           CASE
             WHEN past_due_271_up_amt  > 0 THEN 10
             WHEN past_due_241_270_amt > 0 THEN 9
             WHEN past_due_211_240_amt > 0 THEN 8
             WHEN past_due_181_210_amt > 0 THEN 7
             WHEN past_due_151_180_amt > 0 THEN 6
             WHEN past_due_121_150_amt > 0 THEN 5
             WHEN past_due_91_120_amt  > 0 THEN 4
             WHEN past_due_61_90_amt   > 0 THEN 3
             WHEN past_due_31_60_amt   > 0 THEN 2
             WHEN past_due_1_30_amt    > 0 THEN 1
             ELSE 0
           END AS bucket,
           try_cast(acct_bal_amt AS double) AS bal,
           try_cast(chrgoff_dt AS date) AS co_dt,
           try_cast(chrgoff_amt AS double) AS co_amt,
           clnt_prdct_cd,
           try_cast(cr_lmt_origl_amt AS double) AS cr_lmt_origl_amt,
           coalesce(try_cast(paymt_last_dt AS date),
                    try_to_date(cast(paymt_last_dt AS string), 'ddMMMyyyy')) AS pay_dt,
           coalesce(try_cast(atmtc_paymt_last_dt AS date),
                    try_to_date(cast(atmtc_paymt_last_dt AS string), 'ddMMMyyyy')) AS auto_dt,
           coalesce(try_cast(nsf_last_paymt_dt AS date),
                    try_to_date(cast(nsf_last_paymt_dt AS string), 'ddMMMyyyy')) AS nsf_dt
    FROM {FMT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '{MONTH_WIN_START}' AND eff_dt < '{MONTH_WIN_END}'
)
SELECT {NUM_KEY.format(c="extnl_acct_id")} AS acct_key,   -- numeric-key rule
       ym,
       max(bucket) AS max_bucket,
       max_by(bucket, eff_dt) AS eom_bucket,
       max_by(bal, eff_dt) AS eom_bal,
       min(co_dt) AS mth_co_dt,
       min_by(co_amt, co_dt) AS mth_co_amt,
       min(CASE WHEN bucket >= 1 THEN eff_dt END) AS first_dq_dt,
       min(CASE WHEN bucket = 1 THEN eff_dt END) AS first_b1_dt,
       max_by(clnt_prdct_cd, eff_dt) AS eom_cpc,
       max_by(cr_lmt_origl_amt, eff_dt) AS eom_cr_lmt_origl_amt,
       max(pay_dt) AS pay_dt,
       max(auto_dt) AS auto_dt,
       max(nsf_dt) AS nsf_dt
FROM snap
GROUP BY 1, 2
""")
print(f"[B_ishant/01_accounts.py] built {DB}.uc2_ish_00n_acct_monthly")

print("[B_ishant/01_accounts.py] 00n monthly account count by ym:")
display(spark.sql(f"""
    SELECT ym, count(1) AS accts, count_if(eom_bucket = 1) AS eom_dq1
    FROM {DB}.uc2_ish_00n_acct_monthly GROUP BY ym ORDER BY ym
"""))

# COMMAND ----------

# ---------------------------------------------------------------------
# 01s. The SAS funnel = uc2_ish_01s_ledger.
# Grain: one row per SAS export account (610,183 for 202501).
#
# The three funnel predicates (identical in ours and Ishant's code):
#   wf_dq1    = DLNQT_CD_M1 = 1                          (DQ-1 at Jan EOM)
#   wf_cpc    = CPC_FLAG_NW IN OTHER/OTHERS/COBRAND/PLCC (CPC-eligible; excludes
#                                                         AA/risk-card programs)
#   wf_non_co = CHRGOFF_RSN_M1 blank/PLY/BLANK           (not already charged off)
# in_sas_ledger = wf_dq1 AND wf_cpc AND wf_non_co.
#
# FILTER 1 (Ishant): DQ1-at-Jan-EOM, statement due date in January. The statement
# period (7-Dec-2024 .. 6-Jan-2025) is enforced in 02_windows.py via the statement
# anchor; here we take DLNQT_CD_M1 = 1 as the SAS-side DQ1 definition (M1 = the
# January measurement month).
#
# captured_sas = negative PAYMT_AMT in M1 or M2 (CQ-7, account+month grain).
# ---------------------------------------------------------------------
csv_df = (spark.read.format("csv")
          .option("header", True)
          .option("inferSchema", False)
          .option("mode", "FAILFAST")
          .load(SAS_CSV_PATH))
csv_df.createOrReplaceTempView("_sas_csv")

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_ish_01s_ledger AS
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
)
SELECT cast(f.acct_num AS string) AS acct_key,        -- numeric-key rule
       f.acct_num,
       f.wf_dq1, f.wf_cpc, f.wf_non_co,
       (f.wf_dq1 AND f.wf_cpc AND f.wf_non_co) AS in_sas_ledger,
       -- delinquency codes: M1 = Jan measurement, M2 = Feb, M3 = Mar. DQ1->DQ2
       -- roll = dlnqt_cd_m1 = 1 AND dlnqt_cd_m2 = 2 (computed in 03_roll_impairment)
       f.DLNQT_CD_M1 AS dlnqt_cd_m1, f.DLNQT_CD_M2 AS dlnqt_cd_m2, f.DLNQT_CD_M3 AS dlnqt_cd_m3,
       f.DLNQT_BKT_M1 AS dlnqt_bkt_m1, f.DLNQT_BKT_M2 AS dlnqt_bkt_m2, f.DLNQT_BKT_M3 AS dlnqt_bkt_m3,
       try_cast(f.PAYMT_AMT_M1 AS double) AS paymt_amt_m1,
       try_cast(f.PAYMT_AMT_M2 AS double) AS paymt_amt_m2,
       try_cast(f.PAYMT_AMT_M3 AS double) AS paymt_amt_m3,
       try_cast(f.EOP_BAL_M1 AS double) AS eop_bal_m1,
       try_cast(f.EOP_BAL_M2 AS double) AS eop_bal_m2,
       try_cast(f.EOP_BAL_M3 AS double) AS eop_bal_m3,
       try_cast(f.CR_LMT_M1 AS double) AS cr_lmt_m1,
       -- ECL by measurement month; the roll cohort's M2/M3 ECL is the impairment headline
       try_cast(f.ECL_M0 AS double) AS ecl_m0,
       try_cast(f.ECL_M1 AS double) AS ecl_m1,
       try_cast(f.ECL_M2 AS double) AS ecl_m2,
       try_cast(f.ECL_M3 AS double) AS ecl_m3,
       try_cast(f.ECL_M4 AS double) AS ecl_m4,
       try_cast(f.ECL_LIFTM_M1 AS double) AS ecl_liftm_m1,
       -- IFRS9 stage code by month; STG_CD_M2 drives the stage-2 breakdown
       f.STG_CD_M1 AS stg_cd_m1, f.STG_CD_M2 AS stg_cd_m2, f.STG_CD_M3 AS stg_cd_m3,
       -- 12-month forward charge-off amounts + flags (the NCL headline is 12M)
       try_cast(f.GROSS_LOSS_8M_AMT AS double) AS gross_loss_8m_amt,
       try_cast(f.GROSS_LOSS_10M_AMT AS double) AS gross_loss_10m_amt,
       try_cast(f.GROSS_LOSS_12M_AMT AS double) AS gross_loss_12m_amt,
       try_cast(f.CHRGOFF_8M_AMT AS double) AS chrgoff_8m_amt,
       try_cast(f.CHRGOFF_10M_AMT AS double) AS chrgoff_10m_amt,
       try_cast(f.CHRGOFF_12M_AMT AS double) AS chrgoff_12m_amt,
       try_cast(f.CHRGOFF_AMT_LFTM AS double) AS chrgoff_amt_lftm,
       try_cast(f.GROSS_LOSS_AMT_LFTM AS double) AS gross_loss_amt_lftm,
       f.CO_8M_FLAG AS co_8m_flag, f.CO_10M_FLAG AS co_10m_flag, f.CO_12M_FLAG AS co_12m_flag,
       f.CHRGOFF_RSN_M1 AS chrgoff_rsn_m1,
       f.CPC_FLAG_NW AS cpc_flag_nw,
       -- CPC / department class labels (for the ex-AA / GM / Bronco / CoBrand cuts)
       CASE
         WHEN f.CPC_FLAG_NW IS NULL OR trim(f.CPC_FLAG_NW) = '' THEN 'BLANK'
         ELSE upper(trim(f.CPC_FLAG_NW))
       END AS cpc_class,
       -- captured_sas = a negative PAYMT_AMT in M1 or M2 (payment landed)
       (coalesce(try_cast(f.PAYMT_AMT_M1 AS double), 0) < 0
        OR coalesce(try_cast(f.PAYMT_AMT_M2 AS double), 0) < 0) AS captured_sas
FROM f
""")
print(f"[B_ishant/01_accounts.py] built {DB}.uc2_ish_01s_ledger")

# COMMAND ----------

# ---------------------------------------------------------------------
# Print the funnel inline (screenshot target). Expected vintage-202501 numbers
# are printed as plain lines, NOT asserted - a mismatch prints, never crashes.
# ---------------------------------------------------------------------
_f = spark.sql(f"""
SELECT
    count(1) AS total_accts,
    count_if(wf_dq1) AS dq1,
    count_if(wf_dq1 AND wf_cpc) AS dq1_cpc,
    count_if(in_sas_ledger) AS ledger,
    count_if(captured_sas) AS captured_sas_all,
    count_if(in_sas_ledger AND captured_sas) AS captured_sas_ledger
FROM {DB}.uc2_ish_01s_ledger
""").first()

print("[B_ishant/01_accounts.py] SAS funnel (Ishant Filter 1) - actual vs expected (202501):")
print(f"  total accounts        : {_f['total_accts']:>10,}   expected ~610,183")
print(f"  DQ1 (DLNQT_CD_M1=1)   : {_f['dq1']:>10,}   expected ~202,479")
print(f"  + CPC eligible        : {_f['dq1_cpc']:>10,}   expected ~186,848")
print(f"  + non-chargeoff=ledger: {_f['ledger']:>10,}   expected ~186,013")
print(f"  captured_sas (all)    : {_f['captured_sas_all']:>10,}")
print(f"  captured_sas (ledger) : {_f['captured_sas_ledger']:>10,}   (ref ~125,275)")

print("[B_ishant/01_accounts.py] funnel as a table:")
display(spark.sql(f"""
SELECT 1 AS stage_order, '01. Total accounts'              AS stage, count(1)                       AS accts FROM {DB}.uc2_ish_01s_ledger
UNION ALL
SELECT 2, '02. DQ-1 (DLNQT_CD_M1=1)',            count_if(wf_dq1)                FROM {DB}.uc2_ish_01s_ledger
UNION ALL
SELECT 3, '03. + CPC eligible',                  count_if(wf_dq1 AND wf_cpc)     FROM {DB}.uc2_ish_01s_ledger
UNION ALL
SELECT 4, '04. + non-chargeoff = in_sas_ledger', count_if(in_sas_ledger)         FROM {DB}.uc2_ish_01s_ledger
ORDER BY stage_order
"""))

print("[B_ishant/01_accounts.py] 01_accounts complete: uc2_ish_00n_acct_monthly, uc2_ish_01s_ledger")
