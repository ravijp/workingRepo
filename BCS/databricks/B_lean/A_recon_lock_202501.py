# Databricks notebook source
# MAGIC %md
# MAGIC # A. Reconciliation lock, vintage 202501 (FROZEN; build-only lean copy)
# MAGIC
# MAGIC One-time reconciliation for 202501. Consolidates the round-11 caller-gap
# MAGIC evidence, replicates the export waterfall, probes the PAYMT_AMT sign
# MAGIC convention, and builds the one-time capture-gate delta table. Constants,
# MAGIC not widgets. Runs alone; run A_recon_lock_checks.py after it to certify.
# MAGIC
# MAGIC GO-FORWARD RULE: this is the ONLY place the export's call_type_INB column is
# MAGIC read. The B pipeline rebuilds the caller flag natively; the CSV flag is
# MAGIC reconciliation evidence here, nothing else.

# COMMAND ----------

# ============================ A1. constants ============================
CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
DB = f"{CATALOG}.{SCHEMA}"
CSV_PATH = "/Volumes/cda_model_shared/ecm_cld_model/ecm_cld/collections_zenon/WATERFALL_COLL_CALL_V2_202501.csv"
CALL_TABLE = "062108867742_glue_connectivity_catalog.contactcenter_bdp_db.`call`"
ANCHOR_YM = "202501"

NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"   # THE numeric key rule

# The pinned CSV schema (A5 census of 2026-07-16, 90 columns, exact order).
# All-string by construction. A_recon_lock_checks.py asserts the loaded header
# equals this list exactly (the locked explicit-schema convention).
CSV_COLUMNS = [
    "EXTNL_ACCT_ID", "NEW_ROLL_FLAG", "NO_PRIOR_RECORD_FLAG",
    "DLNQT_CD_M1", "DLNQT_CD_M2", "DLNQT_CD_M3",
    "DLNQT_BKT_M1", "DLNQT_BKT_M2", "DLNQT_BKT_M3",
    "PAYMT_AMT_M1", "PAYMT_AMT_M2", "PAYMT_AMT_M3",
    "CHRGOFF_RVRSL_M1", "CHRGOFF_RVRSL_M2", "CHRGOFF_RVRSL_M3",
    "CHRGOFF_AMT_M1", "CHRGOFF_AMT_M2", "CHRGOFF_AMT_M3",
    "GROSS_LOSS_M1", "GROSS_LOSS_M2", "GROSS_LOSS_M3",
    "PLCY_LOSS_M1", "PLCY_LOSS_M2", "PLCY_LOSS_M3",
    "CR_LMT_M1", "CR_LMT_M2", "CR_LMT_M3",
    "EOP_BAL_M1", "EOP_BAL_M2", "EOP_BAL_M3",
    "CHRGOFF_RSN_M1", "CHRGOFF_RSN_M2", "CHRGOFF_RSN_M3",
    "cpc_M1", "cpc_M2", "cpc_M3",
    "ECL_M1", "ECL_M2", "ECL_M3",
    "ECL_12MO_M1", "ECL_12MO_M2", "ECL_12MO_M3",
    "ECL_LIFTM_M1", "ECL_LIFTM_M2", "ECL_LIFTM_M3",
    "STG_CD_M1", "STG_CD_M2", "STG_CD_M3",
    "WRITE_OFF_M1", "WRITE_OFF_M2", "WRITE_OFF_M3",
    "CO_CURRENT_FLAG", "CO_8M_FLAG", "CO_10M_FLAG", "CO_12M_FLAG",
    "REAGE_EVER_FLAG",
    "GROSS_LOSS_8M_AMT", "GROSS_LOSS_10M_AMT", "GROSS_LOSS_12M_AMT",
    "CHRGOFF_8M_AMT", "CHRGOFF_10M_AMT", "CHRGOFF_12M_AMT",
    "PLCY_LOSS_8M_AMT", "PLCY_LOSS_10M_AMT", "PLCY_LOSS_12M_AMT",
    "imp_M0", "ECL_M0", "ECL_12MO_M0", "ECL_LIFTM_M0", "STG_CD_M0", "WRITE_OFF_M0",
    "imp_M4", "ECL_M4", "ECL_12MO_M4", "ECL_LIFTM_M4", "STG_CD_M4", "WRITE_OFF_M4",
    "CPC_FLAG_NW",
    "hram_flag_refit_M1", "hram_flag_apollo_M1",
    "hram_flag_refit_M2", "hram_flag_apollo_M2",
    "hram_flag_refit_M3", "hram_flag_apollo_M3",
    "chrgoff_amt_lftm", "GROSS_LOSS_AMT_LFTM", "chrgoff_val_flag",
    "call_type_CLBCK", "call_type_INB", "call_type_TRSFR",
]
assert len(CSV_COLUMNS) == 90

print(f"A_recon_lock_202501 (lean): target {DB}")

# COMMAND ----------

# A3. preconditions: the round-10 tables exist
for t in ["uc2_t16_01_populations", "uc2_t16_02_episodes", "uc2_t16_04_outcomes"]:
    assert spark.catalog.tableExists(f"{DB}.{t}"), f"PRECONDITION MISS: {DB}.{t} does not exist"

# COMMAND ----------

# A4. CSV load: all-string schema, FAILFAST, mirror table
csv_df = (spark.read.format("csv")
          .option("header", True)
          .option("inferSchema", False)
          .option("mode", "FAILFAST")
          .load(CSV_PATH))
csv_df.createOrReplaceTempView("_sas_csv")
spark.sql(f"CREATE OR REPLACE TABLE {DB}.uc2_sas_raw_202501 AS SELECT * FROM _sas_csv")

# columns THIS notebook uses must exist (the checks sibling pins the full 90)
_required = ["EXTNL_ACCT_ID", "DLNQT_CD_M1", "CPC_FLAG_NW", "CHRGOFF_RSN_M1",
             "call_type_INB", "PAYMT_AMT_M1", "PAYMT_AMT_M2"]
_missing = [c for c in _required if c.upper() not in {x.upper() for x in csv_df.columns}]
assert not _missing, f"REQUIRED COLUMNS MISSING from the CSV: {_missing} - STOP"

# COMMAND ----------

# A7. the enriched build: waterfall flags as columns (round-10-verified filters)
spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_sas_wf_202501 AS
WITH f AS (
    SELECT r.*,
           try_cast(EXTNL_ACCT_ID AS bigint) AS acct_num,
           {NUM_KEY.format(c="EXTNL_ACCT_ID")} AS acct_key,
           coalesce(try_cast(DLNQT_CD_M1 AS int) = 1, false) AS wf_dq1,
           coalesce(upper(trim(CPC_FLAG_NW)) IN ('OTHER', 'OTHERS', 'COBRAND', 'PLCC'), false) AS wf_cpc,
           (CHRGOFF_RSN_M1 IS NULL OR trim(CHRGOFF_RSN_M1) = ''
            OR upper(trim(CHRGOFF_RSN_M1)) IN ('PLY', 'BLANK')) AS wf_non_co,
           coalesce(upper(trim(call_type_INB)) LIKE '%INB%', false) AS csv_inb
    FROM _sas_csv r
)
SELECT *, (wf_dq1 AND wf_cpc AND wf_non_co) AS in_sas_ledger
FROM f
""")
print(f"built {DB}.uc2_sas_wf_202501")

# COMMAND ----------

# A9. the gap decomposition: persist the flagged set and the recovered 1,942
spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW _sas_flagged AS
    SELECT DISTINCT acct_key, acct_num
    FROM {DB}.uc2_sas_wf_202501
    WHERE in_sas_ledger AND csv_inb
""")
spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW _our_callers AS
    SELECT DISTINCT c.acct_key, try_cast(c.acct_key AS bigint) AS acct_num
    FROM {DB}.uc2_t16_02_episodes c
    JOIN {DB}.uc2_t16_01_populations p ON p.acct_key = c.acct_key
    WHERE c.is_episode_std = 1 AND p.in_ledger_exaa
""")
spark.sql(f"""
    CREATE OR REPLACE TABLE {DB}.uc2_gap1942_202501 AS
    SELECT f.acct_key, f.acct_num
    FROM _sas_flagged f LEFT ANTI JOIN _our_callers o ON o.acct_num = f.acct_num
""")
spark.sql(f"""
    CREATE OR REPLACE TABLE {DB}.uc2_sasflag_202501 AS
    SELECT acct_key, acct_num FROM _sas_flagged
""")
print(f"persisted {DB}.uc2_gap1942_202501 and {DB}.uc2_sasflag_202501")

# COMMAND ----------

# A12. the one-time capture-gate delta table (the ONLY place the two gates meet;
# neither gate is a denominator here). captured_sas per the CQ-7 convention
# confirmed by the A11 sign probe (payment = negative PAYMT_AMT in M1/M2).
spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_capture_delta_202501 AS
WITH ours AS (SELECT acct_num FROM _our_callers),
flagged AS (SELECT acct_num FROM {DB}.uc2_sasflag_202501),
uni AS (
    SELECT coalesce(o.acct_num, f.acct_num) AS acct_num,
           CASE WHEN o.acct_num IS NOT NULL AND f.acct_num IS NOT NULL THEN 'a. both'
                WHEN o.acct_num IS NOT NULL THEN 'b. ours-only (string-keyed)'
                ELSE 'c. flagged-only' END AS construct
    FROM ours o FULL OUTER JOIN flagged f ON f.acct_num = o.acct_num
),
aws AS (
    SELECT try_cast(acct_key AS bigint) AS acct_num, max(any_captured) AS aws_captured
    FROM {DB}.uc2_t16_04_outcomes
    GROUP BY 1
),
sas AS (
    SELECT acct_num,
           (coalesce(try_cast(PAYMT_AMT_M1 AS double), 0) < 0
            OR coalesce(try_cast(PAYMT_AMT_M2 AS double), 0) < 0) AS captured_sas
    FROM {DB}.uc2_sas_wf_202501
)
SELECT u.acct_num, u.construct, a.aws_captured, s.captured_sas
FROM uni u
LEFT JOIN aws a ON a.acct_num = u.acct_num
LEFT JOIN sas s ON s.acct_num = u.acct_num
""")
print(f"built {DB}.uc2_capture_delta_202501")

# COMMAND ----------

print("A_recon_lock_202501 build complete: uc2_sas_raw_202501, uc2_sas_wf_202501, "
      "uc2_gap1942_202501, uc2_sasflag_202501, uc2_capture_delta_202501. "
      "Run A_recon_lock_checks.py once to certify.")
