# Databricks notebook source
# MAGIC %md
# MAGIC # A_recon_lock_checks - run ONCE after A_recon_lock_202501.py to certify.
# MAGIC
# MAGIC Re-reads the tables A built (uc2_sas_raw_202501, uc2_sas_wf_202501,
# MAGIC uc2_gap1942_202501, uc2_sasflag_202501, uc2_capture_delta_202501, the CSV,
# MAGIC the call table, the round-10 tables) and asserts every locked round-12
# MAGIC value. A miss STOPS. It rebuilds NO logic; it only re-measures and checks.

# COMMAND ----------

# paste _checks_common.py here, or import it as a notebook. It defines chk()/fmt().
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
DB = f"{CATALOG}.{SCHEMA}"
CSV_PATH = "/Volumes/cda_model_shared/ecm_cld_model/ecm_cld/collections_zenon/WATERFALL_COLL_CALL_V2_202501.csv"
CALL_TABLE = "062108867742_glue_connectivity_catalog.contactcenter_bdp_db.`call`"
NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"

# locked EXPECTED for A (round-12 record; the A5 census + delta grid included)
EXPECTED = {
    "r10 ledger all": 204323,
    "r10 ledger exaa": 189146,
    "r10 touched b1": 724848,
    "r10 ledger callers (string key)": 9389,
    "r10 ledger episodes (string key)": 11262,
    "csv rows": 610183,
    "csv distinct accounts": 610183,
    "csv id null": 0,
    "csv id non-castable": 0,
    "csv id pad-mismatch": 0,
    "csv column count": 90,
    "wf 01 total": 610183,
    "wf 02 dq1": 202479,
    "wf 03 +cpc": 186848,
    "wf 04 sas ledger": 186013,
    "wf inb 01 total": 34234,
    "wf inb 02 dq1": 12615,
    "wf inb 03 +cpc": 11289,
    "wf inb 04 flagged": 11136,
    "gap: sas flagged": 11136,
    "gap: our callers": 9389,
    "gap: shared": 9194,
    "gap: flagged only": 1942,
    "gap: ours only": 195,
    "1d: numeric coverage of the 1,942": 1942,
    "1d: jan inbound rows recovered": 2220,
    "1d: jan inbound accounts recovered": 1942,
    "1d: jan key mismatches (id-carrying inbound)": 75883,
    "1d: jan inbound rows null acctid": 312795,
    "1d: jan acctid digits-only rows": 1267227,
    "1d: jan acctid null rows": 481838,
    "1d: jan acctid other-shape rows": 0,
    "probe: aws-captured accounts joined to export": 6029,
    "probe: of those, negative PAYMT M1 or M2": 5767,
    "delta grid": {
        "a. both|0|False": 2148,
        "a. both|0|True": 1112,
        "a. both|1|False": 259,
        "a. both|1|True": 5675,
        "b. ours-only (string-keyed)|0|False": 78,
        "b. ours-only (string-keyed)|0|True": 22,
        "b. ours-only (string-keyed)|1|False": 3,
        "b. ours-only (string-keyed)|1|True": 92,
        "c. flagged-only|0|False": 692,
        "c. flagged-only|0|True": 1250,
    },
}

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

# COMMAND ----------

# A3. round-10 preconditions
_r = spark.sql(f"""
    SELECT count_if(in_ledger_all)  AS ledger_all,
           count_if(in_ledger_exaa) AS ledger_exaa,
           count_if(touched_b1)     AS touched_b1
    FROM {DB}.uc2_t16_01_populations
""").first()
chk("r10 ledger all", _r["ledger_all"], EXPECTED["r10 ledger all"])
chk("r10 ledger exaa", _r["ledger_exaa"], EXPECTED["r10 ledger exaa"])
chk("r10 touched b1", _r["touched_b1"], EXPECTED["r10 touched b1"])

_r = spark.sql(f"""
    SELECT count(*) AS episodes, count(DISTINCT c.acct_key) AS callers
    FROM {DB}.uc2_t16_02_episodes c
    JOIN {DB}.uc2_t16_01_populations p ON p.acct_key = c.acct_key
    WHERE c.is_episode_std = 1 AND p.in_ledger_exaa
""").first()
chk("r10 ledger episodes (string key)", _r["episodes"], EXPECTED["r10 ledger episodes (string key)"])
chk("r10 ledger callers (string key)", _r["callers"], EXPECTED["r10 ledger callers (string key)"])

# COMMAND ----------

# A4/A5. CSV grain + the pinned 90-column header
csv_df = (spark.read.format("csv").option("header", True)
          .option("inferSchema", False).option("mode", "FAILFAST").load(CSV_PATH))
csv_df.createOrReplaceTempView("_sas_csv")
_r = spark.sql("SELECT count(*) AS rows, count(DISTINCT EXTNL_ACCT_ID) AS accts FROM _sas_csv").first()
chk("csv rows", _r["rows"], EXPECTED["csv rows"])
chk("csv distinct accounts", _r["accts"], EXPECTED["csv distinct accounts"])

_cols = csv_df.columns
chk("csv column count", len(_cols), EXPECTED["csv column count"])
assert _cols == CSV_COLUMNS, (
    "CSV SCHEMA DRIFT: the loaded header differs from the pinned A5 census - STOP. "
    f"first difference at position {next(i for i, (a, b) in enumerate(zip(_cols + ['<end>'], CSV_COLUMNS + ['<end>'])) if a != b) + 1}")
print("PASS  csv header equals the pinned 90-column census exactly")

# COMMAND ----------

# A6. export key probe (no null / non-castable / padded ids)
_r = spark.sql(f"""
    SELECT count_if(EXTNL_ACCT_ID IS NULL) AS id_null,
           count_if(EXTNL_ACCT_ID IS NOT NULL AND try_cast(EXTNL_ACCT_ID AS bigint) IS NULL) AS non_castable,
           count_if(EXTNL_ACCT_ID IS NOT NULL AND try_cast(EXTNL_ACCT_ID AS bigint) IS NOT NULL
                    AND trim(cast(EXTNL_ACCT_ID AS string)) <> {NUM_KEY.format(c="EXTNL_ACCT_ID")}) AS pad_mismatch
    FROM _sas_csv
""").first()
chk("csv id null", _r["id_null"], EXPECTED["csv id null"])
chk("csv id non-castable", _r["non_castable"], EXPECTED["csv id non-castable"])
chk("csv id pad-mismatch", _r["pad_mismatch"], EXPECTED["csv id pad-mismatch"])

# COMMAND ----------

# A8. the eight waterfall ladder numbers (population + INBOUND)
_r = spark.sql(f"""
    SELECT count(*)                                AS s1,
           count_if(wf_dq1)                        AS s2,
           count_if(wf_dq1 AND wf_cpc)             AS s3,
           count_if(in_sas_ledger)                 AS s4,
           count_if(csv_inb)                       AS i1,
           count_if(csv_inb AND wf_dq1)            AS i2,
           count_if(csv_inb AND wf_dq1 AND wf_cpc) AS i3,
           count_if(csv_inb AND in_sas_ledger)     AS i4
    FROM {DB}.uc2_sas_wf_202501
""").first()
chk("wf 01 total", _r["s1"], EXPECTED["wf 01 total"])
chk("wf 02 dq1", _r["s2"], EXPECTED["wf 02 dq1"])
chk("wf 03 +cpc", _r["s3"], EXPECTED["wf 03 +cpc"])
chk("wf 04 sas ledger", _r["s4"], EXPECTED["wf 04 sas ledger"])
chk("wf inb 01 total", _r["i1"], EXPECTED["wf inb 01 total"])
chk("wf inb 02 dq1", _r["i2"], EXPECTED["wf inb 02 dq1"])
chk("wf inb 03 +cpc", _r["i3"], EXPECTED["wf inb 03 +cpc"])
chk("wf inb 04 flagged", _r["i4"], EXPECTED["wf inb 04 flagged"])

# COMMAND ----------

# A9. gap decomposition (re-derived from the persisted tables + round-10)
spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW _our_callers AS
    SELECT DISTINCT c.acct_key, try_cast(c.acct_key AS bigint) AS acct_num
    FROM {DB}.uc2_t16_02_episodes c
    JOIN {DB}.uc2_t16_01_populations p ON p.acct_key = c.acct_key
    WHERE c.is_episode_std = 1 AND p.in_ledger_exaa
""")
_r = spark.sql(f"""
    SELECT (SELECT count(*) FROM {DB}.uc2_sasflag_202501)  AS flagged,
           (SELECT count(*) FROM _our_callers)             AS ours,
           (SELECT count(*) FROM {DB}.uc2_sasflag_202501 f JOIN _our_callers o ON o.acct_num = f.acct_num)      AS shared,
           (SELECT count(*) FROM {DB}.uc2_sasflag_202501 f LEFT ANTI JOIN _our_callers o ON o.acct_num = f.acct_num) AS flagged_only,
           (SELECT count(*) FROM _our_callers o LEFT ANTI JOIN {DB}.uc2_sasflag_202501 f ON f.acct_num = o.acct_num) AS ours_only
""").first()
chk("gap: sas flagged", _r["flagged"], EXPECTED["gap: sas flagged"])
chk("gap: our callers", _r["ours"], EXPECTED["gap: our callers"])
chk("gap: shared", _r["shared"], EXPECTED["gap: shared"])
chk("gap: flagged only", _r["flagged_only"], EXPECTED["gap: flagged only"])
chk("gap: ours only", _r["ours_only"], EXPECTED["gap: ours only"])

# COMMAND ----------

# A10. numeric-rejoin evidence (immutable under the effdt bound)
spark.sql(f"REFRESH TABLE {CALL_TABLE}")
_r = spark.sql(f"""
    SELECT count(DISTINCT g.acct_key) AS coverage
    FROM {CALL_TABLE} c
    JOIN {DB}.uc2_gap1942_202501 g ON g.acct_num = try_cast(c.acctid AS bigint)
    WHERE c.effdt >= '2024-12-01' AND c.effdt < '2026-07-10'
""").first()
chk("1d: numeric coverage of the 1,942", _r["coverage"], EXPECTED["1d: numeric coverage of the 1,942"])

_r = spark.sql(f"""
    SELECT count(*) AS jan_rows, count(DISTINCT g.acct_key) AS jan_accts
    FROM {CALL_TABLE} c
    JOIN {DB}.uc2_gap1942_202501 g ON g.acct_num = try_cast(c.acctid AS bigint)
    WHERE c.effdt >= '2024-12-01' AND c.effdt < '2026-07-10'
      AND c.initiationmethod = 'INBOUND'
      AND c.`date` >= DATE '2025-01-01' AND c.`date` < DATE '2025-02-01'
""").first()
chk("1d: jan inbound rows recovered", _r["jan_rows"], EXPECTED["1d: jan inbound rows recovered"])
chk("1d: jan inbound accounts recovered", _r["jan_accts"], EXPECTED["1d: jan inbound accounts recovered"])

_r = spark.sql(f"""
    SELECT count_if(initiationmethod = 'INBOUND' AND acctid IS NOT NULL
                    AND (try_cast(acctid AS bigint) IS NULL
                         OR trim(cast(acctid AS string)) <> {NUM_KEY.format(c="acctid")})) AS mismatches,
           count_if(initiationmethod = 'INBOUND' AND acctid IS NULL) AS inb_null_id,
           count_if(acctid IS NULL) AS null_rows,
           count_if(acctid IS NOT NULL AND cast(acctid AS string) rlike '^[0-9]+$') AS digits_rows,
           count_if(acctid IS NOT NULL AND NOT cast(acctid AS string) rlike '^[0-9]+$') AS other_rows
    FROM {CALL_TABLE}
    WHERE `date` >= DATE '2025-01-01' AND `date` < DATE '2025-02-01'
      AND effdt >= '2024-12-01' AND effdt < '2026-07-10'
""").first()
chk("1d: jan key mismatches (id-carrying inbound)", _r["mismatches"], EXPECTED["1d: jan key mismatches (id-carrying inbound)"])
chk("1d: jan inbound rows null acctid", _r["inb_null_id"], EXPECTED["1d: jan inbound rows null acctid"])
chk("1d: jan acctid null rows", _r["null_rows"], EXPECTED["1d: jan acctid null rows"])
chk("1d: jan acctid digits-only rows", _r["digits_rows"], EXPECTED["1d: jan acctid digits-only rows"])
chk("1d: jan acctid other-shape rows", _r["other_rows"], EXPECTED["1d: jan acctid other-shape rows"])

# COMMAND ----------

# A11. PAYMT_AMT sign probe (CQ-7 confirmation on the known-payer set)
_r = spark.sql(f"""
    WITH cap AS (
        SELECT try_cast(acct_key AS bigint) AS acct_num
        FROM {DB}.uc2_t16_04_outcomes
        WHERE any_captured = 1 AND in_ledger_exaa
        GROUP BY 1
    )
    SELECT count(*) AS joined,
           count_if(coalesce(try_cast(w.PAYMT_AMT_M1 AS double), 0) < 0
                    OR coalesce(try_cast(w.PAYMT_AMT_M2 AS double), 0) < 0) AS with_negative
    FROM cap c
    JOIN {DB}.uc2_sas_wf_202501 w ON w.acct_num = c.acct_num
""").first()
chk("probe: aws-captured accounts joined to export", _r["joined"],
    EXPECTED["probe: aws-captured accounts joined to export"])
chk("probe: of those, negative PAYMT M1 or M2", _r["with_negative"],
    EXPECTED["probe: of those, negative PAYMT M1 or M2"])

# COMMAND ----------

# A12. the capture-gate delta grid (construct x aws day-grain x sas month-grain)
_delta = spark.sql(f"""
    SELECT construct, aws_captured, captured_sas, count(*) AS accounts
    FROM {DB}.uc2_capture_delta_202501
    GROUP BY 1, 2, 3
""").collect()
_measured = {f"{r['construct']}|{fmt(r['aws_captured'])}|{fmt(r['captured_sas'])}": r["accounts"]
             for r in _delta}
assert _measured == EXPECTED["delta grid"], \
    f"ANCHOR MISS delta grid: measured {_measured} vs expected {EXPECTED['delta grid']}"
print("PASS  delta grid matches the frozen expectation")

print("A_recon_lock_checks: ALL PASS - the lean A build is certified equivalent to the locked original.")
