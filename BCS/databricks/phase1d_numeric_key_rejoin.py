# Databricks notebook source
# MAGIC %md
# MAGIC # Phase 1d: is the missing-calls mystery a JOIN-KEY FORMAT artifact?
# MAGIC
# MAGIC Phase 1c (2026-07-16): of the 1,942 flagged-only accounts, 930 have NO
# MAGIC id-attached call rows ever under our string join, and none of the 1,942
# MAGIC has a calendar-January-2025 call. But the SAS side attached these accounts
# MAGIC to calls with a NUMERIC join (input(extnl_acct_id, BEST12.)), which is
# MAGIC immune to string-format differences (leading zeros, '123.0', scientific
# MAGIC notation from a double-typed column). If call.acctid sometimes casts to a
# MAGIC string our trim(cast(...)) key cannot match, those calls are invisible to
# MAGIC the ENTIRE tier-16 episode layer - and our 9,389 undercounts the same way.
# MAGIC
# MAGIC This probe re-runs the coverage and month histogram with a NUMERIC join,
# MAGIC and measures the format mismatch directly on January's rows.
# MAGIC
# MAGIC READ-ONLY (temp views only).
# MAGIC
# MAGIC STILL OUTSTANDING and now critical: the export-code cell that DEFINES
# MAGIC call_month and shows which source table and id column the
# MAGIC aws_call_accts_jan_mar_25 build read. The data can narrow this; only the
# MAGIC code closes it.

# COMMAND ----------

CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
CALL_TABLE = "062108867742_glue_connectivity_catalog.contactcenter_bdp_db.call"

RESULTS = []


def chk(name, actual, expected):
    if actual != expected:
        raise AssertionError(f"ANCHOR MISS {name}: got {actual:,}, expected {expected:,}")
    RESULTS.append((name, actual, expected, "PASS"))
    print(f"PASS  {name}: {actual:,}")


def T(name):
    return f"{CATALOG}.{SCHEMA}.{name}"


# COMMAND ----------

# MAGIC %md
# MAGIC ## C1. What type IS acctid here? (schema + value-shape census, January window)

# COMMAND ----------

spark.sql(f"DESCRIBE {CALL_TABLE}").show(50, truncate=False)

spark.sql(f"REFRESH TABLE {CALL_TABLE}")

# Value shapes of acctid on January-2025 rows: digits-only vs anything else,
# and the NULL share (the known ~28% id hole).
spark.sql(f"""
    SELECT CASE
             WHEN acctid IS NULL THEN 'a. NULL'
             WHEN cast(acctid AS string) rlike '^[0-9]+$' THEN 'b. digits-only string'
             WHEN cast(acctid AS string) rlike '^[0-9]+\\\\.0+$' THEN 'c. trailing .0 decimal'
             WHEN upper(cast(acctid AS string)) rlike 'E' THEN 'd. scientific notation'
             ELSE 'e. other shape'
           END AS acctid_shape,
           count(*) AS rows
    FROM {CALL_TABLE}
    WHERE `date` >= DATE '2025-01-01' AND `date` < DATE '2025-02-01'
      AND effdt >= '2024-12-01' AND effdt < '2026-07-10'
    GROUP BY 1 ORDER BY 1
""").show(truncate=False)

# Direct mismatch measure: January INBOUND rows where the string key and the
# numeric key disagree (nonzero = the episode layer loses these rows).
spark.sql(f"""
    SELECT count(*) AS jan_inbound_rows_string_vs_numeric_key_mismatch
    FROM {CALL_TABLE}
    WHERE `date` >= DATE '2025-01-01' AND `date` < DATE '2025-02-01'
      AND effdt >= '2024-12-01' AND effdt < '2026-07-10'
      AND initiationmethod = 'INBOUND'
      AND acctid IS NOT NULL
      AND (try_cast(acctid AS bigint) IS NULL
           OR trim(cast(acctid AS string)) <> cast(try_cast(acctid AS bigint) AS string))
""").show(truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C2. The sets (stop rules: 11,136 / 9,389 / 1,942), keys in BOTH forms

# COMMAND ----------

spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW sas_flagged AS
    SELECT DISTINCT trim(cast(EXTNL_ACCT_ID AS string)) AS acct_key,
           try_cast(EXTNL_ACCT_ID AS bigint) AS acct_long
    FROM {T('waterfall_coll_call_enriched')}
    WHERE cast(DLNQT_CD_M1 AS int) = 1
      AND upper(trim(CPC_FLAG_NW)) IN ('OTHER', 'OTHERS', 'COBRAND', 'PLCC')
      AND upper(trim(call_type_INB)) LIKE '%INB%'
      AND (CHRGOFF_RSN_M1 IS NULL
           OR trim(CHRGOFF_RSN_M1) = ''
           OR upper(trim(CHRGOFF_RSN_M1)) IN ('PLY', 'BLANK'))
""")

spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW our_callers AS
    SELECT DISTINCT c.acct_key
    FROM {T('uc2_t16_02_episodes')} c
    JOIN {T('uc2_t16_01_populations')} p ON p.acct_key = c.acct_key
    WHERE c.is_episode_std = 1
      AND p.in_ledger_exaa
""")

spark.sql("""
    CREATE OR REPLACE TEMP VIEW flagged_only AS
    SELECT f.acct_key, f.acct_long
    FROM sas_flagged f
    LEFT ANTI JOIN our_callers o ON o.acct_key = f.acct_key
""")

chk("SAS-flagged caller set",
    spark.sql("SELECT count(*) n FROM sas_flagged").collect()[0]["n"], 11136)
chk("our ledger callers",
    spark.sql("SELECT count(*) n FROM our_callers").collect()[0]["n"], 9389)
chk("flagged only",
    spark.sql("SELECT count(*) n FROM flagged_only").collect()[0]["n"], 1942)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C3. The NUMERIC rejoin: coverage + month histogram for the 1,942
# MAGIC
# MAGIC Same scan as 1c but joined on try_cast(acctid AS bigint). Compare against
# MAGIC the string-join baseline (1,012 accounts with any row; zero January
# MAGIC INBOUND). Any gain here = rows the string key loses.

# COMMAND ----------

spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW gap_all_num AS
    SELECT g.acct_key,
           date_format(c.`date`, 'yyyy-MM') AS call_ym,
           c.`date` AS call_dt,
           c.initiationmethod
    FROM {CALL_TABLE} c
    JOIN flagged_only g ON g.acct_long = try_cast(c.acctid AS bigint)
    WHERE c.effdt >= '2024-12-01' AND c.effdt < '2026-07-10'
""")

print("NUMERIC-join coverage (string baseline was 1,012):")
spark.sql("SELECT count(DISTINCT acct_key) AS accts_with_any_row_numeric FROM gap_all_num").show(truncate=False)

print("NUMERIC-join January-2025 INBOUND (string baseline was ZERO):")
spark.sql("""
    SELECT count(*) AS jan_inbound_rows,
           count(DISTINCT acct_key) AS jan_inbound_accounts
    FROM gap_all_num
    WHERE initiationmethod = 'INBOUND'
      AND call_dt >= DATE '2025-01-01' AND call_dt < DATE '2025-02-01'
""").show(truncate=False)

print("NUMERIC-join month histogram (compare 1c):")
spark.sql("""
    SELECT call_ym, initiationmethod, count(*) AS rows,
           count(DISTINCT acct_key) AS accounts
    FROM gap_all_num
    GROUP BY 1, 2 ORDER BY 1, 2
""").show(100, truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C4. Context volume: January INBOUND rows invisible to BOTH sides
# MAGIC
# MAGIC NULL-acctid January INBOUND rows (the known ~28% id hole, for scale next
# MAGIC to whatever C1/C3 found).

# COMMAND ----------

spark.sql(f"""
    SELECT count(*) AS jan_inbound_rows_null_acctid
    FROM {CALL_TABLE}
    WHERE `date` >= DATE '2025-01-01' AND `date` < DATE '2025-02-01'
      AND effdt >= '2024-12-01' AND effdt < '2026-07-10'
      AND initiationmethod = 'INBOUND'
      AND acctid IS NULL
""").show(truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C5. Verdict

# COMMAND ----------

print("=" * 78)
print("PHASE 1d VERDICT: numeric-key rejoin")
print("=" * 78)
for name, actual, expected, status in RESULTS:
    print(f"{status:4}  {name:45} {actual:>12,}  (expected {expected:,})")
print("-" * 78)
print("Reading guide:")
print("- If C3's numeric join finds January INBOUND rows the string join could")
print("  not, the gap is a KEY-FORMAT artifact: the episode layer (and the")
print("  9,389) loses rows whose acctid does not cast cleanly to our string")
print("  key. The C1 mismatch count sizes the effect table-wide.")
print("- If the numeric join finds nothing new, the flagged calls are simply")
print("  not in this table today under any account key: the export's source,")
print("  id column, or call_month definition differs. Only the export code")
print("  settles that; screenshot the cell that defines call_month.")
print("=" * 78)
