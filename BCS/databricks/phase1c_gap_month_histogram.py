# Databricks notebook source
# MAGIC %md
# MAGIC # Phase 1c: WHEN did the 1,942 flagged-not-ours accounts actually call?
# MAGIC
# MAGIC Phase 1b (2026-07-16) measured: flagged 11,136 / ours 9,389 / overlap 9,194
# MAGIC / flagged-only 1,942 / ours-only 195. NONE of the 1,942 has a
# MAGIC January-dated INBOUND row in the call table (85% have no rows at all in a
# MAGIC 15-Dec..15-Feb window). So the SAS flag's call_month='M1' cannot be
# MAGIC calendar January by call date for these accounts. This probe drops ALL
# MAGIC date filters and histograms their calls by calendar month, and
# MAGIC characterizes the 195 ours-only accounts (January callers the export did
# MAGIC not flag). Together the two grids identify what 'M1' really is.
# MAGIC
# MAGIC READ-ONLY (temp views only).
# MAGIC
# MAGIC ALSO NEEDED (one screenshot, decisive): the FULL Athena export code that
# MAGIC built zenon.aws_call_accts_jan_mar_25 - specifically the expression that
# MAGIC DEFINES call_month (only the filter line `call_month = 'M1'` is on record).
# MAGIC If call_month is statement-anchored, everything observed follows.
# MAGIC
# MAGIC Caveat stated up front: rows whose acctid is NULL today are invisible to
# MAGIC every account-keyed scan here, ours and the export's alike.

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
# MAGIC ## C2. The three sets (stop rules now the MEASURED 1b values)

# COMMAND ----------

spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW sas_flagged AS
    SELECT DISTINCT trim(cast(EXTNL_ACCT_ID AS string)) AS acct_key
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
    SELECT f.acct_key FROM sas_flagged f
    LEFT ANTI JOIN our_callers o ON o.acct_key = f.acct_key
""")
spark.sql("""
    CREATE OR REPLACE TEMP VIEW ours_only AS
    SELECT o.acct_key FROM our_callers o
    LEFT ANTI JOIN sas_flagged f ON f.acct_key = o.acct_key
""")

chk("SAS-flagged caller set",
    spark.sql("SELECT count(*) n FROM sas_flagged").collect()[0]["n"], 11136)
chk("our ledger callers",
    spark.sql("SELECT count(*) n FROM our_callers").collect()[0]["n"], 9389)
chk("flagged only (1b measured)",
    spark.sql("SELECT count(*) n FROM flagged_only").collect()[0]["n"], 1942)
chk("ours only (1b measured)",
    spark.sql("SELECT count(*) n FROM ours_only").collect()[0]["n"], 195)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C3. THE histogram: every call row for the 1,942, ALL TIME, by calendar month
# MAGIC
# MAGIC No date, method, producttype, or effdt filter. If the flag is real, their
# MAGIC INBOUND calls are SOMEWHERE; this shows where. (Cost note: full call-table
# MAGIC scan with a broadcast semi-join; the transcript table is the expensive one,
# MAGIC not this. If it still struggles, bound `date` to 2024-01-01..2026-01-01.)

# COMMAND ----------

spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW gap_all AS
    SELECT trim(cast(c.acctid AS string)) AS acct_key,
           c.`date` AS call_dt,
           date_format(c.`date`, 'yyyy-MM') AS call_ym,
           c.initiationmethod,
           cast(c.producttype AS string) AS producttype
    FROM {CALL_TABLE} c
    JOIN flagged_only g ON g.acct_key = trim(cast(c.acctid AS string))
""")

print("All rows by calendar month x method:")
spark.sql("""
    SELECT call_ym, initiationmethod, count(*) AS rows,
           count(DISTINCT acct_key) AS accounts
    FROM gap_all
    GROUP BY 1, 2
    ORDER BY 1, 2
""").show(100, truncate=False)

print("Coverage: how many of the 1,942 have ANY row, ever:")
spark.sql("""
    SELECT count(DISTINCT acct_key) AS accts_with_any_row_ever
    FROM gap_all
""").show(truncate=False)

print("Per-account first/last INBOUND call date distribution (month of first):")
spark.sql("""
    SELECT date_format(min_dt, 'yyyy-MM') AS first_inbound_month,
           count(*) AS accounts
    FROM (
        SELECT acct_key, min(call_dt) AS min_dt
        FROM gap_all
        WHERE initiationmethod = 'INBOUND'
        GROUP BY acct_key
    )
    GROUP BY 1 ORDER BY 1
""").show(50, truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C4. The 195 ours-only: January callers the export did NOT flag
# MAGIC
# MAGIC Their January episode dates, by day-of-month band. If the export's M1 is
# MAGIC statement-anchored, these should cluster where a December-statement window
# MAGIC does not reach (the mirror image of C3).

# COMMAND ----------

spark.sql(f"""
    SELECT CASE
             WHEN day(c.call_dt) <= 10 THEN 'a. Jan 1-10'
             WHEN day(c.call_dt) <= 20 THEN 'b. Jan 11-20'
             ELSE 'c. Jan 21-31'
           END AS call_day_band,
           count(*) AS episodes,
           count(DISTINCT c.acct_key) AS accounts
    FROM {T('uc2_t16_02_episodes')} c
    JOIN ours_only o ON o.acct_key = c.acct_key
    WHERE c.is_episode_std = 1
    GROUP BY 1 ORDER BY 1
""").show(truncate=False)

# And what the export CSV knows about these 195 (flag columns as loaded):
spark.sql(f"""
    SELECT coalesce(upper(trim(call_type_INB)), '(null)') AS call_type_INB,
           coalesce(upper(trim(call_type_TRSFR)), '(null)') AS call_type_TRSFR,
           count(*) AS accounts
    FROM {T('waterfall_coll_call_enriched')} w
    JOIN ours_only o ON o.acct_key = trim(cast(w.EXTNL_ACCT_ID AS string))
    GROUP BY 1, 2 ORDER BY accounts DESC
""").show(truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C5. Verdict

# COMMAND ----------

print("=" * 78)
print("PHASE 1c VERDICT: locating the flagged-only accounts' calls in time")
print("=" * 78)
for name, actual, expected, status in RESULTS:
    print(f"{status:4}  {name:45} {actual:>12,}  (expected {expected:,})")
print("-" * 78)
print("Reading guide: if the C3 histogram shows the 1,942's INBOUND calls in")
print("Feb/Mar (or Dec), the export's call_month='M1' is NOT calendar January")
print("and the SAS flag mixes non-January callers into its 11,136; our 9,389")
print("remains the calendar-January count. If C3 shows NO rows ever for most")
print("accounts, the flag's account attachment itself is the question (acctid")
print("backfill or a different id source at export time). Screenshot all grids")
print("plus the export-code cell that DEFINES call_month.")
print("=" * 78)
