# Databricks notebook source
# MAGIC %md
# MAGIC # Phase 1b: why does our call scan miss the 1,747 flagged ledger callers?
# MAGIC
# MAGIC Phase 1 (2026-07-16) showed ALL 11,136 SAS-flagged callers are in the ex-AA
# MAGIC ledger. So the 11,136-vs-9,389 gap is 1,747 LEDGER accounts that carry the
# MAGIC SAS INBOUND flag but produce no row in our episode layer. This probe scans
# MAGIC the RAW call table for exactly those accounts, with NO filters beyond a
# MAGIC wide date window, and classifies why each account is invisible to the
# MAGIC episode scan (late effdt load / business-card rows / no January INBOUND
# MAGIC row at all / an in-cap row that the build should have kept = build drift).
# MAGIC
# MAGIC READ-ONLY (temp views only). Run top to bottom.
# MAGIC
# MAGIC ALSO NEEDED (manual, one screenshot): the notebook cell that CREATED
# MAGIC `uc2_t16_02_episodes` (the q3 cell in the tier-16 build notebook). The
# MAGIC kit file has no effdt condition in its scan WHERE; if the Databricks port
# MAGIC added one (for partition pruning), that alone explains the gap. Screenshot
# MAGIC that cell's code along with this notebook's grids.

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
# MAGIC ## C2. Rebuild the two sets and the gap (stop rules: 11,136 / 9,389 / 1,747)

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
    CREATE OR REPLACE TEMP VIEW gap_accts AS
    SELECT f.acct_key
    FROM sas_flagged f
    LEFT ANTI JOIN our_callers o ON o.acct_key = f.acct_key
""")

n_flagged = spark.sql("SELECT count(*) AS n FROM sas_flagged").collect()[0]["n"]
n_ours = spark.sql("SELECT count(*) AS n FROM our_callers").collect()[0]["n"]
n_gap = spark.sql("SELECT count(*) AS n FROM gap_accts").collect()[0]["n"]
chk("SAS-flagged caller set", n_flagged, 11136)
chk("our ledger callers", n_ours, 9389)
chk("the gap (flagged, not ours)", n_gap, 1747)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C3. Do the gap accounts have ANY row in our episode layer? (report, no assert)
# MAGIC
# MAGIC Expected 0 per the round-9 settle logic. If nonzero, the rows' flag combos
# MAGIC print below and the cause is filter logic, not the scan window.

# COMMAND ----------

spark.sql(f"""
    SELECT count(DISTINCT c.acct_key) AS gap_accts_with_any_episode_row,
           count(*) AS rows
    FROM {T('uc2_t16_02_episodes')} c
    JOIN gap_accts g ON g.acct_key = c.acct_key
""").show(truncate=False)

spark.sql(f"""
    SELECT c.is_biz, c.within_effdt_cap, c.is_episode_std,
           count(*) AS rows, count(DISTINCT c.acct_key) AS accounts
    FROM {T('uc2_t16_02_episodes')} c
    JOIN gap_accts g ON g.acct_key = c.acct_key
    GROUP BY 1, 2, 3
    ORDER BY 1, 2, 3
""").show(truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C4. THE raw probe: every call row for the 1,747, no filters
# MAGIC
# MAGIC Wide call-date window (15 Dec 2024 to 15 Feb 2025) to catch boundary
# MAGIC effects; no initiationmethod, producttype, or effdt condition. NOTE: no
# MAGIC effdt bound means no partition pruning on the call table; the semi-join to
# MAGIC 1,747 accounts keeps the result small but the scan reads the table. If the
# MAGIC cluster struggles, add an effdt bound '2024-12-01' <= effdt < '2026-07-01'
# MAGIC and note it on the screenshot.

# COMMAND ----------

spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW gap_raw AS
    SELECT trim(cast(c.acctid AS string)) AS acct_key,
           c.contactid,
           c.`date` AS call_dt,
           c.initiationmethod,
           cast(c.producttype AS string) AS producttype,
           cast(c.effdt AS string) AS effdt
    FROM {CALL_TABLE} c
    JOIN gap_accts g ON g.acct_key = trim(cast(c.acctid AS string))
    WHERE c.`date` >= DATE '2024-12-15' AND c.`date` < DATE '2025-02-15'
""")

spark.sql("""
    CREATE OR REPLACE TEMP VIEW gap_acct_facts AS
    SELECT g.acct_key,
           count(r.contactid) AS raw_rows,
           count_if(r.initiationmethod = 'INBOUND'
                    AND r.call_dt >= DATE '2025-01-01' AND r.call_dt < DATE '2025-02-01')
               AS jan_inbound_rows,
           count_if(r.initiationmethod = 'INBOUND'
                    AND r.call_dt >= DATE '2025-01-01' AND r.call_dt < DATE '2025-02-01'
                    AND r.effdt >= '2025-01-01' AND r.effdt < '2025-02-02'
                    AND coalesce(r.producttype, '') <> 'BUSINESS_CARD')
               AS jan_inbound_incap_consumer_rows,
           count_if(r.initiationmethod = 'INBOUND'
                    AND r.call_dt >= DATE '2025-01-01' AND r.call_dt < DATE '2025-02-01'
                    AND r.effdt >= '2025-01-01' AND r.effdt < '2025-02-02'
                    AND coalesce(r.producttype, '') = 'BUSINESS_CARD')
               AS jan_inbound_incap_biz_rows,
           count_if(r.initiationmethod = 'INBOUND'
                    AND r.call_dt >= DATE '2025-01-01' AND r.call_dt < DATE '2025-02-01'
                    AND r.effdt >= '2025-02-02')
               AS jan_inbound_late_effdt_rows,
           count_if(r.initiationmethod = 'INBOUND'
                    AND r.call_dt >= DATE '2025-01-01' AND r.call_dt < DATE '2025-02-01'
                    AND r.effdt < '2025-01-01')
               AS jan_inbound_early_effdt_rows,
           count_if(r.initiationmethod <> 'INBOUND') AS non_inbound_rows,
           count_if(r.call_dt < DATE '2025-01-01' OR r.call_dt >= DATE '2025-02-01')
               AS outside_jan_rows
    FROM gap_accts g
    LEFT JOIN gap_raw r ON r.acct_key = g.acct_key
    GROUP BY g.acct_key
""")

probe = spark.sql("""
    SELECT CASE
        WHEN jan_inbound_incap_consumer_rows > 0
            THEN '1. in-cap consumer January INBOUND row EXISTS (episode-build drift)'
        WHEN jan_inbound_incap_biz_rows > 0
            THEN '2. in-cap January INBOUND rows, BUSINESS_CARD only'
        WHEN jan_inbound_late_effdt_rows > 0
            THEN '3. January INBOUND rows, all loaded late (effdt on/after 2025-02-02)'
        WHEN jan_inbound_early_effdt_rows > 0
            THEN '4. January INBOUND rows, effdt before 2025-01-01'
        WHEN jan_inbound_rows = 0 AND non_inbound_rows > 0
            THEN '5. rows in window, none January INBOUND'
        WHEN raw_rows = 0
            THEN '6. NO raw call rows at all (export/flag-side question)'
        ELSE '7. UNCLASSIFIED'
    END AS probe_class,
    count(*) AS accounts,
    round(100.0 * count(*) / sum(count(*)) OVER (), 1) AS share_pct
    FROM gap_acct_facts
    GROUP BY 1
    ORDER BY 1
""")
probe.show(truncate=False)

covered = sum(r["accounts"] for r in probe.collect())
chk("probe classes cover the gap", covered, 1747)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C5. Color: what the missing rows look like

# COMMAND ----------

# effdt months of the gap accounts' January INBOUND rows (when were they loaded?)
spark.sql("""
    SELECT substr(effdt, 1, 7) AS effdt_month, count(*) AS rows,
           count(DISTINCT acct_key) AS accounts
    FROM gap_raw
    WHERE initiationmethod = 'INBOUND'
      AND call_dt >= DATE '2025-01-01' AND call_dt < DATE '2025-02-01'
    GROUP BY 1 ORDER BY 1
""").show(50, truncate=False)

# initiationmethod x producttype over all their rows in the window
spark.sql("""
    SELECT initiationmethod, coalesce(producttype, '(null)') AS producttype,
           count(*) AS rows, count(DISTINCT acct_key) AS accounts
    FROM gap_raw
    GROUP BY 1, 2 ORDER BY rows DESC
""").show(50, truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C6. Verdict

# COMMAND ----------

print("=" * 78)
print("PHASE 1b VERDICT: why the 1,747 are invisible to the episode scan")
print("=" * 78)
for name, actual, expected, status in RESULTS:
    print(f"{status:4}  {name:45} {actual:>12,}  (expected {expected:,})")
print("-" * 78)
probe.show(truncate=False)
print("Reading guide: class 3 = the effdt load-date cap named on 2026-07-13;")
print("class 2 = the business-card exclusion; class 1 = the Databricks episode")
print("build dropped rows the kit keeps (screenshot the q3 build cell!);")
print("classes 5/6 = the export's call_month/M1 definition differs from the")
print("calendar-January call-date window, or the flag itself. Screenshot C3,")
print("C4, C5 grids and this block; results land in bridge round 11.")
print("=" * 78)
