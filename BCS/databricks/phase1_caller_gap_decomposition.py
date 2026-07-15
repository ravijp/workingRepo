# Databricks notebook source
# MAGIC %md
# MAGIC # Phase 1: caller-gap decomposition (the real cause of 11,136 vs 9,389)
# MAGIC
# MAGIC Purpose: classify EVERY SAS-flagged inbound caller (the 11,136 from the
# MAGIC 003-export replication) by why it is or is not in the AWS ex-AA ledger
# MAGIC (189,146). The class counts ARE the answer to "what causes the gap".
# MAGIC
# MAGIC READ-ONLY: this notebook writes no tables (temp views only). Safe to run
# MAGIC on any cluster without touching the main notebook. Run top to bottom;
# MAGIC any anchor miss RAISES and stops the run.
# MAGIC
# MAGIC Inputs (must already exist, built by the earlier runs):
# MAGIC 1. `waterfall_coll_call_enriched`  (the SAS 003-export CSV, 610,183 accounts)
# MAGIC 2. `uc2_t16_00_acct_monthly`       (tier-16 layer 00)
# MAGIC 3. `uc2_t16_01_populations`        (tier-16 layer 01, the round-10 build)
# MAGIC
# MAGIC Trap reminder: the caller constructs are never merged. This notebook
# MAGIC decomposes ONE construct (the SAS INBOUND flag population, 11,136) against
# MAGIC the ledger construct (9,389). The call-day stream (29,114), the touched-B1
# MAGIC callers, and the statement-window callers (19,789) play no part here.
# MAGIC
# MAGIC Expected stop rules (pre-registered from bridge rounds 9-10):
# MAGIC   flagged set = 11,136; classes sum = 11,136; class 6 (in ledger) = 9,389;
# MAGIC   class 7 (unclassified) = 0; classes 1-5 sum = 1,747.

# COMMAND ----------

CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"

RESULTS = []  # (name, actual, expected, "PASS")


def chk(name, actual, expected):
    """Hard-failing anchor check. A miss STOPS the run."""
    if actual != expected:
        raise AssertionError(f"ANCHOR MISS {name}: got {actual:,}, expected {expected:,}")
    RESULTS.append((name, actual, expected, "PASS"))
    print(f"PASS  {name}: {actual:,}")


def T(name):
    return f"{CATALOG}.{SCHEMA}.{name}"


print(f"Target: {CATALOG}.{SCHEMA}")

# COMMAND ----------

# MAGIC %md
# MAGIC ## C2. Preconditions: the three input tables exist and are the verified builds

# COMMAND ----------

for t in ["waterfall_coll_call_enriched", "uc2_t16_00_acct_monthly", "uc2_t16_01_populations"]:
    if not spark.catalog.tableExists(T(t)):
        raise AssertionError(f"PRECONDITION MISS: table {T(t)} does not exist")
    print(f"PASS  table exists: {T(t)}")

wf_accts = spark.sql(
    f"SELECT count(DISTINCT extnl_acct_id) AS n FROM {T('waterfall_coll_call_enriched')}"
).collect()[0]["n"]
chk("export distinct accounts", wf_accts, 610183)

# uc2_t16_01_populations must still be the round-10 build (guards against a
# silent rebuild with different logic)
pop_row = spark.sql(f"""
    SELECT count_if(in_ledger_all)  AS ledger_all,
           count_if(in_ledger_exaa) AS ledger_exaa,
           count_if(touched_b1)     AS touched_b1
    FROM {T('uc2_t16_01_populations')}
""").collect()[0]
chk("populations: cleaned ledger (all)", pop_row["ledger_all"], 204323)
chk("populations: ex-AA ledger", pop_row["ledger_exaa"], 189146)
chk("populations: touched-B1 universe", pop_row["touched_b1"], 724848)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C3. Rebuild the SAS-flagged caller set (verbatim the round-10-verified filters)
# MAGIC
# MAGIC DQ1 (`DLNQT_CD_M1` = 1) + CPC eligible (upper/trim in OTHER / OTHERS /
# MAGIC COBRAND / PLCC) + non-chargeoff reason (NULL / blank / PLY / BLANK) +
# MAGIC `call_type_INB` contains 'INB'. Key = trimmed string, the tier-16
# MAGIC `acct_key` rule. Expected: 11,136 distinct accounts.

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

flagged_n = spark.sql("SELECT count(*) AS n FROM sas_flagged").collect()[0]["n"]
chk("SAS-flagged caller set", flagged_n, 11136)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C4. THE decomposition
# MAGIC
# MAGIC One class per flagged account, first-match-wins, against the tier-16
# MAGIC population layer. Classes 1-5 are the gap; class 6 must be exactly the
# MAGIC 9,389 ledger callers; class 7 must be empty.

# COMMAND ----------

spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW gap_classed AS
    SELECT
        f.acct_key,
        CASE
            WHEN p.acct_key IS NULL      THEN '1. no January account row (AWS snapshot)'
            WHEN NOT p.is_exaa           THEN '2. AA product on the AWS side'
            WHEN NOT p.cleaned           THEN '3. cleanup-removed (charged off before 2025)'
            WHEN p.eom_bucket = 0        THEN '4. EOM cured (SAS cycle code DQ1, AWS ladder current at 31 Jan)'
            WHEN p.eom_bucket >= 2       THEN '5. EOM bucket 2+ at 31 Jan'
            WHEN p.in_ledger_exaa        THEN '6. in the ex-AA ledger'
            ELSE '7. UNCLASSIFIED (must be empty)'
        END AS gap_class,
        p.max_bucket,
        p.eom_bucket,
        p.eom_bal,
        p.eom_cpc,
        p.cpc_class,
        p.first_b1_dt,
        p.touched_b1_class
    FROM sas_flagged f
    LEFT JOIN {T('uc2_t16_01_populations')} p
        ON p.acct_key = f.acct_key
""")

decomp = spark.sql("""
    SELECT gap_class,
           count(*) AS accounts,
           round(100.0 * count(*) / sum(count(*)) OVER (), 1) AS share_pct,
           round(sum(eom_bal), 0) AS jan_eom_balance
    FROM gap_classed
    GROUP BY gap_class
    ORDER BY gap_class
""")
decomp.show(truncate=False)

rows = {r["gap_class"][:1]: r["accounts"] for r in decomp.collect()}
total = sum(rows.values())
in_ledger = rows.get("6", 0)
unclassified = rows.get("7", 0)
gap = total - in_ledger - unclassified

chk("decomposition total", total, 11136)
chk("class 6 = ledger callers", in_ledger, 9389)
chk("class 7 = unclassified", unclassified, 0)
chk("classes 1-5 = the gap", gap, 1747)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C5. Cause color (read-only diagnostics, no asserts)
# MAGIC
# MAGIC These grids NAME the causes behind classes 1-5. Screenshot them with the
# MAGIC C4 grid. Do not transcribe raw account ids into any record.

# COMMAND ----------

# 5a. Classes 4 and 5: where were these accounts during January on the AWS
# ladder? (max_bucket = worst position in the month; first_b1_dt filled means
# the account DID touch bucket 1 mid-month.)
spark.sql("""
    SELECT gap_class,
           max_bucket,
           eom_bucket,
           count(*) AS accounts,
           count_if(first_b1_dt IS NOT NULL) AS touched_b1_in_jan,
           round(sum(eom_bal), 0) AS jan_eom_balance
    FROM gap_classed
    WHERE gap_class LIKE '4.%' OR gap_class LIKE '5.%'
    GROUP BY gap_class, max_bucket, eom_bucket
    ORDER BY gap_class, max_bucket, eom_bucket
""").show(50, truncate=False)

# COMMAND ----------

# 5b. Class 2: the CPC disagreement. The SAS filter already excluded AA, so
# any account here carries a non-AA CPC_FLAG_NW in the export but an AA-family
# eom_cpc on the AWS side. Show the AWS code they carry.
spark.sql("""
    SELECT gap_class, cpc_class, eom_cpc, count(*) AS accounts
    FROM gap_classed
    WHERE gap_class LIKE '2.%'
    GROUP BY gap_class, cpc_class, eom_cpc
    ORDER BY accounts DESC
""").show(50, truncate=False)

# COMMAND ----------

# 5c. Class 1: accounts with no January row in the AWS snapshot. Round 10
# matched all 11,136 into acct_monthly across Dec-2024..Mar-2025, so any
# account here has rows in OTHER months only. Show which months.
spark.sql(f"""
    SELECT m.ym, count(DISTINCT m.acct_key) AS accounts_with_row
    FROM gap_classed g
    JOIN {T('uc2_t16_00_acct_monthly')} m
        ON m.acct_key = g.acct_key
    WHERE g.gap_class LIKE '1.%'
    GROUP BY m.ym
    ORDER BY m.ym
""").show(truncate=False)

# A handful of raw ids for eyeballing ON THE VDI ONLY (never transcribed):
spark.sql("""
    SELECT acct_key FROM gap_classed WHERE gap_class LIKE '1.%' LIMIT 10
""").show(truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## C6. Verdict

# COMMAND ----------

print("=" * 78)
print("PHASE 1 VERDICT: caller-gap decomposition")
print("=" * 78)
for name, actual, expected, status in RESULTS:
    print(f"{status:4}  {name:45} {actual:>12,}  (expected {expected:,})")
print("-" * 78)
decomp.show(truncate=False)
print("Standing note: the replicated flag population is 11,136 vs the recorded")
print("SAS count 11,154. That residual is a SAS-side exact-spelling replication")
print("question ([OPEN, small], owner Ravi) and does not affect this verdict:")
print("the decomposition explains the 11,136-vs-9,389 structure account by")
print("account. Screenshot the C4 + C5 grids and this verdict block; the class")
print("counts land in the bridge-round-11 record.")
print("=" * 78)
