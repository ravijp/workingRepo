# Databricks notebook source
# MAGIC %md
# MAGIC # B_window_probe - confirm the statement-window collapse cause (run-once, aggregates only)
# MAGIC
# MAGIC Context: the first re-anchored run collapsed 04s to 25 episodes / captured_sas = 0,
# MAGIC vs Ishant's VERIFIED ~19,025 in-window accounts on the same 186,013 base.
# MAGIC Diagnosis (WINDOW_COLLAPSE_DIAGNOSIS.md): the OLD stmt_anchor used
# MAGIC max(stmt_last_dt) = one date per account, which for a January call was
# MAGIC usually a FUTURE (Feb/Mar) statement -> datediff < 0 -> in_stmt_window = 0
# MAGIC for nearly everyone. The fix (applied in B02) as-of joins each call to its
# MAGIC OWN cycle's statement (most recent stmt_last_dt <= call_dt).
# MAGIC
# MAGIC This probe CONFIRMS the cause and SIZES the target. Run P1 first.
# MAGIC   - Run it on the OLD 02n (before the B02 fix) to see the defect fingerprint,
# MAGIC     OR after the fix to confirm recovery. Both readings are useful.
# MAGIC Aggregates only: no transcript text, no PII, no row-level export.

# COMMAND ----------

CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
DB = f"{CATALOG}.{SCHEMA}"
FMT = "`634153504162_glue_connection_catalog`.fmt_acct_dba.fmt_acct_c"
EPI = f"{DB}.uc2_t16_02n_episodes"

# COMMAND ----------

# MAGIC %md
# MAGIC ## P1 - reconcile the 22,293-vs-25 gap (THE decisive probe)
# MAGIC On the OLD (pre-fix) 02n: std_episodes ~= 25 while in_window_calls ~= tens of
# MAGIC thousands CONFIRMS the single-max anchor as the cause. On the FIXED 02n:
# MAGIC std_episodes should rise to ~19-22k, matching in_window_calls and Ishant.

# COMMAND ----------

spark.sql(f"""
    SELECT
      count_if(is_episode_std = 1)                                   AS std_episodes,
      count_if(in_stmt_window = 1 AND is_biz = 0 AND within_effdt_cap = 1
               AND acct_key IS NOT NULL AND acct_key <> '')          AS in_window_calls,
      count(DISTINCT CASE WHEN is_episode_std = 1 THEN acct_key END)  AS std_callers
    FROM {EPI}
""").show(truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## P2 - the sign of days_since_stmt_dt (proves the anchor direction)
# MAGIC OLD 02n: band '1 negative (stmt AFTER call)' should carry most of the mass.
# MAGIC FIXED 02n: that mass moves into band '2 in-window 0..55'.

# COMMAND ----------

spark.sql(f"""
    SELECT CASE WHEN days_since_stmt_dt IS NULL THEN '0 no anchor'
                WHEN days_since_stmt_dt <  0     THEN '1 negative (stmt AFTER call)'
                WHEN days_since_stmt_dt < 56     THEN '2 in-window 0..55'
                ELSE                                  '3 >=56' END AS band,
           count(*) AS call_rows,
           count(DISTINCT acct_key) AS accts
    FROM {EPI}
    GROUP BY 1 ORDER BY 1
""").show(truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## P3 - distinct statement dates per account (proves one max collapses a real sequence)
# MAGIC Most accounts should show 3-4 distinct statement dates in the scan window,
# MAGIC so max() discarded the earlier per-cycle statements the calls actually belong to.

# COMMAND ----------

spark.sql(f"""
    SELECT n_stmt_dates, count(*) AS accts
    FROM (
      SELECT cast(try_cast(extnl_acct_id AS bigint) AS string) AS acct_key,
             count(DISTINCT try_cast(stmt_last_dt AS date)) AS n_stmt_dates
      FROM {FMT}
      WHERE sfx_nbr = 0 AND eff_dt >= '20241201' AND eff_dt < '20250401'
        AND stmt_last_dt IS NOT NULL
      GROUP BY 1
    ) GROUP BY 1 ORDER BY 1
""").show(50, truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## P4 - target size after the fix (calls with a prior statement in-window)
# MAGIC Expectation: a count near Ishant's ~19,025 (post-due) + pre-due. If P4 lands
# MAGIC there, the fixed as-of anchor recovers the population to his verified scale.

# COMMAND ----------

spark.sql(f"""
    SELECT count(*) AS calls_with_prior_stmt_in_window,
           count(DISTINCT c.acct_key) AS accts
    FROM {EPI} c
    JOIN (
      SELECT cast(try_cast(extnl_acct_id AS bigint) AS string) AS acct_key,
             try_cast(stmt_last_dt AS date) AS stmt_dt
      FROM {FMT}
      WHERE sfx_nbr = 0 AND eff_dt >= '20241201' AND eff_dt < '20250401'
        AND stmt_last_dt IS NOT NULL
      GROUP BY 1, 2
    ) s
      ON s.acct_key = c.acct_key
     AND s.stmt_dt <= c.call_dt
     AND datediff(c.call_dt, s.stmt_dt) < 56
""").show(truncate=False)

# COMMAND ----------

print("B_window_probe: done. P1 is decisive. Compare in_window_calls vs "
      "std_episodes: equal (~19-22k) after the B02 fix = recovered; a 22k-vs-25 "
      "split (old 02n) = the single-max anchor defect confirmed. P4 sizes the "
      "target against Ishant's verified ~19,025.")
