# Databricks notebook source
# MAGIC %md
# MAGIC # B_stmt_distribution - the Story-B re-anchor distribution (run ONCE)
# MAGIC
# MAGIC A run-once sibling (like the _checks files, but descriptive, not raising).
# MAGIC It reports the statement-frame picture the re-anchor produced. Nothing
# MAGIC here asserts a locked value: the numbers MOVED off January ON PURPOSE, so
# MAGIC every print is measure-mode. It reads the re-anchored 04s table and the
# MAGIC round-12 January-frame record values, and prints them side by side.
# MAGIC
# MAGIC Two sections:
# MAGIC   (a) STORY-B PROOF: leaked vs captured episode counts AND dollars
# MAGIC       (eop_bal_m1, gross_loss_12m_amt) across the 5-day statement buckets,
# MAGIC       with pre-due vs post-due subtotals - showing WHERE leakage
# MAGIC       concentrates in the statement cycle.
# MAGIC   (b) SHIFT EXPLANATION: the OLD calendar-January totals (episodes 13,486 /
# MAGIC       callers 11,136 / captured_sas 8,037 / leaked_sas 1,801 / W_s 1,646,
# MAGIC       round-12 record) printed side by side with the NEW statement-frame
# MAGIC       totals the re-anchored 04s now produces, plus the count of episodes
# MAGIC       DROPPED for falling outside all statement windows (the cause of the
# MAGIC       move).
# MAGIC
# MAGIC Run order: after B02 (statement re-anchor) then B02b (04s carries the
# MAGIC statement columns). Reads only; builds no tables.

# COMMAND ----------

# =====================================================================
# SETUP - keep in sync across B00/B01/B02/B02b/B03 (B00 is the canonical copy).
# B_stmt_distribution reads only the 04s table (no date literals of its own),
# so the derived-window block is not needed here.
# =====================================================================
CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
ANCHOR_YM = "202501"
FMT_CATALOG = "634153504162_glue_connection_catalog"
CC_CATALOG = "062108867742_glue_connectivity_catalog"

try:
    dbutils.widgets.text("CATALOG", CATALOG);           CATALOG = dbutils.widgets.get("CATALOG")
    dbutils.widgets.text("SCHEMA", SCHEMA);             SCHEMA = dbutils.widgets.get("SCHEMA")
    dbutils.widgets.text("ANCHOR_YM", ANCHOR_YM);       ANCHOR_YM = dbutils.widgets.get("ANCHOR_YM")
    dbutils.widgets.text("FMT_CATALOG", FMT_CATALOG);   FMT_CATALOG = dbutils.widgets.get("FMT_CATALOG")
    dbutils.widgets.text("CC_CATALOG", CC_CATALOG);     CC_CATALOG = dbutils.widgets.get("CC_CATALOG")
except NameError:
    pass

DB = f"{CATALOG}.{SCHEMA}"
OUT = f"{DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}"
EPI = f"{DB}.uc2_t16_02n_episodes"

# The OLD calendar-January-frame totals, from the round-12 record
# (bridge-round12-phase2-sas-spine-2026-07-16.md). These are the PRE-re-anchor
# values; the re-anchored table is expected to move OFF them.
JAN_FRAME = {
    "episodes": 13486,
    "callers": 11136,
    "captured_sas accounts": 8037,
    "leaked_sas accounts": 1801,
    "W_s accounts": 1646,
}


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


print(f"SETUP OK: vintage {ANCHOR_YM}; reading {OUT}")
# =====================================================================
# end of SETUP
# =====================================================================

# COMMAND ----------

# preconditions
assert spark.catalog.tableExists(OUT), \
    f"PRECONDITION MISS: {OUT} missing - run B02 (re-anchor) then B02b first"
assert spark.catalog.tableExists(EPI), \
    f"PRECONDITION MISS: {EPI} missing - run B02 first"
# the re-anchor columns must be present, else this is the OLD 04s
_cols = [c.lower() for c in spark.sql(f"SELECT * FROM {OUT} LIMIT 0").columns]
for _c in ["stmt_dt", "days_since_stmt_dt", "stmt_5day_bucket", "pre_due_f", "post_due_f"]:
    assert _c in _cols, \
        f"PRECONDITION MISS: {OUT} has no '{_c}' column - re-run the re-anchored B02/B02b"
print(f"PASS  re-anchor columns present on {OUT}")

# COMMAND ----------

# MAGIC %md
# MAGIC ## (a) STORY-B PROOF: leakage across the statement cycle
# MAGIC
# MAGIC Episode counts and dollars by 5-day statement bucket, split leaked vs
# MAGIC captured. Dollars are ACCOUNT-grain export columns; an account can appear
# MAGIC in more than one bucket, so money compares WITHIN a row only and is NEVER
# MAGIC added down a column. The bucket order runs pre-due (00-04 .. 20-24) then
# MAGIC post-due (25-29 .. 50-54); the due day is 25.

# COMMAND ----------

print("=== (a) STORY-B PROOF: episodes and dollars by 5-day statement bucket ===")
print("    money is within-row only (account-grain columns); do NOT add a dollar")
print("    column down the buckets. leaked = leaked_sas; captured = captured_sas.")
spark.sql(f"""
    WITH ep AS (
        SELECT stmt_5day_bucket, stmt_5day_bucket_start, pre_due_f, post_due_f,
               contactid, acct_key, captured_sas, leaked_sas,
               eop_bal_m1, gross_loss_12m_amt
        FROM {OUT}
    ),
    counts AS (
        SELECT stmt_5day_bucket,
               max(stmt_5day_bucket_start) AS ord,
               max(CASE WHEN post_due_f = 1 THEN 1 ELSE 0 END) AS is_post_due,
               count(*) AS episodes,
               count_if(leaked_sas) AS leaked_eps,
               count_if(captured_sas) AS captured_eps
        FROM ep GROUP BY 1
    ),
    leaked_money AS (
        SELECT stmt_5day_bucket,
               round(sum(eop_bal_m1), 0) AS leaked_eop_bal,
               round(sum(gross_loss_12m_amt), 0) AS leaked_gl12m
        FROM (SELECT DISTINCT stmt_5day_bucket, acct_key, eop_bal_m1, gross_loss_12m_amt
              FROM ep WHERE leaked_sas)
        GROUP BY 1
    ),
    captured_money AS (
        SELECT stmt_5day_bucket,
               round(sum(eop_bal_m1), 0) AS captured_eop_bal,
               round(sum(gross_loss_12m_amt), 0) AS captured_gl12m
        FROM (SELECT DISTINCT stmt_5day_bucket, acct_key, eop_bal_m1, gross_loss_12m_amt
              FROM ep WHERE captured_sas)
        GROUP BY 1
    )
    SELECT c.stmt_5day_bucket,
           CASE WHEN c.stmt_5day_bucket = 'outside 0-55 days' THEN 'n/a'
                WHEN c.is_post_due = 1 THEN 'post-due' ELSE 'pre-due' END AS due_side,
           c.episodes, c.leaked_eps, c.captured_eps,
           lm.leaked_eop_bal, lm.leaked_gl12m,
           km.captured_eop_bal, km.captured_gl12m
    FROM counts c
    LEFT JOIN leaked_money lm ON lm.stmt_5day_bucket = c.stmt_5day_bucket
    LEFT JOIN captured_money km ON km.stmt_5day_bucket = c.stmt_5day_bucket
    ORDER BY (c.stmt_5day_bucket = 'outside 0-55 days'), c.ord
""").show(50, truncate=False)

# COMMAND ----------

print("=== (a) pre-due vs post-due subtotals (episodes; leaked/captured dollars within row) ===")
spark.sql(f"""
    WITH ep AS (
        SELECT CASE WHEN pre_due_f = 1 THEN '1. pre-due (days 00-24)'
                    WHEN post_due_f = 1 THEN '2. post-due (days 25-55)'
                    ELSE '3. outside 0-55 days' END AS due_side,
               acct_key, captured_sas, leaked_sas, eop_bal_m1, gross_loss_12m_amt
        FROM {OUT}
    ),
    counts AS (
        SELECT due_side, count(*) AS episodes,
               count_if(leaked_sas) AS leaked_eps,
               count_if(captured_sas) AS captured_eps
        FROM ep GROUP BY 1
    ),
    leaked_money AS (
        SELECT due_side, round(sum(gross_loss_12m_amt), 0) AS leaked_gl12m
        FROM (SELECT DISTINCT due_side, acct_key, gross_loss_12m_amt FROM ep WHERE leaked_sas)
        GROUP BY 1
    ),
    captured_money AS (
        SELECT due_side, round(sum(gross_loss_12m_amt), 0) AS captured_gl12m
        FROM (SELECT DISTINCT due_side, acct_key, gross_loss_12m_amt FROM ep WHERE captured_sas)
        GROUP BY 1
    )
    SELECT c.due_side, c.episodes, c.leaked_eps, c.captured_eps,
           lm.leaked_gl12m, km.captured_gl12m
    FROM counts c
    LEFT JOIN leaked_money lm ON lm.due_side = c.due_side
    LEFT JOIN captured_money km ON km.due_side = c.due_side
    ORDER BY c.due_side
""").show(50, truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## (b) SHIFT EXPLANATION: January frame vs statement frame, side by side
# MAGIC
# MAGIC The re-anchor DROPPED every episode whose call_dt fell outside all
# MAGIC statement windows. That drop is the cause of the move. The 04s table is
# MAGIC already the re-anchored (in-window) set, so its totals ARE the new frame.
# MAGIC The dropped count is measured on the 02n episode table (standard episodes
# MAGIC before the re-anchor keep flag vs after).

# COMMAND ----------

print("=== (b) SHIFT EXPLANATION: OLD January frame vs NEW statement frame ===")
print("    the numbers moved on purpose; this is descriptive, not a stop rule.")

_new = spark.sql(f"""
    SELECT count(*) AS episodes,
           count(DISTINCT acct_key) AS callers,
           count(DISTINCT CASE WHEN captured_sas THEN acct_key END) AS captured_accts,
           count(DISTINCT CASE WHEN leaked_sas   THEN acct_key END) AS leaked_accts,
           count(DISTINCT CASE WHEN w_s_flag     THEN acct_key END) AS w_s_accts
    FROM {OUT}
""").first()

_rows = [
    ("episodes", JAN_FRAME["episodes"], _new["episodes"]),
    ("callers", JAN_FRAME["callers"], _new["callers"]),
    ("captured_sas accounts", JAN_FRAME["captured_sas accounts"], _new["captured_accts"]),
    ("leaked_sas accounts", JAN_FRAME["leaked_sas accounts"], _new["leaked_accts"]),
    ("W_s accounts", JAN_FRAME["W_s accounts"], _new["w_s_accts"]),
]
print(f"{'metric':28} {'Jan frame (round-12)':>22} {'stmt frame (new)':>18} {'delta':>12}")
for _name, _old, _newv in _rows:
    _d = _newv - _old
    print(f"{_name:28} {fmt(_old):>22} {fmt(_newv):>18} {(('+' if _d >= 0 else '') + fmt(_d)):>12}")

# COMMAND ----------

# the cause of the move: standard first-inbound-per-day episodes that FELL
# OUTSIDE all statement windows (in_stmt_window = 0). These are the call-days
# the re-anchor dropped from the episode/caller population. Grain: 02n episode
# rows (contactid), before the statement keep flag.
print("=== (b) episodes DROPPED for falling outside all statement windows ===")
_drop = spark.sql(f"""
    WITH first_inbound AS (
        -- the first-inbound-per-day survivor WITHOUT the statement keep flag,
        -- reconstructed on 02n so we can count what the re-anchor removed
        SELECT acct_key, contactid, call_dt, in_stmt_window, stmt_dt,
               row_number() OVER (PARTITION BY acct_key, call_dt ORDER BY contactid) AS rn
        FROM {EPI}
        WHERE is_biz = 0 AND within_effdt_cap = 1
          AND acct_key IS NOT NULL AND acct_key <> ''
    ),
    fi AS (SELECT * FROM first_inbound WHERE rn = 1)
    SELECT
        CASE WHEN stmt_dt IS NULL THEN '1. no statement date (no fmt anchor)'
             WHEN in_stmt_window = 0 THEN '2. call-day outside [stmt_dt, stmt_dt+56)'
             ELSE '3. in window (kept)' END AS disposition,
        count(*) AS first_inbound_call_days,
        count(DISTINCT acct_key) AS accounts
    FROM fi
    GROUP BY 1 ORDER BY 1
""")
_drop.show(50, truncate=False)
print("note: rows (1) and (2) are the DROPPED call-days that caused the counts to")
print("      move; row (3) is the kept in-window set that feeds the re-anchored 04s.")

# COMMAND ----------

print("B_stmt_distribution complete: (a) the statement-cycle leakage picture and")
print("(b) the January-vs-statement shift with the dropped-episode cause. All")
print("prints are descriptive (measure-mode). The move off the January values is")
print("intended - the statement-frame numbers are the Story-B deliverable.")
