# Databricks notebook source
# =====================================================================
# 04_transcript_frame.py
# The transcript-eligible sampling frame (feeds the Copilot discovery work).
#
#   Reads the folded ledger uc2_t16_01s_populations_<vintage>, the episode table
#   uc2_t16_02n_episodes, and the roll cut uc2_t16_03r_roll_<vintage>; writes
#   uc2_t16_04t_frame_<vintage>.
#
# WHAT THIS BUILDS
#   uc2_t16_04t_frame_<vintage> - one row per in-window inbound call that has a
#   transcript, tagged with its window bucket (post-due / roll-cohort), the account
#   key, the SAS captured_sas outcome, the days-since-statement position, AND the
#   full review-context column set (below). The transcript AI runs over the roll
#   customers' inbound calls to find solvable intent, with enough per-account /
#   per-call context to read against.
#
# REVIEW-CONTEXT COLUMNS ON THE FRAME (real source columns or derived):
#   Call grain (from uc2_t16_02n_episodes; real call-table columns per dict):
#     contactid, acct_key, call_dt, producttype, queue, routingprofile,
#     department (the call routing dept), segment, stmt_dt, due_dt_derived
#     (= stmt_dt + 25; no explicit due-date column in fmt), days_since_stmt_dt
#   Account grain (from the folded ledger; real fmt / SAS columns):
#     cpc_class, min_due_amt (fmt PAYMT_MIN_DUE_AMT), eop_bal_m1, cr_lmt_m1,
#     utilization_m1, last_pay_vs_min_due, paymt_amt_m1/m2, paymt_last_amt, pay_dt,
#     dlnqt_cd_m1/m2, stg_cd_m1, max_bucket, eom_bucket, captured_sas
#   Frame classification: on_roll_cohort, window_bucket, the 25/31/overall flags,
#     has_transcript
#
# THE ELIGIBLE CALL SET
#   inbound calls (uc2_t16_02n_episodes, all have acctid) that are either
#     - call_31_window_f = 1  (post-due / actionable band), or
#     - on a DQ1->DQ2 roll-cohort account (the impairment-heavy set),
#   joined to the transcript table on contactid, transcript present in scan window.
#
# DIRECTION / ACCTID HANDLING
#   INBOUND calls carry acctid (that is why 02n only keeps INBOUND). OUTBOUND
#   calls have acctid missing in this source, so they cannot be account-joined and
#   are out of scope for this frame. [VERIFY: transfer/callback acctid] transfer
#   and callback inbound contactids may carry a different or missing acctid;
#   whether they should be re-keyed is not resolved here - noted, not dropped
#   silently.
#
#   The frame emits NO transcript text - only contactid + join keys + flags +
#   numeric context, so a screenshot of the sizing is safe. Transcript text pull
#   stays in the masked export step (out of scope for this module).
#
#   Every print/display below is prefixed with this file name for screenshots.
# =====================================================================

# COMMAND ----------

# ---------------------------------------------------------------------
# SETUP - catalog/schema, transcript handle, table handles, scan bounds.
# ---------------------------------------------------------------------
import datetime as _dt

CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
ANCHOR_YM = "202501"
FMT_CATALOG = "634153504162_glue_connection_catalog"
CC_CATALOG = "062108867742_glue_connectivity_catalog"

try:
    dbutils.widgets.text("CATALOG", CATALOG);       CATALOG = dbutils.widgets.get("CATALOG")
    dbutils.widgets.text("SCHEMA", SCHEMA);         SCHEMA = dbutils.widgets.get("SCHEMA")
    dbutils.widgets.text("ANCHOR_YM", ANCHOR_YM);   ANCHOR_YM = dbutils.widgets.get("ANCHOR_YM")
    dbutils.widgets.text("FMT_CATALOG", FMT_CATALOG); FMT_CATALOG = dbutils.widgets.get("FMT_CATALOG")
    dbutils.widgets.text("CC_CATALOG", CC_CATALOG);   CC_CATALOG = dbutils.widgets.get("CC_CATALOG")
except NameError:
    pass

DB = f"{CATALOG}.{SCHEMA}"
TX = f"`{CC_CATALOG}`.contactcenter_bdp_db.transcript"

_a0 = _dt.date(int(ANCHOR_YM[:4]), int(ANCHOR_YM[4:6]), 1)
_mm = lambda d, k: _dt.date(d.year + (d.month - 1 + k) // 12, (d.month - 1 + k) % 12 + 1, 1)

# transcript effdt scan bounds (prev month .. anchor+3) - scan-pruning guard only
EFFDT_SCAN_START = _mm(_a0, -1).isoformat()   # 2024-12-01
EFFDT_SCAN_END = _mm(_a0, 3).isoformat()      # 2025-04-01

T_01S = f"{DB}.uc2_t16_01s_populations_{ANCHOR_YM}"
T_02N = f"{DB}.uc2_t16_02n_episodes"
T_03R = f"{DB}.uc2_t16_03r_roll_{ANCHOR_YM}"
T_04T = f"{DB}.uc2_t16_04t_frame_{ANCHOR_YM}"

print(f"[04_transcript_frame.py] SETUP OK: vintage {ANCHOR_YM}; layers -> {DB}")
# ---------------------------------------------------------------------
# end of SETUP
# ---------------------------------------------------------------------

# COMMAND ----------

# ---------------------------------------------------------------------
# 04t. The sampling frame. Eligible = post-due(31) call OR roll-cohort account,
# with a transcript present. Window bucket labels each eligible call. Full
# review-context columns joined from the folded ledger (account grain) and the
# episode table (call grain).
# ---------------------------------------------------------------------
spark.sql(f"""
CREATE OR REPLACE TABLE {T_04T} AS
WITH roll_accts AS (
    SELECT acct_key FROM {T_03R} WHERE rolled_dq1_dq2
),
ledger AS (
    -- account-grain review context from the folded ledger (real fmt / SAS columns)
    SELECT acct_key, in_sas_ledger, captured_sas, cpc_class,
           dlnqt_cd_m1, dlnqt_cd_m2, stg_cd_m1,
           min_due_amt, eop_bal_m1, cr_lmt_m1, utilization_m1, last_pay_vs_min_due,
           paymt_amt_m1, paymt_amt_m2, paymt_last_amt, pay_dt,
           max_bucket, eom_bucket
    FROM {T_01S}
),
eligible_calls AS (
    -- inbound in-window calls: post-due(31) band, or any call on a roll-cohort account
    SELECT c.acct_key, c.acct_num, c.contactid, c.call_dt, c.call_month,
           c.stmt_dt, c.due_dt_derived, c.days_since_stmt_dt,
           c.producttype, c.queue, c.routingprofile, c.department, c.segment,
           c.call_25_window_f, c.call_31_window_f, c.call_overall_f,
           l.in_sas_ledger, l.captured_sas, l.cpc_class,
           l.dlnqt_cd_m1, l.dlnqt_cd_m2, l.stg_cd_m1,
           l.min_due_amt, l.eop_bal_m1, l.cr_lmt_m1, l.utilization_m1, l.last_pay_vs_min_due,
           l.paymt_amt_m1, l.paymt_amt_m2, l.paymt_last_amt, l.pay_dt,
           l.max_bucket, l.eom_bucket,
           (r.acct_key IS NOT NULL) AS on_roll_cohort,
           CASE
             WHEN r.acct_key IS NOT NULL AND c.call_31_window_f = 1 THEN 'roll-cohort post-due'
             WHEN r.acct_key IS NOT NULL                            THEN 'roll-cohort other-window'
             WHEN c.call_31_window_f = 1                            THEN 'post-due (non-roll)'
             ELSE 'other in-window'
           END AS window_bucket
    FROM {T_02N} c
    JOIN ledger l         ON l.acct_key = c.acct_key
    LEFT JOIN roll_accts r ON r.acct_key = c.acct_key
    WHERE l.in_sas_ledger
      AND (c.call_31_window_f = 1 OR r.acct_key IS NOT NULL)
),
tx_ids AS (
    -- distinct contactids that actually have transcript content in the scan window
    SELECT DISTINCT t.contactid
    FROM {TX} t
    JOIN (SELECT DISTINCT contactid FROM eligible_calls) e ON e.contactid = t.contactid
    WHERE t.content IS NOT NULL
      AND t.effdt >= '{EFFDT_SCAN_START}' AND t.effdt < '{EFFDT_SCAN_END}'
)
SELECT e.contactid,
       e.acct_key,
       e.acct_num,
       e.window_bucket,
       e.on_roll_cohort,
       e.call_25_window_f,
       e.call_31_window_f,
       e.call_overall_f,
       e.days_since_stmt_dt,
       e.stmt_dt,
       e.due_dt_derived,
       e.call_dt,
       e.call_month,
       -- call-context review columns (real call-table columns)
       e.producttype,
       e.queue,
       e.routingprofile,
       e.department,
       e.segment,
       -- account-context review columns (real fmt / SAS columns; derived ratios)
       e.cpc_class,
       e.dlnqt_cd_m1,
       e.dlnqt_cd_m2,
       e.stg_cd_m1,
       e.max_bucket,
       e.eom_bucket,
       e.min_due_amt,
       e.eop_bal_m1,
       e.cr_lmt_m1,
       e.utilization_m1,
       e.last_pay_vs_min_due,
       e.paymt_amt_m1,
       e.paymt_amt_m2,
       e.paymt_last_amt,
       e.pay_dt,
       e.captured_sas,
       (x.contactid IS NOT NULL) AS has_transcript
FROM eligible_calls e
LEFT JOIN tx_ids x ON x.contactid = e.contactid
""")
print(f"[04_transcript_frame.py] built {T_04T}")

# COMMAND ----------

# ---------------------------------------------------------------------
# Frame coverage: eligible calls / accounts and how many have a transcript.
# ---------------------------------------------------------------------
print("[04_transcript_frame.py] transcript-frame coverage:")
display(spark.sql(f"""
SELECT count(1)                          AS eligible_calls,
       count(DISTINCT acct_key)          AS eligible_accounts,
       count_if(has_transcript)          AS calls_with_transcript,
       count(DISTINCT CASE WHEN has_transcript THEN acct_key END) AS accounts_with_transcript,
       count_if(on_roll_cohort)          AS roll_cohort_calls,
       count_if(on_roll_cohort AND has_transcript) AS roll_cohort_calls_with_tx
FROM {T_04T}
"""))

print("[04_transcript_frame.py] eligible calls by window bucket (with-transcript split):")
display(spark.sql(f"""
SELECT window_bucket,
       count(1)                 AS calls,
       count(DISTINCT acct_key) AS accounts,
       count_if(has_transcript) AS calls_with_transcript,
       count_if(captured_sas)   AS captured_sas_calls
FROM {T_04T}
GROUP BY window_bucket ORDER BY window_bucket
"""))

# COMMAND ----------

# ---------------------------------------------------------------------
# The frame that feeds Copilot discovery, with the review-context columns the
# reviewers read against: contactid, acct_key, call dept/queue, cpc_class,
# due date, balance/limit/utilization, min-due, payment signals, stage, bucket.
# NO transcript text emitted here (screenshot-safe sizing view).
# ---------------------------------------------------------------------
print("[04_transcript_frame.py] Copilot review frame (roll-cohort, transcript present) - top 200:")
display(spark.sql(f"""
SELECT contactid, acct_key, window_bucket, captured_sas,
       call_dt, days_since_stmt_dt, stmt_dt, due_dt_derived,
       department, queue, routingprofile, segment, cpc_class,
       dlnqt_cd_m1, dlnqt_cd_m2, stg_cd_m1, max_bucket,
       min_due_amt, eop_bal_m1, cr_lmt_m1, utilization_m1, last_pay_vs_min_due,
       paymt_amt_m1, paymt_amt_m2, paymt_last_amt, pay_dt
FROM {T_04T}
WHERE has_transcript AND on_roll_cohort
ORDER BY days_since_stmt_dt, acct_key
LIMIT 200
"""))

print("[04_transcript_frame.py] [VERIFY: transfer/callback acctid] outbound calls "
      "carry no acctid in this source and are excluded; transfer/callback inbound "
      "contactids may carry a different/missing acctid - not separately resolved here.")

print(f"[04_transcript_frame.py] 04_transcript_frame complete: uc2_t16_04t_frame_{ANCHOR_YM}")
