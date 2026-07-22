# Databricks notebook source
# =====================================================================
# B_ishant / 04_transcript_frame.py
# The transcript-eligible sampling frame (feeds the Copilot discovery work).
#
# WHAT THIS BUILDS
#   uc2_ish_04t_frame - one row per in-window inbound call that has a transcript,
#   tagged with its window bucket (post-due / roll-cohort), the account key, the
#   SAS captured_sas outcome, and the days-since-statement position. This is the
#   sampling frame Namit described: run the transcript AI over the ~2,879 roll
#   customers' inbound calls to find solvable intent.
#
# THE ELIGIBLE CALL SET
#   inbound calls (uc2_ish_02n_calls, all have acctid) that are either
#     - post_due_f = 1  (in the actionable band), or
#     - on a DQ1->DQ2 roll-cohort account (the impairment-heavy set),
#   joined to the transcript table on contactid, transcript present in scan window.
#
# DIRECTION / ACCTID HANDLING
#   INBOUND calls carry acctid (that is why 02n only keeps INBOUND). OUTBOUND
#   calls have acctid missing in this source, so they cannot be account-joined and
#   are out of scope for this frame. [VERIFY: transfer / callback contactids may
#   carry a different or missing acctid; not resolved here - noted, not dropped
#   silently.]
#
#   The frame emits NO transcript text - only contactid + join keys + flags, so a
#   screenshot of the sizing is safe. Transcript text pull stays in the masked
#   export step (out of scope for this module).
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

# transcript effdt scan bounds (Dec24 .. Apr25) - scan-pruning guard only
EFFDT_SCAN_START = _mm(_a0, -1).isoformat()   # 2024-12-01
EFFDT_SCAN_END = _mm(_a0, 3).isoformat()      # 2025-04-01

print(f"[B_ishant/04_transcript_frame.py] SETUP OK: vintage {ANCHOR_YM}; layers -> {DB}")
# ---------------------------------------------------------------------
# end of SETUP
# ---------------------------------------------------------------------

# COMMAND ----------

# ---------------------------------------------------------------------
# Preconditions: the call classification + the roll cohort exist; transcript reachable.
# ---------------------------------------------------------------------
for _t in ["uc2_ish_02n_calls", "uc2_ish_02s_pop", "uc2_ish_03r_roll"]:
    if not spark.catalog.tableExists(f"{DB}.{_t}"):
        raise AssertionError(f"[B_ishant/04_transcript_frame.py] {DB}.{_t} missing - run 01/02/03 first")
if not spark.catalog.tableExists(f"{CC_CATALOG}.contactcenter_bdp_db.transcript"):
    raise AssertionError(f"[B_ishant/04_transcript_frame.py] transcript table not reachable: {TX}")
print("[B_ishant/04_transcript_frame.py] preconditions OK: 02n / 02s / 03r present; transcript reachable")

# COMMAND ----------

# ---------------------------------------------------------------------
# 04t. The sampling frame. Eligible = post-due call OR roll-cohort account,
# with a transcript present. Window bucket labels each eligible call.
# ---------------------------------------------------------------------
spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_ish_04t_frame AS
WITH roll_accts AS (
    SELECT acct_key FROM {DB}.uc2_ish_03r_roll WHERE rolled_dq1_dq2
),
ledger AS (
    SELECT acct_key, in_sas_ledger, captured_sas, cpc_class,
           dlnqt_cd_m1, dlnqt_cd_m2
    FROM {DB}.uc2_ish_02s_pop
),
eligible_calls AS (
    -- inbound in-window calls: post-due band, or any call on a roll-cohort account
    SELECT c.acct_key, c.acct_num, c.contactid, c.call_dt, c.call_month,
           c.stmt_dt, c.days_since_stmt_dt,
           c.pre_due_f, c.post_due_f, c.overall_f,
           l.in_sas_ledger, l.captured_sas, l.cpc_class,
           l.dlnqt_cd_m1, l.dlnqt_cd_m2,
           (r.acct_key IS NOT NULL) AS on_roll_cohort,
           CASE
             WHEN r.acct_key IS NOT NULL AND c.post_due_f = 1 THEN 'roll-cohort post-due'
             WHEN r.acct_key IS NOT NULL                       THEN 'roll-cohort other-window'
             WHEN c.post_due_f = 1                             THEN 'post-due (non-roll)'
             ELSE 'other in-window'
           END AS window_bucket
    FROM {DB}.uc2_ish_02n_calls c
    JOIN ledger l         ON l.acct_key = c.acct_key
    LEFT JOIN roll_accts r ON r.acct_key = c.acct_key
    WHERE l.in_sas_ledger
      AND (c.post_due_f = 1 OR r.acct_key IS NOT NULL)
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
       e.post_due_f,
       e.pre_due_f,
       e.overall_f,
       e.days_since_stmt_dt,
       e.stmt_dt,
       e.call_dt,
       e.call_month,
       e.cpc_class,
       e.dlnqt_cd_m1,
       e.dlnqt_cd_m2,
       e.captured_sas,
       (x.contactid IS NOT NULL) AS has_transcript
FROM eligible_calls e
LEFT JOIN tx_ids x ON x.contactid = e.contactid
""")
print(f"[B_ishant/04_transcript_frame.py] built {DB}.uc2_ish_04t_frame")

# COMMAND ----------

# ---------------------------------------------------------------------
# Frame coverage: eligible calls / accounts and how many have a transcript.
# ---------------------------------------------------------------------
print("[B_ishant/04_transcript_frame.py] transcript-frame coverage:")
display(spark.sql(f"""
SELECT count(1)                          AS eligible_calls,
       count(DISTINCT acct_key)          AS eligible_accounts,
       count_if(has_transcript)          AS calls_with_transcript,
       count(DISTINCT CASE WHEN has_transcript THEN acct_key END) AS accounts_with_transcript,
       count_if(on_roll_cohort)          AS roll_cohort_calls,
       count_if(on_roll_cohort AND has_transcript) AS roll_cohort_calls_with_tx
FROM {DB}.uc2_ish_04t_frame
"""))

print("[B_ishant/04_transcript_frame.py] eligible calls by window bucket (with-transcript split):")
display(spark.sql(f"""
SELECT window_bucket,
       count(1)                 AS calls,
       count(DISTINCT acct_key) AS accounts,
       count_if(has_transcript) AS calls_with_transcript,
       count_if(captured_sas)   AS captured_sas_calls
FROM {DB}.uc2_ish_04t_frame
GROUP BY window_bucket ORDER BY window_bucket
"""))

# COMMAND ----------

# ---------------------------------------------------------------------
# The frame that feeds Copilot discovery: contactid, acct_key, captured_sas,
# window bucket, days-since-statement - roll-cohort with-transcript rows first.
# NO transcript text emitted here (screenshot-safe sizing view).
# ---------------------------------------------------------------------
print("[B_ishant/04_transcript_frame.py] Copilot sampling frame (roll-cohort, transcript present) - top 200:")
display(spark.sql(f"""
SELECT contactid, acct_key, window_bucket, captured_sas,
       days_since_stmt_dt, stmt_dt, call_dt, cpc_class,
       dlnqt_cd_m1, dlnqt_cd_m2
FROM {DB}.uc2_ish_04t_frame
WHERE has_transcript AND on_roll_cohort
ORDER BY days_since_stmt_dt, acct_key
LIMIT 200
"""))

print("[B_ishant/04_transcript_frame.py] [VERIFY: transfer/callback acctid] outbound calls "
      "carry no acctid in this source and are excluded; transfer/callback inbound contactids "
      "may carry a different/missing acctid - not separately resolved here.")

print("[B_ishant/04_transcript_frame.py] 04_transcript_frame complete: uc2_ish_04t_frame")
