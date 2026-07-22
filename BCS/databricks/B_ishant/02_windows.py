# Databricks notebook source
# =====================================================================
# B_ishant / 02_windows.py
# The statement-window classification (Ishant's single-max anchor).
#
# THE ANCHOR (this is the whole point of the client-blessed approach)
#   stmt_anchor = max(stmt_last_dt) per account over the bounded fmt window.
#   ONE governing statement per account (the January-qualifying one). This
#   REPLACES our old per-call as-of join. The population is FIXED at all ledger
#   accounts; inbound calls only CLASSIFY an account by which window they land in.
#
# THE STATEMENT / DUE-DATE WINDOWS (all measured in days from stmt_dt = day 0)
#   pre-due  = [stmt_dt,      stmt_dt + 25)  -> days  0..24  (run-up to due date)
#   post-due = [stmt_dt + 25, stmt_dt + 56)  -> days 25..55  (missed due, pre-roll)
#   overall  = [stmt_dt,      stmt_dt + 56)  -> days  0..55  (whole cycle)
#   Account-level flag = max(window_flag) over the account's inbound calls.
#   post-due is the "actionable" band Anupam confirmed: before the due date a call
#   goes to the Cares team, after it goes to Collections.
#
# WHAT THIS BUILDS
#   1. uc2_ish_02n_calls  - inbound calls, one row per contactid, with pre/post/
#                           overall window flags against the single-max anchor.
#   2. uc2_ish_02s_pop    - the fixed ledger population with account-level
#                           accts_called_pre/post/overall flags joined on.
#   Then it prints the 4-stage x 3-window pivot (the funnel-with-windows table).
#
#   Expected (Ishant, 202501): on the 186,013 ledger base,
#     pre-due 6,778 / post-due 19,025 / overall 23,713 accounts called.
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
SAS_CSV_PATH = "/Volumes/cda_model_shared/ecm_cld_model/ecm_cld/collections_zenon/WATERFALL_COLL_CALL_V2_202501.csv"
FMT_CATALOG = "634153504162_glue_connection_catalog"
CC_CATALOG = "062108867742_glue_connectivity_catalog"

try:
    dbutils.widgets.text("CATALOG", CATALOG);           CATALOG = dbutils.widgets.get("CATALOG")
    dbutils.widgets.text("SCHEMA", SCHEMA);             SCHEMA = dbutils.widgets.get("SCHEMA")
    dbutils.widgets.text("ANCHOR_YM", ANCHOR_YM);       ANCHOR_YM = dbutils.widgets.get("ANCHOR_YM")
    dbutils.widgets.text("SAS_CSV_PATH", SAS_CSV_PATH); SAS_CSV_PATH = dbutils.widgets.get("SAS_CSV_PATH")
    dbutils.widgets.text("FMT_CATALOG", FMT_CATALOG);   FMT_CATALOG = dbutils.widgets.get("FMT_CATALOG")
    dbutils.widgets.text("CC_CATALOG", CC_CATALOG);     CC_CATALOG = dbutils.widgets.get("CC_CATALOG")
except NameError:
    pass

DB = f"{CATALOG}.{SCHEMA}"
FMT = f"`{FMT_CATALOG}`.fmt_acct_dba.fmt_acct_c"
CALL = f"`{CC_CATALOG}`.contactcenter_bdp_db.`call`"
TX = f"`{CC_CATALOG}`.contactcenter_bdp_db.transcript"

_a0 = _dt.date(int(ANCHOR_YM[:4]), int(ANCHOR_YM[4:6]), 1)
_mm = lambda d, k: _dt.date(d.year + (d.month - 1 + k) // 12, (d.month - 1 + k) % 12 + 1, 1)

MONTH_WIN_START = _mm(_a0, -1).strftime("%Y%m%d")   # 20241201 - fmt scan low edge
MONTH_WIN_END = _mm(_a0, 3).strftime("%Y%m%d")      # 20250401 - fmt scan high edge (excl)

NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"

# --- statement period + window constants (Ishant / 21-Jul meeting) ------------
# Statement period = the billing cycles whose due date falls in January.
#   STMT_START     = 2024-12-07  (statement dates on/after this qualify)
#   STMT_END_EXCL  = 2025-01-07  (statement dates before this qualify)
# Window day offsets from stmt_dt (day 0):
STMT_DUE_DAY = 25         # days: payment due-date marker; pre-due ends here
STMT_WINDOW_DAYS = 56     # days: overall window is [stmt_dt, stmt_dt+56)
# Call-table effdt scan bounds (Dec24 .. Apr25) - a scan-pruning guard only.
EFFDT_SCAN_START = _mm(_a0, -1).isoformat()   # 2024-12-01
EFFDT_HARD_END = _mm(_a0, 3).isoformat()      # 2025-04-01

STMT_START = "2024-12-07"
STMT_END_EXCL = "2025-01-07"

print(f"[B_ishant/02_windows.py] SETUP OK: vintage {ANCHOR_YM}; layers -> {DB}")
print(f"[B_ishant/02_windows.py] windows: pre-due [0,{STMT_DUE_DAY}) "
      f"post-due [{STMT_DUE_DAY},{STMT_WINDOW_DAYS}) overall [0,{STMT_WINDOW_DAYS}); "
      f"stmt period {STMT_START}..{STMT_END_EXCL}")
# ---------------------------------------------------------------------
# end of SETUP
# ---------------------------------------------------------------------

# COMMAND ----------

# ---------------------------------------------------------------------
# Preconditions: 01_accounts.py built the ledger + monthly layer.
# ---------------------------------------------------------------------
for _t in ["uc2_ish_00n_acct_monthly", "uc2_ish_01s_ledger"]:
    if not spark.catalog.tableExists(f"{DB}.{_t}"):
        raise AssertionError(f"[B_ishant/02_windows.py] {DB}.{_t} missing - run 01_accounts.py first")
spark.sql(f"REFRESH TABLE {CALL}")
print("[B_ishant/02_windows.py] preconditions OK: 00n / 01s present; call table refreshed")

# COMMAND ----------

# ---------------------------------------------------------------------
# 02n. Inbound calls classified against the SINGLE-MAX statement anchor.
#
#   stmt_anchor: max(stmt_last_dt) per account = ONE statement per account.
#   pre_due_f / post_due_f / overall_f are computed at CALL grain using
#   date_add(stmt_dt, 25) and date_add(stmt_dt, 56).
#   One row per contactid; first-inbound-per-day dedup, business cards dropped.
# ---------------------------------------------------------------------
spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_ish_02n_calls AS
WITH stmt_anchor_raw AS (
    -- every distinct account statement date over the bounded fmt window
    SELECT {NUM_KEY.format(c="extnl_acct_id")} AS acct_key,
           try_cast(extnl_acct_id AS bigint) AS acct_num,
           try_cast(stmt_last_dt AS date) AS stmt_dt
    FROM {FMT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '{MONTH_WIN_START}' AND eff_dt < '{MONTH_WIN_END}'
      AND stmt_last_dt IS NOT NULL
),
stmt_anchor AS (
    -- THE ANCHOR: max(stmt_last_dt) = one governing statement per account.
    -- (Ishant's stmt_anchor CTE. NOT a per-call as-of join.)
    SELECT acct_key, acct_num, max(stmt_dt) AS stmt_dt
    FROM stmt_anchor_raw
    GROUP BY acct_key, acct_num
),
calls_flagged AS (
    SELECT a.acct_key,
           a.acct_num,
           c.contactid,
           c.`date` AS call_dt,
           cast(date_trunc('month', c.`date`) AS date) AS call_month,
           c.initiationtimestamp,
           a.stmt_dt,
           datediff(c.`date`, a.stmt_dt) AS days_since_stmt_dt,
           CASE WHEN coalesce(cast(c.producttype AS string), '') = 'BUSINESS_CARD'
                THEN 1 ELSE 0 END AS is_biz,
           -- pre-due: days 0..24 (run-up to the payment due date)
           CASE WHEN c.`date` >= a.stmt_dt
                 AND c.`date` <  date_add(a.stmt_dt, {STMT_DUE_DAY})
                THEN 1 ELSE 0 END AS pre_due_f,
           -- post-due: days 25..55 (missed due date, before roll to next stage)
           CASE WHEN c.`date` >= date_add(a.stmt_dt, {STMT_DUE_DAY})
                 AND c.`date` <  date_add(a.stmt_dt, {STMT_WINDOW_DAYS})
                THEN 1 ELSE 0 END AS post_due_f,
           -- overall: days 0..55 (the whole statement cycle window)
           CASE WHEN c.`date` >= a.stmt_dt
                 AND c.`date` <  date_add(a.stmt_dt, {STMT_WINDOW_DAYS})
                THEN 1 ELSE 0 END AS overall_f
    FROM {CALL} c
    JOIN stmt_anchor a
      ON try_cast(c.acctid AS bigint) = a.acct_num
    WHERE c.initiationmethod = 'INBOUND'
      AND c.acctid IS NOT NULL
      AND c.effdt >= '{EFFDT_SCAN_START}' AND c.effdt < '{EFFDT_HARD_END}'
),
dedup AS (
    -- first inbound call per (account, day); drop business cards
    SELECT *,
           row_number() OVER (PARTITION BY acct_key, call_dt
                              ORDER BY initiationtimestamp) AS rn
    FROM calls_flagged
    WHERE acct_key IS NOT NULL AND acct_key <> '' AND is_biz = 0
)
SELECT acct_key, acct_num, contactid, call_dt, call_month, stmt_dt,
       days_since_stmt_dt, pre_due_f, post_due_f, overall_f
FROM dedup
WHERE rn = 1
""")
print(f"[B_ishant/02_windows.py] built {DB}.uc2_ish_02n_calls")

print("[B_ishant/02_windows.py] inbound call-level window counts:")
display(spark.sql(f"""
SELECT count(1) AS inbound_calls,
       count(DISTINCT acct_key) AS inbound_accounts,
       count_if(pre_due_f = 1)  AS calls_pre_due,
       count_if(post_due_f = 1) AS calls_post_due,
       count_if(overall_f = 1)  AS calls_overall
FROM {DB}.uc2_ish_02n_calls
"""))

# COMMAND ----------

# ---------------------------------------------------------------------
# 02s. Roll the call-level flags to ACCOUNT level and join to the fixed
# ledger population. accts_called_* = max(window_flag) per account. The
# population is every 01s account; calls only classify.
# ---------------------------------------------------------------------
spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_ish_02s_pop AS
WITH call_acct AS (
    SELECT acct_key,
           count(1) AS inbound_call_cnt,
           count(DISTINCT contactid) AS inbound_contact_cnt,
           max(pre_due_f)  AS accts_called_pre_f,
           max(post_due_f) AS accts_called_post_f,
           max(overall_f)  AS accts_called_overall_f
    FROM {DB}.uc2_ish_02n_calls
    GROUP BY acct_key
)
SELECT p.*,
       coalesce(c.inbound_call_cnt, 0)        AS inbound_call_cnt,
       coalesce(c.inbound_contact_cnt, 0)     AS inbound_contact_cnt,
       coalesce(c.accts_called_pre_f, 0)      AS accts_called_pre_f,
       coalesce(c.accts_called_post_f, 0)     AS accts_called_post_f,
       coalesce(c.accts_called_overall_f, 0)  AS accts_called_overall_f
FROM {DB}.uc2_ish_01s_ledger p
LEFT JOIN call_acct c ON c.acct_key = p.acct_key
""")
print(f"[B_ishant/02_windows.py] built {DB}.uc2_ish_02s_pop")

# COMMAND ----------

# ---------------------------------------------------------------------
# The funnel-with-window pivot: 4 stages x 3 windows (account counts).
# This is the table Namit walked in the 21-Jul meeting. The ledger row's
# post-due cell is the headline 19,025.
# ---------------------------------------------------------------------
print("[B_ishant/02_windows.py] Stage funnel x window pivot (accounts called in each window):")
display(spark.sql(f"""
SELECT stage_order, stage,
       count(1)                            AS accts_in_stage,
       count_if(accts_called_pre_f = 1)    AS called_pre_due,
       count_if(accts_called_post_f = 1)   AS called_post_due,
       count_if(accts_called_overall_f = 1) AS called_overall
FROM (
    SELECT 1 AS stage_order, '01. Total accounts' AS stage,
           accts_called_pre_f, accts_called_post_f, accts_called_overall_f
    FROM {DB}.uc2_ish_02s_pop
    UNION ALL
    SELECT 2, '02. DQ-1 (DLNQT_CD_M1=1)',
           accts_called_pre_f, accts_called_post_f, accts_called_overall_f
    FROM {DB}.uc2_ish_02s_pop WHERE wf_dq1
    UNION ALL
    SELECT 3, '03. + CPC eligible',
           accts_called_pre_f, accts_called_post_f, accts_called_overall_f
    FROM {DB}.uc2_ish_02s_pop WHERE wf_dq1 AND wf_cpc
    UNION ALL
    SELECT 4, '04. + non-chargeoff = ledger',
           accts_called_pre_f, accts_called_post_f, accts_called_overall_f
    FROM {DB}.uc2_ish_02s_pop WHERE in_sas_ledger
) x
GROUP BY stage_order, stage
ORDER BY stage_order
"""))

_led = spark.sql(f"""
SELECT count_if(accts_called_pre_f = 1)     AS pre_due,
       count_if(accts_called_post_f = 1)    AS post_due,
       count_if(accts_called_overall_f = 1) AS overall
FROM {DB}.uc2_ish_02s_pop WHERE in_sas_ledger
""").first()
print("[B_ishant/02_windows.py] ledger-base window counts - actual vs expected (Ishant 202501):")
print(f"  pre-due accounts (day 0-24) : {_led['pre_due']:>8,}   expected ~6,778")
print(f"  post-due accounts (day 25-55): {_led['post_due']:>8,}   expected ~19,025  <- HEADLINE")
print(f"  overall in-window accounts  : {_led['overall']:>8,}   expected ~23,713")

# COMMAND ----------

# ---------------------------------------------------------------------
# [OPEN / VERIFY] The 25-vs-28-day edge (grace period). Anupam raised that the
# actionable window may start at day 28, not 25 (3-day grace between due date and
# next cycle date). Below shows the day-by-day call distribution around day 25 so
# the edge can be re-cut if the client settles on 28. NOT decided - do not hardcode.
# ---------------------------------------------------------------------
print("[B_ishant/02_windows.py] [OPEN: 25-vs-28 edge] daily inbound call counts, days 20-31 since stmt:")
display(spark.sql(f"""
SELECT days_since_stmt_dt, count(1) AS calls, count(DISTINCT acct_key) AS accts
FROM {DB}.uc2_ish_02n_calls
WHERE days_since_stmt_dt BETWEEN 20 AND 31
GROUP BY days_since_stmt_dt ORDER BY days_since_stmt_dt
"""))

print("[B_ishant/02_windows.py] 02_windows complete: uc2_ish_02n_calls, uc2_ish_02s_pop")
