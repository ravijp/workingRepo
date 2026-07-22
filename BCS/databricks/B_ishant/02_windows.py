# Databricks notebook source
# =====================================================================
# 02_windows.py
# The statement-window classification (single-max statement anchor).
#
# THE ANCHOR
#   stmt_anchor = max(stmt_last_dt) per account, taken over the statements
#   whose statement date falls in the statement period [STMT_START, STMT_END_EXCL).
#   That yields ONE governing statement per account: the cycle whose payment due
#   date lands in the anchor month. The population is fixed at all ledger
#   accounts; inbound calls only classify an account by which window they land in.
#
#   Bounding the anchor to the statement period is what keeps each call measured
#   against ITS cycle's statement. max(stmt_last_dt) taken over the whole fmt scan
#   window would return a later cycle's statement (dated after the calls), pushing
#   every call to a negative days-since-statement and out of all windows.
#
# THE STATEMENT / DUE-DATE WINDOWS (all measured in days from stmt_dt = day 0)
#   pre-due  = [stmt_dt,      stmt_dt + 25)  -> days  0..24  (run-up to due date)
#   post-due = [stmt_dt + 25, stmt_dt + 56)  -> days 25..55  (missed due, pre-roll)
#   overall  = [stmt_dt,      stmt_dt + 56)  -> days  0..55  (whole cycle)
#   Account-level flag = max(window_flag) over the account's inbound calls.
#   Before the due date a call routes to the Cares team; after it, to Collections,
#   so post-due is the actionable band.
#
#   Flag names (call grain / account grain):
#     call_25_window_f / call_31_window_f / call_overall_f
#     accts_called_25_f / accts_called_31_f / accts_called_overall_f
#   25 = pre-due (day 0-24, run-up to due date);
#   31 = post-due (day 25-55, the missed-due actionable band).
#
# WHAT THIS BUILDS
#   1. uc2_t16_02n_episodes - inbound calls, one row per contactid, with the
#                             25/31/overall window flags against the single-max
#                             anchor, plus call-context review columns.
#   2. uc2_t16_01s_populations_<vintage> - re-created to fold the account-level
#                             accts_called_25_f/_31_f/_overall_f onto the same
#                             table that carries the funnel/ECL/stage columns.
#   Then it prints the 4-stage x 3-window pivot (the funnel-with-windows table).
#
#   Every print/display below is prefixed with this file name for screenshots.
# =====================================================================

# COMMAND ----------

# ---------------------------------------------------------------------
# SETUP - catalog/schema, source-table handles, window + period constants.
# ---------------------------------------------------------------------
import datetime as _dt

CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
ANCHOR_YM = "202501"
SAS_CSV_PATH = "/Volumes/cda_model_shared/ecm_cld_model/ecm_cld/collections_zenon/WATERFALL_COLL_CALL_V3_202501.csv"
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

# Statement period = the billing cycles whose payment due date falls in the
# anchor month. A statement date on/after STMT_START and before STMT_END_EXCL
# qualifies. The single-max anchor is taken over statements inside this period.
STMT_START = "2024-12-07"       # statement dates on/after this qualify
STMT_END_EXCL = "2025-01-07"    # statement dates before this qualify
# Window day offsets from stmt_dt (day 0):
STMT_DUE_DAY = 25         # days: payment due-date marker; pre-due window ends here
STMT_WINDOW_DAYS = 56     # days: overall window is [stmt_dt, stmt_dt+56)
# Call-table effdt scan bounds (prev month .. anchor+3) - a scan-pruning guard.
EFFDT_SCAN_START = _mm(_a0, -1).isoformat()   # 2024-12-01
EFFDT_HARD_END = _mm(_a0, 3).isoformat()      # 2025-04-01

T_00N = f"{DB}.uc2_t16_00n_acct_monthly"
T_01S = f"{DB}.uc2_t16_01s_populations_{ANCHOR_YM}"
T_02N = f"{DB}.uc2_t16_02n_episodes"

print(f"[02_windows.py] SETUP OK: vintage {ANCHOR_YM}; layers -> {DB}")
print(f"[02_windows.py] windows: 25/pre-due [0,{STMT_DUE_DAY}) "
      f"31/post-due [{STMT_DUE_DAY},{STMT_WINDOW_DAYS}) overall [0,{STMT_WINDOW_DAYS}); "
      f"stmt period {STMT_START}..{STMT_END_EXCL}")
# ---------------------------------------------------------------------
# end of SETUP
# ---------------------------------------------------------------------

# COMMAND ----------

# REFRESH the call table so the scan sees the live loading edge.
spark.sql(f"REFRESH TABLE {CALL}")
print("[02_windows.py] call table refreshed")

# COMMAND ----------

# ---------------------------------------------------------------------
# 02n. Inbound calls classified against the single-max statement anchor.
#
#   stmt_anchor = max(stmt_last_dt) per account over the statement period =
#   one governing statement per account. call_25/31/overall_window_f are
#   computed at call grain using date_add(stmt_dt, 25) and date_add(stmt_dt, 56).
#   One row per contactid; first-inbound-per-day dedup, business cards dropped.
#
#   Review columns (real call-table columns per data-dictionary lines 323-329):
#     producttype   - product line
#     queue         - collection queue / routing bucket (e.g. CARE_GAP)
#     routingprofile- vendor+function+program (e.g. TEL_CARE_PP_01)
#     department    - org work classification (Care / Fraud / Collection)
#     segment       - org/work classification
# ---------------------------------------------------------------------
spark.sql(f"""
CREATE OR REPLACE TABLE {T_02N} AS
WITH stmt_anchor_raw AS (
    -- distinct account statement dates that fall in the statement period
    SELECT {NUM_KEY.format(c="extnl_acct_id")} AS acct_key,
           try_cast(extnl_acct_id AS bigint) AS acct_num,
           try_cast(stmt_last_dt AS date) AS stmt_dt
    FROM {FMT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '{MONTH_WIN_START}' AND eff_dt < '{MONTH_WIN_END}'
      AND stmt_last_dt IS NOT NULL
      AND try_cast(stmt_last_dt AS date) >= DATE '{STMT_START}'
      AND try_cast(stmt_last_dt AS date) <  DATE '{STMT_END_EXCL}'
),
stmt_anchor AS (
    -- THE ANCHOR: max(stmt_last_dt) over the statement period = one governing
    -- statement per account (the cycle due in the anchor month).
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
           -- due date derived = stmt_dt + 25; no explicit due-date column in fmt
           date_add(a.stmt_dt, {STMT_DUE_DAY}) AS due_dt_derived,
           -- call-context review columns (real call-table columns, dict lines 323-329)
           cast(c.producttype AS string)    AS producttype,
           cast(c.queue AS string)          AS queue,
           cast(c.routingprofile AS string) AS routingprofile,
           cast(c.department AS string)     AS department,
           cast(c.segment AS string)        AS segment,
           CASE WHEN coalesce(cast(c.producttype AS string), '') = 'BUSINESS_CARD'
                THEN 1 ELSE 0 END AS is_biz,
           -- 25/pre-due: days 0..24 (run-up to the payment due date)
           CASE WHEN c.`date` >= a.stmt_dt
                 AND c.`date` <  date_add(a.stmt_dt, {STMT_DUE_DAY})
                THEN 1 ELSE 0 END AS call_25_window_f,
           -- 31/post-due: days 25..55 (missed due date, before roll to next stage)
           CASE WHEN c.`date` >= date_add(a.stmt_dt, {STMT_DUE_DAY})
                 AND c.`date` <  date_add(a.stmt_dt, {STMT_WINDOW_DAYS})
                THEN 1 ELSE 0 END AS call_31_window_f,
           -- overall: days 0..55 (the whole statement cycle window)
           CASE WHEN c.`date` >= a.stmt_dt
                 AND c.`date` <  date_add(a.stmt_dt, {STMT_WINDOW_DAYS})
                THEN 1 ELSE 0 END AS call_overall_f,
           -- within_effdt_cap: effdt inside the statement period (carried diagnostic)
           CASE WHEN c.effdt >= '{STMT_START}' AND c.effdt < '{STMT_END_EXCL}'
                THEN 1 ELSE 0 END AS within_effdt_cap
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
       days_since_stmt_dt, due_dt_derived,
       producttype, queue, routingprofile, department, segment,
       is_biz, within_effdt_cap,
       call_25_window_f, call_31_window_f, call_overall_f
FROM dedup
WHERE rn = 1
""")
print(f"[02_windows.py] built {T_02N}")

print("[02_windows.py] inbound call-level window counts:")
display(spark.sql(f"""
SELECT count(1) AS inbound_calls,
       count(DISTINCT acct_key) AS inbound_accounts,
       count_if(call_25_window_f = 1)      AS calls_pre_due_25,
       count_if(call_31_window_f = 1)      AS calls_post_due_31,
       count_if(call_overall_f = 1)        AS calls_overall
FROM {T_02N}
"""))

# COMMAND ----------

# ---------------------------------------------------------------------
# FOLD: roll the call-level flags to account level and re-create the ledger
# table with the accts_called_* flags on the same table as the funnel columns.
# accts_called_* = max(window_flag) per account. The population is every 01s
# account; calls only classify.
#
# TWO GRAINS carried for the post-due(31) count:
#   *_episodes = counted over uc2_t16_02n_episodes (the DEDUPED grain, first
#                inbound per (acct_key, call_dt)); this is our episode definition.
#   *_calls    = counted over ALL inbound in-window calls (PRE-dedup, is_biz=0),
#                no per-day dedup. Re-derived here in all_calls straight from the
#                call table with the same anchor/window logic as 02n's
#                calls_flagged, so nothing is dropped and the dedup above is
#                untouched.
#   *_ind      = account-level any-post-due-call indicator = max(call_31_window_f)
#                (equals accts_called_31_f; aliased, both kept).
# ---------------------------------------------------------------------
spark.sql(f"""
CREATE OR REPLACE TABLE {T_01S} AS
WITH call_acct AS (
    -- DEDUPED episode grain (from uc2_t16_02n_episodes, one row per acct/day)
    SELECT acct_key,
           count(1) AS inbound_call_cnt_stmt_window,
           count(DISTINCT contactid) AS inbound_contact_cnt_stmt_window,
           count_if(call_31_window_f = 1) AS inbound_call_stmnt_dt_25_plus_31_episodes,
           max(call_25_window_f)  AS accts_called_25_f,       -- 25 = pre-due (day 0-24)
           max(call_31_window_f)  AS accts_called_31_f,       -- 31 = post-due (day 25-55)
           max(call_overall_f)    AS accts_called_overall_f,
           max(call_31_window_f)  AS inbound_call_stmnt_dt_25_plus_31_ind
    FROM {T_02N}
    GROUP BY acct_key
),
stmt_anchor_raw AS (
    -- same anchor derivation as the 02n build (single-max statement anchor)
    SELECT {NUM_KEY.format(c="extnl_acct_id")} AS acct_key,
           try_cast(extnl_acct_id AS bigint) AS acct_num,
           try_cast(stmt_last_dt AS date) AS stmt_dt
    FROM {FMT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '{MONTH_WIN_START}' AND eff_dt < '{MONTH_WIN_END}'
      AND stmt_last_dt IS NOT NULL
      AND try_cast(stmt_last_dt AS date) >= DATE '{STMT_START}'
      AND try_cast(stmt_last_dt AS date) <  DATE '{STMT_END_EXCL}'
),
stmt_anchor AS (
    SELECT acct_key, acct_num, max(stmt_dt) AS stmt_dt
    FROM stmt_anchor_raw
    GROUP BY acct_key, acct_num
),
all_calls AS (
    -- ALL inbound in-window calls, PRE-dedup (no per-day dedup), is_biz = 0.
    -- Same anchor/window/scan logic as 02n's calls_flagged; used only for the
    -- all-calls post-due count. The 02n dedup above is unchanged.
    SELECT a.acct_key,
           count_if(c.`date` >= date_add(a.stmt_dt, {STMT_DUE_DAY})
                    AND c.`date` < date_add(a.stmt_dt, {STMT_WINDOW_DAYS}))
             AS inbound_call_stmnt_dt_25_plus_31_calls
    FROM {CALL} c
    JOIN stmt_anchor a
      ON try_cast(c.acctid AS bigint) = a.acct_num
    WHERE c.initiationmethod = 'INBOUND'
      AND c.acctid IS NOT NULL
      AND c.effdt >= '{EFFDT_SCAN_START}' AND c.effdt < '{EFFDT_HARD_END}'
      AND coalesce(cast(c.producttype AS string), '') <> 'BUSINESS_CARD'
      AND a.acct_key IS NOT NULL AND a.acct_key <> ''
    GROUP BY a.acct_key
)
SELECT p.*,
       coalesce(c.inbound_call_cnt_stmt_window, 0)     AS inbound_call_cnt_stmt_window,
       coalesce(c.inbound_contact_cnt_stmt_window, 0)  AS inbound_contact_cnt_stmt_window,
       coalesce(c.accts_called_25_f, 0)                AS accts_called_25_f,
       coalesce(c.accts_called_31_f, 0)                AS accts_called_31_f,
       coalesce(c.accts_called_overall_f, 0)           AS accts_called_overall_f,
       -- post-due(31) counts, two grains + indicator
       coalesce(ac.inbound_call_stmnt_dt_25_plus_31_calls, 0)
         AS inbound_call_stmnt_dt_25_plus_31_calls,     -- ALL calls, no per-day dedup
       coalesce(c.inbound_call_stmnt_dt_25_plus_31_episodes, 0)
         AS inbound_call_stmnt_dt_25_plus_31_episodes,  -- DEDUPED episodes (our grain)
       coalesce(c.inbound_call_stmnt_dt_25_plus_31_ind, 0)
         AS inbound_call_stmnt_dt_25_plus_31_ind        -- any-post-due-call indicator
FROM {T_01S} p
LEFT JOIN call_acct c  ON c.acct_key = p.acct_key
LEFT JOIN all_calls ac ON ac.acct_key = p.acct_key
""")
print(f"[02_windows.py] folded window flags into {T_01S} (funnel + accts_called_* on one table)")
print(f"[02_windows.py] post-due(31) counts carried at two grains: "
      "_calls (all-calls, no dedup) and _episodes (deduped), plus _ind (indicator)")

_pd = spark.sql(f"""
SELECT sum(inbound_call_stmnt_dt_25_plus_31_calls)    AS post_due_31_calls_all,
       sum(inbound_call_stmnt_dt_25_plus_31_episodes) AS post_due_31_episodes,
       sum(inbound_call_stmnt_dt_25_plus_31_ind)      AS post_due_31_accts_ind
FROM {T_01S} WHERE in_sas_ledger
""").first()
print("[02_windows.py] ledger post-due(31) totals:")
print(f"  all-calls (no dedup)    : {_pd['post_due_31_calls_all']:>10,}")
print(f"  deduped episodes        : {_pd['post_due_31_episodes']:>10,}")
print(f"  accounts (indicator)    : {_pd['post_due_31_accts_ind']:>10,}")

# COMMAND ----------

# ---------------------------------------------------------------------
# The funnel-with-window pivot: 4 stages x 3 windows (account counts).
# The ledger row's post-due(31) cell is the actionable-band headline.
# ---------------------------------------------------------------------
print("[02_windows.py] Stage funnel x window pivot (accounts called in each window):")
display(spark.sql(f"""
SELECT stage_order, stage,
       count(1)                              AS accts_in_stage,
       count_if(accts_called_25_f = 1)       AS called_pre_due_25,
       count_if(accts_called_31_f = 1)       AS called_post_due_31,
       count_if(accts_called_overall_f = 1)  AS called_overall
FROM (
    SELECT 1 AS stage_order, '01. Total accounts' AS stage,
           accts_called_25_f, accts_called_31_f, accts_called_overall_f
    FROM {T_01S}
    UNION ALL
    SELECT 2, '02. DQ-1 (DLNQT_CD_M1=1)',
           accts_called_25_f, accts_called_31_f, accts_called_overall_f
    FROM {T_01S} WHERE wf_dq1
    UNION ALL
    SELECT 3, '03. + CPC eligible',
           accts_called_25_f, accts_called_31_f, accts_called_overall_f
    FROM {T_01S} WHERE wf_dq1 AND wf_cpc
    UNION ALL
    SELECT 4, '04. + non-chargeoff = ledger',
           accts_called_25_f, accts_called_31_f, accts_called_overall_f
    FROM {T_01S} WHERE in_sas_ledger
) x
GROUP BY stage_order, stage
ORDER BY stage_order
"""))

_led = spark.sql(f"""
SELECT count_if(accts_called_25_f = 1)      AS pre_due,
       count_if(accts_called_31_f = 1)      AS post_due,
       count_if(accts_called_overall_f = 1) AS overall
FROM {T_01S} WHERE in_sas_ledger
""").first()
print("[02_windows.py] ledger-base window counts:")
print(f"  pre-due accounts (25 / day 0-24)  : {_led['pre_due']:>8,}")
print(f"  post-due accounts (31 / day 25-55): {_led['post_due']:>8,}")
print(f"  overall in-window accounts        : {_led['overall']:>8,}")

# COMMAND ----------

# ---------------------------------------------------------------------
# [OPEN: 25-vs-28 day edge] The actionable window may start at day 28, not 25,
# if the client applies a 3-day grace period between the due date and the next
# cycle date. Below shows the day-by-day call distribution around day 25 so the
# edge can be re-cut if the client settles on 28. Not decided; not hardcoded.
# ---------------------------------------------------------------------
print("[02_windows.py] [OPEN: 25-vs-28 edge] daily inbound call counts, days 20-31 since stmt:")
display(spark.sql(f"""
SELECT days_since_stmt_dt, count(1) AS calls, count(DISTINCT acct_key) AS accts
FROM {T_02N}
WHERE days_since_stmt_dt BETWEEN 20 AND 31
GROUP BY days_since_stmt_dt ORDER BY days_since_stmt_dt
"""))

print(f"[02_windows.py] 02_windows complete: uc2_t16_02n_episodes, "
      f"uc2_t16_01s_populations_{ANCHOR_YM} (folded)")
