# Databricks notebook source
# MAGIC %md
# MAGIC # B02. The key fix: rebuild the AWS layers on the numeric key (runs FIRST)
# MAGIC
# MAGIC WHAT CHANGES: one line - acct_key moves from `trim(cast(id AS string))` to
# MAGIC `cast(try_cast(id AS bigint) AS string)` in every layer. Everything else is
# MAGIC the verified tier-16 logic verbatim.
# MAGIC OUTPUT TABLES (new names; round-10 uc2_t16_00..04 stay frozen):
# MAGIC uc2_t16_00n_acct_monthly, uc2_t16_01n_populations, uc2_t16_02n_episodes,
# MAGIC uc2_t16_03n_signals, uc2_t16_04n_outcomes.
# MAGIC DISCLOSED DEVIATION (D8): the 02n call scan adds an effdt bound
# MAGIC [EFFDT_SCAN_START, 2026-07-10) plus REFRESH TABLE. Standard episodes need
# MAGIC the stricter in-column cap, so no anchor moves.
# MAGIC STORY-B RE-ANCHOR (2026-07-21, fixed same day): the 02n episode build now
# MAGIC anchors each call to ITS OWN cycle's statement. A stmt_dates CTE reads ALL
# MAGIC distinct statement dates per account from fmt (bounded, sfx_nbr=0); each
# MAGIC call as-of joins the most-recent statement ON OR BEFORE it, then computes
# MAGIC days_since_stmt_dt / stmt_5day_bucket / pre_due_f / post_due_f, and keeps
# MAGIC ONLY episodes whose call_dt falls in [stmt_dt, stmt_dt + 56d). Call-days
# MAGIC outside any statement window DROP from the episode population - this is
# MAGIC what makes the caller/episode counts MOVE off the January values (by
# MAGIC design). The numeric key, is_biz=0 / within-scan filters, the 03n regexes,
# MAGIC the ex-AA gate and the aws diagnostics are byte-identical; 00n/01n
# MAGIC population logic is unchanged (account grain, frame-independent).
# MAGIC Run B02_checks.py once after this to certify all population/caller anchors.

# COMMAND ----------

# =====================================================================
# SETUP - keep in sync across B00/B01/B02/B02b/B03 (B00 is the canonical copy).
# =====================================================================
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

PRV_YM = _mm(_a0, -1).strftime("%Y%m")
FEB_YM = _mm(_a0, 1).strftime("%Y%m")
MAR_YM = _mm(_a0, 2).strftime("%Y%m")
MONTH_WIN_START = _mm(_a0, -1).strftime("%Y%m%d")
MONTH_WIN_END = _mm(_a0, 3).strftime("%Y%m%d")
CALL_WIN_START = _a0.isoformat()
CALL_WIN_END = _mm(_a0, 1).isoformat()
EFFDT_CAP_START = _a0.isoformat()
EFFDT_CAP_END = (_mm(_a0, 1) + _dt.timedelta(days=1)).isoformat()
CLEANUP_DATE = _a0.isoformat()
ANCHOR_EOM = (_mm(_a0, 1) - _dt.timedelta(days=1)).isoformat()
FEB_START = _mm(_a0, 1).isoformat()
MAR_START = _mm(_a0, 2).isoformat()
APR_START = _mm(_a0, 3).isoformat()
CO8_END = (_mm(_a0, 9) - _dt.timedelta(days=1)).isoformat()
CO10_END = (_mm(_a0, 11) - _dt.timedelta(days=1)).isoformat()
CO12_END = (_mm(_a0, 13) - _dt.timedelta(days=1)).isoformat()
FWD_CO_START = _a0.strftime("%Y%m%d")
FWD_CO_END = _mm(_a0, 12).strftime("%Y%m%d")
SNAP_DAILY_START = _mm(_a0, -7).strftime("%Y%m%d")
SNAP_DAILY_END = _mm(_a0, 1).strftime("%Y%m%d")
EFFDT_SCAN_START = _mm(_a0, -1).isoformat()
EFFDT_HARD_END = "2026-07-10"   # not vintage-derived: the live-loading-edge guard

NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"

# --- statement-cycle re-anchor constants (Story B, 2026-07-21) ------------
# The inbound analysis is re-anchored from calendar January to each account's
# STATEMENT CYCLE. stmt_dt = statement date (day 0, the bill lands), sourced
# from fmt_acct_c.stmt_last_dt. Day 0..~25 = PRE-DUE run-up to the payment due
# date (due day ~= 25). Day 25..~56 = POST-DUE, from the missed due date until
# the NEXT statement lands ~31 days later. days_since_stmt_dt = datediff(call_dt,
# stmt_dt); 5-day buckets over 0..55. An episode is kept only if call_dt is in
# [stmt_dt, stmt_dt + STMT_WINDOW_DAYS).
STMT_DUE_DAY = 25            # unit: days since stmt_dt; the payment due date marker
STMT_WINDOW_DAYS = 56        # unit: days; the [stmt_dt, stmt_dt+56) episode window
STMT_BUCKET_WIDTH = 5        # unit: days; 5-day bucket width over 0..55

print(f"SETUP OK: vintage {ANCHOR_YM}; layers -> {DB}")
# =====================================================================
# end of SETUP
# =====================================================================

# COMMAND ----------

# K1. preconditions: sources + round-10 tables + notebook A's recon tables
for _t in [f"{FMT_CATALOG}.fmt_acct_dba.fmt_acct_c",
           f"{CC_CATALOG}.contactcenter_bdp_db.call",
           f"{CC_CATALOG}.contactcenter_bdp_db.transcript",
           f"{DB}.uc2_t16_01_populations", f"{DB}.uc2_t16_02_episodes"]:
    assert spark.catalog.tableExists(_t), f"PRECONDITION MISS: {_t} not reachable"
if ANCHOR_YM == "202501":
    for _t in ["uc2_gap1942_202501", "uc2_sasflag_202501"]:
        assert spark.catalog.tableExists(f"{DB}.{_t}"), \
            f"PRECONDITION MISS: {DB}.{_t} missing - run A_recon_lock_202501 first"

# COMMAND ----------

# MAGIC %md
# MAGIC ## K3. Build `uc2_t16_00n_acct_monthly` (the expensive scan; numeric key)

# COMMAND ----------

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_00n_acct_monthly AS
WITH snap AS (
    SELECT extnl_acct_id,
           substr(eff_dt, 1, 6) AS ym,
           eff_dt,
           CASE
             WHEN past_due_271_up_amt  > 0 THEN 10
             WHEN past_due_241_270_amt > 0 THEN 9
             WHEN past_due_211_240_amt > 0 THEN 8
             WHEN past_due_181_210_amt > 0 THEN 7
             WHEN past_due_151_180_amt > 0 THEN 6
             WHEN past_due_121_150_amt > 0 THEN 5
             WHEN past_due_91_120_amt  > 0 THEN 4
             WHEN past_due_61_90_amt   > 0 THEN 3
             WHEN past_due_31_60_amt   > 0 THEN 2
             WHEN past_due_1_30_amt    > 0 THEN 1
             ELSE 0
           END AS bucket,
           try_cast(acct_bal_amt AS double) AS bal,
           try_cast(chrgoff_dt AS date) AS co_dt,
           try_cast(chrgoff_amt AS double) AS co_amt,
           clnt_prdct_cd,
           try_cast(cr_lmt_origl_amt AS double) AS cr_lmt_origl_amt,
           coalesce(try_cast(paymt_last_dt AS date),
                    try_to_date(cast(paymt_last_dt AS string), 'ddMMMyyyy')) AS pay_dt,
           coalesce(try_cast(atmtc_paymt_last_dt AS date),
                    try_to_date(cast(atmtc_paymt_last_dt AS string), 'ddMMMyyyy')) AS auto_dt,
           coalesce(try_cast(nsf_last_paymt_dt AS date),
                    try_to_date(cast(nsf_last_paymt_dt AS string), 'ddMMMyyyy')) AS nsf_dt
    FROM {FMT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '{MONTH_WIN_START}' AND eff_dt < '{MONTH_WIN_END}'
)
SELECT {NUM_KEY.format(c="extnl_acct_id")} AS acct_key,   -- THE KEY CHANGE (D2)
       ym,
       max(bucket) AS max_bucket,
       max_by(bucket, eff_dt) AS eom_bucket,
       max_by(bal, eff_dt) AS eom_bal,
       min(co_dt) AS mth_co_dt,
       min_by(co_amt, co_dt) AS mth_co_amt,
       min(CASE WHEN bucket >= 1 THEN eff_dt END) AS first_dq_dt,
       min(CASE WHEN bucket = 1 THEN eff_dt END) AS first_b1_dt,
       max_by(clnt_prdct_cd, eff_dt) AS eom_cpc,
       max_by(cr_lmt_origl_amt, eff_dt) AS eom_cr_lmt_origl_amt,
       max(pay_dt) AS pay_dt,
       max(auto_dt) AS auto_dt,
       max(nsf_dt) AS nsf_dt
FROM snap
GROUP BY 1, 2
""")
print(f"built {DB}.uc2_t16_00n_acct_monthly")

# COMMAND ----------

# MAGIC %md
# MAGIC ## K4. Build `uc2_t16_01n_populations` (tier-16 layer 01 verbatim on 00n)

# COMMAND ----------

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_01n_populations AS
WITH jan AS (SELECT * FROM {DB}.uc2_t16_00n_acct_monthly WHERE ym = '{ANCHOR_YM}'),
prv AS (SELECT * FROM {DB}.uc2_t16_00n_acct_monthly WHERE ym = '{PRV_YM}'),
feb AS (SELECT * FROM {DB}.uc2_t16_00n_acct_monthly WHERE ym = '{FEB_YM}'),
mar AS (SELECT * FROM {DB}.uc2_t16_00n_acct_monthly WHERE ym = '{MAR_YM}'),
future_co AS (
    SELECT {NUM_KEY.format(c="extnl_acct_id")} AS acct_key,
           min(try_cast(chrgoff_dt AS date)) AS co_dt_future,
           min_by(try_cast(chrgoff_amt AS double), try_cast(chrgoff_dt AS date)) AS co_amt
    FROM {FMT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '{FWD_CO_START}' AND eff_dt < '{FWD_CO_END}'
      AND chrgoff_dt IS NOT NULL
      AND {NUM_KEY.format(c="extnl_acct_id")} IN (SELECT acct_key FROM jan WHERE max_bucket >= 1)
    GROUP BY 1
),
pop_base AS (
    SELECT j.acct_key,
           j.max_bucket, j.eom_bucket, j.eom_bal, j.mth_co_dt AS jan_co_dt,
           j.first_dq_dt, j.first_b1_dt, j.eom_cpc, j.eom_cr_lmt_origl_amt,
           p.max_bucket AS prev_max_bucket,
           p.eom_bucket AS prev_eom_bucket,
           (j.mth_co_dt IS NULL OR j.mth_co_dt >= DATE '{CLEANUP_DATE}') AS cleaned,
           (j.eom_cpc IS NULL OR trim(j.eom_cpc) = ''
            OR j.eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                                  'AA3','AC3','AM3','AA4','AC4','AM4',
                                  'BGC','BGM','CGM','GMR',
                                  'FBS','IBS','U1C','U2C','U3C')) AS is_exaa,
           CASE
             WHEN j.eom_cpc IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                                 'AA3','AC3','AM3','AA4','AC4','AM4')     THEN 'AA'
             WHEN j.eom_cpc IN ('BGC','BGM','CGM','GMR')                 THEN 'GM'
             WHEN j.eom_cpc IN ('FBS','IBS','U1C','U2C','U3C')           THEN 'Bronco'
             WHEN j.eom_cpc IN ('BHA','BJT','BJC','BFR','BWY','BBB')     THEN 'Biz'
             WHEN j.eom_cpc IN ('GAP','GP2','ONV','ON2','BRP','BR2','ATH','AT2',
                                 'GPC','G2C','ONC','O2C','BRC','B2C','ATC','A2C')
                                                                          THEN 'CoBrand'
             WHEN j.eom_cpc IN ('8GP','8ON','8BR','8AT','9GP','9ON','9BR','9AT')
                                                                          THEN 'PLCC'
             ELSE 'OTHER'
           END AS cpc_class,
           CASE
             WHEN coalesce(p.eom_bucket, 0) >= 1
               THEN 'd. carried-in (past due at Dec-31 EOM)'
             WHEN j.first_dq_dt IS NULL THEN NULL
             WHEN cast(substr(j.first_dq_dt, 7, 2) AS int) <= 10
               THEN 'a. runway >= 21 days (entry day 1-10)'
             WHEN cast(substr(j.first_dq_dt, 7, 2) AS int) <= 20
               THEN 'b. runway 11-20 days (entry day 11-20)'
             ELSE 'c. runway <= 10 days (entry day 21-31)'
           END AS runway_band,
           CASE
             WHEN f.mth_co_dt >= DATE '{FEB_START}' AND f.mth_co_dt < DATE '{MAR_START}'
               THEN 'e. charged off in Feb'
             WHEN f.acct_key IS NULL THEN 'f. no Feb row'
             WHEN f.eom_bucket = 0 THEN 'a. Feb EOM bucket 0 (cured)'
             WHEN f.eom_bucket = 1 THEN 'b. Feb EOM bucket 1 (stayed)'
             WHEN f.eom_bucket = 2 THEN 'c. Feb EOM bucket 2 (rolled)'
             ELSE 'd. Feb EOM bucket 3+ (rolled deeper)'
           END AS feb_position_b14,
           CASE
             WHEN f.mth_co_dt >= DATE '{FEB_START}' AND f.mth_co_dt < DATE '{MAR_START}' THEN 'co'
             WHEN f.acct_key IS NULL THEN 'gone'
             ELSE cast(f.eom_bucket AS string)
           END AS feb_pos,
           CASE
             WHEN m.mth_co_dt >= DATE '{FEB_START}' AND m.mth_co_dt < DATE '{APR_START}' THEN 'co'
             WHEN m.acct_key IS NULL THEN 'gone'
             ELSE cast(m.eom_bucket AS string)
           END AS mar_pos,
           fc.co_dt_future,
           fc.co_amt,
           (fc.co_dt_future >= DATE '{ANCHOR_EOM}' AND fc.co_dt_future < DATE '{CO8_END}')  AS co_8m,
           (fc.co_dt_future >= DATE '{ANCHOR_EOM}' AND fc.co_dt_future < DATE '{CO10_END}') AS co_10m,
           (fc.co_dt_future >= DATE '{ANCHOR_EOM}' AND fc.co_dt_future < DATE '{CO12_END}') AS co_12m
    FROM jan j
    LEFT JOIN prv p ON p.acct_key = j.acct_key
    LEFT JOIN feb f ON f.acct_key = j.acct_key
    LEFT JOIN mar m ON m.acct_key = j.acct_key
    LEFT JOIN future_co fc ON fc.acct_key = j.acct_key
)
SELECT *,
       (eom_bucket = 1 AND cleaned)             AS in_ledger_all,
       (eom_bucket = 1 AND cleaned AND is_exaa) AS in_ledger_exaa,
       (first_b1_dt IS NOT NULL AND cleaned AND is_exaa) AS touched_b1,
       CASE
         WHEN NOT (first_b1_dt IS NOT NULL AND cleaned AND is_exaa) THEN NULL
         WHEN jan_co_dt >= DATE '{CLEANUP_DATE}' AND jan_co_dt < DATE '{FEB_START}'
           THEN 'd. charged off in January'
         WHEN eom_bucket = 0
           THEN 'a. current at 31 Jan (cured in month)'
         WHEN eom_bucket = 1
           THEN 'b. bucket 1 at 31 Jan'
         WHEN eom_bucket >= 2
           THEN 'c. bucket 2+ at 31 Jan (rolled past DQ1 within January)'
       END AS touched_b1_class
FROM pop_base
""")
print(f"built {DB}.uc2_t16_01n_populations")

# COMMAND ----------

# MAGIC %md
# MAGIC ## K5. REFRESH the call table (live loading edge) before the 02n build

# COMMAND ----------

spark.sql(f"REFRESH TABLE {CALL}")

# COMMAND ----------

# MAGIC %md
# MAGIC ## K6. Build `uc2_t16_02n_episodes` (numeric key; had_zero_pad diagnostic; D8 bounded scan; STORY-B statement re-anchor)
# MAGIC
# MAGIC RE-ANCHOR: a stmt_dates CTE reads ALL distinct statement dates per account
# MAGIC (over the bounded fmt window, sfx_nbr=0); each call as-of joins the
# MAGIC most-recent statement on or before it (QUALIFY row_number ... = 1).
# MAGIC days_since_stmt_dt = datediff(call_dt, stmt_dt); the 5-day
# MAGIC bucket floor(days/5)*5 is labelled with due-date meaning (pre-due 00-24,
# MAGIC post-due 25-55, due day = 25), with an "outside 0-55" sentinel. An episode
# MAGIC is standard (is_episode_std = 1) ONLY if it is first-inbound-per-day AND
# MAGIC in_stmt_window = 1 (call_dt in [stmt_dt, stmt_dt+56)). Call-days outside
# MAGIC any statement window DROP from the episode population - the intended move.

# COMMAND ----------

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_02n_episodes AS
stmt_dates AS (
    -- ALL distinct statement dates per account over the bounded fmt window
    -- (sfx_nbr=0), one per billing cycle. Source is fmt_acct_c.stmt_last_dt
    -- (NOT a SAS csv column). Grain: one row per (acct_key, statement date).
    -- FIX 2026-07-21: the prior stmt_anchor used max(stmt_last_dt) = ONE date
    -- per account, which for a January call was almost always a Feb/Mar
    -- statement (AFTER the call), so datediff < 0 and in_stmt_window = 0 for
    -- nearly everyone - only ~25 coincidental survivors reached 04s and
    -- captured_sas went to 0. Each call must be measured against ITS OWN
    -- cycle's statement (the most recent statement on or before the call),
    -- not the account's latest statement. See WINDOW_COLLAPSE_DIAGNOSIS.md.
    SELECT {NUM_KEY.format(c="extnl_acct_id")} AS acct_key,
           try_cast(stmt_last_dt AS date) AS stmt_dt
    FROM {FMT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '{MONTH_WIN_START}' AND eff_dt < '{MONTH_WIN_END}'
      AND stmt_last_dt IS NOT NULL
    GROUP BY 1, 2
),
calls_flagged AS (
    SELECT {NUM_KEY.format(c="acctid")} AS acct_key,      -- THE KEY CHANGE (D2)
           contactid,
           `date` AS call_dt,
           cast(date_trunc('month', `date`) AS date) AS call_month,
           initiationtimestamp,
           CASE WHEN coalesce(cast(producttype AS string), '') = 'BUSINESS_CARD'
                THEN 1 ELSE 0 END AS is_biz,
           CASE WHEN effdt >= '{EFFDT_CAP_START}' AND effdt < '{EFFDT_CAP_END}'
                THEN 1 ELSE 0 END AS within_effdt_cap,
           CASE WHEN try_cast(acctid AS bigint) IS NOT NULL
                 AND trim(cast(acctid AS string)) <> {NUM_KEY.format(c="acctid")}
                THEN 1 ELSE 0 END AS had_zero_pad          -- diagnostic: the rows the string key lost
    FROM {CALL}
    WHERE initiationmethod = 'INBOUND'
      AND `date` >= DATE '{CALL_WIN_START}' AND `date` < DATE '{CALL_WIN_END}'
      AND acctid IS NOT NULL
      AND effdt >= '{EFFDT_SCAN_START}' AND effdt < '{EFFDT_HARD_END}'   -- D8 bounded scan
),
calls_anchored AS (
    -- attach the statement anchor and derive the statement-cycle attributes.
    -- days_since_stmt_dt unit = days; in_stmt_window is the re-anchor keep flag.
    SELECT c.*,
           a.stmt_dt,
           datediff(c.call_dt, a.stmt_dt) AS days_since_stmt_dt,
           CASE WHEN a.stmt_dt IS NOT NULL
                 AND datediff(c.call_dt, a.stmt_dt) >= 0
                 AND datediff(c.call_dt, a.stmt_dt) < {STMT_WINDOW_DAYS}
                THEN 1 ELSE 0 END AS in_stmt_window,        -- re-anchor keep flag: call in [stmt_dt, stmt_dt+56)
           CASE WHEN a.stmt_dt IS NOT NULL
                 AND datediff(c.call_dt, a.stmt_dt) >= 0
                 AND datediff(c.call_dt, a.stmt_dt) < {STMT_DUE_DAY}
                THEN 1 ELSE 0 END AS pre_due_f,             -- days 0-24: run-up to the payment due date
           CASE WHEN a.stmt_dt IS NOT NULL
                 AND datediff(c.call_dt, a.stmt_dt) >= {STMT_DUE_DAY}
                 AND datediff(c.call_dt, a.stmt_dt) < {STMT_WINDOW_DAYS}
                THEN 1 ELSE 0 END AS post_due_f,            -- days 25-55: past-due, until the next statement lands
           CASE WHEN a.stmt_dt IS NULL
                 OR datediff(c.call_dt, a.stmt_dt) < 0
                 OR datediff(c.call_dt, a.stmt_dt) >= {STMT_WINDOW_DAYS}
                THEN 'outside 0-55 days'
                ELSE lpad(cast(floor(datediff(c.call_dt, a.stmt_dt) / {STMT_BUCKET_WIDTH})
                               * {STMT_BUCKET_WIDTH} AS string), 2, '0')
                     || '-'
                     || lpad(cast(floor(datediff(c.call_dt, a.stmt_dt) / {STMT_BUCKET_WIDTH})
                               * {STMT_BUCKET_WIDTH} + {STMT_BUCKET_WIDTH} - 1 AS string), 2, '0')
                     || CASE WHEN datediff(c.call_dt, a.stmt_dt) < {STMT_DUE_DAY}
                             THEN ' pre-due' ELSE ' post-due' END
           END AS stmt_5day_bucket,                          -- e.g. '00-04 pre-due' .. '50-54 post-due'; due_day=25
           CASE WHEN a.stmt_dt IS NULL
                 OR datediff(c.call_dt, a.stmt_dt) < 0
                 OR datediff(c.call_dt, a.stmt_dt) >= {STMT_WINDOW_DAYS}
                THEN NULL
                ELSE cast(floor(datediff(c.call_dt, a.stmt_dt) / {STMT_BUCKET_WIDTH})
                          * {STMT_BUCKET_WIDTH} AS int)
           END AS stmt_5day_bucket_start                     -- unit: days; the bucket's low edge, for ordering
    FROM calls_flagged c
    -- AS-OF join: attach the most-recent statement ON OR BEFORE the call, so
    -- each call is measured against its OWN cycle's statement (FIX 2026-07-21).
    -- A call with no prior statement gets stmt_dt = NULL -> in_stmt_window = 0
    -- (correctly excluded). One statement per contactid via the QUALIFY pick.
    LEFT JOIN stmt_dates a
      ON a.acct_key = c.acct_key
     AND a.stmt_dt <= c.call_dt
    QUALIFY row_number() OVER (PARTITION BY c.contactid
                               ORDER BY a.stmt_dt DESC) = 1
),
episodes_std AS (
    -- first-inbound-per-day survivor, RE-ANCHORED to the statement window.
    -- Only in_stmt_window = 1 call-days can be a standard episode.
    SELECT contactid
    FROM (
        SELECT contactid,
               row_number() OVER (PARTITION BY acct_key, call_dt
                                  ORDER BY initiationtimestamp) AS rn
        FROM calls_anchored
        WHERE acct_key IS NOT NULL AND acct_key <> ''
          AND is_biz = 0
          AND within_effdt_cap = 1
          AND in_stmt_window = 1                             -- RE-ANCHOR: drop call-days outside any statement window
    )
    WHERE rn = 1
)
SELECT c.acct_key, c.contactid, c.call_dt, c.call_month,
       c.is_biz, c.within_effdt_cap, c.had_zero_pad,
       c.stmt_dt, c.days_since_stmt_dt, c.in_stmt_window,
       c.pre_due_f, c.post_due_f, c.stmt_5day_bucket, c.stmt_5day_bucket_start,
       CASE WHEN e.contactid IS NOT NULL THEN 1 ELSE 0 END AS is_episode_std
FROM calls_anchored c
LEFT JOIN episodes_std e ON e.contactid = c.contactid
""")
print(f"built {DB}.uc2_t16_02n_episodes")

# COMMAND ----------

# MAGIC %md
# MAGIC ## K7. Build `uc2_t16_03n_signals` (the ONE transcript pass)

# COMMAND ----------

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_03n_signals AS
WITH drivers AS (
    SELECT DISTINCT e.contactid
    FROM {DB}.uc2_t16_02n_episodes e
    LEFT JOIN {DB}.uc2_t16_01n_populations p ON p.acct_key = e.acct_key
    WHERE e.is_episode_std = 1
      AND (p.acct_key IS NULL OR p.is_exaa)
),
tx AS (
    SELECT t.contactid,
           max(CASE WHEN regexp_like(lower(t.content),
                     'passed away|death certificate|executor|deceased|calling on behalf') THEN 1 ELSE 0 END) AS deceased_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'pay|paid|payment|settle|payment plan|arrangement|work something out') THEN 1 ELSE 0 END) AS pay_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'settle|payment plan|arrangement|work something out') THEN 1 ELSE 0 END) AS plan_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford') THEN 1 ELSE 0 END) AS hard_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'dispute|not my charge|didn.t authorize|did not authorize|unauthorized|fraud|identity theft') THEN 1 ELSE 0 END) AS dispute_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'i.ll pay|i will pay|going to pay|gonna pay|pay (on|by|this|next)|when i get paid|payday|after my paycheck') THEN 1 ELSE 0 END) AS promise_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'bank routing|routing number|check number|checkbook|a check for|that check|on the check') THEN 1 ELSE 0 END) AS exec_f
    FROM {TX} t
    JOIN drivers d ON t.contactid = d.contactid
    WHERE t.effdt >= '{EFFDT_CAP_START}' AND t.effdt < '{EFFDT_CAP_END}'
      AND t.content IS NOT NULL
      AND t.participantid = 'CUSTOMER'
      AND regexp_like(lower(t.content),
            'pay|paid|payment|settle|arrangement|work something out|passed away|death certificate|executor|deceased|calling on behalf|hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford|dispute|not my charge|didn.t authorize|did not authorize|unauthorized|fraud|identity theft|i.ll pay|i will pay|going to pay|gonna pay|when i get paid|payday|after my paycheck|bank routing|routing number|check number|checkbook|a check for|that check|on the check')
    GROUP BY 1
)
SELECT contactid,
       deceased_f, promise_f, pay_f, plan_f, hard_f, dispute_f, exec_f,
       CASE
         WHEN deceased_f > 0 THEN 'a. deceased or estate'
         WHEN promise_f  > 0 THEN 'b. future-dated promise'
         WHEN pay_f > 0 AND plan_f = 0 THEN 'c. payment talk, no promise'
         WHEN plan_f     > 0 THEN 'd. plan or settlement talk'
         WHEN hard_f     > 0 THEN 'e. hardship talk'
         WHEN dispute_f  > 0 THEN 'f. dispute or fraud talk'
         ELSE 'g. no payment-related language'
       END AS language_group
FROM tx
""")
print(f"built {DB}.uc2_t16_03n_signals")

# COMMAND ----------

# MAGIC %md
# MAGIC ## K8. Build `uc2_t16_04n_outcomes` (has_tx diagnostic; AWS day-grain gate is a DIAGNOSTIC)

# COMMAND ----------

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_t16_04n_outcomes AS
WITH episodes_exaa AS (
    SELECT c.acct_key, c.contactid, c.call_dt, c.call_month,
           c.stmt_dt, c.days_since_stmt_dt,
           c.pre_due_f, c.post_due_f, c.stmt_5day_bucket, c.stmt_5day_bucket_start
    FROM {DB}.uc2_t16_02n_episodes c
    LEFT JOIN {DB}.uc2_t16_01n_populations p ON p.acct_key = c.acct_key
    WHERE c.is_episode_std = 1
      AND (p.acct_key IS NULL OR p.is_exaa)
),
pay_lead AS (
    SELECT acct_key,
           to_date(concat(ym, '01'), 'yyyyMMdd') AS m,
           pay_dt, auto_dt, nsf_dt,
           lead(pay_dt)  OVER (PARTITION BY acct_key ORDER BY ym) AS next_pay_dt,
           lead(auto_dt) OVER (PARTITION BY acct_key ORDER BY ym) AS next_auto_dt,
           lead(nsf_dt)  OVER (PARTITION BY acct_key ORDER BY ym) AS next_nsf_dt
    FROM {DB}.uc2_t16_00n_acct_monthly
    WHERE acct_key IN (SELECT acct_key FROM episodes_exaa)
),
ep AS (
    SELECT e.acct_key, e.contactid, e.call_dt, e.call_month,
           e.stmt_dt, e.days_since_stmt_dt,
           e.pre_due_f, e.post_due_f, e.stmt_5day_bucket, e.stmt_5day_bucket_start,
           CASE WHEN
                  (s.pay_dt IS NOT NULL
                   AND s.pay_dt >= e.call_dt
                   AND s.pay_dt <= date_add(e.call_dt, 30)
                   AND (s.auto_dt IS NULL OR s.auto_dt <> s.pay_dt)
                   AND (s.nsf_dt IS NULL OR s.nsf_dt <> s.pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.pay_dt))
                OR
                  (s.next_pay_dt IS NOT NULL
                   AND s.next_pay_dt >= e.call_dt
                   AND s.next_pay_dt <= date_add(e.call_dt, 30)
                   AND (s.next_auto_dt IS NULL OR s.next_auto_dt <> s.next_pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.next_pay_dt))
                THEN 1 ELSE 0 END AS captured
    FROM episodes_exaa e
    LEFT JOIN pay_lead s
      ON s.acct_key = e.acct_key
     AND s.m = e.call_month
),
snap_daily AS (
    SELECT {NUM_KEY.format(c="extnl_acct_id")} AS acct_key,
           eff_dt,
           to_date(eff_dt, 'yyyyMMdd') AS snap_dt,
           CASE
             WHEN past_due_271_up_amt  > 0 THEN 10
             WHEN past_due_241_270_amt > 0 THEN 9
             WHEN past_due_211_240_amt > 0 THEN 8
             WHEN past_due_181_210_amt > 0 THEN 7
             WHEN past_due_151_180_amt > 0 THEN 6
             WHEN past_due_121_150_amt > 0 THEN 5
             WHEN past_due_91_120_amt  > 0 THEN 4
             WHEN past_due_61_90_amt   > 0 THEN 3
             WHEN past_due_31_60_amt   > 0 THEN 2
             WHEN past_due_1_30_amt    > 0 THEN 1
             ELSE 0
           END AS bucket,
           try_cast(chrgoff_dt AS date) AS co_dt
    FROM {FMT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '{SNAP_DAILY_START}' AND eff_dt < '{SNAP_DAILY_END}'
      AND {NUM_KEY.format(c="extnl_acct_id")} IN (SELECT acct_key FROM episodes_exaa)
),
callday AS (
    SELECT e.acct_key, e.call_dt,
           max_by(s.bucket, s.eff_dt) AS callday_bucket,
           max_by(s.co_dt, s.eff_dt) AS callday_co_dt
    FROM (SELECT DISTINCT acct_key, call_dt FROM episodes_exaa) e
    JOIN snap_daily s
      ON s.acct_key = e.acct_key
     AND s.snap_dt <= e.call_dt
    GROUP BY 1, 2
),
esig AS (
    SELECT e.acct_key, e.contactid, e.call_dt, e.captured,
           e.stmt_dt, e.days_since_stmt_dt,
           e.pre_due_f, e.post_due_f, e.stmt_5day_bucket, e.stmt_5day_bucket_start,
           coalesce(x.language_group, 'g. no payment-related language') AS language_group,
           coalesce(x.pay_f, 0)      AS pay_f,
           coalesce(x.deceased_f, 0) AS deceased_f,
           coalesce(x.exec_f, 0)     AS exec_f,
           CASE WHEN x.contactid IS NOT NULL THEN 1 ELSE 0 END AS has_tx,
           cd.callday_bucket,
           (cd.callday_bucket = 1
            AND (cd.callday_co_dt IS NULL OR cd.callday_co_dt >= DATE '{CLEANUP_DATE}'))
               AS is_addressable
    FROM ep e
    LEFT JOIN {DB}.uc2_t16_03n_signals x ON x.contactid = e.contactid
    LEFT JOIN callday cd ON cd.acct_key = e.acct_key AND cd.call_dt = e.call_dt
),
esig_acct AS (
    SELECT *,
           max(captured)   OVER (PARTITION BY acct_key) AS any_captured,
           max(CASE WHEN captured = 0 AND pay_f > 0 THEN 1 ELSE 0 END)
                           OVER (PARTITION BY acct_key) AS any_leaked_intent,
           max(deceased_f) OVER (PARTITION BY acct_key) AS deceased_acct
    FROM esig
)
SELECT a.acct_key, a.contactid, a.call_dt,
       a.stmt_dt, a.days_since_stmt_dt,
       a.pre_due_f, a.post_due_f, a.stmt_5day_bucket, a.stmt_5day_bucket_start,
       a.captured, a.language_group, a.pay_f, a.deceased_f, a.exec_f, a.has_tx,
       a.callday_bucket, a.is_addressable,
       a.any_captured, a.any_leaked_intent, a.deceased_acct,
       CASE
         WHEN a.any_captured = 1 THEN 'b. captured (>= 1 paid-30d episode)'
         WHEN a.any_leaked_intent = 1 THEN 'c. leaked-intent (intent, no payment 30d)'
         ELSE 'd. other-caller'
       END AS caller_class,
       (a.any_captured = 0 AND a.any_leaked_intent = 1) AS leaked_acct,
       (a.any_captured = 0 AND a.any_leaked_intent = 1
        AND coalesce(p.in_ledger_exaa, false) AND a.deceased_acct = 0) AS w_flag,
       coalesce(p.in_ledger_all, false)  AS in_ledger_all,
       coalesce(p.in_ledger_exaa, false) AS in_ledger_exaa,
       coalesce(p.touched_b1, false)     AS touched_b1,
       p.touched_b1_class,
       p.eom_bal AS jan_eom_bal, p.cpc_class, p.runway_band,
       p.feb_position_b14, p.feb_pos, p.mar_pos,
       p.co_dt_future, p.co_amt, p.co_8m, p.co_10m, p.co_12m
FROM esig_acct a
LEFT JOIN {DB}.uc2_t16_01n_populations p ON p.acct_key = a.acct_key
""")
print(f"built {DB}.uc2_t16_04n_outcomes")

# COMMAND ----------

print("B02_keyfix_aws_layers build complete: uc2_t16_00n/01n/02n/03n/04n. "
      "Run B02_checks.py once to certify the population anchors, call-table "
      "evidence ties, the re-anchor, and the 202501 recovery reconciliation.")
