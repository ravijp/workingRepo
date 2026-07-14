-- Tier 16 | LAYER 04: outcomes. Builds on 01 + 02 + 03. Grain: one row per
-- STANDARD January episode on an ex-AA account (the b19 driver set), carrying:
--   * captured        : the clean 30-day payment capture gate, verbatim
--                       (payment within 30 days of the call, not an autopay
--                       date, not an NSF date, next-month dates checked too)
--   * language_group  : the v2 lexicon class (03; missing transcript -> 'g')
--   * callday_bucket / is_addressable : the as-of-call-day bucket join
--                       (bucket 1 at call + no pre-2025 charge-off)
--   * caller_class    : account-level b14 class among callers (captured /
--                       leaked-intent / other-caller); 'a. non-caller' rows
--                       exist only on the 01 side (see the README worked example)
--   * leaked_acct, deceased_acct, w_flag : the strict leak list and W routing.
--     Strict leak list = account never captured AND >= 1 uncaptured episode
--     with payment language. W = leaked, in the ex-AA ledger, NO deceased
--     language (deceased accounts route to estate handling, not W).
--   * account attributes from 01 (ledger flags, balance, runway, positions,
--     CPC class, CO windows) repeated on every episode row for easy grouping.
--
-- DEPENDS ON: 00 (payment-date lead), 01, 02, 03, plus ONE extra base scan:
-- the day-grain call-day snapshot (snap_daily, 2024-06-01 lookback), which a
-- monthly layer cannot provide; it is semi-joined to episode accounts first,
-- per the verified resource fixes.
--
-- TIE-OUT ANCHORS (STOP RULE; any miss after restructuring = STOP):
--   * count_if(in_ledger_exaa)                      = 11,262 episodes;
--     distinct acct_key among them                  = 9,389 EXACTLY
--   * count_if(is_addressable)                      = 29,114 episodes EXACTLY
--   * distinct acct_key with w_flag                 = 1,765 EXACTLY;
--     their Jan EOM balance (one row per account)   = $7,690,886
--   * the addressable work list                     = 1,863 accounts
--     (walkthrough v5 anchor; compute per the recorded addressable-partition
--     definition, INSIGHT 8.3)
--   * language partition over in_ledger_exaa episodes sums to 11,262 with the
--     m4 row values (a 498 / b 1,374 / c 5,164 / d 442 / e 79 / f 274 / g 3,431)
--
-- WITH TABLE ACCESS, uncomment:
-- CREATE TABLE <schema>.uc2_t16_04_outcomes AS
WITH acct_monthly AS (
    -- TABLE MODE (later): keep; STITCH MODE (today): paste layer 00 here.
    SELECT * FROM "<schema>"."uc2_t16_00_acct_monthly"
),
populations AS (
    -- TABLE MODE (later): keep; STITCH MODE (today): paste layer 01's CTEs.
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
),
calls AS (
    -- TABLE MODE (later): keep; STITCH MODE (today): paste layer 02's CTEs.
    SELECT * FROM "<schema>"."uc2_t16_02_episodes"
),
signals AS (
    -- TABLE MODE (later): keep; STITCH MODE (today): paste layer 03's CTEs.
    SELECT * FROM "<schema>"."uc2_t16_03_signals"
),
-- standard episodes on ex-AA accounts: the driver set (matches 03's prune).
-- NULL-safe ex-AA intersection, matching the verified b19 LEFT JOIN: a
-- calling account with NO anchor-month account row is KEPT (unknown cpc is
-- kept as "others" by the NULL-safe rule).
episodes_exaa AS (
    SELECT c.acct_key, c.contactid, c.call_dt, c.call_month
    FROM calls c
    LEFT JOIN populations p ON p.acct_key = c.acct_key
    WHERE c.is_episode_std = 1
      AND (p.acct_key IS NULL OR p.is_exaa)
),
-- payment-date lead over the monthly account layer (no new base scan);
-- lead(m) of the January row = the February row, as in the verified kit
pay_lead AS (
    SELECT acct_key,
           cast(date_parse(concat(ym, '01'), '%Y%m%d') AS date) AS m,
           pay_dt, auto_dt, nsf_dt,
           lead(pay_dt)  OVER (PARTITION BY acct_key ORDER BY ym) AS next_pay_dt,
           lead(auto_dt) OVER (PARTITION BY acct_key ORDER BY ym) AS next_auto_dt,
           lead(nsf_dt)  OVER (PARTITION BY acct_key ORDER BY ym) AS next_nsf_dt
    FROM acct_monthly
    -- logic-safe prune (lead partitions per account; non-calling accounts
    -- contribute nothing downstream), per the verified resource fixes
    WHERE acct_key IN (SELECT acct_key FROM episodes_exaa)
),
-- the capture gate, verbatim
ep AS (
    SELECT e.acct_key, e.contactid, e.call_dt, e.call_month,
           CASE WHEN
                  (s.pay_dt IS NOT NULL
                   AND s.pay_dt >= e.call_dt
                   AND s.pay_dt <= date_add('day', 30, e.call_dt)
                   AND (s.auto_dt IS NULL OR s.auto_dt <> s.pay_dt)
                   AND (s.nsf_dt IS NULL OR s.nsf_dt <> s.pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.pay_dt))
                OR
                  (s.next_pay_dt IS NOT NULL
                   AND s.next_pay_dt >= e.call_dt
                   AND s.next_pay_dt <= date_add('day', 30, e.call_dt)
                   AND (s.next_auto_dt IS NULL OR s.next_auto_dt <> s.next_pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.next_pay_dt))
                THEN 1 ELSE 0 END AS captured
    FROM episodes_exaa e
    LEFT JOIN pay_lead s
      ON s.acct_key = e.acct_key
     AND s.m = e.call_month
),
-- the ONE extra base scan: daily snapshots for the as-of-call-day bucket,
-- semi-joined to episode accounts BEFORE anything heavy runs (verified fix)
snap_daily AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           eff_dt,
           date(date_parse(eff_dt, '%Y%m%d')) AS snap_dt,
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
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20240601' AND eff_dt < '20250201'   -- PARAM: lookback floor
      AND trim(cast(extnl_acct_id AS varchar)) IN (SELECT acct_key FROM episodes_exaa)
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
-- episode + signal + call-day view, then account-level flags as window
-- functions (NO second transcript touch; the round-9 m4 pattern)
esig AS (
    SELECT e.acct_key, e.contactid, e.call_dt, e.captured,
           coalesce(x.language_group, 'g. no payment-related language') AS language_group,
           coalesce(x.pay_f, 0)      AS pay_f,
           coalesce(x.deceased_f, 0) AS deceased_f,
           coalesce(x.exec_f, 0)     AS exec_f,
           cd.callday_bucket,
           (cd.callday_bucket = 1
            AND (cd.callday_co_dt IS NULL OR cd.callday_co_dt >= DATE '2025-01-01'))
               AS is_addressable
    FROM ep e
    LEFT JOIN signals x ON x.contactid = e.contactid
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
       a.captured, a.language_group, a.pay_f, a.deceased_f, a.exec_f,
       a.callday_bucket, a.is_addressable,
       a.any_captured, a.any_leaked_intent, a.deceased_acct,
       -- b14 caller class among callers ('a. non-caller' lives on the 01 side)
       CASE
         WHEN a.any_captured = 1 THEN 'b. captured (>= 1 paid-30d episode)'
         WHEN a.any_leaked_intent = 1 THEN 'c. leaked-intent (intent, no payment 30d)'
         ELSE 'd. other-caller'
       END AS caller_class,
       -- strict leak list and W routing
       (a.any_captured = 0 AND a.any_leaked_intent = 1) AS leaked_acct,
       (a.any_captured = 0 AND a.any_leaked_intent = 1
        AND coalesce(p.in_ledger_exaa, false) AND a.deceased_acct = 0) AS w_flag,
       -- account attributes from 01, repeated per episode for easy grouping;
       -- NULL / false when the calling account has no anchor-month row
       coalesce(p.in_ledger_all, false)  AS in_ledger_all,
       coalesce(p.in_ledger_exaa, false) AS in_ledger_exaa,
       coalesce(p.touched_b1, false)     AS touched_b1,
       p.touched_b1_class,
       p.eom_bal AS jan_eom_bal, p.cpc_class, p.runway_band,
       p.feb_position_b14, p.feb_pos, p.mar_pos,
       p.co_dt_future, p.co_amt, p.co_8m, p.co_10m, p.co_12m
FROM esig_acct a
LEFT JOIN populations p ON p.acct_key = a.acct_key
