-- Tier 14 | EX-AA VARIANT of b17_callday_population.
-- Exclusion: applied on the ACCOUNT side of the call-day join, using the
-- account's January EOM cpc (eom_cpc, max_by(clnt_prdct_cd, eff_dt) over
-- January snapshots only - the same monthly CTE pattern as b14, added here
-- as a small separate CTE since b17's own `snap` is daily-grain across a
-- 2024-06-through-2025-01 window and has no single monthly-EOM concept).
-- The exclusion is NOT applied to the daily as-of-call-day bucket logic
-- itself (callday/spell), which is untouched. NULL-safe form: NULL or blank
-- cpc is kept as "others".
-- Expected tie-outs: episodes <= 48,106 (the brief's overall run-order
-- ceiling for b17-family queries); the 'a. bucket 1 at call' cell
-- <= 32,712 (the original unfiltered b17 cell - ex-AA removes accounts
-- only, so strictly less-or-equal). Sanity: total episodes stays
-- <= the original's own <= 122,606 bound (b7's delinquent-in-month sum),
-- now also bounded by ex-AA b7_exaa's own class total.
-- Everything else is unchanged from b17_callday_population.sql: the daily
-- snapshot lookback floor (2024-06-01), the episode dedup, the as-of-call-
-- day bucket/spell logic, and the bucket/days-since bands.
WITH snap AS (
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
      AND eff_dt >= '20240601' AND eff_dt < '20250201'
),
cpc_monthly AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250201'
    GROUP BY 1
),
inb AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key, contactid,
           "date" AS call_dt, initiationtimestamp
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2025-01-01' AND "date" < DATE '2025-02-01'
      AND effdt >= '2025-01-01' AND effdt < '2025-02-02'
      AND coalesce(cast(producttype AS varchar), '') <> 'BUSINESS_CARD'
),
episodes AS (
    SELECT acct_key, call_dt
    FROM (
        SELECT acct_key, call_dt,
               row_number() OVER (PARTITION BY acct_key, call_dt
                                  ORDER BY initiationtimestamp) AS rn
        FROM inb
        WHERE acct_key IS NOT NULL AND acct_key <> ''
    )
    WHERE rn = 1
),
episodes_exaa AS (
    SELECT e.acct_key, e.call_dt
    FROM episodes e
    LEFT JOIN cpc_monthly c ON c.acct_key = e.acct_key
    WHERE (c.eom_cpc IS NULL OR trim(c.eom_cpc) = ''
           OR c.eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                                 'AA3','AC3','AM3','AA4','AC4','AM4',
                                 'BGC','BGM','CGM','GMR',
                                 'FBS','IBS','U1C','U2C','U3C'))
),
callday AS (
    SELECT e.acct_key, e.call_dt,
           max_by(s.bucket, s.eff_dt) AS callday_bucket,
           max_by(s.co_dt, s.eff_dt) AS callday_co_dt,
           max(CASE WHEN s.bucket = 0 THEN s.snap_dt END) AS last_current_dt
    FROM episodes_exaa e
    JOIN snap s
      ON s.acct_key = e.acct_key
     AND s.snap_dt <= e.call_dt
    GROUP BY 1, 2
),
kept AS (
    SELECT acct_key, call_dt, callday_bucket, last_current_dt
    FROM callday
    WHERE callday_bucket >= 1
      AND (callday_co_dt IS NULL OR callday_co_dt >= DATE '2025-01-01')
),
spell AS (
    SELECT k.acct_key, k.call_dt, k.callday_bucket,
           min(CASE WHEN s.bucket >= 1
                     AND s.snap_dt <= k.call_dt
                     AND s.snap_dt > coalesce(k.last_current_dt, DATE '1900-01-01')
                    THEN s.snap_dt END) AS spell_start_dt
    FROM kept k
    JOIN snap s ON s.acct_key = k.acct_key
    GROUP BY 1, 2, 3
)
SELECT CASE
         WHEN callday_bucket = 1 THEN 'a. bucket 1 at call'
         WHEN callday_bucket <= 6 THEN 'b. bucket 2-6 at call'
         ELSE 'c. bucket 7+ at call'
       END AS b17_bucket_at_call,
       CASE
         WHEN date_diff('day', spell_start_dt, call_dt) <= 10 THEN 'a. 0-10 days delinquent'
         WHEN date_diff('day', spell_start_dt, call_dt) <= 20 THEN 'b. 11-20 days delinquent'
         ELSE 'c. 21+ days delinquent'
       END AS b17_days_since_first_dq,
       count(*) AS b17_episodes,
       count(DISTINCT acct_key) AS b17_accounts
FROM spell
GROUP BY 1, 2
ORDER BY 1, 2
