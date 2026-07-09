-- Tier 11 | Call-day population: the operable January call stream, by what
-- is knowable AT CALL TIME. January inbound episodes (f1's convention: first
-- inbound leg per account per day, business-card excluded) joined to the
-- account's LATEST daily snapshot on or before the call date - the as-of-call
-- read, not the month-max tag. Kept: episodes whose call-day snapshot shows
-- past-due bucket >= 1; pre-2025 chrgoff_dt stock excluded (the b12 cleanup,
-- read off the call-day snapshot). Rows: bucket at call (1 / 2-6 / 7+) x
-- days since the current delinquency spell's first delinquent snapshot
-- (0-10 / 11-20 / 21+). Spell start = the first bucket >= 1 snapshot after
-- the account's last bucket-0 snapshot at or before the call; snapshot
-- lookback floor 2024-06-01 (v11/b2's floor - spells older than that band as
-- 21+ anyway). No transcript pass. Sanity expectation, not a hard tie-out:
-- total episodes <= 122,606 (b7's delinquent-in-month sum); the shortfall is
-- calls made before the account's first delinquent snapshot of the month.
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
callday AS (
    SELECT e.acct_key, e.call_dt,
           max_by(s.bucket, s.eff_dt) AS callday_bucket,
           max_by(s.co_dt, s.eff_dt) AS callday_co_dt,
           max(CASE WHEN s.bucket = 0 THEN s.snap_dt END) AS last_current_dt
    FROM episodes e
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
