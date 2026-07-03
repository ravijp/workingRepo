-- Tier 5 | Inbound episodes between DQ1 entry and charge-off (W3 entrants)
-- For accounts that entered DQ1 in 2024-07 .. 2025-06 and went on to charge off
-- (by 2026-02): how many inbound episodes (first inbound per account per day)
-- did they place between entry and the charge-off date? Sizes the intervention
-- surface on the loss path: '0 calls' losses are unreachable inbound; the 1-2
-- call band is where a single missed capture is the whole story.
-- Entrant detection uses a 3-month lookback buffer (snapshots from 2024-04) so
-- a bucket-1 month preceded by silence is not miscounted as an entry.
WITH snap AS (
    SELECT extnl_acct_id,
           date_trunc('month', date(date_parse(eff_dt, '%Y%m%d'))) AS m,
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
      AND date(date_parse(eff_dt, '%Y%m%d')) >= DATE '2024-04-01'
      AND date(date_parse(eff_dt, '%Y%m%d')) < DATE '2026-03-01'
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, min(co_dt) AS co_dt
    FROM snap GROUP BY 1, 2
),
entry AS (
    SELECT extnl_acct_id, m, bucket,
           lag(bucket) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS prev_bucket,
           min(co_dt) OVER (PARTITION BY extnl_acct_id) AS co_dt
    FROM monthly
),
cohort AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           cast(m AS date) AS start_m, co_dt
    FROM entry
    WHERE bucket = 1
      AND coalesce(prev_bucket, 0) = 0
      AND m >= DATE '2024-07-01' AND m < DATE '2025-07-01'
      AND co_dt IS NOT NULL
      AND co_dt >= cast(m AS date)
      AND co_dt < DATE '2026-03-01'
),
episodes AS (
    SELECT acct_key, call_dt
    FROM (
        SELECT trim(cast(acctid AS varchar)) AS acct_key, "date" AS call_dt,
               row_number() OVER (PARTITION BY trim(cast(acctid AS varchar)), "date"
                                  ORDER BY initiationtimestamp) AS rn
        FROM "contactcenter_bdp_db"."call"
        WHERE initiationmethod = 'INBOUND'
          AND "date" >= DATE '2024-07-01' AND "date" < DATE '2026-03-01'
          AND acctid IS NOT NULL
    )
    WHERE rn = 1
),
per_acct AS (
    SELECT c.acct_key,
           count(e.call_dt) AS n
    FROM cohort c
    LEFT JOIN episodes e
      ON c.acct_key = e.acct_key
     AND e.call_dt >= c.start_m
     AND e.call_dt <= c.co_dt
    GROUP BY 1
)
SELECT CASE
         WHEN n = 0  THEN 'a. 0 calls'
         WHEN n = 1  THEN 'b. 1 call'
         WHEN n = 2  THEN 'c. 2 calls'
         WHEN n <= 5 THEN 'd. 3-5 calls'
         WHEN n <= 10 THEN 'e. 6-10 calls'
         ELSE 'f. 11+ calls'
       END AS episode_band,
       count(*) AS accounts,
       round(100.0 * count(*) / sum(count(*)) OVER (), 1) AS pct_of_accounts,
       sum(n) AS total_episodes
FROM per_acct
GROUP BY 1
ORDER BY 1
