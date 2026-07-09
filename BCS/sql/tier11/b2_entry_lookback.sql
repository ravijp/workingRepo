-- Tier 11 | Entry-definition lookback: why v11's entrant count is what it is
-- v11 calls an account a 2025-01 DQ1 entrant when its month-MAX bucket is 1
-- and the lag over PRESENT months (lookback from 2024-06) is 0 or absent.
-- That lag is not "202412": if the account's last sighting was, say, 202409,
-- the lag reads the 202409 bucket. A strict-202412 prior is tighter. The
-- difference a - b below is accounts whose last sighting predates December.
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
           END AS bucket
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20240601' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket
    FROM snap GROUP BY 1, 2
),
entry AS (
    SELECT extnl_acct_id, m, bucket,
           lag(bucket) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS prev_bucket,
           lag(m)      OVER (PARTITION BY extnl_acct_id ORDER BY m) AS prev_m
    FROM monthly
),
cand AS (
    SELECT * FROM entry
    WHERE m = DATE '2025-01-01' AND bucket = 1
)
SELECT 'a. v11 definition: last present month bucket = 0 or no prior row (lookback 2024-06)' AS entry_variant,
       count(*) AS accounts_monthmax_b1
FROM cand WHERE coalesce(prev_bucket, 0) = 0
UNION ALL SELECT 'b. strict prior month: 202412 present with month-MAX = 0', count(*)
FROM cand WHERE prev_m = DATE '2024-12-01' AND prev_bucket = 0
UNION ALL SELECT 'c. v11 entrants whose last present month is not 202412', count(*)
FROM cand WHERE coalesce(prev_bucket, 0) = 0
  AND (prev_m IS NULL OR prev_m <> DATE '2024-12-01')
UNION ALL SELECT 'd. of c: no prior row at all since 2024-06', count(*)
FROM cand WHERE coalesce(prev_bucket, 0) = 0 AND prev_m IS NULL
ORDER BY 1
