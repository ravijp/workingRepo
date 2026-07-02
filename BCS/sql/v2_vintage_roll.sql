-- Validation | Roll / cure curve for a new-delinquency vintage
-- Cohort: accounts entering DQ1 (bucket 0 -> 1) 11 months before the newest
-- account snapshot; tracked over the following months. This produces the
-- EMPIRICAL stage-migration rates (the multipliers a provision read needs),
-- instead of assumed ones.
-- Note: accounts can drop out of later snapshots (sold, closed); the per-month
-- base is the accounts still visible that month.
WITH latest AS (
    SELECT max(date_trunc('month', date(date_parse(eff_dt, '%Y%m%d')))) AS d
    FROM "fmt_acct_dba"."fmt_acct_c" WHERE sfx_nbr = 0
),
snap AS (
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
    CROSS JOIN latest
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) > date_add('month', -14, latest.d)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, min(co_dt) AS co_dt
    FROM snap GROUP BY 1, 2
),
entry AS (
    SELECT extnl_acct_id, m, bucket,
           lag(bucket) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS prev_bucket
    FROM monthly
),
cohort AS (
    SELECT e.extnl_acct_id, e.m AS start_m
    FROM entry e
    CROSS JOIN latest
    WHERE e.bucket = 1
      AND coalesce(e.prev_bucket, 0) = 0
      AND e.m = date_add('month', -11, latest.d)
),
path AS (
    SELECT c.extnl_acct_id,
           date_diff('month', c.start_m, s.m) AS month_on_book,
           CASE
             WHEN s.co_dt IS NOT NULL AND date_trunc('month', s.co_dt) <= s.m THEN 'chargedoff'
             WHEN s.bucket = 0 THEN 'current'
             WHEN s.bucket <= 3 THEN 'dq1_3'
             ELSE 'dq4_plus'
           END AS state
    FROM cohort c
    JOIN monthly s
      ON c.extnl_acct_id = s.extnl_acct_id
     AND s.m >= c.start_m
)
SELECT month_on_book,
       count(*) AS accounts_visible,
       round(100.0 * count_if(state = 'current') / count(*), 1) AS pct_current,
       round(100.0 * count_if(state = 'dq1_3') / count(*), 1) AS pct_dq1_3,
       round(100.0 * count_if(state = 'dq4_plus') / count(*), 1) AS pct_dq4_plus,
       round(100.0 * count_if(state = 'chargedoff') / count(*), 1) AS pct_charged_off
FROM path
GROUP BY 1
ORDER BY 1
