-- Validation | Deep-bucket resets to current (re-age / cure signal, last 6 months)
-- There is no explicit re-age flag in these tables. This measures the raw
-- month-over-month transitions from DQ2+ straight to current. Some are real
-- payments (cure), some are program re-ages; splitting the two needs the
-- payment fields or a program table. Sizes the question first.
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
           END AS bucket
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN latest
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) > date_add('month', -7, latest.d)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket
    FROM snap GROUP BY 1, 2
),
trans AS (
    SELECT extnl_acct_id, m, bucket,
           lag(bucket) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS prev_bucket
    FROM monthly
)
SELECT prev_bucket AS from_bucket,
       count(*) AS account_months_observed,
       count_if(bucket = 0) AS reset_to_current,
       round(100.0 * count_if(bucket = 0) / count(*), 2) AS pct_reset_to_current
FROM trans
WHERE prev_bucket >= 2
GROUP BY 1
ORDER BY 1
