-- Tier 6 | Leading-edge roll matrix: where each bucket goes next month (last 12 account months)
-- The measured migration rates, account-level and month-over-month - the
-- number a provision-shaped sizing needs instead of pooled balance-flow
-- percentages. For each bucket: the share that cures to current, improves,
-- holds, rolls deeper, or charges off in the NEXT month.
-- Consecutive-month transitions only; accounts already charged off are
-- excluded from the base. Anchored to the newest account month.
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
      AND date(date_parse(eff_dt, '%Y%m%d')) > date_add('month', -13, latest.d)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, min(co_dt) AS co_dt
    FROM snap GROUP BY 1, 2
),
seq AS (
    SELECT extnl_acct_id, m, bucket,
           min(co_dt) OVER (PARTITION BY extnl_acct_id) AS co_dt,
           lead(bucket) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_bucket,
           lead(m) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_m
    FROM monthly
),
trans AS (
    SELECT bucket AS from_bucket, next_bucket,
           CASE WHEN co_dt IS NOT NULL AND date_trunc('month', co_dt) <= next_m
                THEN 1 ELSE 0 END AS co_by_next
    FROM seq
    WHERE next_m = date_add('month', 1, m)
      AND (co_dt IS NULL OR date_trunc('month', co_dt) > m)
)
SELECT from_bucket,
       count(*) AS account_months,
       round(100.0 * count_if(co_by_next = 0 AND next_bucket = 0) / count(*), 2) AS pct_to_current,
       round(100.0 * count_if(co_by_next = 0 AND next_bucket > 0
                              AND next_bucket < from_bucket) / count(*), 2) AS pct_improved,
       round(100.0 * count_if(co_by_next = 0 AND next_bucket = from_bucket
                              AND from_bucket > 0) / count(*), 2) AS pct_same,
       round(100.0 * count_if(co_by_next = 0 AND next_bucket > from_bucket) / count(*), 2) AS pct_deeper,
       round(100.0 * count_if(co_by_next = 1) / count(*), 2) AS pct_charged_off
FROM trans
GROUP BY 1
ORDER BY 1
