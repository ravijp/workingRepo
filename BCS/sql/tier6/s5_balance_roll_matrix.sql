-- Tier 6 | Dollar-weighted roll matrix: where each bucket's BALANCE goes next month
-- The account-count roll matrix (s2) says how many accounts move; this says
-- how many DOLLARS move. Balance-weighted migration is what a provision-shaped
-- read multiplies, and big-balance accounts do not roll like small ones.
-- Same construction as s2: last 12 complete account months, consecutive-month
-- transitions, already-charged-off accounts excluded from the base.
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
           try_cast(chrgoff_dt AS date) AS co_dt,
           try_cast(acct_bal_amt AS double) AS bal
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN latest
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) > date_add('month', -13, latest.d)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, min(co_dt) AS co_dt,
           max(bal) AS bal
    FROM snap GROUP BY 1, 2
),
seq AS (
    SELECT extnl_acct_id, m, bucket, bal,
           min(co_dt) OVER (PARTITION BY extnl_acct_id) AS co_dt,
           lead(bucket) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_bucket,
           lead(m) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_m
    FROM monthly
),
trans AS (
    SELECT bucket AS from_bucket, next_bucket, bal,
           CASE WHEN co_dt IS NOT NULL AND date_trunc('month', co_dt) <= next_m
                THEN 1 ELSE 0 END AS co_by_next
    FROM seq
    WHERE next_m = date_add('month', 1, m)
      AND (co_dt IS NULL OR date_trunc('month', co_dt) > m)
      AND bal IS NOT NULL
)
SELECT from_bucket,
       round(sum(bal), 0) AS total_balance,
       round(100.0 * sum(CASE WHEN co_by_next = 0 AND next_bucket = 0 THEN bal ELSE 0 END)
             / sum(bal), 2) AS pct_bal_to_current,
       round(100.0 * sum(CASE WHEN co_by_next = 0 AND next_bucket > 0
                               AND next_bucket < from_bucket THEN bal ELSE 0 END)
             / sum(bal), 2) AS pct_bal_improved,
       round(100.0 * sum(CASE WHEN co_by_next = 0 AND next_bucket = from_bucket
                               AND from_bucket > 0 THEN bal ELSE 0 END)
             / sum(bal), 2) AS pct_bal_same,
       round(100.0 * sum(CASE WHEN co_by_next = 0 AND next_bucket > from_bucket
                          THEN bal ELSE 0 END)
             / sum(bal), 2) AS pct_bal_deeper,
       round(100.0 * sum(CASE WHEN co_by_next = 1 THEN bal ELSE 0 END)
             / sum(bal), 2) AS pct_bal_charged_off
FROM trans
GROUP BY 1
ORDER BY 1
