-- Tier 2 | FALLBACK for t2_dpd_buckets: ladder capped at 181-210
-- Used automatically if the deeper past-due columns (211+) are absent.
WITH latest AS (SELECT max(eff_dt) AS d FROM "fmt_acct_dba"."fmt_acct_c")
SELECT CASE
         WHEN past_due_181_210_amt > 0 THEN 7
         WHEN past_due_151_180_amt > 0 THEN 6
         WHEN past_due_121_150_amt > 0 THEN 5
         WHEN past_due_91_120_amt  > 0 THEN 4
         WHEN past_due_61_90_amt   > 0 THEN 3
         WHEN past_due_31_60_amt   > 0 THEN 2
         WHEN past_due_1_30_amt    > 0 THEN 1
         ELSE 0
       END AS dpd_bucket,
       count(*) AS accounts,
       round(sum(try_cast(acct_bal_amt AS double)), 0) AS total_balance
FROM "fmt_acct_dba"."fmt_acct_c", latest
WHERE sfx_nbr = 0
  AND eff_dt = latest.d
GROUP BY 1
ORDER BY 1
