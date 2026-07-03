-- Tier 6 | Balance profile by delinquency bucket (latest snapshot)
-- The per-account dollar input a sizing model needs: average and median
-- balance per bucket, not just the total. Charged-off accounts are excluded
-- so deep buckets show live balances, not written-off stock (t2_dpd_buckets
-- keeps them in, which is why its bucket 10 is huge).
-- Anchored to the newest account snapshot.
WITH latest AS (SELECT max(eff_dt) AS d FROM "fmt_acct_dba"."fmt_acct_c")
SELECT CASE
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
       END AS dpd_bucket,
       count(*) AS accounts,
       round(avg(try_cast(acct_bal_amt AS double)), 0) AS avg_balance,
       round(approx_percentile(try_cast(acct_bal_amt AS double), 0.5), 0) AS median_balance,
       round(sum(try_cast(acct_bal_amt AS double)), 0) AS total_balance
FROM "fmt_acct_dba"."fmt_acct_c", latest
WHERE sfx_nbr = 0
  AND eff_dt = latest.d
  AND chrgoff_dt IS NULL
GROUP BY 1
ORDER BY 1
