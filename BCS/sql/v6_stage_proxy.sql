-- Validation | Provision-stage proxy from the DQ ladder (latest snapshot)
-- Maps the bucket ladder onto an IFRS9-style staging shape:
--   stage-1 proxy = current + DQ1, stage-2 proxy = DQ2-3,
--   stage-3 proxy = DQ4+ (90+ days past due).
-- Accounts already charged off are excluded (write-off, not a stage).
-- Real staging has overrides (high-risk flags, stickiness); those fields are
-- not in these tables, so treat this as the shape, not the booked number.
-- Pair with v2_vintage_roll: balance x empirical roll-to-loss rate per stage
-- gives an evidence-based expected-loss ladder.
WITH latest AS (SELECT max(eff_dt) AS d FROM "fmt_acct_dba"."fmt_acct_c")
SELECT CASE
         WHEN past_due_91_120_amt > 0 OR past_due_121_150_amt > 0
           OR past_due_151_180_amt > 0 OR past_due_181_210_amt > 0
           OR past_due_211_240_amt > 0 OR past_due_241_270_amt > 0
           OR past_due_271_up_amt > 0
           THEN 'c. stage-3 proxy (90+ days)'
         WHEN past_due_31_60_amt > 0 OR past_due_61_90_amt > 0
           THEN 'b. stage-2 proxy (DQ2-3)'
         ELSE 'a. stage-1 proxy (current + DQ1)'
       END AS stage_proxy,
       count(*) AS accounts,
       round(sum(try_cast(acct_bal_amt AS double)), 0) AS total_balance,
       round(avg(try_cast(acct_bal_amt AS double)), 0) AS avg_balance
FROM "fmt_acct_dba"."fmt_acct_c"
CROSS JOIN latest
WHERE sfx_nbr = 0
  AND eff_dt = latest.d
  AND chrgoff_dt IS NULL
GROUP BY 1
ORDER BY 1
