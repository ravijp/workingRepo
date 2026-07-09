-- Tier 11 | Charge-off scope probes around 202501
-- Ishant's grand total (433,914) includes a straight-to-CO class next to the
-- DQ-code entrants. These probes size the charge-off populations Athena can
-- see in the same month: straight-to-CO with no delinquency bucket all
-- month, CO while delinquent, pre-2025 CO stock still carrying rows, and
-- deep-bucket accounts with no charge-off date at all.
WITH snap AS (
    SELECT extnl_acct_id,
           substr(eff_dt, 1, 6) AS ym,
           eff_dt,
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
           try_cast(acct_bal_amt AS double) AS bal,
           try_cast(chrgoff_dt AS date) AS co_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           max_by(bal, eff_dt) AS eom_bal,
           min(co_dt) AS co_dt
    FROM snap GROUP BY 1, 2
)
SELECT 'a. charged off in Jan 2025, month-MAX bucket = 0 (straight-to-CO; Ishant CO_VINTAGE analogue)' AS co_scope_step,
       count(*) AS co_accounts,
       round(sum(eom_bal), 0) AS co_eom_balance
FROM monthly
WHERE co_dt >= DATE '2025-01-01' AND co_dt < DATE '2025-02-01' AND max_bucket = 0
UNION ALL SELECT 'b. charged off in Jan 2025, month-MAX bucket >= 1', count(*), round(sum(eom_bal), 0)
FROM monthly
WHERE co_dt >= DATE '2025-01-01' AND co_dt < DATE '2025-02-01' AND max_bucket >= 1
UNION ALL SELECT 'c. charged off before Jan 2025, still carrying 202501 rows (stock)', count(*), round(sum(eom_bal), 0)
FROM monthly
WHERE co_dt < DATE '2025-01-01'
UNION ALL SELECT 'd. no charge-off date, month-END bucket = 10 (271+ past due)', count(*), round(sum(eom_bal), 0)
FROM monthly
WHERE co_dt IS NULL AND eom_bucket = 10
ORDER BY 1
