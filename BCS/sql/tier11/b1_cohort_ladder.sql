-- Tier 11 | The 202501 cohort bridge: month-max entrants vs month-end DQ1
-- Decomposes the Athena 493,139 (month-MAX DQ1 entrants, v11) vs SAS/ASP
-- 186,412 (month-END DLNQT_CD=1) gap into measured steps. Grain is stated
-- on every step: month-MAX = worst bucket on any snapshot in the month;
-- month-END = bucket on the last snapshot of the month (max_by(bucket, eff_dt)).
-- Prior month here is strictly 202412 (Ishant's definition); the v11 lag
-- definition is reconciled separately in b2_entry_lookback.
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
           try_cast(acct_bal_amt AS double) AS bal
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20241201' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           max_by(bal, eff_dt) AS eom_bal
    FROM snap GROUP BY 1, 2
),
base AS (
    SELECT j.extnl_acct_id, j.max_bucket, j.eom_bucket, j.eom_bal,
           p.max_bucket AS prev_max_bucket,
           p.eom_bucket AS prev_eom_bucket,
           (p.extnl_acct_id IS NOT NULL) AS has_prior_row
    FROM (SELECT * FROM monthly WHERE ym = '202501') j
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202412') p
      ON j.extnl_acct_id = p.extnl_acct_id
)
SELECT 'a. all accounts with a 202501 snapshot' AS cohort_step,
       count(*) AS accounts_202501,
       round(sum(eom_bal), 0) AS eom_balance_202501
FROM base
UNION ALL SELECT 'b. month-MAX bucket >= 1', count(*), round(sum(eom_bal), 0)
FROM base WHERE max_bucket >= 1
UNION ALL SELECT 'c. month-MAX bucket = 1', count(*), round(sum(eom_bal), 0)
FROM base WHERE max_bucket = 1
UNION ALL SELECT 'd. c + 202412 month-MAX = 0 or no row (entrant, month-MAX grain, ~v11)', count(*), round(sum(eom_bal), 0)
FROM base WHERE max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0
UNION ALL SELECT 'e. d + month-END bucket >= 1 (still delinquent at month end)', count(*), round(sum(eom_bal), 0)
FROM base WHERE max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0 AND eom_bucket >= 1
UNION ALL SELECT 'f. d + month-END bucket = 1', count(*), round(sum(eom_bal), 0)
FROM base WHERE max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0 AND eom_bucket = 1
UNION ALL SELECT 'g. month-END bucket = 1, all accounts', count(*), round(sum(eom_bal), 0)
FROM base WHERE eom_bucket = 1
UNION ALL SELECT 'h. month-END bucket = 1 + 202412 month-END = 0 or no row (entrant, month-END grain; Ishant-comparable)', count(*), round(sum(eom_bal), 0)
FROM base WHERE eom_bucket = 1 AND coalesce(prev_eom_bucket, 0) = 0
UNION ALL SELECT 'h1. of h: 202412 month-END = 0 (new roll)', count(*), round(sum(eom_bal), 0)
FROM base WHERE eom_bucket = 1 AND prev_eom_bucket = 0
UNION ALL SELECT 'h2. of h: no 202412 row (no prior record)', count(*), round(sum(eom_bal), 0)
FROM base WHERE eom_bucket = 1 AND NOT has_prior_row
UNION ALL SELECT 'h3. excluded from h: 202412 month-END >= 1 (already delinquent)', count(*), round(sum(eom_bal), 0)
FROM base WHERE eom_bucket = 1 AND prev_eom_bucket >= 1
UNION ALL SELECT 'i. month-END bucket >= 1 + entrant on month-END grain (any-code entrant; vs Ishant grand total)', count(*), round(sum(eom_bal), 0)
FROM base WHERE eom_bucket >= 1 AND coalesce(prev_eom_bucket, 0) = 0
ORDER BY 1
