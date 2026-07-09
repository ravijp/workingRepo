-- Tier 11 | M2/M3 roll mirror: the January EOM bucket-1 stock, two months on
-- Mirrors Ishant's DLNQT_CD_M1=1 x M2 x M3 pivot on our side. His M1=1 rows:
-- M2=0: 105,715 cured / M2=1: 15,002 / M2=2: 65,585 (of which M3=3: 41,975) /
-- deeper tiny. His base 186,412 is stock under ex-AA + CHRGOFF_RSN filters;
-- ours is 207,006 unfiltered, so compare SHARES not counts. Grain: month-END
-- (last snapshot per month) on all three months; 'co' = charge-off date lands
-- in the window; 'gone' = the account has no row that month.
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
      AND eff_dt >= '20241201' AND eff_dt < '20250401'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           max_by(bal, eff_dt) AS eom_bal,
           min(co_dt) AS co_dt
    FROM snap GROUP BY 1, 2
),
base AS (
    SELECT j.extnl_acct_id, j.eom_bucket, j.eom_bal,
           p.eom_bucket AS prev_eom_bucket,
           (p.extnl_acct_id IS NOT NULL) AS has_prior_row,
           m2.eom_bucket AS m2_eom_bucket,
           m2.co_dt AS m2_co_dt,
           (m2.extnl_acct_id IS NOT NULL) AS has_m2_row,
           m3.eom_bucket AS m3_eom_bucket,
           m3.co_dt AS m3_co_dt,
           (m3.extnl_acct_id IS NOT NULL) AS has_m3_row
    FROM (SELECT * FROM monthly WHERE ym = '202501') j
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202412') p
      ON j.extnl_acct_id = p.extnl_acct_id
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202502') m2
      ON j.extnl_acct_id = m2.extnl_acct_id
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202503') m3
      ON j.extnl_acct_id = m3.extnl_acct_id
    WHERE j.eom_bucket = 1
)
SELECT CASE
         WHEN m2_co_dt >= DATE '2025-02-01' AND m2_co_dt < DATE '2025-03-01' THEN 'co'
         WHEN NOT has_m2_row THEN 'gone'
         ELSE cast(m2_eom_bucket AS varchar)
       END AS m2_bucket_202502,
       CASE
         WHEN m3_co_dt >= DATE '2025-02-01' AND m3_co_dt < DATE '2025-04-01' THEN 'co'
         WHEN NOT has_m3_row THEN 'gone'
         ELSE cast(m3_eom_bucket AS varchar)
       END AS m3_bucket_202503,
       CASE
         WHEN coalesce(prev_eom_bucket, 0) = 0 THEN 'a. entrant (202412 EOM = 0 or no row)'
         ELSE 'b. already delinquent (202412 EOM >= 1)'
       END AS roll_entry_class,
       count(*) AS roll_accounts,
       round(sum(eom_bal), 0) AS roll_eom_balance_jan
FROM base
GROUP BY 1, 2, 3
ORDER BY 1, 2, 3
