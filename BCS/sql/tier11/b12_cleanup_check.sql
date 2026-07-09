-- Tier 11 | Cleanup check: the pre-2025 charge-off exclusion, made visible
-- b7-b11 drop accounts whose January snapshot carries a charge-off date
-- BEFORE 2025-01-01 (the classed CTE: co_dt IS NULL OR co_dt >= DATE
-- '2025-01-01'). These are old charged-off accounts, mostly charge-off
-- reversals and re-aged stock, whose chrgoff_dt field still holds the
-- historical date; they are not January delinquency. This query shows the
-- exclusion exactly, and splits the 31,810 deeper-path month-end entrants
-- into their two causes. Expected from prior runs: a=152,177, b=1,064,
-- c=151,113, d=54,829, e=1,619, f=53,210, g=31,810 (h+i split is new).
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
           try_cast(chrgoff_dt AS date) AS co_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20241201' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           min(co_dt) AS co_dt
    FROM snap GROUP BY 1, 2
),
base AS (
    SELECT j.extnl_acct_id, j.max_bucket, j.eom_bucket, j.co_dt,
           p.max_bucket AS prev_max_bucket,
           p.eom_bucket AS prev_eom_bucket
    FROM (SELECT * FROM monthly WHERE ym = '202501') j
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202412') p
      ON j.extnl_acct_id = p.extnl_acct_id
)
SELECT 'a. strict entrants still bucket 1 at Jan 31 (ladder step f)' AS b12_check_step,
       count(*) AS b12_accounts
FROM base
WHERE max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0 AND eom_bucket = 1
UNION ALL SELECT 'b. of a: charge-off date before 2025-01-01 (excluded by cleanup)', count(*)
FROM base
WHERE max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0 AND eom_bucket = 1
  AND co_dt < DATE '2025-01-01'
UNION ALL SELECT 'c. of a: kept (= class b in b7-b11)', count(*)
FROM base
WHERE max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0 AND eom_bucket = 1
  AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
UNION ALL SELECT 'd. bucket-1 stock remainder at Jan 31 (EOM=1, not the strict entrant path)', count(*)
FROM base
WHERE eom_bucket = 1 AND NOT (max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0)
UNION ALL SELECT 'e. of d: charge-off date before 2025-01-01 (excluded by cleanup)', count(*)
FROM base
WHERE eom_bucket = 1 AND NOT (max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0)
  AND co_dt < DATE '2025-01-01'
UNION ALL SELECT 'f. of d: kept (= class c in b7-b11)', count(*)
FROM base
WHERE eom_bucket = 1 AND NOT (max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0)
  AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
UNION ALL SELECT 'g. of d: Dec 31 current (month-END entrants off the strict path)', count(*)
FROM base
WHERE eom_bucket = 1 AND NOT (max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0)
  AND coalesce(prev_eom_bucket, 0) = 0
UNION ALL SELECT 'h. of g: touched bucket 2+ during January (deeper path, paid back to 1)', count(*)
FROM base
WHERE eom_bucket = 1 AND coalesce(prev_eom_bucket, 0) = 0
  AND max_bucket >= 2
UNION ALL SELECT 'i. of g: bucket 1 all January, but past due inside December (cured by Dec 31)', count(*)
FROM base
WHERE eom_bucket = 1 AND coalesce(prev_eom_bucket, 0) = 0
  AND max_bucket = 1 AND coalesce(prev_max_bucket, 0) >= 1
ORDER BY 1
