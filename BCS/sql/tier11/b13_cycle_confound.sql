-- Tier 11 | Cycle confound: cure-by-EOM as a function of entry day-of-month
-- The 69/31 cured-vs-still split is a calendar-month cut: an account entering
-- DQ1 on Jan 3 had four weeks of in-month runway, one entering Jan 27 had four
-- days (Anupam 07-08: reporting is EOM, collections runs on cycles). The
-- account table is daily, so entry day is direct: the first January snapshot
-- with a past-due bucket >= 1. One row per entry day for the cleaned entrant
-- cohort (b7-b11's classes a+b, expected 492,074 total): entrants, cured by
-- Jan 31, still DQ1 at Jan 31, % cured. If % cured falls with entry day, part
-- of the split is timing, not behavior; report the controlled split either way.
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
           min(co_dt) AS co_dt,
           min(CASE WHEN bucket >= 1 THEN eff_dt END) AS first_dq_dt
    FROM snap GROUP BY 1, 2
),
base AS (
    SELECT j.extnl_acct_id, j.max_bucket, j.eom_bucket, j.co_dt, j.first_dq_dt,
           p.max_bucket AS prev_max_bucket
    FROM (SELECT * FROM monthly WHERE ym = '202501') j
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202412') p
      ON j.extnl_acct_id = p.extnl_acct_id
),
entrants AS (
    SELECT cast(substr(first_dq_dt, 7, 2) AS integer) AS entry_day,
           eom_bucket
    FROM base
    WHERE max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0
      AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
)
SELECT entry_day AS b13_entry_day,
       31 - entry_day AS b13_days_runway,
       count(*) AS b13_entrants,
       count_if(eom_bucket = 0) AS b13_cured_by_eom,
       count_if(eom_bucket >= 1) AS b13_still_dq_eom,
       round(100.0 * count_if(eom_bucket = 0) / count(*), 1) AS b13_pct_cured
FROM entrants
GROUP BY 1
ORDER BY 1
