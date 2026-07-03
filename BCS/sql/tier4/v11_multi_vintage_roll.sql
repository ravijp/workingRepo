-- Validation | Roll / cure curve across THREE vintages (2024-09, 2025-01, 2025-05)
-- One vintage is one seasonal draw. This tracks three DQ1-entry cohorts spread
-- across the funnel window, month on book 0-9, in one result - if the curves
-- agree, the migration rates are stable enough to quote; if they diverge, the
-- divergence IS the finding. Entrant detection keeps a 3-month lookback buffer.
WITH snap AS (
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
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= DATE '2024-06-01'
      AND date(date_parse(eff_dt, '%Y%m%d')) < DATE '2026-03-01'
      AND eff_dt >= '20240601' AND eff_dt < '20260301'
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, min(co_dt) AS co_dt
    FROM snap GROUP BY 1, 2
),
entry AS (
    SELECT extnl_acct_id, m, bucket,
           lag(bucket) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS prev_bucket
    FROM monthly
),
cohort AS (
    SELECT extnl_acct_id, m AS start_m
    FROM entry
    WHERE bucket = 1
      AND coalesce(prev_bucket, 0) = 0
      AND m IN (DATE '2024-09-01', DATE '2025-01-01', DATE '2025-05-01')
),
path AS (
    SELECT cast(c.start_m AS date) AS vintage,
           date_diff('month', c.start_m, s.m) AS month_on_book,
           CASE
             WHEN s.co_dt IS NOT NULL AND date_trunc('month', s.co_dt) <= s.m THEN 'co'
             WHEN s.bucket = 0 THEN 'cur'
             ELSE 'dq'
           END AS state
    FROM cohort c
    JOIN monthly s
      ON c.extnl_acct_id = s.extnl_acct_id
     AND s.m >= c.start_m
     AND s.m <= date_add('month', 9, c.start_m)
)
SELECT vintage,
       month_on_book,
       count(*) AS accounts_visible,
       round(100.0 * count_if(state = 'cur') / count(*), 1) AS pct_current,
       round(100.0 * count_if(state = 'dq') / count(*), 1) AS pct_delinquent,
       round(100.0 * count_if(state = 'co') / count(*), 1) AS pct_charged_off
FROM path
GROUP BY 1, 2
ORDER BY 1, 2
