-- Tier 11 | Snapshot cadence: what "month-end" means in this table
-- The Athena month-END read takes the LAST snapshot date per account in the
-- month, not the calendar end of month. This measures how many snapshot
-- dates an account carries in 202501 and where those last dates land. If
-- the last dates cluster on one end-of-month date, our month-END grain is
-- Ishant-comparable; if they scatter, "month-end" means different days for
-- different accounts and part of the gap can live here.
WITH snap AS (
    SELECT extnl_acct_id, eff_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250201'
),
per_acct AS (
    SELECT extnl_acct_id,
           count(DISTINCT eff_dt) AS n_snapshots,
           max(eff_dt) AS last_dt
    FROM snap GROUP BY 1
),
cadence AS (
    SELECT '1. snapshots per account in 202501' AS snapshot_metric,
           cast(n_snapshots AS varchar) AS metric_value,
           count(*) AS accounts_202501_b3
    FROM per_acct GROUP BY 2
),
last_dates AS (
    SELECT '2. last eff_dt in month' AS snapshot_metric,
           last_dt AS metric_value,
           count(*) AS accounts_202501_b3,
           row_number() OVER (ORDER BY count(*) DESC) AS rk
    FROM per_acct GROUP BY 2
)
SELECT snapshot_metric, metric_value, accounts_202501_b3 FROM cadence
UNION ALL
SELECT snapshot_metric, metric_value, accounts_202501_b3
FROM last_dates WHERE rk <= 15
ORDER BY 1, 2
