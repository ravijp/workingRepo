-- Tier 1 | Account-id fill on call legs, by month (last 24 complete months)
-- Lifetime fill hides the trend: old months with poor fill weaken any
-- history-based cohort join. Anchored to the newest call date; the
-- in-progress month is excluded.
WITH mx AS (SELECT date_trunc('month', max("date")) AS m1 FROM "contactcenter_bdp_db"."call")
SELECT cast(date_trunc('month', "date") AS date) AS month,
       count(*) AS call_rows,
       round(100.0 * count_if(acctid IS NOT NULL) / count(*), 1) AS pct_with_acctid
FROM "contactcenter_bdp_db"."call", mx
WHERE "date" >= date_add('month', -24, mx.m1)
  AND "date" < mx.m1
GROUP BY 1
ORDER BY 1
