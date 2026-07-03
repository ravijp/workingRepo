-- Tier 2 | Monthly call volume by initiation method (last 12 complete months)
-- Is call volume trending up or down, and what is the inbound share?
-- Anchored to the newest call date; the in-progress month is excluded so the
-- last point is not an artificial cliff.
WITH mx AS (SELECT date_trunc('month', max("date")) AS m1 FROM "contactcenter_bdp_db"."call")
SELECT cast(date_trunc('month', "date") AS date) AS month,
       coalesce(initiationmethod, '(blank)') AS initiationmethod,
       count(*) AS calls
FROM "contactcenter_bdp_db"."call", mx
WHERE "date" >= date_add('month', -12, mx.m1)
  AND "date" < mx.m1
GROUP BY 1, 2
ORDER BY 1
