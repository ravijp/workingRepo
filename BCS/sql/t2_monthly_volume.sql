-- Tier 2 | Monthly call volume by initiation method (last 12 months of table data)
-- Is call volume trending up or down, and what is the inbound share?
-- Window is anchored to the newest call date in the table, not to today.
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call")
SELECT cast(date_trunc('month', "date") AS date) AS month,
       coalesce(initiationmethod, '(blank)') AS initiationmethod,
       count(*) AS calls
FROM "contactcenter_bdp_db"."call", mx
WHERE "date" > date_add('month', -12, mx.d)
GROUP BY 1, 2
ORDER BY 1
