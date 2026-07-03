-- Tier 2 | Inbound calls by vendor (last 6 months)
-- How is inbound volume spread across servicing vendors?
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call")
SELECT coalesce(cast(vendor AS varchar), '(blank)') AS vendor,
       count(*) AS calls
FROM "contactcenter_bdp_db"."call", mx
WHERE "date" > date_add('month', -6, mx.d)
  AND initiationmethod = 'INBOUND'
GROUP BY 1
ORDER BY 2 DESC
LIMIT 15
