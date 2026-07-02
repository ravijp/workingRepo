-- Tier 2 | Inbound calls by product type (last 6 months)
-- Which product lines drive inbound volume?
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call")
SELECT coalesce(cast(producttype AS varchar), '(blank)') AS producttype,
       count(*) AS calls
FROM "contactcenter_bdp_db"."call", mx
WHERE "date" > date_add('month', -6, mx.d)
  AND initiationmethod = 'INBOUND'
GROUP BY 1
ORDER BY 2 DESC
LIMIT 15
