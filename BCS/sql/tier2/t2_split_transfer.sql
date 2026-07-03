-- Tier 2 | Transfer type on inbound calls (last 6 months)
-- How much inbound volume ends in a transfer, and of which kind?
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call")
SELECT coalesce(cast(transfertype AS varchar), '(blank)') AS transfertype,
       count(*) AS calls
FROM "contactcenter_bdp_db"."call", mx
WHERE "date" > date_add('month', -6, mx.d)
  AND initiationmethod = 'INBOUND'
GROUP BY 1
ORDER BY 2 DESC
LIMIT 10
