-- Tier 2 | Top queues for inbound calls (last 6 months)
-- Which queues take the volume (collections vs care vs fraud)?
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call")
SELECT coalesce(cast(queue AS varchar), '(blank)') AS queue,
       count(*) AS calls
FROM "contactcenter_bdp_db"."call", mx
WHERE "date" > date_add('month', -6, mx.d)
  AND initiationmethod = 'INBOUND'
GROUP BY 1
ORDER BY 2 DESC
LIMIT 15
