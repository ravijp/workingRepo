-- Tier 2 | Authentication outcome on inbound calls (last 6 months)
-- How often do callers fail authentication (a friction signal)?
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call")
SELECT coalesce(cast(authenticationstatus AS varchar), '(blank)') AS authenticationstatus,
       count(*) AS calls
FROM "contactcenter_bdp_db"."call", mx
WHERE "date" > date_add('month', -6, mx.d)
  AND initiationmethod = 'INBOUND'
GROUP BY 1
ORDER BY 2 DESC
LIMIT 10
