-- Tier 1 | Calls by initiation method
-- What share of call legs is inbound vs outbound vs transfer?
SELECT coalesce(initiationmethod, '(blank)') AS initiationmethod,
       count(*) AS calls
FROM "contactcenter_bdp_db"."call"
GROUP BY 1
ORDER BY 2 DESC
