-- Tier 2 | Utterances per transcribed call
-- How long are the conversations, measured in spoken turns?
WITH per_call AS (
    SELECT contactid, count(*) AS n
    FROM "contactcenter_bdp_db"."transcript"
    GROUP BY 1
)
SELECT CASE
         WHEN n <= 10  THEN 'a. 1-10'
         WHEN n <= 25  THEN 'b. 11-25'
         WHEN n <= 50  THEN 'c. 26-50'
         WHEN n <= 100 THEN 'd. 51-100'
         WHEN n <= 200 THEN 'e. 101-200'
         ELSE 'f. 200+'
       END AS turns_bucket,
       count(*) AS calls
FROM per_call
GROUP BY 1
ORDER BY 1
