-- Tier 3 | Repeat calling by account (inbound, last 6 months)
-- Do accounts call once or repeatedly? Repeat calls signal friction / unresolved need.
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call"),
inb AS (
    SELECT acctid, count(*) AS n
    FROM "contactcenter_bdp_db"."call", mx
    WHERE "date" > date_add('month', -6, mx.d)
      AND initiationmethod = 'INBOUND'
      AND acctid IS NOT NULL
    GROUP BY 1
)
SELECT CASE
         WHEN n = 1  THEN 'a. 1 call'
         WHEN n = 2  THEN 'b. 2 calls'
         WHEN n = 3  THEN 'c. 3 calls'
         WHEN n <= 5 THEN 'd. 4-5 calls'
         WHEN n <= 10 THEN 'e. 6-10 calls'
         ELSE 'f. 11+ calls'
       END AS calls_per_account,
       count(*) AS accounts,
       sum(n) AS calls
FROM inb
GROUP BY 1
ORDER BY 1
