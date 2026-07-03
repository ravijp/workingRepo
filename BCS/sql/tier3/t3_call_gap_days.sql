-- Tier 3 | Days between consecutive inbound calls per account (last 6 months)
-- The cadence of repeat calling: same-day and 1-3-day gaps read as friction
-- (unresolved on the first attempt); month-scale gaps read as cycle-driven
-- contact. Decides how tight an 'episode' window should be and sizes the
-- callback burden.
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call" WHERE effdt < cast(date_add('day', -1, current_date) AS varchar)),
inb AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key, "date" AS call_dt
    FROM "contactcenter_bdp_db"."call", mx
    WHERE "date" > date_add('month', -6, mx.d)
      AND effdt >= '2025-11-01' AND effdt < cast(date_add('day', -1, current_date) AS varchar)
      AND initiationmethod = 'INBOUND'
      AND acctid IS NOT NULL
),
gaps AS (
    SELECT date_diff('day',
               lag(call_dt) OVER (PARTITION BY acct_key ORDER BY call_dt),
               call_dt) AS gap_days
    FROM inb
)
SELECT CASE
         WHEN gap_days = 0 THEN 'a. same day'
         WHEN gap_days <= 3 THEN 'b. 1-3 days'
         WHEN gap_days <= 7 THEN 'c. 4-7 days'
         WHEN gap_days <= 14 THEN 'd. 8-14 days'
         WHEN gap_days <= 30 THEN 'e. 15-30 days'
         ELSE 'f. 31+ days'
       END AS gap_band,
       count(*) AS call_pairs,
       round(100.0 * count(*) / sum(count(*)) OVER (), 1) AS pct_of_pairs
FROM gaps
WHERE gap_days IS NOT NULL
GROUP BY 1
ORDER BY 1
