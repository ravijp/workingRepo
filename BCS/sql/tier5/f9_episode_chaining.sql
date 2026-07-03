-- Tier 5 | Episode chaining: gap to the account's previous episode (W3)
-- The funnel treats each account-day as one episode. This measures how those
-- episodes cluster: the share arriving within 3 / 7 / 14 days of the previous
-- one decides whether multi-day chains should collapse into one episode (and
-- at what window) - the open episode-definition question, answered with data.
-- Same pinned W3 window and dedup as f1_funnel_waterfall.
WITH inb AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key, contactid,
           "date" AS call_dt, initiationtimestamp
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2024-07-01' AND "date" < DATE '2025-07-01'
      AND effdt >= '2024-07-01' AND effdt < '2025-07-02'
),
episodes AS (
    SELECT acct_key, call_dt
    FROM (
        SELECT acct_key, call_dt,
               row_number() OVER (PARTITION BY acct_key, call_dt
                                  ORDER BY initiationtimestamp) AS rn
        FROM inb
        WHERE acct_key IS NOT NULL AND acct_key <> ''
    )
    WHERE rn = 1
),
gaps AS (
    SELECT date_diff('day',
               lag(call_dt) OVER (PARTITION BY acct_key ORDER BY call_dt),
               call_dt) AS gap_days
    FROM episodes
)
SELECT CASE
         WHEN gap_days IS NULL THEN 'a. first episode in window'
         WHEN gap_days <= 3 THEN 'b. 1-3 days after previous'
         WHEN gap_days <= 7 THEN 'c. 4-7 days'
         WHEN gap_days <= 14 THEN 'd. 8-14 days'
         WHEN gap_days <= 30 THEN 'e. 15-30 days'
         ELSE 'f. 31+ days'
       END AS gap_band,
       count(*) AS episodes,
       round(100.0 * count(*) / sum(count(*)) OVER (), 1) AS pct_of_episodes
FROM gaps
GROUP BY 1
ORDER BY 1
