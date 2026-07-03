-- Tier 2 | Abandon and transfer rates on inbound calls, by month (last 6 months)
-- Are callers being lost to abandons or transfer loops, and is it getting worse?
WITH mx AS (SELECT date_trunc('month', max("date")) AS m1 FROM "contactcenter_bdp_db"."call")
SELECT cast(date_trunc('month', "date") AS date) AS month,
       count(*) AS inbound_calls,
       round(100.0 * count_if(try_cast(abandoned AS integer) = 1) / count(*), 2) AS pct_abandoned,
       round(100.0 * count_if(transfertype IS NOT NULL
                              AND transfertype <> 'Not A Transfer') / count(*), 2) AS pct_transferred
FROM "contactcenter_bdp_db"."call", mx
WHERE "date" >= date_add('month', -6, mx.m1)
  AND "date" < mx.m1
  AND initiationmethod = 'INBOUND'
GROUP BY 1
ORDER BY 1
