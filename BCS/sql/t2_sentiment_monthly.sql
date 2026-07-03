-- Tier 2 | Average call-level sentiment on inbound calls, by month (last 6 months)
-- How do customer and agent sentiment trend? (Sentiment scored on a subset of calls.)
WITH mx AS (SELECT date_trunc('month', max("date")) AS m1 FROM "contactcenter_bdp_db"."call")
SELECT cast(date_trunc('month', "date") AS date) AS month,
       round(avg(try_cast(overallcustomersentiment AS double)), 3) AS avg_customer_sentiment,
       round(avg(try_cast(overallagentsentiment AS double)), 3) AS avg_agent_sentiment,
       count_if(overallcustomersentiment IS NOT NULL) AS scored_calls
FROM "contactcenter_bdp_db"."call", mx
WHERE "date" >= date_add('month', -6, mx.m1)
  AND "date" < mx.m1
  AND initiationmethod = 'INBOUND'
GROUP BY 1
ORDER BY 1
