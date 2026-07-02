-- Tier 2 | Call duration from transcript timestamps
-- How many minutes of conversation does a transcribed call hold?
WITH per_call AS (
    SELECT contactid,
           max(try_cast(endmillis AS bigint)) / 60000.0 AS minutes
    FROM "contactcenter_bdp_db"."transcript"
    GROUP BY 1
)
SELECT count(*) AS transcribed_calls,
       round(avg(minutes), 1) AS avg_minutes,
       round(approx_percentile(minutes, 0.5), 1) AS median_minutes,
       round(approx_percentile(minutes, 0.9), 1) AS p90_minutes,
       round(max(minutes), 1) AS max_minutes
FROM per_call
