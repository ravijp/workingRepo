-- Tier 2 | Handle-time profile of inbound handled calls (last 6 months)
-- How long does an inbound call take, end to end (talk, hold, wrap-up)?
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call")
SELECT count(*) AS handled_calls,
       round(avg(try_cast(totalhandletime AS double)), 0) AS avg_handle_s,
       round(approx_percentile(try_cast(totalhandletime AS double), 0.5), 0) AS median_handle_s,
       round(approx_percentile(try_cast(totalhandletime AS double), 0.9), 0) AS p90_handle_s,
       round(avg(try_cast(customerholdtime AS double)), 0) AS avg_hold_s,
       round(avg(try_cast(aftercontactworktime AS double)), 0) AS avg_wrapup_s
FROM "contactcenter_bdp_db"."call", mx
WHERE "date" > date_add('month', -6, mx.d)
  AND initiationmethod = 'INBOUND'
  AND try_cast(handled AS integer) = 1
