-- Tier 5 | Calibration companion: call legs by initiation method, by month (full history)
-- The method mix shifted hard over time (recent outbound volume is a fraction
-- of its lifetime share, and queue-transfer legs appear/disappear). This shows
-- WHEN each method's era starts and ends, so no window is read across a mix
-- break. Companion to f0_period_calibration (which is inbound-only).
SELECT cast(date_trunc('month', "date") AS date) AS call_month,
       coalesce(initiationmethod, '(blank)') AS initiationmethod,
       count(*) AS calls
FROM "contactcenter_bdp_db"."call"
WHERE effdt < cast(date_add('day', -1, current_date) AS varchar)
GROUP BY 1, 2
ORDER BY 1, 2
