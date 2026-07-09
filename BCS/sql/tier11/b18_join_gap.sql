-- Tier 11 | Join gap: the verified-only join hole on the target population,
-- weekly. ALL January 2025 inbound CALL ROWS (leg grain, not episodes, no
-- product exclusion), split verified-joined vs not using f4_match_by_auth's
-- logic: joined = acctid present AND resolving to the account master
-- (sfx_nbr = 0). Among joined calls, the call-day bucket flag comes from
-- b17's join: the account's latest daily snapshot on or before the call date
-- (snapshot floor 2024-12-01 - enough for any January call day). Rows: call
-- week x join status; column: call counts. Sanity expectation: the joined
-- share should sit in the C47 band (~72-79% acctid fill, ~21% of inbound
-- unjoinable) - this is the historical baseline for pilot-time joinability.
WITH inb AS (
    SELECT contactid, acctid, "date" AS call_dt
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2025-01-01' AND "date" < DATE '2025-02-01'
      AND effdt >= '2025-01-01' AND effdt < '2025-02-02'
),
acct AS (
    SELECT DISTINCT trim(cast(extnl_acct_id AS varchar)) AS acct_key
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
),
matched AS (
    SELECT i.contactid, i.call_dt,
           trim(cast(i.acctid AS varchar)) AS acct_key,
           (a.acct_key IS NOT NULL) AS joined
    FROM inb i
    LEFT JOIN acct a
      ON trim(cast(i.acctid AS varchar)) = a.acct_key
),
snap AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           eff_dt,
           date(date_parse(eff_dt, '%Y%m%d')) AS snap_dt,
           CASE
             WHEN past_due_271_up_amt  > 0 THEN 10
             WHEN past_due_241_270_amt > 0 THEN 9
             WHEN past_due_211_240_amt > 0 THEN 8
             WHEN past_due_181_210_amt > 0 THEN 7
             WHEN past_due_151_180_amt > 0 THEN 6
             WHEN past_due_121_150_amt > 0 THEN 5
             WHEN past_due_91_120_amt  > 0 THEN 4
             WHEN past_due_61_90_amt   > 0 THEN 3
             WHEN past_due_31_60_amt   > 0 THEN 2
             WHEN past_due_1_30_amt    > 0 THEN 1
             ELSE 0
           END AS bucket
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20241201' AND eff_dt < '20250201'
),
callday AS (
    SELECT m.acct_key, m.call_dt,
           max_by(s.bucket, s.eff_dt) AS callday_bucket
    FROM (SELECT DISTINCT acct_key, call_dt FROM matched WHERE joined) m
    JOIN snap s
      ON s.acct_key = m.acct_key
     AND s.snap_dt <= m.call_dt
    GROUP BY 1, 2
)
SELECT cast(date_trunc('week', m.call_dt) AS date) AS b18_call_week,
       CASE
         WHEN NOT m.joined THEN 'a. not verified-joined (no or unresolvable acctid)'
         WHEN c.callday_bucket >= 1 THEN 'b. joined, call-day bucket >= 1'
         ELSE 'c. joined, call-day bucket 0 or no snapshot'
       END AS b18_join_status,
       count(*) AS b18_calls
FROM matched m
LEFT JOIN callday c
  ON m.joined AND c.acct_key = m.acct_key AND c.call_dt = m.call_dt
GROUP BY 1, 2
ORDER BY 1, 2
