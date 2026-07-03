-- Tier 3 | Top queues for delinquent callers (same-month join, last 3 account months)
-- Do delinquent accounts ring collections queues or care queues? Decides the
-- queue filter for any collections-scoped call analysis. Window pinned to
-- the months the account table actually covers (same-month join).
WITH latest AS (
    SELECT max(date_trunc('month', date(date_parse(eff_dt, '%Y%m%d')))) AS m
    FROM "fmt_acct_dba"."fmt_acct_c" WHERE sfx_nbr = 0
),
snap AS (
    SELECT extnl_acct_id,
           date_trunc('month', date(date_parse(eff_dt, '%Y%m%d'))) AS m,
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
    CROSS JOIN latest
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) > date_add('month', -4, latest.m)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket
    FROM snap GROUP BY 1, 2
),
inb AS (
    SELECT acctid,
           coalesce(cast(queue AS varchar), '(blank)') AS queue,
           cast(date_trunc('month', "date") AS date) AS call_month
    FROM "contactcenter_bdp_db"."call"
    CROSS JOIN latest
    WHERE cast(date_trunc('month', "date") AS date)
          BETWEEN cast(date_add('month', -2, latest.m) AS date)
              AND cast(latest.m AS date)
      AND initiationmethod = 'INBOUND'
      AND acctid IS NOT NULL
)
SELECT inb.queue,
       count(*) AS delinquent_calls
FROM inb
JOIN monthly s
  ON trim(cast(inb.acctid AS varchar)) = trim(cast(s.extnl_acct_id AS varchar))
 AND inb.call_month = cast(s.m AS date)
WHERE s.bucket >= 1
GROUP BY 1
ORDER BY 2 DESC
LIMIT 15
