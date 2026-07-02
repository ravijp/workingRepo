-- Tier 3 | Inbound calls by caller delinquency bucket (last 6 months)
-- How much inbound volume comes from delinquent accounts, and does caller mood
-- worsen with bucket depth? Bucket read at the newest account snapshot, which may
-- trail the call dates - treat as approximate.
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call"),
latest AS (SELECT max(eff_dt) AS d FROM "fmt_acct_dba"."fmt_acct_c"),
acct AS (
    SELECT extnl_acct_id,
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
           END AS dpd_bucket
    FROM "fmt_acct_dba"."fmt_acct_c", latest
    WHERE sfx_nbr = 0
      AND eff_dt = latest.d
),
inb AS (
    SELECT acctid, overallcustomersentiment AS cs
    FROM "contactcenter_bdp_db"."call", mx
    WHERE "date" > date_add('month', -6, mx.d)
      AND initiationmethod = 'INBOUND'
      AND acctid IS NOT NULL
)
SELECT a.dpd_bucket,
       count(*) AS calls,
       round(avg(try_cast(inb.cs AS double)), 3) AS avg_customer_sentiment
FROM inb
JOIN acct a
  ON trim(cast(inb.acctid AS varchar)) = trim(cast(a.extnl_acct_id AS varchar))
GROUP BY 1
ORDER BY 1
