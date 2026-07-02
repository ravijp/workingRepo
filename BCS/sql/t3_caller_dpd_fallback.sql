-- Tier 3 | FALLBACK for t3_caller_dpd: ladder capped at 181-210
-- Used automatically if the deeper past-due columns (211+) are absent.
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call"),
latest AS (SELECT max(eff_dt) AS d FROM "fmt_acct_dba"."fmt_acct_c"),
acct AS (
    SELECT extnl_acct_id,
           CASE
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
