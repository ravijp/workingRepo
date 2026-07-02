-- Validation | Where on the DQ ladder do inbound calls actually land?
-- Tests the assumption "most inbound calls concentrate in DQ1".
-- Joins each inbound call to the caller's delinquency bucket in the SAME month
-- (not the latest snapshot), last 6 call months.
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call"),
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
           END AS dpd_bucket
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN mx
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) > date_add('month', -8, mx.d)
),
monthly AS (
    SELECT extnl_acct_id, m, max(dpd_bucket) AS dpd_bucket
    FROM snap GROUP BY 1, 2
),
inb AS (
    SELECT acctid, cast(date_trunc('month', "date") AS date) AS call_month
    FROM "contactcenter_bdp_db"."call"
    CROSS JOIN mx
    WHERE "date" > date_add('month', -6, mx.d)
      AND initiationmethod = 'INBOUND'
      AND acctid IS NOT NULL
)
SELECT s.dpd_bucket,
       count(*) AS inbound_calls,
       count(DISTINCT i.acctid) AS accounts,
       round(100.0 * count(*) / sum(count(*)) OVER (), 1) AS pct_of_matched_calls
FROM inb i
JOIN monthly s
  ON trim(cast(i.acctid AS varchar)) = trim(cast(s.extnl_acct_id AS varchar))
 AND i.call_month = cast(s.m AS date)
GROUP BY 1
ORDER BY 1
