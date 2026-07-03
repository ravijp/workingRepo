-- Tier 5 | Bias gate: transcript coverage by delinquency bucket (W2: 2025-10 .. 2026-02)
-- The funnel's transcript and language gates assume coverage does not depend on
-- how delinquent the caller is. If deep-bucket calls are transcribed less often,
-- the funnel undercounts exactly where the dollars are. Same-month bucket join
-- (the honest read), inbound calls with an account id only.
WITH snap AS (
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
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= DATE '2025-10-01'
      AND date(date_parse(eff_dt, '%Y%m%d')) < DATE '2026-03-01'
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket
    FROM snap GROUP BY 1, 2
),
inb AS (
    SELECT contactid, acctid,
           cast(date_trunc('month', "date") AS date) AS call_month
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2025-10-01' AND "date" < DATE '2026-03-01'
      AND acctid IS NOT NULL
),
t AS (SELECT DISTINCT contactid FROM "contactcenter_bdp_db"."transcript")
SELECT s.bucket AS dpd_bucket,
       count(*) AS calls,
       count(t.contactid) AS with_transcript,
       round(100.0 * count(t.contactid) / count(*), 1) AS pct_with_transcript
FROM inb
JOIN monthly s
  ON trim(cast(inb.acctid AS varchar)) = trim(cast(s.extnl_acct_id AS varchar))
 AND inb.call_month = cast(s.m AS date)
LEFT JOIN t ON inb.contactid = t.contactid
GROUP BY 1
ORDER BY 1
