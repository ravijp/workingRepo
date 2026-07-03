-- Validation | Where on the DQ ladder do inbound calls actually land? (last 6 complete ACCOUNT months)
-- Tests the assumption "most inbound calls concentrate in DQ1". Joins each
-- inbound call to the caller's delinquency bucket in the SAME month (not the
-- latest snapshot). Window anchored to the ACCOUNT table's clock (its newest
-- complete month): call months past the account copy's edge cannot match and
-- used to silently drop out, which biased levels. Self-heals on refresh.
WITH am AS (
    SELECT date_add('month', -1,
               max(date_trunc('month', date(date_parse(eff_dt, '%Y%m%d'))))) AS m1
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
           END AS dpd_bucket
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN am
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= cast(date_add('month', -5, am.m1) AS date)
      AND date(date_parse(eff_dt, '%Y%m%d')) < cast(date_add('month', 1, am.m1) AS date)
),
monthly AS (
    SELECT extnl_acct_id, m, max(dpd_bucket) AS dpd_bucket
    FROM snap GROUP BY 1, 2
),
inb AS (
    SELECT acctid, cast(date_trunc('month', "date") AS date) AS call_month
    FROM "contactcenter_bdp_db"."call"
    CROSS JOIN am
    WHERE "date" >= cast(date_add('month', -5, am.m1) AS date)
      AND "date" < cast(date_add('month', 1, am.m1) AS date)
      AND effdt >= '2025-06-01' AND effdt < '2026-04-01'
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
