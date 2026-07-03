-- Validation | Inbound vs outbound contact mix by delinquency bucket (last 6 complete ACCOUNT months)
-- Tests "inbound-only misses most contact": for each bucket, how much of the
-- contact volume is inbound vs outbound vs transfer? Keeps the outbound world
-- separate (as it should be) but sizes it.
-- Window anchored to the ACCOUNT table's clock (its newest complete month) so
-- the same-month join is complete in every window month. Self-heals on refresh.
-- Known finding: OUTBOUND legs carry no account id here, so the outbound
-- column reads ~0 - that absence is itself the result.
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
           END AS bucket
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN am
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= cast(date_add('month', -5, am.m1) AS date)
      AND date(date_parse(eff_dt, '%Y%m%d')) < cast(date_add('month', 1, am.m1) AS date)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket
    FROM snap GROUP BY 1, 2
),
calls AS (
    SELECT acctid, initiationmethod,
           cast(date_trunc('month', "date") AS date) AS call_month
    FROM "contactcenter_bdp_db"."call"
    CROSS JOIN am
    WHERE "date" >= cast(date_add('month', -5, am.m1) AS date)
      AND "date" < cast(date_add('month', 1, am.m1) AS date)
      AND acctid IS NOT NULL
)
SELECT s.bucket AS dpd_bucket,
       count(*) AS total_call_legs,
       count_if(c.initiationmethod = 'INBOUND') AS inbound,
       count_if(c.initiationmethod = 'OUTBOUND') AS outbound,
       count_if(c.initiationmethod NOT IN ('INBOUND', 'OUTBOUND')) AS other_method,
       round(1.0 * count_if(c.initiationmethod = 'OUTBOUND')
             / greatest(count_if(c.initiationmethod = 'INBOUND'), 1), 2) AS outbound_per_inbound
FROM calls c
JOIN monthly s
  ON trim(cast(c.acctid AS varchar)) = trim(cast(s.extnl_acct_id AS varchar))
 AND c.call_month = cast(s.m AS date)
GROUP BY 1
ORDER BY 1
