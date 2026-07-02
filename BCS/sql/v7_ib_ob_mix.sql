-- Validation | Inbound vs outbound contact mix by delinquency bucket (last 6 months)
-- Tests "inbound-only misses most contact": for each bucket, how much of the
-- contact volume is inbound vs outbound vs transfer? Keeps the outbound world
-- separate (as it should be) but sizes it.
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
           END AS bucket
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN mx
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) > date_add('month', -8, mx.d)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket
    FROM snap GROUP BY 1, 2
),
calls AS (
    SELECT acctid, initiationmethod,
           cast(date_trunc('month', "date") AS date) AS call_month
    FROM "contactcenter_bdp_db"."call"
    CROSS JOIN mx
    WHERE "date" > date_add('month', -6, mx.d)
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
