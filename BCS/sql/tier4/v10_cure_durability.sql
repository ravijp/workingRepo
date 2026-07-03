-- Validation | Cure durability: how many cures come back within 3 months?
-- A cure that re-delinquents is worth less than a durable one, and program
-- re-ages masquerade as cures. For each bucket-to-current transition, the
-- share re-delinquent (bucket 1+) within the following 3 months.
-- Cure months anchored so every cure keeps 3 complete follow-up months
-- inside the account copy's edge.
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
      AND date(date_parse(eff_dt, '%Y%m%d')) >= cast(date_add('month', -8, am.m1) AS date)
      AND date(date_parse(eff_dt, '%Y%m%d')) < cast(date_add('month', 1, am.m1) AS date)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket
    FROM snap GROUP BY 1, 2
),
seq AS (
    SELECT extnl_acct_id, m, bucket,
           lead(bucket, 1) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS b1,
           lead(m, 1) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS m1_next,
           lead(bucket, 2) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS b2,
           lead(bucket, 3) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS b3,
           lead(bucket, 4) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS b4
    FROM monthly
),
cures AS (
    SELECT s.bucket AS from_bucket,
           CASE WHEN greatest(coalesce(s.b2, 0), coalesce(s.b3, 0), coalesce(s.b4, 0)) >= 1
                THEN 1 ELSE 0 END AS redelinquent_3m
    FROM seq s
    CROSS JOIN am
    WHERE s.bucket >= 1
      AND s.b1 = 0
      AND s.m1_next = date_add('month', 1, s.m)
      AND s.m >= date_add('month', -8, am.m1)
      AND s.m <= date_add('month', -5, am.m1)
)
SELECT from_bucket,
       count(*) AS cures,
       sum(redelinquent_3m) AS redelinquent_3m,
       round(100.0 * sum(redelinquent_3m) / count(*), 1) AS pct_redelinquent_3m
FROM cures
GROUP BY 1
ORDER BY 1
