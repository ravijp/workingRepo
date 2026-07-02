-- Validation | Outcomes for callers vs non-callers in the same DQ1 vintage
-- The forward frame: take ALL accounts entering DQ1 in the vintage month
-- (11 months before the newest snapshot), split by "had an inbound call in the
-- first 3 months of delinquency", compare charge-off and cure rates.
-- This avoids conditioning on the charged-off population.
-- Association only, not causation: callers self-select.
WITH latest AS (
    SELECT max(date_trunc('month', date(date_parse(eff_dt, '%Y%m%d')))) AS d
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
           END AS bucket,
           try_cast(chrgoff_dt AS date) AS co_dt,
           try_cast(acct_bal_amt AS double) AS bal
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN latest
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) > date_add('month', -14, latest.d)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, min(co_dt) AS co_dt, max(bal) AS bal
    FROM snap GROUP BY 1, 2
),
entry AS (
    SELECT extnl_acct_id, m, bucket,
           lag(bucket) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS prev_bucket
    FROM monthly
),
cohort AS (
    SELECT e.extnl_acct_id, e.m AS start_m
    FROM entry e
    CROSS JOIN latest
    WHERE e.bucket = 1
      AND coalesce(e.prev_bucket, 0) = 0
      AND e.m = date_add('month', -11, latest.d)
),
outcome AS (
    SELECT c.extnl_acct_id,
           c.start_m,
           max(CASE WHEN s.co_dt IS NOT NULL AND date_trunc('month', s.co_dt) <= s.m
                    THEN 1 ELSE 0 END) AS charged_off,
           max(CASE WHEN s.bucket = 0 AND s.m > c.start_m THEN 1 ELSE 0 END) AS ever_cured
    FROM cohort c
    JOIN monthly s
      ON c.extnl_acct_id = s.extnl_acct_id
     AND s.m >= c.start_m
    GROUP BY 1, 2
),
entry_bal AS (
    SELECT c.extnl_acct_id, s.bal
    FROM cohort c
    JOIN monthly s
      ON c.extnl_acct_id = s.extnl_acct_id
     AND s.m = c.start_m
),
calls AS (
    SELECT DISTINCT trim(cast(acctid AS varchar)) AS acct_key
    FROM "contactcenter_bdp_db"."call"
    CROSS JOIN latest
    WHERE initiationmethod = 'INBOUND'
      AND acctid IS NOT NULL
      AND cast(date_trunc('month', "date") AS date)
          BETWEEN cast(date_add('month', -11, latest.d) AS date)
              AND cast(date_add('month', -8, latest.d) AS date)
)
SELECT CASE WHEN k.acct_key IS NOT NULL
            THEN 'a. inbound call in first 3 months'
            ELSE 'b. no inbound call' END AS caller_group,
       count(*) AS accounts,
       round(100.0 * sum(o.charged_off) / count(*), 1) AS pct_charged_off,
       round(100.0 * sum(o.ever_cured) / count(*), 1) AS pct_ever_back_to_current,
       round(avg(b.bal), 0) AS avg_entry_balance,
       round(sum(b.bal), 0) AS total_entry_balance
FROM outcome o
LEFT JOIN entry_bal b ON o.extnl_acct_id = b.extnl_acct_id
LEFT JOIN calls k
  ON trim(cast(o.extnl_acct_id AS varchar)) = k.acct_key
GROUP BY 1
ORDER BY 1
