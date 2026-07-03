-- Validation | Entry balance by eventual outcome for the DQ1 vintage
-- Tests the assumption "outstanding balance at DQ1 = balance at risk".
-- Most DQ1 balance cures; this shows how entry dollars actually split by
-- outcome, i.e. the balance-weighted roll rate.
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
per_acct AS (
    SELECT c.extnl_acct_id,
           max(CASE WHEN s.co_dt IS NOT NULL AND date_trunc('month', s.co_dt) <= s.m
                    THEN 1 ELSE 0 END) AS charged_off,
           max(CASE WHEN s.bucket = 0 AND s.m > c.start_m THEN 1 ELSE 0 END) AS ever_cured,
           max(CASE WHEN s.m = c.start_m THEN s.bal END) AS entry_bal
    FROM cohort c
    JOIN monthly s
      ON c.extnl_acct_id = s.extnl_acct_id
     AND s.m >= c.start_m
    GROUP BY 1
)
SELECT CASE
         WHEN charged_off = 1 THEN 'c. charged off in window'
         WHEN ever_cured = 1 THEN 'a. cured (back to current)'
         ELSE 'b. still delinquent'
       END AS outcome,
       count(*) AS accounts,
       round(sum(entry_bal), 0) AS entry_balance,
       round(100.0 * sum(entry_bal) / sum(sum(entry_bal)) OVER (), 1) AS pct_of_entry_balance
FROM per_acct
GROUP BY 1
ORDER BY 1
