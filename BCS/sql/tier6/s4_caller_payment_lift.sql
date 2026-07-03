-- Tier 6 | Payment in-month or next: delinquent callers vs non-callers, by bucket
-- The self-cure baseline. Of delinquent account-months, what share shows a
-- payment dated in that month or the following one, split by whether the
-- account placed an inbound call that month? (Counting only the following
-- month misses same-month captures - a call on the 3rd collecting on the 5th.)
-- The non-caller column is the do-nothing baseline every capture claim must
-- beat; the caller minus non-caller gap is raw association (callers
-- self-select), not a causal lift.
-- Observation months: 6 account months ending two months before the newest
-- account month, so every observation has a complete following month.
WITH latest AS (
    SELECT max(date_trunc('month', date(date_parse(eff_dt, '%Y%m%d')))) AS d
    FROM "fmt_acct_dba"."fmt_acct_c" WHERE sfx_nbr = 0
),
snap AS (
    SELECT extnl_acct_id, eff_dt,
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
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(paymt_last_dt, '%d%b%Y') AS date))) AS pay_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN latest
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) > date_add('month', -9, latest.d)
),
monthly AS (
    SELECT extnl_acct_id, m,
           max(bucket) AS bucket,
           max_by(pay_dt, eff_dt) AS pay_dt
    FROM snap GROUP BY 1, 2
),
seq AS (
    SELECT extnl_acct_id, m, bucket, pay_dt,
           lead(pay_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_pay_dt,
           lead(m) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_m
    FROM monthly
),
obs AS (
    SELECT s.extnl_acct_id, s.m, s.bucket,
           CASE WHEN (s.pay_dt >= cast(s.m AS date)
                      AND s.pay_dt < date_add('month', 1, cast(s.m AS date)))
                  OR (s.next_pay_dt >= cast(s.m AS date)
                      AND s.next_pay_dt < date_add('month', 2, cast(s.m AS date)))
                THEN 1 ELSE 0 END AS paid_next_month
    FROM seq s
    CROSS JOIN latest
    WHERE s.bucket >= 1
      AND s.next_m = date_add('month', 1, s.m)
      AND s.m >= date_add('month', -7, latest.d)
      AND s.m <= date_add('month', -2, latest.d)
),
callers AS (
    SELECT DISTINCT trim(cast(acctid AS varchar)) AS acct_key,
           cast(date_trunc('month', "date") AS date) AS call_month
    FROM "contactcenter_bdp_db"."call"
    CROSS JOIN latest
    WHERE initiationmethod = 'INBOUND'
      AND acctid IS NOT NULL
      AND cast(date_trunc('month', "date") AS date)
          >= cast(date_add('month', -7, latest.d) AS date)
      AND cast(date_trunc('month', "date") AS date)
          <= cast(date_add('month', -2, latest.d) AS date)
),
flagged AS (
    SELECT o.bucket, o.paid_next_month,
           CASE WHEN k.acct_key IS NOT NULL THEN 1 ELSE 0 END AS called
    FROM obs o
    LEFT JOIN callers k
      ON trim(cast(o.extnl_acct_id AS varchar)) = k.acct_key
     AND cast(o.m AS date) = k.call_month
)
SELECT bucket AS dpd_bucket,
       count_if(called = 1) AS caller_account_months,
       round(100.0 * count_if(called = 1 AND paid_next_month = 1)
             / greatest(count_if(called = 1), 1), 1) AS caller_pct_paid,
       count_if(called = 0) AS noncaller_account_months,
       round(100.0 * count_if(called = 0 AND paid_next_month = 1)
             / greatest(count_if(called = 0), 1), 1) AS noncaller_pct_paid
FROM flagged
GROUP BY 1
ORDER BY 1
