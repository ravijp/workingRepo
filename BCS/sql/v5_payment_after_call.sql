-- Validation | Payment-after-call proxy, by delinquency bucket (last 6 months)
-- The capture / leakage proxy at scale: of inbound calls from delinquent
-- accounts, what share shows NO payment within 30 days of the call?
-- Payment read from the NEXT month-end snapshot's last-payment date, so it is
-- a proxy, not a ledger join. If paymt_last_dt fails to parse as a date the
-- pct columns go null: check the parse before trusting the numbers.
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
           END AS bucket,
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(paymt_last_dt, '%d%b%Y') AS date))) AS pay_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN mx
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) > date_add('month', -8, mx.d)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, max(pay_dt) AS pay_dt
    FROM snap GROUP BY 1, 2
),
inb AS (
    SELECT acctid, "date" AS call_dt,
           cast(date_trunc('month', "date") AS date) AS call_month
    FROM "contactcenter_bdp_db"."call"
    CROSS JOIN mx
    WHERE "date" > date_add('month', -6, mx.d)
      AND initiationmethod = 'INBOUND'
      AND acctid IS NOT NULL
),
j AS (
    SELECT s.bucket, i.call_dt, nxt.pay_dt
    FROM inb i
    JOIN monthly s
      ON trim(cast(i.acctid AS varchar)) = trim(cast(s.extnl_acct_id AS varchar))
     AND i.call_month = cast(s.m AS date)
    LEFT JOIN monthly nxt
      ON trim(cast(i.acctid AS varchar)) = trim(cast(nxt.extnl_acct_id AS varchar))
     AND cast(nxt.m AS date) = date_add('month', 1, i.call_month)
    WHERE s.bucket >= 1
)
SELECT bucket AS dpd_bucket,
       count(*) AS delinquent_inbound_calls,
       round(100.0 * count_if(pay_dt IS NOT NULL
                              AND pay_dt >= call_dt
                              AND pay_dt <= date_add('day', 30, call_dt)) / count(*), 1)
           AS pct_payment_within_30d,
       round(100.0 * count_if(pay_dt IS NULL
                              OR pay_dt < call_dt
                              OR pay_dt > date_add('day', 30, call_dt)) / count(*), 1)
           AS pct_no_payment_30d
FROM j
GROUP BY 1
ORDER BY 1
