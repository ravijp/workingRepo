-- Tier 6 | Payment size by delinquency bucket (last 6 complete account months)
-- When a delinquent account pays, how much? The per-capture dollar input:
-- an incremental captured payment is one cycle's payment, not the balance.
-- An account-month counts as paying when its month-end last-payment date
-- falls inside that month; the amount is the month-end last-payment amount
-- (a proxy: one payment per month is captured, the last one).
-- Anchored to the newest account month; the in-progress month is excluded.
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
                    try(cast(date_parse(paymt_last_dt, '%d%b%Y') AS date))) AS pay_dt,
           try_cast(paymt_last_amt AS double) AS pay_amt
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN latest
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= date_add('month', -6, latest.d)
      AND date(date_parse(eff_dt, '%Y%m%d')) < latest.d
),
monthly AS (
    SELECT extnl_acct_id, m,
           max(bucket) AS bucket,
           max_by(pay_dt, eff_dt) AS pay_dt,
           max_by(pay_amt, eff_dt) AS pay_amt
    FROM snap GROUP BY 1, 2
),
delq AS (
    SELECT bucket, pay_amt,
           CASE WHEN pay_dt >= cast(m AS date)
                 AND pay_dt < date_add('month', 1, cast(m AS date))
                THEN 1 ELSE 0 END AS paid_in_month
    FROM monthly
    WHERE bucket >= 1
)
SELECT bucket AS dpd_bucket,
       count(*) AS delinquent_account_months,
       count_if(paid_in_month = 1) AS with_payment_in_month,
       round(100.0 * count_if(paid_in_month = 1) / count(*), 1) AS pct_with_payment,
       round(avg(CASE WHEN paid_in_month = 1 THEN pay_amt END), 0) AS avg_payment,
       round(approx_percentile(CASE WHEN paid_in_month = 1 THEN pay_amt END, 0.5), 0) AS median_payment
FROM delq
GROUP BY 1
ORDER BY 1
