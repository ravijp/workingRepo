-- Tier 6 | Capture-gate contamination: autopay-dated and NSF-marked payments
-- RUN BEFORE QUOTING THE FUNNEL. The capture gate counts a payment near the
-- call; but a payment that is just autopay firing is not a capture, and a
-- bounced (NSF) payment is not money. This sizes both contaminations among
-- delinquent account-months with a payment, by bucket - the correction band
-- for every payment-after-call number, and the empirical case for the
-- autopay/NSF-clean gate the f1 headline applies.
-- Doubles as the column probe: if atmtc_paymt_last_dt / nsf_last_paymt_dt do
-- not exist in this copy, THIS query errors alone and f1's fallback is the
-- documented route. Last 6 complete account months.
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
           coalesce(try_cast(atmtc_paymt_last_dt AS date),
                    try(cast(date_parse(atmtc_paymt_last_dt, '%d%b%Y') AS date))) AS auto_dt,
           coalesce(try_cast(nsf_last_paymt_dt AS date),
                    try(cast(date_parse(nsf_last_paymt_dt, '%d%b%Y') AS date))) AS nsf_dt
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
           max_by(auto_dt, eff_dt) AS auto_dt,
           max_by(nsf_dt, eff_dt) AS nsf_dt
    FROM snap GROUP BY 1, 2
),
delq AS (
    SELECT bucket,
           CASE WHEN pay_dt >= cast(m AS date)
                 AND pay_dt < date_add('month', 1, cast(m AS date))
                THEN 1 ELSE 0 END AS paid_in_month,
           CASE WHEN auto_dt IS NOT NULL AND auto_dt = pay_dt THEN 1 ELSE 0 END AS autopay_dated,
           CASE WHEN nsf_dt IS NOT NULL
                 AND nsf_dt >= cast(m AS date)
                 AND nsf_dt < date_add('month', 1, cast(m AS date))
                THEN 1 ELSE 0 END AS nsf_in_month
    FROM monthly
    WHERE bucket >= 1
)
SELECT bucket AS dpd_bucket,
       count(*) AS delinquent_account_months,
       count_if(paid_in_month = 1) AS with_payment_in_month,
       round(100.0 * count_if(paid_in_month = 1 AND autopay_dated = 1)
             / greatest(count_if(paid_in_month = 1), 1), 1) AS pct_payment_autopay_dated,
       round(100.0 * count_if(paid_in_month = 1 AND nsf_in_month = 1)
             / greatest(count_if(paid_in_month = 1), 1), 1) AS pct_payment_nsf_marked
FROM delq
GROUP BY 1
ORDER BY 1
