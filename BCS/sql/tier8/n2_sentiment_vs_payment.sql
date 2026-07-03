-- Tier 8 | Does the platform's own NLP predict payment? (call-level sentiment vs outcome)
-- The transcripts already carry model-produced NLP: a call-level customer
-- sentiment score. This joins that score to the 30-day payment outcome on
-- delinquent-account inbound calls. If sentiment bands separate payment rates,
-- the precomputed NLP earns a slot as a funnel signal next to the lexicon;
-- if they do not, that is worth knowing before anyone builds on mood.
-- Call month anchored two months before the newest account month (complete
-- following month for the payment check).
WITH am AS (
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
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(paymt_last_dt, '%d%b%Y') AS date))) AS pay_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN am
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= cast(date_add('month', -2, am.d) AS date)
      AND date(date_parse(eff_dt, '%Y%m%d')) < cast(am.d AS date)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, max(pay_dt) AS pay_dt
    FROM snap GROUP BY 1, 2
),
delq AS (
    SELECT try_cast(c.overallcustomersentiment AS double) AS cs,
           CASE WHEN (s.pay_dt IS NOT NULL
                      AND s.pay_dt >= c."date"
                      AND s.pay_dt <= date_add('day', 30, c."date"))
                  OR (nxt.pay_dt IS NOT NULL
                      AND nxt.pay_dt >= c."date"
                      AND nxt.pay_dt <= date_add('day', 30, c."date"))
                THEN 1 ELSE 0 END AS paid_30d
    FROM "contactcenter_bdp_db"."call" c
    CROSS JOIN am
    JOIN monthly s
      ON trim(cast(c.acctid AS varchar)) = trim(cast(s.extnl_acct_id AS varchar))
     AND cast(s.m AS date) = cast(date_add('month', -2, am.d) AS date)
    LEFT JOIN monthly nxt
      ON trim(cast(c.acctid AS varchar)) = trim(cast(nxt.extnl_acct_id AS varchar))
     AND cast(nxt.m AS date) = cast(date_add('month', -1, am.d) AS date)
    WHERE cast(date_trunc('month', c."date") AS date)
          = cast(date_add('month', -2, am.d) AS date)
      AND c.initiationmethod = 'INBOUND'
      AND c.acctid IS NOT NULL
      AND s.bucket >= 1
)
SELECT CASE
         WHEN cs IS NULL THEN 'd. unscored'
         WHEN cs <= -1 THEN 'a. negative (score <= -1)'
         WHEN cs < 1 THEN 'b. neutral (-1 to 1)'
         ELSE 'c. positive (score >= 1)'
       END AS sentiment_band,
       count(*) AS delinquent_calls,
       round(100.0 * count_if(paid_30d = 1) / count(*), 1) AS pct_payment_within_30d,
       round(100.0 * count_if(paid_30d = 0) / count(*), 1) AS pct_no_payment_30d
FROM delq
GROUP BY 1
ORDER BY 1
