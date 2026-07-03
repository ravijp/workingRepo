-- Tier 7 | Does payment language predict payment? (one account month)
-- The lexicon validation: of delinquent-account inbound calls, compare the
-- 30-day payment rate for calls WITH customer payment/plan language vs calls
-- without it (and calls with no transcript at all). If the language flag does
-- not separate the payment rates, the transcript gate in the funnel is noise.
-- Call month anchored two months before the newest account month, so the
-- 30-day payment window sits inside a complete following month.
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
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(paymt_last_dt, '%d%b%Y') AS date))) AS pay_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN latest
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= date_add('month', -2, latest.d)
      AND date(date_parse(eff_dt, '%Y%m%d')) < latest.d
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, max(pay_dt) AS pay_dt
    FROM snap GROUP BY 1, 2
),
inb AS (
    SELECT c.contactid, c."date" AS call_dt,
           trim(cast(c.acctid AS varchar)) AS acct_key
    FROM "contactcenter_bdp_db"."call" c
    CROSS JOIN latest
    WHERE cast(date_trunc('month', c."date") AS date)
          = cast(date_add('month', -2, latest.d) AS date)
      AND c.initiationmethod = 'INBOUND'
      AND c.acctid IS NOT NULL
),
delq AS (
    SELECT i.contactid, i.call_dt, i.acct_key, nxt.pay_dt
    FROM inb i
    CROSS JOIN latest
    JOIN monthly s
      ON i.acct_key = trim(cast(s.extnl_acct_id AS varchar))
     AND cast(s.m AS date) = cast(date_add('month', -2, latest.d) AS date)
    LEFT JOIN monthly nxt
      ON i.acct_key = trim(cast(nxt.extnl_acct_id AS varchar))
     AND cast(nxt.m AS date) = cast(date_add('month', -1, latest.d) AS date)
    WHERE s.bucket >= 1
),
tx AS (
    SELECT t.contactid,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND regexp_like(lower(t.content),
                        'pay|paid|payment|settle|payment plan|arrangement|work something out'))
               AS pay_utts
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM delq) d
      ON t.contactid = d.contactid
    GROUP BY 1
),
flagged AS (
    SELECT CASE
             WHEN t.contactid IS NULL THEN 'c. no transcript'
             WHEN t.pay_utts > 0 THEN 'a. payment/plan language'
             ELSE 'b. no payment/plan language'
           END AS language_group,
           CASE WHEN d.pay_dt IS NOT NULL
                 AND d.pay_dt >= d.call_dt
                 AND d.pay_dt <= date_add('day', 30, d.call_dt)
                THEN 1 ELSE 0 END AS paid_30d
    FROM delq d
    LEFT JOIN tx t ON d.contactid = t.contactid
)
SELECT language_group,
       count(*) AS delinquent_calls,
       round(100.0 * count_if(paid_30d = 1) / count(*), 1) AS pct_payment_within_30d,
       round(100.0 * count_if(paid_30d = 0) / count(*), 1) AS pct_no_payment_30d
FROM flagged
GROUP BY 1
ORDER BY 1
