-- Tier 7 | Promise language and whether the payment followed (kept-promise proxy)
-- The promise-to-pay table lives in another system, but promises are audible.
-- This splits delinquent calls into future-dated promise language ('I'll pay
-- Friday', 'when I get paid'), payment talk without a promise, and no payment
-- talk - and reads the 30-day payment rate for each. Promise calls with no
-- payment after are the broken-promise proxy the capture story needs until
-- the operational promise data lands.
-- Call month anchored two months before the newest account month.
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
                    try(cast(date_parse(try_cast(paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS pay_dt
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
    SELECT c.contactid,
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
      AND c.effdt >= '2025-10-01' AND c.effdt < '2026-04-01'
      AND c.initiationmethod = 'INBOUND'
      AND c.acctid IS NOT NULL
      AND s.bucket >= 1
),
tx AS (
    SELECT t.contactid,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'i.ll pay|i will pay|going to pay|gonna pay|pay (on|by|this|next)|when i get paid|payday|after my paycheck'))
               AS promise_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'pay|paid|payment|settle|payment plan|arrangement|work something out'))
               AS pay_n
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM delq) d ON t.contactid = d.contactid
     AND t.effdt >= '2025-09-01' AND t.effdt < '2026-05-01'
    WHERE t.content IS NOT NULL
    GROUP BY 1
)
SELECT CASE
         WHEN x.promise_n > 0 THEN 'a. future-dated promise language'
         WHEN x.pay_n > 0 THEN 'b. payment talk, no promise'
         ELSE 'c. no payment talk'
       END AS promise_group,
       count(*) AS delinquent_calls,
       round(100.0 * count(*) / sum(count(*)) OVER (), 1) AS pct_of_calls,
       round(100.0 * count_if(d.paid_30d = 1) / count(*), 1) AS pct_payment_within_30d,
       round(100.0 * count_if(d.paid_30d = 0) / count(*), 1) AS pct_no_payment_30d
FROM delq d
JOIN tx x ON d.contactid = x.contactid
GROUP BY 1
ORDER BY 1
