-- Tier 9 | Transfers on the money path: delinquent vs current callers, and payment after
-- Transfers are separate legs chained by initialcontactid (lower bound - the
-- field is sparsely filled). This measures whether delinquent calls get
-- transferred more than current ones, and whether a transferred delinquent
-- call pays less afterwards - the transfer-drop leakage cause, sized.
-- Calls span the 5 complete account months before the newest complete month.
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
           END AS bucket,
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(try_cast(paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS pay_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN am
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= cast(date_add('month', -5, am.m1) AS date)
      AND date(date_parse(eff_dt, '%Y%m%d')) < cast(date_add('month', 1, am.m1) AS date)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, max(pay_dt) AS pay_dt
    FROM snap GROUP BY 1, 2
),
inb AS (
    SELECT contactid, acctid, "date" AS call_dt,
           cast(date_trunc('month', "date") AS date) AS call_month
    FROM "contactcenter_bdp_db"."call"
    CROSS JOIN am
    WHERE "date" >= cast(date_add('month', -5, am.m1) AS date)
      AND "date" < cast(am.m1 AS date)
      AND effdt >= '2025-06-01' AND effdt < '2026-04-01'
      AND initiationmethod = 'INBOUND'
      AND acctid IS NOT NULL
),
tl AS (
    SELECT DISTINCT initialcontactid
    FROM "contactcenter_bdp_db"."call"
    CROSS JOIN am
    WHERE "date" >= cast(date_add('month', -5, am.m1) AS date)
      AND "date" < cast(date_add('month', 1, am.m1) AS date)
      AND effdt >= '2025-06-01' AND effdt < '2026-04-01'
      AND initiationmethod IN ('TRANSFER', 'QUEUE_TRANSFER')
      AND initialcontactid IS NOT NULL
),
j AS (
    SELECT CASE WHEN s.bucket >= 1 THEN 'delinquent' ELSE 'current' END AS caller,
           CASE WHEN t.initialcontactid IS NOT NULL THEN 'transferred' ELSE 'not transferred' END AS xfer,
           CASE WHEN (s.pay_dt IS NOT NULL
                      AND s.pay_dt >= i.call_dt
                      AND s.pay_dt <= date_add('day', 30, i.call_dt))
                  OR (nxt.pay_dt IS NOT NULL
                      AND nxt.pay_dt >= i.call_dt
                      AND nxt.pay_dt <= date_add('day', 30, i.call_dt))
                THEN 1 ELSE 0 END AS paid
    FROM inb i
    JOIN monthly s
      ON trim(cast(i.acctid AS varchar)) = trim(cast(s.extnl_acct_id AS varchar))
     AND i.call_month = cast(s.m AS date)
    LEFT JOIN tl t ON i.contactid = t.initialcontactid
    LEFT JOIN monthly nxt
      ON trim(cast(i.acctid AS varchar)) = trim(cast(nxt.extnl_acct_id AS varchar))
     AND cast(nxt.m AS date) = date_add('month', 1, i.call_month)
)
SELECT CASE
         WHEN caller = 'current' AND xfer = 'not transferred' THEN 'a. current - not transferred'
         WHEN caller = 'current' THEN 'b. current - transferred'
         WHEN xfer = 'not transferred' THEN 'c. delinquent - not transferred'
         ELSE 'd. delinquent - transferred'
       END AS transfer_group,
       count(*) AS inbound_calls,
       round(100.0 * count_if(paid = 1) / count(*), 1) AS pct_payment_within_30d,
       round(100.0 * count_if(paid = 0) / count(*), 1) AS pct_no_payment_30d
FROM j
GROUP BY 1
ORDER BY 1
