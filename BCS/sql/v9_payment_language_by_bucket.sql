-- Validation | Payment language: current vs delinquent callers (newest account month)
-- The unconditioned payment-talk rate is an upper bound (payment words are
-- common on ordinary service calls). This splits customer language by the
-- caller's same-month delinquency state, in the newest month the account
-- table covers.
WITH latest AS (
    SELECT max(date_trunc('month', date(date_parse(eff_dt, '%Y%m%d')))) AS m
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
           END AS bucket
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN latest
    WHERE sfx_nbr = 0
      AND date_trunc('month', date(date_parse(eff_dt, '%Y%m%d'))) = latest.m
),
acct AS (
    SELECT extnl_acct_id, max(bucket) AS bucket FROM snap GROUP BY 1
),
inb AS (
    SELECT c.contactid, a.bucket
    FROM "contactcenter_bdp_db"."call" c
    CROSS JOIN latest
    JOIN acct a
      ON trim(cast(c.acctid AS varchar)) = trim(cast(a.extnl_acct_id AS varchar))
    WHERE cast(date_trunc('month', c."date") AS date) = cast(latest.m AS date)
      AND c.initiationmethod = 'INBOUND'
      AND c.acctid IS NOT NULL
),
cust AS (
    SELECT t.contactid, i.bucket, lower(t.content) AS content
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN inb i ON t.contactid = i.contactid
    WHERE t.participantid = 'CUSTOMER'
      AND t.content IS NOT NULL
),
per_call AS (
    SELECT contactid,
           max(bucket) AS bucket,
           count_if(regexp_like(content, 'pay|paid|payment')) AS pay_n,
           count_if(regexp_like(content, 'settle|payment plan|arrangement|work something out')) AS plan_n,
           count_if(regexp_like(content, 'hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford')) AS hard_n,
           count_if(regexp_like(content, 'lawyer|attorney|dispute|complaint|supervisor')) AS esc_n
    FROM cust
    GROUP BY 1
)
SELECT CASE WHEN bucket >= 1 THEN 'b. delinquent (bucket 1+)'
            ELSE 'a. current (bucket 0)' END AS caller_group,
       count(*) AS calls_scanned,
       round(100.0 * count_if(pay_n > 0) / count(*), 1) AS pct_calls_mentioning_payment,
       round(100.0 * count_if(plan_n > 0) / count(*), 1) AS pct_calls_mentioning_plan_or_settlement,
       round(100.0 * count_if(hard_n > 0) / count(*), 1) AS pct_calls_mentioning_hardship,
       round(100.0 * count_if(esc_n > 0) / count(*), 1) AS pct_calls_mentioning_escalation
FROM per_call
GROUP BY 1
ORDER BY 1
