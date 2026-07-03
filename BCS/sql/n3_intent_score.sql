-- Tier 8 | Can rules rank calls? A composed intent score vs the payment outcome
-- The pre-production question for any live transcript read: do simple,
-- deterministic signals COMBINED rank calls by capture likelihood? Per
-- delinquent-account inbound call, an additive score:
--   +1 customer payment language   +2 customer plan/settlement language
--   +1 customer raises payment before the agent
--   +1 customer sentiment ends positive (final third of customer turns)
--   -1 hardship language           -1 escalation language
-- If payment rate climbs with the score, a rules engine can already rank -
-- and a proper model read starts from a proven floor, not a hope.
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
    SELECT c.contactid,
           CASE WHEN nxt.pay_dt IS NOT NULL
                 AND nxt.pay_dt >= c."date"
                 AND nxt.pay_dt <= date_add('day', 30, c."date")
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
),
cust AS (
    SELECT t.contactid,
           lower(t.content) AS content,
           try_cast(t.beginmillis AS bigint) AS b,
           t.sentiment,
           ntile(3) OVER (PARTITION BY t.contactid
                          ORDER BY try_cast(t.beginmillis AS bigint)) AS cust_third
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM delq) d ON t.contactid = d.contactid
    WHERE t.participantid = 'CUSTOMER'
      AND t.content IS NOT NULL
),
cust_call AS (
    SELECT contactid,
           count_if(regexp_like(content, 'pay|paid|payment')) AS pay_n,
           count_if(regexp_like(content, 'settle|payment plan|arrangement|work something out')) AS plan_n,
           count_if(regexp_like(content, 'hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford')) AS hard_n,
           count_if(regexp_like(content, 'lawyer|attorney|dispute|complaint|supervisor')) AS esc_n,
           min(CASE WHEN regexp_like(content, 'pay|paid|payment') THEN b END) AS cust_first,
           count_if(cust_third = 3 AND sentiment = 'POSITIVE') AS final_pos,
           count_if(cust_third = 3 AND sentiment = 'NEGATIVE') AS final_neg
    FROM cust
    GROUP BY 1
),
agent_call AS (
    SELECT t.contactid,
           min(try_cast(t.beginmillis AS bigint)) AS agent_first
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM delq) d ON t.contactid = d.contactid
    WHERE t.participantid = 'AGENT'
      AND t.content IS NOT NULL
      AND regexp_like(lower(t.content), 'pay|paid|payment')
    GROUP BY 1
),
scored AS (
    SELECT d.paid_30d,
           (CASE WHEN p.pay_n > 0 THEN 1 ELSE 0 END)
         + (CASE WHEN p.plan_n > 0 THEN 2 ELSE 0 END)
         + (CASE WHEN p.cust_first IS NOT NULL
                  AND (a.agent_first IS NULL OR p.cust_first <= a.agent_first)
                 THEN 1 ELSE 0 END)
         + (CASE WHEN p.final_pos > p.final_neg THEN 1 ELSE 0 END)
         - (CASE WHEN p.hard_n > 0 THEN 1 ELSE 0 END)
         - (CASE WHEN p.esc_n > 0 THEN 1 ELSE 0 END) AS intent_score
    FROM delq d
    JOIN cust_call p ON d.contactid = p.contactid
    LEFT JOIN agent_call a ON d.contactid = a.contactid
)
SELECT intent_score,
       count(*) AS calls,
       round(100.0 * count_if(paid_30d = 1) / count(*), 1) AS pct_payment_within_30d
FROM scored
GROUP BY 1
ORDER BY 1
