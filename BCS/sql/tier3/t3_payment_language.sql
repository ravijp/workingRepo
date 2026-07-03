-- Tier 3 | Payment language in customer utterances (inbound calls, last 1 month)
-- In what share of transcribed inbound calls does the customer talk about paying,
-- a plan/settlement, hardship, or escalation? A first, crude read of caller intent.
-- Window kept to 1 month to bound the text scan.
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call"),
inb AS (
    SELECT contactid
    FROM "contactcenter_bdp_db"."call", mx
    WHERE "date" > date_add('month', -1, mx.d)
      AND initiationmethod = 'INBOUND'
),
cust AS (
    SELECT t.contactid, lower(t.content) AS content
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN inb ON t.contactid = inb.contactid
    WHERE t.participantid = 'CUSTOMER'
      AND t.content IS NOT NULL
),
per_call AS (
    SELECT contactid,
           count(*) AS utter,
           count_if(regexp_like(content, 'pay|paid|payment')) AS pay_n,
           count_if(regexp_like(content, 'settle|payment plan|arrangement|work something out')) AS plan_n,
           count_if(regexp_like(content, 'hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford')) AS hard_n,
           count_if(regexp_like(content, 'lawyer|attorney|dispute|complaint|supervisor')) AS esc_n
    FROM cust
    GROUP BY 1
)
SELECT count(*) AS calls_scanned,
       sum(utter) AS customer_utterances,
       round(100.0 * count_if(pay_n > 0) / count(*), 1) AS pct_calls_mentioning_payment,
       round(100.0 * count_if(plan_n > 0) / count(*), 1) AS pct_calls_mentioning_plan_or_settlement,
       round(100.0 * count_if(hard_n > 0) / count(*), 1) AS pct_calls_mentioning_hardship,
       round(100.0 * count_if(esc_n > 0) / count(*), 1) AS pct_calls_mentioning_escalation
FROM per_call
