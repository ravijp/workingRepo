-- Tier 5 | The funnel by call-month vintage (W3: 2024-07 .. 2025-06)
-- Stability check on f1: the same cumulative gates, one row per call month.
-- A funnel read off one pooled year can hide a drifting month; this shows
-- whether the stage-to-stage drops are stable vintage by vintage.
-- Columns are cumulative episode counts: matched >= delinquent >= ... >= chargeoff_8m.
-- Same windows, dedup, and gates as f1 (payment counts from the call month's
-- OR the next month's snapshot; raw-payment variant of the gate - f1's
-- headline additionally applies the autopay/NSF exclusion, s6 sizes the gap).
WITH snap AS (
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
           try_cast(chrgoff_dt AS date) AS co_dt,
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(paymt_last_dt, '%d%b%Y') AS date))) AS pay_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= DATE '2024-07-01'
      AND date(date_parse(eff_dt, '%Y%m%d')) < DATE '2026-03-01'
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, min(co_dt) AS co_dt,
           max(pay_dt) AS pay_dt
    FROM snap GROUP BY 1, 2
),
monthly2 AS (
    SELECT extnl_acct_id, m, bucket, pay_dt,
           min(co_dt) OVER (PARTITION BY extnl_acct_id) AS acct_co_dt,
           lead(pay_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_pay_dt
    FROM monthly
),
inb AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key, contactid,
           "date" AS call_dt, initiationtimestamp
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2024-07-01' AND "date" < DATE '2025-07-01'
),
episodes AS (
    SELECT acct_key, contactid, call_dt,
           cast(date_trunc('month', call_dt) AS date) AS call_month
    FROM (
        SELECT acct_key, contactid, call_dt,
               row_number() OVER (PARTITION BY acct_key, call_dt
                                  ORDER BY initiationtimestamp) AS rn
        FROM inb
        WHERE acct_key IS NOT NULL AND acct_key <> ''
    )
    WHERE rn = 1
),
matched AS (
    SELECT e.acct_key, e.contactid, e.call_dt, e.call_month,
           s.bucket, s.acct_co_dt, s.pay_dt, s.next_pay_dt,
           CASE WHEN s.bucket >= 1
                 AND (s.acct_co_dt IS NULL OR s.acct_co_dt > e.call_dt)
                THEN 1 ELSE 0 END AS is_delq
    FROM episodes e
    LEFT JOIN monthly2 s
      ON e.acct_key = trim(cast(s.extnl_acct_id AS varchar))
     AND e.call_month = cast(s.m AS date)
),
tx AS (
    SELECT t.contactid,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND regexp_like(lower(t.content),
                        'pay|paid|payment|settle|payment plan|arrangement|work something out'))
               AS pay_utts
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM matched WHERE is_delq = 1) d
      ON t.contactid = d.contactid
    GROUP BY 1
),
ep AS (
    SELECT m.call_month,
           CASE
             WHEN m.bucket IS NULL THEN 2
             WHEN m.is_delq = 0 THEN 3
             WHEN t.contactid IS NULL THEN 4
             WHEN t.pay_utts = 0 THEN 5
             WHEN (m.pay_dt IS NOT NULL
                   AND m.pay_dt >= m.call_dt
                   AND m.pay_dt <= date_add('day', 30, m.call_dt))
               OR (m.next_pay_dt IS NOT NULL
                   AND m.next_pay_dt >= m.call_dt
                   AND m.next_pay_dt <= date_add('day', 30, m.call_dt)) THEN 6
             WHEN m.acct_co_dt IS NULL
                  OR m.acct_co_dt > date_add('month', 8, m.call_dt) THEN 7
             ELSE 8
           END AS deepest
    FROM matched m
    LEFT JOIN tx t ON m.contactid = t.contactid
)
SELECT call_month,
       count(*) AS episodes,
       count_if(deepest >= 3) AS matched,
       count_if(deepest >= 4) AS delinquent,
       count_if(deepest >= 5) AS with_transcript,
       count_if(deepest >= 6) AS pay_language,
       count_if(deepest >= 7) AS no_payment_30d,
       count_if(deepest >= 8) AS chargeoff_8m
FROM ep
GROUP BY 1
ORDER BY 1
