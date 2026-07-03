-- Tier 10 | Follow-through: leaked vs captured episodes, three months later (W3)
-- The outcome curve the value story rests on. Take delinquent, transcribed,
-- payment-intent episodes (the f1 chain through stage f) and split them by
-- whether a payment followed within 30 days (captured) or not (leaked). Then
-- read the account's position three months after the call: current, still
-- early-bucket, deep, charged off, or gone from the table. If captured
-- episodes sit visibly better at month three, capture is worth balance-shaped
-- money, not just one payment. Association, not causation - same caveat as
-- every caller split. Same pinned W3 window as f1.
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
           s.acct_co_dt, s.pay_dt, s.next_pay_dt
    FROM episodes e
    JOIN monthly2 s
      ON e.acct_key = trim(cast(s.extnl_acct_id AS varchar))
     AND e.call_month = cast(s.m AS date)
    WHERE s.bucket >= 1
      AND (s.acct_co_dt IS NULL OR s.acct_co_dt > e.call_dt)
),
tx AS (
    SELECT t.contactid,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND regexp_like(lower(t.content),
                        'pay|paid|payment|settle|payment plan|arrangement|work something out'))
               AS pay_utts
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM matched) d
      ON t.contactid = d.contactid
    GROUP BY 1
),
intent AS (
    SELECT m.acct_key, m.call_dt, m.call_month, m.acct_co_dt,
           CASE WHEN (m.pay_dt IS NOT NULL
                      AND m.pay_dt >= m.call_dt
                      AND m.pay_dt <= date_add('day', 30, m.call_dt))
                  OR (m.next_pay_dt IS NOT NULL
                      AND m.next_pay_dt >= m.call_dt
                      AND m.next_pay_dt <= date_add('day', 30, m.call_dt))
                THEN 1 ELSE 0 END AS captured
    FROM matched m
    JOIN tx t ON m.contactid = t.contactid
    WHERE t.pay_utts > 0
),
outcome AS (
    SELECT i.captured,
           CASE
             WHEN i.acct_co_dt IS NOT NULL
                  AND i.acct_co_dt <= date_add('month', 4, i.call_month) THEN 'co'
             WHEN m3.bucket IS NULL THEN 'gone'
             WHEN m3.bucket = 0 THEN 'cur'
             WHEN m3.bucket <= 3 THEN 'dq1_3'
             ELSE 'dq4_plus'
           END AS state_3m
    FROM intent i
    LEFT JOIN monthly2 m3
      ON i.acct_key = trim(cast(m3.extnl_acct_id AS varchar))
     AND cast(m3.m AS date) = date_add('month', 3, i.call_month)
)
SELECT CASE WHEN captured = 1 THEN 'a. captured (payment within 30d)'
            ELSE 'b. leaked (no payment within 30d)' END AS episode_group,
       count(*) AS intent_episodes,
       round(100.0 * count_if(state_3m = 'cur') / count(*), 1) AS pct_current_3m,
       round(100.0 * count_if(state_3m = 'dq1_3') / count(*), 1) AS pct_dq1_3_3m,
       round(100.0 * count_if(state_3m = 'dq4_plus') / count(*), 1) AS pct_dq4_plus_3m,
       round(100.0 * count_if(state_3m = 'co') / count(*), 1) AS pct_chargedoff_3m,
       round(100.0 * count_if(state_3m = 'gone') / count(*), 1) AS pct_not_visible_3m
FROM outcome
GROUP BY captured
ORDER BY 1
