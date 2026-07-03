-- Tier 5 | Leaked episodes and their dollars, by call month (W3: 2024-07 .. 2025-06)
-- The dollar trend of the funnel end: per call month, the episodes that reached
-- 'no payment within 30 days' (leaked), the accounts and month-end balances
-- behind them, and how many of those accounts (and dollars) charged off within
-- 8 months of the call. An account that leaks in several months appears in
-- each of those months. Same pinned W3 window, dedup, and gates as f1.
WITH snap AS (
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
           try_cast(chrgoff_dt AS date) AS co_dt,
           try_cast(chrgoff_amt AS double) AS co_amt,
           try_cast(acct_bal_amt AS double) AS bal,
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(try_cast(paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS pay_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= DATE '2024-07-01'
      AND date(date_parse(eff_dt, '%Y%m%d')) < DATE '2026-03-01'
      AND eff_dt >= '20240701' AND eff_dt < '20260301'
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, min(co_dt) AS co_dt,
           max(co_amt) AS co_amt, max_by(bal, eff_dt) AS bal, max(pay_dt) AS pay_dt
    FROM snap GROUP BY 1, 2
),
monthly2 AS (
    SELECT extnl_acct_id, m, bucket, bal, pay_dt,
           min(co_dt) OVER (PARTITION BY extnl_acct_id) AS acct_co_dt,
           max(co_amt) OVER (PARTITION BY extnl_acct_id) AS acct_co_amt,
           lead(pay_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_pay_dt
    FROM monthly
),
inb AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key, contactid,
           "date" AS call_dt, initiationtimestamp
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2024-07-01' AND "date" < DATE '2025-07-01'
      AND effdt >= '2024-07-01' AND effdt < '2025-07-02'
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
           s.bal, s.acct_co_dt, s.acct_co_amt, s.pay_dt, s.next_pay_dt
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
     AND t.effdt >= '2024-07-01' AND t.effdt < '2025-07-02'
    GROUP BY 1
),
leaked AS (
    SELECT m.call_month, m.acct_key, m.bal, m.acct_co_amt,
           CASE WHEN m.acct_co_dt IS NOT NULL
                 AND m.acct_co_dt > m.call_dt
                 AND m.acct_co_dt <= date_add('month', 8, m.call_dt)
                THEN 1 ELSE 0 END AS co_8m
    FROM matched m
    JOIN tx t ON m.contactid = t.contactid
    WHERE t.pay_utts > 0
      AND NOT ((m.pay_dt IS NOT NULL
                AND m.pay_dt >= m.call_dt
                AND m.pay_dt <= date_add('day', 30, m.call_dt))
            OR (m.next_pay_dt IS NOT NULL
                AND m.next_pay_dt >= m.call_dt
                AND m.next_pay_dt <= date_add('day', 30, m.call_dt)))
),
per_am AS (
    SELECT call_month, acct_key,
           count(*) AS eps,
           max(bal) AS bal,
           max(co_8m) AS co,
           max(acct_co_amt) AS co_amt
    FROM leaked
    GROUP BY 1, 2
)
SELECT call_month,
       sum(eps) AS leaked_episodes,
       count(*) AS leaked_accounts,
       round(sum(bal), 0) AS leaked_balance,
       sum(co) AS chargeoff_accounts,
       round(sum(CASE WHEN co = 1 THEN co_amt END), 0) AS chargeoff_dollars
FROM per_am
GROUP BY 1
ORDER BY 1
