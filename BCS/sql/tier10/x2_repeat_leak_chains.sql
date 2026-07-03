-- Tier 10 | Repeat leaks: accounts that leaked more than once, and what became of them
-- One leaked episode can be bad luck; a chain of them is a process failure.
-- Per account: leaked episodes (delinquent + intent + no payment, the f1
-- stage-g definition) within 90 days of its FIRST leak, banded, and the share
-- charged off within 8 months of that first leak. If charge-off climbs with
-- the leak count, repeat leakage is the compounding story - and the earliest
-- leak is the intervention point. Same pinned W3 window as f1.
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
                    try(cast(date_parse(try_cast(paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS pay_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= DATE '2024-07-01'
      AND date(date_parse(eff_dt, '%Y%m%d')) < DATE '2026-03-01'
      AND eff_dt >= '20240701' AND eff_dt < '20260301'
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
    SELECT e.acct_key, e.contactid, e.call_dt,
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
     AND t.effdt >= '2024-07-01' AND t.effdt < '2025-07-02'
    GROUP BY 1
),
leaked AS (
    SELECT m.acct_key, m.call_dt, m.acct_co_dt
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
per_acct AS (
    SELECT acct_key,
           min(call_dt) AS first_leak,
           max(acct_co_dt) AS acct_co_dt
    FROM leaked
    GROUP BY 1
),
counted AS (
    SELECT p.acct_key,
           count(l.call_dt) AS leaks_90d,
           max(CASE WHEN p.acct_co_dt IS NOT NULL
                     AND p.acct_co_dt > p.first_leak
                     AND p.acct_co_dt <= date_add('month', 8, p.first_leak)
                    THEN 1 ELSE 0 END) AS co_8m
    FROM per_acct p
    JOIN leaked l
      ON l.acct_key = p.acct_key
     AND l.call_dt >= p.first_leak
     AND l.call_dt <= date_add('day', 90, p.first_leak)
    GROUP BY 1
)
SELECT CASE
         WHEN leaks_90d = 1 THEN 'a. 1 leaked episode'
         WHEN leaks_90d = 2 THEN 'b. 2 leaked episodes'
         WHEN leaks_90d <= 4 THEN 'c. 3-4 leaked episodes'
         ELSE 'd. 5+ leaked episodes'
       END AS leak_band,
       count(*) AS accounts,
       round(100.0 * count(*) / sum(count(*)) OVER (), 1) AS pct_of_accounts,
       round(100.0 * sum(co_8m) / count(*), 1) AS pct_chargedoff_8m
FROM counted
GROUP BY 1
ORDER BY 1
