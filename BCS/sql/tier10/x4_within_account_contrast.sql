-- Tier 10 | The within-account contrast: same account, captured vs leaked episode
-- The causality answer the caller/non-caller splits cannot give. Take accounts
-- that had BOTH a captured and a leaked intent episode inside the window (the
-- f1 chain through the language gate), and compare what the account looked
-- like the month after each episode type - within the same account, so the
-- who-calls selection effect cancels. If the after-capture month is visibly
-- better than the after-leak month for the same customers, capture moves the
-- account, not just one payment. Self-controlled, still not a randomized
-- trial (timing within the delinquency spell differs) - the strongest
-- non-experimental read these tables allow. Same pinned W3 window as f1.
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
           lead(bucket) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_bucket,
           lead(m) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_m,
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
           s.bucket, s.next_bucket, s.next_m, s.m, s.acct_co_dt,
           CASE WHEN (s.pay_dt IS NOT NULL
                      AND s.pay_dt >= e.call_dt
                      AND s.pay_dt <= date_add('day', 30, e.call_dt))
                  OR (s.next_pay_dt IS NOT NULL
                      AND s.next_pay_dt >= e.call_dt
                      AND s.next_pay_dt <= date_add('day', 30, e.call_dt))
                THEN 1 ELSE 0 END AS captured
    FROM episodes e
    JOIN monthly2 s
      ON e.acct_key = trim(cast(s.extnl_acct_id AS varchar))
     AND e.call_month = cast(s.m AS date)
    WHERE s.bucket >= 1
      AND (s.acct_co_dt IS NULL OR s.acct_co_dt > e.call_dt)
      AND s.next_m = date_add('month', 1, s.m)
),
tx AS (
    SELECT t.contactid,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND regexp_like(lower(t.content),
                        'pay|paid|payment|settle|payment plan|arrangement|work something out'))
               AS pay_utts
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM matched) d ON t.contactid = d.contactid
     AND t.effdt >= '2024-07-01' AND t.effdt < '2025-07-02'
    GROUP BY 1
),
intent AS (
    SELECT m.acct_key, m.captured, m.bucket, m.next_bucket
    FROM matched m
    JOIN tx t ON m.contactid = t.contactid
    WHERE t.pay_utts > 0
),
both_kinds AS (
    SELECT acct_key
    FROM intent
    GROUP BY 1
    HAVING max(captured) = 1 AND min(captured) = 0
)
SELECT CASE WHEN i.captured = 1 THEN 'a. captured episodes (same accounts)'
            ELSE 'b. leaked episodes (same accounts)' END AS episode_group,
       count(*) AS intent_episodes,
       count(DISTINCT i.acct_key) AS accounts,
       round(100.0 * count_if(i.next_bucket = 0) / count(*), 1) AS pct_current_next_month,
       round(100.0 * count_if(i.next_bucket > i.bucket) / count(*), 1) AS pct_deeper_next_month
FROM intent i
JOIN both_kinds b ON i.acct_key = b.acct_key
GROUP BY 1
ORDER BY 1
