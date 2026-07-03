-- Tier 5 | Sensitivity: the no-payment gate at 7 / 14 / 30 / 60 days (W3)
-- The funnel's leakage pool depends on one assumption: 'no payment within 30
-- days'. This reruns that gate at 7, 14, 30, and 60 days on the same episodes
-- (delinquent, transcribed, payment/plan language - the f1 chain through
-- stage g). The 7-day read is the tight self-cure bracket the value maths
-- wants; if the pool halves at 60 days, slow payers are being counted as
-- leaked; if it barely moves, the 30-day read is robust.
-- SINGLE-PASS by construction: Athena inlines a WITH-CTE at every reference,
-- so the four windows are computed as one aggregate row and unpivoted -
-- never four UNION arms re-running the pipeline (the first cut of this query
-- read 272B rows that way and hit the 30-minute cap).
-- Account slice is 15 months (f8 needs a 60-day payment runway, not the
-- funnel's 8-month charge-off runway); the already-charged-off exclusion
-- reads the call-month row's own chrgoff_dt. Payment checks read the call
-- month's and following months' snapshots (raw-payment gate; f1 carries the
-- autopay/NSF-clean gate).
-- Same pinned W3 window and dedup as f1_funnel_waterfall.
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
      AND date(date_parse(eff_dt, '%Y%m%d')) < DATE '2025-10-01'
      AND eff_dt >= '20240701' AND eff_dt < '20251001'
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, min(co_dt) AS co_dt,
           max(pay_dt) AS pay_dt
    FROM snap GROUP BY 1, 2
),
monthly2 AS (
    SELECT extnl_acct_id, m, bucket, co_dt, pay_dt AS pay0,
           lead(pay_dt, 1) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS pay1,
           lead(pay_dt, 2) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS pay2
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
tx AS (
    SELECT t.contactid,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND regexp_like(lower(t.content),
                        'pay|paid|payment|settle|payment plan|arrangement|work something out'))
               AS pay_utts
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM episodes) d ON t.contactid = d.contactid
     AND t.effdt >= '2024-07-01' AND t.effdt < '2025-07-02'
    GROUP BY 1
),
flags AS (
    SELECT
        CASE WHEN (s.pay0 IS NOT NULL AND s.pay0 >= e.call_dt
                   AND s.pay0 <= date_add('day', 7, e.call_dt))
               OR (s.pay1 IS NOT NULL AND s.pay1 >= e.call_dt
                   AND s.pay1 <= date_add('day', 7, e.call_dt))
             THEN 1 ELSE 0 END AS paid_7,
        CASE WHEN (s.pay0 IS NOT NULL AND s.pay0 >= e.call_dt
                   AND s.pay0 <= date_add('day', 14, e.call_dt))
               OR (s.pay1 IS NOT NULL AND s.pay1 >= e.call_dt
                   AND s.pay1 <= date_add('day', 14, e.call_dt))
             THEN 1 ELSE 0 END AS paid_14,
        CASE WHEN (s.pay0 IS NOT NULL AND s.pay0 >= e.call_dt
                   AND s.pay0 <= date_add('day', 30, e.call_dt))
               OR (s.pay1 IS NOT NULL AND s.pay1 >= e.call_dt
                   AND s.pay1 <= date_add('day', 30, e.call_dt))
             THEN 1 ELSE 0 END AS paid_30,
        CASE WHEN (s.pay0 IS NOT NULL AND s.pay0 >= e.call_dt
                   AND s.pay0 <= date_add('day', 60, e.call_dt))
               OR (s.pay1 IS NOT NULL AND s.pay1 >= e.call_dt
                   AND s.pay1 <= date_add('day', 60, e.call_dt))
               OR (s.pay2 IS NOT NULL AND s.pay2 >= e.call_dt
                   AND s.pay2 <= date_add('day', 60, e.call_dt))
             THEN 1 ELSE 0 END AS paid_60
    FROM episodes e
    JOIN monthly2 s
      ON e.acct_key = trim(cast(s.extnl_acct_id AS varchar))
     AND e.call_month = cast(s.m AS date)
    JOIN tx t
      ON e.contactid = t.contactid
    WHERE s.bucket >= 1
      AND (s.co_dt IS NULL OR s.co_dt > e.call_dt)
      AND t.pay_utts > 0
),
agg AS (
    SELECT count(*) AS n,
           count_if(paid_7 = 0) AS no7,
           count_if(paid_14 = 0) AS no14,
           count_if(paid_30 = 0) AS no30,
           count_if(paid_60 = 0) AS no60
    FROM flags
)
SELECT w.pay_window_days,
       a.n AS intent_episodes,
       w.no_pay AS episodes_no_payment,
       round(100.0 * w.no_pay / greatest(a.n, 1), 1) AS pct_no_payment
FROM agg a
CROSS JOIN UNNEST(
    ARRAY[7, 14, 30, 60],
    ARRAY[a.no7, a.no14, a.no30, a.no60]
) AS w (pay_window_days, no_pay)
ORDER BY 1
