-- Tier 11 | Within-cohort contrast v2: x4's episode-scored design on the ledger
-- Supersedes b16. b16 had two faults. Shape: it scored every January episode
-- against the account's single Feb-EOM position, so both outcome rows carried
-- the same accounts and the same distribution - no contrast was expressible.
-- Population: it required the PAID episodes to also pass the transcript +
-- intent-language gate (its intent CTE inner-joined tx and kept pay_utts > 0
-- before both_kinds), so an account whose paid-30d episode had no transcript
-- or no payment language dropped out - 168 accounts against a hard 217 floor
-- (b8 2,394 minus b14 2,177).
-- v2 mirrors x4's outcome logic: episodes of ledger accounts (Jan-31 EOM
-- bucket 1, b12 cleanup, per b14/b15) across January-March 2025, each episode
-- scored on ITS OWN next-month bucket move (month-MAX bucket, x4's grain),
-- contrast paid-30d vs no-pay-30d episodes within accounts that have both
-- kinds in the window. No transcript pass: episode kind is the payment gate
-- alone, so the paid side matches b14's captured rule. The gate stays the
-- cohort series' CLEAN gate (b8/b9: autopay- and NSF-dated payments do not
-- count); x4 itself used the raw gate - the two guards are not interchangeable.
WITH snap AS (
    SELECT extnl_acct_id,
           eff_dt,
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
                    try(cast(date_parse(try_cast(paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS pay_dt,
           coalesce(try_cast(atmtc_paymt_last_dt AS date),
                    try(cast(date_parse(try_cast(atmtc_paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS auto_dt,
           coalesce(try_cast(nsf_last_paymt_dt AS date),
                    try(cast(date_parse(try_cast(nsf_last_paymt_dt AS varchar), '%d%b%Y') AS date))) AS nsf_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250501'
),
monthly AS (
    SELECT extnl_acct_id, m,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           min(co_dt) AS co_dt,
           max(pay_dt) AS pay_dt,
           max(auto_dt) AS auto_dt,
           max(nsf_dt) AS nsf_dt
    FROM snap GROUP BY 1, 2
),
ledger AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key
    FROM monthly
    WHERE m = DATE '2025-01-01'
      AND eom_bucket = 1
      AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
),
monthly2 AS (
    SELECT extnl_acct_id, m, max_bucket, pay_dt, auto_dt, nsf_dt,
           min(co_dt) OVER (PARTITION BY extnl_acct_id) AS acct_co_dt,
           lead(max_bucket) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_bucket,
           lead(m) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_m,
           lead(pay_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_pay_dt,
           lead(auto_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_auto_dt,
           lead(nsf_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_nsf_dt
    FROM monthly
),
inb AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key, contactid,
           "date" AS call_dt, initiationtimestamp
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2025-01-01' AND "date" < DATE '2025-04-01'
      AND effdt >= '2025-01-01' AND effdt < '2025-04-02'
      AND coalesce(cast(producttype AS varchar), '') <> 'BUSINESS_CARD'
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
    SELECT l.acct_key, e.contactid, e.call_dt,
           s.max_bucket AS bucket, s.next_bucket,
           CASE WHEN
                  (s.pay_dt IS NOT NULL
                   AND s.pay_dt >= e.call_dt
                   AND s.pay_dt <= date_add('day', 30, e.call_dt)
                   AND (s.auto_dt IS NULL OR s.auto_dt <> s.pay_dt)
                   AND (s.nsf_dt IS NULL OR s.nsf_dt <> s.pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.pay_dt))
                OR
                  (s.next_pay_dt IS NOT NULL
                   AND s.next_pay_dt >= e.call_dt
                   AND s.next_pay_dt <= date_add('day', 30, e.call_dt)
                   AND (s.next_auto_dt IS NULL OR s.next_auto_dt <> s.next_pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.next_pay_dt))
                THEN 1 ELSE 0 END AS captured
    FROM ledger l
    JOIN episodes e ON e.acct_key = l.acct_key
    JOIN monthly2 s
      ON e.acct_key = trim(cast(s.extnl_acct_id AS varchar))
     AND e.call_month = cast(s.m AS date)
    WHERE s.max_bucket >= 1
      AND (s.acct_co_dt IS NULL OR s.acct_co_dt > e.call_dt)
      AND s.next_m = date_add('month', 1, s.m)
),
both_kinds AS (
    SELECT acct_key
    FROM matched
    GROUP BY 1
    HAVING max(captured) = 1 AND min(captured) = 0
)
SELECT CASE WHEN i.captured = 1 THEN 'a. paid-30d episodes (same accounts)'
            ELSE 'b. no-pay-30d episodes (same accounts)' END AS b16v2_episode_group,
       count(*) AS b16v2_episodes,
       count(DISTINCT i.acct_key) AS b16v2_accounts,
       round(100.0 * count_if(i.next_bucket = 0) / count(*), 1) AS b16v2_pct_current_next_month,
       round(100.0 * count_if(i.next_bucket = i.bucket) / count(*), 1) AS b16v2_pct_same_bucket_next_month,
       round(100.0 * count_if(i.next_bucket > i.bucket) / count(*), 1) AS b16v2_pct_deeper_next_month
FROM matched i
JOIN both_kinds b ON i.acct_key = b.acct_key
GROUP BY 1
ORDER BY 1
