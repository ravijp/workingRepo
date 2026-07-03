-- Tier 5 | The leakage funnel, end to end (W3: call months 2024-07 .. 2025-06)
-- The headline result. One waterfall, episodes + accounts + dollars per stage:
--   a. inbound call legs
--   b. legs with an account id (the acctid gap, made visible)
--   c. episodes = first inbound per account per day (the reference dedup)
--   d. episode matched to the caller's SAME-MONTH account snapshot
--   e. account delinquent that month (bucket 1+, not yet charged off)
--   f. episode has a transcript
--   g. customer payment / plan / settlement language in the transcript
--   h. no CLEAN payment within 30 days (see gate below)
--   i. account charged off within 8 months of the call
-- episodes_strict = the same gates with a STRICT intent lexicon (plan /
-- settlement / future-dated promise) instead of the loose one, so the
-- leakage pool reads as a band [strict, loose], not one number.
-- Capture gate: a payment counts if the call month's OR the next month's
-- last-payment date falls within 30 days of the call (one-sided next-month
-- checks misclassify same-month payers as leaked), EXCLUDING payments that
-- are autopay-dated or NSF-marked (a bounced or automatic payment is not a
-- capture). If the autopay/NSF columns are absent in this copy, the console
-- errors - paste the fallback file instead (raw-payment gate) and note it.
-- W3 ends 2025-06 so every episode keeps >= 8 account months of outcome
-- runway before the account copy's edge (newest snapshot 2026-03-07).
-- Balance = the account's month-end balance at its first matched episode,
-- held constant down the funnel, so the dollar column is a true waterfall.
-- Scope: business-card legs excluded; blank producttype kept (largely
-- consumer); no issuing-partner exclusions applied (values unverified) -
-- f4_scope_split sizes what that scope choice moves.
-- Cost note: if the console kills the account scan, restrict snap to
-- month-end snapshots (see the run plan) and rerun - buckets then read
-- month-end instead of month-max.
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
           try_cast(acct_bal_amt AS double) AS bal,
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(paymt_last_dt, '%d%b%Y') AS date))) AS pay_dt,
           coalesce(try_cast(atmtc_paymt_last_dt AS date),
                    try(cast(date_parse(atmtc_paymt_last_dt, '%d%b%Y') AS date))) AS auto_dt,
           coalesce(try_cast(nsf_last_paymt_dt AS date),
                    try(cast(date_parse(nsf_last_paymt_dt, '%d%b%Y') AS date))) AS nsf_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= DATE '2024-07-01'
      AND date(date_parse(eff_dt, '%Y%m%d')) < DATE '2026-03-01'
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, min(co_dt) AS co_dt,
           max_by(bal, eff_dt) AS bal,
           max(pay_dt) AS pay_dt,
           max(auto_dt) AS auto_dt,
           max(nsf_dt) AS nsf_dt
    FROM snap GROUP BY 1, 2
),
monthly2 AS (
    SELECT extnl_acct_id, m, bucket, bal, pay_dt, auto_dt, nsf_dt,
           min(co_dt) OVER (PARTITION BY extnl_acct_id) AS acct_co_dt,
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
      AND "date" >= DATE '2024-07-01' AND "date" < DATE '2025-07-01'
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
    SELECT e.acct_key, e.contactid, e.call_dt, e.call_month,
           s.bucket, s.bal, s.acct_co_dt,
           CASE WHEN s.bucket >= 1
                 AND (s.acct_co_dt IS NULL OR s.acct_co_dt > e.call_dt)
                THEN 1 ELSE 0 END AS is_delq,
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
               AS pay_utts,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND regexp_like(lower(t.content),
                        'settle|settlement|payment plan|arrangement|work something out|i.ll pay|i will pay|going to pay|gonna pay|when i get paid|payday'))
               AS strict_utts
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM matched WHERE is_delq = 1) d
      ON t.contactid = d.contactid
    GROUP BY 1
),
ep AS (
    SELECT m.acct_key,
           min_by(m.bal, CASE WHEN m.bal IS NOT NULL THEN m.call_dt END)
               OVER (PARTITION BY m.acct_key) AS acct_bal,
           CASE
             WHEN m.bucket IS NULL THEN 2
             WHEN m.is_delq = 0 THEN 3
             WHEN t.contactid IS NULL THEN 4
             WHEN t.pay_utts = 0 THEN 5
             WHEN m.captured = 1 THEN 6
             WHEN m.acct_co_dt IS NULL
                  OR m.acct_co_dt > date_add('month', 8, m.call_dt) THEN 7
             ELSE 8
           END AS deepest,
           CASE
             WHEN m.bucket IS NULL THEN 2
             WHEN m.is_delq = 0 THEN 3
             WHEN t.contactid IS NULL THEN 4
             WHEN t.strict_utts = 0 THEN 5
             WHEN m.captured = 1 THEN 6
             WHEN m.acct_co_dt IS NULL
                  OR m.acct_co_dt > date_add('month', 8, m.call_dt) THEN 7
             ELSE 8
           END AS deepest_strict
    FROM matched m
    LEFT JOIN tx t ON m.contactid = t.contactid
),
exploded AS (
    SELECT e.acct_key, e.acct_bal, s.stage_no,
           CASE WHEN e.deepest_strict >= s.stage_no THEN 1 ELSE 0 END AS strict_ok
    FROM ep e
    CROSS JOIN UNNEST(sequence(2, e.deepest)) AS s (stage_no)
),
acct_stage AS (
    SELECT stage_no, acct_key, count(*) AS eps,
           sum(strict_ok) AS eps_strict, max(acct_bal) AS bal
    FROM exploded GROUP BY 1, 2
)
SELECT 'a. inbound call legs' AS stage,
       count(*) AS episodes,
       CAST(NULL AS bigint) AS episodes_strict,
       CAST(NULL AS bigint) AS accounts,
       CAST(NULL AS double) AS balance_dollars
FROM inb
UNION ALL
SELECT 'b. legs with an account id',
       count(*),
       CAST(NULL AS bigint),
       count(DISTINCT acct_key),
       CAST(NULL AS double)
FROM inb
WHERE acct_key IS NOT NULL AND acct_key <> ''
UNION ALL
SELECT CASE stage_no
         WHEN 2 THEN 'c. episodes (first inbound per account per day)'
         WHEN 3 THEN 'd. matched to same-month account snapshot'
         WHEN 4 THEN 'e. delinquent in call month (bucket 1+)'
         WHEN 5 THEN 'f. has transcript'
         WHEN 6 THEN 'g. customer payment or plan language'
         WHEN 7 THEN 'h. no clean payment within 30 days'
         WHEN 8 THEN 'i. charged off within 8 months'
       END AS stage,
       sum(eps) AS episodes,
       sum(eps_strict) AS episodes_strict,
       count(*) AS accounts,
       round(sum(CASE WHEN stage_no >= 3 THEN bal END), 0) AS balance_dollars
FROM acct_stage
GROUP BY stage_no
ORDER BY 1
