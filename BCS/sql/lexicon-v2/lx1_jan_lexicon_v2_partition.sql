-- Tier 13 | Lexicon v2, January calibration: the v1-to-v2 displacement matrix
-- One transcript pass over the DQ1 cohort's January episodes (bridge classes
-- a, b, c — b10's population, verbatim CTEs). Each episode gets BOTH its v1
-- language group (b10's priority ladder: promise > payment-talk-no-plan >
-- plan > hardship > dispute > none) AND its v2 group (the deceased-estate
-- lexicon, m4's measured five phrases, as a TOP-priority routing category
-- above the same ladder). Two stacked output blocks:
--   A. v1 marginals — must reproduce b10's 18 recorded cells EXACTLY
--      (episodes, accounts, pct paid-30d, pct CO12), and the count columns
--      re-derive b8 stages 2-5 by summation. Any drift = STOP and diagnose
--      before lx2/lx3/lx4 run.
--   B. v2 x v1 displacement — where the deceased routing pulls episodes
--      from. Classes b+c must reproduce m4 exactly (614 deceased episodes;
--      displacement 293 none / 223 pay-talk / 56 promise / 28 plan /
--      9 dispute / 5 hardship; v2 marginals 614/1,518/5,660/514/89/326/
--      3,823). Class a's deceased row is a NEW measurement (m4 never
--      scanned the cured class) — report it labeled, month-max grain,
--      never folded silently into the m4 story.
-- Capture gate: f1's clean gate (autopay/NSF excluded), verbatim from b10.
-- CO windows: forward chrgoff_dt scan, 8M < 2025-09-01, 12M < 2026-01-01.
-- ONE transcript pass; aggregates only; no window functions over utterances
-- (b11's lesson). Episodes without a transcript land in the 'none' group
-- (b10 convention) and are counted in lx1_eps_no_transcript.
WITH snap AS (
    SELECT extnl_acct_id,
           substr(eff_dt, 1, 6) AS ym,
           eff_dt,
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
           try_cast(chrgoff_dt AS date) AS co_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20241201' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           min(co_dt) AS co_dt
    FROM snap GROUP BY 1, 2
),
base AS (
    SELECT j.extnl_acct_id, j.max_bucket, j.eom_bucket, j.co_dt,
           p.max_bucket AS prev_max_bucket
    FROM (SELECT * FROM monthly WHERE ym = '202501') j
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202412') p
      ON j.extnl_acct_id = p.extnl_acct_id
),
classed AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           CASE
             WHEN max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0
                  AND eom_bucket = 0
               THEN 'a. month-MAX B1 entrant, cured by EOM (invisible to ASP)'
             WHEN max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0
                  AND eom_bucket >= 1
               THEN 'b. month-MAX B1 entrant, still DQ1 at EOM'
             WHEN eom_bucket = 1
                  AND NOT (max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0)
               THEN 'c. EOM bucket 1 stock, not a month-max-B1 entrant'
             ELSE 'd. other delinquent in month (month-MAX >= 2, EOM <> 1)'
           END AS bridge_class
    FROM base
    WHERE max_bucket >= 1
      AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
),
cohort AS (
    SELECT * FROM classed WHERE bridge_class NOT LIKE 'd.%'
),
future_co AS (
    SELECT extnl_acct_id, min(try_cast(chrgoff_dt AS date)) AS co_dt_future
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20260101'
      AND chrgoff_dt IS NOT NULL
    GROUP BY 1
),
pay_snap AS (
    SELECT extnl_acct_id, eff_dt,
           date_trunc('month', date(date_parse(eff_dt, '%Y%m%d'))) AS m,
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(try_cast(paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS pay_dt,
           coalesce(try_cast(atmtc_paymt_last_dt AS date),
                    try(cast(date_parse(try_cast(atmtc_paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS auto_dt,
           coalesce(try_cast(nsf_last_paymt_dt AS date),
                    try(cast(date_parse(try_cast(nsf_last_paymt_dt AS varchar), '%d%b%Y') AS date))) AS nsf_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250301'
),
pay_monthly AS (
    SELECT extnl_acct_id, m,
           max(pay_dt) AS pay_dt,
           max(auto_dt) AS auto_dt,
           max(nsf_dt) AS nsf_dt
    FROM pay_snap GROUP BY 1, 2
),
pay_monthly2 AS (
    SELECT extnl_acct_id, m, pay_dt, auto_dt, nsf_dt,
           lead(pay_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_pay_dt,
           lead(auto_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_auto_dt,
           lead(nsf_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_nsf_dt
    FROM pay_monthly
),
inb AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key, contactid,
           "date" AS call_dt, initiationtimestamp
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2025-01-01' AND "date" < DATE '2025-02-01'
      AND effdt >= '2025-01-01' AND effdt < '2025-02-02'
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
ep AS (
    SELECT c.acct_key, c.bridge_class, e.contactid, e.call_dt,
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
    FROM cohort c
    JOIN episodes e ON e.acct_key = c.acct_key
    LEFT JOIN pay_monthly2 s
      ON e.acct_key = trim(cast(s.extnl_acct_id AS varchar))
     AND e.call_month = cast(s.m AS date)
),
tx AS (
    SELECT t.contactid,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'passed away|death certificate|executor|deceased|calling on behalf'))
               AS deceased_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'pay|paid|payment|settle|payment plan|arrangement|work something out'))
               AS pay_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'settle|payment plan|arrangement|work something out'))
               AS plan_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford'))
               AS hard_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'dispute|not my charge|didn.t authorize|did not authorize|unauthorized|fraud|identity theft'))
               AS dispute_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'i.ll pay|i will pay|going to pay|gonna pay|pay (on|by|this|next)|when i get paid|payday|after my paycheck'))
               AS promise_n
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM ep) d
      ON t.contactid = d.contactid
     AND t.effdt >= '2025-01-01' AND t.effdt < '2025-02-02'
    WHERE t.content IS NOT NULL
    GROUP BY 1
),
ep2 AS (
    SELECT e.acct_key, e.bridge_class, e.captured,
           (x.contactid IS NOT NULL) AS has_tx,
           CASE
             WHEN coalesce(x.promise_n, 0) > 0 THEN 'a. future-dated promise'
             WHEN coalesce(x.pay_n, 0) > 0
                  AND coalesce(x.plan_n, 0) = 0 THEN 'b. payment talk, no promise'
             WHEN coalesce(x.plan_n, 0) > 0 THEN 'c. plan or settlement talk'
             WHEN coalesce(x.hard_n, 0) > 0 THEN 'd. hardship talk'
             WHEN coalesce(x.dispute_n, 0) > 0 THEN 'e. dispute or fraud talk'
             ELSE 'f. no payment-related language'
           END AS v1_group,
           CASE
             WHEN coalesce(x.deceased_n, 0) > 0 THEN 'a. deceased or estate'
             WHEN coalesce(x.promise_n, 0) > 0 THEN 'b. future-dated promise'
             WHEN coalesce(x.pay_n, 0) > 0
                  AND coalesce(x.plan_n, 0) = 0 THEN 'c. payment talk, no promise'
             WHEN coalesce(x.plan_n, 0) > 0 THEN 'd. plan or settlement talk'
             WHEN coalesce(x.hard_n, 0) > 0 THEN 'e. hardship talk'
             WHEN coalesce(x.dispute_n, 0) > 0 THEN 'f. dispute or fraud talk'
             ELSE 'g. no payment-related language'
           END AS v2_group,
           (f.co_dt_future >= DATE '2025-01-01'
            AND f.co_dt_future < DATE '2025-09-01') AS co_8m,
           (f.co_dt_future >= DATE '2025-01-01'
            AND f.co_dt_future < DATE '2026-01-01') AS co_12m
    FROM ep e
    LEFT JOIN tx x ON e.contactid = x.contactid
    LEFT JOIN future_co f
      ON trim(cast(f.extnl_acct_id AS varchar)) = e.acct_key
)
SELECT 'A. v1 marginal (b10 tie-out)' AS lx1_block,
       '(all v2 groups)' AS lx1_v2_group,
       v1_group AS lx1_v1_group,
       bridge_class AS lx1_bridge_class,
       count(*) AS lx1_episodes,
       count(DISTINCT acct_key) AS lx1_accounts,
       count_if(captured = 1) AS lx1_eps_paid30d,
       count_if(captured = 0) AS lx1_eps_no_pay30d,
       count_if(captured = 0 AND coalesce(co_8m, false)) AS lx1_eps_no_pay30d_co8m,
       count_if(NOT has_tx) AS lx1_eps_no_transcript,
       count(DISTINCT CASE WHEN coalesce(co_8m, false) THEN acct_key END) AS lx1_co8m_accounts,
       count(DISTINCT CASE WHEN coalesce(co_12m, false) THEN acct_key END) AS lx1_co12m_accounts,
       round(100.0 * sum(captured) / count(*), 1) AS lx1_pct_paid30d,
       round(100.0 * count(DISTINCT CASE WHEN coalesce(co_12m, false) THEN acct_key END)
             / count(DISTINCT acct_key), 1) AS lx1_pct_co12m
FROM ep2
GROUP BY 3, 4
UNION ALL
SELECT 'B. v2 x v1 displacement' AS lx1_block,
       v2_group AS lx1_v2_group,
       v1_group AS lx1_v1_group,
       bridge_class AS lx1_bridge_class,
       count(*) AS lx1_episodes,
       count(DISTINCT acct_key) AS lx1_accounts,
       count_if(captured = 1) AS lx1_eps_paid30d,
       count_if(captured = 0) AS lx1_eps_no_pay30d,
       count_if(captured = 0 AND coalesce(co_8m, false)) AS lx1_eps_no_pay30d_co8m,
       count_if(NOT has_tx) AS lx1_eps_no_transcript,
       count(DISTINCT CASE WHEN coalesce(co_8m, false) THEN acct_key END) AS lx1_co8m_accounts,
       count(DISTINCT CASE WHEN coalesce(co_12m, false) THEN acct_key END) AS lx1_co12m_accounts,
       round(100.0 * sum(captured) / count(*), 1) AS lx1_pct_paid30d,
       round(100.0 * count(DISTINCT CASE WHEN coalesce(co_12m, false) THEN acct_key END)
             / count(DISTINCT acct_key), 1) AS lx1_pct_co12m
FROM ep2
GROUP BY 2, 3, 4
ORDER BY 1, 2, 3, 4
