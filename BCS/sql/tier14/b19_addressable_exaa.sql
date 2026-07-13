-- Tier 14 | NEW QUERY (b19): the addressable January call stream, ex-AA -
-- bucket-1-at-call episodes by v2 language group x captured.
-- Assembled from verified templates under the keeper's signed-off design
-- (2026-07-13): the episode base is b14's inb + episodes CTEs (they retain
-- contactid through the first-inbound-per-account-per-day dedup - b17's own
-- episodes CTE drops it, which is why this file could not be a b17
-- derivative); the ex-AA filter is b17_exaa's cpc_monthly + episodes_exaa
-- pattern verbatim; the bucket-at-call read lifts b17's as-of-call-day
-- snapshot join (callday), WITHOUT threading contactid into it - its output
-- joins back to the contactid-bearing episode rows on (acct_key, call_dt);
-- the capture gate is b14's pay_snap/pay_monthly/pay_monthly2/ep logic
-- verbatim; the language partition is lx1's v2 priority CASE verbatim
-- (deceased or estate first, then promise, payment talk, plan, hardship,
-- dispute, none). ONE transcript pass, aggregates only.
-- RESOURCE FIX (keeper, 2026-07-13, after "exhausted resources at this
-- scale factor"): v1 stacked b17's 8-month whole-book daily snapshot join
-- AND b14's whole-book payment window function in one query. Fix: the
-- episode CTEs moved to the TOP of the WITH chain, and both whole-book
-- account scans (snap, pay_snap) now semi-join to the January episode
-- accounts before anything heavy runs. Logic unchanged: accounts with no
-- January inbound episode contributed nothing to the output.
-- Omitted vs b17: the last_current_dt column and the spell CTE (days-since-
-- delinquent bands are not part of b19's output; nothing else consumes them).
-- Kept: episodes whose call-day snapshot shows bucket = 1 (bucket-1-at-call
-- only, per the signed-off spec) and b17's pre-2025 chrgoff_dt exclusion.
-- NULL-safe exclusion form: NULL or blank cpc is kept as "others".
-- Expected tie-outs (pre-registered): total episodes <= 32,712 (the original
-- b17's unfiltered 'a. bucket 1 at call' total; strictly less expected -
-- ex-AA plus construct differences). ABOVE 32,712 = STOP, route to the
-- keeper. Language-group x captured cells partition the total EXACTLY;
-- deceased episodes never appear inside intent groups (by construction of
-- the priority CASE); distinct accounts <= episodes in every cell.
WITH inb AS (
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
cpc_monthly AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250201'
      AND trim(cast(extnl_acct_id AS varchar)) IN (SELECT acct_key FROM episodes)
    GROUP BY 1
),
episodes_exaa AS (
    SELECT e.acct_key, e.contactid, e.call_dt, e.call_month
    FROM episodes e
    LEFT JOIN cpc_monthly c ON c.acct_key = e.acct_key
    WHERE (c.eom_cpc IS NULL OR trim(c.eom_cpc) = ''
           OR c.eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                                 'AA3','AC3','AM3','AA4','AC4','AM4',
                                 'BGC','BGM','CGM','GMR',
                                 'FBS','IBS','U1C','U2C','U3C'))
),
snap AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           eff_dt,
           date(date_parse(eff_dt, '%Y%m%d')) AS snap_dt,
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
      AND eff_dt >= '20240601' AND eff_dt < '20250201'
      AND trim(cast(extnl_acct_id AS varchar)) IN (SELECT acct_key FROM episodes_exaa)
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
      AND trim(cast(extnl_acct_id AS varchar)) IN (SELECT acct_key FROM episodes_exaa)
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
callday AS (
    SELECT e.acct_key, e.call_dt,
           max_by(s.bucket, s.eff_dt) AS callday_bucket,
           max_by(s.co_dt, s.eff_dt) AS callday_co_dt
    FROM episodes_exaa e
    JOIN snap s
      ON s.acct_key = e.acct_key
     AND s.snap_dt <= e.call_dt
    GROUP BY 1, 2
),
kept AS (
    SELECT acct_key, call_dt
    FROM callday
    WHERE callday_bucket = 1
      AND (callday_co_dt IS NULL OR callday_co_dt >= DATE '2025-01-01')
),
ep AS (
    SELECT e.acct_key, e.contactid,
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
    FROM kept k
    JOIN episodes_exaa e
      ON e.acct_key = k.acct_key
     AND e.call_dt = k.call_dt
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
    SELECT e.acct_key, e.captured,
           CASE
             WHEN coalesce(x.deceased_n, 0) > 0 THEN 'a. deceased or estate'
             WHEN coalesce(x.promise_n, 0) > 0 THEN 'b. future-dated promise'
             WHEN coalesce(x.pay_n, 0) > 0
                  AND coalesce(x.plan_n, 0) = 0 THEN 'c. payment talk, no promise'
             WHEN coalesce(x.plan_n, 0) > 0 THEN 'd. plan or settlement talk'
             WHEN coalesce(x.hard_n, 0) > 0 THEN 'e. hardship talk'
             WHEN coalesce(x.dispute_n, 0) > 0 THEN 'f. dispute or fraud talk'
             ELSE 'g. no payment-related language'
           END AS language_group
    FROM ep e
    LEFT JOIN tx x ON e.contactid = x.contactid
)
SELECT language_group AS b19_language_group,
       captured AS b19_captured,
       count(*) AS b19_episodes,
       count(DISTINCT acct_key) AS b19_accounts
FROM ep2
GROUP BY 1, 2
ORDER BY 1, 2
