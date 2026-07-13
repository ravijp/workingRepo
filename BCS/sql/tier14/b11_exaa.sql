-- Tier 14 | EX-AA VARIANT of b11_cohort_signal_timing.
-- Exclusion: clnt_prdct_cd at the account's January EOM snapshot (eom_cpc,
-- max_by(clnt_prdct_cd, eff_dt) over the January row), matched against the
-- SAS AA+GM+Bronco code set (23 codes). Applied on the account population
-- CTE (classed), so cohort classes a/b/c (class d still dropped, unchanged)
-- are all ex-AA. NULL-safe form: NULL or blank cpc is kept as "others".
-- Expected tie-out: signals 1-3 still partition the ex-AA episode base per
-- class (same identity as the original); every signal's episode count
-- <= the corresponding original b11 cell.
-- Everything else is unchanged from b11_cohort_signal_timing.sql: the
-- bucket ladder, bridge-class definitions, cohort filter, the ONE
-- transcript pass (aggregates only), the seven signal definitions, and the
-- first-third-of-call-time logic.
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
           try_cast(chrgoff_dt AS date) AS co_dt,
           clnt_prdct_cd
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20241201' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           min(co_dt) AS co_dt,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc
    FROM snap GROUP BY 1, 2
),
base AS (
    SELECT j.extnl_acct_id, j.max_bucket, j.eom_bucket, j.co_dt, j.eom_cpc,
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
      AND (eom_cpc IS NULL OR trim(eom_cpc) = ''
           OR eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                               'AA3','AC3','AM3','AA4','AC4','AM4',
                               'BGC','BGM','CGM','GMR',
                               'FBS','IBS','U1C','U2C','U3C'))
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
           min(CASE WHEN t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                     AND regexp_like(lower(t.content),
                         'pay|paid|payment|settle|payment plan|arrangement|work something out')
                    THEN try_cast(t.beginmillis AS bigint) END) AS cust_first,
           min(CASE WHEN t.participantid = 'AGENT' AND t.content IS NOT NULL
                     AND regexp_like(lower(t.content),
                         'pay|paid|payment|settle|payment plan|arrangement|work something out')
                    THEN try_cast(t.beginmillis AS bigint) END) AS agent_first,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND regexp_like(lower(t.content), 'pay|paid|payment')) AS cust_pay_n,
           count_if(t.participantid = 'AGENT' AND t.content IS NOT NULL
                    AND regexp_like(lower(t.content),
                        'payment plan|arrangement|settle|work something out|hardship program|assistance program|payment program'))
               AS agent_offer_n,
           min(try_cast(t.beginmillis AS bigint)) AS call_start_ms,
           max(try_cast(t.beginmillis AS bigint)) AS call_end_ms
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM ep) d
      ON t.contactid = d.contactid
     AND t.effdt >= '2025-01-01' AND t.effdt < '2025-02-02'
    GROUP BY 1
),
ep2 AS (
    SELECT e.bridge_class, e.captured,
           x.cust_first, x.agent_first,
           coalesce(x.cust_pay_n, 0) AS cust_pay_n,
           coalesce(x.agent_offer_n, 0) AS agent_offer_n,
           coalesce(x.cust_first IS NOT NULL
                    AND x.call_end_ms > x.call_start_ms
                    AND (x.cust_first - x.call_start_ms) * 3
                        <= (x.call_end_ms - x.call_start_ms), false) AS intent_early
    FROM ep e
    LEFT JOIN tx x ON e.contactid = x.contactid
)
SELECT '1. customer raises payment first' AS b11_signal,
       bridge_class AS b11_bridge_class,
       count(*) AS b11_episodes,
       round(100.0 * sum(captured) / count(*), 1) AS b11_pct_captured
FROM ep2
WHERE NOT (cust_first IS NULL AND agent_first IS NULL)
  AND (agent_first IS NULL OR (cust_first IS NOT NULL AND cust_first <= agent_first))
GROUP BY 2
UNION ALL SELECT '2. agent raises payment first', bridge_class,
       count(*), round(100.0 * sum(captured) / count(*), 1)
FROM ep2
WHERE NOT (cust_first IS NULL AND agent_first IS NULL)
  AND NOT (agent_first IS NULL OR (cust_first IS NOT NULL AND cust_first <= agent_first))
GROUP BY 2
UNION ALL SELECT '3. no payment mention', bridge_class,
       count(*), round(100.0 * sum(captured) / count(*), 1)
FROM ep2
WHERE cust_first IS NULL AND agent_first IS NULL
GROUP BY 2
UNION ALL SELECT '4. intent in first third of call', bridge_class,
       count(*), round(100.0 * sum(captured) / count(*), 1)
FROM ep2
WHERE intent_early
GROUP BY 2
UNION ALL SELECT '5. intent later in call only', bridge_class,
       count(*), round(100.0 * sum(captured) / count(*), 1)
FROM ep2
WHERE cust_first IS NOT NULL AND NOT intent_early
GROUP BY 2
UNION ALL SELECT '6. customer intent + agent offer', bridge_class,
       count(*), round(100.0 * sum(captured) / count(*), 1)
FROM ep2
WHERE cust_pay_n > 0 AND agent_offer_n > 0
GROUP BY 2
UNION ALL SELECT '7. customer intent, no agent offer', bridge_class,
       count(*), round(100.0 * sum(captured) / count(*), 1)
FROM ep2
WHERE cust_pay_n > 0 AND agent_offer_n = 0
GROUP BY 2
ORDER BY 1, 2
