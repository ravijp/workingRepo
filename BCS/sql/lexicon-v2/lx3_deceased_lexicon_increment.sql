-- Tier 13 | Lexicon v2, January calibration: deceased-lexicon extension increments
-- Same population and chain as lx1 (b10's cohort, verbatim). Measures the
-- m1/m2 mined phrases the m4 five-phrase base list does NOT cover, each as
-- an INCREMENT: episodes matched by the candidate AND NOT by the base list.
-- Candidates (word-boundary anchored so 'he passed' does not fire inside
-- 'she passed'): has passed / she passed / he passed / on behalf / a death /
-- the death. An episode that also says 'passed away' or any other base
-- phrase is already routed and never counts as an increment.
-- Candidates join the v2 routing group ONLY if they pass the protocol
-- acceptance bar (support >= 40 class-b episodes AND capture lift >= 2x or
-- CO12 spread >= 10 points vs class base); below bar = recorded, not used.
-- Per-candidate rows OVERLAP (one episode can match several candidates);
-- only the 'any candidate' union row is additive against the base.
-- Tie-outs: the base rows reproduce lx1's block-B deceased cells per class
-- (episodes exact; classes b+c sum to 614). Increment rows are NEW
-- measurement — no recorded anchor; internal partition check only.
-- ONE transcript pass; aggregates only (b11's lesson). The UNNEST fan-out
-- runs over the small per-episode table, never over utterances.
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
               AS base_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content), '\bhas passed'))
               AS hp_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content), '\bshe passed'))
               AS sp_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content), '\bhe passed'))
               AS hep_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content), '\bon behalf'))
               AS ob_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content), '\ba death'))
               AS ad_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content), '\bthe death'))
               AS td_n
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM ep) d
      ON t.contactid = d.contactid
     AND t.effdt >= '2025-01-01' AND t.effdt < '2025-02-02'
    WHERE t.content IS NOT NULL
    GROUP BY 1
),
ep2 AS (
    SELECT e.acct_key, e.bridge_class, e.captured,
           coalesce(x.base_n, 0) > 0 AS base_f,
           coalesce(x.hp_n, 0) > 0 AS hp_f,
           coalesce(x.sp_n, 0) > 0 AS sp_f,
           coalesce(x.hep_n, 0) > 0 AS hep_f,
           coalesce(x.ob_n, 0) > 0 AS ob_f,
           coalesce(x.ad_n, 0) > 0 AS ad_f,
           coalesce(x.td_n, 0) > 0 AS td_f,
           (f.co_dt_future >= DATE '2025-01-01'
            AND f.co_dt_future < DATE '2025-09-01') AS co_8m,
           (f.co_dt_future >= DATE '2025-01-01'
            AND f.co_dt_future < DATE '2026-01-01') AS co_12m
    FROM ep e
    LEFT JOIN tx x ON e.contactid = x.contactid
    LEFT JOIN future_co f
      ON trim(cast(f.extnl_acct_id AS varchar)) = e.acct_key
),
labeled AS (
    SELECT u.lbl, e.bridge_class, e.acct_key, e.captured, e.co_8m, e.co_12m,
           CASE u.lbl
             WHEN 'a. base list (m4 five, lx1 tie-out)' THEN e.base_f
             WHEN 'b. has passed (increment)'  THEN e.hp_f  AND NOT e.base_f
             WHEN 'c. she passed (increment)'  THEN e.sp_f  AND NOT e.base_f
             WHEN 'd. he passed (increment)'   THEN e.hep_f AND NOT e.base_f
             WHEN 'e. on behalf (increment)'   THEN e.ob_f  AND NOT e.base_f
             WHEN 'f. a death (increment)'     THEN e.ad_f  AND NOT e.base_f
             WHEN 'g. the death (increment)'   THEN e.td_f  AND NOT e.base_f
             WHEN 'h. any candidate (union increment)'
               THEN (e.hp_f OR e.sp_f OR e.hep_f OR e.ob_f OR e.ad_f OR e.td_f)
                    AND NOT e.base_f
           END AS matched
    FROM ep2 e
    CROSS JOIN UNNEST(ARRAY[
        'a. base list (m4 five, lx1 tie-out)',
        'b. has passed (increment)',
        'c. she passed (increment)',
        'd. he passed (increment)',
        'e. on behalf (increment)',
        'f. a death (increment)',
        'g. the death (increment)',
        'h. any candidate (union increment)']) AS u (lbl)
)
SELECT lbl AS lx3_candidate,
       bridge_class AS lx3_bridge_class,
       count(*) AS lx3_episodes,
       count(DISTINCT acct_key) AS lx3_accounts,
       count_if(captured = 1) AS lx3_eps_paid30d,
       count_if(captured = 0) AS lx3_eps_no_pay30d,
       count(DISTINCT CASE WHEN coalesce(co_8m, false) THEN acct_key END) AS lx3_co8m_accounts,
       count(DISTINCT CASE WHEN coalesce(co_12m, false) THEN acct_key END) AS lx3_co12m_accounts,
       round(100.0 * sum(captured) / count(*), 1) AS lx3_pct_paid30d,
       round(100.0 * count(DISTINCT CASE WHEN coalesce(co_12m, false) THEN acct_key END)
             / count(DISTINCT acct_key), 1) AS lx3_pct_co12m
FROM labeled
WHERE matched
GROUP BY 1, 2
ORDER BY 1, 2
