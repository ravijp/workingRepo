-- Tier 14 | EX-AA VARIANT of lx4_funnel_v2_monthly (original in sql\lexicon-v2\).
-- Exclusion: clnt_prdct_cd per account-month (eom_cpc, max_by(clnt_prdct_cd,
-- eff_dt) within the same monthly GROUP BY that already produces bucket/
-- co_dt/pay_dt), matched against the SAS AA+GM+Bronco code set (23 codes).
-- Applied on the account side of the match (the delinquent-in-month
-- membership test in `matched`, via monthly2.eom_cpc), so is_delq = 1 only
-- for ex-AA account-months. NULL-safe form: NULL or blank cpc is kept as
-- "others". Transcript CTEs (tx) are untouched; one-pass discipline stands.
-- RUN LAST (heaviest query; after all other tier-14 files tie out).
-- Expected tie-outs (pre-registered, per the keeper's scope note): every
-- month's episodes/matched/delinquent/with_transcript/pay_language/
-- no_payment_30d/chargeoff_8m and the six lx4_ columns are all
-- <= the recorded unfiltered lx4 values (12-row grid in
-- uc2-anchoring\sprint\lexicon-v2\tie-outs.md) - ex-AA removes account-
-- months only, never adds. Row count = 12 months (unchanged from the
-- original; the window and month grid are untouched). Internal identities
-- per month (unchanged from the original, now on the ex-AA population):
-- pay_language minus lx4_pay_language_net_dec <= lx4_deceased_eps;
-- no_payment_30d minus lx4_no_payment_30d_net_dec <= lx4_deceased_eps;
-- lx4_leaked_with_exec <= lx4_exec_eps. No exact reproduction expected
-- anywhere (new face) except the row count. 2024-07 remains a boundary
-- artifact month: quote nothing from it.
-- Everything else is unchanged from lx4_funnel_v2_monthly.sql: the window
-- (2024-07 through 2026-03 snapshots, W3 call months 2024-07..2025-06), the
-- cumulative gate ladder (deepest 2-8), the raw-payment 30-day gate (no
-- autopay/NSF exclusion, matching f2 not f1), the deceased and execution
-- lexicons, and the eight-plus-six output columns.
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
                    try(cast(date_parse(try_cast(paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS pay_dt,
           clnt_prdct_cd,
           eff_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= DATE '2024-07-01'
      AND date(date_parse(eff_dt, '%Y%m%d')) < DATE '2026-03-01'
      AND eff_dt >= '20240701' AND eff_dt < '20260301'
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, min(co_dt) AS co_dt,
           max(pay_dt) AS pay_dt,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc
    FROM snap GROUP BY 1, 2
),
monthly2 AS (
    SELECT extnl_acct_id, m, bucket, pay_dt, eom_cpc,
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
    SELECT e.acct_key, e.contactid, e.call_dt, e.call_month,
           s.bucket, s.acct_co_dt, s.pay_dt, s.next_pay_dt,
           CASE WHEN s.bucket >= 1
                 AND (s.acct_co_dt IS NULL OR s.acct_co_dt > e.call_dt)
                 AND (s.eom_cpc IS NULL OR trim(s.eom_cpc) = ''
                      OR s.eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                                           'AA3','AC3','AM3','AA4','AC4','AM4',
                                           'BGC','BGM','CGM','GMR',
                                           'FBS','IBS','U1C','U2C','U3C'))
                THEN 1 ELSE 0 END AS is_delq
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
                        'passed away|death certificate|executor|deceased|calling on behalf'))
               AS deceased_n,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND regexp_like(lower(t.content),
                        'bank routing|routing number|check number|checkbook|a check for|that check|on the check'))
               AS exec_n
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM matched WHERE is_delq = 1) d
      ON t.contactid = d.contactid
     AND t.effdt >= '2024-07-01' AND t.effdt < '2025-07-02'
    GROUP BY 1
),
ep AS (
    SELECT m.call_month,
           CASE
             WHEN m.bucket IS NULL THEN 2
             WHEN m.is_delq = 0 THEN 3
             WHEN t.contactid IS NULL THEN 4
             WHEN t.pay_utts = 0 THEN 5
             WHEN (m.pay_dt IS NOT NULL
                   AND m.pay_dt >= m.call_dt
                   AND m.pay_dt <= date_add('day', 30, m.call_dt))
               OR (m.next_pay_dt IS NOT NULL
                   AND m.next_pay_dt >= m.call_dt
                   AND m.next_pay_dt <= date_add('day', 30, m.call_dt)) THEN 6
             WHEN m.acct_co_dt IS NULL
                  OR m.acct_co_dt > date_add('month', 8, m.call_dt) THEN 7
             ELSE 8
           END AS deepest,
           coalesce(t.deceased_n, 0) > 0 AS is_dec,
           coalesce(t.exec_n, 0) > 0 AS has_exec
    FROM matched m
    LEFT JOIN tx t ON m.contactid = t.contactid
)
SELECT call_month,
       count(*) AS episodes,
       count_if(deepest >= 3) AS matched,
       count_if(deepest >= 4) AS delinquent,
       count_if(deepest >= 5) AS with_transcript,
       count_if(deepest >= 6) AS pay_language,
       count_if(deepest >= 7) AS no_payment_30d,
       count_if(deepest >= 8) AS chargeoff_8m,
       count_if(deepest >= 5 AND is_dec) AS lx4_deceased_eps,
       count_if(deepest >= 6 AND NOT is_dec) AS lx4_pay_language_net_dec,
       count_if(deepest >= 7 AND NOT is_dec) AS lx4_no_payment_30d_net_dec,
       count_if(deepest >= 8 AND NOT is_dec) AS lx4_chargeoff_8m_net_dec,
       count_if(deepest >= 5 AND has_exec) AS lx4_exec_eps,
       count_if(deepest >= 7 AND has_exec) AS lx4_leaked_with_exec
FROM ep
GROUP BY 1
ORDER BY 1
