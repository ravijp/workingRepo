-- Tier 15 | ADDITIVE VARIANT of lx4_exaa.sql, drafted 2026-07-13; logic
-- unchanged, columns appended.
-- PRE-REGISTERED TIE-OUT (STOP RULE): 12 rows returned (one per call
-- month, 2024-07 through 2025-06 unchanged); episodes/matched/delinquent/
-- with_transcript/pay_language/no_payment_30d/chargeoff_8m columns
-- reproduce the tier-14 recorded ex-AA values EXACTLY: leaked (no_
-- payment_30d, summed across all 12 months) = 118,069 EXACTLY; net-of-
-- deceased leaked (lx4_no_payment_30d_net_dec, summed) = 113,192 EXACTLY;
-- chargeoff_8m summed = 35,490 EXACTLY; lx4_chargeoff_8m_net_dec summed =
-- 32,019 EXACTLY. Any mismatch = STOP, route to the keeper before trusting
-- the new columns below. 2024-07 remains a boundary artifact month per
-- the tier-14 header: quote nothing from it, new columns included.
-- ADDITION: snap already scans acct_bal_amt-adjacent columns are NOT
-- selected today, so acct_bal_amt and chrgoff_amt are added to snap's
-- SELECT list (same FROM/WHERE, no new scan), then threaded through monthly
-- (one more max_by each, same GROUP BY 1, 2: max_by(bal, eff_dt) AS
-- eom_bal, max_by(try_cast(chrgoff_amt AS double), eff_dt) AS eom_co_amt)
-- and through monthly2 (both columns simply added to the SELECT list,
-- the two existing window functions - min(co_dt) OVER, lead(pay_dt) OVER -
-- are untouched and do not need to reference the new columns). matched
-- picks up s.eom_bal and s.eom_co_amt from monthly2 the same way it
-- already picks up s.bucket/s.pay_dt; ep carries them through to the
-- final SELECT. Balance here is the account's balance IN THE CALL MONTH
-- (the row from monthly2 matched to e.call_month via the existing s.m =
-- e.call_month join condition in `matched`), i.e. the same month-grain
-- the query already uses throughout - NOT the account's balance at time
-- of eventual charge-off.
-- New final-SELECT columns, appended at the end, per month:
-- lx4_chargeoff_8m_amt = sum of eom_co_amt where deepest >= 8 (mirrors
-- chargeoff_8m's count_if predicate exactly); lx4_chargeoff_8m_amt_net_dec
-- = same predicate AND NOT is_dec (mirrors lx4_chargeoff_8m_net_dec);
-- lx4_leaked_bal = sum of eom_bal where deepest >= 7 (mirrors
-- no_payment_30d's count_if predicate); lx4_leaked_bal_net_dec = same
-- predicate AND NOT is_dec (mirrors lx4_no_payment_30d_net_dec).
-- PLAUSIBILITY BOUNDS: all four sums >= 0; lx4_leaked_bal >=
-- lx4_leaked_bal_net_dec per month (net-of-deceased is a subset);
-- lx4_chargeoff_8m_amt >= lx4_chargeoff_8m_amt_net_dec per month, same
-- reason; the four new columns are monotone with their corresponding
-- existing counts (no_payment_30d, chargeoff_8m, and their net_dec pairs)
-- month over month in the same direction, since both are sums restricted
-- by the identical deepest/is_dec predicates over the same row set.
-- RESOURCE DISCIPLINE: no new scans, no change to the call-CTEs-first /
-- semi-join structure from the tier-14 resource fix (inb, episodes stay at
-- the top of the WITH chain; snap still semi-joins to episodes' accounts
-- only); the two added columns ride the same window/group passes already
-- in place.
WITH inb AS (
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
snap AS (
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
           eff_dt,
           try_cast(acct_bal_amt AS double) AS bal,
           chrgoff_amt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= DATE '2024-07-01'
      AND date(date_parse(eff_dt, '%Y%m%d')) < DATE '2026-03-01'
      AND eff_dt >= '20240701' AND eff_dt < '20260301'
      AND trim(cast(extnl_acct_id AS varchar)) IN (SELECT acct_key FROM episodes)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, min(co_dt) AS co_dt,
           max(pay_dt) AS pay_dt,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc,
           max_by(bal, eff_dt) AS eom_bal,
           max_by(try_cast(chrgoff_amt AS double), eff_dt) AS eom_co_amt
    FROM snap GROUP BY 1, 2
),
monthly2 AS (
    SELECT extnl_acct_id, m, bucket, pay_dt, eom_cpc, eom_bal, eom_co_amt,
           min(co_dt) OVER (PARTITION BY extnl_acct_id) AS acct_co_dt,
           lead(pay_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_pay_dt
    FROM monthly
),
matched AS (
    SELECT e.acct_key, e.contactid, e.call_dt, e.call_month,
           s.bucket, s.acct_co_dt, s.pay_dt, s.next_pay_dt,
           s.eom_bal, s.eom_co_amt,
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
           coalesce(t.exec_n, 0) > 0 AS has_exec,
           m.eom_bal,
           m.eom_co_amt
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
       count_if(deepest >= 7 AND has_exec) AS lx4_leaked_with_exec,
       round(sum(CASE WHEN deepest >= 8 THEN eom_co_amt END), 0) AS lx4_chargeoff_8m_amt,
       round(sum(CASE WHEN deepest >= 8 AND NOT is_dec THEN eom_co_amt END), 0) AS lx4_chargeoff_8m_amt_net_dec,
       round(sum(CASE WHEN deepest >= 7 THEN eom_bal END), 0) AS lx4_leaked_bal,
       round(sum(CASE WHEN deepest >= 7 AND NOT is_dec THEN eom_bal END), 0) AS lx4_leaked_bal_net_dec
FROM ep
GROUP BY 1
ORDER BY 1
