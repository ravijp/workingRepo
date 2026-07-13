-- Tier 15 | ADDITIVE VARIANT of lx4_exaa.sql, drafted 2026-07-13, header
-- rewritten P17. Logic unchanged from tier 14; four money columns appended.
--
-- WHY lx4 IS THE ODD ONE OUT (NOT RE-ANCHORED). The tier-15 spec re-anchors
-- the charge-off windows to 2025-01-31 (CO8/CO10/CO12) for the fixed-Jan-
-- anchor files (b15, m4, b19, b14), which read the book AS OF 31 Jan 2025 and
-- scan a forward future_co CTE. lx4 is different by construction: it is the
-- YEAR-LONG monthly funnel, 12 rows, one per call month (2024-07 through
-- 2025-06), grain = episode, built on the "deepest" ladder. It has NO
-- future_co CTE and NO fixed 31-Jan anchor. Its charge-off measure is
-- chargeoff_8m = "the account charged off within 8 MONTHS OF THE CALL"
-- (deepest >= 8, i.e. m.acct_co_dt <= date_add('month', 8, m.call_dt) in ep),
-- a PER-CALL-MONTH ROLLING window. Each of the 12 months has its own call
-- date, so a single fixed 31-Jan CO8/10/12 window is meaningless here.
-- Therefore the spec's re-anchor rule DOES NOT APPLY to this file. The
-- chargeoff_8m count and its net-of-deceased variant are left EXACTLY as they
-- were (rolling, correct as-is); nothing is re-baselined.
--
-- BECAUSE lx4 IS NOT RE-ANCHORED, ALL ITS TIE-OUTS SURVIVE AS HARD STOP
-- RULES (they do not depend on any CO-window change):
--   * 12 rows returned (one per call month, 2024-07 through 2025-06). STOP.
--   * leaked (no_payment_30d, summed across all 12 months) = 118,069 EXACTLY. STOP.
--   * net-of-deceased leaked (lx4_no_payment_30d_net_dec, summed) = 113,192 EXACTLY. STOP.
--   * chargeoff_8m summed = 35,490 EXACTLY. STOP.
--   * lx4_chargeoff_8m_net_dec summed = 32,019 EXACTLY. STOP.
-- Also unchanged and expected: every month's episodes/matched/delinquent/
-- with_transcript/pay_language/no_payment_30d/chargeoff_8m and the six lx4_
-- columns reproduce the tier-14 recorded ex-AA values EXACTLY (the count
-- columns are untouched). Any mismatch on ANY of the above = STOP, route to
-- the keeper before trusting the new money columns below. 2024-07 remains a
-- boundary artifact month: quote nothing from it, new columns included.
--
-- THE ONE CHANGE FROM TIER 14: four money columns appended to the final
-- SELECT. No count column, ladder, gate, join, or window is touched.
-- acct_bal_amt (decimal(17,3), the Jan/call-month EOM balance source) and
-- chrgoff_amt (decimal(17,3), the GROSS charge-off dollar) are added to
-- snap's SELECT list (same FROM/WHERE, no new scan; they ride rows already
-- read), threaded through monthly (one more max_by each, same GROUP BY 1, 2:
-- max_by(try_cast(acct_bal_amt AS double), eff_dt) AS eom_bal via the `bal`
-- alias, max_by(try_cast(chrgoff_amt AS double), eff_dt) AS eom_co_amt) and
-- through monthly2 (both columns added to the SELECT list; the two existing
-- window functions - min(co_dt) OVER, lead(pay_dt) OVER - are untouched and
-- do not reference the new columns). matched picks up s.eom_bal and
-- s.eom_co_amt the same way it already picks up s.bucket/s.pay_dt; ep carries
-- them to the final SELECT.
--
-- Balance is the account's balance IN THE CALL MONTH: eom_bal is the row from
-- monthly2 matched to e.call_month via the existing s.m = e.call_month join
-- condition in `matched`, i.e. the same month grain the query uses
-- throughout - NOT the account's balance at time of eventual charge-off. The
-- CO dollar eom_co_amt is chrgoff_amt at that same call-month row. This is
-- the RIGHT money grain for a rolling per-call-month funnel; there is no
-- 31-Jan balance to take here, unlike the fixed-anchor files.
--
-- New final-SELECT columns, appended at the end, per month (each mirrors an
-- existing count column's predicate EXACTLY, so it moves with that count):
--   lx4_chargeoff_8m_amt         = sum(eom_co_amt) where deepest >= 8
--                                  (mirrors chargeoff_8m). GROSS CO $.
--   lx4_chargeoff_8m_amt_net_dec = same AND NOT is_dec
--                                  (mirrors lx4_chargeoff_8m_net_dec).
--   lx4_leaked_bal               = sum(eom_bal) where deepest >= 7
--                                  (mirrors no_payment_30d). Call-month balance.
--   lx4_leaked_bal_net_dec       = same AND NOT is_dec
--                                  (mirrors lx4_no_payment_30d_net_dec).
-- GRAIN NOTE: these four sums are over EPISODES, matching the funnel's own
-- episode-count semantics (an account with two episodes in a call month
-- contributes twice, exactly as it does to no_payment_30d / chargeoff_8m).
-- The spec's one-row-per-account dedup CTE is for the account-level money
-- reads in m4/b19; it deliberately does NOT apply to this funnel's leak and
-- charge-off dollar columns, which are episode-attributed by design.
--
-- PLAUSIBILITY BOUNDS (state, not tie-outs): all four sums >= 0;
-- lx4_leaked_bal >= lx4_leaked_bal_net_dec per month (net-of-deceased is a
-- subset); lx4_chargeoff_8m_amt >= lx4_chargeoff_8m_amt_net_dec per month,
-- same reason; each money column moves month over month in the same
-- direction as its mirror count, since both are restricted by the identical
-- deepest/is_dec predicate over the same row set.
--
-- RESOURCE DISCIPLINE: no new scans; the call-CTEs-first / semi-join
-- structure from the tier-14 resource fix is untouched (inb, episodes stay at
-- the top of the WITH chain; snap still semi-joins to episodes' accounts
-- only). The two added columns ride the same window/group passes already in
-- place. The ex-AA exclusion, the deepest ladder, the delinquency gate, the
-- deceased/execution lexicon regexes, and the raw-payment 30-day gate are all
-- byte-for-byte unchanged from lx4_exaa.sql.
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
