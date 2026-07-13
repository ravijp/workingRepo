-- Tier 15 | m4_exaa_bal, SINGLE-QUERY rewrite 2026-07-14 (P17) to fix the
-- out-of-memory failure WITHOUT creating any table (Ravi cannot CREATE TABLE).
-- Same numbers as the drafted version; only the execution shape changed.
--
-- WHY IT OOM'd, and the three single-query fixes applied:
--  (1) TWO transcript references. The draft referenced the tx CTE twice (once
--      in `caller` for pay_n, once in `ep2` for all six flags), so the planner
--      could scan/materialize the 4.9B-row transcript TWICE. FIX: tx is now
--      referenced EXACTLY ONCE. The caller-level any_captured / any_leaked_intent
--      are computed as window functions over the single episode-signal CTE
--      (esig), partitioned by acct_key, so no second transcript touch.
--  (2) participantid = 'CUSTOMER' sat INSIDE each count_if, so regexp_like ran
--      on agent rows too, then was discarded. FIX: participantid = 'CUSTOMER'
--      moved to the tx WHERE clause, plus a single coarse pre-filter rlike that
--      lets only utterances containing ANY relevant token reach the six precise
--      passes. This drops the overwhelming majority of utterances before the
--      expensive regex work. The six count_if flags are unchanged in meaning
--      (a row that fails the coarse gate matches none of them anyway).
--  (3) Boolean presence, not raw counts, carried forward (max(...) per
--      contactid) -> a narrow one-row-per-contactid intermediate.
--
-- SURVIVING HARD STOP: m4_episodes across all language-group rows = 11,262
-- EXACTLY. CO-count tie-outs are re-baselined (31-Jan anchor), NOTES only.
-- New columns: m4_co_10m_accounts, m4_jan_eom_balance, m4_co8/10/12_amt,
-- m4_jan_bal_co8/10/12. Windows anchored at 31 Jan (CO8 [01-31,09-30),
-- CO10 [..,11-30), CO12 [..,2026-01-31)). Balance/CO-dollar sums are deduped
-- one row per account per language group (acct_bal_grp).
--
-- COARSE PRE-FILTER NOTE: the WHERE rlike is the OR-union of all six lexicon
-- alternations. Any utterance matching a specific flag necessarily matches the
-- union, so gating on the union cannot drop a would-be match. Verified by
-- construction: the union string below is the concatenation of the six
-- per-flag alternations with '|' between them.
WITH snap AS (
    SELECT extnl_acct_id, substr(eff_dt, 1, 6) AS ym, eff_dt,
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
           try_cast(acct_bal_amt AS double) AS bal,
           try_cast(chrgoff_dt AS date) AS co_dt,
           clnt_prdct_cd
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20241201' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max_by(bucket, eff_dt) AS eom_bucket,
           max_by(bal, eff_dt) AS eom_bal,
           min(co_dt) AS co_dt,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc
    FROM snap GROUP BY 1, 2
),
ledger AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key, eom_bal
    FROM (SELECT * FROM monthly WHERE ym = '202501')
    WHERE eom_bucket = 1
      AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
      AND (eom_cpc IS NULL OR trim(eom_cpc) = ''
           OR eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                               'AA3','AC3','AM3','AA4','AC4','AM4',
                               'BGC','BGM','CGM','GMR',
                               'FBS','IBS','U1C','U2C','U3C'))
),
future_co AS (
    SELECT extnl_acct_id,
           min(try_cast(chrgoff_dt AS date)) AS co_dt_future,
           min_by(try_cast(chrgoff_amt AS double), try_cast(chrgoff_dt AS date)) AS co_amt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20260101'
      AND chrgoff_dt IS NOT NULL
    GROUP BY 1
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
        SELECT i.acct_key, i.contactid, i.call_dt,
               row_number() OVER (PARTITION BY i.acct_key, i.call_dt
                                  ORDER BY i.initiationtimestamp) AS rn
        FROM inb i
        JOIN ledger l ON l.acct_key = i.acct_key
        WHERE i.acct_key IS NOT NULL AND i.acct_key <> ''
    )
    WHERE rn = 1
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
    SELECT extnl_acct_id, m, max(pay_dt) AS pay_dt, max(auto_dt) AS auto_dt, max(nsf_dt) AS nsf_dt
    FROM pay_snap GROUP BY 1, 2
),
pay_monthly2 AS (
    SELECT extnl_acct_id, m, pay_dt, auto_dt, nsf_dt,
           lead(pay_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_pay_dt,
           lead(auto_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_auto_dt,
           lead(nsf_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_nsf_dt
    FROM pay_monthly
),
ep AS (
    SELECT l.acct_key, e.contactid,
           CASE WHEN
                  (s.pay_dt IS NOT NULL AND s.pay_dt >= e.call_dt
                   AND s.pay_dt <= date_add('day', 30, e.call_dt)
                   AND (s.auto_dt IS NULL OR s.auto_dt <> s.pay_dt)
                   AND (s.nsf_dt IS NULL OR s.nsf_dt <> s.pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.pay_dt))
                OR
                  (s.next_pay_dt IS NOT NULL AND s.next_pay_dt >= e.call_dt
                   AND s.next_pay_dt <= date_add('day', 30, e.call_dt)
                   AND (s.next_auto_dt IS NULL OR s.next_auto_dt <> s.next_pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.next_pay_dt))
                THEN 1 ELSE 0 END AS captured
    FROM ledger l
    JOIN episodes e ON e.acct_key = l.acct_key
    LEFT JOIN pay_monthly2 s
      ON e.acct_key = trim(cast(s.extnl_acct_id AS varchar))
     AND e.call_month = cast(s.m AS date)
),
-- The ONE transcript scan. participantid + a coarse union pre-filter are in the
-- WHERE, so regexp_like runs only on customer utterances that already contain a
-- relevant token. Referenced exactly once (below).
tx AS (
    SELECT t.contactid,
           max(CASE WHEN regexp_like(lower(t.content),
                     'passed away|death certificate|executor|deceased|calling on behalf') THEN 1 ELSE 0 END) AS deceased_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'pay|paid|payment|settle|payment plan|arrangement|work something out') THEN 1 ELSE 0 END) AS pay_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'settle|payment plan|arrangement|work something out') THEN 1 ELSE 0 END) AS plan_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford') THEN 1 ELSE 0 END) AS hard_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'dispute|not my charge|didn.t authorize|did not authorize|unauthorized|fraud|identity theft') THEN 1 ELSE 0 END) AS dispute_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'i.ll pay|i will pay|going to pay|gonna pay|pay (on|by|this|next)|when i get paid|payday|after my paycheck') THEN 1 ELSE 0 END) AS promise_f
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM ep) d ON t.contactid = d.contactid
    WHERE t.effdt >= '2025-01-01' AND t.effdt < '2025-02-02'
      AND t.content IS NOT NULL
      AND t.participantid = 'CUSTOMER'
      AND regexp_like(lower(t.content),
            'pay|paid|payment|settle|arrangement|work something out|passed away|death certificate|executor|deceased|calling on behalf|hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford|dispute|not my charge|didn.t authorize|did not authorize|unauthorized|fraud|identity theft|i.ll pay|i will pay|going to pay|gonna pay|when i get paid|payday|after my paycheck')
    GROUP BY 1
),
-- episode + signal + CO flags + captured, in ONE pass (tx referenced once here).
esig AS (
    SELECT e.acct_key, e.contactid, e.captured,
           CASE
             WHEN coalesce(x.deceased_f, 0) > 0 THEN 'a. deceased or estate'
             WHEN coalesce(x.promise_f, 0)  > 0 THEN 'b. future-dated promise'
             WHEN coalesce(x.pay_f, 0) > 0 AND coalesce(x.plan_f, 0) = 0 THEN 'c. payment talk, no promise'
             WHEN coalesce(x.plan_f, 0)     > 0 THEN 'd. plan or settlement talk'
             WHEN coalesce(x.hard_f, 0)     > 0 THEN 'e. hardship talk'
             WHEN coalesce(x.dispute_f, 0)  > 0 THEN 'f. dispute or fraud talk'
             ELSE 'g. no payment-related language'
           END AS language_group,
           coalesce(x.pay_f, 0) AS pay_f,
           (f.co_dt_future >= DATE '2025-01-31' AND f.co_dt_future < DATE '2025-09-30') AS co_8m,
           (f.co_dt_future >= DATE '2025-01-31' AND f.co_dt_future < DATE '2025-11-30') AS co_10m,
           (f.co_dt_future >= DATE '2025-01-31' AND f.co_dt_future < DATE '2026-01-31') AS co_12m,
           f.co_amt
    FROM ep e
    LEFT JOIN tx x ON x.contactid = e.contactid
    LEFT JOIN future_co f ON trim(cast(f.extnl_acct_id AS varchar)) = e.acct_key
),
-- account-level flags via window functions over esig (NO second transcript touch):
-- any_captured and any_leaked_intent (captured=0 AND pay_f>0 on some episode).
esig_acct AS (
    SELECT acct_key, contactid, captured, language_group, co_8m, co_10m, co_12m, co_amt,
           max(captured) OVER (PARTITION BY acct_key) AS any_captured,
           max(CASE WHEN captured = 0 AND pay_f > 0 THEN 1 ELSE 0 END)
               OVER (PARTITION BY acct_key) AS any_leaked_intent
    FROM esig
),
ep2 AS (
    SELECT acct_key, captured, language_group, co_8m, co_10m, co_12m, co_amt,
           (any_captured = 0 AND any_leaked_intent = 1) AS leaked_intent_acct
    FROM esig_acct
),
acct_group AS (   -- dedup to one row per (language_group, acct_key)
    SELECT language_group, acct_key,
           bool_or(co_8m)  AS acct_co_8m,
           bool_or(co_10m) AS acct_co_10m,
           bool_or(co_12m) AS acct_co_12m
    FROM ep2 GROUP BY 1, 2
),
acct_bal AS (
    SELECT ag.language_group, ag.acct_key, ag.acct_co_8m, ag.acct_co_10m, ag.acct_co_12m,
           l.eom_bal, f.co_amt
    FROM acct_group ag
    JOIN ledger l ON l.acct_key = ag.acct_key
    LEFT JOIN future_co f ON trim(cast(f.extnl_acct_id AS varchar)) = ag.acct_key
),
acct_bal_grp AS (
    SELECT language_group,
           round(sum(eom_bal), 0) AS m4_jan_eom_balance,
           round(sum(CASE WHEN acct_co_8m  THEN co_amt END), 0) AS m4_co8_amt,
           round(sum(CASE WHEN acct_co_10m THEN co_amt END), 0) AS m4_co10_amt,
           round(sum(CASE WHEN acct_co_12m THEN co_amt END), 0) AS m4_co12_amt,
           round(sum(CASE WHEN acct_co_8m  THEN eom_bal END), 0) AS m4_jan_bal_co8,
           round(sum(CASE WHEN acct_co_10m THEN eom_bal END), 0) AS m4_jan_bal_co10,
           round(sum(CASE WHEN acct_co_12m THEN eom_bal END), 0) AS m4_jan_bal_co12
    FROM acct_bal GROUP BY 1
)
SELECT e.language_group AS m4_group,
       count(*) AS m4_episodes,
       count(DISTINCT e.acct_key) AS m4_accounts,
       round(100.0 * sum(e.captured) / count(*), 1) AS m4_pct_paid_30d,
       count(DISTINCT CASE WHEN coalesce(e.co_8m, false)  THEN e.acct_key END) AS m4_co_8m_accounts,
       count(DISTINCT CASE WHEN coalesce(e.co_10m, false) THEN e.acct_key END) AS m4_co_10m_accounts,
       count(DISTINCT CASE WHEN coalesce(e.co_12m, false) THEN e.acct_key END) AS m4_co_12m_accounts,
       count(DISTINCT CASE WHEN coalesce(e.leaked_intent_acct, false) THEN e.acct_key END) AS m4_leaked_intent_accounts,
       g.m4_jan_eom_balance, g.m4_co8_amt, g.m4_co10_amt, g.m4_co12_amt,
       g.m4_jan_bal_co8, g.m4_jan_bal_co10, g.m4_jan_bal_co12
FROM ep2 e
JOIN acct_bal_grp g ON g.language_group = e.language_group
GROUP BY 1, 9, 10, 11, 12, 13, 14, 15
ORDER BY 1
