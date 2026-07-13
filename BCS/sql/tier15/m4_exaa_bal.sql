-- Tier 15 | ADDITIVE VARIANT of m4_exaa.sql, drafted 2026-07-13; logic
-- unchanged, columns appended.
-- PRE-REGISTERED TIE-OUT (STOP RULE): m4_episodes across ALL language-group
-- rows sums to 11,262 EXACTLY on re-run. Any mismatch = STOP, route to the
-- keeper before trusting the new columns below.
-- ADDITION: the original snap/monthly CTEs did not carry acct_bal_amt or
-- chrgoff_amt (m4 had no balance column at all). One more expression is
-- added to the existing January-window snap SELECT (bal, co_amt) and
-- threaded through the existing GROUP BY in monthly (max_by(bal, eff_dt)
-- AS eom_bal, matching b14's eom_bal derivation exactly, over the SAME
-- 2024-12-01/2025-02-01 scan already in place) and into base/ledger. NO
-- NEW SCAN is added. chrgoff_amt is threaded the same way as a per-account
-- eom_co_amt = max_by(try_cast(chrgoff_amt AS double), eff_dt) inside the
-- same January monthly aggregation (this is the account's own charge-off
-- amount if it charged off within the Jan snapshot window; for the
-- forward CO8/CO12 flags, which are driven off future_co's forward scan,
-- the amount is threaded from future_co the same way as b15/b14: a paired
-- min_by(chrgoff_amt, chrgoff_dt) added to that CTE's SELECT).
-- DEDUP NOTE: m4_accounts is COUNT(DISTINCT acct_key) per language_group,
-- i.e. one account can appear in multiple episodes within a group but is
-- counted once. m4_jan_eom_balance is computed at the SAME account-dedup
-- level: it sums l.eom_bal once per DISTINCT acct_key within the group
-- (via the ledger_bal subquery keyed on acct_key, joined back after
-- dedup), not once per episode row, so it will NOT double count an
-- account with 2+ episodes in the group. This required two new CTEs
-- (acct_group, acct_bal, acct_bal_grp) that sit alongside ep2 rather than
-- inside it: acct_group dedups ep2 down to one row per (language_group,
-- acct_key); acct_bal joins that to ledger.eom_bal and future_co.co_amt;
-- acct_bal_grp aggregates to one row per group; the final SELECT joins
-- that back onto the untouched ep2-grain query. No existing CTE's SELECT
-- list, join, or WHERE clause is edited - this is the one deviation from
-- "same CTEs only" the spec allows for ("if account dedup ... makes a
-- clean per-group balance impossible without restructuring, say so and
-- compute at the same account-dedup level as m4_accounts") - done here.
-- New final-SELECT columns, appended at the end: m4_jan_eom_balance (sum
-- of eom_bal, one row per distinct account in the group), m4_co8_amt,
-- m4_co12_amt (sum of the future-charge-off amount for accounts flagged
-- co_8m / co_12m respectively, deduped the same way), m4_jan_bal_co8,
-- m4_jan_bal_co12 (Jan EOM balance restricted to those same two flag
-- sets, same dedup).
-- PLAUSIBILITY BOUNDS: all four amount/balance sums >= 0; m4_co12_amt >=
-- m4_co8_amt and m4_jan_bal_co12 >= m4_jan_bal_co8 per group (CO12 window
-- is a superset of CO8); m4_jan_bal_co8/co12 <= m4_jan_eom_balance per
-- group; m4_co8_amt/co12_amt are on the same dollar scale as
-- m4_jan_bal_co8/co12 for the same accounts (order-of-magnitude check).
-- RESOURCE DISCIPLINE: no new base-table scans; snap/monthly/future_co
-- reuse the exact WHERE clauses already in place; the only change to
-- scan cost is one extra max_by expression evaluated per group in each
-- existing aggregation.
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
           try_cast(acct_bal_amt AS double) AS bal,
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
           max_by(bal, eff_dt) AS eom_bal,
           min(co_dt) AS co_dt,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc
    FROM snap GROUP BY 1, 2
),
base AS (
    SELECT j.extnl_acct_id, j.max_bucket, j.eom_bucket, j.eom_bal, j.co_dt, j.eom_cpc,
           p.max_bucket AS prev_max_bucket
    FROM (SELECT * FROM monthly WHERE ym = '202501') j
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202412') p
      ON j.extnl_acct_id = p.extnl_acct_id
),
ledger AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           eom_bal
    FROM base
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
           min_by(try_cast(chrgoff_amt AS double), try_cast(chrgoff_dt AS date))
             FILTER (WHERE chrgoff_dt IS NOT NULL) AS co_amt
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
    SELECT l.acct_key, e.contactid,
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
caller AS (
    SELECT e.acct_key,
           max(e.captured) AS any_captured,
           max(CASE WHEN e.captured = 0 AND coalesce(x.pay_utts_n, 0) > 0
                    THEN 1 ELSE 0 END) AS any_leaked_intent
    FROM ep e
    LEFT JOIN (SELECT contactid, pay_n AS pay_utts_n FROM tx) x
      ON e.contactid = x.contactid
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
           END AS language_group,
           (f.co_dt_future >= DATE '2025-01-01'
            AND f.co_dt_future < DATE '2025-09-01') AS co_8m,
           (f.co_dt_future >= DATE '2025-01-01'
            AND f.co_dt_future < DATE '2026-01-01') AS co_12m,
           (k.any_captured = 0 AND k.any_leaked_intent = 1) AS leaked_intent_acct
    FROM ep e
    LEFT JOIN tx x ON e.contactid = x.contactid
    LEFT JOIN future_co f
      ON trim(cast(f.extnl_acct_id AS varchar)) = e.acct_key
    LEFT JOIN caller k ON e.acct_key = k.acct_key
),
-- account-level dedup (same grain as m4_accounts = COUNT(DISTINCT acct_key)
-- per language_group): one row per (language_group, acct_key), carrying the
-- account's Jan EOM balance and forward charge-off amount/flags, so the
-- final SUM cannot double count an account with multiple episodes.
acct_group AS (
    SELECT language_group, acct_key,
           bool_or(co_8m) AS acct_co_8m,
           bool_or(co_12m) AS acct_co_12m
    FROM ep2
    GROUP BY 1, 2
),
acct_bal AS (
    SELECT ag.language_group, ag.acct_key, ag.acct_co_8m, ag.acct_co_12m,
           l.eom_bal, f.co_amt
    FROM acct_group ag
    JOIN ledger l ON l.acct_key = ag.acct_key
    LEFT JOIN future_co f
      ON trim(cast(f.extnl_acct_id AS varchar)) = ag.acct_key
),
acct_bal_grp AS (
    SELECT language_group,
           round(sum(eom_bal), 0) AS m4_jan_eom_balance,
           round(sum(CASE WHEN acct_co_8m THEN co_amt END), 0) AS m4_co8_amt,
           round(sum(CASE WHEN acct_co_12m THEN co_amt END), 0) AS m4_co12_amt,
           round(sum(CASE WHEN acct_co_8m THEN eom_bal END), 0) AS m4_jan_bal_co8,
           round(sum(CASE WHEN acct_co_12m THEN eom_bal END), 0) AS m4_jan_bal_co12
    FROM acct_bal
    GROUP BY 1
)
SELECT e.language_group AS m4_group,
       count(*) AS m4_episodes,
       count(DISTINCT e.acct_key) AS m4_accounts,
       round(100.0 * sum(e.captured) / count(*), 1) AS m4_pct_paid_30d,
       count(DISTINCT CASE WHEN coalesce(e.co_8m, false) THEN e.acct_key END) AS m4_co_8m_accounts,
       count(DISTINCT CASE WHEN coalesce(e.co_12m, false) THEN e.acct_key END) AS m4_co_12m_accounts,
       count(DISTINCT CASE WHEN coalesce(e.leaked_intent_acct, false) THEN e.acct_key END) AS m4_leaked_intent_accounts,
       g.m4_jan_eom_balance,
       g.m4_co8_amt,
       g.m4_co12_amt,
       g.m4_jan_bal_co8,
       g.m4_jan_bal_co12
FROM ep2 e
JOIN acct_bal_grp g ON g.language_group = e.language_group
GROUP BY 1, 8, 9, 10, 11, 12
ORDER BY 1
