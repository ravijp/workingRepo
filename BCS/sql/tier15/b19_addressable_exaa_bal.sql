-- Tier 15 | ADDITIVE VARIANT of b19_addressable_exaa.sql, drafted 2026-07-13;
-- logic unchanged, columns appended. This is the addressable January call
-- stream: bucket-1-at-call ex-AA episodes by v2 language group x captured.
--
-- SURVIVING HARD STOP (does NOT depend on the CO window): total episodes =
-- 29,114 EXACTLY (sum of all b19_episodes across every language_group x
-- captured cell). The language_group x captured cells partition that total
-- EXACTLY; distinct accounts <= episodes in every cell. Any mismatch = STOP,
-- route to the keeper before trusting the new columns below.
--
-- CO WINDOWS RE-ANCHORED to 31 Jan 2025 (per Ravi, NOT the old Jan-01 anchor).
-- The forward charge-off scan (future_co CTE) computes co_dt_future =
-- min(try_cast(chrgoff_dt AS date)) over 2025 daily rows and, paired to it,
-- co_amt = min_by(try_cast(chrgoff_amt AS double), try_cast(chrgoff_dt AS
-- date)) FILTER (WHERE chrgoff_dt IS NOT NULL). Windows:
--   CO8  : co_dt_future >= DATE '2025-01-31' AND co_dt_future < DATE '2025-09-30'
--   CO10 : co_dt_future >= DATE '2025-01-31' AND co_dt_future < DATE '2025-11-30'
--   CO12 : co_dt_future >= DATE '2025-01-31' AND co_dt_future < DATE '2026-01-31'
-- The CO-window account counts and dollar sums are NEW measurements on the
-- re-anchored windows: they have no pre-registered value and are NOT tie-outs.
-- Plausibility bounds only (state, do not enforce): co12 >= co10 >= co8 for
-- both the account counts and the dollar sums per row; all balance/CO-dollar
-- sums >= 0; coN_amt is the same order of magnitude as the balance on those
-- accounts (charge-off amount ~ balance scale, not equal).
--
-- NEW COLUMNS appended to the final SELECT:
--   b19_jan_eom_balance = round(sum(account Jan EOM balance), 0)
--   b19_co_8m / _co_10m / _co_12m   = ACCOUNT COUNT charging off in each window
--   b19_co8_amt / _co10_amt / _co12_amt = round(sum(gross chrgoff_amt), 0)
-- Jan EOM balance is sourced from a small new jan_bal CTE: max_by(
-- try_cast(acct_bal_amt AS double), eff_dt) over the JANUARY snapshots only
-- (eff_dt in [20250101, 20250201)), semi-joined to episodes_exaa. It is NOT
-- read off the existing snap CTE (which spans 2024-06..2025-01) so there is
-- no risk of picking up a pre-January EOM balance. future_co is likewise
-- semi-joined to episodes_exaa. Neither adds a new full base-table scan: both
-- are bounded to the January episode accounts, matching the file's resource
-- discipline (the episode CTEs run first, both new CTEs semi-join to them).
--
-- DEDUP NOTE: the output grain is episode (language_group x captured), so an
-- account can appear in multiple episodes within one cell. Balance and CO
-- dollars MUST be summed one row per account within each (language_group,
-- captured) group, never once per episode. Reuses the m4 acct_group/acct_bal/
-- acct_bal_grp pattern: acct_group dedups ep2 to one row per (language_group,
-- captured, acct_key) and bool_or's the CO flags; acct_bal joins jan_bal.
-- jan_eom_bal and future_co.co_amt onto that; acct_bal_grp aggregates to one
-- row per (language_group, captured); the final SELECT joins that back onto
-- the untouched ep2-grain query. b19_accounts and the b19_co_Nm account counts
-- use count(DISTINCT acct_key). No existing CTE's SELECT list, join, or WHERE
-- clause is edited; the ex-AA exclusion, the bucket-1-at-call filter, the
-- payment gate, and the language priority CASE are all unchanged.
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
-- NEW (additive): the account's January EOM balance. Read directly off the
-- January snapshots only (eff_dt in [20250101, 20250201)), semi-joined to the
-- episode accounts, so it cannot pick up a pre-January balance. Not derived
-- from snap (which spans 2024-06..2025-01). No new full scan.
jan_bal AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           max_by(try_cast(acct_bal_amt AS double), eff_dt) AS jan_eom_bal
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250201'
      AND trim(cast(extnl_acct_id AS varchar)) IN (SELECT acct_key FROM episodes_exaa)
    GROUP BY 1
),
-- NEW (additive): forward charge-off scan over 2025 daily rows, semi-joined to
-- the episode accounts to bound cost. co_amt is the gross charge-off dollar
-- paired to the earliest charge-off date.
future_co AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           min(try_cast(chrgoff_dt AS date)) AS co_dt_future,
           -- co_amt = the charge-off amount on the row with the earliest
           -- charge-off date. The CTE's WHERE already restricts to
           -- chrgoff_dt IS NOT NULL, so no FILTER clause is needed (kept out
           -- to avoid a syntax path not exercised by the verified tier-14 kit).
           min_by(try_cast(chrgoff_amt AS double), try_cast(chrgoff_dt AS date)) AS co_amt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20260101'
      AND chrgoff_dt IS NOT NULL
      AND trim(cast(extnl_acct_id AS varchar)) IN (SELECT acct_key FROM episodes_exaa)
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
),
-- account-level dedup at the output grain (language_group, captured): one row
-- per (language_group, captured, acct_key), carrying the account's forward
-- charge-off window flags on the re-anchored windows, so the final SUM cannot
-- double count an account with multiple episodes in the same cell.
acct_group AS (
    SELECT e.language_group, e.captured, e.acct_key,
           bool_or(f.co_dt_future >= DATE '2025-01-31'
                   AND f.co_dt_future < DATE '2025-09-30') AS acct_co_8m,
           bool_or(f.co_dt_future >= DATE '2025-01-31'
                   AND f.co_dt_future < DATE '2025-11-30') AS acct_co_10m,
           bool_or(f.co_dt_future >= DATE '2025-01-31'
                   AND f.co_dt_future < DATE '2026-01-31') AS acct_co_12m
    FROM ep2 e
    LEFT JOIN future_co f ON f.acct_key = e.acct_key
    GROUP BY 1, 2, 3
),
acct_bal AS (
    SELECT ag.language_group, ag.captured, ag.acct_key,
           ag.acct_co_8m, ag.acct_co_10m, ag.acct_co_12m,
           b.jan_eom_bal, f.co_amt
    FROM acct_group ag
    LEFT JOIN jan_bal b ON b.acct_key = ag.acct_key
    LEFT JOIN future_co f ON f.acct_key = ag.acct_key
),
acct_bal_grp AS (
    SELECT language_group, captured,
           round(sum(jan_eom_bal), 0) AS b19_jan_eom_balance,
           count(DISTINCT CASE WHEN acct_co_8m  THEN acct_key END) AS b19_co_8m,
           count(DISTINCT CASE WHEN acct_co_10m THEN acct_key END) AS b19_co_10m,
           count(DISTINCT CASE WHEN acct_co_12m THEN acct_key END) AS b19_co_12m,
           round(sum(CASE WHEN acct_co_8m  THEN co_amt END), 0) AS b19_co8_amt,
           round(sum(CASE WHEN acct_co_10m THEN co_amt END), 0) AS b19_co10_amt,
           round(sum(CASE WHEN acct_co_12m THEN co_amt END), 0) AS b19_co12_amt
    FROM acct_bal
    GROUP BY 1, 2
)
SELECT e.language_group AS b19_language_group,
       e.captured AS b19_captured,
       count(*) AS b19_episodes,
       count(DISTINCT e.acct_key) AS b19_accounts,
       g.b19_jan_eom_balance,
       g.b19_co_8m,
       g.b19_co_10m,
       g.b19_co_12m,
       g.b19_co8_amt,
       g.b19_co10_amt,
       g.b19_co12_amt
FROM ep2 e
JOIN acct_bal_grp g
  ON g.language_group = e.language_group
 AND g.captured = e.captured
GROUP BY 1, 2, 5, 6, 7, 8, 9, 10, 11
ORDER BY 1, 2
