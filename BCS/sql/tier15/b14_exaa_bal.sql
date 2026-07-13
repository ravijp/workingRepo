-- Tier 15 | ADDITIVE VARIANT of b14_exaa.sql, drafted 2026-07-13; logic
-- unchanged, columns appended.
-- WHAT THIS FILE ADDS (Ravi point 4): the same 31-Jan-anchored balance and
-- forward charge-off treatment carried by the other tier-15 _bal files, on
-- the ledger transitions grain (feb_position x caller_class x runway_band,
-- account-level, one ledger row per account).
-- RE-ANCHORED CO WINDOWS: the forward charge-off scan (future_co CTE, the
-- same shape as b15) computes co_dt_future = min(try_cast(chrgoff_dt AS
-- date)) over 2025 daily rows and, paired to it, co_amt = min_by(
-- try_cast(chrgoff_amt AS double), try_cast(chrgoff_dt AS date)) filtered to
-- rows with a charge-off date, so co_amt is the amount on the SAME row as the
-- earliest co_dt. All three CO windows are anchored at 2025-01-31 per Ravi
-- (NOT 2025-01-01):
--   CO8  : co_dt_future >= DATE '2025-01-31' AND co_dt_future < DATE '2025-09-30'
--   CO10 : co_dt_future >= DATE '2025-01-31' AND co_dt_future < DATE '2025-11-30'
--   CO12 : co_dt_future >= DATE '2025-01-31' AND co_dt_future < DATE '2026-01-31'
-- SURVIVING HARD STOP TIE-OUTS (do NOT depend on the CO window): total
-- accounts (b14_accounts summed across ALL rows) = 189,146 EXACTLY; total
-- balance (b14_jan_eom_balance summed) = ~457,943,987 (rounding tolerance
-- ~$5). Any mismatch on either = STOP, route to the keeper before running
-- any other tier-15 file and before trusting the new columns below. This
-- query MUST still run FIRST AND ALONE per the tier-14 header.
-- RE-BASELINED (NOT a STOP): the CO-window counts and dollars below are new
-- measurements on the 31-Jan anchor. There is no prior recorded value to tie
-- them to. State the plausibility bounds only, do not gate on an exact match.
-- NEW COLUMNS, appended after the existing ones:
--   b14_co_amt   : the FEB charge-off dollar (sum of feb_co_amt gated to
--                  feb_position = 'e. charged off in Feb'). KEPT from the
--                  prior draft as a useful extra; it is the Feb monthly-row
--                  charge-off amount, distinct from the forward-scan columns.
--   b14_co_8m / _co_10m / _co_12m   : ACCOUNT counts with co_dt_future in the
--                  CO8 / CO10 / CO12 window (count_if on the forward scan).
--   b14_co8_amt / _co10_amt / _co12_amt : round(sum(co_amt in window), 0),
--                  the GROSS forward charge-off dollar in each window.
-- The Feb-CO plumbing (eom_co_amt in monthly, feb_co_amt in base/ledger) is
-- unchanged from the prior draft and only feeds b14_co_amt.
-- PLAUSIBILITY BOUNDS (not tie-outs): all balance/CO-dollar sums >= 0;
-- b14_co_12m >= b14_co_10m >= b14_co_8m per row and likewise b14_co12_amt >=
-- b14_co10_amt >= b14_co8_amt (the CO12 window contains CO10 contains CO8);
-- each coN_amt is balance-scale on those accounts (charge-off amount ~
-- balance scale, not equal); b14_co_amt >= 0 and is 0 (or NULL-summed-to-0)
-- for every row whose feb_position is NOT 'e. charged off in Feb'.
-- ACCOUNT-LEVEL GRAIN: the ledger holds one row per account, so plain sums
-- (count(*), sum(eom_bal), sum(co_amt), count_if) are correct here with no
-- per-episode dedup. The acct_group/acct_bal dedup CTE used by the
-- episode-grained files (m4, b19) is not needed and is not added.
-- RESOURCE DISCIPLINE: future_co is the same forward scan already used by
-- tier 14 (fmt_acct_c, 2025-01-01 to 2026-01-01, chrgoff_dt IS NOT NULL);
-- its SELECT list only gains co_amt via a second aggregate over the
-- already-scanned rows. chrgoff_amt and acct_bal_amt ride scans already in
-- place. No new base-table scans; the query stays run first and alone.
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
      AND eff_dt >= '20241201' AND eff_dt < '20250301'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           max_by(bal, eff_dt) AS eom_bal,
           min(co_dt) AS co_dt,
           min(CASE WHEN bucket >= 1 THEN eff_dt END) AS first_dq_dt,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc,
           max_by(try_cast(chrgoff_amt AS double), eff_dt) AS eom_co_amt
    FROM snap GROUP BY 1, 2
),
base AS (
    SELECT j.extnl_acct_id, j.max_bucket, j.eom_bucket, j.eom_bal, j.co_dt,
           j.first_dq_dt, j.eom_cpc,
           p.max_bucket AS prev_max_bucket,
           p.eom_bucket AS prev_eom_bucket,
           f.eom_bucket AS feb_eom_bucket,
           f.co_dt AS feb_co_dt,
           f.eom_co_amt AS feb_co_amt,
           (f.extnl_acct_id IS NOT NULL) AS has_feb_row
    FROM (SELECT * FROM monthly WHERE ym = '202501') j
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202412') p
      ON j.extnl_acct_id = p.extnl_acct_id
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202502') f
      ON j.extnl_acct_id = f.extnl_acct_id
),
ledger AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           eom_bal,
           feb_co_amt,
           CASE
             WHEN coalesce(prev_eom_bucket, 0) >= 1
               THEN 'd. carried-in (past due at Dec-31 EOM)'
             WHEN cast(substr(first_dq_dt, 7, 2) AS integer) <= 10
               THEN 'a. runway >= 21 days (entry day 1-10)'
             WHEN cast(substr(first_dq_dt, 7, 2) AS integer) <= 20
               THEN 'b. runway 11-20 days (entry day 11-20)'
             ELSE 'c. runway <= 10 days (entry day 21-31)'
           END AS runway_band,
           CASE
             WHEN feb_co_dt >= DATE '2025-02-01' AND feb_co_dt < DATE '2025-03-01'
               THEN 'e. charged off in Feb'
             WHEN NOT has_feb_row THEN 'f. no Feb row'
             WHEN feb_eom_bucket = 0 THEN 'a. Feb EOM bucket 0 (cured)'
             WHEN feb_eom_bucket = 1 THEN 'b. Feb EOM bucket 1 (stayed)'
             WHEN feb_eom_bucket = 2 THEN 'c. Feb EOM bucket 2 (rolled)'
             ELSE 'd. Feb EOM bucket 3+ (rolled deeper)'
           END AS feb_position
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
           -- co_amt = the charge-off amount on the row with the earliest
           -- charge-off date. The CTE's WHERE already restricts to
           -- chrgoff_dt IS NOT NULL, so no FILTER clause is needed (kept out
           -- to avoid a syntax path not exercised by the verified tier-14 kit).
           min_by(try_cast(chrgoff_amt AS double), try_cast(chrgoff_dt AS date)) AS co_amt
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
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND regexp_like(lower(t.content),
                        'pay|paid|payment|settle|payment plan|arrangement|work something out'))
               AS pay_utts
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM ep) d
      ON t.contactid = d.contactid
     AND t.effdt >= '2025-01-01' AND t.effdt < '2025-02-02'
    GROUP BY 1
),
caller AS (
    SELECT e.acct_key,
           max(e.captured) AS any_captured,
           max(CASE WHEN e.captured = 0 AND coalesce(x.pay_utts, 0) > 0
                    THEN 1 ELSE 0 END) AS any_leaked_intent
    FROM ep e
    LEFT JOIN tx x ON e.contactid = x.contactid
    GROUP BY 1
)
SELECT l.feb_position AS b14_feb_position,
       CASE
         WHEN k.acct_key IS NULL THEN 'a. non-caller'
         WHEN k.any_captured = 1 THEN 'b. captured (>= 1 paid-30d episode)'
         WHEN k.any_leaked_intent = 1 THEN 'c. leaked-intent (intent, no payment 30d)'
         ELSE 'd. other-caller'
       END AS b14_caller_class,
       l.runway_band AS b14_runway_band,
       count(*) AS b14_accounts,
       round(sum(l.eom_bal), 0) AS b14_jan_eom_balance,
       round(sum(CASE WHEN l.feb_position = 'e. charged off in Feb'
                      THEN l.feb_co_amt END), 0) AS b14_co_amt,
       count_if(f.co_dt_future >= DATE '2025-01-31'
                AND f.co_dt_future < DATE '2025-09-30') AS b14_co_8m,
       count_if(f.co_dt_future >= DATE '2025-01-31'
                AND f.co_dt_future < DATE '2025-11-30') AS b14_co_10m,
       count_if(f.co_dt_future >= DATE '2025-01-31'
                AND f.co_dt_future < DATE '2026-01-31') AS b14_co_12m,
       round(sum(CASE WHEN f.co_dt_future >= DATE '2025-01-31'
                       AND f.co_dt_future < DATE '2025-09-30'
                      THEN f.co_amt END), 0) AS b14_co8_amt,
       round(sum(CASE WHEN f.co_dt_future >= DATE '2025-01-31'
                       AND f.co_dt_future < DATE '2025-11-30'
                      THEN f.co_amt END), 0) AS b14_co10_amt,
       round(sum(CASE WHEN f.co_dt_future >= DATE '2025-01-31'
                       AND f.co_dt_future < DATE '2026-01-31'
                      THEN f.co_amt END), 0) AS b14_co12_amt
FROM ledger l
LEFT JOIN caller k ON l.acct_key = k.acct_key
LEFT JOIN future_co f
  ON trim(cast(f.extnl_acct_id AS varchar)) = l.acct_key
GROUP BY 1, 2, 3
ORDER BY 1, 2, 3
