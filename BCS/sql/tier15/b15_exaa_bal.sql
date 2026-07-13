-- Tier 15 | ADDITIVE VARIANT of b15_exaa.sql, drafted 2026-07-13; logic
-- unchanged, balance and charge-off dollar columns appended.
-- RE-ANCHORED CO WINDOWS (per Ravi, P17): the forward charge-off scan is now
-- anchored at 31 Jan 2025, NOT 1 Jan 2025. The three windows on
-- future_co.co_dt_future are:
--   CO8  : co_dt_future >= DATE '2025-01-31' AND co_dt_future < DATE '2025-09-30'
--   CO10 : co_dt_future >= DATE '2025-01-31' AND co_dt_future < DATE '2025-11-30'
--   CO12 : co_dt_future >= DATE '2025-01-31' AND co_dt_future < DATE '2026-01-31'
-- This anchor applies to the co_8m/co_10m/co_12m COUNTS and to every
-- charge-off dollar and window-restricted balance column below.
-- CO-COUNT TIE-OUTS RE-BASELINED: the tier-14 recorded counts (b15_co_8m 853,
-- b15_co_10m 943, b15_co_12m 1,006) were computed on the OLD Jan-01 anchor.
-- With the anchor moved to Jan-31 they will NOT reproduce exactly. Expect
-- values near the old numbers but not equal. These are NOTES, not STOP rules.
-- SURVIVING HARD STOPS (do NOT depend on the CO window; any mismatch = STOP,
-- route to the keeper before trusting the new columns):
--   * b15_accounts across ALL class rows sums to 2,177 EXACTLY.
--   * deceased-flag 'a' rows sum to 210 EXACTLY.
--   * b15_jan_eom_balance sums to ~9,752,424 (rounding tolerance).
-- NEW COLUMNS (appended to the final SELECT, in existing group order):
--   * b15_co8_amt / b15_co10_amt / b15_co12_amt = round(sum(chrgoff_amt), 0)
--     over the CO8 / CO10 / CO12 window (GROSS charge-off dollars).
--   * b15_jan_bal_co8 / b15_jan_bal_co10 / b15_jan_bal_co12 = round(sum of the
--     Jan EOM balance, 0) restricted to the CO8 / CO10 / CO12 window, i.e. the
--     same l.eom_bal already summed unconditionally into b15_jan_eom_balance,
--     but limited to accounts whose earliest 2025 charge-off lands in-window.
-- All three windows (8/10/12) now carry a count, a gross CO dollar, and a
-- window-restricted Jan EOM balance.
-- CO DOLLAR SOURCE: future_co finds co_dt_future = min(co_dt) = the account's
-- earliest 2025 charge-off date, and pairs to it co_amt = min_by(chrgoff_amt,
-- chrgoff_dt) FILTER (WHERE chrgoff_dt IS NOT NULL), so co_amt is the
-- charge-off amount on the SAME row as the earliest co_dt. Output is
-- account-level (JOIN ledger l on acct_key, LEFT JOIN future_co), one ledger
-- row per account, so a plain sum is correct here; no per-episode dedup needed.
-- PLAUSIBILITY BOUNDS (state only, not tie-outs): every balance / CO-dollar
-- sum >= 0; per row b15_co12_amt >= b15_co10_amt >= b15_co8_amt and
-- b15_jan_bal_co12 >= b15_jan_bal_co10 >= b15_jan_bal_co8 (each wider CO window
-- contains the narrower one); each b15_jan_bal_coN <= b15_jan_eom_balance for
-- that row; coN_amt sits in the same order of magnitude as the balance on
-- those accounts (charge-off amount ~ balance scale, not equal).
-- RESOURCE DISCIPLINE: no new base-table scans. co_amt rides the same forward
-- scan already in place (a second aggregate over rows already read);
-- chrgoff_amt / acct_bal_amt are columns on rows the scans already touch.
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
      AND eff_dt >= '20241201' AND eff_dt < '20250401'
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
    SELECT j.extnl_acct_id, j.max_bucket, j.eom_bucket, j.eom_bal, j.co_dt,
           j.eom_cpc,
           p.max_bucket AS prev_max_bucket,
           m2.eom_bucket AS m2_eom_bucket,
           m2.co_dt AS m2_co_dt,
           (m2.extnl_acct_id IS NOT NULL) AS has_m2_row,
           m3.eom_bucket AS m3_eom_bucket,
           m3.co_dt AS m3_co_dt,
           (m3.extnl_acct_id IS NOT NULL) AS has_m3_row
    FROM (SELECT * FROM monthly WHERE ym = '202501') j
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202412') p
      ON j.extnl_acct_id = p.extnl_acct_id
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202502') m2
      ON j.extnl_acct_id = m2.extnl_acct_id
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202503') m3
      ON j.extnl_acct_id = m3.extnl_acct_id
),
ledger AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           eom_bal,
           CASE
             WHEN eom_cpc IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                               'AA3','AC3','AM3','AA4','AC4','AM4') THEN 'a. AA'
             WHEN eom_cpc IN ('BGC','BGM','CGM','GMR')              THEN 'b. GM'
             WHEN eom_cpc IN ('FBS','IBS','U1C','U2C','U3C')        THEN 'c. Bronco'
             ELSE 'd. others'
           END AS pc_class,
           CASE
             WHEN m2_co_dt >= DATE '2025-02-01' AND m2_co_dt < DATE '2025-03-01' THEN 'co'
             WHEN NOT has_m2_row THEN 'gone'
             ELSE cast(m2_eom_bucket AS varchar)
           END AS feb_position,
           CASE
             WHEN m3_co_dt >= DATE '2025-02-01' AND m3_co_dt < DATE '2025-04-01' THEN 'co'
             WHEN NOT has_m3_row THEN 'gone'
             ELSE cast(m3_eom_bucket AS varchar)
           END AS mar_position
    FROM base
    WHERE eom_bucket = 1
      AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
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
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'pay|paid|payment|settle|payment plan|arrangement|work something out'))
               AS pay_utts,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'passed away|death certificate|executor|deceased|calling on behalf'))
               AS deceased_utts
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM ep) d
      ON t.contactid = d.contactid
     AND t.effdt >= '2025-01-01' AND t.effdt < '2025-02-02'
    WHERE t.content IS NOT NULL
    GROUP BY 1
),
leaklist AS (
    SELECT e.acct_key,
           max(CASE WHEN coalesce(x.deceased_utts, 0) > 0 THEN 1 ELSE 0 END) AS deceased_flag
    FROM ep e
    LEFT JOIN tx x ON e.contactid = x.contactid
    GROUP BY 1
    HAVING max(e.captured) = 0
       AND max(CASE WHEN e.captured = 0 AND coalesce(x.pay_utts, 0) > 0
                    THEN 1 ELSE 0 END) = 1
)
SELECT CASE WHEN k.deceased_flag = 1 THEN 'a. deceased or estate language'
            ELSE 'b. no deceased language' END AS b15_deceased_flag,
       l.pc_class AS b15_pc_class,
       l.feb_position AS b15_feb_position,
       l.mar_position AS b15_mar_position,
       count(*) AS b15_accounts,
       count_if(f.co_dt_future >= DATE '2025-01-31'
                AND f.co_dt_future < DATE '2025-09-30') AS b15_co_8m,
       count_if(f.co_dt_future >= DATE '2025-01-31'
                AND f.co_dt_future < DATE '2025-11-30') AS b15_co_10m,
       count_if(f.co_dt_future >= DATE '2025-01-31'
                AND f.co_dt_future < DATE '2026-01-31') AS b15_co_12m,
       round(sum(l.eom_bal), 0) AS b15_jan_eom_balance,
       round(sum(CASE WHEN f.co_dt_future >= DATE '2025-01-31'
                       AND f.co_dt_future < DATE '2025-09-30'
                      THEN f.co_amt END), 0) AS b15_co8_amt,
       round(sum(CASE WHEN f.co_dt_future >= DATE '2025-01-31'
                       AND f.co_dt_future < DATE '2025-11-30'
                      THEN f.co_amt END), 0) AS b15_co10_amt,
       round(sum(CASE WHEN f.co_dt_future >= DATE '2025-01-31'
                       AND f.co_dt_future < DATE '2026-01-31'
                      THEN f.co_amt END), 0) AS b15_co12_amt,
       round(sum(CASE WHEN f.co_dt_future >= DATE '2025-01-31'
                       AND f.co_dt_future < DATE '2025-09-30'
                      THEN l.eom_bal END), 0) AS b15_jan_bal_co8,
       round(sum(CASE WHEN f.co_dt_future >= DATE '2025-01-31'
                       AND f.co_dt_future < DATE '2025-11-30'
                      THEN l.eom_bal END), 0) AS b15_jan_bal_co10,
       round(sum(CASE WHEN f.co_dt_future >= DATE '2025-01-31'
                       AND f.co_dt_future < DATE '2026-01-31'
                      THEN l.eom_bal END), 0) AS b15_jan_bal_co12
FROM leaklist k
JOIN ledger l ON k.acct_key = l.acct_key
LEFT JOIN future_co f
  ON trim(cast(f.extnl_acct_id AS varchar)) = k.acct_key
GROUP BY 1, 2, 3, 4
ORDER BY 1, 2, 3, 4
