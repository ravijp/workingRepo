-- Tier 11 | The leak list followed forward: 2,394 leaked-intent accounts
-- Account grain. Population: the ledger's leaked-intent accounts (b8's intent
-- lexicon on January episodes, no clean payment within 30 days on ANY
-- January episode, month-end classes b + c). Expected tie-outs: 2,394
-- accounts, 880 charged off within 8 months; the 12-month charge-off on this
-- list is a NEW number. Rows: deceased flag (m4's estate lexicon, kept as a
-- column so both views render) x Feb-EOM position x Mar-EOM position, with
-- accounts, CO at 8/10/12 months, and the Jan-31 EOM balance. Feb/Mar
-- positions follow b6's convention: 'co' when the charge-off date lands in
-- the window, 'gone' when the account has no row that month. ONE transcript
-- pass (intent + deceased counted in the same scan), aggregates only.
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
           try_cast(chrgoff_dt AS date) AS co_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20241201' AND eff_dt < '20250401'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           max_by(bal, eff_dt) AS eom_bal,
           min(co_dt) AS co_dt
    FROM snap GROUP BY 1, 2
),
base AS (
    SELECT j.extnl_acct_id, j.max_bucket, j.eom_bucket, j.eom_bal, j.co_dt,
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
       l.feb_position AS b15_feb_position,
       l.mar_position AS b15_mar_position,
       count(*) AS b15_accounts,
       count_if(f.co_dt_future >= DATE '2025-01-01'
                AND f.co_dt_future < DATE '2025-09-01') AS b15_co_8m,
       count_if(f.co_dt_future >= DATE '2025-01-01'
                AND f.co_dt_future < DATE '2025-11-01') AS b15_co_10m,
       count_if(f.co_dt_future >= DATE '2025-01-01'
                AND f.co_dt_future < DATE '2026-01-01') AS b15_co_12m,
       round(sum(l.eom_bal), 0) AS b15_jan_eom_balance
FROM leaklist k
JOIN ledger l ON k.acct_key = l.acct_key
LEFT JOIN future_co f
  ON trim(cast(f.extnl_acct_id AS varchar)) = k.acct_key
GROUP BY 1, 2, 3
ORDER BY 1, 2, 3
