-- Tier 5 | Charge-off dollars: funnel-end callers vs other callers vs non-callers
-- BACKWARD / AUDIT FRAME: this reads the realized loss ledger backwards.
-- Size and reconcile on it; never read rates off it (outcome-conditioned).
-- Of the accounts that charged off between 2024-08 and 2026-02, how many (and
-- how many dollars) had a funnel-end episode on the path - a delinquent-month
-- inbound call with payment/plan language and no payment within 30 days, at most
-- 8 months before the charge-off? The funnel-end share is the slice of realized
-- losses the inbound channel touched and did not capture.
-- pct_dollars_bk_dcsd_fraud = share of each group's dollars on accounts whose
-- status text reads bankruptcy / deceased-like / fraud (FR/ST/BK regex, a
-- partial proxy) - the unsolvable slice the client nets out of its own
-- policy-loss number; quote the funnel-end dollars net of it.
-- Calls observed 2024-07 .. 2026-01 (calls need a same-month snapshot and a
-- 30-day payment runway inside the account edge). The last charge-off months
-- see a shorter call lookback, and unmatched calls (no account id) cannot
-- count as calls here - group c is 'no MATCHED inbound call observed'.
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
           try_cast(chrgoff_amt AS double) AS co_amt,
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(paymt_last_dt, '%d%b%Y') AS date))) AS pay_dt,
           CASE
             WHEN regexp_like(upper(acct_status_rsn_txt), 'FR') THEN 'FR'
             WHEN regexp_like(upper(acct_status_rsn_txt), 'ST') THEN 'ST'
             WHEN regexp_like(upper(acct_status_rsn_txt), 'BK') THEN 'BK'
             ELSE NULL
           END AS status_flag
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= DATE '2024-07-01'
      AND date(date_parse(eff_dt, '%Y%m%d')) < DATE '2026-03-01'
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, min(co_dt) AS co_dt,
           max(co_amt) AS co_amt, max(pay_dt) AS pay_dt,
           max(status_flag) AS status_flag
    FROM snap GROUP BY 1, 2
),
monthly2 AS (
    SELECT extnl_acct_id, m, bucket, pay_dt,
           min(co_dt) OVER (PARTITION BY extnl_acct_id) AS acct_co_dt,
           lead(pay_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_pay_dt
    FROM monthly
),
co_accts AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           min(co_dt) AS co_dt, max(co_amt) AS co_amt,
           max(status_flag) AS status_flag
    FROM monthly
    WHERE co_dt IS NOT NULL
    GROUP BY 1
    HAVING min(co_dt) >= DATE '2024-08-01' AND min(co_dt) < DATE '2026-03-01'
),
inb AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key, contactid,
           "date" AS call_dt, initiationtimestamp
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2024-07-01' AND "date" < DATE '2026-02-01'
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
    SELECT e.acct_key, e.contactid, e.call_dt,
           s.pay_dt, s.next_pay_dt,
           CASE WHEN s.bucket >= 1
                 AND (s.acct_co_dt IS NULL OR s.acct_co_dt > e.call_dt)
                THEN 1 ELSE 0 END AS is_delq
    FROM episodes e
    JOIN monthly2 s
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
    JOIN (SELECT DISTINCT contactid FROM matched WHERE is_delq = 1) d
      ON t.contactid = d.contactid
    GROUP BY 1
),
funnel_end AS (
    SELECT m.acct_key, m.call_dt
    FROM matched m
    JOIN tx t ON m.contactid = t.contactid
    WHERE m.is_delq = 1
      AND t.pay_utts > 0
      AND NOT ((m.pay_dt IS NOT NULL
                AND m.pay_dt >= m.call_dt
                AND m.pay_dt <= date_add('day', 30, m.call_dt))
            OR (m.next_pay_dt IS NOT NULL
                AND m.next_pay_dt >= m.call_dt
                AND m.next_pay_dt <= date_add('day', 30, m.call_dt)))
),
per_co AS (
    SELECT c.acct_key, c.co_amt, c.status_flag,
           max(CASE WHEN f.call_dt IS NOT NULL AND f.call_dt < c.co_dt
                     AND c.co_dt <= date_add('month', 8, f.call_dt)
                    THEN 1 ELSE 0 END) AS funnel_end_caller,
           max(CASE WHEN e.call_dt IS NOT NULL AND e.call_dt < c.co_dt
                     AND c.co_dt <= date_add('month', 8, e.call_dt)
                    THEN 1 ELSE 0 END) AS any_caller
    FROM co_accts c
    LEFT JOIN episodes e ON c.acct_key = e.acct_key
    LEFT JOIN funnel_end f ON c.acct_key = f.acct_key
    GROUP BY 1, 2, 3
)
SELECT CASE
         WHEN funnel_end_caller = 1 THEN 'a. funnel-end caller (leaked episode on the path)'
         WHEN any_caller = 1 THEN 'b. other inbound caller (8 months before charge-off)'
         ELSE 'c. no matched inbound call observed'
       END AS caller_group,
       count(*) AS accounts,
       round(sum(co_amt), 0) AS chargeoff_dollars,
       round(avg(co_amt), 0) AS avg_chargeoff,
       round(100.0 * sum(co_amt) / sum(sum(co_amt)) OVER (), 1) AS pct_of_dollars,
       round(100.0 * sum(CASE WHEN status_flag IS NOT NULL THEN co_amt ELSE 0 END)
             / greatest(sum(co_amt), 1), 1) AS pct_dollars_bk_dcsd_fraud
FROM per_co
GROUP BY 1
ORDER BY 1
