-- Tier 16 | LAYER 00: the account layer. ONE scan of fmt_acct_dba.fmt_acct_c
-- over a parameterized month window -> one row per account per month.
-- Everything downstream (01/04) reads accounts from here, never from the
-- base table again (the two documented exceptions: 01's future_co forward
-- charge-off scan, which needs a 12-month horizon beyond this window and is
-- heavily pruned by chrgoff_dt IS NOT NULL, and 04's day-grain call-day
-- snapshot, which cannot be derived from a monthly layer).
--
-- DEPENDS ON: nothing (base table only).
-- FEEDS: 01_populations (slices by ym), 04_outcomes (payment-date lead).
--
-- PARAMETERS (edit these literals together; they are the ONLY knobs):
--   window start  '20241201'  (month BEFORE the anchor month, for carried-in)
--   window end    '20250401'  (exclusive; anchor month + 2 for Feb/Mar reads)
-- The anchor month itself ('202501') is a 01-layer parameter.
--
-- TIE-OUT ANCHORS (check via 01 after any change here; STOP on any miss):
--   * cleaned January bucket-1 ledger, all products : 204,323 accounts
--   * cleaned January bucket-1 ledger, ex-AA        : 189,146 accounts,
--     Jan EOM balance $457,943,987 (rounding tolerance ~$5)
--   * touched-bucket-1 universe (ex-AA, cleaned)    : 724,848 accounts
-- Any anchor miss after restructuring = STOP, re-verify before trusting
-- any new number.
--
-- COLUMN NOTES (all logic character-faithful to the verified tier-14/15 kit):
--   bucket CASE       = the ten-rung past-due-amount ladder, verbatim.
--   eom_*             = max_by(col, eff_dt) over the month = month-end value.
--   mth_co_dt         = min(chrgoff_dt) within the month (the cleanup rule
--                       tests the ANCHOR month's value: mth_co_dt IS NULL OR
--                       mth_co_dt >= DATE '2025-01-01').
--   mth_co_amt        = charge-off amount on the earliest in-month CO row
--                       (informational; the canonical forward CO dollar is
--                       01's future_co, the verified shape).
--   first_dq_dt       = first snapshot day with bucket >= 1 (runway band).
--   first_b1_dt       = first snapshot day with bucket = 1 (touched-B1 flag).
--   pay/auto/nsf dt   = max last-payment / autopay / NSF dates in the month,
--                       dual-format parse verbatim from the verified kit.
--
-- WITH TABLE ACCESS (Databricks or CTAS), uncomment ONE of:
-- CREATE TABLE <schema>.uc2_t16_00_acct_monthly AS
-- (Databricks: spark.sql("""<this WITH block>""").write.saveAsTable("<schema>.uc2_t16_00_acct_monthly"))
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
           try_cast(chrgoff_amt AS double) AS co_amt,
           clnt_prdct_cd,
           try_cast(cr_lmt_origl_amt AS double) AS cr_lmt_origl_amt,
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(try_cast(paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS pay_dt,
           coalesce(try_cast(atmtc_paymt_last_dt AS date),
                    try(cast(date_parse(try_cast(atmtc_paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS auto_dt,
           coalesce(try_cast(nsf_last_paymt_dt AS date),
                    try(cast(date_parse(try_cast(nsf_last_paymt_dt AS varchar), '%d%b%Y') AS date))) AS nsf_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20241201' AND eff_dt < '20250401'   -- PARAM: month window
),
acct_monthly AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           max_by(bal, eff_dt) AS eom_bal,
           min(co_dt) AS mth_co_dt,
           min_by(co_amt, co_dt) AS mth_co_amt,
           min(CASE WHEN bucket >= 1 THEN eff_dt END) AS first_dq_dt,
           min(CASE WHEN bucket = 1 THEN eff_dt END) AS first_b1_dt,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc,
           max_by(cr_lmt_origl_amt, eff_dt) AS eom_cr_lmt_origl_amt,
           max(pay_dt) AS pay_dt,
           max(auto_dt) AS auto_dt,
           max(nsf_dt) AS nsf_dt
    FROM snap
    GROUP BY extnl_acct_id, ym
)
SELECT * FROM acct_monthly
