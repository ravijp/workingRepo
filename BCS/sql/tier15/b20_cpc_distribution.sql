-- Tier 15 | NEW QUERY, drafted 2026-07-13, REVISED 2026-07-13 (P17).
-- Purpose: the product-class (CPC) distribution of the CLEANED January
-- bucket-1 ledger, with NO ex-AA filter, so AA / GM / Bronco appear as their
-- OWN rows (Ravi's point 3: the earlier _exaa variant filtered AA/GM/Bronco
-- out before grouping, leaving only CoBrand / Biz / PLCC / OTHER; this
-- version keeps every account so the full CPC decomposition is visible).
-- Reuses b14_ledger's snap/monthly/ledger CTEs verbatim (January window,
-- eff_dt >= '20250101' AND < '20250201'; cleaned rule co_dt IS NULL OR
-- co_dt >= 2025-01-01; eom_cpc kept). The ONLY change vs b14's ledger is
-- that the ex-AA exclusion predicate is REMOVED, so the base is the full
-- cleaned 204,323-account ledger. Final SELECT groups eom_cpc under the full
-- client CPC mapping (AA tested first, in the order given; ELSE 'OTHER' also
-- catches NULL / blank cpc and the bare code '2').
-- Credit-limit column CONFIRMED against the live Athena schema
-- (athena-filter-check-transcription-2026-07-13.md, DESCRIBE fmt_acct_c row
-- 74: cr_lmt_origl_amt decimal(17,3)). The earlier [VERIFY] is CLOSED; no
-- DESCRIBE needed before running.
-- PRE-REGISTERED TIE-OUTS (STOP RULE):
--   - accounts summed across ALL cpc_class rows = 204,323 EXACTLY (the
--     cleaned ledger; matches athena-filter-check round 3: AA 15,177 +
--     others 189,146).
--   - the AA row = 15,177 EXACTLY, balance ~73,744,823 (P-C round 3).
--   - the OTHER-plus-non-AA rows sum to 189,146 accounts / ~457,943,985
--     balance (the ex-AA ledger).
--   - jan_eom_balance summed across all rows = ~531,688,808 (tolerance $5).
--   - GM and Bronco rows: zero or absent (CQ-8 / P-C: AA-and-others only at
--     code 1). A NONZERO GM/Bronco row is NOT an error here (this is the
--     unfiltered ledger) but IS a surprise vs the SAS read: flag, do not stop.
-- RESOURCE DISCIPLINE: single January account scan (snap/monthly), same as
-- b14; no additional table scans.
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
           clnt_prdct_cd,
           try_cast(cr_lmt_origl_amt AS double) AS cr_lmt_origl_amt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           max_by(bal, eff_dt) AS eom_bal,
           min(co_dt) AS co_dt,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc,
           max_by(cr_lmt_origl_amt, eff_dt) AS eom_cr_lmt_origl_amt
    FROM snap GROUP BY 1, 2
),
ledger AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           eom_bal,
           eom_cpc,
           eom_cr_lmt_origl_amt
    FROM monthly
    WHERE ym = '202501'
      AND eom_bucket = 1
      AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
    -- NOTE: no ex-AA exclusion here (intentional; point 3). This is the
    -- full cleaned 204,323-account ledger.
)
SELECT CASE
         WHEN eom_cpc IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                           'AA3','AC3','AM3','AA4','AC4','AM4')     THEN 'AA'
         WHEN eom_cpc IN ('BGC','BGM','CGM','GMR')                 THEN 'GM'
         WHEN eom_cpc IN ('FBS','IBS','U1C','U2C','U3C')           THEN 'Bronco'
         WHEN eom_cpc IN ('BHA','BJT','BJC','BFR','BWY','BBB')     THEN 'Biz'
         WHEN eom_cpc IN ('GAP','GP2','ONV','ON2','BRP','BR2','ATH','AT2',
                           'GPC','G2C','ONC','O2C','BRC','B2C','ATC','A2C')
                                                                     THEN 'CoBrand'
         WHEN eom_cpc IN ('8GP','8ON','8BR','8AT','9GP','9ON','9BR','9AT')
                                                                     THEN 'PLCC'
         ELSE 'OTHER'
       END AS cpc_class,
       count(*) AS accounts,
       round(sum(eom_bal), 0) AS jan_eom_balance,
       round(sum(eom_cr_lmt_origl_amt), 0) AS orig_credit_limit_total,
       round(avg(eom_cr_lmt_origl_amt), 0) AS orig_credit_limit_avg
FROM ledger
GROUP BY 1
ORDER BY 2 DESC
