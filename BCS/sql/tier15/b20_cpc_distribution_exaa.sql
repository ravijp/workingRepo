-- Tier 15 | NEW QUERY, drafted 2026-07-13.
-- Purpose: the product-class (CPC) distribution of the final ex-AA ledger
-- (189,146). Reuses b14_exaa.sql's snap/monthly/ledger CTEs verbatim
-- (January window, eff_dt >= '20250101' AND < '20250201'; cleaned rule;
-- eom_cpc kept), with the standard NULL-safe 23-code ex-AA exclusion applied
-- so the base is exactly the 189,146-account ledger. Final SELECT re-groups
-- eom_cpc under the FULL client CPC mapping (AA tested first, in the order
-- given; ELSE 'OTHER' also catches NULL/blank cpc).
-- Credit-limit column: CR_LMT_ORIGL_AMT per data-dictionary.md line 118/430
-- (fmt_acct_c "Balances, APR, credit limit" section; SAS name = dictionary
-- name, no renaming noted for this table). [VERIFY: run DESCRIBE
-- fmt_acct_dba.fmt_acct_c; and confirm cr_lmt_origl_amt is present with this
-- exact casing before trusting orig_credit_limit_* below] -- one-line check:
--   DESCRIBE "fmt_acct_dba"."fmt_acct_c";  -- grep output for cr_lmt_origl_amt
-- PRE-REGISTERED TIE-OUTS (STOP RULE):
--   - accounts summed across all cpc_class rows = 189,146 EXACTLY.
--   - jan_eom_balance summed across all rows ~= 457,943,987 (tolerance $5).
--   - AA, GM, and Bronco rows are ZERO by construction (the ex-AA exclusion
--     already removed every account carrying those codes at EOM). Any
--     nonzero accounts in AA/GM/Bronco = STOP, route to the keeper.
--   - Biz appears only via its 6 non-AA-overlapping codes (BHA/BJT/BJC/BFR/
--     BWY/BBB); BA5/BC5 are excluded upstream as part of the AA code set and
--     will not appear under Biz here.
-- RESOURCE DISCIPLINE: single January account scan (snap/monthly), same as
-- b14_exaa; no additional table scans.
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
      AND (eom_cpc IS NULL OR trim(eom_cpc) = ''
           OR eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                               'AA3','AC3','AM3','AA4','AC4','AM4',
                               'BGC','BGM','CGM','GMR',
                               'FBS','IBS','U1C','U2C','U3C'))
)
SELECT CASE
         WHEN eom_cpc IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                           'AA3','AC3','AM3','AA4','AC4','AM4')     THEN 'AA'
         WHEN eom_cpc IN ('BGC','BGM','CGM','GMR')                 THEN 'GM'
         WHEN eom_cpc IN ('FBS','IBS','U1C','U2C','U3C')           THEN 'Bronco'
         WHEN eom_cpc IN ('BA5','BC5','BHA','BJT','BJC','BFR','BWY','BBB')
                                                                     THEN 'Biz'
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
ORDER BY 1
