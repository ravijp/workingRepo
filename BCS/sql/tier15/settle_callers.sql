-- Tier 15 | CALLER-GAP DIAGNOSTIC, REDESIGNED 2026-07-13 (P17), replacing
-- the book-level C/B2/A ladder (Ravi's point 5: that ladder returned raw
-- book totals ~831k and never SAID why our 9,389 ex-AA January callers
-- disagree with the SAS flag's 11,154). This version is an ACCOUNT-LEVEL
-- SET DIFFERENCE on ONE comparable population, tagging every disagreeing
-- account with the reason it disagrees.
--
-- HOW THE TWO SIDES ARE RECONSTRUCTED, both in Athena, both on the same
-- January call table (contactcenter_bdp_db.call):
--   OURS      = INBOUND, January, id present, effdt load-date cap
--               (< 2025-02-02), business-card call rows EXCLUDED. This is
--               the exact filter set b7/b14 use.
--   SAS-STYLE = the export behind zenon.aws_call_accts_jan_mar_25 as
--               characterized in CQ-11 / the export code (Ravi 2026-07-13):
--               INBOUND, call_month='M1' (January), id present, NO
--               producttype exclusion, NO effdt cap. Reconstructed here so
--               the SAS flag's own membership can be diffed account by
--               account without leaving Athena.
-- Both are then INTERSECTED with the ex-AA January bucket-1 ledger
-- (b14_exaa's 189,146 population), because the 9,389-vs-11,154 comparison
-- is a comparison of CALLERS AMONG THAT LEDGER. Accounts that call but are
-- not in the ledger are reported separately (row f) so nothing is hidden.
--
-- OUTPUT: one row per (membership, reason). membership is 'in both' /
-- 'ours only' / 'sas-style only'; reason names WHY an account sits in a
-- one-sided class:
--   - ours-only accounts exist only if OURS is broader on some axis; by
--     construction OURS is a strict subset of SAS-STYLE among ledger
--     accounts (we add filters, we never relax any), so 'ours only' should
--     be ZERO. A nonzero 'ours only' = an unmodeled difference: investigate.
--   - sas-style-only accounts are the gap. Each is tagged with the FIRST
--     applicable reason: 'business-card call only' (every January inbound
--     call from this account was on a business-card product, so our
--     exclusion drops it) > 'dropped by effdt cap' (its only January
--     inbound calls arrived in the data on/after 2025-02-02) > 'other'
--     (neither of the two named filters explains it; residual to chase).
-- The two named reasons are the two filter differences the export code
-- exposed; the residual 'other' is what neither explains and is the true
-- open question.
--
-- PRE-REGISTERED TIE-OUTS (STOP RULE):
--   - (in both) + (sas-style only) accounts, summed, = the SAS-style
--     caller count among the ex-AA ledger. This should land at / very near
--     11,154 (the CQ-1 flagged ex-AA count) IF the SAS flag's perimeter
--     matches this reconstruction. A large miss (> a few hundred) means the
--     SAS flag population is not what the export code implies: STOP, route
--     to the keeper (do not attribute reasons off a mismatched base).
--   - (in both) + (ours only) accounts, summed, = 9,389 EXACTLY (our ex-AA
--     January caller count of record, b7_exaa/b14_exaa). Any miss = our
--     caller reconstruction here drifted from b7/b14: STOP.
--   - 'ours only' = 0 (see above). Nonzero = investigate before trusting
--     the reason split.
-- RESOURCE DISCIPLINE: one January account scan (ledger), two passes over
-- the January call table (one per side; both are single scans with the same
-- WHERE shape); no window functions over the call rows.
WITH snap AS (
    SELECT extnl_acct_id,
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
           try_cast(chrgoff_dt AS date) AS co_dt,
           clnt_prdct_cd
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id,
           max_by(bucket, eff_dt) AS eom_bucket,
           min(co_dt) AS co_dt,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc
    FROM snap GROUP BY 1
),
ledger AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key
    FROM monthly
    WHERE eom_bucket = 1
      AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
      AND (eom_cpc IS NULL OR trim(eom_cpc) = ''
           OR eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                               'AA3','AC3','AM3','AA4','AC4','AM4',
                               'BGC','BGM','CGM','GMR',
                               'FBS','IBS','U1C','U2C','U3C'))
),
-- One scan of the January inbound call table, carrying the two filter axes
-- (business-card membership, effdt-cap membership) as per-row flags so both
-- sides and both reasons are derivable in a single pass.
jan_calls AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key,
           CASE WHEN coalesce(cast(producttype AS varchar), '') = 'BUSINESS_CARD'
                THEN 1 ELSE 0 END AS is_biz,
           CASE WHEN effdt >= '2025-01-01' AND effdt < '2025-02-02'
                THEN 1 ELSE 0 END AS within_effdt_cap
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2025-01-01' AND "date" < DATE '2025-02-01'
      AND acctid IS NOT NULL
),
-- Collapse to one row per calling account with the axes it EVER satisfied.
-- max(...) over the flags = "the account has at least one call that is X".
call_acct AS (
    SELECT acct_key,
           max(CASE WHEN is_biz = 0 THEN 1 ELSE 0 END)             AS has_nonbiz_call,
           max(CASE WHEN within_effdt_cap = 1 THEN 1 ELSE 0 END)   AS has_incap_call,
           max(CASE WHEN is_biz = 0 AND within_effdt_cap = 1
                    THEN 1 ELSE 0 END)                             AS has_ours_call
    FROM jan_calls
    WHERE acct_key IS NOT NULL AND acct_key <> ''
    GROUP BY 1
),
-- Restrict to the ex-AA ledger; classify membership and reason.
classed AS (
    SELECT l.acct_key,
           -- OURS: at least one January inbound call that is non-business-card
           -- AND within the effdt cap (the exact b7/b14 filter set).
           coalesce(c.has_ours_call, 0)                            AS in_ours,
           -- SAS-STYLE: at least one January inbound call at all (id present),
           -- no producttype exclusion, no effdt cap (the export logic).
           CASE WHEN c.acct_key IS NOT NULL THEN 1 ELSE 0 END       AS in_sas,
           c.has_nonbiz_call, c.has_incap_call
    FROM ledger l
    LEFT JOIN call_acct c ON c.acct_key = l.acct_key
)
SELECT CASE
         WHEN in_ours = 1 AND in_sas = 1 THEN '1. in both'
         WHEN in_ours = 1 AND in_sas = 0 THEN '2. ours only (should be 0)'
         WHEN in_ours = 0 AND in_sas = 1 THEN '3. sas-style only'
         ELSE '4. neither (ledger non-caller)'
       END AS membership,
       CASE
         WHEN in_ours = 0 AND in_sas = 1 AND coalesce(has_nonbiz_call, 0) = 0
              THEN 'a. business-card calls only (our exclusion drops it)'
         WHEN in_ours = 0 AND in_sas = 1 AND coalesce(has_incap_call, 0) = 0
              THEN 'b. all January calls arrived after the effdt cap'
         WHEN in_ours = 0 AND in_sas = 1
              THEN 'c. other (neither named filter explains it) [residual]'
         ELSE '-'
       END AS reason,
       count(*) AS accounts
FROM classed
GROUP BY 1, 2
ORDER BY 1, 2
