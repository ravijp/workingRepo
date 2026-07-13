-- Tier 15 | NEW DIAGNOSTIC PAIR (drafted 2026-07-13; keeper-corrected same
-- day): isolates two candidate causes of the 9,389 (Athena, January, ex-AA
-- ledger callers) vs 11,154 (SAS caller flag on the 186,412 slice)
-- difference: the BUSINESS_CARD product-type exclusion, and the Jan-only vs
-- Jan-Mar call window.
--
-- IMPORTANT SCOPE NOTE (keeper correction): these two statements count
-- distinct calling accounts over the WHOLE call table, NOT joined to the
-- bucket-1 ledger. Their outputs are book-level caller counts (expected in
-- the hundreds of thousands) and are NOT directly comparable to 9,389 or
-- 11,154. They isolate the DIRECTION and RELATIVE SIZE of each effect:
--   (A) vs the same statement with the exclusion re-added = the
--       business-card share of January inbound callers;
--   (B) vs (A-with-exclusion) = the multiplier from widening one month to
--       three, at book level.
-- The population-joined settle that directly tests the flag-builder's
-- January-only claim runs on the SAS side, one line, per the CQ-11 note
-- ("call_month='01' gives a January-only rerun"):
--
--   /* SAS, not Athena: */
--   -- proc sql;
--   --   select count(distinct wf.extnl_acct_id) as jan_only_inb_flagged
--   --   from zenon.waterfall_acct_coll_v1_202501 wf
--   --   where wf.DLNQT_CD_M1='1' and wf.cpc_M1='others'
--   --     and wf.CHRGOFF_RSN_M1 in ('blank','PLY')
--   --     and input(wf.extnl_acct_id, BEST12.) in
--   --         (select extnl_acct_id from zenon.aws_call_accts_jan_mar_25
--   --           where call_month='01' and initiationmethod='INBOUND');
--   -- quit;
--
--   Expected: if the January-only count lands near 9,389 (same order,
--   allowing the slice-vs-ex-AA population difference and the unshared
--   business-card handling), the flag as run (11,154) was Jan-Mar and the
--   walkthrough's reading stands; if it stays near 11,154, the import
--   table itself is January-only despite its name and grain, and the
--   record needs a correction.
--
-- PLAUSIBILITY BOUNDS: (A) and (B) are book-level counts, far above
-- 11,154; (B) > (A-with-exclusion) by construction (wider window);
-- (A) >= the same January statement with the exclusion re-added.
-- RESOURCE DISCIPLINE: each statement is a single scan over
-- contactcenter_bdp_db.call filtered to its own date window, matching the
-- scan shape of b14_exaa's `inb` CTE exactly (no join, no window
-- function); nothing here touches fmt_acct_c.

-- ============================================================
-- (A) January-window inbound distinct accounts, WITHOUT the
-- BUSINESS_CARD exclusion (same "date"/effdt window as b14_exaa's inb,
-- exclusion clause removed).
-- ============================================================
SELECT count(DISTINCT trim(cast(acctid AS varchar))) AS settle_a_accounts
FROM "contactcenter_bdp_db"."call"
WHERE initiationmethod = 'INBOUND'
  AND "date" >= DATE '2025-01-01' AND "date" < DATE '2025-02-01'
  AND effdt >= '2025-01-01' AND effdt < '2025-02-02';

-- ============================================================
-- (B) Jan-Mar-window inbound distinct accounts, WITH the BUSINESS_CARD
-- exclusion (same exclusion clause as b14_exaa's inb, window widened to
-- "date" < 2025-04-01 / effdt < 2025-04-02).
-- ============================================================
SELECT count(DISTINCT trim(cast(acctid AS varchar))) AS settle_b_accounts
FROM "contactcenter_bdp_db"."call"
WHERE initiationmethod = 'INBOUND'
  AND "date" >= DATE '2025-01-01' AND "date" < DATE '2025-04-01'
  AND effdt >= '2025-01-01' AND effdt < '2025-04-02'
  AND coalesce(cast(producttype AS varchar), '') <> 'BUSINESS_CARD';
