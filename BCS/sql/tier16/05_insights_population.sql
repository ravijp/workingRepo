-- Tier 16 | INSIGHTS 05: the January base (walkthrough section 1 + 1.2).
-- Numbered blocks; run ONE block at a time. Every block is a standalone
-- statement over the LAYER tables.
--
-- DUAL-MODE RULE (applies to every block in this file):
--   MODE 1 (tables exist): run the block as-is; its placeholder CTE(s) at the
--     top point at the saved uc2_t16_* tables. Fill in <schema>.
--   MODE 2 (no table access): DELETE the placeholder CTE(s) and paste the
--     referenced layer file's WITH-chain in their place, in dependency order
--     (00 before 01 before 02/03/04), per the README stitch recipe: drop each
--     upstream file's final bare SELECT or wrap it as a CTE named exactly like
--     the placeholder it replaces, keep ONE `WITH` at the very top, join
--     chains with commas. Blocks here need only `populations` (= layer 01,
--     which itself needs layer 00) - a stitch already proven by the verified
--     tier-14/15 originals these blocks reproduce.
-- After ANY stitch or rebuild, run block 9.0 (the anchor sweep in
-- 09_insights_diagnostics.sql) BEFORE quoting any number. Any miss = STOP.

-- ============================================================================
-- INSIGHT 5.1: the two-side population walk, AWS cells (walkthrough section 1
-- table). One row of count_ifs over the January account layer.
-- Logic: the walk cells are pure flag counts on layer 01 (cleanup rule, ex-AA
-- rule, CPC mapping all live in the layer, verbatim from the verified kit).
-- EXPECTED (walkthrough section 1 / filter-check record; STOP on miss):
--   book_jan_accounts    ~ 45,105,572  (accounts with >= 1 January snapshot)
--   raw_b1_eom           =    207,006  (bucket 1 at 31 Jan, before cleanup)
--   precleanup_drop      =      2,683  (bucket-1 rows charged off before 2025)
--   ledger_all           =    204,323 EXACTLY; ledger_all_bal ~ $531,688,811
--   aa_ledger            =     15,177 EXACTLY; aa_ledger_bal  ~ $73,744,823
--   ledger_exaa          =    189,146 EXACTLY; ledger_exaa_bal = $457,943,987
--     (balance tolerance ~$5; account counts are the exact checks)
-- The SAS cells of the same walk (45,362,305 / 202,479 / 15,226 / 186,412 /
-- $454.2M / ECL $93.5M) are SAS-side and out of scope here (README section 5).
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
)
SELECT count(*)                                        AS book_jan_accounts,
       count_if(eom_bucket = 1)                        AS raw_b1_eom,
       count_if(eom_bucket = 1 AND NOT cleaned)        AS precleanup_drop,
       count_if(in_ledger_all)                         AS ledger_all,
       round(sum(CASE WHEN in_ledger_all THEN eom_bal END), 0)  AS ledger_all_bal,
       count_if(in_ledger_all AND cpc_class = 'AA')    AS aa_ledger,
       round(sum(CASE WHEN in_ledger_all AND cpc_class = 'AA'
                      THEN eom_bal END), 0)            AS aa_ledger_bal,
       count_if(in_ledger_exaa)                        AS ledger_exaa,
       round(sum(CASE WHEN in_ledger_exaa THEN eom_bal END), 0) AS ledger_exaa_bal
FROM populations

-- ============================================================================
-- INSIGHT 5.2: CPC distribution of the cleaned ledger, b20 face (walkthrough
-- 1.2). NO ex-AA filter: the full 204,323 by product class, with balance and
-- original credit limit. Logic verbatim from tier15/b20_cpc_distribution.sql
-- (the CPC mapping is layer 01's cpc_class column, AA tested first).
-- EXPECTED (bridge round 9, EXACT; STOP on miss):
--   OTHER   78,027 / $354,209,096 / avg limit $5,481
--   PLCC    57,664 / $21,046,624  / avg limit $794
--   CoBrand 53,455 / $82,688,265  / avg limit $4,184
--   AA      15,177 / $73,744,823  / avg limit $8,737
--   accounts sum = 204,323 EXACTLY; GM / Bronco / Biz rows absent (zero).
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
)
SELECT cpc_class,
       count(*) AS accounts,
       round(sum(eom_bal), 0) AS jan_eom_balance,
       round(sum(eom_cr_lmt_origl_amt), 0) AS orig_credit_limit_total,
       round(avg(eom_cr_lmt_origl_amt), 0) AS orig_credit_limit_avg
FROM populations
WHERE in_ledger_all
GROUP BY 1
ORDER BY 2 DESC

-- ============================================================================
-- INSIGHT 5.3: charge-off shares of the cleaned ledger at 8 / 10 / 12 months,
-- ORIGINAL Jan-01-anchored windows (walkthrough section 1 cross-check row:
-- 19.8% / 23.5% / 26.4% vs SAS 20.4% / 24.4% / 27.5%). Computed inline from
-- co_dt_future because layer 01's co_8m/10m/12m flags are 31-Jan-anchored
-- (the round-9 re-baseline); the walkthrough section-1 shares use the OLD
-- anchor, so this block recomputes the old windows verbatim from b9:
--   CO8 [2025-01-01, 2025-09-01) CO10 [.., 2025-11-01) CO12 [.., 2026-01-01).
-- EXPECTED (b9 / bridge round 3, derived shares): pct_co8 ~ 19.8 /
-- pct_co10 ~ 23.5 / pct_co12 ~ 26.4 on the 204,323 base.
-- NOTE: the FULL b9 grid (classes a/b/c x non-caller/captured/leaked, before
-- the AA filter) is NOT derivable from the layers: layer 03/04 exclude AA
-- accounts from the transcript/capture path by design. To reproduce it, run
-- ../tier11/b9_cohort_outcomes.sql verbatim (standalone, no layer dependency).
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
)
SELECT count(*) AS ledger_all,
       count_if(co_dt_future >= DATE '2025-01-01' AND co_dt_future < DATE '2025-09-01') AS co8_accounts,
       count_if(co_dt_future >= DATE '2025-01-01' AND co_dt_future < DATE '2025-11-01') AS co10_accounts,
       count_if(co_dt_future >= DATE '2025-01-01' AND co_dt_future < DATE '2026-01-01') AS co12_accounts,
       round(100.0 * count_if(co_dt_future >= DATE '2025-01-01' AND co_dt_future < DATE '2025-09-01') / count(*), 1) AS pct_co8,
       round(100.0 * count_if(co_dt_future >= DATE '2025-01-01' AND co_dt_future < DATE '2025-11-01') / count(*), 1) AS pct_co10,
       round(100.0 * count_if(co_dt_future >= DATE '2025-01-01' AND co_dt_future < DATE '2026-01-01') / count(*), 1) AS pct_co12
FROM populations
WHERE in_ledger_all

-- ============================================================================
-- INSIGHT 5.4: what happened by February, share form (walkthrough 1.1 first
-- check row + section 2 insight line: cure 55.6% / stay 8.0% / roll 36.0% on
-- the ex-AA 189,146). Rollup of layer 01's feb_position_b14 (the full grid
-- with balance and CO dollars is INSIGHT 6.1/6.2).
-- EXPECTED (b14_exaa, bridge round 8; STOP on miss):
--   cured 105,215 (55.6%) / stayed 15,054 (8.0%) / rolled 68,093 (36.0%) /
--   deeper 111 (0.1%) / charged off in Feb 673 (0.4%); accounts sum 189,146.
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
)
SELECT feb_position_b14,
       count(*) AS accounts,
       round(100.0 * count(*) / sum(count(*)) OVER (), 1) AS share_pct
FROM populations
WHERE in_ledger_exaa
GROUP BY 1
ORDER BY 1
