-- Tier 16 | INSIGHTS 09: diagnostics and guards (walkthrough 1.1, Appendix
-- 2.15 / 2.16, and the layer anchor sweep). Numbered blocks; run ONE at a time.
--
-- DUAL-MODE RULE (every block):
--   MODE 1 (tables exist): run as-is against the uc2_t16_* tables; fill <schema>.
--   MODE 2 (no tables): delete the placeholder CTE(s) and paste the layer
--     chains per the README stitch recipe (`populations` = 00 -> 01;
--     `calls` = 02, standalone; `outcomes` = the full 04 stitch, heavy).
--
-- ============================================================================
-- INSIGHT 9.0: THE ANCHOR SWEEP. Run this FIRST after ANY stitch, rebuild,
-- window change, or engine move, BEFORE quoting any number from any other
-- block. Any miss = STOP: find the cause; quote nothing downstream.
-- EXPECTED (EXACT unless noted):
--   ledger_all 204,323 | ledger_exaa 189,146 | ledger_exaa_bal $457,943,987
--   (tolerance ~$5) | aa_ledger 15,177 | touched_b1 724,848
--   (classes a 464,023 / b 186,714 / c 69,513 / d 4,598).
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
)
SELECT count_if(in_ledger_all)  AS ledger_all,
       count_if(in_ledger_exaa) AS ledger_exaa,
       round(sum(CASE WHEN in_ledger_exaa THEN eom_bal END), 0) AS ledger_exaa_bal,
       count_if(in_ledger_all AND cpc_class = 'AA') AS aa_ledger,
       count_if(touched_b1) AS touched_b1,
       count_if(touched_b1_class LIKE 'a.%') AS touched_a,
       count_if(touched_b1_class LIKE 'b.%') AS touched_b,
       count_if(touched_b1_class LIKE 'c.%') AS touched_c,
       count_if(touched_b1_class LIKE 'd.%') AS touched_d
FROM populations

-- ============================================================================
-- INSIGHT 9.1: the caller-gap set difference, settle_callers face
-- (walkthrough 1.1 caller row; bridge round 9). Both caller definitions
-- reconstructed from layer 02's per-row flags (is_biz, within_effdt_cap -
-- carried as COLUMNS for exactly this purpose), intersected with the ex-AA
-- ledger. OURS = >= 1 January inbound call that is non-business-card AND
-- within the effdt cap; SAS-STYLE = >= 1 January inbound call at all (id
-- present). Reason CASE verbatim from tier15/settle_callers.sql.
-- EXPECTED (bridge round 9, EXACT; STOP on miss):
--   '1. in both' = 9,389; '4. neither' = 179,757; rows 2 and 3 ABSENT (zero);
--   9,389 + 179,757 = 189,146. The ~1,765 extra SAS-flagged accounts live
--   OUTSIDE this ledger (population membership, not caller logic).
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
),
calls AS (
    SELECT * FROM "<schema>"."uc2_t16_02_episodes"
),
call_acct AS (
    SELECT acct_key,
           max(CASE WHEN is_biz = 0 THEN 1 ELSE 0 END)           AS has_nonbiz_call,
           max(CASE WHEN within_effdt_cap = 1 THEN 1 ELSE 0 END) AS has_incap_call,
           max(CASE WHEN is_biz = 0 AND within_effdt_cap = 1
                    THEN 1 ELSE 0 END)                           AS has_ours_call
    FROM calls
    WHERE acct_key IS NOT NULL AND acct_key <> ''
    GROUP BY 1
),
classed AS (
    SELECT p.acct_key,
           coalesce(c.has_ours_call, 0)                       AS in_ours,
           CASE WHEN c.acct_key IS NOT NULL THEN 1 ELSE 0 END AS in_sas,
           c.has_nonbiz_call, c.has_incap_call
    FROM populations p
    LEFT JOIN call_acct c ON c.acct_key = p.acct_key
    WHERE p.in_ledger_exaa
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

-- ============================================================================
-- INSIGHT 9.2: the ledger caller pair, two independent constructions
-- (the 'two independent queries produce the same 9,389' claim, walkthrough
-- 1.1). Construction A counts from layer 04 (episode grain); construction B
-- counts from layer 02 x layer 01 with no 04 involvement. Both rows must
-- match each other AND the anchors.
-- EXPECTED: episodes 11,262 / callers 9,389 EXACTLY on BOTH rows.
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
),
calls AS (
    SELECT * FROM "<schema>"."uc2_t16_02_episodes"
),
outcomes AS (
    SELECT * FROM "<schema>"."uc2_t16_04_outcomes"
)
SELECT 'A. via layer 04 (outcomes)' AS construction,
       count(*) AS episodes,
       count(DISTINCT acct_key) AS callers
FROM outcomes
WHERE in_ledger_exaa
UNION ALL
SELECT 'B. via layer 02 x layer 01',
       count(*),
       count(DISTINCT c.acct_key)
FROM calls c
JOIN populations p ON p.acct_key = c.acct_key
WHERE c.is_episode_std = 1
  AND p.in_ledger_exaa
ORDER BY 1

-- ============================================================================
-- INSIGHT 9.3: the January verification-join hole, b18 face (walkthrough
-- Appendix 2.15; the ~28% undercount evidence, reference R2). NOT DERIVABLE
-- FROM THE LAYERS: layer 02 keeps only call rows WITH an account id (acctid
-- IS NOT NULL), and the join-hole measurement is precisely about the rows
-- WITHOUT one. RUN ../tier11/b18_join_gap.sql VERBATIM (standalone; leg
-- grain, no product exclusion, weekly).
-- EXPECTED (bridge round 7): weekly not-joinable shares 27.5 / 27.6 / 27.6 /
-- 28.3 / 28.7% (stable); e.g. week 2025-01-06: 90,964 not joinable / 13,196
-- joined-behind / 225,208 joined-current.
--
-- ============================================================================
-- INSIGHT 9.4: the same-customer guard, b16_v2 face (walkthrough Appendix
-- 2.16 check 1; reference R3). NOT DERIVABLE FROM THE LAYERS: it scores
-- January-MARCH episodes (layer 02 is January-only) against each episode's
-- own next-month bucket, and its March capture gate needs April payment
-- dates (outside layer 00's window). RUN
-- ../tier11/b16_v2_within_cohort_contrast.sql VERBATIM (standalone; before
-- the AA filter, ledger accounts, Jan-Mar calls).
-- EXPECTED (bridge round 7): paid-30d episodes 3,802 vs no-pay 3,137 on the
-- same 2,411 accounts; current next month 45.2% vs 27.4%; deeper 14.2% vs
-- 39.8%. (The platform-wide version, x4 - 31,804 accounts, 56%/42% and
-- 10%/36% - is story-run era: out of scope, see README section 5.)
