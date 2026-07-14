-- Tier 16 | INSIGHTS 07: the calls, the words, and the work list
-- (walkthrough section 4 + Appendix 2.1/2.2/2.3/2.5/2.4/2.7).
-- Numbered blocks; run ONE block at a time.
--
-- DUAL-MODE RULE (every block):
--   MODE 1 (tables exist): run as-is against the uc2_t16_* tables; fill <schema>.
--   MODE 2 (no tables): delete the placeholder CTE(s), paste the layer chains
--     per the README stitch recipe. `populations` = 00 -> 01. `calls` = 02
--     (standalone). `signals` = 00 -> 01 -> 02 -> 03 (contains the ONE
--     transcript pass). `outcomes` = the full 00 -> 01 -> 02 -> 03 -> 04
--     stitch wrapped as an `outcomes` CTE. Blocks that reference BOTH
--     `outcomes` and `signals` (7.5) still paste 03 ONCE: the 04 stitch
--     already contains 03's chain ending in a `signals` CTE, so both names
--     resolve from one paste. NEVER paste the transcript scan twice into one
--     statement (the round-9 m4 OOM lesson); run heavy stitches one at a time.
-- After any stitch or rebuild: run block 9.0 (anchor sweep) first. Miss = STOP.

-- ============================================================================
-- INSIGHT 7.1: who calls - the January classes and their call rates, b7 face
-- (walkthrough Appendix 2.1). Population: every ex-AA cleaned account
-- delinquent at some point in January (max_bucket >= 1), classed by the
-- verbatim b7 bridge CASE recomputed from layer 01's bucket columns.
-- EXPECTED (b7_exaa, bridge round 8; STOP on miss):
--   a. entrant cured:  301,827 accounts / 44,447 callers / 51,670 episodes / 14.7%
--   b. entrant still:  139,654 / 6,627 / 8,059 / 4.7%
--   c. older stock:     49,492 / 2,762 / 3,203 / 5.6%
--   d. other delinquent: 472,914 / 38,934 / 45,200 / 8.2%
--   check: b + c = 189,146 (the ledger) EXACTLY.
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
),
calls AS (
    SELECT * FROM "<schema>"."uc2_t16_02_episodes"
),
per_acct AS (
    SELECT acct_key, count(*) AS n_episodes
    FROM calls
    WHERE is_episode_std = 1
    GROUP BY 1
)
SELECT CASE
         WHEN p.max_bucket = 1 AND coalesce(p.prev_max_bucket, 0) = 0
              AND p.eom_bucket = 0
           THEN 'a. month-MAX B1 entrant, cured by EOM'
         WHEN p.max_bucket = 1 AND coalesce(p.prev_max_bucket, 0) = 0
              AND p.eom_bucket >= 1
           THEN 'b. month-MAX B1 entrant, still DQ1 at EOM'
         WHEN p.eom_bucket = 1
              AND NOT (p.max_bucket = 1 AND coalesce(p.prev_max_bucket, 0) = 0)
           THEN 'c. EOM bucket 1 stock, not a month-max-B1 entrant'
         ELSE 'd. other delinquent in month (month-MAX >= 2, EOM <> 1)'
       END AS bridge_class,
       count(*) AS class_accounts,
       count(e.acct_key) AS jan_inbound_callers,
       coalesce(sum(e.n_episodes), 0) AS jan_inbound_episodes,
       round(100.0 * count(e.acct_key) / count(*), 1) AS pct_accounts_calling
FROM populations p
LEFT JOIN per_acct e ON e.acct_key = p.acct_key
WHERE p.max_bucket >= 1 AND p.cleaned AND p.is_exaa
GROUP BY 1
ORDER BY 1

-- ============================================================================
-- INSIGHT 7.2: the funnel by class, b8 face, stages 2 / 4 / 5 / 6
-- (walkthrough 4.1 = the b+c row sums; Appendix 2.2 = by class). Episode
-- grain from layer 04, classed by the b7/b8 bridge CASE via populations.
-- Stage 6 uses the ORIGINAL Jan-01-anchored CO8 window (b8 verbatim:
-- [2025-01-01, 2025-09-01)), computed inline from co_dt_future.
-- OUT OF LAYER SCOPE: stage 3 ('episode has a transcript', 49,597 / 7,609 /
-- 2,957; ledger 10,566 / 8,822 accounts). Layer 03 keeps only contactids with
-- matching language, so transcript EXISTENCE is not in the layers. For stage 3
-- run ../tier14/b8_exaa.sql verbatim (standalone; its tx CTE is a transcript
-- existence pass, proven at this scale).
-- EXPECTED (b8_exaa, bridge round 8; STOP on miss):
--   stage 2 episodes/accounts: a 51,670/44,447; b 8,059/6,627; c 3,203/2,762
--   stage 4: a 39,654/35,531; b 4,984/4,345; c 2,251/2,013
--   stage 5: a 3,491/3,303; b 1,928/1,742; c 422/386
--   stage 6: a 34/27; b 710/652; c 149/130
--   ledger sums (b + c): 11,262/9,389; 7,235/6,358; 2,350/2,128; 859/782.
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
),
outcomes AS (
    SELECT * FROM "<schema>"."uc2_t16_04_outcomes"
),
cohort_ep AS (
    SELECT o.acct_key, o.contactid, o.captured, o.pay_f,
           (o.co_dt_future >= DATE '2025-01-01'
            AND o.co_dt_future < DATE '2025-09-01') AS co_8m_orig,
           CASE
             WHEN p.max_bucket = 1 AND coalesce(p.prev_max_bucket, 0) = 0
                  AND p.eom_bucket = 0
               THEN 'a. month-MAX B1 entrant, cured by EOM'
             WHEN p.max_bucket = 1 AND coalesce(p.prev_max_bucket, 0) = 0
                  AND p.eom_bucket >= 1
               THEN 'b. month-MAX B1 entrant, still DQ1 at EOM'
             WHEN p.eom_bucket = 1
                  AND NOT (p.max_bucket = 1 AND coalesce(p.prev_max_bucket, 0) = 0)
               THEN 'c. EOM bucket 1 stock, not a month-max-B1 entrant'
             ELSE 'd. other'
           END AS bridge_class
    FROM outcomes o
    JOIN populations p ON p.acct_key = o.acct_key
    WHERE p.max_bucket >= 1 AND p.cleaned AND p.is_exaa
)
SELECT '2. called inbound in January' AS b8_stage, bridge_class,
       count(*) AS episodes, count(DISTINCT acct_key) AS accounts
FROM cohort_ep WHERE bridge_class NOT LIKE 'd.%' GROUP BY 2
UNION ALL SELECT '4. customer payment/plan language', bridge_class,
       count(*), count(DISTINCT acct_key)
FROM cohort_ep WHERE bridge_class NOT LIKE 'd.%' AND pay_f > 0 GROUP BY 2
UNION ALL SELECT '5. of 4: leaked (no clean payment in 30d)', bridge_class,
       count(*), count(DISTINCT acct_key)
FROM cohort_ep WHERE bridge_class NOT LIKE 'd.%' AND pay_f > 0 AND captured = 0 GROUP BY 2
UNION ALL SELECT '6. of 5: charged off within 8 months (Jan-01 anchor)', bridge_class,
       count(*), count(DISTINCT acct_key)
FROM cohort_ep WHERE bridge_class NOT LIKE 'd.%' AND pay_f > 0 AND captured = 0
  AND coalesce(co_8m_orig, false) GROUP BY 2
ORDER BY 1, 2

-- ============================================================================
-- INSIGHT 7.3: the callers classed by what happened after the call, account
-- level with balance and CO12 dollars (walkthrough 4.2). Layer 01 accounts,
-- caller class from layer 04 (non-caller = no episode row).
-- EXPECTED (b14_exaa row sums + the 14-July dollar run; STOP on miss):
--   a. non-caller:    179,757 / $424.3M  / cured 55.5% / rolled 36.2% / co12$ ~$144.7M
--   b. captured:        6,029 / $20.3M   / 71.6% / 14.5% / ~$4.3M
--   c. leaked-intent:   1,929 / $8,266,309 / 31.7% / 63.1% / ~$4.5M
--   d. other-caller:    1,431 / $5.0M    / 31.8% / 64.0% / ~$3.1M
--   sums: 189,146 accounts / $457,943,987 / ~$156.6M EXACTLY.
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
),
outcomes AS (
    SELECT * FROM "<schema>"."uc2_t16_04_outcomes"
),
callers AS (
    SELECT acct_key, max_by(caller_class, contactid) AS caller_class
    FROM outcomes
    WHERE in_ledger_exaa
    GROUP BY 1
)
SELECT coalesce(k.caller_class, 'a. non-caller') AS caller_class,
       count(*) AS accounts,
       round(sum(p.eom_bal), 0) AS jan_eom_balance,
       round(100.0 * count_if(p.feb_position_b14 LIKE 'a.%') / count(*), 1) AS pct_cured_feb,
       round(100.0 * count_if(p.feb_position_b14 LIKE 'c.%') / count(*), 1) AS pct_rolled_feb,
       count_if(p.co_12m) AS co_12m_accounts,
       round(sum(CASE WHEN p.co_12m THEN p.co_amt END), 0) AS co12_amt
FROM populations p
LEFT JOIN callers k ON k.acct_key = p.acct_key
WHERE p.in_ledger_exaa
GROUP BY 1
ORDER BY 1

-- ============================================================================
-- INSIGHT 7.4: the language groups on the ledger's conversations, m4 face
-- with balance and CO dollars (walkthrough 4.3; round-9 m4_exaa_bal grid).
-- Grain: in_ledger_exaa episodes from layer 04; v2 language_group; balances
-- and CO dollars deduped ONE ROW PER ACCOUNT within each group (an account
-- with episodes in two groups appears in both rows: NEVER sum the balance
-- column down for a total). Two CO12 account columns: the 31-Jan-anchored
-- layer flag (ties the round-9 grid) and the original Jan-01 anchor (ties
-- walkthrough 4.3's CO12 counts).
-- EXPECTED (m4_exaa_bal, round 9; STOP on miss): episodes
--   a 498 / b 1,374 / c 5,164 / d 442 / e 79 / f 274 / g 3,431, sum 11,262
--   EXACTLY; accounts a 469 .. g 3,082; co_12m (31-Jan) a 329 / b 236 /
--   c 1,020 / d 169 / e 37 / f 77 / g 744; co_12m_orig (Jan-01) a 331 /
--   b 243 / c 1,040 / d 184 / e 37 / f 77 / g 763 (walkthrough 4.3);
--   leaked-intent accounts a 164 / c 1,356; balances a $1,675,990 ..
--   c $17,173,381; co12_amt a $1,244,785 / c $5,672,621; every group
--   monotone co8 <= co10 <= co12.
WITH outcomes AS (
    SELECT * FROM "<schema>"."uc2_t16_04_outcomes"
),
led AS (
    SELECT * FROM outcomes WHERE in_ledger_exaa
),
acct_grp AS (   -- one row per (language_group, acct_key); flags are account-constant
    SELECT language_group, acct_key,
           max(jan_eom_bal) AS jan_eom_bal,
           bool_or(co_8m)  AS co_8m,
           bool_or(co_10m) AS co_10m,
           bool_or(co_12m) AS co_12m,
           max(co_amt) AS co_amt
    FROM led
    GROUP BY 1, 2
),
grp_bal AS (
    SELECT language_group,
           round(sum(jan_eom_bal), 0) AS jan_eom_balance,
           round(sum(CASE WHEN co_8m  THEN co_amt END), 0) AS co8_amt,
           round(sum(CASE WHEN co_10m THEN co_amt END), 0) AS co10_amt,
           round(sum(CASE WHEN co_12m THEN co_amt END), 0) AS co12_amt,
           round(sum(CASE WHEN co_8m  THEN jan_eom_bal END), 0) AS jan_bal_co8,
           round(sum(CASE WHEN co_10m THEN jan_eom_bal END), 0) AS jan_bal_co10,
           round(sum(CASE WHEN co_12m THEN jan_eom_bal END), 0) AS jan_bal_co12
    FROM acct_grp
    GROUP BY 1
)
SELECT e.language_group AS m4_group,
       count(*) AS episodes,
       count(DISTINCT e.acct_key) AS accounts,
       round(100.0 * sum(e.captured) / count(*), 1) AS pct_paid_30d,
       count(DISTINCT CASE WHEN coalesce(e.co_8m, false)  THEN e.acct_key END) AS co_8m_accounts,
       count(DISTINCT CASE WHEN coalesce(e.co_10m, false) THEN e.acct_key END) AS co_10m_accounts,
       count(DISTINCT CASE WHEN coalesce(e.co_12m, false) THEN e.acct_key END) AS co_12m_accounts,
       count(DISTINCT CASE WHEN e.co_dt_future >= DATE '2025-01-01'
                            AND e.co_dt_future < DATE '2026-01-01'
                           THEN e.acct_key END) AS co_12m_accounts_orig,
       count(DISTINCT CASE WHEN e.leaked_acct THEN e.acct_key END) AS leaked_intent_accounts,
       g.jan_eom_balance, g.co8_amt, g.co10_amt, g.co12_amt,
       g.jan_bal_co8, g.jan_bal_co10, g.jan_bal_co12
FROM led e
JOIN grp_bal g ON g.language_group = e.language_group
GROUP BY 1, 10, 11, 12, 13, 14, 15, 16
ORDER BY 1

-- ============================================================================
-- INSIGHT 7.5: language groups by January class, b10 face (walkthrough 4.3
-- second table + Appendix 2.5). The ORIGINAL v1 six-group ladder (no deceased
-- group), recomputed verbatim from layer 03's raw flags (deceased episodes
-- fall wherever their payment flags put them, as in b10). CO12 per episode at
-- the ORIGINAL Jan-01 anchor. Cohort classes a/b/c only (b10 drops d).
-- EXPECTED (b10_exaa, bridge round 8; STOP on miss): class-b episodes
--   promise 981 / payment-talk 3,769 / plan 234 / hardship 65 / dispute 237 /
--   none 2,773, sum 8,059 EXACTLY; class totals a 51,670 / c 3,203; capture
--   and CO12 percentages per the recorded 18-row grid (e.g. b-promise 68 /
--   16.1; b-none 47.6 / 29.6).
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
),
outcomes AS (
    SELECT * FROM "<schema>"."uc2_t16_04_outcomes"
),
signals AS (
    SELECT * FROM "<schema>"."uc2_t16_03_signals"
),
cohort_ep AS (
    SELECT o.acct_key, o.contactid, o.captured,
           (o.co_dt_future >= DATE '2025-01-01'
            AND o.co_dt_future < DATE '2026-01-01') AS co_12m_orig,
           CASE
             WHEN coalesce(x.promise_f, 0) > 0 THEN 'a. future-dated promise'
             WHEN coalesce(x.pay_f, 0) > 0
                  AND coalesce(x.plan_f, 0) = 0 THEN 'b. payment talk, no promise'
             WHEN coalesce(x.plan_f, 0)    > 0 THEN 'c. plan or settlement talk'
             WHEN coalesce(x.hard_f, 0)    > 0 THEN 'd. hardship talk'
             WHEN coalesce(x.dispute_f, 0) > 0 THEN 'e. dispute or fraud talk'
             ELSE 'f. no payment-related language'
           END AS language_group_v1,
           CASE
             WHEN p.max_bucket = 1 AND coalesce(p.prev_max_bucket, 0) = 0
                  AND p.eom_bucket = 0 THEN 'a. entrant, cured by EOM'
             WHEN p.max_bucket = 1 AND coalesce(p.prev_max_bucket, 0) = 0
                  AND p.eom_bucket >= 1 THEN 'b. entrant, still DQ1 at EOM'
             WHEN p.eom_bucket = 1 THEN 'c. EOM bucket 1 stock (older)'
             ELSE 'd. other'
           END AS bridge_class
    FROM outcomes o
    JOIN populations p ON p.acct_key = o.acct_key
    LEFT JOIN signals x ON x.contactid = o.contactid
    WHERE p.max_bucket >= 1 AND p.cleaned AND p.is_exaa
)
SELECT language_group_v1, bridge_class,
       count(*) AS episodes,
       count(DISTINCT acct_key) AS accounts,
       round(100.0 * sum(captured) / count(*), 1) AS pct_captured,
       round(100.0 * count_if(co_12m_orig) / count(*), 1) AS pct_co12_per_episode
FROM cohort_ep
WHERE bridge_class NOT LIKE 'd.%'
GROUP BY 1, 2
ORDER BY 1, 2

-- ============================================================================
-- INSIGHT 7.6: the work list W - build steps and follow-forward (walkthrough
-- 4.4 + Appendix 2.4, ex-AA rows). Strict rule = layer 04's leaked_acct
-- (>= 1 leaked conversation AND no captured conversation that month);
-- deceased routing = deceased_acct; W = w_flag. Two statements: 7.6a the
-- build steps, 7.6b the follow-forward grid.
-- OUT OF LAYER SCOPE: the strict list's AA rows (248 accounts; the 2,177
-- before-AA total). Layers 03/04 exclude AA accounts from the transcript path
-- by design. For the full 2,177 grid run ../tier15/b15_exaa_bal.sql verbatim.
-- EXPECTED 7.6a (b15/b14/m4, rounds 8-9; STOP on miss):
--   strict ex-AA list 1,929; deceased routed 164 ($575,422; CO12_orig 122);
--   W = 1,765 / $7,690,886 EXACTLY / CO8_orig 637 (36.1%) / CO12_orig 776
--   (44.0%) / co12_amt (31-Jan) $4,098,105 / jan_bal_co12 ~ $4.53M.
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
),
outcomes AS (
    SELECT * FROM "<schema>"."uc2_t16_04_outcomes"
),
strict_acct AS (   -- one row per strict-leak ex-AA-ledger account
    SELECT o.acct_key,
           max(o.deceased_acct) AS deceased_acct,
           max(o.w_flag) AS w_flag
    FROM outcomes o
    WHERE o.leaked_acct AND o.in_ledger_exaa
    GROUP BY 1
),
priced AS (
    SELECT s.acct_key, s.deceased_acct, s.w_flag,
           p.eom_bal, p.co_amt, p.co_dt_future, p.co_8m, p.co_10m, p.co_12m,
           p.feb_pos, p.mar_pos
    FROM strict_acct s
    JOIN populations p ON p.acct_key = s.acct_key
)
SELECT step, accounts, jan_eom_balance, co8_orig, co12_orig, co12_amt, jan_bal_co12
FROM (
    SELECT '1. strict leaked-intent, ex-AA ledger' AS step,
           count(*) AS accounts,
           round(sum(eom_bal), 0) AS jan_eom_balance,
           count_if(co_dt_future >= DATE '2025-01-01' AND co_dt_future < DATE '2025-09-01') AS co8_orig,
           count_if(co_dt_future >= DATE '2025-01-01' AND co_dt_future < DATE '2026-01-01') AS co12_orig,
           round(sum(CASE WHEN co_12m THEN co_amt END), 0) AS co12_amt,
           round(sum(CASE WHEN co_12m THEN eom_bal END), 0) AS jan_bal_co12
    FROM priced
    UNION ALL
    SELECT '2. deceased-estate, routed out', count(*), round(sum(eom_bal), 0),
           count_if(co_dt_future >= DATE '2025-01-01' AND co_dt_future < DATE '2025-09-01'),
           count_if(co_dt_future >= DATE '2025-01-01' AND co_dt_future < DATE '2026-01-01'),
           round(sum(CASE WHEN co_12m THEN co_amt END), 0),
           round(sum(CASE WHEN co_12m THEN eom_bal END), 0)
    FROM priced WHERE deceased_acct = 1
    UNION ALL
    SELECT '3. W, the work list (no deceased language)', count(*), round(sum(eom_bal), 0),
           count_if(co_dt_future >= DATE '2025-01-01' AND co_dt_future < DATE '2025-09-01'),
           count_if(co_dt_future >= DATE '2025-01-01' AND co_dt_future < DATE '2026-01-01'),
           round(sum(CASE WHEN co_12m THEN co_amt END), 0),
           round(sum(CASE WHEN co_12m THEN eom_bal END), 0)
    FROM priced WHERE w_flag
)
ORDER BY step

-- INSIGHT 7.6b: the follow-forward grid on the strict ex-AA list (Appendix
-- 2.4's 'others' rows): deceased flag x Feb position x Mar position, with CO
-- windows at BOTH anchors and the 31-Jan dollar columns.
-- EXPECTED: accounts sum = 1,929 (= 2,177 - 248 AA); the deceased rows sum to
-- 164; cells reproduce the round-9 b15_exaa_bal 'd. others' rows (31-Jan
-- columns) and the walkthrough Appendix 2.4 'others' rows (orig-anchor CO
-- counts, e.g. No/others/2/3 = 764 accounts / CO12_orig 634).
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
),
outcomes AS (
    SELECT * FROM "<schema>"."uc2_t16_04_outcomes"
),
strict_acct AS (
    SELECT o.acct_key, max(o.deceased_acct) AS deceased_acct
    FROM outcomes o
    WHERE o.leaked_acct AND o.in_ledger_exaa
    GROUP BY 1
)
SELECT CASE WHEN s.deceased_acct = 1 THEN 'a. deceased language'
            ELSE 'b. no deceased language' END AS deceased_flag,
       p.feb_pos, p.mar_pos,
       count(*) AS accounts,
       round(sum(p.eom_bal), 0) AS jan_eom_balance,
       count_if(p.co_dt_future >= DATE '2025-01-01' AND p.co_dt_future < DATE '2025-09-01') AS co8_orig,
       count_if(p.co_dt_future >= DATE '2025-01-01' AND p.co_dt_future < DATE '2026-01-01') AS co12_orig,
       count_if(p.co_8m)  AS co_8m,
       count_if(p.co_12m) AS co_12m,
       round(sum(CASE WHEN p.co_8m  THEN p.co_amt END), 0) AS co8_amt,
       round(sum(CASE WHEN p.co_12m THEN p.co_amt END), 0) AS co12_amt,
       round(sum(CASE WHEN p.co_12m THEN p.eom_bal END), 0) AS jan_bal_co12
FROM strict_acct s
JOIN populations p ON p.acct_key = s.acct_key
GROUP BY 1, 2, 3
ORDER BY 1, 2, 3

-- ============================================================================
-- INSIGHT 7.7: accounts that TOUCHED bucket 1 in January, b21 face
-- (walkthrough 4.2's wider view). Population: layer 01's touched_b1 flag
-- (first_b1_dt set, cleaned, ex-AA), classed by touched_b1_class; callers and
-- episodes from layer 02. CO windows are the 31-Jan layer flags.
-- EXPECTED (b21, bridge round 9; STOP on miss):
--   a. current at 31 Jan: 464,023 / 68,444 callers / 79,389 eps / 15% /
--      $910,932,039 / co12 20,075 / co12$ $65,481,294
--   b. bucket 1 at 31 Jan: 186,714 / 9,330 / 11,197 / 5% / $455,458,836 /
--      47,967 / $156,567,263
--   c. bucket 2+ at 31 Jan: 69,513 / 2,854 / 3,421 / 4% / $209,009,131 /
--      39,695 / $135,679,384
--   d. charged off in January: 4,598 / 122 / 132 / 3% / $6,146,344 / 54 / $185,165
--   ADDITIVE RECONCILIATION (the corrected stop rule): row b + the 2,432
--   eom_bucket=1 Jan-CO accounts inside row d = 189,146 / 9,389 / 11,262
--   EXACTLY; full population 724,848.
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
),
calls AS (
    SELECT * FROM "<schema>"."uc2_t16_02_episodes"
),
per_acct AS (
    SELECT acct_key, count(*) AS n_episodes
    FROM calls
    WHERE is_episode_std = 1
    GROUP BY 1
)
SELECT p.touched_b1_class AS class,
       count(*) AS accounts,
       count(e.acct_key) AS callers,
       coalesce(sum(e.n_episodes), 0) AS episodes,
       round(100.0 * count(e.acct_key) / count(*), 1) AS pct_accounts_calling,
       round(sum(p.eom_bal), 0) AS jan_eom_balance,
       count_if(p.co_8m)  AS co_8m,
       count_if(p.co_10m) AS co_10m,
       count_if(p.co_12m) AS co_12m,
       round(sum(CASE WHEN p.co_8m  THEN p.co_amt END), 0) AS co8_amt,
       round(sum(CASE WHEN p.co_10m THEN p.co_amt END), 0) AS co10_amt,
       round(sum(CASE WHEN p.co_12m THEN p.co_amt END), 0) AS co12_amt
FROM populations p
LEFT JOIN per_acct e ON e.acct_key = p.acct_key
WHERE p.touched_b1
GROUP BY 1
ORDER BY 1

-- ============================================================================
-- INSIGHT 7.8: the year, month by month, lx4 face (walkthrough 4.5 + Appendix
-- 2.7). NOT DERIVABLE FROM THE LAYERS: the layer windows are anchored on
-- January (00 spans 2024-12..2025-03; 02/03 are January-only), and lx4 needs
-- twelve call months of episodes and transcripts (Jul 2024 - Jun 2025).
-- RUN ../tier15/lx4_exaa_bal.sql VERBATIM (standalone; no layer dependency;
-- heavy - a year of transcript scanning, one sitting).
-- EXPECTED (bridge rounds 8-9, EXACT; STOP on miss): column sums
--   no_payment_30d 118,069; no_payment_30d_net_dec 113,192;
--   chargeoff_8m 35,490; chargeoff_8m_net_dec 32,019;
--   mechanics 72,671; leaked-despite-mechanics 4,525;
--   behind-in-month 1,034,682; January row 108,009 / 12,061 / 11,578 /
--   3,455 / 3,125 / 8,164 / 457; leaked_bal sums $345.6M (Jan $34.4M);
--   co8 dollars $13.1M / net-dec $12.2M. July 2024 is the recorded boundary
--   artifact: never quote it alone.

-- ============================================================================
-- INSIGHT 7.9: call timing and who raises payment first, b11 face
-- (walkthrough 6.1 'act on the early signal' + Appendix 2.6). NOT DERIVABLE
-- FROM THE LAYERS: the seven signals need utterance TIMING (beginmillis
-- positions, agent-vs-customer order, agent offers), which layer 03 does not
-- carry (presence flags only, by design). RUN ../tier14/b11_exaa.sql VERBATIM
-- (standalone; its own single transcript pass with timing aggregates - one
-- sitting, never combined with another transcript stitch).
-- EXPECTED (b11_exaa, bridge round 8; STOP on miss): signals 1-3 partition
-- each class's episodes EXACTLY (51,670 / 8,059 / 3,203); signals 4+5 sum to
-- each class's intent episodes (39,654 / 4,984 / 2,251); class-b cells:
-- early-intent 3,976 @ 63.5% vs late 1,008 @ 52.6%; intent+offer 420 @ 43.6%
-- vs no-offer 4,534 @ 63.2%.
