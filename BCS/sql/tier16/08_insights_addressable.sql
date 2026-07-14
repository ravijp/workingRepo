-- Tier 16 | INSIGHTS 08: the addressable moment (walkthrough section 5).
-- Numbered blocks; run ONE block at a time.
--
-- DUAL-MODE RULE (every block):
--   MODE 1 (tables exist): run as-is against the uc2_t16_* tables; fill <schema>.
--   MODE 2 (no tables): delete the `outcomes` placeholder and paste the full
--     00 -> 01 -> 02 -> 03 -> 04 stitch wrapped as `outcomes` (README recipe).
--     The stitch contains the ONE transcript pass AND the day-grain call-day
--     scan: heavy, one statement per sitting, never pasted twice.
-- After any stitch or rebuild: run block 9.0 (anchor sweep) first. Miss = STOP.

-- ============================================================================
-- INSIGHT 8.1: the addressable stream size check (walkthrough 5.1/5.2 bridge).
-- is_addressable = bucket 1 on the call day AND no pre-2025 charge-off, the
-- verified b19 gate carried as a layer-04 column.
-- EXPECTED (b19/b17 cross-check, bridge round 8; STOP on miss):
--   addressable episodes = 29,114 EXACTLY; accounts ~ 28,277.
-- OUT OF LAYER SCOPE: the FULL live-stream grid (42,870 episodes across
-- buckets 1 / 2-6 / 7+ x days-since-first-DQ bands, walkthrough 5.1). The
-- day bands need the delinquency SPELL START from daily snapshots, which the
-- layers do not carry. RUN ../tier14/b17_exaa.sql VERBATIM (standalone).
-- Its anchors: bucket-1 cells 18,112 / 6,530 / 4,472 (sum 29,114, matching
-- this block to the episode); stream total 42,870; 62.2% of bucket-1 calls
-- within 10 days of falling behind.
WITH outcomes AS (
    SELECT * FROM "<schema>"."uc2_t16_04_outcomes"
)
SELECT count(*) AS addressable_episodes,
       count(DISTINCT acct_key) AS addressable_accounts
FROM outcomes
WHERE is_addressable

-- ============================================================================
-- INSIGHT 8.2: the 29,114 read by language x captured, with balance and CO
-- dollars - b19_addressable_exaa_bal face (walkthrough 5.2 + the round-9
-- grid). Balances and CO dollars deduped one row per account within each
-- (language_group, captured) cell (the m4 dedup pattern; accounts repeat
-- across cells, so never sum the balance column down). CO windows are the
-- 31-Jan layer flags.
-- EXPECTED (bridge round 9, 14 cells, EXACT episode partition; STOP on miss):
--   episodes: a - 292 / a 1 253; b - 346 / b 1 4,433; c - 1,324 / c 1 14,049;
--   d - 193 / d 1 460; e - 35 / e 1 53; f - 144 / f 1 177;
--   g - 1,097 / g 1 6,258; sum = 29,114 EXACTLY.
--   Dollar spot-checks: (a, not captured) $936,327 bal / co12 238 / $786,946;
--   (c, captured) $24,121,189 / 625 / $2,582,326.
--   Capture-rate view (walkthrough 5.2): the same cells regrouped give
--   deceased 46.4% / promise 92.8% / pay-talk 91.4% / plan 70.4% /
--   hardship 60.2% / dispute 55.1% / none 85.1%; total 25,683 / 3,431 (88.2%).
WITH outcomes AS (
    SELECT * FROM "<schema>"."uc2_t16_04_outcomes"
),
adr AS (
    SELECT * FROM outcomes WHERE is_addressable
),
acct_grp AS (
    SELECT language_group, captured, acct_key,
           max(jan_eom_bal) AS jan_eom_bal,
           bool_or(co_8m)  AS co_8m,
           bool_or(co_10m) AS co_10m,
           bool_or(co_12m) AS co_12m,
           max(co_amt) AS co_amt
    FROM adr
    GROUP BY 1, 2, 3
),
grp_bal AS (
    SELECT language_group, captured,
           round(sum(jan_eom_bal), 0) AS jan_eom_balance,
           count_if(co_8m)  AS co_8m,
           count_if(co_10m) AS co_10m,
           count_if(co_12m) AS co_12m,
           round(sum(CASE WHEN co_8m  THEN co_amt END), 0) AS co8_amt,
           round(sum(CASE WHEN co_10m THEN co_amt END), 0) AS co10_amt,
           round(sum(CASE WHEN co_12m THEN co_amt END), 0) AS co12_amt
    FROM acct_grp
    GROUP BY 1, 2
)
SELECT e.language_group, e.captured,
       count(*) AS episodes,
       count(DISTINCT e.acct_key) AS accounts,
       g.jan_eom_balance, g.co_8m, g.co_10m, g.co_12m,
       g.co8_amt, g.co10_amt, g.co12_amt
FROM adr e
JOIN grp_bal g
  ON g.language_group = e.language_group AND g.captured = e.captured
GROUP BY 1, 2, 5, 6, 7, 8, 9, 10, 11
ORDER BY 1, 2

-- ============================================================================
-- INSIGHT 8.3: the walk-down to the addressable number (walkthrough 5.3).
-- Intent = the promise + payment-talk + plan groups (b/c/d), the recorded
-- addressable-partition definition. Dollar cells dedup one row per account
-- within the intent-not-captured set.
-- EXPECTED (derived from the 8.2 cells, re-added in verification; STOP on miss):
--   bucket-1 episodes 29,114 (~28,277 accounts);
--   intent present 20,805; captured 18,942;
--   intent NOT captured = 1,863 EXACTLY, from ~1,799 accounts,
--   on ~$6.95M of January balance with ~$3.88M charged off within 12 months
--   (31-Jan windows; = the round-9 intent rows 346+1,324+193 eps,
--   $1,280,420 + $4,390,537 + $1,278,049 bal, $658,477 + $2,424,431 +
--   $792,394 co12$, minus nothing - the three not-captured intent cells);
--   deceased routed: 545 episodes, 292 of them not captured.
WITH outcomes AS (
    SELECT * FROM "<schema>"."uc2_t16_04_outcomes"
),
adr AS (
    SELECT *,
           (language_group IN ('b. future-dated promise',
                               'c. payment talk, no promise',
                               'd. plan or settlement talk')) AS has_intent
    FROM outcomes WHERE is_addressable
),
leak_acct AS (   -- one row per account inside the intent-not-captured cell
    SELECT acct_key,
           max(jan_eom_bal) AS jan_eom_bal,
           bool_or(co_12m) AS co_12m,
           max(co_amt) AS co_amt
    FROM adr
    WHERE has_intent AND captured = 0
    GROUP BY 1
)
SELECT '1. bucket 1 on the call day' AS step,
       count(*) AS episodes, count(DISTINCT acct_key) AS accounts,
       NULL AS jan_eom_balance, NULL AS co12_amt
FROM adr
UNION ALL
SELECT '2. payment intent present', count(*), count(DISTINCT acct_key), NULL, NULL
FROM adr WHERE has_intent
UNION ALL
SELECT '3. of 2: captured within 30 days', count(*), count(DISTINCT acct_key), NULL, NULL
FROM adr WHERE has_intent AND captured = 1
UNION ALL
SELECT '4. of 2: NOT captured (the addressable moment)',
       count(*), count(DISTINCT acct_key),
       (SELECT round(sum(jan_eom_bal), 0) FROM leak_acct),
       (SELECT round(sum(CASE WHEN co_12m THEN co_amt END), 0) FROM leak_acct)
FROM adr WHERE has_intent AND captured = 0
UNION ALL
SELECT '5. deceased or estate, routed out', count(*), count(DISTINCT acct_key), NULL, NULL
FROM adr WHERE language_group = 'a. deceased or estate'
UNION ALL
SELECT '6. of 5: not captured', count(*), count(DISTINCT acct_key), NULL, NULL
FROM adr WHERE language_group = 'a. deceased or estate' AND captured = 0
ORDER BY 1
