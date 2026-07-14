-- Tier 16 | INSIGHTS 06: motion, runway, and entrants (walkthrough section 2).
-- Numbered blocks; run ONE block at a time.
--
-- DUAL-MODE RULE (every block):
--   MODE 1 (tables exist): run as-is against the uc2_t16_* tables; fill <schema>.
--   MODE 2 (no tables): delete the placeholder CTE(s), paste the layer chains
--     per the README stitch recipe (00 -> 01 for `populations`; the FULL
--     00 -> 01 -> 02 -> 03 -> 04 stitch, wrapped as `outcomes`, wherever an
--     `outcomes` placeholder appears - that stitch includes the ONE transcript
--     pass, so run it in a sitting where the transcript scan is affordable,
--     and never paste it twice into one statement).
-- After any stitch or rebuild: run block 9.0 (anchor sweep) first. Miss = STOP.

-- ============================================================================
-- INSIGHT 6.1: the full February transitions grid with balance and CO dollars,
-- b14_exaa_bal face (walkthrough 2.2 / 2.3 / 4.2 and Appendix 2.3). Grain: Feb
-- position x caller class x runway band over the ex-AA ledger. The caller
-- class comes from layer 04 rolled up to account level (caller_class is
-- constant per account; max_by is just a picker); non-callers exist only on
-- the 01 side, hence the LEFT JOIN + coalesce. CO windows here are the
-- 31-Jan-anchored layer flags (the round-9 re-baseline). The Feb charge-off
-- dollar (co_amt_feb) populates only on the 'e. charged off in Feb' rows,
-- exactly as the verified b14_exaa_bal grid does.
-- EXPECTED (bridge round 9, b14_exaa_bal, 66 rows; STOP on miss):
--   accounts across ALL rows = 189,146 EXACTLY;
--   jan_eom_balance across all rows = $457,943,987 (tolerance ~$5);
--   co12_amt across all rows ~ $156.6M (the section 2.2 total, derived);
--   row values reproduce the recorded 66-row grid cell for cell.
-- Label note: layer labels read 'a. Feb EOM bucket 0 (cured)' etc.; the
-- recorded grid says 'a. cured' - same partition, cosmetic difference.
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
SELECT p.feb_position_b14,
       coalesce(k.caller_class, 'a. non-caller') AS caller_class,
       p.runway_band,
       count(*) AS accounts,
       round(sum(p.eom_bal), 0) AS jan_eom_balance,
       round(sum(CASE WHEN p.feb_position_b14 LIKE 'e.%' THEN p.co_amt END), 0) AS co_amt_feb,
       count_if(p.co_8m)  AS co_8m,
       count_if(p.co_10m) AS co_10m,
       count_if(p.co_12m) AS co_12m,
       round(sum(CASE WHEN p.co_8m  THEN p.co_amt END), 0) AS co8_amt,
       round(sum(CASE WHEN p.co_10m THEN p.co_amt END), 0) AS co10_amt,
       round(sum(CASE WHEN p.co_12m THEN p.co_amt END), 0) AS co12_amt
FROM populations p
LEFT JOIN callers k ON k.acct_key = p.acct_key
WHERE p.in_ledger_exaa
GROUP BY 1, 2, 3
ORDER BY 1, 2, 3

-- ============================================================================
-- INSIGHT 6.2: February outcome rollup with balance and CO12 dollars
-- (walkthrough 2.2 main table). Same population as 6.1, grouped by Feb
-- position only. No caller join needed.
-- EXPECTED (derived row sums of the round-8/9 b14 grids; STOP on miss):
--   cured 105,215 / $203.0M bal / ~$16.9M co12$;
--   stayed 15,054 / $71.1M / ~$14.7M;
--   rolled 68,093 / $180.9M / ~$122.3M;
--   deeper 111 / $0.6M / ~$0.4M;
--   charged off in Feb 673 / $2.4M / ~$2.3M (its co_amt_feb ~ $2.42M);
--   totals 189,146 / $457,943,987 / ~$156.6M.
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
)
SELECT feb_position_b14,
       count(*) AS accounts,
       round(100.0 * count(*) / sum(count(*)) OVER (), 1) AS share_pct,
       round(sum(eom_bal), 0) AS jan_eom_balance,
       round(sum(CASE WHEN feb_position_b14 LIKE 'e.%' THEN co_amt END), 0) AS co_amt_feb,
       count_if(co_12m) AS co_12m_accounts,
       round(sum(CASE WHEN co_12m THEN co_amt END), 0) AS co12_amt
FROM populations
WHERE in_ledger_exaa
GROUP BY 1
ORDER BY 1

-- ============================================================================
-- INSIGHT 6.3: January entrants by entry-day band, b13 face (walkthrough 2.1).
-- Population: the WHOLE cleaned book, BEFORE the AA filter (b13 was never
-- re-run ex-AA; section 1 records that the cut does not change the pattern).
-- Entrant = month-max bucket exactly 1 AND not delinquent in December
-- (b13's max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0), cleaned.
-- Entry day = the day-of-month of the first delinquent January snapshot
-- (layer 01's first_dq_dt, 'YYYYMMDD' string; day = substr(_, 7, 2)).
-- EXPECTED (b13, bridge round 6; STOP on miss):
--   day 1-10: 217,607 entrants / 174,380 cured (80.1%);
--   day 11-20: 162,076 / 110,217 (68.0%);
--   day 21-31: 112,391 / 56,364 (50.1%);
--   total 492,074 / 340,961 cured / 151,113 still behind (69.3%).
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
)
SELECT CASE
         WHEN cast(substr(first_dq_dt, 7, 2) AS integer) <= 10 THEN 'a. day 1-10 (21+ days runway)'
         WHEN cast(substr(first_dq_dt, 7, 2) AS integer) <= 20 THEN 'b. day 11-20 (11-20 days runway)'
         ELSE 'c. day 21-31 (<= 10 days runway)'
       END AS entry_band,
       count(*) AS entrants,
       count_if(eom_bucket = 0) AS cured_by_eom,
       count_if(eom_bucket >= 1) AS still_behind_eom,
       round(100.0 * count_if(eom_bucket = 0) / count(*), 1) AS pct_cured
FROM populations
WHERE max_bucket = 1
  AND coalesce(prev_max_bucket, 0) = 0
  AND cleaned
GROUP BY 1
ORDER BY 1

-- ============================================================================
-- INSIGHT 6.4: February outcome by entry timing on the ex-AA ledger
-- (walkthrough 2.3), plus the entrant/carried-in split (167,951 / 21,195).
-- Groups = layer 01's runway_band (the b14 runway CASE, verbatim).
-- EXPECTED (b14_exaa row sums, derived; STOP on miss):
--   day 1-10:   46,827 / $101.7M / cured 49.5% / rolled 45.2%;
--   day 11-20:  56,127 / $124.7M / 56.2% / 37.6%;
--   day 21-31:  64,997 / $146.0M / 65.2% / 27.9%;
--   carried-in: 21,195 / $85.6M  / 38.4% / 36.1%;
--   totals 189,146 and $457,943,987 EXACTLY; the three entrant bands sum
--   to 167,951 (88.8%), carried-in 21,195 (11.2%).
-- 'Cured' = Feb EOM bucket 0; 'rolled' = Feb EOM bucket 2 (the b14 rule).
WITH populations AS (
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
)
SELECT runway_band,
       count(*) AS accounts,
       round(sum(eom_bal), 0) AS jan_eom_balance,
       round(100.0 * count_if(feb_position_b14 LIKE 'a.%') / count(*), 1) AS pct_cured_feb,
       round(100.0 * count_if(feb_position_b14 LIKE 'c.%') / count(*), 1) AS pct_rolled_feb
FROM populations
WHERE in_ledger_exaa
GROUP BY 1
ORDER BY 1
