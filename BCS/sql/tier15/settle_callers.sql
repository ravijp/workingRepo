-- Tier 15 | CALLER-GAP DIAGNOSTIC LADDER (rewritten 2026-07-13 after the
-- export code behind zenon.aws_call_accts_jan_mar_25 surfaced).
-- SETTLED: that export filters call_month='M1' before saving, so the SAS
-- caller flag is JANUARY-ONLY (the table name is misleading). The earlier
-- Jan-vs-Jan-Mar window explanation of 11,154 vs 9,389 is dead; the
-- remaining candidate causes are (1) the export includes business-card
-- call rows (ours excludes), (2) the export has no effdt load-date cap
-- (ours drops rows arriving after 2025-02-02), (3) the cycle-code vs
-- amount-bucket population boundary between his slice and our ledger.
-- This ladder sizes (1) and (2) at book level; the residual is (3).
--
-- Run in order. All three are single scans of contactcenter_bdp_db.call,
-- January "date" window. NOTE ON COST: statement C intentionally has NO
-- effdt filter (replicating the export), so it scans without partition
-- pruning, as the original export did.
--
-- PRE-REGISTERED EXPECTATIONS:
--   C  (replicates the export logic)  ~= 831,261 distinct inbound accounts
--      (the export pivot's own INBOUND count; small drift only, from
--      empty-string id handling). Far off = our call-table read and the
--      export disagree at base level: STOP, route to the keeper.
--   B2 (C + the effdt cap)            <= C; C - B2 = late-arriving rows.
--   A  (B2 + business-card exclusion) <= B2; B2 - A = business-card share.
--   These are BOOK-LEVEL counts (hundreds of thousands), not comparable
--   to 9,389 / 11,154 directly; the two deltas transfer proportionally to
--   the ledger-level gap, and what they do not explain is the population
--   boundary (3).

-- ============================================================
-- (C) Replicate the SAS-side export logic exactly: January calls,
-- id present, no producttype exclusion, no effdt cap.
-- ============================================================
SELECT count(DISTINCT trim(cast(acctid AS varchar))) AS settle_c_accounts
FROM "contactcenter_bdp_db"."call"
WHERE initiationmethod = 'INBOUND'
  AND "date" >= DATE '2025-01-01' AND "date" <= DATE '2025-01-31'
  AND acctid IS NOT NULL;

-- ============================================================
-- (B2) Same as C, plus our effdt load-date cap (isolates the
-- late-arriving-row effect).
-- ============================================================
SELECT count(DISTINCT trim(cast(acctid AS varchar))) AS settle_b2_accounts
FROM "contactcenter_bdp_db"."call"
WHERE initiationmethod = 'INBOUND'
  AND "date" >= DATE '2025-01-01' AND "date" <= DATE '2025-01-31'
  AND acctid IS NOT NULL
  AND effdt >= '2025-01-01' AND effdt < '2025-02-02';

-- ============================================================
-- (A) Same as B2, plus our business-card exclusion (isolates the
-- business-card share; this is our standard inb filter set).
-- ============================================================
SELECT count(DISTINCT trim(cast(acctid AS varchar))) AS settle_a_accounts
FROM "contactcenter_bdp_db"."call"
WHERE initiationmethod = 'INBOUND'
  AND "date" >= DATE '2025-01-01' AND "date" <= DATE '2025-01-31'
  AND acctid IS NOT NULL
  AND effdt >= '2025-01-01' AND effdt < '2025-02-02'
  AND coalesce(cast(producttype AS varchar), '') <> 'BUSINESS_CARD';
