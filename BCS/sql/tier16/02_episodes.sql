-- Tier 16 | LAYER 02: the call layer. ONE scan of contactcenter_bdp_db.call
-- over the parameterized call month -> one row per January inbound call row
-- (id present), with the two diagnostic axes carried as COLUMNS, not filters:
--   * is_biz            : the call row is producttype BUSINESS_CARD
--   * within_effdt_cap  : the row's effdt landed before the load-date cap
--   * is_episode_std    : the row is the STANDARD episode (first inbound call
--     per account per day, ordered by initiationtimestamp, chosen AMONG rows
--     passing the standard filters: non-business-card AND within the cap).
-- The standard filters are applied BEFORE the dedup, exactly as the verified
-- kit does (filter-then-dedup); rows failing them can never be the standard
-- episode, but they stay in the output so settle-callers-style diagnostics
-- (why does another construction count more callers?) run off this same layer.
--
-- DEPENDS ON: nothing (call table only). Deliberately population-free: it
-- knows nothing about ledgers or ex-AA; 04 intersects it with 01.
-- FEEDS: 03_signals (episode contactids), 04_outcomes.
--
-- PARAMETERS: call window DATE '2025-01-01' .. DATE '2025-02-01' (exclusive);
-- effdt load-date cap '2025-01-01' .. '2025-02-02' (exclusive).
--
-- TIE-OUT ANCHORS (checked at layer 04, where the ledger join exists;
-- STOP RULE): standard episodes on accounts in the ex-AA ledger =
-- 11,262 episodes / 9,389 distinct calling accounts EXACTLY.
--
-- WITH TABLE ACCESS, uncomment:
-- CREATE TABLE <schema>.uc2_t16_02_episodes AS
WITH calls_flagged AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key,
           contactid,
           "date" AS call_dt,
           cast(date_trunc('month', "date") AS date) AS call_month,
           initiationtimestamp,
           CASE WHEN coalesce(cast(producttype AS varchar), '') = 'BUSINESS_CARD'
                THEN 1 ELSE 0 END AS is_biz,
           CASE WHEN effdt >= '2025-01-01' AND effdt < '2025-02-02'   -- PARAM: effdt cap
                THEN 1 ELSE 0 END AS within_effdt_cap
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2025-01-01' AND "date" < DATE '2025-02-01' -- PARAM: call window
      AND acctid IS NOT NULL
),
episodes_std AS (
    SELECT contactid
    FROM (
        SELECT contactid,
               row_number() OVER (PARTITION BY acct_key, call_dt
                                  ORDER BY initiationtimestamp) AS rn
        FROM calls_flagged
        WHERE acct_key IS NOT NULL AND acct_key <> ''
          AND is_biz = 0
          AND within_effdt_cap = 1
    )
    WHERE rn = 1
)
SELECT c.acct_key, c.contactid, c.call_dt, c.call_month,
       c.is_biz, c.within_effdt_cap,
       CASE WHEN e.contactid IS NOT NULL THEN 1 ELSE 0 END AS is_episode_std
FROM calls_flagged c
LEFT JOIN episodes_std e ON e.contactid = c.contactid
