-- Tier 16 | LAYER 03: transcript signals. ONE pass over
-- contactcenter_bdp_db.transcript per contactid -> one row per contactid that
-- (a) is a standard episode on an ex-AA account and (b) has at least one
-- customer utterance matching the coarse pre-filter. Contactids with NO row
-- here have no matching language: downstream must COALESCE missing flags to 0
-- and missing language_group to 'g. no payment-related language'.
--
-- RESPECTS THE ROUND-9 m4 LESSONS (the OOM fixes), all three:
--   (1) the transcript table is referenced EXACTLY ONCE in the whole layer
--       chain (this file; nothing downstream touches it again);
--   (2) participantid = 'CUSTOMER' sits in the WHERE, plus ONE coarse
--       union-of-all-lexicons regexp_like pre-filter, so the precise regexes
--       run only on customer utterances already holding a relevant token;
--   (3) boolean presence (max(CASE ...)), not raw counts, per contactid.
-- The coarse union below is the concatenation of ALL per-flag alternations
-- (six v2 lexicons + the execution lexicon), so gating on it cannot drop a
-- would-be match for any flag.
--
-- DEPENDS ON: 01_populations (ex-AA prune) and 02_episodes (episode
-- contactids). The prune matches the verified kit: transcripts are read only
-- for standard episodes on ex-AA accounts (the b19 driver set, which contains
-- the ledger-caller set m4 used).
-- FEEDS: 04_outcomes.
--
-- PARAMETERS: transcript effdt window '2025-01-01' .. '2025-02-02' (exclusive),
-- matching the call layer's cap.
--
-- TIE-OUT ANCHORS (checked at layer 04; STOP RULE): the language partition
-- over ledger episodes sums to 11,262 EXACTLY (m4: a 498 / b 1,374 / c 5,164 /
-- d 442 / e 79 / f 274 / g 3,431); deceased episodes never appear inside
-- intent groups (priority CASE by construction).
--
-- FLAGS (regexes verbatim from the verified v2 lexicon):
--   deceased_f, promise_f, pay_f, plan_f, hard_f, dispute_f, exec_f
-- plus language_group = the v2 priority CASE (deceased first, then promise,
-- payment talk, plan, hardship, dispute; 'g' handled downstream via coalesce).
--
-- WITH TABLE ACCESS, uncomment:
-- CREATE TABLE <schema>.uc2_t16_03_signals AS
WITH populations AS (
    -- TABLE MODE (later): keep this SELECT, pointing at the saved 01 table.
    -- STITCH MODE (today): delete this placeholder CTE and paste layer 00's
    -- and layer 01's CTE chains here (README recipe).
    SELECT * FROM "<schema>"."uc2_t16_01_populations"
),
calls AS (
    -- TABLE MODE (later): point at the saved 02 table.
    -- STITCH MODE (today): paste layer 02's CTE chain, wrapping its final
    -- SELECT as `calls AS (SELECT ... FROM calls_flagged c LEFT JOIN episodes_std ...)`.
    -- Same placeholder name as layer 04, so a full stitch pastes 02 once.
    SELECT * FROM "<schema>"."uc2_t16_02_episodes"
),
drivers AS (
    -- NULL-safe ex-AA intersection, matching the verified b19 LEFT JOIN:
    -- a calling account with NO anchor-month account row is KEPT (its cpc is
    -- unknown, and the NULL-safe rule keeps NULL/blank cpc as "others").
    SELECT DISTINCT e.contactid
    FROM calls e
    LEFT JOIN populations p ON p.acct_key = e.acct_key
    WHERE e.is_episode_std = 1
      AND (p.acct_key IS NULL OR p.is_exaa)
),
tx AS (
    SELECT t.contactid,
           max(CASE WHEN regexp_like(lower(t.content),
                     'passed away|death certificate|executor|deceased|calling on behalf') THEN 1 ELSE 0 END) AS deceased_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'pay|paid|payment|settle|payment plan|arrangement|work something out') THEN 1 ELSE 0 END) AS pay_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'settle|payment plan|arrangement|work something out') THEN 1 ELSE 0 END) AS plan_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford') THEN 1 ELSE 0 END) AS hard_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'dispute|not my charge|didn.t authorize|did not authorize|unauthorized|fraud|identity theft') THEN 1 ELSE 0 END) AS dispute_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'i.ll pay|i will pay|going to pay|gonna pay|pay (on|by|this|next)|when i get paid|payday|after my paycheck') THEN 1 ELSE 0 END) AS promise_f,
           max(CASE WHEN regexp_like(lower(t.content),
                     'bank routing|routing number|check number|checkbook|a check for|that check|on the check') THEN 1 ELSE 0 END) AS exec_f
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN drivers d ON t.contactid = d.contactid
    WHERE t.effdt >= '2025-01-01' AND t.effdt < '2025-02-02'   -- PARAM: effdt window
      AND t.content IS NOT NULL
      AND t.participantid = 'CUSTOMER'
      AND regexp_like(lower(t.content),
            'pay|paid|payment|settle|arrangement|work something out|passed away|death certificate|executor|deceased|calling on behalf|hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford|dispute|not my charge|didn.t authorize|did not authorize|unauthorized|fraud|identity theft|i.ll pay|i will pay|going to pay|gonna pay|when i get paid|payday|after my paycheck|bank routing|routing number|check number|checkbook|a check for|that check|on the check')
    GROUP BY 1
)
SELECT contactid,
       deceased_f, promise_f, pay_f, plan_f, hard_f, dispute_f, exec_f,
       -- the v2 priority CASE, verbatim ordering (deceased always wins)
       CASE
         WHEN deceased_f > 0 THEN 'a. deceased or estate'
         WHEN promise_f  > 0 THEN 'b. future-dated promise'
         WHEN pay_f > 0 AND plan_f = 0 THEN 'c. payment talk, no promise'
         WHEN plan_f     > 0 THEN 'd. plan or settlement talk'
         WHEN hard_f     > 0 THEN 'e. hardship talk'
         WHEN dispute_f  > 0 THEN 'f. dispute or fraud talk'
         ELSE 'g. no payment-related language'
       END AS language_group
FROM tx
