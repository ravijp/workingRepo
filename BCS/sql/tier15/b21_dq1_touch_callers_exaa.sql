-- Tier 15 | NEW QUERY, drafted 2026-07-13.
-- Purpose (Ravi's ask): count January inbound calls for accounts that
-- TOUCHED bucket 1 at ANY point in January 2025, not just those still
-- bucket-1 at month-end. The current b14_exaa ledger misses an account that
-- was DQ1 mid-month and DQ2+ (or charged off) by 31 Jan. Reuses b14_exaa's
-- snap/monthly pattern over January, keeping max_bucket, eom_bucket, AND a
-- new touched_b1 test (>= 1 January snapshot with bucket = 1, via
-- first_b1_dt = min(eff_dt) WHERE bucket = 1, IS NOT NULL). Population:
-- touched_b1 accounts, cleaned (co_dt IS NULL OR co_dt >= 2025-01-01),
-- ex-AA (NULL-safe eom_cpc predicate, same 23-code set). Joins January
-- inbound episodes (b14_exaa's inb/episodes CTEs verbatim), semi-joined to
-- this population for cost.
-- PRE-REGISTERED TIE-OUTS (STOP RULE):
--   - Row b ('bucket 1 at 31 Jan') MUST reproduce the b14_exaa ledger
--     numbers EXACTLY: 189,146 accounts / 9,389 callers / 11,262 episodes.
--     This is the strongest cross-check; any mismatch = STOP, route to the
--     keeper before trusting rows a/c/d.
--   - Row a ('current at 31 Jan, cured in month') pct_accounts_calling
--     should sit near b7_exaa's class-a caller rate (14.7%).
--   - Row c is the NEW information: the DQ1-to-deeper-within-January
--     callers the current bucket-1-at-EOM ledger does not count.
-- RESOURCE DISCIPLINE: one January account scan (snap/monthly, as in
-- b14_exaa); episodes CTE built first and semi-joined against the
-- touched_b1 population to bound cost, matching the tier-14 fixes.
WITH snap AS (
    SELECT extnl_acct_id,
           substr(eff_dt, 1, 6) AS ym,
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
           try_cast(acct_bal_amt AS double) AS bal,
           try_cast(chrgoff_dt AS date) AS co_dt,
           clnt_prdct_cd
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           max_by(bal, eff_dt) AS eom_bal,
           min(co_dt) AS co_dt,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc,
           min(CASE WHEN bucket = 1 THEN eff_dt END) AS first_b1_dt
    FROM snap GROUP BY 1, 2
),
population AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           eom_bucket,
           eom_bal,
           co_dt
    FROM monthly
    WHERE ym = '202501'
      AND first_b1_dt IS NOT NULL
      AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
      AND (eom_cpc IS NULL OR trim(eom_cpc) = ''
           OR eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                               'AA3','AC3','AM3','AA4','AC4','AM4',
                               'BGC','BGM','CGM','GMR',
                               'FBS','IBS','U1C','U2C','U3C'))
),
classed AS (
    SELECT acct_key, eom_bal,
           CASE
             WHEN co_dt >= DATE '2025-01-01' AND co_dt < DATE '2025-02-01'
               THEN 'd. charged off in January'
             WHEN eom_bucket = 0
               THEN 'a. current at 31 Jan (cured in month)'
             WHEN eom_bucket = 1
               THEN 'b. bucket 1 at 31 Jan'
             WHEN eom_bucket >= 2
               THEN 'c. bucket 2+ at 31 Jan (rolled past DQ1 within January)'
           END AS class
    FROM population
),
inb AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key, contactid,
           "date" AS call_dt, initiationtimestamp
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2025-01-01' AND "date" < DATE '2025-02-01'
      AND effdt >= '2025-01-01' AND effdt < '2025-02-02'
      AND coalesce(cast(producttype AS varchar), '') <> 'BUSINESS_CARD'
),
episodes AS (
    SELECT acct_key, call_dt
    FROM (
        SELECT i.acct_key, i.call_dt,
               row_number() OVER (PARTITION BY i.acct_key, i.call_dt
                                  ORDER BY i.initiationtimestamp) AS rn
        FROM inb i
        JOIN classed c ON c.acct_key = i.acct_key
        WHERE i.acct_key IS NOT NULL AND i.acct_key <> ''
    )
    WHERE rn = 1
),
per_acct AS (
    SELECT acct_key, count(*) AS n_episodes
    FROM episodes GROUP BY 1
)
SELECT c.class,
       count(*) AS accounts,
       count(e.acct_key) AS callers,
       coalesce(sum(e.n_episodes), 0) AS episodes,
       round(100.0 * count(e.acct_key) / count(*), 1) AS pct_accounts_calling,
       round(sum(c.eom_bal), 0) AS jan_eom_balance
FROM classed c
LEFT JOIN per_acct e ON c.acct_key = e.acct_key
GROUP BY 1
ORDER BY 1
