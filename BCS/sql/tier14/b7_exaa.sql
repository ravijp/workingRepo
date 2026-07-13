-- Tier 14 | EX-AA VARIANT of b7_call_overlay.
-- Exclusion: clnt_prdct_cd at the account's January EOM snapshot (eom_cpc,
-- max_by(clnt_prdct_cd, eff_dt) over the January row), matched against the
-- SAS AA+GM+Bronco code set (23 codes). Applied on the account population
-- CTE (classed), so every downstream bridge_class (a/b/c/d) is ex-AA.
-- NULL-safe form: NULL or blank cpc is kept as "others".
-- Expected tie-out: run AFTER b14_exaa ties out 189,146. b7_exaa's classes
-- a+b+c+d should sum to the ex-AA delinquent-in-month population (a bound,
-- not an exact tie to 189,146 - b7's population is max_bucket >= 1 in
-- January, broader than b14's month-END bucket-1 ledger).
-- Everything else is unchanged from b7_call_overlay.sql: the bucket ladder,
-- the bridge-class definitions, the call/episode logic.
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
           try_cast(chrgoff_dt AS date) AS co_dt,
           clnt_prdct_cd
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20241201' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           min(co_dt) AS co_dt,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc
    FROM snap GROUP BY 1, 2
),
base AS (
    SELECT j.extnl_acct_id, j.max_bucket, j.eom_bucket, j.co_dt, j.eom_cpc,
           p.max_bucket AS prev_max_bucket
    FROM (SELECT * FROM monthly WHERE ym = '202501') j
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202412') p
      ON j.extnl_acct_id = p.extnl_acct_id
),
classed AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           CASE
             WHEN max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0
                  AND eom_bucket = 0
               THEN 'a. month-MAX B1 entrant, cured by EOM (invisible to ASP)'
             WHEN max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0
                  AND eom_bucket >= 1
               THEN 'b. month-MAX B1 entrant, still DQ1 at EOM'
             WHEN eom_bucket = 1
                  AND NOT (max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0)
               THEN 'c. EOM bucket 1 stock, not a month-max-B1 entrant'
             ELSE 'd. other delinquent in month (month-MAX >= 2, EOM <> 1)'
           END AS bridge_class
    FROM base
    WHERE max_bucket >= 1
      AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
      AND (eom_cpc IS NULL OR trim(eom_cpc) = ''
           OR eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                               'AA3','AC3','AM3','AA4','AC4','AM4',
                               'BGC','BGM','CGM','GMR',
                               'FBS','IBS','U1C','U2C','U3C'))
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
        SELECT acct_key, call_dt,
               row_number() OVER (PARTITION BY acct_key, call_dt
                                  ORDER BY initiationtimestamp) AS rn
        FROM inb
        WHERE acct_key IS NOT NULL AND acct_key <> ''
    )
    WHERE rn = 1
),
per_acct AS (
    SELECT acct_key, count(*) AS n_episodes
    FROM episodes GROUP BY 1
)
SELECT c.bridge_class,
       count(*) AS class_accounts,
       count(e.acct_key) AS jan_inbound_callers,
       coalesce(sum(e.n_episodes), 0) AS jan_inbound_episodes,
       round(100.0 * count(e.acct_key) / count(*), 1) AS pct_accounts_calling
FROM classed c
LEFT JOIN per_acct e ON c.acct_key = e.acct_key
GROUP BY 1
ORDER BY 1
