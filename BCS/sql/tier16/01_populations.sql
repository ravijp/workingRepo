-- Tier 16 | LAYER 01: populations. Builds ON layer 00. One row per account
-- that has a January (anchor-month) row, carrying every population flag and
-- classification the analyses use, so downstream queries filter on flags
-- instead of re-deriving logic:
--   * cleaned          : the cleanup rule (Jan mth_co_dt IS NULL OR >= 2025-01-01)
--   * is_exaa          : the NULL-safe 23-code ex-AA test (NULL/blank kept)
--   * cpc_class        : the full client CPC mapping, AA tested first
--   * in_ledger_all    : cleaned January EOM bucket-1 ledger, all products
--   * in_ledger_exaa   : the ex-AA ledger (the 189,146)
--   * touched_b1 / touched_b1_class : the touched-bucket-1 universe and its
--     EOM-outcome class (current / B1 / B2+ / charged off in January)
--   * runway_band, feb_position_b14, feb_pos, mar_pos : motion labels
--   * co_dt_future, co_amt, co_8m/co_10m/co_12m : the 31-Jan-anchored
--     forward charge-off windows and the gross CO dollar
--
-- DEPENDS ON: 00_acct_monthly (the `acct_monthly` reference below), plus ONE
-- extra base scan: future_co, the verified forward charge-off read (needs a
-- 12-month horizon past the 00 window; pruned by chrgoff_dt IS NOT NULL and
-- semi-joined to January delinquency-touched accounts).
-- FEEDS: 03_signals (ex-AA prune), 04_outcomes (all account attributes).
--
-- PARAMETERS: anchor month '202501' (and its neighbors '202412', '202502',
-- '202503'); cleanup date DATE '2025-01-01'; CO windows anchored 31 Jan 2025:
--   CO8  [2025-01-31, 2025-09-30)   CO10 [.., 2025-11-30)   CO12 [.., 2026-01-31)
--
-- TIE-OUT ANCHORS (STOP RULE; any miss after restructuring = STOP):
--   * count_if(in_ledger_all)  = 204,323 EXACTLY
--       (AA row of cpc_class = 15,177 / ~$73,744,823)
--   * count_if(in_ledger_exaa) = 189,146 EXACTLY;
--       sum eom_bal over them  = $457,943,987 (tolerance ~$5)
--   * count_if(touched_b1)     = 724,848 EXACTLY
--       (class rows: a 464,023 / b 186,714 / c 69,513 / d 4,598;
--        b + the 2,432 eom_bucket=1 Jan-CO accounts in d = 189,146)
--
-- NOTE on runway_band: computed exactly as the verified ledger CASE for any
-- account delinquent in the month; NULL for accounts never delinquent in
-- January and not carried-in (the verified query never evaluated those rows).
--
-- WITH TABLE ACCESS, uncomment:
-- CREATE TABLE <schema>.uc2_t16_01_populations AS
WITH acct_monthly AS (
    -- TABLE MODE (later): keep this SELECT, pointing at the saved 00 table.
    -- STITCH MODE (today): DELETE this whole placeholder CTE and paste layer
    -- 00's CTEs (snap, acct_monthly) here in its place; the references below
    -- then read 00's own acct_monthly directly. See README recipe.
    SELECT * FROM "<schema>"."uc2_t16_00_acct_monthly"
),
jan AS (SELECT * FROM acct_monthly WHERE ym = '202501'),   -- PARAM: anchor month
prv AS (SELECT * FROM acct_monthly WHERE ym = '202412'),
feb AS (SELECT * FROM acct_monthly WHERE ym = '202502'),
mar AS (SELECT * FROM acct_monthly WHERE ym = '202503'),
future_co AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           min(try_cast(chrgoff_dt AS date)) AS co_dt_future,
           -- co_amt = the charge-off amount on the row with the earliest
           -- charge-off date. The CTE's WHERE already restricts to
           -- chrgoff_dt IS NOT NULL, so no FILTER clause is needed.
           min_by(try_cast(chrgoff_amt AS double), try_cast(chrgoff_dt AS date)) AS co_amt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20260101'    -- PARAM: forward CO scan
      AND chrgoff_dt IS NOT NULL
      AND trim(cast(extnl_acct_id AS varchar)) IN
          (SELECT acct_key FROM jan WHERE max_bucket >= 1)
    GROUP BY extnl_acct_id
),
pop_base AS (
    SELECT j.acct_key,
           j.max_bucket, j.eom_bucket, j.eom_bal, j.mth_co_dt AS jan_co_dt,
           j.first_dq_dt, j.first_b1_dt, j.eom_cpc, j.eom_cr_lmt_origl_amt,
           p.max_bucket AS prev_max_bucket,
           p.eom_bucket AS prev_eom_bucket,
           -- the cleanup rule, verbatim
           (j.mth_co_dt IS NULL OR j.mth_co_dt >= DATE '2025-01-01') AS cleaned,
           -- the NULL-safe ex-AA test, verbatim (NULL/blank kept as "others")
           (j.eom_cpc IS NULL OR trim(j.eom_cpc) = ''
            OR j.eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                                  'AA3','AC3','AM3','AA4','AC4','AM4',
                                  'BGC','BGM','CGM','GMR',
                                  'FBS','IBS','U1C','U2C','U3C')) AS is_exaa,
           -- the full client CPC mapping, AA tested first, verbatim
           CASE
             WHEN j.eom_cpc IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                                 'AA3','AC3','AM3','AA4','AC4','AM4')     THEN 'AA'
             WHEN j.eom_cpc IN ('BGC','BGM','CGM','GMR')                 THEN 'GM'
             WHEN j.eom_cpc IN ('FBS','IBS','U1C','U2C','U3C')           THEN 'Bronco'
             WHEN j.eom_cpc IN ('BHA','BJT','BJC','BFR','BWY','BBB')     THEN 'Biz'
             WHEN j.eom_cpc IN ('GAP','GP2','ONV','ON2','BRP','BR2','ATH','AT2',
                                 'GPC','G2C','ONC','O2C','BRC','B2C','ATC','A2C')
                                                                          THEN 'CoBrand'
             WHEN j.eom_cpc IN ('8GP','8ON','8BR','8AT','9GP','9ON','9BR','9AT')
                                                                          THEN 'PLCC'
             ELSE 'OTHER'
           END AS cpc_class,
           -- runway band, verbatim ordering; NULL guard added only for rows
           -- the verified query never evaluated (not delinquent in month)
           CASE
             WHEN coalesce(p.eom_bucket, 0) >= 1
               THEN 'd. carried-in (past due at Dec-31 EOM)'
             WHEN j.first_dq_dt IS NULL THEN NULL
             WHEN cast(substr(j.first_dq_dt, 7, 2) AS integer) <= 10
               THEN 'a. runway >= 21 days (entry day 1-10)'
             WHEN cast(substr(j.first_dq_dt, 7, 2) AS integer) <= 20
               THEN 'b. runway 11-20 days (entry day 11-20)'
             ELSE 'c. runway <= 10 days (entry day 21-31)'
           END AS runway_band,
           -- Feb position, b14 labels, verbatim
           CASE
             WHEN f.mth_co_dt >= DATE '2025-02-01' AND f.mth_co_dt < DATE '2025-03-01'
               THEN 'e. charged off in Feb'
             WHEN f.acct_key IS NULL THEN 'f. no Feb row'
             WHEN f.eom_bucket = 0 THEN 'a. Feb EOM bucket 0 (cured)'
             WHEN f.eom_bucket = 1 THEN 'b. Feb EOM bucket 1 (stayed)'
             WHEN f.eom_bucket = 2 THEN 'c. Feb EOM bucket 2 (rolled)'
             ELSE 'd. Feb EOM bucket 3+ (rolled deeper)'
           END AS feb_position_b14,
           -- Feb / Mar positions, b15 compact codes, verbatim
           CASE
             WHEN f.mth_co_dt >= DATE '2025-02-01' AND f.mth_co_dt < DATE '2025-03-01' THEN 'co'
             WHEN f.acct_key IS NULL THEN 'gone'
             ELSE cast(f.eom_bucket AS varchar)
           END AS feb_pos,
           CASE
             WHEN m.mth_co_dt >= DATE '2025-02-01' AND m.mth_co_dt < DATE '2025-04-01' THEN 'co'
             WHEN m.acct_key IS NULL THEN 'gone'
             ELSE cast(m.eom_bucket AS varchar)
           END AS mar_pos,
           fc.co_dt_future,
           fc.co_amt,
           (fc.co_dt_future >= DATE '2025-01-31' AND fc.co_dt_future < DATE '2025-09-30') AS co_8m,
           (fc.co_dt_future >= DATE '2025-01-31' AND fc.co_dt_future < DATE '2025-11-30') AS co_10m,
           (fc.co_dt_future >= DATE '2025-01-31' AND fc.co_dt_future < DATE '2026-01-31') AS co_12m
    FROM jan j
    LEFT JOIN prv p ON p.acct_key = j.acct_key
    LEFT JOIN feb f ON f.acct_key = j.acct_key
    LEFT JOIN mar m ON m.acct_key = j.acct_key
    LEFT JOIN future_co fc ON fc.acct_key = j.acct_key
)
SELECT *,
       (eom_bucket = 1 AND cleaned)            AS in_ledger_all,
       (eom_bucket = 1 AND cleaned AND is_exaa) AS in_ledger_exaa,
       (first_b1_dt IS NOT NULL AND cleaned AND is_exaa) AS touched_b1,
       -- touched-B1 EOM-outcome class, b21 verbatim (NULL when not touched_b1)
       CASE
         WHEN NOT (first_b1_dt IS NOT NULL AND cleaned AND is_exaa) THEN NULL
         WHEN jan_co_dt >= DATE '2025-01-01' AND jan_co_dt < DATE '2025-02-01'
           THEN 'd. charged off in January'
         WHEN eom_bucket = 0
           THEN 'a. current at 31 Jan (cured in month)'
         WHEN eom_bucket = 1
           THEN 'b. bucket 1 at 31 Jan'
         WHEN eom_bucket >= 2
           THEN 'c. bucket 2+ at 31 Jan (rolled past DQ1 within January)'
       END AS touched_b1_class
FROM pop_base
