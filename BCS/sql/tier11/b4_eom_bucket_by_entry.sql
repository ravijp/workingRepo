-- Tier 11 | Month-END bucket distribution by entry class, 202501
-- Set against Ishant's DLNQT_CD_M1 distribution (his codes 1-7). On a pure
-- days-past-due ladder a new-roll entrant can only land at month-END bucket
-- 1 (or 0 if cured within the month); deeper new-roll rows here would mean
-- this ladder and ASP DLNQT_CD disagree on semantics. Ishant's deep entries
-- (M1=5 at 42k) must come from no-prior-record accounts or a different code
-- meaning - this query shows which side our data takes.
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
           try_cast(acct_bal_amt AS double) AS bal
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20241201' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           max_by(bal, eff_dt) AS eom_bal
    FROM snap GROUP BY 1, 2
),
base AS (
    SELECT j.extnl_acct_id, j.max_bucket, j.eom_bucket, j.eom_bal,
           p.max_bucket AS prev_max_bucket,
           p.eom_bucket AS prev_eom_bucket,
           (p.extnl_acct_id IS NOT NULL) AS has_prior_row
    FROM (SELECT * FROM monthly WHERE ym = '202501') j
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202412') p
      ON j.extnl_acct_id = p.extnl_acct_id
)
SELECT eom_bucket AS eom_bucket_202501,
       CASE
         WHEN NOT has_prior_row THEN 'c. no prior row'
         WHEN prev_eom_bucket = 0 THEN 'a. new roll (202412 EOM = 0)'
         ELSE 'b. already delinquent (202412 EOM >= 1)'
       END AS entry_class_202501,
       count(*) AS bridge_accounts,
       round(sum(eom_bal), 0) AS bridge_eom_balance
FROM base
WHERE eom_bucket >= 1
GROUP BY 1, 2
ORDER BY 1, 2
