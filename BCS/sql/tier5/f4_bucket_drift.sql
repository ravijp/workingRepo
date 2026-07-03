-- Tier 5 | Bucket-drift gate: month-max bucket vs the bucket as of the call date
-- Every join in this kit tags a call with its account's MONTH-MAX bucket -
-- but the account has ~14 snapshots a month, and a call on the 2nd can be
-- tagged with a bucket the account only reached on the 28th, after the call.
-- This measures that drift on one complete account month: per month-max
-- bucket, the share of episodes where the latest snapshot ON OR BEFORE the
-- call date says current, a lower bucket, or the same-or-higher bucket.
-- A small drift validates the month-grain kit; a large one is its error band
-- (worst near charge-off, where a month moves an account across the line).
WITH am AS (
    SELECT date_add('month', -1,
               max(date_trunc('month', date(date_parse(eff_dt, '%Y%m%d'))))) AS m1
    FROM "fmt_acct_dba"."fmt_acct_c" WHERE sfx_nbr = 0
),
snap AS (
    SELECT extnl_acct_id,
           date(date_parse(eff_dt, '%Y%m%d')) AS snap_dt,
           date_trunc('month', date(date_parse(eff_dt, '%Y%m%d'))) AS m,
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
           END AS bucket
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN am
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= cast(date_add('month', -1, am.m1) AS date)
      AND date(date_parse(eff_dt, '%Y%m%d')) < cast(date_add('month', 1, am.m1) AS date)
),
monthmax AS (
    SELECT extnl_acct_id, max(bucket) AS mm_bucket
    FROM snap
    CROSS JOIN am
    WHERE m = am.m1
    GROUP BY 1
),
episodes AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key, "date" AS call_dt
    FROM (
        SELECT acctid, "date",
               row_number() OVER (PARTITION BY trim(cast(acctid AS varchar)), "date"
                                  ORDER BY initiationtimestamp) AS rn
        FROM "contactcenter_bdp_db"."call"
        CROSS JOIN am
        WHERE cast(date_trunc('month', "date") AS date) = cast(am.m1 AS date)
          AND initiationmethod = 'INBOUND'
          AND acctid IS NOT NULL
    )
    WHERE rn = 1
),
asof AS (
    SELECT e.acct_key, e.call_dt,
           max_by(s.bucket, s.snap_dt) AS asof_bucket
    FROM episodes e
    JOIN snap s
      ON e.acct_key = trim(cast(s.extnl_acct_id AS varchar))
     AND s.snap_dt <= e.call_dt
     AND s.snap_dt >= date_add('day', -45, e.call_dt)
    GROUP BY 1, 2
),
j AS (
    SELECT m.mm_bucket, a.asof_bucket
    FROM asof a
    JOIN monthmax m ON a.acct_key = trim(cast(m.extnl_acct_id AS varchar))
)
SELECT mm_bucket AS dpd_bucket,
       count(*) AS episodes,
       round(100.0 * count_if(asof_bucket = mm_bucket) / count(*), 1) AS pct_asof_same,
       round(100.0 * count_if(asof_bucket = 0 AND mm_bucket >= 1) / count(*), 1) AS pct_asof_current,
       round(100.0 * count_if(asof_bucket > 0 AND asof_bucket < mm_bucket) / count(*), 1) AS pct_asof_lower,
       round(100.0 * count_if(asof_bucket > mm_bucket) / count(*), 1) AS pct_asof_higher
FROM j
GROUP BY 1
ORDER BY 1
