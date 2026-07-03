-- Tier 9 | Abandons on the money path: delinquent callers lost before an agent
-- An abandoned delinquent call is intent that never reached a human. Per
-- bucket: the abandon rate, and of the abandoned, how many called back within
-- 7 days (self-recovered) vs went silent (lost intent - the purest friction
-- leakage in the data).
-- One complete account month for the base calls; the 7-day recontact search
-- runs on the call table alone so it can cross the month edge.
WITH am AS (
    SELECT date_add('month', -1,
               max(date_trunc('month', date(date_parse(eff_dt, '%Y%m%d'))))) AS m1
    FROM "fmt_acct_dba"."fmt_acct_c" WHERE sfx_nbr = 0
),
snap AS (
    SELECT extnl_acct_id,
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
      AND date_trunc('month', date(date_parse(eff_dt, '%Y%m%d'))) = am.m1
),
acct AS (
    SELECT extnl_acct_id, max(bucket) AS bucket FROM snap GROUP BY 1
),
inb AS (
    SELECT c.contactid, c."date" AS call_dt,
           trim(cast(c.acctid AS varchar)) AS acct_key,
           a.bucket,
           CASE WHEN try_cast(c.abandoned AS integer) = 1 THEN 1 ELSE 0 END AS abandoned
    FROM "contactcenter_bdp_db"."call" c
    CROSS JOIN am
    JOIN acct a
      ON trim(cast(c.acctid AS varchar)) = trim(cast(a.extnl_acct_id AS varchar))
    WHERE cast(date_trunc('month', c."date") AS date) = cast(am.m1 AS date)
      AND c.effdt >= '2025-10-01' AND c.effdt < '2026-04-01'
      AND c.initiationmethod = 'INBOUND'
      AND c.acctid IS NOT NULL
),
recontact AS (
    SELECT DISTINCT i.contactid
    FROM inb i
    JOIN "contactcenter_bdp_db"."call" c2
      ON trim(cast(c2.acctid AS varchar)) = i.acct_key
     AND c2.initiationmethod = 'INBOUND'
     AND c2."date" > i.call_dt
     AND c2."date" <= date_add('day', 7, i.call_dt)
     AND c2.effdt >= '2025-10-01' AND c2.effdt < '2026-04-01'
    WHERE i.abandoned = 1
)
SELECT i.bucket AS dpd_bucket,
       count(*) AS inbound_calls,
       sum(i.abandoned) AS abandoned_calls,
       round(100.0 * sum(i.abandoned) / count(*), 2) AS pct_abandoned,
       round(100.0 * count_if(i.abandoned = 1 AND r.contactid IS NOT NULL)
             / greatest(sum(i.abandoned), 1), 1) AS pct_recontact_7d
FROM inb i
LEFT JOIN recontact r ON i.contactid = r.contactid
GROUP BY 1
ORDER BY 1
