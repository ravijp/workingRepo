-- Tier 7 | Call effort by delinquency bucket (one account month)
-- How much conversation does each bucket take? Handle time from the call log,
-- turns and minutes from the transcripts, per same-month bucket. Deep-bucket
-- calls that run long and still leak are the cost side of the capture case.
-- Month anchored one month before the newest account month (the newest is a
-- partial copy).
WITH latest AS (
    SELECT max(date_trunc('month', date(date_parse(eff_dt, '%Y%m%d')))) AS d
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
    CROSS JOIN latest
    WHERE sfx_nbr = 0
      AND date_trunc('month', date(date_parse(eff_dt, '%Y%m%d')))
          = date_add('month', -1, latest.d)
),
acct AS (
    SELECT extnl_acct_id, max(bucket) AS bucket FROM snap GROUP BY 1
),
inb AS (
    SELECT c.contactid, a.bucket,
           try_cast(c.totalhandletime AS double) AS handle_s
    FROM "contactcenter_bdp_db"."call" c
    CROSS JOIN latest
    JOIN acct a
      ON trim(cast(c.acctid AS varchar)) = trim(cast(a.extnl_acct_id AS varchar))
    WHERE cast(date_trunc('month', c."date") AS date)
          = cast(date_add('month', -1, latest.d) AS date)
      AND c.effdt >= '2025-10-01' AND c.effdt < '2026-04-01'
      AND c.initiationmethod = 'INBOUND'
      AND c.acctid IS NOT NULL
),
tx AS (
    SELECT t.contactid,
           count_if(t.participantid = 'CUSTOMER') AS customer_turns,
           max(try_cast(t.endmillis AS bigint)) / 60000.0 AS minutes
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM inb) i ON t.contactid = i.contactid
     AND t.effdt >= '2025-09-01' AND t.effdt < '2026-05-01'
    GROUP BY 1
)
SELECT i.bucket AS dpd_bucket,
       count(*) AS calls,
       round(avg(i.handle_s), 0) AS avg_handle_s,
       round(avg(t.customer_turns), 1) AS avg_customer_turns,
       round(avg(t.minutes), 1) AS avg_minutes
FROM inb i
LEFT JOIN tx t ON i.contactid = t.contactid
GROUP BY 1
ORDER BY 1
