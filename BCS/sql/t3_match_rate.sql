-- Tier 3 | Call-to-account match rate (inbound, last 6 months)
-- Of recent inbound calls, how many carry an account id, and how many of those ids
-- resolve to the account master? This gates every cross-table analysis.
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call"),
inb AS (
    SELECT contactid, acctid
    FROM "contactcenter_bdp_db"."call", mx
    WHERE "date" > date_add('month', -6, mx.d)
      AND initiationmethod = 'INBOUND'
),
acct AS (
    SELECT DISTINCT extnl_acct_id
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
)
SELECT count(*) AS inbound_calls,
       count_if(inb.acctid IS NOT NULL) AS calls_with_acctid,
       count(a.extnl_acct_id) AS calls_matched_to_account,
       round(100.0 * count_if(inb.acctid IS NOT NULL) / count(*), 1) AS pct_with_acctid,
       round(100.0 * count(a.extnl_acct_id) / count(*), 1) AS pct_matched
FROM inb
LEFT JOIN acct a
  ON trim(cast(inb.acctid AS varchar)) = trim(cast(a.extnl_acct_id AS varchar))
