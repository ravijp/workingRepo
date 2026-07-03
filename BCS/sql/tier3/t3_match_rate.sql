-- Tier 3 | Call-to-account match rate (inbound, last 6 complete ACCOUNT months)
-- Of inbound calls, how many carry an account id, and how many of those ids
-- resolve to the account master? This gates every cross-table analysis.
-- Window anchored to the ACCOUNT table's clock (its newest complete month),
-- not the call table's: the account copy trails the calls, and call months
-- past its edge under-match against a master that cannot know newly opened
-- accounts. Self-heals when the account copy refreshes.
-- f4_match_by_auth splits this rate by authentication outcome.
WITH am AS (
    SELECT date_add('month', -1,
               max(date_trunc('month', date(date_parse(eff_dt, '%Y%m%d'))))) AS m1
    FROM "fmt_acct_dba"."fmt_acct_c" WHERE sfx_nbr = 0
),
inb AS (
    SELECT contactid, acctid
    FROM "contactcenter_bdp_db"."call"
    CROSS JOIN am
    WHERE "date" >= cast(date_add('month', -5, am.m1) AS date)
      AND "date" < cast(date_add('month', 1, am.m1) AS date)
      AND effdt >= '2025-06-01' AND effdt < '2026-04-01'
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
