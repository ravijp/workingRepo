-- Tier 1 | Call table size, period, and account-link fill
-- How many call legs, over what period, and how often is the account id populated?
SELECT count(*) AS row_count,
       count(DISTINCT contactid) AS distinct_calls,
       min("date") AS first_call_date,
       max("date") AS last_call_date,
       round(100.0 * count_if(acctid IS NOT NULL) / count(*), 1) AS pct_with_acctid,
       round(100.0 * count_if(partyid IS NOT NULL) / count(*), 1) AS pct_with_partyid
FROM "contactcenter_bdp_db"."call"
