-- Tier 1 | Account table size and freshness
-- How big is the account master, and how fresh is its newest monthly snapshot?
-- eff_dt is a yyyymmdd string; sfx_nbr = 0 keeps the current card per account.
SELECT count(*) AS row_count,
       count(DISTINCT extnl_acct_id) AS distinct_accounts,
       min(eff_dt) AS oldest_snapshot,
       max(eff_dt) AS newest_snapshot,
       count(DISTINCT eff_dt) AS monthly_snapshots
FROM "fmt_acct_dba"."fmt_acct_c"
WHERE sfx_nbr = 0
