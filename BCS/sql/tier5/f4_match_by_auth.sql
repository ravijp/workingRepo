-- Tier 5 | Bias gate: account match rate by authentication outcome (W2: 2025-10 .. 2026-02)
-- The funnel only sees calls that resolve to an account. If unmatched calls
-- concentrate among callers who FAILED authentication - the very friction
-- population an inbound-capture case targets - every funnel rate is biased
-- toward the easy callers. This splits acctid fill and account-master match
-- by authenticationstatus so that bias is measured, not assumed.
WITH inb AS (
    SELECT contactid, acctid,
           coalesce(cast(authenticationstatus AS varchar), '(blank)') AS auth
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2025-10-01' AND "date" < DATE '2026-03-01'
),
acct AS (
    SELECT DISTINCT extnl_acct_id
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
)
SELECT inb.auth AS authenticationstatus,
       count(*) AS calls,
       count_if(inb.acctid IS NOT NULL) AS with_acctid,
       round(100.0 * count_if(inb.acctid IS NOT NULL) / count(*), 1) AS pct_with_acctid,
       count(a.extnl_acct_id) AS matched_master,
       round(100.0 * count(a.extnl_acct_id) / count(*), 1) AS pct_matched_master
FROM inb
LEFT JOIN acct a
  ON trim(cast(inb.acctid AS varchar)) = trim(cast(a.extnl_acct_id AS varchar))
GROUP BY 1
ORDER BY 2 DESC
