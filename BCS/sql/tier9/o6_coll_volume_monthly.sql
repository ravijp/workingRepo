-- Tier 9 | Control total: inbound volume on collections queues, by month
-- The funnel's denominator is a candidate LOWER BOUND until this table's
-- inbound universe reconciles against the workforce-management call counts.
-- This produces the reconcilable number: monthly inbound calls landing on
-- collections queues (queue name starting COLL), with the account-id fill
-- beside it. Compare against the ops-reported monthly volumes offline; a
-- large gap means calls exist that this table does not carry, and every
-- funnel count inherits that caveat. Last 6 complete call months.
WITH mx AS (SELECT date_trunc('month', max("date")) AS m1 FROM "contactcenter_bdp_db"."call" WHERE effdt < cast(date_add('day', -1, current_date) AS varchar))
SELECT cast(date_trunc('month', "date") AS date) AS month,
       count(*) AS coll_queue_calls,
       count_if(acctid IS NOT NULL
                AND trim(cast(acctid AS varchar)) <> '') AS with_acctid,
       round(100.0 * count_if(acctid IS NOT NULL
                              AND trim(cast(acctid AS varchar)) <> '') / count(*), 1)
           AS pct_with_acctid
FROM "contactcenter_bdp_db"."call", mx
WHERE "date" >= date_add('month', -6, mx.m1)
  AND "date" < mx.m1
  AND effdt >= '2025-11-01' AND effdt < cast(date_add('day', -1, current_date) AS varchar)
  AND initiationmethod = 'INBOUND'
  AND upper(coalesce(cast(queue AS varchar), '')) LIKE 'COLL%'
GROUP BY 1
ORDER BY 1
