-- Tier 2 | Charge-offs by month (last 24 months of table history)
-- How many accounts charge off per month, and for how much?
-- One row per account (min charge-off date, max amount) to undo the monthly-snapshot repetition.
WITH co AS (
    SELECT extnl_acct_id,
           min(try_cast(chrgoff_dt AS date)) AS co_dt,
           max(try_cast(chrgoff_amt AS double)) AS amt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND chrgoff_dt IS NOT NULL
    GROUP BY 1
),
mxco AS (SELECT date_trunc('month', max(co_dt)) AS m1 FROM co)
SELECT cast(date_trunc('month', co_dt) AS date) AS month,
       count(*) AS accounts_charged_off,
       round(sum(amt), 0) AS chargeoff_amount
FROM co, mxco
WHERE co_dt >= date_add('month', -24, mxco.m1)
  AND co_dt < mxco.m1
GROUP BY 1
ORDER BY 1
