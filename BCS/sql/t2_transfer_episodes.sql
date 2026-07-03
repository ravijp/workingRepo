-- Tier 2 | Inbound calls that spawn transfer legs, by month (last 6 complete months)
-- The correct transfer read: transfertype on INBOUND legs is always
-- 'Not A Transfer' because transfers are separate TRANSFER / QUEUE_TRANSFER
-- legs. This chains those legs back to their inbound origin via
-- initialcontactid. initialcontactid fill was low in profiling, so treat
-- the result as a LOWER BOUND on transfer intensity.
WITH mx AS (SELECT date_trunc('month', max("date")) AS m1 FROM "contactcenter_bdp_db"."call"),
inb AS (
    SELECT contactid, cast(date_trunc('month', "date") AS date) AS month
    FROM "contactcenter_bdp_db"."call", mx
    WHERE "date" >= date_add('month', -6, mx.m1)
      AND "date" < mx.m1
      AND initiationmethod = 'INBOUND'
),
tl AS (
    SELECT initialcontactid, count(*) AS legs
    FROM "contactcenter_bdp_db"."call", mx
    WHERE "date" >= date_add('month', -6, mx.m1)
      AND "date" < mx.m1
      AND initiationmethod IN ('TRANSFER', 'QUEUE_TRANSFER')
      AND initialcontactid IS NOT NULL
    GROUP BY 1
)
SELECT inb.month,
       count(*) AS inbound_calls,
       count(tl.initialcontactid) AS with_transfer_leg,
       round(100.0 * count(tl.initialcontactid) / count(*), 1) AS pct_with_transfer_leg,
       coalesce(sum(tl.legs), 0) AS transfer_legs
FROM inb
LEFT JOIN tl ON inb.contactid = tl.initialcontactid
GROUP BY 1
ORDER BY 1
