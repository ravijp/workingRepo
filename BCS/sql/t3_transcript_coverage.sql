-- Tier 3 | Transcript coverage of inbound calls, by month (last 6 months)
-- What share of inbound calls has at least one transcript row?
-- Coverage is known to be partial; this measures it month by month.
WITH mx AS (SELECT date_trunc('month', max("date")) AS m1 FROM "contactcenter_bdp_db"."call"),
inb AS (
    SELECT contactid, cast(date_trunc('month', "date") AS date) AS month
    FROM "contactcenter_bdp_db"."call", mx
    WHERE "date" >= date_add('month', -6, mx.m1)
      AND "date" < mx.m1
      AND initiationmethod = 'INBOUND'
),
t AS (SELECT DISTINCT contactid FROM "contactcenter_bdp_db"."transcript")
SELECT inb.month,
       count(*) AS inbound_calls,
       count(t.contactid) AS with_transcript,
       round(100.0 * count(t.contactid) / count(*), 1) AS pct_with_transcript
FROM inb
LEFT JOIN t ON inb.contactid = t.contactid
GROUP BY 1
ORDER BY 1
