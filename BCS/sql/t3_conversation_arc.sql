-- Tier 3 | Customer sentiment across the call (inbound, last 1 month)
-- Split each transcribed call into thirds by time: does customer sentiment
-- recover by the end of the call?
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call"),
inb AS (
    SELECT contactid
    FROM "contactcenter_bdp_db"."call", mx
    WHERE "date" > date_add('month', -1, mx.d)
      AND initiationmethod = 'INBOUND'
),
cust AS (
    SELECT t.contactid, t.sentiment,
           try_cast(t.beginmillis AS bigint) AS beginmillis
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN inb ON t.contactid = inb.contactid
    WHERE t.participantid = 'CUSTOMER'
),
pos AS (
    SELECT contactid, sentiment,
           ntile(3) OVER (PARTITION BY contactid ORDER BY beginmillis) AS call_third
    FROM cust
)
SELECT call_third,
       coalesce(sentiment, '(blank)') AS sentiment,
       count(*) AS utterances
FROM pos
GROUP BY 1, 2
ORDER BY 1, 2
