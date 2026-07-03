-- Tier 3 | Transcript coverage by initiation method (last 6 months)
-- Transcribed distinct calls exceed total inbound calls, so transcripts
-- must also cover outbound and transfer legs. This sizes coverage per
-- method - it decides whether outbound conversations are readable too.
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call"),
c AS (
    SELECT contactid, coalesce(cast(initiationmethod AS varchar), '(blank)') AS im
    FROM "contactcenter_bdp_db"."call", mx
    WHERE "date" > date_add('month', -6, mx.d)
),
t AS (SELECT DISTINCT contactid FROM "contactcenter_bdp_db"."transcript")
SELECT c.im AS initiationmethod,
       count(*) AS calls,
       count(t.contactid) AS with_transcript,
       round(100.0 * count(t.contactid) / count(*), 1) AS pct_with_transcript
FROM c
LEFT JOIN t ON c.contactid = t.contactid
GROUP BY 1
ORDER BY 2 DESC
