-- Tier 3 | The invisible pool: where unmatched calls live, and whether they talk payment
-- A quarter of inbound calls carry no account id and vanish from every joined
-- read. Transcripts join on contactid - no account id needed - so this reads
-- the invisible pool directly: which queues it rings, and how often the
-- customer talks payment there. If collections-relevant queues run high AND
-- payment talk is rich, the funnel undercounts exactly its target population
-- (pairs with f4_match_by_auth: that says WHO fails to match, this says WHERE
-- and WITH WHAT INTENT). Last complete call month to bound the text scan.
WITH mx AS (SELECT date_trunc('month', max("date")) AS m1 FROM "contactcenter_bdp_db"."call"),
unm AS (
    SELECT contactid,
           coalesce(cast(queue AS varchar), '(blank)') AS queue
    FROM "contactcenter_bdp_db"."call", mx
    WHERE "date" >= date_add('month', -1, mx.m1)
      AND "date" < mx.m1
      AND initiationmethod = 'INBOUND'
      AND (acctid IS NULL OR trim(cast(acctid AS varchar)) = '')
),
tx AS (
    SELECT t.contactid,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND regexp_like(lower(t.content),
                        'pay|paid|payment|settle|payment plan|arrangement|work something out'))
               AS pay_utts
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN unm u ON t.contactid = u.contactid
    GROUP BY 1
)
SELECT u.queue,
       count(*) AS unmatched_calls,
       round(100.0 * count(x.contactid) / count(*), 1) AS pct_with_transcript,
       round(100.0 * count_if(x.pay_utts > 0)
             / greatest(count(x.contactid), 1), 1) AS pct_payment_language
FROM unm u
LEFT JOIN tx x ON u.contactid = x.contactid
GROUP BY 1
ORDER BY 2 DESC
LIMIT 15
