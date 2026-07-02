-- Tier 1 | Transcript table size and history depth
-- How many utterances and distinct calls do transcripts cover, and how far back does the history go?
SELECT count(*) AS row_count,
       count(DISTINCT contactid) AS distinct_calls,
       min(calldate) AS first_calldate,
       max(calldate) AS last_calldate,
       round(100.0 * count_if(content IS NOT NULL AND content <> '') / count(*), 1) AS pct_with_text
FROM "contactcenter_bdp_db"."transcript"
