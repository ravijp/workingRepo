-- Tier 1 | Who speaks in transcripts
-- Utterance split by participant; any value beyond AGENT / CUSTOMER shows up here.
SELECT coalesce(participantid, '(blank)') AS participantid,
       count(*) AS utterances,
       count(DISTINCT contactid) AS calls,
       round(avg(length(content)), 0) AS avg_chars_per_utterance
FROM "contactcenter_bdp_db"."transcript"
GROUP BY 1
ORDER BY 2 DESC
