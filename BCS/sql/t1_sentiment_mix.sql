-- Tier 1 | Utterance sentiment by speaker
-- How does per-utterance sentiment (POSITIVE / NEUTRAL / NEGATIVE) split for customers vs agents?
SELECT coalesce(participantid, '(blank)') AS participantid,
       coalesce(sentiment, '(blank)') AS sentiment,
       count(*) AS utterances
FROM "contactcenter_bdp_db"."transcript"
GROUP BY 1, 2
ORDER BY 1, 3 DESC
