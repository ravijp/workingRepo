-- Tier 12 | Mined language: customer phrases that separate charged-off from
-- surviving accounts, still-DQ1 class (the CO-vs-not companion to m1)
-- Same episode base as m1 (class b: month-MAX B1 entrants still DQ1 at
-- Jan 31), but the outcome is the ACCOUNT's 12-month charge-off (forward
-- chrgoff_dt scan, END_12M = 202512) instead of the 30-day capture gate.
-- Phrases here mark risk, not capture mechanics; a phrase high on both m1's
-- leak side and this CO side is the strongest lexicon candidate.
-- ONE transcript pass; no window functions over utterances; window functions
-- only over the tiny counted-phrase table. Support >= 40 episodes; all-filler
-- phrases dropped. Association, not causation - words mark accounts.
WITH snap AS (
    SELECT extnl_acct_id,
           substr(eff_dt, 1, 6) AS ym,
           eff_dt,
           CASE
             WHEN past_due_271_up_amt  > 0 THEN 10
             WHEN past_due_241_270_amt > 0 THEN 9
             WHEN past_due_211_240_amt > 0 THEN 8
             WHEN past_due_181_210_amt > 0 THEN 7
             WHEN past_due_151_180_amt > 0 THEN 6
             WHEN past_due_121_150_amt > 0 THEN 5
             WHEN past_due_91_120_amt  > 0 THEN 4
             WHEN past_due_61_90_amt   > 0 THEN 3
             WHEN past_due_31_60_amt   > 0 THEN 2
             WHEN past_due_1_30_amt    > 0 THEN 1
             ELSE 0
           END AS bucket,
           try_cast(chrgoff_dt AS date) AS co_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20241201' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket) AS max_bucket,
           max_by(bucket, eff_dt) AS eom_bucket,
           min(co_dt) AS co_dt
    FROM snap GROUP BY 1, 2
),
base_acct AS (
    SELECT j.extnl_acct_id, j.max_bucket, j.eom_bucket, j.co_dt,
           p.max_bucket AS prev_max_bucket
    FROM (SELECT * FROM monthly WHERE ym = '202501') j
    LEFT JOIN (SELECT * FROM monthly WHERE ym = '202412') p
      ON j.extnl_acct_id = p.extnl_acct_id
),
cohort AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key
    FROM base_acct
    WHERE max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0
      AND eom_bucket >= 1
      AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
),
future_co AS (
    SELECT extnl_acct_id, min(try_cast(chrgoff_dt AS date)) AS co_dt_future
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20260101'
      AND chrgoff_dt IS NOT NULL
    GROUP BY 1
),
inb AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key, contactid,
           "date" AS call_dt, initiationtimestamp
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2025-01-01' AND "date" < DATE '2025-02-01'
      AND effdt >= '2025-01-01' AND effdt < '2025-02-02'
      AND coalesce(cast(producttype AS varchar), '') <> 'BUSINESS_CARD'
),
episodes AS (
    SELECT acct_key, contactid, call_dt
    FROM (
        SELECT acct_key, contactid, call_dt,
               row_number() OVER (PARTITION BY acct_key, call_dt
                                  ORDER BY initiationtimestamp) AS rn
        FROM inb
        WHERE acct_key IS NOT NULL AND acct_key <> ''
    )
    WHERE rn = 1
),
ep AS (
    SELECT e.contactid,
           CASE WHEN f.co_dt_future >= DATE '2025-01-01'
                 AND f.co_dt_future < DATE '2026-01-01'
                THEN 1 ELSE 0 END AS co_12m
    FROM cohort c
    JOIN episodes e ON e.acct_key = c.acct_key
    LEFT JOIN future_co f
      ON trim(cast(f.extnl_acct_id AS varchar)) = c.acct_key
),
totals AS (
    SELECT count_if(co_12m = 1) AS n_co,
           count_if(co_12m = 0) AS n_ok
    FROM ep
),
words AS (
    SELECT d.contactid, d.co_12m,
           filter(split(regexp_replace(lower(t.content), '[^a-z ]', ' '), ' '),
                  x -> x <> '') AS w
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN ep d ON t.contactid = d.contactid
     AND t.effdt >= '2025-01-01' AND t.effdt < '2025-02-02'
    WHERE t.participantid = 'CUSTOMER'
      AND t.content IS NOT NULL
),
phrases AS (
    SELECT contactid, co_12m, array_join(bg, ' ') AS phrase,
           cardinality(bg) AS phrase_len
    FROM words
    CROSS JOIN UNNEST(ngrams(w, 2) || CASE WHEN cardinality(w) >= 3
                                           THEN ngrams(w, 3)
                                           ELSE CAST(ARRAY[] AS array(array(varchar))) END) AS t (bg)
    WHERE cardinality(w) >= 2
      AND cardinality(filter(bg, x -> contains(
            ARRAY['the','a','an','i','you','to','of','and','is','it','that','my','me',
                  'on','for','in','we','be','this','have','do','was','so','are','not',
                  'with','your','um','uh','okay','yeah','know','like','just','can','get'],
            x))) < cardinality(bg)
),
counted AS (
    SELECT phrase, max(phrase_len) AS phrase_len,
           count(DISTINCT CASE WHEN co_12m = 1 THEN contactid END) AS eps_co,
           count(DISTINCT CASE WHEN co_12m = 0 THEN contactid END) AS eps_ok
    FROM phrases
    GROUP BY 1
    HAVING count(DISTINCT contactid) >= 40
),
rated AS (
    SELECT c.phrase, c.phrase_len, c.eps_co, c.eps_ok,
           round(100.0 * c.eps_co / greatest(b.n_co, 1), 2) AS pct_of_co,
           round(100.0 * c.eps_ok / greatest(b.n_ok, 1), 2) AS pct_of_ok
    FROM counted c
    CROSS JOIN totals b
),
ranked AS (
    SELECT phrase, phrase_len, eps_co, eps_ok, pct_of_co, pct_of_ok,
           round(pct_of_co / greatest(pct_of_ok, 0.05), 2) AS lift_co,
           round(pct_of_ok / greatest(pct_of_co, 0.05), 2) AS lift_ok,
           row_number() OVER (ORDER BY pct_of_co / greatest(pct_of_ok, 0.05) DESC) AS r_co,
           row_number() OVER (ORDER BY pct_of_ok / greatest(pct_of_co, 0.05) DESC) AS r_ok
    FROM rated
)
SELECT CASE WHEN r_co <= 20 THEN 'a. charge-off marker'
            ELSE 'b. survival marker' END AS m2_side,
       phrase AS m2_phrase,
       phrase_len AS m2_words,
       eps_co AS m2_eps_co12,
       eps_ok AS m2_eps_no_co,
       pct_of_co AS m2_pct_of_co12,
       pct_of_ok AS m2_pct_of_no_co,
       CASE WHEN r_co <= 20 THEN lift_co ELSE lift_ok END AS m2_lift
FROM ranked
WHERE r_co <= 20 OR r_ok <= 20
ORDER BY m2_side, m2_lift DESC
