-- Tier 12 | Mined language: customer phrases that separate captured from
-- leaked episodes, still-DQ1 class (n1 extended to the cohort)
-- n1 learned bigrams platform-wide; this mines the January cohort's hardest
-- class: month-MAX B1 entrants still DQ1 at Jan 31 (b10/b11's class b, ~9,013
-- episodes). Bigrams AND trigrams in one explosion (ngrams 2 || ngrams 3),
-- ranked by capture-vs-leak lift. Winners are lexicon candidates for the
-- hybrid upgrade: every phrase proposed here or by the Copilot read gets
-- measured at full count before it enters the story.
-- ONE transcript pass by construction (b11's memory lesson): the words CTE is
-- referenced exactly once; no window functions over utterances; the only
-- window functions run over the tiny counted-phrase table. Support >= 40
-- episodes; phrases made only of filler words are dropped.
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
    -- class b only: month-MAX B1 entrant, still DQ1 at EOM, cleanup applied
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key
    FROM base_acct
    WHERE max_bucket = 1 AND coalesce(prev_max_bucket, 0) = 0
      AND eom_bucket >= 1
      AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
),
pay_snap AS (
    SELECT extnl_acct_id, eff_dt,
           date_trunc('month', date(date_parse(eff_dt, '%Y%m%d'))) AS m,
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(try_cast(paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS pay_dt,
           coalesce(try_cast(atmtc_paymt_last_dt AS date),
                    try(cast(date_parse(try_cast(atmtc_paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS auto_dt,
           coalesce(try_cast(nsf_last_paymt_dt AS date),
                    try(cast(date_parse(try_cast(nsf_last_paymt_dt AS varchar), '%d%b%Y') AS date))) AS nsf_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250301'
),
pay_monthly AS (
    SELECT extnl_acct_id, m,
           max(pay_dt) AS pay_dt,
           max(auto_dt) AS auto_dt,
           max(nsf_dt) AS nsf_dt
    FROM pay_snap GROUP BY 1, 2
),
pay_monthly2 AS (
    SELECT extnl_acct_id, m, pay_dt, auto_dt, nsf_dt,
           lead(pay_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_pay_dt,
           lead(auto_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_auto_dt,
           lead(nsf_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_nsf_dt
    FROM pay_monthly
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
    SELECT acct_key, contactid, call_dt,
           cast(date_trunc('month', call_dt) AS date) AS call_month
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
           CASE WHEN
                  (s.pay_dt IS NOT NULL
                   AND s.pay_dt >= e.call_dt
                   AND s.pay_dt <= date_add('day', 30, e.call_dt)
                   AND (s.auto_dt IS NULL OR s.auto_dt <> s.pay_dt)
                   AND (s.nsf_dt IS NULL OR s.nsf_dt <> s.pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.pay_dt))
                OR
                  (s.next_pay_dt IS NOT NULL
                   AND s.next_pay_dt >= e.call_dt
                   AND s.next_pay_dt <= date_add('day', 30, e.call_dt)
                   AND (s.next_auto_dt IS NULL OR s.next_auto_dt <> s.next_pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.next_pay_dt))
                THEN 1 ELSE 0 END AS captured
    FROM cohort c
    JOIN episodes e ON e.acct_key = c.acct_key
    LEFT JOIN pay_monthly2 s
      ON e.acct_key = trim(cast(s.extnl_acct_id AS varchar))
     AND e.call_month = cast(s.m AS date)
),
totals AS (
    SELECT count_if(captured = 1) AS n_captured,
           count_if(captured = 0) AS n_leaked
    FROM ep
),
words AS (
    SELECT d.contactid, d.captured,
           filter(split(regexp_replace(lower(t.content), '[^a-z ]', ' '), ' '),
                  x -> x <> '') AS w
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN ep d ON t.contactid = d.contactid
     AND t.effdt >= '2025-01-01' AND t.effdt < '2025-02-02'
    WHERE t.participantid = 'CUSTOMER'
      AND t.content IS NOT NULL
),
phrases AS (
    SELECT contactid, captured, array_join(bg, ' ') AS phrase,
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
           count(DISTINCT CASE WHEN captured = 1 THEN contactid END) AS eps_captured,
           count(DISTINCT CASE WHEN captured = 0 THEN contactid END) AS eps_leaked
    FROM phrases
    GROUP BY 1
    HAVING count(DISTINCT contactid) >= 40
),
rated AS (
    SELECT c.phrase, c.phrase_len, c.eps_captured, c.eps_leaked,
           round(100.0 * c.eps_captured / greatest(b.n_captured, 1), 2) AS pct_of_captured,
           round(100.0 * c.eps_leaked / greatest(b.n_leaked, 1), 2) AS pct_of_leaked
    FROM counted c
    CROSS JOIN totals b
),
ranked AS (
    SELECT phrase, phrase_len, eps_captured, eps_leaked,
           pct_of_captured, pct_of_leaked,
           round(pct_of_leaked / greatest(pct_of_captured, 0.05), 2) AS lift_leak,
           round(pct_of_captured / greatest(pct_of_leaked, 0.05), 2) AS lift_capture,
           row_number() OVER (ORDER BY pct_of_leaked / greatest(pct_of_captured, 0.05) DESC) AS r_leak,
           row_number() OVER (ORDER BY pct_of_captured / greatest(pct_of_leaked, 0.05) DESC) AS r_cap
    FROM rated
)
SELECT CASE WHEN r_leak <= 20 THEN 'a. leak-marker'
            ELSE 'b. capture-marker' END AS m1_side,
       phrase AS m1_phrase,
       phrase_len AS m1_words,
       eps_captured AS m1_eps_captured,
       eps_leaked AS m1_eps_leaked,
       pct_of_captured AS m1_pct_of_captured,
       pct_of_leaked AS m1_pct_of_leaked,
       CASE WHEN r_leak <= 20 THEN lift_leak ELSE lift_capture END AS m1_lift
FROM ranked
WHERE r_leak <= 20 OR r_cap <= 20
ORDER BY m1_side, m1_lift DESC
