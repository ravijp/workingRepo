-- Tier 8 | Learned lexicon: customer bigrams that separate paid from leaked calls
-- Instead of hand-writing payment lexicons, learn them: for delinquent-account
-- inbound calls, split by payment-within-30-days vs not, and rank customer
-- bigrams by how strongly they mark each outcome. 'a. leak-marker' phrases
-- appear disproportionately on calls with NO payment after; 'b. payment-marker'
-- phrases on calls that paid. The winners are candidate rules for a live
-- transcript read - measured on outcomes, not intuition.
-- Call month anchored two months before the newest account month (complete
-- following month for the payment check). Bigrams need >= 250 calls support;
-- pairs of pure filler words are dropped.
WITH am AS (
    SELECT max(date_trunc('month', date(date_parse(eff_dt, '%Y%m%d')))) AS d
    FROM "fmt_acct_dba"."fmt_acct_c" WHERE sfx_nbr = 0
),
snap AS (
    SELECT extnl_acct_id,
           date_trunc('month', date(date_parse(eff_dt, '%Y%m%d'))) AS m,
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
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(paymt_last_dt, '%d%b%Y') AS date))) AS pay_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    CROSS JOIN am
    WHERE sfx_nbr = 0
      AND date(date_parse(eff_dt, '%Y%m%d')) >= cast(date_add('month', -2, am.d) AS date)
      AND date(date_parse(eff_dt, '%Y%m%d')) < cast(am.d AS date)
),
monthly AS (
    SELECT extnl_acct_id, m, max(bucket) AS bucket, max(pay_dt) AS pay_dt
    FROM snap GROUP BY 1, 2
),
delq AS (
    SELECT c.contactid,
           CASE WHEN nxt.pay_dt IS NOT NULL
                 AND nxt.pay_dt >= c."date"
                 AND nxt.pay_dt <= date_add('day', 30, c."date")
                THEN 1 ELSE 0 END AS paid_30d
    FROM "contactcenter_bdp_db"."call" c
    CROSS JOIN am
    JOIN monthly s
      ON trim(cast(c.acctid AS varchar)) = trim(cast(s.extnl_acct_id AS varchar))
     AND cast(s.m AS date) = cast(date_add('month', -2, am.d) AS date)
    LEFT JOIN monthly nxt
      ON trim(cast(c.acctid AS varchar)) = trim(cast(nxt.extnl_acct_id AS varchar))
     AND cast(nxt.m AS date) = cast(date_add('month', -1, am.d) AS date)
    WHERE cast(date_trunc('month', c."date") AS date)
          = cast(date_add('month', -2, am.d) AS date)
      AND c.initiationmethod = 'INBOUND'
      AND c.acctid IS NOT NULL
      AND s.bucket >= 1
),
base AS (
    SELECT count(DISTINCT contactid) AS n_calls,
           count(DISTINCT CASE WHEN paid_30d = 1 THEN contactid END) AS n_paid,
           count(DISTINCT CASE WHEN paid_30d = 0 THEN contactid END) AS n_leaked
    FROM delq
),
words AS (
    SELECT d.contactid, d.paid_30d,
           filter(split(regexp_replace(lower(t.content), '[^a-z ]', ' '), ' '),
                  x -> x <> '') AS w
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN delq d ON t.contactid = d.contactid
    WHERE t.participantid = 'CUSTOMER'
      AND t.content IS NOT NULL
),
bigrams AS (
    SELECT contactid, paid_30d, array_join(bg, ' ') AS bigram,
           element_at(bg, 1) AS w1, element_at(bg, 2) AS w2
    FROM words
    CROSS JOIN UNNEST(ngrams(w, 2)) AS t (bg)
    WHERE cardinality(w) >= 2
),
counted AS (
    SELECT bigram,
           count(DISTINCT CASE WHEN paid_30d = 1 THEN contactid END) AS calls_paid,
           count(DISTINCT CASE WHEN paid_30d = 0 THEN contactid END) AS calls_leaked
    FROM bigrams
    WHERE NOT (w1 IN ('the','a','an','i','you','to','of','and','is','it','that','my','me',
                      'on','for','in','we','be','this','have','do','was','so','are','not',
                      'with','your','um','uh','okay','yeah','know','like','just','can','get')
           AND w2 IN ('the','a','an','i','you','to','of','and','is','it','that','my','me',
                      'on','for','in','we','be','this','have','do','was','so','are','not',
                      'with','your','um','uh','okay','yeah','know','like','just','can','get'))
    GROUP BY 1
    HAVING count(DISTINCT contactid) >= 250
),
rated AS (
    SELECT c.bigram, c.calls_paid, c.calls_leaked,
           round(100.0 * c.calls_paid / greatest(b.n_paid, 1), 2) AS pct_of_paid_calls,
           round(100.0 * c.calls_leaked / greatest(b.n_leaked, 1), 2) AS pct_of_leaked_calls
    FROM counted c
    CROSS JOIN base b
)
SELECT * FROM (
    SELECT 'a. leak-marker' AS side, bigram, calls_paid, calls_leaked,
           pct_of_paid_calls, pct_of_leaked_calls,
           round(pct_of_leaked_calls / greatest(pct_of_paid_calls, 0.05), 2) AS lift
    FROM rated
    ORDER BY lift DESC
    LIMIT 15
)
UNION ALL
SELECT * FROM (
    SELECT 'b. payment-marker' AS side, bigram, calls_paid, calls_leaked,
           pct_of_paid_calls, pct_of_leaked_calls,
           round(pct_of_paid_calls / greatest(pct_of_leaked_calls, 0.05), 2) AS lift
    FROM rated
    ORDER BY lift DESC
    LIMIT 15
)
ORDER BY side, lift DESC
