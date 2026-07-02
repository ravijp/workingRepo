-- Tier 3 | Who opens and who closes the call (inbound, last 1 month)
-- Conversation shape: who speaks first and last, and how many turns each side takes.
WITH mx AS (SELECT max("date") AS d FROM "contactcenter_bdp_db"."call"),
inb AS (
    SELECT contactid
    FROM "contactcenter_bdp_db"."call", mx
    WHERE "date" > date_add('month', -1, mx.d)
      AND initiationmethod = 'INBOUND'
),
j AS (
    SELECT t.contactid, t.participantid,
           try_cast(t.beginmillis AS bigint) AS b,
           try_cast(t.endmillis AS bigint) AS e
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN inb ON t.contactid = inb.contactid
),
per_call AS (
    SELECT contactid,
           min_by(participantid, b) AS first_speaker,
           max_by(participantid, e) AS last_speaker,
           count_if(participantid = 'CUSTOMER') AS customer_turns,
           count_if(participantid = 'AGENT') AS agent_turns
    FROM j
    GROUP BY 1
)
SELECT first_speaker,
       last_speaker,
       count(*) AS calls,
       round(avg(customer_turns), 1) AS avg_customer_turns,
       round(avg(agent_turns), 1) AS avg_agent_turns
FROM per_call
GROUP BY 1, 2
ORDER BY 3 DESC
LIMIT 10
