-- Tier 5 | Period calibration: inbound volume, account-id fill, transcript coverage by month (full history)
-- Run FIRST. Every other query sits on one of three pinned windows:
--   W1 ops     = 2026-01-01 .. 2026-07-01  (call-only stats; fresh, complete months)
--   W2 joined  = 2025-10-01 .. 2026-03-01  (same-month account joins; the account
--                copy's newest snapshot is 2026-03-07, so W2 ends at 2026-02)
--   W3 outcome = 2024-07-01 .. 2025-07-01  (funnel + vintages; every call month
--                keeps >= 8 account months of outcome runway before that edge)
-- This query confirms the windows: check that W3 months show the flat ~72-78%
-- acctid fill and healthy transcript coverage. A broken month here means shift
-- the window BEFORE running the f-series funnel queries.
-- The in-progress final month is included but partial: read it accordingly.
WITH inb AS (
    SELECT contactid, acctid,
           cast(date_trunc('month', "date") AS date) AS call_month
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
),
t AS (SELECT DISTINCT contactid FROM "contactcenter_bdp_db"."transcript")
SELECT inb.call_month,
       count(*) AS inbound_calls,
       round(100.0 * count_if(inb.acctid IS NOT NULL) / count(*), 1) AS pct_with_acctid,
       round(100.0 * count(t.contactid) / count(*), 1) AS pct_with_transcript
FROM inb
LEFT JOIN t ON inb.contactid = t.contactid
GROUP BY 1
ORDER BY 1
