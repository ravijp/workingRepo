-- Tier 12 | PII surface probe: is there a non-sensitive transcript twin?
-- On the Oracle/PCDS side, tables come in twins: *_HI_SNSTV (raw 16-digit
-- card number, PII-gated) vs *_NON_PRV (masked, accessible) - validated on
-- V_REAGEEVENTS in the 07-03..08 chat. Before ANY transcript sample leaves
-- Athena for the Copilot read, we need to know whether the Athena transcript
-- surface has an equivalent split, and which tables/views are in scope.
-- Metadata only: information_schema, no data rows scanned. This probe does
-- NOT authorize an export; the export itself waits on the A-066 governance
-- answer.
SELECT table_schema AS m3_table_schema,
       table_name   AS m3_table_name,
       CASE
         WHEN lower(table_name) LIKE '%snstv%'   THEN 'a. sensitive-twin name pattern'
         WHEN lower(table_name) LIKE '%non_prv%' THEN 'b. non-private-twin name pattern'
         WHEN lower(table_name) LIKE '%prv%'     THEN 'c. other prv pattern'
         WHEN lower(table_name) LIKE '%transcript%' THEN 'd. transcript-named'
         ELSE 'e. contact-center schema member'
       END AS m3_pattern
FROM information_schema.tables
WHERE lower(table_name) LIKE '%snstv%'
   OR lower(table_name) LIKE '%non_prv%'
   OR lower(table_name) LIKE '%prv%'
   OR lower(table_name) LIKE '%transcript%'
   OR table_schema = 'contactcenter_bdp_db'
ORDER BY m3_pattern, m3_table_schema, m3_table_name
