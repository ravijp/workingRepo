# Databricks notebook source
# MAGIC %md
# MAGIC # UC2 warehouse — base tables + insight layer
# MAGIC
# MAGIC Build three base tables ONCE, then run every insight as a short query on
# MAGIC top. Ports the verified tier-15 Athena/Trino SQL to Spark SQL and persists
# MAGIC with `CREATE OR REPLACE TABLE ... USING DELTA` into a temp catalog/schema
# MAGIC (Namit's pattern). Design: uc2-anchoring/uc2-warehouse-design.md.
# MAGIC
# MAGIC Base tables (built once, cells 1-3):
# MAGIC - `uc2_acct_month`     account x month spine (position, balance, cpc_class, charge-off, payment)
# MAGIC - `uc2_episode`        call spine (first-inbound-per-day, call-day position, capture)
# MAGIC - `uc2_episode_signal` transcript read (v2 lexicon flags + language_group)
# MAGIC
# MAGIC Every insight (cells 10+) is a JOIN + GROUP BY on these. CO 8/10/12 windows
# MAGIC are applied in the INSIGHT layer (a filter on uc2_acct_month.chrgoff_dt), so
# MAGIC changing the anchor never rebuilds a base table.
# MAGIC
# MAGIC Tie-out cells (1t/2t/3t) fail loud if a ported number drifts from the
# MAGIC tier-15 reference: 189,146 / 457,943,987 / 11,262 / 29,114.

# COMMAND ----------

# MAGIC %md ## Cell 0 — widgets / parameters

# COMMAND ----------

# Source catalog (Glue-federated; holds fmt_acct_c, the call table, the transcript table)
dbutils.widgets.text("source_catalog", "062108867742_glue_connectivity_catalog", "Source Catalog")
# Source schemas under that catalog. [VERIFY] on first run:
#   SHOW SCHEMAS IN 062108867742_glue_connectivity_catalog;
#   then SHOW TABLES to find fmt_acct_c / call / transcript.
# Athena had them as fmt_acct_dba and contactcenter_bdp_db; confirm the UC names.
dbutils.widgets.text("acct_schema", "fmt_acct_dba", "Account Schema [VERIFY]")
dbutils.widgets.text("cc_schema", "contactcenter_bdp_db", "Contact-center Schema [VERIFY]")

# Temp catalog / schema where UC2 base tables are written (Namit's working area)
dbutils.widgets.text("temp_catalog", "cda_model_shared", "Temp Catalog")
dbutils.widgets.text("temp_schema", "ecm_cld_model", "Temp Schema")

# Output S3 (also the SAS -> S3 landing zone for the later enrichment table)
dbutils.widgets.text("output_location", "s3://355538383407-edpss/ECM_CLD", "Output Location")

src   = dbutils.widgets.get("source_catalog")
acsch = dbutils.widgets.get("acct_schema")
ccsch = dbutils.widgets.get("cc_schema")
tcat  = dbutils.widgets.get("temp_catalog")
tsch  = dbutils.widgets.get("temp_schema")
out   = dbutils.widgets.get("output_location")

# Fully-qualified source table names (backtick `call`: reserved word in Spark SQL)
ACCT = f"`{src}`.`{acsch}`.fmt_acct_c"
CALL = f"`{src}`.`{ccsch}`.`call`"
TX   = f"`{src}`.`{ccsch}`.transcript"
# Base-table home
DB   = f"`{tcat}`.`{tsch}`"

print("ACCT =", ACCT)
print("CALL =", CALL)
print("TX   =", TX)
print("base tables ->", DB)

# COMMAND ----------

# MAGIC %md
# MAGIC ## Cell 1 — build `uc2_acct_month` (the account x month spine)
# MAGIC
# MAGIC One row per `extnl_acct_id x month`, Dec-2024 through 2026-01 (spans the Feb/Mar
# MAGIC transition logic AND the forward charge-off horizon in a single scan). Carries the
# MAGIC bucket ladder, cpc_class + is_ex_aa (computed ONCE here), balance, credit limit,
# MAGIC raw charge-off date/amount, payment-gate fields, and the pre-2025 cleanup flag.
# MAGIC Ported from tier-15 b14/b15 spine; Trino date fns -> Spark.

# COMMAND ----------

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_acct_month USING DELTA AS
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
           try_cast(acct_bal_amt AS double)       AS bal,
           try_cast(cr_lmt_origl_amt AS double)   AS cr_lmt_origl_amt,
           try_cast(cr_lmt_amt AS double)         AS cr_lmt_amt,
           try_cast(chrgoff_dt AS date)           AS co_dt,
           try_cast(chrgoff_amt AS double)        AS co_amt,
           try_cast(paymt_last_dt AS date)        AS pay_dt,
           try_cast(atmtc_paymt_last_dt AS date)  AS auto_dt,
           try_cast(nsf_last_paymt_dt AS date)    AS nsf_dt,
           clnt_prdct_cd
    FROM {ACCT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '20241201' AND eff_dt < '20260201'
),
monthly AS (
    SELECT extnl_acct_id, ym,
           max(bucket)                              AS max_bucket,
           max_by(bucket, eff_dt)                   AS eom_bucket,
           max(eff_dt)                              AS eff_dt_eom,
           max_by(bal, eff_dt)                      AS eom_bal,
           max_by(cr_lmt_origl_amt, eff_dt)         AS cr_lmt_origl_amt,
           max_by(cr_lmt_amt, eff_dt)               AS cr_lmt_amt,
           min(co_dt)                               AS co_dt,
           max_by(co_amt, eff_dt)                   AS co_amt,
           max_by(pay_dt, eff_dt)                   AS pay_dt,
           max_by(auto_dt, eff_dt)                  AS auto_dt,
           max_by(nsf_dt, eff_dt)                   AS nsf_dt,
           min(CASE WHEN bucket >= 1 THEN eff_dt END) AS first_dq_dt,
           min(CASE WHEN bucket  = 1 THEN eff_dt END) AS first_b1_dt,
           max_by(clnt_prdct_cd, eff_dt)            AS eom_cpc
    FROM snap GROUP BY extnl_acct_id, ym
)
SELECT trim(cast(extnl_acct_id AS string)) AS extnl_acct_id,
       ym, eff_dt_eom,
       eom_bucket, max_bucket, first_dq_dt, first_b1_dt,
       eom_bal, cr_lmt_origl_amt, cr_lmt_amt,
       eom_cpc,
       CASE
         WHEN eom_cpc IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                           'AA3','AC3','AM3','AA4','AC4','AM4')     THEN 'AA'
         WHEN eom_cpc IN ('BGC','BGM','CGM','GMR')                 THEN 'GM'
         WHEN eom_cpc IN ('FBS','IBS','U1C','U2C','U3C')           THEN 'Bronco'
         WHEN eom_cpc IN ('BHA','BJT','BJC','BFR','BWY','BBB')     THEN 'Biz'
         WHEN eom_cpc IN ('GAP','GP2','ONV','ON2','BRP','BR2','ATH','AT2',
                           'GPC','G2C','ONC','O2C','BRC','B2C','ATC','A2C')
                                                                     THEN 'CoBrand'
         WHEN eom_cpc IN ('8GP','8ON','8BR','8AT','9GP','9ON','9BR','9AT')
                                                                     THEN 'PLCC'
         ELSE 'OTHER'
       END AS cpc_class,
       (eom_cpc IS NULL OR trim(eom_cpc) = ''
        OR eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                           'AA3','AC3','AM3','AA4','AC4','AM4',
                           'BGC','BGM','CGM','GMR',
                           'FBS','IBS','U1C','U2C','U3C')) AS is_ex_aa,
       co_dt AS chrgoff_dt, co_amt AS chrgoff_amt,
       pay_dt, auto_dt, nsf_dt,
       (co_dt IS NOT NULL AND co_dt < DATE '2025-01-01') AS co_before_2025
FROM monthly
""")
print("built", f"{tcat}.{tsch}.uc2_acct_month")

# COMMAND ----------

# MAGIC %md ## Cell 1t — tie-out: the cleaned Jan bucket-1 ledger
# MAGIC Cleaned = drop pre-2025 charge-offs. ex-AA count MUST be 189,146 / balance
# MAGIC 457,943,987; full cleaned ledger 204,323. These are the tier-15 stop-gate.

# COMMAND ----------

_t = spark.sql(f"""
SELECT
  count_if(NOT co_before_2025 AND is_ex_aa)                                   AS exaa_accts,
  round(sum(CASE WHEN NOT co_before_2025 AND is_ex_aa THEN eom_bal END), 0)   AS exaa_bal,
  count_if(NOT co_before_2025)                                                AS cleaned_accts
FROM {DB}.uc2_acct_month
WHERE ym = '202501' AND eom_bucket = 1
""").first()
print(dict(_t.asDict()))
assert _t["exaa_accts"] == 189146, f"ex-AA ledger {_t['exaa_accts']} != 189146 — STOP, dialect/port drift"
assert abs(_t["exaa_bal"] - 457943987) <= 5, f"ex-AA balance {_t['exaa_bal']} off > $5 — STOP"
assert _t["cleaned_accts"] == 204323, f"cleaned ledger {_t['cleaned_accts']} != 204323 — STOP"
print("PASS: uc2_acct_month ties out (189,146 / 457,943,987 / 204,323)")

# COMMAND ----------

# MAGIC %md
# MAGIC ## Cell 2 — build `uc2_episode` (the call spine)
# MAGIC
# MAGIC First-inbound-per-account-per-day (business-card excluded, effdt-capped), with the
# MAGIC call-day position (bucket + days-since-first-dq) and the 30-day capture result
# MAGIC computed once against uc2_acct_month's payment fields. Ports b17/b19's call-day
# MAGIC snapshot join and b14/b15's capture gate. Uses a daily-snapshot read for the
# MAGIC call-day bucket (the account x month spine is month-grain and cannot give the
# MAGIC as-of-call-day position), semi-joined to the episode accounts.
# MAGIC
# MAGIC PORT NOTE (verify at cell 2t): the tier-15 capture gate read `pay_monthly` =
# MAGIC `max(pay_dt)` over the month's DAILY rows. Here the gate reads uc2_acct_month's
# MAGIC `pay_dt = max_by(pay_dt, eff_dt)` (the last-of-month value). For paymt_last_dt,
# MAGIC a monotone "last payment date" field, max-over-month and max-by-EOM are the SAME
# MAGIC value, so this is equivalent. The captured=6,029 assertion in 2t pins it: if the
# MAGIC substitution ever diverged, that count would move and the cell fails loud.

# COMMAND ----------

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_episode USING DELTA AS
WITH inb AS (
    SELECT trim(cast(acctid AS string)) AS extnl_acct_id, contactid,
           `date` AS call_dt, initiationtimestamp
    FROM {CALL}
    WHERE initiationmethod = 'INBOUND'
      AND `date` >= DATE '2025-01-01' AND `date` < DATE '2025-02-01'
      AND effdt >= '2025-01-01' AND effdt < '2025-02-02'
      AND coalesce(cast(producttype AS string), '') <> 'BUSINESS_CARD'
),
episodes AS (
    SELECT extnl_acct_id, contactid, call_dt,
           trunc(call_dt, 'MM') AS call_month
    FROM (
        SELECT extnl_acct_id, contactid, call_dt,
               row_number() OVER (PARTITION BY extnl_acct_id, call_dt
                                  ORDER BY initiationtimestamp) AS rn
        FROM inb
        WHERE extnl_acct_id IS NOT NULL AND extnl_acct_id <> ''
    )
    WHERE rn = 1
),
-- daily snapshot lookback for the call-day position (month spine is too coarse)
snap AS (
    SELECT trim(cast(extnl_acct_id AS string)) AS extnl_acct_id,
           eff_dt,
           to_date(eff_dt, 'yyyyMMdd') AS snap_dt,
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
    FROM {ACCT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '20240601' AND eff_dt < '20250201'
      AND trim(cast(extnl_acct_id AS string)) IN (SELECT extnl_acct_id FROM episodes)
),
callday AS (
    SELECT e.extnl_acct_id, e.call_dt,
           max_by(s.bucket, s.eff_dt)  AS callday_bucket,
           max_by(s.co_dt, s.eff_dt)   AS callday_co_dt,
           max(CASE WHEN s.bucket = 0 THEN s.snap_dt END) AS last_current_dt
    FROM episodes e
    JOIN snap s ON s.extnl_acct_id = e.extnl_acct_id AND s.snap_dt <= e.call_dt
    GROUP BY e.extnl_acct_id, e.call_dt
),
spell AS (
    SELECT c.extnl_acct_id, c.call_dt, c.callday_bucket, c.callday_co_dt,
           min(CASE WHEN s.bucket >= 1
                     AND s.snap_dt <= c.call_dt
                     AND s.snap_dt > coalesce(c.last_current_dt, DATE '1900-01-01')
                    THEN s.snap_dt END) AS spell_start_dt
    FROM callday c
    JOIN snap s ON s.extnl_acct_id = c.extnl_acct_id
    GROUP BY c.extnl_acct_id, c.call_dt, c.callday_bucket, c.callday_co_dt
),
pay AS (
    SELECT extnl_acct_id, ym, pay_dt, auto_dt, nsf_dt,
           trunc(to_date(concat(ym, '01'), 'yyyyMMdd'), 'MM') AS m
    FROM {DB}.uc2_acct_month
),
pay2 AS (
    SELECT extnl_acct_id, m, pay_dt, auto_dt, nsf_dt,
           lead(pay_dt)  OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_pay_dt,
           lead(auto_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_auto_dt,
           lead(nsf_dt)  OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_nsf_dt
    FROM pay
)
SELECT e.extnl_acct_id, e.contactid, e.call_dt, e.call_month,
       sp.callday_bucket, sp.callday_co_dt,
       datediff(e.call_dt, sp.spell_start_dt) AS days_since_first_dq,
       CASE WHEN
              (s.pay_dt IS NOT NULL
               AND s.pay_dt >= e.call_dt AND s.pay_dt <= date_add(e.call_dt, 30)
               AND (s.auto_dt IS NULL OR s.auto_dt <> s.pay_dt)
               AND (s.nsf_dt  IS NULL OR s.nsf_dt  <> s.pay_dt)
               AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.pay_dt))
            OR
              (s.next_pay_dt IS NOT NULL
               AND s.next_pay_dt >= e.call_dt AND s.next_pay_dt <= date_add(e.call_dt, 30)
               AND (s.next_auto_dt IS NULL OR s.next_auto_dt <> s.next_pay_dt)
               AND (s.next_nsf_dt  IS NULL OR s.next_nsf_dt  <> s.next_pay_dt))
            THEN 1 ELSE 0 END AS captured
FROM episodes e
JOIN spell sp ON sp.extnl_acct_id = e.extnl_acct_id AND sp.call_dt = e.call_dt
LEFT JOIN pay2 s ON s.extnl_acct_id = e.extnl_acct_id AND s.m = e.call_month
""")
print("built", f"{tcat}.{tsch}.uc2_episode")

# COMMAND ----------

# MAGIC %md ## Cell 2t — tie-out: episodes, capture gate, and the call-day stream
# MAGIC Three checks, because they pin three separate pieces of the port:
# MAGIC - ledger episodes = 11,262 (the first-inbound-per-day dedup + effdt cap + biz exclusion).
# MAGIC - ledger CAPTURED accounts = 6,029 (the 30-day payment gate — this is the one that
# MAGIC   would drift silently if the pay_dt source substitution changed behavior, so it is
# MAGIC   an explicit STOP, not just the episode count).
# MAGIC - call-day bucket-1 stream = 29,114 (the as-of-call-day position logic).

# COMMAND ----------

_t = spark.sql(f"""
WITH led AS (
  SELECT extnl_acct_id FROM {DB}.uc2_acct_month
  WHERE ym = '202501' AND eom_bucket = 1 AND NOT co_before_2025 AND is_ex_aa
),
ep_exaa AS (
  SELECT e.* FROM {DB}.uc2_episode e
  JOIN (SELECT extnl_acct_id, is_ex_aa FROM {DB}.uc2_acct_month WHERE ym='202501') a
    ON a.extnl_acct_id = e.extnl_acct_id
  WHERE a.is_ex_aa
),
ledger_ep AS (
  SELECT e.extnl_acct_id, e.captured
  FROM ep_exaa e JOIN led l ON l.extnl_acct_id = e.extnl_acct_id
),
caller AS (   -- account-level: captured if any ledger episode captured
  SELECT extnl_acct_id, max(captured) AS any_captured
  FROM ledger_ep GROUP BY extnl_acct_id
)
SELECT
  (SELECT count(*) FROM ledger_ep)                                    AS ledger_episodes,
  (SELECT count_if(any_captured = 1) FROM caller)                     AS captured_accts,
  (SELECT count(*) FROM ep_exaa WHERE callday_bucket = 1
        AND (callday_co_dt IS NULL OR callday_co_dt >= DATE '2025-01-01')) AS callday_b1_stream
""").first()
print(dict(_t.asDict()))
assert _t["ledger_episodes"] == 11262, f"ledger episodes {_t['ledger_episodes']} != 11262 — STOP"
assert _t["captured_accts"]  == 6029,  f"captured accts {_t['captured_accts']} != 6029 — capture gate drifted, STOP"
assert _t["callday_b1_stream"] == 29114, f"call-day b1 stream {_t['callday_b1_stream']} != 29114 — STOP"
print("PASS: uc2_episode ties out (episodes 11,262 / captured 6,029 / call-day b1 29,114)")

# COMMAND ----------

# MAGIC %md
# MAGIC ## Cell 3 — build `uc2_episode_signal` (the transcript read)
# MAGIC
# MAGIC The ONE transcript scan. One row per contactid over the ledger+addressable
# MAGIC episode set, with v2 lexicon flag counts and the resolved language_group
# MAGIC (deceased-first priority CASE, verbatim from m4). No insight query touches the
# MAGIC transcript table again.

# COMMAND ----------

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_episode_signal USING DELTA AS
WITH tx AS (
    SELECT t.contactid,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND t.content rlike 'passed away|death certificate|executor|deceased|calling on behalf') AS deceased_n,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND lower(t.content) rlike 'pay|paid|payment|settle|payment plan|arrangement|work something out') AS pay_n,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND lower(t.content) rlike 'settle|payment plan|arrangement|work something out') AS plan_n,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND lower(t.content) rlike 'hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford') AS hard_n,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND lower(t.content) rlike 'dispute|not my charge|didn.t authorize|did not authorize|unauthorized|fraud|identity theft') AS dispute_n,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND lower(t.content) rlike "i.ll pay|i will pay|going to pay|gonna pay|pay (on|by|this|next)|when i get paid|payday|after my paycheck") AS promise_n,
           count_if(t.participantid = 'CUSTOMER' AND t.content IS NOT NULL
                    AND lower(t.content) rlike 'bank routing|routing number|check number|checkbook|a check for|that check|on the check') AS exec_n
    FROM {TX} t
    JOIN (SELECT DISTINCT contactid FROM {DB}.uc2_episode) d ON t.contactid = d.contactid
    WHERE t.effdt >= '2025-01-01' AND t.effdt < '2025-02-02'
      AND t.content IS NOT NULL
    GROUP BY t.contactid
)
SELECT contactid, deceased_n, pay_n, plan_n, hard_n, dispute_n, promise_n, exec_n,
       CASE
         WHEN coalesce(deceased_n,0) > 0 THEN 'a. deceased or estate'
         WHEN coalesce(promise_n,0)  > 0 THEN 'b. future-dated promise'
         WHEN coalesce(pay_n,0) > 0 AND coalesce(plan_n,0) = 0 THEN 'c. payment talk, no promise'
         WHEN coalesce(plan_n,0)    > 0 THEN 'd. plan or settlement talk'
         WHEN coalesce(hard_n,0)    > 0 THEN 'e. hardship talk'
         WHEN coalesce(dispute_n,0) > 0 THEN 'f. dispute or fraud talk'
         ELSE 'g. no payment-related language'
       END AS language_group
FROM tx
""")
print("built", f"{tcat}.{tsch}.uc2_episode_signal")

# COMMAND ----------

# MAGIC %md ## Cell 3t — tie-out: signal joins to episodes, groups partition
# MAGIC Every signal row joins a distinct episode contactid; language groups are exhaustive.

# COMMAND ----------

_t = spark.sql(f"""
SELECT
  (SELECT count(*) FROM {DB}.uc2_episode_signal)                                     AS signal_rows,
  (SELECT count(DISTINCT contactid) FROM {DB}.uc2_episode_signal)                    AS distinct_cid,
  (SELECT count(*) FROM {DB}.uc2_episode_signal WHERE language_group IS NULL)        AS null_group
""").first()
print(dict(_t.asDict()))
assert _t["signal_rows"] == _t["distinct_cid"], "signal not unique per contactid — STOP"
assert _t["null_group"] == 0, "some episodes fell outside every language group — STOP"
print("PASS: uc2_episode_signal is one row per contactid, groups exhaustive")

# COMMAND ----------

# MAGIC %md
# MAGIC ## Cells 10+ — INSIGHT layer
# MAGIC
# MAGIC Each insight is a short JOIN + GROUP BY on the three base tables. CO 8/10/12
# MAGIC windows are applied HERE (a filter on uc2_acct_month.chrgoff_dt), anchored at
# MAGIC 31 Jan 2025: CO8 = [2025-01-31, 2025-09-30), CO10 = [.., 2025-11-30),
# MAGIC CO12 = [.., 2026-01-31). Reproduce the tier-15 numbers before retiring tier-15.
# MAGIC
# MAGIC Below: cell 15 (cpc distribution / former b20) as the worked example. The rest
# MAGIC (ledger motion, leak list, language, addressable, year funnel, touched-DQ1,
# MAGIC caller gap, priced views) follow the same pattern and get added as each is
# MAGIC ported and tie-checked.

# COMMAND ----------

# MAGIC %md ## Cell 15 — CPC distribution (former b20), full cleaned ledger, no ex-AA filter

# COMMAND ----------

display(spark.sql(f"""
SELECT cpc_class,
       count(*)                              AS accounts,
       round(sum(eom_bal), 0)                AS jan_eom_balance,
       round(sum(cr_lmt_origl_amt), 0)       AS orig_credit_limit_total,
       round(avg(cr_lmt_origl_amt), 0)       AS orig_credit_limit_avg
FROM {DB}.uc2_acct_month
WHERE ym = '202501' AND eom_bucket = 1 AND NOT co_before_2025
GROUP BY cpc_class
ORDER BY accounts DESC
"""))
