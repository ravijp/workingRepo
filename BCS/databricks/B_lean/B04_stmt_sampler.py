# Databricks notebook source
# MAGIC %md
# MAGIC # B04. The statement-frame masked-excerpt sampler (Story B as core)
# MAGIC
# MAGIC =====================================================================
# MAGIC Replaces the calendar-January stratum sampler. Under Story B the sampling
# MAGIC axis is the STATEMENT-TIMING 5-day bucket (days since stmt_dt) x capture
# MAGIC side, on the LOCKED population; the strata are the descriptive B05 pools
# MAGIC (leaked_core, leaked_exec, leaked_promise, captured_contrast, captured_exec,
# MAGIC captured_promise, silent_relaxed, reference) - NOT the old A/B/C/D letters.
# MAGIC It selects a small set of transcripts (~30 per wave), masks them, and
# MAGIC emits a display() grid so a human + the bank-tenant assistant can read
# MAGIC them and score Prompt R (the agent-behavior rubric) BY statement bucket
# MAGIC and leaked-vs-captured. B04 does NOT count populations or size the
# MAGIC opportunity (that is SQL on B05 / the labels table). It stays lean:
# MAGIC build the pool, pick deterministically, mask, export. The locked-value
# MAGIC stop rules live in the sibling B04_checks.py.
# MAGIC
# MAGIC GOVERNANCE (Tier-1, from COMPARISON_REPORT.md):
# MAGIC   * Digit mask ON by default. ONE owner-gated UNMASK edit point (marked
# MAGIC     in the turns_rows CTE). The keeper does not flip it.
# MAGIC   * NO HTML writer. NO write to any personal Workspace path. NO DBFS
# MAGIC     browser-reachable fallback. NO "_unmasked" filename.
# MAGIC   * Output = the masked display() grid + optional CSV to output/csv/ only.
# MAGIC   * Row-level text: this is the package's ONLY row-level-text query; it
# MAGIC     touches the transcript table for the ~30 sampled contactids only.
# MAGIC     NEVER screenshot excerpt content; NEVER import it into the story JSON;
# MAGIC     excerpts travel only inside the batch files.
# MAGIC
# MAGIC FLAGS: payment = captured_sas (locked). spoken promise = transcript
# MAGIC promise_f (locked, the FULL 03n regex). A recorded-PTP flag from
# MAGIC V_COLL_PRMS_DTL_TBL is NOT joined here (unverified); a clean, isolated
# MAGIC seam is left for it (see X2b), to slot in AFTER the owner runs the PTP
# MAGIC probe (ptp-table-probe-spec-2026-07-21.md).
# MAGIC =====================================================================

# COMMAND ----------

# =====================================================================
# SETUP - keep in sync with B00_setup.py (the canonical copy).
# =====================================================================
import datetime as _dt
import time as _time

CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
ANCHOR_YM = "202501"
SAS_CSV_PATH = "/Volumes/cda_model_shared/ecm_cld_model/ecm_cld/collections_zenon/WATERFALL_COLL_CALL_V2_202501.csv"
FMT_CATALOG = "634153504162_glue_connection_catalog"      # fmt_acct_c lives here
CC_CATALOG = "062108867742_glue_connectivity_catalog"     # call + transcript live here

try:
    dbutils.widgets.text("CATALOG", CATALOG);           CATALOG = dbutils.widgets.get("CATALOG")
    dbutils.widgets.text("SCHEMA", SCHEMA);             SCHEMA = dbutils.widgets.get("SCHEMA")
    dbutils.widgets.text("ANCHOR_YM", ANCHOR_YM);       ANCHOR_YM = dbutils.widgets.get("ANCHOR_YM")
    dbutils.widgets.text("SAS_CSV_PATH", SAS_CSV_PATH); SAS_CSV_PATH = dbutils.widgets.get("SAS_CSV_PATH")
    dbutils.widgets.text("FMT_CATALOG", FMT_CATALOG);   FMT_CATALOG = dbutils.widgets.get("FMT_CATALOG")
    dbutils.widgets.text("CC_CATALOG", CC_CATALOG);     CC_CATALOG = dbutils.widgets.get("CC_CATALOG")
except NameError:
    pass  # dbutils absent (plain Spark); the constants above apply

DB = f"{CATALOG}.{SCHEMA}"
FMT = f"`{FMT_CATALOG}`.fmt_acct_dba.fmt_acct_c"
CALL = f"`{CC_CATALOG}`.contactcenter_bdp_db.`call`"
TX = f"`{CC_CATALOG}`.contactcenter_bdp_db.transcript"

# --- derived windows (all from ANCHOR_YM; 202501 literals in comments) ---
_a0 = _dt.date(int(ANCHOR_YM[:4]), int(ANCHOR_YM[4:6]), 1)
_mm = lambda d, k: _dt.date(d.year + (d.month - 1 + k) // 12, (d.month - 1 + k) % 12 + 1, 1)

# the transcript scan is bounded to the statement-frame span (Dec24-Mar25) and
# KEEPS the EFFDT_HARD_END live-loading-edge guard (2026-07-10). The episode
# window is statement-anchored upstream; here we only need a bounded scan that
# covers the sampled contactids' transcripts.
EFFDT_SCAN_START = _mm(_a0, -1).isoformat()                              # 2024-12-01
EFFDT_SCAN_END = _mm(_a0, 3).isoformat()                                # 2025-04-01
EFFDT_HARD_END = "2026-07-10"   # not vintage-derived: the live-loading-edge guard

NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"   # THE numeric key rule

# --- wave + quota widgets (waves re-shape without a code edit) ------------
# WAVE selects a disjoint pick_rn window per stratum: [(WAVE-1)*Q+1, WAVE*Q].
# QUOTAS is the per-stratum wave-1 shape (design 4.4, owner sets the final
# dict). Strata omitted from QUOTAS drop out of the wave with no code change.
WAVE = 1
QUOTAS = {
    "leaked_core": 12,        # PRIMARY - the Story-B read, bucket-balanced
    "leaked_exec": 4,         # small pool; sample near-exhaustively over waves
    "leaked_promise": 4,      # dedicated stratum
    "captured_contrast": 8,   # the falsifier baseline, mirrors leaked_core buckets
    # captured_exec / captured_promise / silent_relaxed / reference: held to
    # later waves (add a key here to pull them into a wave; design 4.4/8).
}
try:
    dbutils.widgets.text("WAVE", str(WAVE)); WAVE = int(dbutils.widgets.get("WAVE"))
except NameError:
    pass
QUOTA_TOTAL = sum(QUOTAS.values())

# --- anchor-excerpt drift detector (design 4.6) ---------------------------
# From wave 2 on, 2 anchor excerpts are repeated from wave 1 as a cross-session
# consistency check. Pin their contactids here (comma-separated widget). Empty
# on wave 1. They are emitted at the top of the export with is_anchor = 1.
ANCHOR_CONTACTIDS = ""   # e.g. "abc123,def456"; leave empty for wave 1
try:
    dbutils.widgets.text("ANCHOR_CONTACTIDS", ANCHOR_CONTACTIDS)
    ANCHOR_CONTACTIDS = dbutils.widgets.get("ANCHOR_CONTACTIDS")
except NameError:
    pass
_ANCHORS = [a.strip() for a in ANCHOR_CONTACTIDS.split(",") if a.strip()]

# --- output plumbing (lean provenance; NO locked-value asserts here) -------
RESULTS = []          # (name, actual, expected, status)
GRIDS = {}            # name -> (columns, rows); reprinted in the record block
_T0 = _time.time()
_TSEC = [_time.time()]


def sec(title):
    now = _time.time()
    print(f"\n===== {title} =====  (prev section {int(now - _TSEC[0])}s, elapsed {int(now - _T0)}s)")
    _TSEC[0] = now


def fmt(v):
    if v is None:
        return "NULL"
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return f"{v:,}"
    if isinstance(v, float):
        return f"{v:,.0f}"
    return str(v)


def grid(name, df):
    """Lossless, transcription-friendly grid. The SQL behind df carries ORDER BY."""
    rows = df.collect()
    cols = df.columns
    print(f"{name}: {len(rows)} rows, all shown")
    header = "| idx | " + " | ".join(cols) + " |"
    rule = "|" + "---|" * (len(cols) + 1)
    for i, r in enumerate(rows):
        if i % 40 == 0:
            print(f"-- {name}: rows {i + 1}..{min(i + 40, len(rows))} of {len(rows)} --")
            print(header)
            print(rule)
        print("| " + str(i + 1) + " | " + " | ".join(fmt(r[c]) for c in cols) + " |")
    GRIDS[name] = (cols, rows)


def record_block(notebook):
    """Lean record block: what was sampled. No locked-value asserts."""
    print("=" * 78)
    print(f"RECORD BLOCK  {notebook}  vintage {ANCHOR_YM}  wave {WAVE}  platform Databricks")
    print("=" * 78)
    for name, actual, expected, status in RESULTS:
        tail = "" if expected is None else f"  (expected {fmt(expected)})"
        print(f"{status:8}  {name} = {fmt(actual)}{tail}")
    for gname, (cols, rows) in GRIDS.items():
        print(f"\n### {gname} ({len(rows)} rows)")
        print("| " + " | ".join(map(str, cols)) + " |")
        print("|" + "---|" * len(cols))
        for r in rows:
            vals = r if isinstance(r, tuple) else [r[c] for c in cols]
            print("| " + " | ".join(fmt(v) for v in vals) + " |")
    print("=" * 78)


def flush_metrics(notebook):
    """On-platform provenance: the sitting's measured values survive across
    sittings in a small Delta table."""
    stage = [(notebook, n,
              float(a) if isinstance(a, (int, float)) and not isinstance(a, bool) else None,
              fmt(a),
              float(e) if isinstance(e, (int, float)) and not isinstance(e, bool) else None,
              s, ANCHOR_YM) for n, a, e, s in RESULTS]
    spark.createDataFrame(
        stage,
        "notebook string, name string, value double, value_str string, expected double, status string, vintage string"
    ).createOrReplaceTempView("_metrics_stage")
    spark.sql(f"""CREATE TABLE IF NOT EXISTS {DB}.uc2_run_metrics
                  (run_ts timestamp, notebook string, name string, value double,
                   value_str string, expected double, status string, vintage string)""")
    spark.sql(f"INSERT INTO {DB}.uc2_run_metrics SELECT current_timestamp(), * FROM _metrics_stage")
    print(f"metrics appended to {DB}.uc2_run_metrics ({len(stage)} rows)")


print(f"SETUP OK: vintage {ANCHOR_YM}; wave {WAVE}; layers -> {DB}")
print(f"  quotas = {QUOTAS} (total {QUOTA_TOTAL}); anchors = {_ANCHORS}")
# =====================================================================
# end of SETUP
# =====================================================================

# COMMAND ----------

# MAGIC %md
# MAGIC ## X1. Preconditions

# COMMAND ----------

sec("X1 preconditions")

OUT = f"{DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}"
SIG = f"{DB}.uc2_t16_03n_signals"
assert spark.catalog.tableExists(OUT), f"PRECONDITION MISS: {OUT} missing - run B02b first"
print(f"PASS  table exists: {OUT}")
assert spark.catalog.tableExists(SIG), f"PRECONDITION MISS: {SIG} missing - run B02 first"
print(f"PASS  table exists: {SIG}")
assert spark.catalog.tableExists(f"{CC_CATALOG}.contactcenter_bdp_db.transcript"), \
    "PRECONDITION MISS: transcript table not reachable"
print("PASS  transcript table reachable")
# the re-anchor columns must be present, else this is the OLD 04s
_cols = [c.lower() for c in spark.sql(f"SELECT * FROM {OUT} LIMIT 0").columns]
for _c in ["stmt_dt", "days_since_stmt_dt", "stmt_5day_bucket", "stmt_5day_bucket_start"]:
    assert _c in _cols, \
        f"PRECONDITION MISS: {OUT} has no '{_c}' column - re-run the re-anchored B02/B02b"
print("PASS  statement-frame columns present on the 04s table")

# COMMAND ----------

# MAGIC %md
# MAGIC ## X2. Build the statement-window pool on the LOCKED gate
# MAGIC
# MAGIC No dlnqt_cd_m2 filter on the whole pool; the locked in_sas_ledger /
# MAGIC captured_sas / leaked_sas / w_s_flag gate IS the population and is read
# MAGIC as-is (never recomputed). The 04s table is already re-anchored (in-window
# MAGIC episodes only). promise_f lives on 03n (contactid grain) - joined in as
# MAGIC the LOCKED spoken-promise flag; a missing 03n row means all flags 0.

# COMMAND ----------

sec("X2 pool build")

# transcript-exists semi-join, bounded (design 7.2 X4). One pass over the
# 04s contactids; HARD_END guard retained.
spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW cb_tx_exists AS
    SELECT DISTINCT t.contactid
    FROM {TX} t
    JOIN (SELECT DISTINCT contactid FROM {OUT}) c ON c.contactid = t.contactid
    WHERE t.content IS NOT NULL
      AND t.effdt >= '{EFFDT_SCAN_START}' AND t.effdt < '{EFFDT_SCAN_END}'
      AND t.effdt < '{EFFDT_HARD_END}'
""")

# the pool with the LOCKED gate columns read as-is + promise_f from 03n + the
# statement bucket, and the descriptive B05-pool stratum. Sub-strata (exec /
# promise) are tested BEFORE their core so the specific pool wins; leaked
# before captured is moot (the gate is disjoint).
spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW cb_pool AS
    SELECT o.acct_key, o.contactid, o.call_dt,
           o.stmt_dt, o.days_since_stmt_dt,
           o.stmt_5day_bucket, o.stmt_5day_bucket_start,
           o.pre_due_f, o.post_due_f,
           o.language_group, o.pay_f, o.exec_f, o.deceased_acct,
           o.captured_sas, o.leaked_sas, o.w_s_flag,
           o.eop_bal_m1, o.gross_loss_12m_amt,
           o.aws_caller_class,
           coalesce(x.promise_f, 0) AS promise_f,   -- LOCKED spoken-promise flag (03n full regex)
           CASE WHEN w.contactid IS NOT NULL THEN 1 ELSE 0 END AS has_tx,
           CASE
             WHEN o.leaked_sas AND o.pay_f > 0 AND o.deceased_acct = 0
                  AND o.exec_f = 1 AND w.contactid IS NOT NULL
               THEN 'leaked_exec'
             WHEN o.leaked_sas AND o.pay_f > 0 AND o.deceased_acct = 0
                  AND o.language_group = 'b. future-dated promise' AND w.contactid IS NOT NULL
               THEN 'leaked_promise'
             WHEN o.leaked_sas AND o.pay_f > 0 AND o.deceased_acct = 0
                  AND w.contactid IS NOT NULL
               THEN 'leaked_core'
             WHEN o.captured_sas AND o.pay_f > 0 AND o.exec_f = 1 AND w.contactid IS NOT NULL
               THEN 'captured_exec'
             WHEN o.captured_sas AND o.pay_f > 0
                  AND o.language_group = 'b. future-dated promise' AND w.contactid IS NOT NULL
               THEN 'captured_promise'
             WHEN o.captured_sas AND o.pay_f > 0 AND w.contactid IS NOT NULL
               THEN 'captured_contrast'
             WHEN NOT o.captured_sas
                  AND o.language_group = 'g. no payment-related language' AND w.contactid IS NOT NULL
               THEN 'silent_relaxed'
             WHEN w.contactid IS NOT NULL
               THEN 'reference'
           END AS stratum
    FROM {OUT} o
    LEFT JOIN {SIG} x ON x.contactid = o.contactid
    LEFT JOIN cb_tx_exists w ON w.contactid = o.contactid
""")

grid("stratum pool sizes (episodes; sanity, not a record number)", spark.sql("""
    SELECT stratum, count(*) AS episodes, count(DISTINCT acct_key) AS accounts
    FROM cb_pool
    WHERE stratum IS NOT NULL
    GROUP BY 1 ORDER BY 1
"""))

# COMMAND ----------

# MAGIC %md
# MAGIC ## X2b. RECORDED-PTP SEAM (isolated; NOT wired yet)
# MAGIC
# MAGIC The spoken promise above (promise_f) is "the customer SPOKE a promise",
# MAGIC NOT "a promise was RECORDED". A recorded-PTP flag from the system table
# MAGIC V_COLL_PRMS_DTL_TBL (schema SRC_COLL_DBA) would give the missing middle
# MAGIC of the chain intent(spoken) -> PTP(recorded) -> payment(kept). That table
# MAGIC is UNVERIFIED for our 202501 window, so it is NOT joined here. This is the
# MAGIC single, clearly-marked place it will slot in AFTER the owner runs the
# MAGIC probe in ptp-table-probe-spec-2026-07-21.md.
# MAGIC
# MAGIC To wire it (owner's call, once P1/P4/P5 of the probe are healthy):
# MAGIC   1. resolve the catalog path (mirror the fmt/call convention);
# MAGIC   2. build a bounded ptp_flag view keyed by the numeric key rule
# MAGIC      (cast(try_cast(<idcol> AS bigint) AS string)), promise-made date in
# MAGIC      [2024-12-01, 2025-04-01);
# MAGIC   3. LEFT JOIN it below and set recorded_ptp_f from it.
# MAGIC Until then recorded_ptp_f is a NULL placeholder so the export schema is
# MAGIC stable and B06 can see the column from wave 1 on.

# COMMAND ----------

# recorded-PTP seam: the single placeholder column. NULL until the owner wires
# the V_COLL_PRMS_DTL_TBL join above (do NOT join that table now - unverified).
# See ptp-table-probe-spec-2026-07-21.md.
spark.sql("""
    CREATE OR REPLACE TEMP VIEW cb_pool_ptp AS
    SELECT p.*,
           CAST(NULL AS int) AS recorded_ptp_f   -- SEAM: recorded PTP from V_COLL_PRMS_DTL_TBL, wired post-probe
    FROM cb_pool p
""")

# COMMAND ----------

# MAGIC %md
# MAGIC ## X3. Deterministic bucket-balanced pick + wave window + batch
# MAGIC
# MAGIC Two-stage row_number (design 2.3): within (stratum, bucket) by
# MAGIC xxhash64(contactid), then round-robin across buckets before the quota
# MAGIC cap. Waves are disjoint pick_rn windows [(WAVE-1)*Q+1, WAVE*Q] per
# MAGIC stratum. Batch round-robin 1..3. The "outside 0-55 days" bucket is
# MAGIC excluded from sampling (no statement-timing signal). Anchors (wave > 1)
# MAGIC are unioned back at is_anchor = 1 with pick_rn = 0 so they always render.

# COMMAND ----------

sec("X3 pick and batch")

# per-stratum quota, driven from the QUOTAS dict (waves re-shape without edits)
_quota_when = "\n".join(
    f"                WHEN stratum = '{_s}' THEN {_q}" for _s, _q in QUOTAS.items()
) or "                WHEN 1 = 0 THEN 0"
_wave_lo = f"(({WAVE} - 1) * quota + 1)"
_wave_hi = f"({WAVE} * quota)"
_anchor_in = ("'" + "','".join(_ANCHORS) + "'") if _ANCHORS else "''"

spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW cb_sampled AS
    WITH sampleable AS (
        SELECT *,
               CASE
{_quota_when}
                 ELSE 0
               END AS quota
        FROM cb_pool_ptp
        WHERE stratum IS NOT NULL
          AND contactid IS NOT NULL
          AND stmt_5day_bucket <> 'outside 0-55 days'
    ),
    quota_pool AS (SELECT * FROM sampleable WHERE quota > 0),
    bucket_ranked AS (
        SELECT *,
               row_number() OVER (PARTITION BY stratum, stmt_5day_bucket
                                  ORDER BY xxhash64(contactid)) AS bucket_pick_rn
        FROM quota_pool
    ),
    ranked AS (
        SELECT *,
               row_number() OVER (PARTITION BY stratum
                                  ORDER BY bucket_pick_rn, stmt_5day_bucket_start,
                                           xxhash64(contactid)) AS pick_rn
        FROM bucket_ranked
    ),
    waved AS (
        SELECT * FROM ranked
        WHERE pick_rn >= {_wave_lo} AND pick_rn <= {_wave_hi}
    ),
    anchors AS (
        -- wave > 1 drift detector: the pinned contactids, is_anchor = 1,
        -- stratum + bucket preserved; pick_rn 0 keeps them at the top
        SELECT *, 0 AS pick_rn, 1 AS is_anchor
        FROM (SELECT *, 0 AS bucket_pick_rn FROM sampleable
              WHERE contactid IN ({_anchor_in}))
    ),
    picked AS (
        SELECT acct_key, contactid, call_dt, stmt_dt, days_since_stmt_dt,
               stmt_5day_bucket, stmt_5day_bucket_start, pre_due_f, post_due_f,
               language_group, pay_f, exec_f, promise_f, recorded_ptp_f,
               captured_sas, leaked_sas, w_s_flag, aws_caller_class,
               stratum, pick_rn, 0 AS is_anchor
        FROM waved
        UNION ALL
        SELECT acct_key, contactid, call_dt, stmt_dt, days_since_stmt_dt,
               stmt_5day_bucket, stmt_5day_bucket_start, pre_due_f, post_due_f,
               language_group, pay_f, exec_f, promise_f, recorded_ptp_f,
               captured_sas, leaked_sas, w_s_flag, aws_caller_class,
               stratum, pick_rn, is_anchor
        FROM anchors
    )
    SELECT *,
           ((row_number() OVER (ORDER BY is_anchor DESC, stratum, pick_rn) - 1) % 3) + 1 AS batch_nbr
    FROM picked
""")

_pick_df = spark.sql("""
    SELECT stratum, count(*) AS picked, count_if(is_anchor = 1) AS anchors
    FROM cb_sampled GROUP BY 1 ORDER BY 1
""")
grid("picked per stratum", _pick_df)

# cheap guards only (protect the export itself; the locked ties are in the sibling)
for _row in _pick_df.collect():
    _q = QUOTAS.get(_row["stratum"], 0)
    _non_anchor = _row["picked"] - _row["anchors"]
    assert _non_anchor <= _q, \
        f"quota exceeded for {_row['stratum']}: {_non_anchor} > {_q}"
_r = spark.sql("SELECT count(*) AS n, count_if(is_anchor = 1) AS a FROM cb_sampled").first()
assert (_r["n"] - _r["a"]) <= QUOTA_TOTAL, \
    f"sampled {_r['n'] - _r['a']} non-anchor exceeds quota total {QUOTA_TOTAL}"
RESULTS.append(("sampled excerpts (incl anchors)", _r["n"], None, "MEASURED"))
print(f"MEASURED  sampled excerpts (incl anchors) = {fmt(_r['n'])} (anchors {fmt(_r['a'])})")

# the bucket spread of the sample (the point of the bucket-balanced pick)
grid("sample spread by stratum x statement bucket", spark.sql("""
    SELECT stratum, stmt_5day_bucket, count(*) AS picked
    FROM cb_sampled
    GROUP BY 1, 2 ORDER BY 1, stmt_5day_bucket
"""))

# COMMAND ----------

# MAGIC %md
# MAGIC ## X4. Masked export (the ONLY row-level-text query)
# MAGIC
# MAGIC RESTORED turn ordering: turns render in beginmillis order via
# MAGIC array_join(transform(array_sort(collect_list(struct(beginmillis, line))),
# MAGIC x -> x.line), char(10)). Digit mask ON (the single UNMASK edit point).
# MAGIC 8000-char cap; text_state = PARTIAL when the cap clips a call. The
# MAGIC excerpt id prefixes off the stratum name (never a bare letter that can
# MAGIC change meaning): LC/LE/LP/KC/KE/KP/GR/RF.

# COMMAND ----------

sec("X4 masked export")

export_df = spark.sql(f"""
    WITH turns_rows AS (
        SELECT t.contactid, t.beginmillis,
               concat(t.participantid, ': ',
                      -- ============================================================
                      -- REDACTION (digit masking) - ON by default.
                      -- What it does: regexp_replace collapses every run of 3+
                      -- consecutive digits in the transcript text to '###'. This
                      -- redacts account numbers, card numbers, SSNs, phone numbers,
                      -- and dollar amounts before any excerpt leaves in a grid/CSV.
                      -- Words are untouched, so phrase/intent discovery is unaffected.
                      -- Why: excerpts are the only row-level text this pipeline
                      -- emits; masking removes PII spill from clipboard/CSV files.
                      -- THE ONLY UNMASK POINT (owner-gated): to export UNMASKED
                      -- text, replace the regexp_replace(...) line below with just
                      --     t.content
                      -- and NOTHING else. Do this ONLY on the owner's explicit
                      -- instruction (it is a governance decision, not a code tidy).
                      -- Leave masking ON unless told otherwise.
                      -- ============================================================
                      regexp_replace(t.content, '[0-9]{{3,}}', '###')
               ) AS line
        FROM {TX} t
        JOIN (SELECT DISTINCT contactid FROM cb_sampled) s ON s.contactid = t.contactid
        WHERE t.content IS NOT NULL
          AND t.effdt >= '{EFFDT_SCAN_START}' AND t.effdt < '{EFFDT_SCAN_END}'
          AND t.effdt < '{EFFDT_HARD_END}'
    ),
    turns AS (
        SELECT contactid,
               array_join(transform(array_sort(collect_list(struct(beginmillis, line))),
                                    x -> x.line), char(10)) AS convo
        FROM turns_rows
        GROUP BY 1
    ),
    prefix AS (
        SELECT 'leaked_core' AS stratum, 'LC' AS px UNION ALL
        SELECT 'leaked_exec', 'LE' UNION ALL
        SELECT 'leaked_promise', 'LP' UNION ALL
        SELECT 'captured_contrast', 'KC' UNION ALL
        SELECT 'captured_exec', 'KE' UNION ALL
        SELECT 'captured_promise', 'KP' UNION ALL
        SELECT 'silent_relaxed', 'GR' UNION ALL
        SELECT 'reference', 'RF'
    )
    SELECT s.batch_nbr AS cb_batch,
           concat(coalesce(px.px, 'XX'),
                  CASE WHEN s.is_anchor = 1 THEN 'A' ELSE '' END,
                  cast(s.pick_rn AS string)) AS cb_excerpt_id,
           s.stratum AS cb_stratum,
           s.is_anchor AS cb_is_anchor,
           s.stmt_5day_bucket AS cb_stmt_5day_bucket,
           s.days_since_stmt_dt AS cb_days_since_stmt_dt,
           CASE WHEN s.pre_due_f = 1 THEN 'pre-due'
                WHEN s.post_due_f = 1 THEN 'post-due' ELSE 'outside' END AS cb_due_side,
           CASE WHEN s.leaked_sas THEN 'leaked' WHEN s.captured_sas THEN 'captured'
                ELSE 'other' END AS cb_outcome,
           s.contactid AS cb_contactid,
           s.call_dt AS cb_call_dt,
           s.stmt_dt AS cb_stmt_dt,
           s.language_group AS cb_language_group,
           s.promise_f AS cb_spoken_promise_f,       -- LOCKED spoken-promise flag
           s.recorded_ptp_f AS cb_recorded_ptp_f,    -- SEAM placeholder (NULL until PTP probe wires it)
           s.aws_caller_class AS cb_aws_class_label, -- carried label only
           CASE WHEN length(v.convo) > 8000 THEN 'PARTIAL' ELSE 'full' END AS text_state,
           substr(v.convo, 1, 8000) AS cb_transcript_masked
    FROM cb_sampled s
    LEFT JOIN turns v ON v.contactid = s.contactid
    LEFT JOIN prefix px ON px.stratum = s.stratum
    ORDER BY s.is_anchor DESC, s.batch_nbr, s.stratum,
             s.stmt_5day_bucket_start, s.pick_rn
    LIMIT 100
""")
_n = export_df.count()
RESULTS.append(("export rows", _n, None, "MEASURED"))
print(f"MEASURED  export rows = {fmt(_n)}")
print("GOVERNANCE: masked grid only. Download the grid below as CSV into")
print("output/csv/ if you need a file - NO other write. DO NOT screenshot the")
print("excerpt content; excerpts travel only inside the batch files.")
display(export_df)

# COMMAND ----------

# MAGIC %md
# MAGIC ## X5. Record block + on-platform metrics (lean provenance)

# COMMAND ----------

sec("X5 record")

record_block("B04_stmt_sampler")
flush_metrics("B04_stmt_sampler")
print("B04_stmt_sampler complete. Run B04_checks.py ONCE to certify the B05 pool")
print("ties before the wave ships.")
