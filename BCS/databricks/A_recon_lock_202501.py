# Databricks notebook source
# MAGIC %md
# MAGIC # A. Reconciliation lock, vintage 202501 (FROZEN after the first verified run)
# MAGIC
# MAGIC One-time reconciliation notebook for the 202501 vintage. It consolidates the
# MAGIC round-11 caller-gap evidence (phases 1-1d) into permanent raising asserts,
# MAGIC replicates the export waterfall from the CSV, probes the PAYMT_AMT sign
# MAGIC convention, and builds the one-time capture-gate delta table. Constants,
# MAGIC not widgets. Runs alone, pasted as one cell or imported as a notebook.
# MAGIC
# MAGIC FROZEN flag: the first verified run happens with FROZEN = False (measured
# MAGIC cells print instead of asserting). After the run is verified against the
# MAGIC screenshots, the None entries in EXPECTED are filled with the measured
# MAGIC values, FROZEN flips to True, and this notebook becomes a pure regression
# MAGIC test. Never edited after that except as dated corrections.
# MAGIC
# MAGIC GO-FORWARD RULE: this notebook is the ONLY place the export CSV's
# MAGIC call_type_INB column is read. The go-forward pipeline (the B files)
# MAGIC rebuilds the caller flag natively from the call table with the numeric
# MAGIC key; the CSV flag is reconciliation evidence here, nothing else.
# MAGIC
# MAGIC CALLER CONSTRUCTS, never merged (each keeps its definition sentence):
# MAGIC our string-keyed ledger callers 9,389 (historical); the numeric-keyed
# MAGIC ledger callers (measured in the B pipeline); the CSV INBOUND flag 11,136
# MAGIC replicated / 11,154 recorded; the call-day bucket-1 stream 29,114
# MAGIC (re-measures after the key fix); the statement-window callers 19,789
# MAGIC (recorded, not used here).
# MAGIC
# MAGIC Platform: Databricks. Counts are accounts unless a cell says rows or
# MAGIC episodes. Windows are stated per cell.

# COMMAND ----------

# ============================ A1. constants ============================
CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
DB = f"{CATALOG}.{SCHEMA}"
CSV_PATH = "/Volumes/cda_model_shared/ecm_cld_model/ecm_cld/collections_zenon/WATERFALL_COLL_CALL_V2_202501.csv"
CALL_TABLE = "062108867742_glue_connectivity_catalog.contactcenter_bdp_db.`call`"
ANCHOR_YM = "202501"          # constant here by design; the B files parameterize

FROZEN = False   # flip to True at freeze, after EXPECTED is filled and verified

# EXPECTED: every pre-registered value. None = measure mode (fill at freeze).
EXPECTED = {
    # round-10 build preconditions (this notebook must read the verified build)
    "r10 ledger all": 204323,
    "r10 ledger exaa": 189146,
    "r10 touched b1": 724848,
    "r10 ledger callers (string key)": 9389,
    "r10 ledger episodes (string key)": 11262,
    # the export CSV contract
    "csv rows": 610183,
    "csv distinct accounts": 610183,
    "csv id null": 0,
    "csv id non-castable": 0,
    "csv id pad-mismatch": 0,
    # the SAS waterfall replication (round-10 verified, from this same file)
    "wf 01 total": 610183,
    "wf 02 dq1": 202479,
    "wf 03 +cpc": 186848,
    "wf 04 sas ledger": 186013,
    "wf inb 01 total": 34234,
    "wf inb 02 dq1": 12615,
    "wf inb 03 +cpc": 11289,
    "wf inb 04 flagged": 11136,
    # the round-11 gap decomposition (permanent evidence)
    "gap: sas flagged": 11136,
    "gap: our callers": 9389,
    "gap: shared": 9194,
    "gap: flagged only": 1942,
    "gap: ours only": 195,
    # the phase-1d numeric-rejoin evidence (immutable under the effdt bound)
    "1d: numeric coverage of the 1,942": 1942,
    "1d: jan inbound rows recovered": 2220,
    "1d: jan inbound accounts recovered": 1942,
    "1d: jan key mismatches (id-carrying inbound)": 75883,
    "1d: jan inbound rows null acctid": 312795,
    "1d: jan acctid digits-only rows": 1267227,
    "1d: jan acctid null rows": 481838,
    "1d: jan acctid other-shape rows": 0,
    # measured at first run, locked at freeze
    "probe: aws-captured accounts joined to export": None,
    "probe: of those, negative PAYMT M1 or M2": None,
    "delta grid": None,   # dict {"construct|aws|sas": accounts} at freeze
    "csv column count": None,
}

# COMMAND ----------

# ==================== A2. helpers and output plumbing ====================
# Analysis helpers: chk() and NUM_KEY only.
# Output plumbing per the output-ergonomics instruction (2026-07-16):
# lossless grids, transposed wide pulls, record block, on-platform metrics,
# failure context, section timing. Results cross as screenshots; output must
# be lossless and transcription-friendly.
import time as _time

NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"   # THE numeric key rule

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
        return f"{v:,.0f}"          # thousands separators, no scientific notation
    return str(v)


def chk(name, actual, expected, tol=0, ctx=None):
    """Raising anchor check; a miss STOPS the run. expected=None -> measure mode.
    ctx: optional DataFrame printed BEFORE raising, so the failure screenshot
    carries its own diagnosis."""
    if expected is None:
        RESULTS.append((name, actual, None, "MEASURED"))
        print(f"MEASURED  {name} = {fmt(actual)}")
        return
    ok = (abs(actual - expected) <= tol) if tol else (actual == expected)
    if not ok:
        if ctx is not None:
            print(f"CONTEXT for the failing check '{name}':")
            ctx.show(500, truncate=False)
        raise AssertionError(f"ANCHOR MISS {name}: got {fmt(actual)}, expected {fmt(expected)}"
                             + (f" (tol {tol})" if tol else ""))
    RESULTS.append((name, actual, expected, "PASS"))
    print(f"PASS  {name} = {fmt(actual)}")


def grid(name, df):
    """Lossless, transcription-friendly grid: markdown text, running index,
    40-row chunk banners. The SQL behind df must carry its own ORDER BY."""
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


def kv(name, df):
    """Transpose a wide one-row pull: one metric per line, never a wide table."""
    r = df.first()
    print(f"{name}: (one row, transposed)")
    for c in df.columns:
        print(f"  {c} = {fmt(r[c])}")
    GRIDS[name] = (["metric", "value"], [(c, r[c]) for c in df.columns])


def record_block(notebook):
    """The block that gets screenshotted and transcribed verbatim."""
    print("=" * 78)
    print(f"RECORD BLOCK  {notebook}  vintage {ANCHOR_YM}  platform Databricks")
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
    """On-platform provenance: every chk and measured value survives across
    sittings in a small Delta table. Screenshots stay the transfer channel."""
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


print(f"A_recon_lock_202501: target {DB}; FROZEN = {FROZEN}")

# COMMAND ----------

# MAGIC %md
# MAGIC ## A3. Preconditions: the round-10 tables exist and are the verified builds

# COMMAND ----------

sec("A3 preconditions")

for t in ["uc2_t16_01_populations", "uc2_t16_02_episodes", "uc2_t16_04_outcomes"]:
    assert spark.catalog.tableExists(f"{DB}.{t}"), f"PRECONDITION MISS: {DB}.{t} does not exist"
    print(f"PASS  table exists: {DB}.{t}")

_r = spark.sql(f"""
    SELECT count_if(in_ledger_all)  AS ledger_all,
           count_if(in_ledger_exaa) AS ledger_exaa,
           count_if(touched_b1)     AS touched_b1
    FROM {DB}.uc2_t16_01_populations
""").first()
chk("r10 ledger all", _r["ledger_all"], EXPECTED["r10 ledger all"])
chk("r10 ledger exaa", _r["ledger_exaa"], EXPECTED["r10 ledger exaa"])
chk("r10 touched b1", _r["touched_b1"], EXPECTED["r10 touched b1"])

_r = spark.sql(f"""
    SELECT count(*) AS episodes, count(DISTINCT c.acct_key) AS callers
    FROM {DB}.uc2_t16_02_episodes c
    JOIN {DB}.uc2_t16_01_populations p ON p.acct_key = c.acct_key
    WHERE c.is_episode_std = 1 AND p.in_ledger_exaa
""").first()
chk("r10 ledger episodes (string key)", _r["episodes"], EXPECTED["r10 ledger episodes (string key)"])
chk("r10 ledger callers (string key)", _r["callers"], EXPECTED["r10 ledger callers (string key)"])

# COMMAND ----------

# MAGIC %md
# MAGIC ## A4. The CSV load: all-string schema, FAILFAST, grain proof
# MAGIC
# MAGIC First run: header read, every column a string (inferSchema off), FAILFAST
# MAGIC on malformed rows. At freeze, the explicit all-string StructType hard-coded
# MAGIC from the A5 census replaces the header read (the locked explicit-schema
# MAGIC convention, without fabricating a header we have not seen).

# COMMAND ----------

sec("A4 CSV load")

csv_df = (spark.read.format("csv")
          .option("header", True)
          .option("inferSchema", False)     # every column arrives as string
          .option("mode", "FAILFAST")       # malformed row = STOP
          .load(CSV_PATH))
csv_df.createOrReplaceTempView("_sas_csv")

_r = spark.sql("SELECT count(*) AS rows, count(DISTINCT EXTNL_ACCT_ID) AS accts FROM _sas_csv").first()
chk("csv rows", _r["rows"], EXPECTED["csv rows"])
chk("csv distinct accounts", _r["accts"], EXPECTED["csv distinct accounts"])
# rows == distinct accounts proves the grain (one row per account, no null ids)

spark.sql(f"CREATE OR REPLACE TABLE {DB}.uc2_sas_raw_202501 AS SELECT * FROM _sas_csv")
print(f"built {DB}.uc2_sas_raw_202501 (all-string mirror of the CSV)")

# COMMAND ----------

# MAGIC %md
# MAGIC ## A5. Schema census (evidence cell)
# MAGIC
# MAGIC The printed list is the authority for the B01 typed-column extents. At
# MAGIC freeze it locks into EXPECTED["csv column count"] and the B files' pinned
# MAGIC explicit schema.

# COMMAND ----------

sec("A5 schema census")

_cols = csv_df.columns
print(f"CSV columns: {len(_cols)} (one per line, transcribe in order)")
for i, c in enumerate(_cols, 1):
    print(f"  {i:3}  {c}")
chk("csv column count", len(_cols), EXPECTED["csv column count"])

# the columns THIS notebook uses must exist now (B01's fuller list pins later)
_required = ["EXTNL_ACCT_ID", "DLNQT_CD_M1", "CPC_FLAG_NW", "CHRGOFF_RSN_M1",
             "call_type_INB", "PAYMT_AMT_M1", "PAYMT_AMT_M2"]
_missing = [c for c in _required if c.upper() not in {x.upper() for x in _cols}]
assert not _missing, f"REQUIRED COLUMNS MISSING from the CSV: {_missing} - STOP"
print(f"PASS  required columns present: {_required}")

# COMMAND ----------

# MAGIC %md
# MAGIC ## A6. Export key probe (before any key is used)
# MAGIC
# MAGIC The numeric key rule is safe on a source only if every non-null id casts
# MAGIC to bigint and no id differs from its numeric rendering (no padding).

# COMMAND ----------

sec("A6 export key probe")

_r = spark.sql(f"""
    SELECT count_if(EXTNL_ACCT_ID IS NULL) AS id_null,
           count_if(EXTNL_ACCT_ID IS NOT NULL AND try_cast(EXTNL_ACCT_ID AS bigint) IS NULL) AS non_castable,
           count_if(EXTNL_ACCT_ID IS NOT NULL AND try_cast(EXTNL_ACCT_ID AS bigint) IS NOT NULL
                    AND trim(cast(EXTNL_ACCT_ID AS string)) <> {NUM_KEY.format(c="EXTNL_ACCT_ID")}) AS pad_mismatch
    FROM _sas_csv
""").first()
chk("csv id null", _r["id_null"], EXPECTED["csv id null"])
chk("csv id non-castable", _r["non_castable"], EXPECTED["csv id non-castable"])
chk("csv id pad-mismatch", _r["pad_mismatch"], EXPECTED["csv id pad-mismatch"])

# COMMAND ----------

# MAGIC %md
# MAGIC ## A7. The enriched build: waterfall flags as columns
# MAGIC
# MAGIC Filter spellings verbatim from the round-10-verified replication
# MAGIC (upper/trim; OTHERS and BLANK variants included).

# COMMAND ----------

sec("A7 enriched build")

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_sas_wf_202501 AS
WITH f AS (
    SELECT r.*,
           try_cast(EXTNL_ACCT_ID AS bigint) AS acct_num,
           {NUM_KEY.format(c="EXTNL_ACCT_ID")} AS acct_key,
           coalesce(try_cast(DLNQT_CD_M1 AS int) = 1, false) AS wf_dq1,
           coalesce(upper(trim(CPC_FLAG_NW)) IN ('OTHER', 'OTHERS', 'COBRAND', 'PLCC'), false) AS wf_cpc,
           (CHRGOFF_RSN_M1 IS NULL OR trim(CHRGOFF_RSN_M1) = ''
            OR upper(trim(CHRGOFF_RSN_M1)) IN ('PLY', 'BLANK')) AS wf_non_co,
           coalesce(upper(trim(call_type_INB)) LIKE '%INB%', false) AS csv_inb
    FROM _sas_csv r
)
SELECT *, (wf_dq1 AND wf_cpc AND wf_non_co) AS in_sas_ledger
FROM f
""")
print(f"built {DB}.uc2_sas_wf_202501")

# COMMAND ----------

# MAGIC %md
# MAGIC ## A8. Waterfall asserts (the eight ladder numbers)
# MAGIC
# MAGIC Population ladder is SAS-native. The INBOUND ladder is the CSV flag
# MAGIC (AWS-origin): reconciliation evidence, frozen HERE and nowhere else.
# MAGIC Recorded SAS-side count is 11,154 vs 11,136 replicated: ~18-account
# MAGIC residual [OPEN], superseded by the B01 native flag rebuild - kept on
# MAGIC record, no longer chased.

# COMMAND ----------

sec("A8 waterfall asserts")

_r = spark.sql(f"""
    SELECT count(*)                                        AS s1,
           count_if(wf_dq1)                                AS s2,
           count_if(wf_dq1 AND wf_cpc)                     AS s3,
           count_if(in_sas_ledger)                         AS s4,
           count_if(csv_inb)                               AS i1,
           count_if(csv_inb AND wf_dq1)                    AS i2,
           count_if(csv_inb AND wf_dq1 AND wf_cpc)         AS i3,
           count_if(csv_inb AND in_sas_ledger)             AS i4
    FROM {DB}.uc2_sas_wf_202501
""").first()
chk("wf 01 total", _r["s1"], EXPECTED["wf 01 total"])
chk("wf 02 dq1", _r["s2"], EXPECTED["wf 02 dq1"])
chk("wf 03 +cpc", _r["s3"], EXPECTED["wf 03 +cpc"])
chk("wf 04 sas ledger", _r["s4"], EXPECTED["wf 04 sas ledger"])
chk("wf inb 01 total", _r["i1"], EXPECTED["wf inb 01 total"])
chk("wf inb 02 dq1", _r["i2"], EXPECTED["wf inb 02 dq1"])
chk("wf inb 03 +cpc", _r["i3"], EXPECTED["wf inb 03 +cpc"])
chk("wf inb 04 flagged", _r["i4"], EXPECTED["wf inb 04 flagged"])
print("NOTE: replicated flagged = 11,136 vs recorded 11,154; ~18 residual [OPEN],")
print("superseded by the native flag rebuild (B01); kept on record, not chased.")

# COMMAND ----------

# MAGIC %md
# MAGIC ## A9. The gap decomposition, consolidated (permanent evidence)
# MAGIC
# MAGIC The round-11 set relation, re-derived from this build and persisted:
# MAGIC `uc2_gap1942_202501` (the recovered callers) and `uc2_sasflag_202501`
# MAGIC (the flagged set). The B02 reconciliation cells read THESE tables,
# MAGIC never the CSV.

# COMMAND ----------

sec("A9 gap decomposition")

spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW _sas_flagged AS
    SELECT DISTINCT acct_key, acct_num
    FROM {DB}.uc2_sas_wf_202501
    WHERE in_sas_ledger AND csv_inb
""")
spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW _our_callers AS
    SELECT DISTINCT c.acct_key, try_cast(c.acct_key AS bigint) AS acct_num
    FROM {DB}.uc2_t16_02_episodes c
    JOIN {DB}.uc2_t16_01_populations p ON p.acct_key = c.acct_key
    WHERE c.is_episode_std = 1 AND p.in_ledger_exaa
""")

_r = spark.sql("""
    SELECT (SELECT count(*) FROM _sas_flagged)                                            AS flagged,
           (SELECT count(*) FROM _our_callers)                                            AS ours,
           (SELECT count(*) FROM _sas_flagged f JOIN _our_callers o ON o.acct_num = f.acct_num)      AS shared,
           (SELECT count(*) FROM _sas_flagged f LEFT ANTI JOIN _our_callers o ON o.acct_num = f.acct_num) AS flagged_only,
           (SELECT count(*) FROM _our_callers o LEFT ANTI JOIN _sas_flagged f ON f.acct_num = o.acct_num) AS ours_only
""").first()
chk("gap: sas flagged", _r["flagged"], EXPECTED["gap: sas flagged"])
chk("gap: our callers", _r["ours"], EXPECTED["gap: our callers"])
chk("gap: shared", _r["shared"], EXPECTED["gap: shared"])
chk("gap: flagged only", _r["flagged_only"], EXPECTED["gap: flagged only"])
chk("gap: ours only", _r["ours_only"], EXPECTED["gap: ours only"])

spark.sql(f"""
    CREATE OR REPLACE TABLE {DB}.uc2_gap1942_202501 AS
    SELECT f.acct_key, f.acct_num
    FROM _sas_flagged f LEFT ANTI JOIN _our_callers o ON o.acct_num = f.acct_num
""")
spark.sql(f"""
    CREATE OR REPLACE TABLE {DB}.uc2_sasflag_202501 AS
    SELECT acct_key, acct_num FROM _sas_flagged
""")
print(f"persisted {DB}.uc2_gap1942_202501 and {DB}.uc2_sasflag_202501")

# COMMAND ----------

# MAGIC %md
# MAGIC ## A10. The numeric-rejoin evidence (phase-1d rerun, permanent asserts)
# MAGIC
# MAGIC Same predicate as phase 1d: call dates in January 2025, effdt bounded
# MAGIC [2024-12-01, 2026-07-10). effdt is a load date, so rows below the bound
# MAGIC are immutable and these counts can never legitimately change.

# COMMAND ----------

sec("A10 numeric-rejoin evidence")

spark.sql(f"REFRESH TABLE {CALL_TABLE}")

_r = spark.sql(f"""
    SELECT count(DISTINCT g.acct_key) AS coverage
    FROM {CALL_TABLE} c
    JOIN {DB}.uc2_gap1942_202501 g ON g.acct_num = try_cast(c.acctid AS bigint)
    WHERE c.effdt >= '2024-12-01' AND c.effdt < '2026-07-10'
""").first()
chk("1d: numeric coverage of the 1,942", _r["coverage"], EXPECTED["1d: numeric coverage of the 1,942"])

_r = spark.sql(f"""
    SELECT count(*) AS jan_rows, count(DISTINCT g.acct_key) AS jan_accts
    FROM {CALL_TABLE} c
    JOIN {DB}.uc2_gap1942_202501 g ON g.acct_num = try_cast(c.acctid AS bigint)
    WHERE c.effdt >= '2024-12-01' AND c.effdt < '2026-07-10'
      AND c.initiationmethod = 'INBOUND'
      AND c.`date` >= DATE '2025-01-01' AND c.`date` < DATE '2025-02-01'
""").first()
chk("1d: jan inbound rows recovered", _r["jan_rows"], EXPECTED["1d: jan inbound rows recovered"])
chk("1d: jan inbound accounts recovered", _r["jan_accts"], EXPECTED["1d: jan inbound accounts recovered"])

_r = spark.sql(f"""
    SELECT count_if(initiationmethod = 'INBOUND' AND acctid IS NOT NULL
                    AND (try_cast(acctid AS bigint) IS NULL
                         OR trim(cast(acctid AS string)) <> {NUM_KEY.format(c="acctid")})) AS mismatches,
           count_if(initiationmethod = 'INBOUND' AND acctid IS NULL) AS inb_null_id,
           count_if(acctid IS NULL) AS null_rows,
           count_if(acctid IS NOT NULL AND cast(acctid AS string) rlike '^[0-9]+$') AS digits_rows,
           count_if(acctid IS NOT NULL AND NOT cast(acctid AS string) rlike '^[0-9]+$') AS other_rows
    FROM {CALL_TABLE}
    WHERE `date` >= DATE '2025-01-01' AND `date` < DATE '2025-02-01'
      AND effdt >= '2024-12-01' AND effdt < '2026-07-10'
""").first()
chk("1d: jan key mismatches (id-carrying inbound)", _r["mismatches"], EXPECTED["1d: jan key mismatches (id-carrying inbound)"])
chk("1d: jan inbound rows null acctid", _r["inb_null_id"], EXPECTED["1d: jan inbound rows null acctid"])
chk("1d: jan acctid null rows", _r["null_rows"], EXPECTED["1d: jan acctid null rows"])
chk("1d: jan acctid digits-only rows", _r["digits_rows"], EXPECTED["1d: jan acctid digits-only rows"])
chk("1d: jan acctid other-shape rows", _r["other_rows"], EXPECTED["1d: jan acctid other-shape rows"])

# COMMAND ----------

# MAGIC %md
# MAGIC ## A11. PAYMT_AMT sign probe (a pre-registered CONFIRMATION, not a discovery)
# MAGIC
# MAGIC PRE-REGISTERED (CQ-7, cq-results record, 11 July): a true payment shows as
# MAGIC a NEGATIVE PAYMT_AMT in M1/M2. This probe confirms it; a contradiction is a
# MAGIC STOP-AND-INVESTIGATE finding, never a free choice. On confirmation the gate
# MAGIC predicate is: captured_sas = PAYMT_AMT_M1 < 0 OR PAYMT_AMT_M2 < 0.

# COMMAND ----------

sec("A11 PAYMT_AMT sign probe")

print("PRE-REGISTERED EXPECTATION (CQ-7): payment = NEGATIVE PAYMT_AMT in M1/M2.")

grid("PAYMT sign distribution (all 610,183 and the SAS ledger)", spark.sql(f"""
    SELECT col, scope,
           CASE WHEN v IS NULL THEN 'd. null or non-numeric'
                WHEN v < 0 THEN 'a. negative'
                WHEN v = 0 THEN 'b. zero'
                ELSE 'c. positive' END AS sign_class,
           count(*) AS accounts, round(sum(v), 0) AS total_amt
    FROM (
        SELECT 'M1' AS col, 'a. all accounts' AS scope, try_cast(PAYMT_AMT_M1 AS double) AS v FROM {DB}.uc2_sas_wf_202501
        UNION ALL
        SELECT 'M1', 'b. sas ledger', try_cast(PAYMT_AMT_M1 AS double) FROM {DB}.uc2_sas_wf_202501 WHERE in_sas_ledger
        UNION ALL
        SELECT 'M2', 'a. all accounts', try_cast(PAYMT_AMT_M2 AS double) FROM {DB}.uc2_sas_wf_202501
        UNION ALL
        SELECT 'M2', 'b. sas ledger', try_cast(PAYMT_AMT_M2 AS double) FROM {DB}.uc2_sas_wf_202501 WHERE in_sas_ledger
    )
    GROUP BY col, scope, sign_class
    ORDER BY col, scope, sign_class
"""))

# direction cross-check on a known-payer set: accounts the AWS day-grain gate
# captured must overwhelmingly show negative M1/M2 under the CQ-7 convention
_r = spark.sql(f"""
    WITH cap AS (
        SELECT try_cast(acct_key AS bigint) AS acct_num
        FROM {DB}.uc2_t16_04_outcomes
        WHERE any_captured = 1 AND in_ledger_exaa
        GROUP BY 1
    )
    SELECT count(*) AS joined,
           count_if(coalesce(try_cast(w.PAYMT_AMT_M1 AS double), 0) < 0
                    OR coalesce(try_cast(w.PAYMT_AMT_M2 AS double), 0) < 0) AS with_negative
    FROM cap c
    JOIN {DB}.uc2_sas_wf_202501 w ON w.acct_num = c.acct_num
""").first()
chk("probe: aws-captured accounts joined to export", _r["joined"],
    EXPECTED["probe: aws-captured accounts joined to export"])
chk("probe: of those, negative PAYMT M1 or M2", _r["with_negative"],
    EXPECTED["probe: of those, negative PAYMT M1 or M2"])
print("READ RULE: if the negative share on the known-payer set is not dominant,")
print("STOP AND INVESTIGATE before any captured_sas number is read anywhere.")

# COMMAND ----------

# MAGIC %md
# MAGIC ## A12. The one-time capture-gate delta table
# MAGIC
# MAGIC `uc2_capture_delta_202501`: one row per account in (our 9,389 UNION the
# MAGIC flagged 11,136). aws_captured = the day-grain 30-day gate (round-10 build;
# MAGIC NULL for the 1,942 our string join never saw). captured_sas = the
# MAGIC month-grain account-grain gate (CQ-7 convention). This table is the ONLY
# MAGIC place the two gates ever meet; neither is a denominator here.

# COMMAND ----------

sec("A12 capture-gate delta table")

spark.sql(f"""
CREATE OR REPLACE TABLE {DB}.uc2_capture_delta_202501 AS
WITH ours AS (SELECT acct_num FROM _our_callers),
flagged AS (SELECT acct_num FROM {DB}.uc2_sasflag_202501),
uni AS (
    SELECT coalesce(o.acct_num, f.acct_num) AS acct_num,
           CASE WHEN o.acct_num IS NOT NULL AND f.acct_num IS NOT NULL THEN 'a. both'
                WHEN o.acct_num IS NOT NULL THEN 'b. ours-only (string-keyed)'
                ELSE 'c. flagged-only' END AS construct
    FROM ours o FULL OUTER JOIN flagged f ON f.acct_num = o.acct_num
),
aws AS (
    SELECT try_cast(acct_key AS bigint) AS acct_num, max(any_captured) AS aws_captured
    FROM {DB}.uc2_t16_04_outcomes
    GROUP BY 1
),
sas AS (
    -- captured_sas per the CQ-7 convention confirmed in A11
    SELECT acct_num,
           (coalesce(try_cast(PAYMT_AMT_M1 AS double), 0) < 0
            OR coalesce(try_cast(PAYMT_AMT_M2 AS double), 0) < 0) AS captured_sas
    FROM {DB}.uc2_sas_wf_202501
)
SELECT u.acct_num, u.construct, a.aws_captured, s.captured_sas
FROM uni u
LEFT JOIN aws a ON a.acct_num = u.acct_num
LEFT JOIN sas s ON s.acct_num = u.acct_num
""")
print(f"built {DB}.uc2_capture_delta_202501")

_delta_df = spark.sql(f"""
    SELECT construct, aws_captured, captured_sas, count(*) AS accounts
    FROM {DB}.uc2_capture_delta_202501
    GROUP BY 1, 2, 3
    ORDER BY 1, 2, 3
""")
grid("capture-gate delta cross-tab (construct x aws day-grain x sas month-grain)", _delta_df)

# frozen compare: EXPECTED["delta grid"] is a dict {"construct|aws|sas": accounts}
_measured = {f"{r['construct']}|{fmt(r['aws_captured'])}|{fmt(r['captured_sas'])}": r["accounts"]
             for r in _delta_df.collect()}
if EXPECTED["delta grid"] is None:
    print("delta grid: MEASURED (lock the dict above into EXPECTED['delta grid'] at freeze)")
else:
    assert _measured == EXPECTED["delta grid"], \
        f"ANCHOR MISS delta grid: measured {_measured} vs expected {EXPECTED['delta grid']}"
    print("PASS  delta grid matches the frozen expectation")

# COMMAND ----------

# MAGIC %md
# MAGIC ## A13. Verdict and record block

# COMMAND ----------

sec("A13 verdict")

if FROZEN:
    _open = [k for k, v in EXPECTED.items() if v is None]
    assert not _open, f"FROZEN but EXPECTED still has measure-mode entries: {_open}"
    print("FROZEN run: every value asserted, nothing measured.")
else:
    print("FIRST-RUN MODE (FROZEN = False). To freeze after verification:")
    print("  1. fill every None in EXPECTED with the verified measured value")
    print("  2. replace the A4 header read with the explicit all-string StructType from the A5 census")
    print("  3. set FROZEN = True and re-run clean")

record_block("A_recon_lock_202501")
flush_metrics("A_recon_lock_202501")
