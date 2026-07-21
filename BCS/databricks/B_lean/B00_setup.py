# Databricks notebook source
# MAGIC %md
# MAGIC # B00. Setup (canonical LEAN SETUP block)
# MAGIC
# MAGIC Inlined verbatim at the top of every B core file (B01/B02/B02b/B03) so each
# MAGIC runs alone. Edit here, re-paste into every B file ("SETUP - keep in sync").
# MAGIC RUN ORDER: A_recon_lock_202501 once -> B02 -> B01 -> B02b -> B03.
# MAGIC NUMERIC KEY RULE: acct_num = try_cast(id AS bigint); acct_key = cast(acct_num AS string).
# MAGIC ANCHOR/STOP-RULE CHECKS DO NOT LIVE HERE. They are in the *_checks siblings,
# MAGIC run once after each core file to certify a fresh environment.

# COMMAND ----------

# =====================================================================
# SETUP - keep in sync across B00/B01/B02/B02b/B03 (this is the canonical copy).
# =====================================================================
import datetime as _dt

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

PRV_YM = _mm(_a0, -1).strftime("%Y%m")                                   # 202412
FEB_YM = _mm(_a0, 1).strftime("%Y%m")                                    # 202502
MAR_YM = _mm(_a0, 2).strftime("%Y%m")                                    # 202503
MONTH_WIN_START = _mm(_a0, -1).strftime("%Y%m%d")                        # 20241201
MONTH_WIN_END = _mm(_a0, 3).strftime("%Y%m%d")                           # 20250401 (excl)
CALL_WIN_START = _a0.isoformat()                                         # 2025-01-01
CALL_WIN_END = _mm(_a0, 1).isoformat()                                   # 2025-02-01
EFFDT_CAP_START = _a0.isoformat()                                        # 2025-01-01
EFFDT_CAP_END = (_mm(_a0, 1) + _dt.timedelta(days=1)).isoformat()        # 2025-02-02
CLEANUP_DATE = _a0.isoformat()                                           # 2025-01-01
ANCHOR_EOM = (_mm(_a0, 1) - _dt.timedelta(days=1)).isoformat()           # 2025-01-31
FEB_START = _mm(_a0, 1).isoformat()                                      # 2025-02-01
MAR_START = _mm(_a0, 2).isoformat()                                      # 2025-03-01
APR_START = _mm(_a0, 3).isoformat()                                      # 2025-04-01
CO8_END = (_mm(_a0, 9) - _dt.timedelta(days=1)).isoformat()             # 2025-09-30
CO10_END = (_mm(_a0, 11) - _dt.timedelta(days=1)).isoformat()           # 2025-11-30
CO12_END = (_mm(_a0, 13) - _dt.timedelta(days=1)).isoformat()           # 2026-01-31
FWD_CO_START = _a0.strftime("%Y%m%d")                                    # 20250101
FWD_CO_END = _mm(_a0, 12).strftime("%Y%m%d")                             # 20260101
SNAP_DAILY_START = _mm(_a0, -7).strftime("%Y%m%d")                       # 20240601
SNAP_DAILY_END = _mm(_a0, 1).strftime("%Y%m%d")                          # 20250201
EFFDT_SCAN_START = _mm(_a0, -1).isoformat()                              # 2024-12-01
# NOT vintage-derived: the bounded-scan guard on the call table's live loading
# edge (round-11 lesson) and the frozen predicate behind the 75,883 tie.
EFFDT_HARD_END = "2026-07-10"

NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"   # THE numeric key rule

print(f"SETUP OK: vintage {ANCHOR_YM}; layers -> {DB}")
# =====================================================================
# end of SETUP
# =====================================================================

# COMMAND ----------

# Source preflight (cheap; run once per sitting)
for _t in [f"{FMT_CATALOG}.fmt_acct_dba.fmt_acct_c",
           f"{CC_CATALOG}.contactcenter_bdp_db.call",
           f"{CC_CATALOG}.contactcenter_bdp_db.transcript"]:
    assert spark.catalog.tableExists(_t), f"PRECONDITION MISS: source {_t} not reachable - fix the catalog widget"
    print(f"PASS  source reachable: {_t}")
