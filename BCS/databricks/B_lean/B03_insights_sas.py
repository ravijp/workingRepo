# Databricks notebook source
# MAGIC %md
# MAGIC # B03. Insights, SAS-denominated (numbered blocks, each with its tie-out)
# MAGIC
# MAGIC Every denominator is the SAS spine (01s) or the 04s episode table; every
# MAGIC dollar is one of the export's own columns. Values are MEASURED here; they
# MAGIC enter records only after verification. Run B03_checks.py once after this to
# MAGIC certify the funnel anchors.
# MAGIC
# MAGIC STANDING RULES (kit README, verbatim):
# MAGIC * Balance / CO-dollar sums at episode grain double-count accounts with
# MAGIC   several episodes: first collapse to one row per (group, acct_key).
# MAGIC * An account can sit in two language groups: never add per-group balances
# MAGIC   down to a ledger total.
# MAGIC Plus: the 202,479 / 186,013 / 186,412 triplet is three constructions of
# MAGIC "the SAS population" - each use below names which one it is. 186,013
# MAGIC (export replication) and 186,412 (SAS-recorded slice) are disclosed side
# MAGIC by side, NEVER asserted equal. Nothing here references the 19,789
# MAGIC statement-window population or its recorded metrics (a separate caller
# MAGIC construct, never printed next to these).

# COMMAND ----------

# =====================================================================
# SETUP - keep in sync across B00/B01/B02/B02b/B03 (B00 is the canonical copy).
# =====================================================================
import datetime as _dt

CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
ANCHOR_YM = "202501"
SAS_CSV_PATH = "/Volumes/cda_model_shared/ecm_cld_model/ecm_cld/collections_zenon/WATERFALL_COLL_CALL_V2_202501.csv"
FMT_CATALOG = "634153504162_glue_connection_catalog"
CC_CATALOG = "062108867742_glue_connectivity_catalog"

try:
    dbutils.widgets.text("CATALOG", CATALOG);           CATALOG = dbutils.widgets.get("CATALOG")
    dbutils.widgets.text("SCHEMA", SCHEMA);             SCHEMA = dbutils.widgets.get("SCHEMA")
    dbutils.widgets.text("ANCHOR_YM", ANCHOR_YM);       ANCHOR_YM = dbutils.widgets.get("ANCHOR_YM")
    dbutils.widgets.text("SAS_CSV_PATH", SAS_CSV_PATH); SAS_CSV_PATH = dbutils.widgets.get("SAS_CSV_PATH")
    dbutils.widgets.text("FMT_CATALOG", FMT_CATALOG);   FMT_CATALOG = dbutils.widgets.get("FMT_CATALOG")
    dbutils.widgets.text("CC_CATALOG", CC_CATALOG);     CC_CATALOG = dbutils.widgets.get("CC_CATALOG")
except NameError:
    pass

DB = f"{CATALOG}.{SCHEMA}"
FMT = f"`{FMT_CATALOG}`.fmt_acct_dba.fmt_acct_c"
CALL = f"`{CC_CATALOG}`.contactcenter_bdp_db.`call`"
TX = f"`{CC_CATALOG}`.contactcenter_bdp_db.transcript"

_a0 = _dt.date(int(ANCHOR_YM[:4]), int(ANCHOR_YM[4:6]), 1)
_mm = lambda d, k: _dt.date(d.year + (d.month - 1 + k) // 12, (d.month - 1 + k) % 12 + 1, 1)

PRV_YM = _mm(_a0, -1).strftime("%Y%m")
FEB_YM = _mm(_a0, 1).strftime("%Y%m")
MAR_YM = _mm(_a0, 2).strftime("%Y%m")
MONTH_WIN_START = _mm(_a0, -1).strftime("%Y%m%d")
MONTH_WIN_END = _mm(_a0, 3).strftime("%Y%m%d")
CALL_WIN_START = _a0.isoformat()
CALL_WIN_END = _mm(_a0, 1).isoformat()
EFFDT_CAP_START = _a0.isoformat()
EFFDT_CAP_END = (_mm(_a0, 1) + _dt.timedelta(days=1)).isoformat()
CLEANUP_DATE = _a0.isoformat()
ANCHOR_EOM = (_mm(_a0, 1) - _dt.timedelta(days=1)).isoformat()
FEB_START = _mm(_a0, 1).isoformat()
MAR_START = _mm(_a0, 2).isoformat()
APR_START = _mm(_a0, 3).isoformat()
CO8_END = (_mm(_a0, 9) - _dt.timedelta(days=1)).isoformat()
CO10_END = (_mm(_a0, 11) - _dt.timedelta(days=1)).isoformat()
CO12_END = (_mm(_a0, 13) - _dt.timedelta(days=1)).isoformat()
FWD_CO_START = _a0.strftime("%Y%m%d")
FWD_CO_END = _mm(_a0, 12).strftime("%Y%m%d")
SNAP_DAILY_START = _mm(_a0, -7).strftime("%Y%m%d")
SNAP_DAILY_END = _mm(_a0, 1).strftime("%Y%m%d")
EFFDT_SCAN_START = _mm(_a0, -1).isoformat()
EFFDT_HARD_END = "2026-07-10"   # not vintage-derived: the live-loading-edge guard

NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"

print(f"SETUP OK: vintage {ANCHOR_YM}; layers -> {DB}")
# =====================================================================
# end of SETUP
# =====================================================================

# COMMAND ----------

# I0. preconditions
POP = f"{DB}.uc2_t16_01s_populations_{ANCHOR_YM}"
OUT = f"{DB}.uc2_t16_04s_outcomes_{ANCHOR_YM}"
for _t in [POP, OUT]:
    assert spark.catalog.tableExists(_t), f"PRECONDITION MISS: {_t} missing - run B01/B02b first"

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 1. The population walk with money (waterfall stages x EOP_BAL x ECL)
# MAGIC
# MAGIC 186,013 below is the EXPORT REPLICATION of the SAS slice. The SAS-RECORDED
# MAGIC slice is 186,412 accounts / $454.2M EOP balance / $93.5M ECL (client-side
# MAGIC pivot record). Different constructions, quoted side by side, never
# MAGIC asserted equal.

# COMMAND ----------

spark.sql(f"""
    SELECT '01. export total (610,183 construction)' AS stage,
           count(*) AS accounts,
           round(sum(eop_bal_m1), 0) AS eop_bal_m1,
           round(sum(ecl_m1), 0) AS ecl_m1
    FROM {POP}
    UNION ALL
    SELECT '02. DQ1 (DLNQT_CD_M1 = 1; the 202,479 construction)',
           count(*), round(sum(eop_bal_m1), 0), round(sum(ecl_m1), 0)
    FROM {POP} WHERE wf_dq1
    UNION ALL
    SELECT '03. + CPC eligible',
           count(*), round(sum(eop_bal_m1), 0), round(sum(ecl_m1), 0)
    FROM {POP} WHERE wf_dq1 AND wf_cpc
    UNION ALL
    SELECT '04. + non-chargeoff = the SAS ledger (the 186,013 construction)',
           count(*), round(sum(eop_bal_m1), 0), round(sum(ecl_m1), 0)
    FROM {POP} WHERE in_sas_ledger
    ORDER BY stage
""").show(50, truncate=False)
print("SIDE-BY-SIDE (disclosure): SAS-recorded slice = 186,412 / $454.2M / ECL $93.5M.")

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 2. The funnel, SAS-denominated
# MAGIC
# MAGIC Caller flag = inb_native (any January INBOUND id-resolved call, numeric
# MAGIC key). Episodes and language from the fixed 04s layer. Leaked money =
# MAGIC export columns at account grain.

# COMMAND ----------

_r = spark.sql(f"""
    SELECT count_if(in_sas_ledger) AS ledger,
           count_if(in_sas_ledger AND inb_native) AS called
    FROM {POP}
""").first()
_s = spark.sql(f"""
    SELECT count(DISTINCT acct_key) AS ep_callers,
           count(*) AS episodes,
           count(DISTINCT CASE WHEN pay_f > 0 THEN acct_key END) AS intent_accts,
           count(DISTINCT CASE WHEN leaked_sas THEN acct_key END) AS leaked_accts
    FROM {OUT}
""").first()
spark.sql(f"""
    SELECT '1. the SAS ledger (186,013 construction)' AS step, {_r["ledger"]} AS value
    UNION ALL SELECT '2. called in January (inb_native)', {_r["called"]}
    UNION ALL SELECT '3. accounts with standard episodes (04s)', {_s["ep_callers"]}
    UNION ALL SELECT '4. standard episodes (04s)', {_s["episodes"]}
    UNION ALL SELECT '5. accounts with payment language on >= 1 episode', {_s["intent_accts"]}
    UNION ALL SELECT '6. leaked_sas accounts (no account payment M1/M2)', {_s["leaked_accts"]}
    ORDER BY step
""").show(50, truncate=False)

spark.sql(f"""
    SELECT count(*) AS leaked_accounts,
           round(sum(eop_bal_m1), 0) AS eop_bal_m1,
           round(sum(gross_loss_12m_amt), 0) AS gross_loss_12m,
           round(sum(chrgoff_12m_amt), 0) AS chrgoff_12m
    FROM (SELECT DISTINCT acct_key, eop_bal_m1, gross_loss_12m_amt, chrgoff_12m_amt
          FROM {OUT} WHERE leaked_sas)
""").show(50, truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 3. Language groups with export dollars
# MAGIC
# MAGIC Episode counts are the clean split (they partition). Account and money
# MAGIC columns compare WITHIN a row only: an account can sit in two groups, so
# MAGIC per-group money is NEVER added down a column. Money is collapsed to one
# MAGIC row per (group, account) before summing (the kit's standing dedup rules,
# MAGIC verbatim).

# COMMAND ----------

spark.sql(f"""
    WITH acct_grp AS (
        SELECT DISTINCT language_group, acct_key, eop_bal_m1, ecl_liftm_m1, gross_loss_12m_amt
        FROM {OUT}
    ),
    ep AS (
        SELECT language_group, count(*) AS episodes, count(DISTINCT acct_key) AS accounts
        FROM {OUT} GROUP BY 1
    ),
    money AS (
        SELECT language_group,
               round(sum(eop_bal_m1), 0) AS eop_bal_m1,
               round(sum(ecl_liftm_m1), 0) AS ecl_liftm_m1,
               round(sum(gross_loss_12m_amt), 0) AS gross_loss_12m
        FROM acct_grp GROUP BY 1
    )
    SELECT e.language_group, e.episodes, e.accounts,
           m.eop_bal_m1, m.ecl_liftm_m1, m.gross_loss_12m
    FROM ep e JOIN money m ON m.language_group = e.language_group
    ORDER BY e.language_group
""").show(50, truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 4. W_s valued in the client's own columns
# MAGIC
# MAGIC W_s = leaked_sas accounts (no account payment M1/M2, >= 1 payment-language
# MAGIC episode), deceased-language accounts routed out; in_sas_ledger by
# MAGIC construction. Money at account grain, export columns.

# COMMAND ----------

spark.sql(f"""
    WITH acct AS (
        SELECT DISTINCT acct_key, leaked_sas, w_s_flag, deceased_acct,
               eop_bal_m1, ecl_liftm_m1, gross_loss_12m_amt, chrgoff_12m_amt
        FROM {OUT}
    )
    SELECT '1. leaked_sas accounts' AS step, count(*) AS accounts,
           round(sum(eop_bal_m1), 0) AS eop_bal_m1,
           round(sum(ecl_liftm_m1), 0) AS ecl_liftm_m1,
           round(sum(gross_loss_12m_amt), 0) AS gross_loss_12m,
           round(sum(chrgoff_12m_amt), 0) AS chrgoff_12m
    FROM acct WHERE leaked_sas
    UNION ALL
    SELECT '2. deceased or estate, routed out', count(*),
           round(sum(eop_bal_m1), 0), round(sum(ecl_liftm_m1), 0),
           round(sum(gross_loss_12m_amt), 0), round(sum(chrgoff_12m_amt), 0)
    FROM acct WHERE leaked_sas AND deceased_acct = 1
    UNION ALL
    SELECT '3. W_s, the work list', count(*),
           round(sum(eop_bal_m1), 0), round(sum(ecl_liftm_m1), 0),
           round(sum(gross_loss_12m_amt), 0), round(sum(chrgoff_12m_amt), 0)
    FROM acct WHERE w_s_flag
    ORDER BY step
""").show(50, truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 5. The addressable moment, re-denominated
# MAGIC
# MAGIC The call-day walk-down on the fixed 04s episodes. DISCLOSED CONSTRUCTION
# MAGIC CHANGE: the capture split is the ACCOUNT-grain month-grain captured_sas,
# MAGIC not the old episode-grain 30-day gate; the old steps 3/4 (episode-grain)
# MAGIC do not exist under this gate. Money deduped at account grain.

# COMMAND ----------

spark.sql(f"""
    WITH addr AS (SELECT * FROM {OUT} WHERE is_addressable),
    intent_acct AS (
        SELECT DISTINCT acct_key, captured_sas, eop_bal_m1, gross_loss_12m_amt
        FROM addr WHERE pay_f > 0
    )
    SELECT '1. bucket 1 on the call day (episodes)' AS step,
           (SELECT count(*) FROM addr) AS value
    UNION ALL SELECT '2. from accounts', (SELECT count(DISTINCT acct_key) FROM addr)
    UNION ALL SELECT '3. payment-language episodes', (SELECT count(*) FROM addr WHERE pay_f > 0)
    UNION ALL SELECT '4. payment-language accounts', (SELECT count(*) FROM intent_acct)
    UNION ALL SELECT '5. of 4: account captured_sas (payment in call month or next)',
           (SELECT count(*) FROM intent_acct WHERE captured_sas)
    UNION ALL SELECT '6. of 4: account NOT captured_sas (the addressable moment)',
           (SELECT count(*) FROM intent_acct WHERE NOT captured_sas)
    ORDER BY step
""").show(50, truncate=False)
spark.sql(f"""
    WITH addr AS (SELECT * FROM {OUT} WHERE is_addressable),
    intent_acct AS (
        SELECT DISTINCT acct_key, captured_sas, eop_bal_m1, gross_loss_12m_amt
        FROM addr WHERE pay_f > 0
    )
    SELECT count(*) AS accounts,
           round(sum(eop_bal_m1), 0) AS eop_bal_m1,
           round(sum(gross_loss_12m_amt), 0) AS gross_loss_12m
    FROM intent_acct WHERE NOT captured_sas
""").show(50, truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 6. The ECL step M1 -> M2 by caller class (captured_sas classes)
# MAGIC
# MAGIC Population = the SAS ledger (186,013 construction); ECL = the export's
# MAGIC own columns; classes = the captured_sas account classes ('a. non-caller'
# MAGIC = ledger accounts with no 04s episodes). Accounts with a null ECL in
# MAGIC either month are counted separately and excluded from the step sum.

# COMMAND ----------

spark.sql(f"""
    WITH cls AS (
        SELECT acct_key, max(caller_class_sas) AS caller_class
        FROM {OUT} GROUP BY 1
    )
    SELECT coalesce(k.caller_class, 'a. non-caller') AS caller_class,
           count(*) AS accounts,
           count_if(p.ecl_m1 IS NOT NULL AND p.ecl_m2 IS NOT NULL) AS accounts_with_both_ecl,
           round(sum(CASE WHEN p.ecl_m1 IS NOT NULL AND p.ecl_m2 IS NOT NULL
                          THEN p.ecl_m2 - p.ecl_m1 END), 0) AS ecl_step_m1_to_m2,
           round(sum(p.ecl_m1), 0) AS ecl_m1,
           round(sum(p.ecl_m2), 0) AS ecl_m2
    FROM {POP} p
    LEFT JOIN cls k ON k.acct_key = p.acct_key
    WHERE p.in_sas_ledger
    GROUP BY 1 ORDER BY 1
""").show(50, truncate=False)

# COMMAND ----------

# MAGIC %md
# MAGIC ## Block 7. The ONE SAS x AWS continuity bridge
# MAGIC
# MAGIC The single table where the two populations meet: the SAS ledger
# MAGIC (186,013 construction) against the AWS ex-AA ledger (189,146), account
# MAGIC counts and gate rates over the export universe, plus the one count that
# MAGIC lives outside it (AWS-ledger accounts with no export row). Nothing else
# MAGIC ever mixes the two populations in one table.

# COMMAND ----------

spark.sql(f"""
    SELECT in_sas_ledger, aws_in_ledger_exaa,
           count(*) AS accounts,
           count_if(captured_sas) AS captured_sas_accts,
           count_if(aws_captured) AS aws_captured_accts
    FROM {POP}
    GROUP BY 1, 2 ORDER BY 1, 2
""").show(50, truncate=False)
_r = spark.sql(f"""
    SELECT count(*) AS n
    FROM {DB}.uc2_t16_01n_populations p
    LEFT ANTI JOIN {POP} s ON s.acct_key = p.acct_key
    WHERE p.in_ledger_exaa
""").first()
print(f"AWS ex-AA ledger accounts with NO export row = {_r['n']:,}")

# COMMAND ----------

print("B03_insights_sas complete: blocks 1-7 printed. Run B03_checks.py once to "
      "certify the funnel anchors (called / callers-with-episodes / intent / leaked).")
