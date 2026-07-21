# Databricks notebook source
# MAGIC %md
# MAGIC # B02_checks - run ONCE after B02_keyfix_aws_layers.py to certify.
# MAGIC
# MAGIC Re-reads the n-layers (00n/01n/02n/03n/04n), the round-10 string-keyed
# MAGIC tables, the fmt/call sources, and notebook A's persisted recon tables. Runs
# MAGIC the fmt-side key probe, the population anchor sweep, the call-table evidence
# MAGIC ties, the re-anchor (implication stop-rules + measured locked values,
# MAGIC language partition, caller classes, W steps), and the 202501 recovery
# MAGIC reconciliation (gained / recovered / shortfall-cause / flagged-overlap
# MAGIC arithmetic). A miss STOPS. Rebuilds no layer.

# COMMAND ----------

import datetime as _dt


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


def chk(name, actual, expected, tol=0, ctx=None):
    if expected is None:
        print(f"MEASURED  {name} = {fmt(actual)}")
        return
    ok = (abs(actual - expected) <= tol) if tol else (actual == expected)
    if not ok:
        if ctx is not None:
            print(f"CONTEXT for the failing check '{name}':")
            ctx.show(500, truncate=False)
        raise AssertionError(f"ANCHOR MISS {name}: got {fmt(actual)}, expected {fmt(expected)}"
                             + (f" (tol {tol})" if tol else ""))
    print(f"PASS  {name} = {fmt(actual)}")

# COMMAND ----------

CATALOG = "cda_model_shared"
SCHEMA = "ecm_cld_model"
ANCHOR_YM = "202501"
FMT_CATALOG = "634153504162_glue_connection_catalog"
CC_CATALOG = "062108867742_glue_connectivity_catalog"
DB = f"{CATALOG}.{SCHEMA}"
FMT = f"`{FMT_CATALOG}`.fmt_acct_dba.fmt_acct_c"
CALL = f"`{CC_CATALOG}`.contactcenter_bdp_db.`call`"
NUM_KEY = "cast(try_cast({c} AS bigint) AS string)"

_a0 = _dt.date(int(ANCHOR_YM[:4]), int(ANCHOR_YM[4:6]), 1)
_mm = lambda d, k: _dt.date(d.year + (d.month - 1 + k) // 12, (d.month - 1 + k) % 12 + 1, 1)
MONTH_WIN_START = _mm(_a0, -1).strftime("%Y%m%d")   # 20241201
MONTH_WIN_END = _mm(_a0, 3).strftime("%Y%m%d")      # 20250401
CALL_WIN_START = _a0.isoformat()                    # 2025-01-01
CALL_WIN_END = _mm(_a0, 1).isoformat()              # 2025-02-01
EFFDT_SCAN_START = _mm(_a0, -1).isoformat()         # 2024-12-01
EFFDT_HARD_END = "2026-07-10"

E = {
    "ledger all": 204323,
    "ledger AA row": 15177,
    "ledger exaa": 189146,
    "ledger exaa balance": 457943987,   # +/- 5 tolerance
    "touched b1": 724848,
    "touched a. cured": 464023,
    "touched b. bucket 1": 186714,
    "touched c. rolled past": 69513,
    "touched d. jan chargeoff": 4598,
    "jan acctid null rows": 481838,
    "jan acctid digits-only rows": 1267227,
    "jan acctid other-shape rows": 0,
    "jan key mismatches (id-carrying inbound)": 75883,
    "hist ledger callers (string key)": 9389,
    "hist ledger episodes (string key)": 11262,
    "hist callday b1 stream": 29114,
    "ledger callers (numeric key)": 11347,
    "ledger episodes (numeric key)": 13788,
    "addressable episodes (callday b1 stream)": 36594,
    "addressable work list episodes": 2709,
    "addressable work list accounts": 2543,
    "language partition": {
        "a. deceased or estate": 534,
        "b. future-dated promise": 1725,
        "c. payment talk, no promise": 6185,
        "d. plan or settlement talk": 511,
        "e. hardship talk": 99,
        "f. dispute or fraud talk": 306,
        "g. no payment-related language": 4428,
    },
    "caller classes (aws gate)": {
        "a. non-caller": 177799,
        "b. captured (>= 1 paid-30d episode)": 7101,
        "c. leaked-intent (intent, no payment 30d)": 2451,
        "d. other-caller": 1795,
    },
    "W strict leaked accounts": 2451,
    "W deceased routed": 172,
    "W accounts": 2279,
    "W balance": 9277926,
    "gained callers": 1958,
    "gap1942 recovered": 1942,
    "gained outside 1942": 16,
    "flagged overlap (202501 recon)": 11136,
}

# COMMAND ----------

# K2. fmt-side key probe (the numeric key is a no-op on the population side)
_r = spark.sql(f"""
    SELECT count_if(extnl_acct_id IS NOT NULL AND try_cast(extnl_acct_id AS bigint) IS NULL) AS non_castable,
           count_if(extnl_acct_id IS NOT NULL AND try_cast(extnl_acct_id AS bigint) IS NOT NULL
                    AND trim(cast(extnl_acct_id AS string)) <> {NUM_KEY.format(c="extnl_acct_id")}) AS pad_mismatch
    FROM {FMT}
    WHERE sfx_nbr = 0
      AND eff_dt >= '{MONTH_WIN_START}' AND eff_dt < '{MONTH_WIN_END}'
""").first()
chk("fmt id non-castable (00 window)", _r["non_castable"], 0)
chk("fmt id pad-mismatch (00 window)", _r["pad_mismatch"], 0)

# COMMAND ----------

# K4. the population anchor sweep - all assert-exact
_r = spark.sql(f"""
    SELECT count(*) AS rows, count(DISTINCT acct_key) AS accts,
           count_if(in_ledger_all)                                  AS ledger_all,
           count_if(in_ledger_all AND cpc_class = 'AA')             AS ledger_aa,
           count_if(in_ledger_exaa)                                 AS ledger_exaa,
           round(sum(CASE WHEN in_ledger_exaa THEN eom_bal END), 0) AS exaa_bal,
           count_if(touched_b1)                                     AS touched_b1,
           count_if(touched_b1_class LIKE 'a.%')                    AS t_a,
           count_if(touched_b1_class LIKE 'b.%')                    AS t_b,
           count_if(touched_b1_class LIKE 'c.%')                    AS t_c,
           count_if(touched_b1_class LIKE 'd.%')                    AS t_d
    FROM {DB}.uc2_t16_01n_populations
""").first()
chk("01n grain (rows = distinct accounts)", _r["rows"], _r["accts"])
chk("ledger all", _r["ledger_all"], E["ledger all"])
chk("ledger AA row", _r["ledger_aa"], E["ledger AA row"])
chk("ledger exaa", _r["ledger_exaa"], E["ledger exaa"])
chk("ledger exaa balance", int(_r["exaa_bal"] or 0), E["ledger exaa balance"], tol=5)
chk("touched b1", _r["touched_b1"], E["touched b1"])
chk("touched a. cured", _r["t_a"], E["touched a. cured"])
chk("touched b. bucket 1", _r["t_b"], E["touched b. bucket 1"])
chk("touched c. rolled past", _r["t_c"], E["touched c. rolled past"])
chk("touched d. jan chargeoff", _r["t_d"], E["touched d. jan chargeoff"])

# COMMAND ----------

# K5. call-table key probe (immutable under the effdt bound)
spark.sql(f"REFRESH TABLE {CALL}")
_r = spark.sql(f"""
    SELECT count_if(acctid IS NULL) AS null_rows,
           count_if(acctid IS NOT NULL AND cast(acctid AS string) rlike '^[0-9]+$') AS digits_rows,
           count_if(acctid IS NOT NULL AND NOT cast(acctid AS string) rlike '^[0-9]+$') AS other_rows,
           count_if(initiationmethod = 'INBOUND' AND acctid IS NOT NULL
                    AND (try_cast(acctid AS bigint) IS NULL
                         OR trim(cast(acctid AS string)) <> {NUM_KEY.format(c="acctid")})) AS mismatches
    FROM {CALL}
    WHERE `date` >= DATE '{CALL_WIN_START}' AND `date` < DATE '{CALL_WIN_END}'
      AND effdt >= '{EFFDT_SCAN_START}' AND effdt < '{EFFDT_HARD_END}'
""").first()
chk("jan acctid null rows", _r["null_rows"], E["jan acctid null rows"])
chk("jan acctid digits-only rows", _r["digits_rows"], E["jan acctid digits-only rows"])
chk("jan acctid other-shape rows", _r["other_rows"], E["jan acctid other-shape rows"])
chk("jan key mismatches (id-carrying inbound)", _r["mismatches"], E["jan key mismatches (id-carrying inbound)"])

# COMMAND ----------

# K9. re-anchor. Old and new caller sets for the implication stop-rules.
spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW _old_callers AS
    SELECT DISTINCT c.acct_key
    FROM {DB}.uc2_t16_02_episodes c
    JOIN {DB}.uc2_t16_01_populations p ON p.acct_key = c.acct_key
    WHERE c.is_episode_std = 1 AND p.in_ledger_exaa
""")
spark.sql(f"""
    CREATE OR REPLACE TEMP VIEW _new_callers AS
    SELECT DISTINCT acct_key
    FROM {DB}.uc2_t16_04n_outcomes
    WHERE in_ledger_exaa
""")

# implication 1: no old caller may disappear (numeric normalization cannot unmatch)
_dropped_df = spark.sql("SELECT o.acct_key FROM _old_callers o LEFT ANTI JOIN _new_callers n ON n.acct_key = o.acct_key")
chk("dropped old callers (must be zero)", _dropped_df.count(), 0, ctx=_dropped_df)

# measured: numeric-keyed ledger callers and episodes (episodes can only grow)
_r = spark.sql(f"""
    SELECT count_if(in_ledger_exaa) AS episodes,
           count(DISTINCT CASE WHEN in_ledger_exaa THEN acct_key END) AS callers
    FROM {DB}.uc2_t16_04n_outcomes
""").first()
if _r["episodes"] < E["hist ledger episodes (string key)"]:
    raise AssertionError(f"IMPLICATION MISS: episodes {fmt(_r['episodes'])} < historical {fmt(E['hist ledger episodes (string key)'])}")
chk("ledger episodes (numeric key)", _r["episodes"], E["ledger episodes (numeric key)"])
chk("ledger callers (numeric key)", _r["callers"], E["ledger callers (numeric key)"])

# measured: the call-day bucket-1 stream + the work list
_r = spark.sql(f"""
    SELECT count_if(is_addressable) AS addr_episodes,
           count_if(is_addressable AND pay_f > 0 AND captured = 0) AS wl_episodes,
           count(DISTINCT CASE WHEN is_addressable AND pay_f > 0 AND captured = 0 THEN acct_key END) AS wl_accounts
    FROM {DB}.uc2_t16_04n_outcomes
""").first()
if _r["addr_episodes"] < E["hist callday b1 stream"]:
    raise AssertionError(f"IMPLICATION MISS: call-day stream {fmt(_r['addr_episodes'])} < historical {fmt(E['hist callday b1 stream'])}")
chk("addressable episodes (callday b1 stream)", _r["addr_episodes"], E["addressable episodes (callday b1 stream)"])
chk("addressable work list episodes", _r["wl_episodes"], E["addressable work list episodes"])
chk("addressable work list accounts", _r["wl_accounts"], E["addressable work list accounts"])

# COMMAND ----------

# K9. language partition over ledger episodes (must sum to the episode count)
_lang = spark.sql(f"""
    SELECT language_group, count(*) AS episodes
    FROM {DB}.uc2_t16_04n_outcomes
    WHERE in_ledger_exaa
    GROUP BY 1 ORDER BY 1
""").collect()
_exp_lang = E["language partition"]
_lang_total = 0
for _row in _lang:
    chk(f"lang: {_row['language_group']}", _row["episodes"], _exp_lang.get(_row["language_group"]))
    _lang_total += _row["episodes"]
_r = spark.sql(f"SELECT count_if(in_ledger_exaa) AS n FROM {DB}.uc2_t16_04n_outcomes").first()
chk("language partition re-adds to ledger episodes", _lang_total, _r["n"])

# COMMAND ----------

# K9. caller classes (AWS day-grain gate); 'a. non-caller' lives on the 01n side
_cls = spark.sql(f"""
    WITH callers AS (
        SELECT acct_key, max_by(caller_class, contactid) AS caller_class
        FROM {DB}.uc2_t16_04n_outcomes
        WHERE in_ledger_exaa
        GROUP BY 1
    )
    SELECT coalesce(k.caller_class, 'a. non-caller') AS caller_class, count(*) AS accounts
    FROM {DB}.uc2_t16_01n_populations p
    LEFT JOIN callers k ON k.acct_key = p.acct_key
    WHERE p.in_ledger_exaa
    GROUP BY 1 ORDER BY 1
""").collect()
_exp_cls = E["caller classes (aws gate)"]
_cls_total = 0
for _row in _cls:
    chk(f"class: {_row['caller_class']}", _row["accounts"], _exp_cls.get(_row["caller_class"]))
    _cls_total += _row["accounts"]
chk("caller classes re-add to the ex-AA ledger", _cls_total, E["ledger exaa"])

# COMMAND ----------

# K9. W steps (strict leaked-intent, deceased routed out, work list + balance)
_r = spark.sql(f"""
    SELECT count(DISTINCT CASE WHEN leaked_acct AND in_ledger_exaa THEN acct_key END) AS leaked,
           count(DISTINCT CASE WHEN leaked_acct AND in_ledger_exaa AND deceased_acct = 1 THEN acct_key END) AS routed,
           count(DISTINCT CASE WHEN w_flag THEN acct_key END) AS w_accts
    FROM {DB}.uc2_t16_04n_outcomes
""").first()
chk("W strict leaked accounts", _r["leaked"], E["W strict leaked accounts"])
chk("W deceased routed", _r["routed"], E["W deceased routed"])
chk("W accounts", _r["w_accts"], E["W accounts"])
_r = spark.sql(f"""
    SELECT round(sum(jan_eom_bal), 0) AS bal
    FROM (SELECT DISTINCT acct_key, jan_eom_bal FROM {DB}.uc2_t16_04n_outcomes WHERE w_flag)
""").first()
chk("W balance", int(_r["bal"] or 0), E["W balance"])

# COMMAND ----------

# K9r. the 202501 recovery reconciliation (reads A's persisted tables, never the CSV)
if ANCHOR_YM == "202501":
    _gained = spark.sql("SELECT count(*) AS n FROM _new_callers n LEFT ANTI JOIN _old_callers o ON o.acct_key = n.acct_key").first()["n"]
    chk("gained callers", _gained, E["gained callers"])

    _recovered = spark.sql(f"""
        SELECT count(*) AS n FROM {DB}.uc2_gap1942_202501 g
        JOIN _new_callers n ON n.acct_key = g.acct_key
    """).first()["n"]
    chk("gap1942 recovered", _recovered, E["gap1942 recovered"])

    # shortfall causes: an unexplained ('z.') account STOPS the run
    _cause_df = spark.sql(f"""
        WITH short AS (
            SELECT g.acct_key FROM {DB}.uc2_gap1942_202501 g
            LEFT ANTI JOIN _new_callers n ON n.acct_key = g.acct_key
        ),
        r AS (
            SELECT s.acct_key, e.is_biz, e.within_effdt_cap
            FROM short s
            LEFT JOIN {DB}.uc2_t16_02n_episodes e ON e.acct_key = s.acct_key
        ),
        classed AS (
            SELECT acct_key,
                   CASE
                     WHEN max(CASE WHEN is_biz = 0 AND within_effdt_cap = 1 THEN 1 ELSE 0 END) = 1
                       THEN 'z. eligible row exists yet not a caller (unexplained - STOP)'
                     WHEN max(is_biz) IS NULL
                       THEN 'z. no 02n rows at all (unexplained - STOP)'
                     WHEN min(is_biz) = 1 THEN 'a. only business-card rows'
                     WHEN max(is_biz) = 0 THEN 'b. only out-of-effdt-cap rows'
                     ELSE 'c. mixed business-card / out-of-cap rows'
                   END AS cause
            FROM r GROUP BY acct_key
        )
        SELECT cause, count(*) AS accounts FROM classed GROUP BY 1 ORDER BY 1
    """)
    _unexplained = sum(r["accounts"] for r in _cause_df.collect() if r["cause"].startswith("z."))
    chk("gap1942 shortfall unexplained (must be zero)", _unexplained, 0, ctx=_cause_df)

    _outside = spark.sql(f"""
        SELECT count(*) AS n
        FROM _new_callers n
        LEFT ANTI JOIN _old_callers o ON o.acct_key = n.acct_key
        LEFT ANTI JOIN {DB}.uc2_gap1942_202501 g ON g.acct_key = n.acct_key
    """).first()["n"]
    chk("gained outside 1942", _outside, E["gained outside 1942"])

    _overlap = spark.sql(f"""
        SELECT count(*) AS n FROM {DB}.uc2_sasflag_202501 f
        JOIN _new_callers n ON n.acct_key = f.acct_key
    """).first()["n"]
    chk("flagged overlap arithmetic tie (= 9,194 + recovered)", _overlap, 9194 + _recovered)
    chk("flagged overlap (202501 recon)", _overlap, E["flagged overlap (202501 recon)"])

print("B02_checks: ALL PASS - the lean B02 build is certified equivalent to the locked original.")
