# Databricks notebook source
# MAGIC %md
# MAGIC # B02_checks - run ONCE after B02_keyfix_aws_layers.py.
# MAGIC
# MAGIC RE-ANCHOR NOTE (2026-07-21): the checks in this file are SPLIT.
# MAGIC
# MAGIC   RAISING (still assert; a miss STOPS) - the FRAME-INDEPENDENT anchors,
# MAGIC   built on 00n/01n/fmt/call, which the statement-window episode filter does
# MAGIC   NOT touch: K2 fmt key probe, K4 population sweep (204,323 / 189,146 /
# MAGIC   ex-AA balance / touched_b1 + class split), K5 call-table key probe. If any
# MAGIC   of these moves, it is a REAL defect - investigate.
# MAGIC
# MAGIC   MEASURE MODE (report vs the January reference; never STOPS) - the
# MAGIC   FRAME-DEPENDENT numbers, built on the re-anchored 02n/04n episode set:
# MAGIC   ledger callers/episodes, the addressable stream, the work list, the
# MAGIC   language partition, caller classes, W steps, and the whole K9r recovery
# MAGIC   reconciliation. These MOVE by design (only in-window episodes survive), so
# MAGIC   asserting the January values would false-STOP. In particular the "dropped
# MAGIC   old callers = 0" implication no longer holds: the re-anchor deliberately
# MAGIC   drops out-of-window callers, so that count is now the size of the shift.
# MAGIC
# MAGIC Rebuilds no layer.

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
    """RAISING - frame-independent anchors only. A miss STOPS."""
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


def shift(name, actual, ref):
    """MEASURE MODE - frame-dependent counts that move under the re-anchor.
    Prints the fresh statement-frame value + the January reference + the delta.
    Never raises."""
    if ref is None:
        print(f"MEASURED  {name} = {fmt(actual)}")
        return
    d = actual - ref
    print(f"MEASURED  {name} = {fmt(actual)}   (Jan ref {fmt(ref)}, delta {'+' if d >= 0 else ''}{fmt(d)})")

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

# RE-ANCHOR: the "no old caller disappears" implication NO LONGER HOLDS - the
# statement-window filter deliberately drops out-of-window callers. Report the
# drop count as the size of the shift (measure mode, never STOPS).
_dropped = spark.sql("SELECT o.acct_key FROM _old_callers o LEFT ANTI JOIN _new_callers n ON n.acct_key = o.acct_key").count()
shift("old callers dropped by the re-anchor (out-of-window)", _dropped, 0)

# frame-dependent: numeric-keyed ledger callers and episodes - MOVE (only
# in-window episodes survive), report vs the January reference.
_r = spark.sql(f"""
    SELECT count_if(in_ledger_exaa) AS episodes,
           count(DISTINCT CASE WHEN in_ledger_exaa THEN acct_key END) AS callers
    FROM {DB}.uc2_t16_04n_outcomes
""").first()
shift("ledger episodes (statement frame)", _r["episodes"], E["ledger episodes (numeric key)"])
shift("ledger callers (statement frame)", _r["callers"], E["ledger callers (numeric key)"])

# frame-dependent: the call-day bucket-1 stream + the work list.
_r = spark.sql(f"""
    SELECT count_if(is_addressable) AS addr_episodes,
           count_if(is_addressable AND pay_f > 0 AND captured = 0) AS wl_episodes,
           count(DISTINCT CASE WHEN is_addressable AND pay_f > 0 AND captured = 0 THEN acct_key END) AS wl_accounts
    FROM {DB}.uc2_t16_04n_outcomes
""").first()
shift("addressable episodes (statement frame)", _r["addr_episodes"], E["addressable episodes (callday b1 stream)"])
shift("addressable work list episodes (statement frame)", _r["wl_episodes"], E["addressable work list episodes"])
shift("addressable work list accounts (statement frame)", _r["wl_accounts"], E["addressable work list accounts"])

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
    # frame-dependent counts -> measure vs the Jan reference
    shift(f"lang: {_row['language_group']}", _row["episodes"], _exp_lang.get(_row["language_group"]))
    _lang_total += _row["episodes"]
_r = spark.sql(f"SELECT count_if(in_ledger_exaa) AS n FROM {DB}.uc2_t16_04n_outcomes").first()
# internal-consistency tie (partition sums to the total) - frame-independent, KEEP raising
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
    # 'a. non-caller' grows and the caller classes shrink as callers drop out of
    # the window - frame-dependent, measure vs the Jan reference.
    shift(f"class: {_row['caller_class']}", _row["accounts"], _exp_cls.get(_row["caller_class"]))
    _cls_total += _row["accounts"]
# the classes still partition the WHOLE ex-AA ledger (non-caller + callers), a
# frame-independent tie - KEEP raising.
chk("caller classes re-add to the ex-AA ledger", _cls_total, E["ledger exaa"])

# COMMAND ----------

# K9. W steps (strict leaked-intent, deceased routed out, work list + balance)
_r = spark.sql(f"""
    SELECT count(DISTINCT CASE WHEN leaked_acct AND in_ledger_exaa THEN acct_key END) AS leaked,
           count(DISTINCT CASE WHEN leaked_acct AND in_ledger_exaa AND deceased_acct = 1 THEN acct_key END) AS routed,
           count(DISTINCT CASE WHEN w_flag THEN acct_key END) AS w_accts
    FROM {DB}.uc2_t16_04n_outcomes
""").first()
# W steps are frame-dependent (leaked-intent is defined on the in-window caller
# set) - measure vs the Jan reference.
shift("W strict leaked accounts (statement frame)", _r["leaked"], E["W strict leaked accounts"])
shift("W deceased routed (statement frame)", _r["routed"], E["W deceased routed"])
shift("W accounts (statement frame)", _r["w_accts"], E["W accounts"])
_r = spark.sql(f"""
    SELECT round(sum(jan_eom_bal), 0) AS bal
    FROM (SELECT DISTINCT acct_key, jan_eom_bal FROM {DB}.uc2_t16_04n_outcomes WHERE w_flag)
""").first()
shift("W balance (statement frame)", int(_r["bal"] or 0), E["W balance"])

# COMMAND ----------

# K9r. the 202501 recovery reconciliation. This block compared the OLD-frame
# caller set to the NEW-frame set; the re-anchor deliberately changes the new
# set (window filter), so NONE of these hold as January asserts. The entire
# block is measure-mode: it now describes how the caller set moved, not a
# recovery tie. (The recovery reconciliation itself was a one-time round-12
# artifact; it is preserved here as a shift picture, not a gate.)
if ANCHOR_YM == "202501":
    _gained = spark.sql("SELECT count(*) AS n FROM _new_callers n LEFT ANTI JOIN _old_callers o ON o.acct_key = n.acct_key").first()["n"]
    shift("new-frame callers not in the old frame", _gained, E["gained callers"])

    _recovered = spark.sql(f"""
        SELECT count(*) AS n FROM {DB}.uc2_gap1942_202501 g
        JOIN _new_callers n ON n.acct_key = g.acct_key
    """).first()["n"]
    shift("gap1942 accounts still present in the new frame", _recovered, E["gap1942 recovered"])

    _overlap = spark.sql(f"""
        SELECT count(*) AS n FROM {DB}.uc2_sasflag_202501 f
        JOIN _new_callers n ON n.acct_key = f.acct_key
    """).first()["n"]
    shift("flagged accounts present in the new frame", _overlap, E["flagged overlap (202501 recon)"])

print("B02_checks: DONE. Frame-independent anchors (K2/K4/K5 + the partition ties) "
      "asserted and PASS if reached here. Frame-dependent counts (K9/K9r) reported "
      "in measure mode vs the January reference - they move by design under the "
      "re-anchor. A moved FRAME-INDEPENDENT anchor would have STOPPED above.")
