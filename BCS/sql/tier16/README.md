# Tier 16: the layered UC2 query kit

Five layer files that build on each other, like the SAS 00/01/02 program chain
but with clean layer boundaries, plus five insight files that reproduce every
AWS-side number and table in the technical walkthrough
(uc2-technical-walkthrough-2026-07-13). Each layer scans its base table ONCE
and exposes flags and classifications as columns; downstream queries filter on
flags instead of re-deriving logic. All logic is character-faithful to the
verified tier-14/15 kit (bucket ladder, cleanup rule, ex-AA list, episode
rule, capture gate, lexicon regexes copied verbatim).

## 1. The layer map

| File | Grain | Base scan | What it gives |
| --- | --- | --- | --- |
| `00_acct_monthly.sql` | account x month | `fmt_acct_dba.fmt_acct_c` (one scan, param window) | eom/max bucket, first-DQ and first-B1 days, EOM balance, in-month charge-off date and amount, CPC, original credit limit, payment/autopay/NSF last-dates |
| `01_populations.sql` | account (anchor month) | none new except `future_co` (pruned forward charge-off scan) | cleaned flag, ex-AA flag, CPC class, ledger flags (all / ex-AA), touched-B1 flag and class, runway band, Feb/Mar positions, 31-Jan-anchored CO8/10/12 windows and CO dollars |
| `02_episodes.sql` | call row (January inbound) | `contactcenter_bdp_db.call` (one scan) | business-card and effdt-cap flags AS COLUMNS, standard-episode flag (first inbound call per account per day among filtered rows) |
| `03_signals.sql` | contactid | `contactcenter_bdp_db.transcript` (ONE pass, ever) | seven boolean lexicon flags (deceased, promise, pay, plan, hardship, dispute, execution) + the v2 priority language_group |
| `04_outcomes.sql` | episode (standard, ex-AA) | one day-grain snapshot scan (call-day bucket) | capture gate, caller classes, strict leak list and W routing, call-day bucket / addressable flag, all 01 attributes repeated per episode |

Dependency chain: `00 -> 01 -> {03, 04}`, `02 -> {03, 04}`, `03 -> 04`.

Three base scans live outside layer 00 by design, each documented in its
file header: 01's `future_co` (needs a 12-month forward horizon; pruned by
`chrgoff_dt IS NOT NULL` plus a semi-join), 04's `snap_daily` (day-grain
as-of-call-day bucket, underivable from a monthly layer; semi-joined to
episode accounts), and 02/03's own single scans of the call and transcript
tables.

The insight files (each holds numbered standalone blocks over the layers,
one block per walkthrough table, every block carrying its expected anchors):

| File | Walkthrough sections it serves |
| --- | --- |
| `05_insights_population.sql` | 1, 1.1 (AWS cells), 1.2, App 2.14 pointer |
| `06_insights_motion.sql` | 2.1, 2.2, 2.3, App 2.3 |
| `07_insights_calls.sql` | 4.1 to 4.5, App 2.1 / 2.2 / 2.4 / 2.5 / 2.6 / 2.7 |
| `08_insights_addressable.sql` | 5.1, 5.2, 5.3 |
| `09_insights_diagnostics.sql` | anchor sweep (block 9.0), 1.1 caller checks, App 2.15 / 2.16 pointers |

## 2. The coverage matrix: every walkthrough number to its query

Three kinds of row. INSIGHT n.n = run that block in the tier-16 file.
VERBATIM = the number is reproducible but not from the layers; run the named
tier-11/14/15 original standalone (the block or matrix row says which and
why). OUT OF SCOPE = SAS-side or story-run era; see section 5, do not
attempt to fake it from these layers.

| Walkthrough section / table | How to reproduce | Key anchors |
| --- | --- | --- |
| 1 population walk, AWS cells | INSIGHT 5.1 | 204,323 / 189,146 / $457,943,987 |
| 1 population walk, SAS cells | OUT OF SCOPE (SAS) | 186,412 / $454.2M / ECL $93.5M |
| 1 CO-share cross-check row (19.8 / 23.5 / 26.4) | INSIGHT 5.3 (original Jan-01 windows, recomputed inline) | on the 204,323 base |
| 1.1 first check row (Feb shares) | INSIGHT 5.4 | 55.6 / 8.0 / 36.0 on 189,146 |
| 1.1 caller row + set difference | INSIGHT 9.1 | 9,389 + 179,757 = 189,146; rows 2/3 zero |
| 1.1 "two independent queries" claim | INSIGHT 9.2 | 11,262 / 9,389 on both rows |
| 1.2 CPC distribution | INSIGHT 5.2 | 78,027 / 57,664 / 53,455 / 15,177 = 204,323 |
| 2.1 January entrants by entry day | INSIGHT 6.3 | 492,074 entrants / 340,961 cured |
| 2.2 February outcomes (rollup, with $ ) | INSIGHT 6.2 | 105,215 / 15,054 / 68,093 / 111 / 673; ~$156.6M CO12$ |
| 2.3 outcome by runway band | INSIGHT 6.4 | 167,951 entrants + 21,195 carried-in |
| 3 (the roll price, ECL, stages, HRAM) | OUT OF SCOPE (SAS) | 65,585 rollers / +$29.0M / 57.7% |
| 4.1 the funnel, ledger row | INSIGHT 7.2 (stages 2/4/5/6) | 11,262 / 7,235 / 2,350 / 859 |
| 4.1 stage 3 (has a transcript) | VERBATIM ../tier14/b8_exaa.sql | ledger 10,566 episodes / 8,822 accounts |
| 4.2 caller classes with $ | INSIGHT 7.3 | 179,757 / 6,029 / 1,929 / 1,431 |
| 4.2 wider touched-B1 view | INSIGHT 7.7 | 724,848 (a 464,023 / b 186,714 / c 69,513 / d 4,598) |
| 4.3 language groups with $ | INSIGHT 7.4 | episodes 498/1,374/5,164/442/79/274/3,431 = 11,262 |
| 4.3 by January class (v1 groups) | INSIGHT 7.5 | class-b sum 8,059 |
| 4.4 W build + value | INSIGHT 7.6a | 1,929 -> 164 deceased -> W 1,765 / $7,690,886 / co12$ $4,098,105 |
| 4.4 / App 2.4 follow-forward, ex-AA rows | INSIGHT 7.6b | sum 1,929; deceased rows 164 |
| App 2.4 AA rows (248; the 2,177 total) | VERBATIM ../tier15/b15_exaa_bal.sql | 2,177 = 1,929 + 248 |
| 4.5 / App 2.7 the year, month by month | VERBATIM ../tier15/lx4_exaa_bal.sql (see INSIGHT 7.8 notes) | 118,069 / 113,192 / 35,490 / 32,019 |
| 5.1 live stream, size check | INSIGHT 8.1 | 29,114 addressable episodes |
| 5.1 full live grid (42,870, day bands) | VERBATIM ../tier14/b17_exaa.sql | 18,112 / 6,530 / 4,472; 62.2% within 10 days |
| 5.2 language x captured with $ | INSIGHT 8.2 | 14 cells, episodes sum 29,114 |
| 5.3 walk-down to the addressable number | INSIGHT 8.3 | 29,114 -> 20,805 -> 18,942 -> 1,863 |
| 6.1 lever bases | derived from INSIGHTS 6.x / 7.6 / 8.3 (no new query) | as above |
| 6.1 timing signal / App 2.6 | VERBATIM ../tier14/b11_exaa.sql (see INSIGHT 7.9 notes) | class-b early 3,976 @ 63.5% vs late 1,008 @ 52.6% |
| 6.2 dollar lenses, AWS cells | INSIGHTS 7.6 / 7.8 pointer / 8.3 | $156.6M / $4.10M / $12.2M / $3.88M |
| 6.2 dollar lenses, SAS + year-gross cells | OUT OF SCOPE (SAS lens 1-2; story-run lens 4) | +$29.0M; $332.0M / $172.9M |
| App 2.1 who calls | INSIGHT 7.1 | b + c = 189,146; callers 6,627 + 2,762 |
| App 2.2 funnel by class | INSIGHT 7.2 (+ b8 for stage 3) | stage-2 a 51,670 / b 8,059 / c 3,203 |
| App 2.3 full February transitions | INSIGHT 6.1 | 66 rows; 189,146 / $457,943,987 / ~$156.6M |
| App 2.5 language by class in full | INSIGHT 7.5 | 18 rows; b sums 8,059 |
| App 2.8 - 2.13 (all marked SAS) | OUT OF SCOPE (SAS) | see section 5 |
| App 2.14 CO windows by class and caller | VERBATIM ../tier11/b9_cohort_outcomes.sql (see INSIGHT 5.3 notes) | 19.8 / 23.5 / 26.4 on 204,323 |
| App 2.15 verification-join hole | VERBATIM ../tier11/b18_join_gap.sql (see INSIGHT 9.3 notes) | weekly 27.5 to 28.7% |
| App 2.16 same-customer check 1 | VERBATIM ../tier11/b16_v2_within_cohort_contrast.sql (see INSIGHT 9.4 notes) | 3,802 vs 3,137 on 2,411 accounts |
| App 2.16 platform-wide row (x4) | OUT OF SCOPE (story run) | 31,804 accounts |
| App 2.17 year funnel with balances | OUT OF SCOPE (story run, f1) | $332.0M / $172.9M |
| App 2.18 loss-conversion checks | OUT OF SCOPE (story run, x1/x2/x3/f8/x5) | 27.9% vs 3.4% |
| App 2.19 / 2.20 mined phrases | OUT OF SCOPE (story run, m1/m2) | top-20 lists |
| App 3 / 4 (code, dictionaries) | documentation; the rules live verbatim in layers 00-04 | n/a |

Rounded or derived cells in the walkthrough (percent shares, $M roundings,
sums of grid rows) reproduce by arithmetic on the block outputs; the blocks
return the unrounded cells.

## 3. Running the kit both ways (the dual-mode rule)

Every layer file is written as `WITH <chain> SELECT ...`. Downstream files
start with placeholder CTEs (`populations AS (SELECT * FROM "<schema>"."uc2_t16_01_populations")`)
that stand in for the upstream layers. Every insight block repeats this rule
in its file header.

### Mode 1: tables exist (Databricks or CTAS)

1. In each layer file, uncomment the `CREATE TABLE <schema>.uc2_t16_NN_... AS`
   header line (Athena: `CREATE TABLE ... AS` with your workgroup's output
   location; Databricks: `spark.sql("""<the WITH block>""").write.saveAsTable(...)`
   or `CREATE TABLE ... AS` directly in a notebook cell).
2. Leave the placeholder CTEs in downstream files exactly as they are; they
   already point at the table names. Fill in `<schema>`.
3. Build in order: 00, 01, 02, 03, 04. Re-run the anchor checks after EACH
   layer lands (section 6) before building the next.
4. Insight blocks then run as-is.
5. Rebuilds: 00 is the only expensive layer; 01/03/04 rebuild in minutes off
   it. If the month window moves, rebuild 00 first and re-anchor everything.

### Mode 2: no table access (stitch as CTEs)

The stitch recipe, mechanical:

1. Open the statement you want to run (a layer file or an insight block).
2. For each placeholder CTE at the top, DELETE the placeholder and paste the
   upstream file's CTE chain in its place (drop the upstream file's final
   `SELECT * FROM ...` line; its own placeholders get pasted-through the
   same way, deepest layer first). The upstream chain's last CTE has exactly
   the name the placeholder had (`acct_monthly`, `populations`; 02's and
   03's results are their final bare SELECTs, wrapped per step 3), so the
   references below resolve unchanged.
3. Wrapping rule: where a file's result is its final bare SELECT, wrap that
   SELECT as a CTE named exactly like the placeholder it replaces, e.g.
   `calls AS (SELECT c.acct_key, ... FROM calls_flagged c LEFT JOIN episodes_std e ...)`
   (internal CTE names are chosen to avoid clashing with the wrapper names).
4. Keep only ONE `WITH` keyword at the very top; every subsequent chain joins
   with a comma.
5. Run. Check the anchors (block 9.0 + section 6) BEFORE reading any new
   number.

Paste order for a full 04 stitch: 00 (snap, acct_monthly), then 01
(jan .. pop_base, with 01's final SELECT wrapped as a `populations` CTE),
then 02 (calls_flagged, episodes_std, final SELECT wrapped as a `calls` CTE),
then 03 (drivers, tx, final SELECT wrapped as a `signals` CTE), then 04's own
CTEs and final SELECT.

Which stitch each insight needs is stated at the top of each block's file:
`populations` = 00 -> 01. `calls` = 02 alone. `signals` = 00 -> 01 -> 02 -> 03.
`outcomes` = the full 00 -> 01 -> 02 -> 03 -> 04 stitch.

Cost warnings, learned the hard way:

* The `outcomes` and `signals` stitches contain the ONE transcript pass, and
  the `outcomes` stitch also carries the day-grain call-day scan. Heavy: run
  one such statement per sitting, and NEVER paste the transcript scan twice
  into one statement (the round-9 m4 OOM lesson). A block that references
  both `outcomes` and `signals` still pastes 03 once: the 04 stitch already
  contains 03's chain ending in a `signals` CTE, so both names resolve.
* `populations`-only and `calls`-only stitches are cheap and equal already
  proven tier-14/15 originals (the layer files' anchor headers say which),
  so a clean stitch is also a regression test.

### Worked example, end to end: February motion by caller class (b14 face)

Target: INSIGHT 6.1 in Mode 2. The block references `populations` and
`outcomes`. Assembly:

```sql
WITH
-- 1. layer 00 pasted whole, minus its final SELECT:
snap AS ( ... 00's snap body ... ),
acct_monthly AS ( ... 00's acct_monthly body ... ),
-- 2. layer 01 pasted, its placeholder acct_monthly CTE DELETED (the name
--    now resolves to 00's chain above), its final SELECT wrapped:
jan AS (SELECT * FROM acct_monthly WHERE ym = '202501'),
prv AS ( ... ), feb AS ( ... ), mar AS ( ... ),
future_co AS ( ... ),
pop_base AS ( ... ),
populations AS (SELECT *, (eom_bucket = 1 AND cleaned) AS in_ledger_all,
                ... 01's final SELECT body ... FROM pop_base),
-- 3. layer 02 pasted, final SELECT wrapped as `calls`:
calls_flagged AS ( ... ), episodes_std AS ( ... ),
calls AS (SELECT c.acct_key, ... , CASE WHEN e.contactid IS NOT NULL THEN 1
          ELSE 0 END AS is_episode_std
          FROM calls_flagged c LEFT JOIN episodes_std e ON e.contactid = c.contactid),
-- 4. layer 03 pasted, its populations/calls placeholders deleted, final
--    SELECT wrapped as `signals`:
drivers AS ( ... ), tx AS ( ... ),
signals AS (SELECT contactid, deceased_f, ... , CASE ... END AS language_group FROM tx),
-- 5. layer 04 pasted, all four placeholders deleted, final SELECT wrapped:
episodes_exaa AS ( ... ), pay_lead AS ( ... ), ep AS ( ... ),
snap_daily AS ( ... ), callday AS ( ... ), esig AS ( ... ), esig_acct AS ( ... ),
outcomes AS (SELECT a.acct_key, a.contactid, ... FROM esig_acct a
             LEFT JOIN populations p ON p.acct_key = a.acct_key),
-- 6. the insight block's own body, placeholders already satisfied:
callers AS (
    SELECT acct_key, max_by(caller_class, contactid) AS caller_class
    FROM outcomes
    WHERE in_ledger_exaa
    GROUP BY 1
)
SELECT p.feb_position_b14,
       coalesce(k.caller_class, 'a. non-caller') AS caller_class,
       p.runway_band,
       count(*) AS accounts,
       round(sum(p.eom_bal), 0) AS jan_eom_balance
FROM populations p
LEFT JOIN callers k ON k.acct_key = p.acct_key
WHERE p.in_ledger_exaa
GROUP BY 1, 2, 3
ORDER BY 1, 2, 3
```

Tie-out: accounts across all rows = 189,146; the grid reproduces the
verified b14_exaa_bal 66-row shape. This stitched combination equals the
already-proven tier-15 original, which is the point: a correct stitch
changes nothing but the plumbing.

(`caller_class` is constant per account in 04, so `max_by` is just a picker.)

Two standing dedup rules for any ad-hoc rollup you write yourself:

* Balance / CO-dollar sums at episode grain double-count accounts with
  several episodes: first collapse to one row per (group, acct_key), as
  INSIGHT 7.4's `acct_grp` does.
* An account can sit in two language groups: never add per-group balances
  down to a ledger total.

## 4. Athena (Trino) to Databricks (Spark SQL) migration notes

| Construct used here | Athena/Trino | Databricks/Spark |
| --- | --- | --- |
| `max_by` / `min_by` | native | Spark 3.3+ only (Databricks DBR 11+); on older runtimes rewrite as a `row_number()` pick |
| `date_parse(eff_dt, '%Y%m%d')` | Trino format tokens | `to_date(eff_dt, 'yyyyMMdd')`; every `%Y%m%d` / `%d%b%Y` format string must be translated (`'%d%b%Y'` -> `'ddMMMyyyy'`) |
| `"date"` (quoted column on the call table) | double quotes | backticks: `` `date` `` (Spark treats double quotes as string literals by default) |
| `regexp_like(x, p)` | native | `rlike` / `regexp_like` exists in Spark 3.2+; `x rlike p` is the safe spelling |
| `FILTER (WHERE ...)` on aggregates | supported | NOT supported in Spark SQL; the kit deliberately avoids it (uses CASE inside aggregates) so nothing to change |
| `try_cast` | native | Spark 3.2+ native; identical semantics for these uses |
| `try(...)` wrapper | Trino only | no direct equivalent; the `try(cast(date_parse(...)))` payment-date fallback becomes `to_date(col, 'ddMMMyyyy')` which already returns NULL on failure |
| `count_if(x)` | native | `count_if` exists in Spark 3.0+; or `sum(case when x then 1 else 0 end)` |
| `date_add('day', 30, d)` | Trino 3-arg | Spark: `date_add(d, 30)` |
| `date_diff('day', a, b)` | Trino 3-arg | Spark: `datediff(b, a)` (argument order flips) |
| `substr(eff_dt, 1, 6)` | 1-based | identical in Spark |
| `bool_or` | native | Spark 3.0+: `bool_or` / `any` |

The string month keys (`ym = '202501'`, `eff_dt >= '20241201'`) and the
quoted schema.table names port unchanged. The `min_by(co_amt, co_dt)` NULL
behavior (rows with NULL ordering key ignored) matches between engines, but
re-anchor anyway (section 6).

## 5. Out of scope: numbers this kit CANNOT reproduce (do not fake them)

Two families. Both are quoted in the walkthrough from their own recorded
sources; neither is derivable from the AWS tables this kit reads. If a
request lands on one of these, point at the source record; never approximate
it from the layers.

### 5a. SAS-side numbers (the client's own system)

Source: the SAS pivot record and the CQ-1 to CQ-11 screenshot grids
(repo `uc2-anchoring/records/`, cq-results record), run on the client's SAS
side against the IFRS9 / pricing views. The AWS tables carry no ECL, stage,
or HRAM columns, so there is nothing to compute here.

* The walkthrough section 1 SAS column: 45,362,305 / 202,479 / 15,226 /
  186,412 / $454.2M / ECL $93.5M.
* ALL of section 3: the 65,585 rollers, +$29.0M reserve step, ECL by month,
  IFRS 9 stages, the HRAM splits, CO12 57.7%.
* Appendix 2.8 to 2.13 (each is marked "(SAS)" in its title).
* The 11,154 SAS-side caller count (our side of that check IS covered:
  INSIGHT 9.1).
* Section 6.2 lens 1 and lens 2 dollar cells (ECL-denominated).

### 5b. Story-run era numbers (4 July run, pre-kit queries, never ported)

Source: the story-run record (f1, x1, x2, x3, x5, f8) and the mined-phrase
record (m1, m2), repo `uc2-anchoring/records/`. These predate the verified
tier-11+ kit; their exact populations and filters were never re-baselined,
so re-running them from these layers would produce near-miss numbers that
look like errors. Quote the records.

* Appendix 2.17, the year funnel with balances (f1): 16.1M legs down to
  $332.0M leaked / $172.9M CO8, and section 6.2 lens 4.
* Appendix 2.18, the loss-conversion checks (x1 / x2 / x3 / f8 / x5):
  27.9% vs 3.4%, the repeat-miss ladder, the late-payer curve.
* Appendix 2.16's platform-wide row (x4): 31,804 accounts, 56%/42%, 10%/36%.
* Appendix 2.19 / 2.20 mined phrases (m1 / m2).

### 5c. Reproducible, but only by a standalone original (not these layers)

Not out of scope, but not layer-derivable either; the matrix (section 2) and
the pointing insight blocks name each one: b8 stage 3 (transcript existence),
b9 full grid (pre-AA capture path), b11 (utterance timing), b13 is covered
by INSIGHT 6.3 but its record grain lives in tier-11 too, b17 full live grid
(delinquency spell start), b16_v2 (Jan-Mar scoring), b18 (rows without an
account id), lx4 (twelve call months), b15 AA rows (AA transcript path).
Each runs verbatim from `../tier11`, `../tier14`, or `../tier15` with no
layer dependency.

## 6. The verification rule (non-negotiable)

After ANY migration, restructuring, window change, or engine move, run
INSIGHT 9.0 (the anchor sweep) and re-check the layer anchors BEFORE trusting
a new number. Never let the anchors drift silently. Any miss = STOP, find the
cause, do not quote downstream results.

| Layer | Anchor | Expected |
| --- | --- | --- |
| 00/01 | cleaned January bucket-1 ledger, all products | 204,323 accounts (AA row 15,177 / ~$73,744,823) |
| 00/01 | ex-AA ledger accounts / Jan EOM balance | 189,146 / $457,943,987 (tolerance ~$5) |
| 00/01 | touched-B1 universe (ex-AA, cleaned) | 724,848 (classes a 464,023 / b 186,714 / c 69,513 / d 4,598; b + 2,432 Jan-CO eom-B1 accounts = 189,146) |
| 02/04 | ledger callers / standard episodes | 9,389 / 11,262 |
| 04 | language partition over ledger episodes | sums to 11,262 (m4 rows: 498 / 1,374 / 5,164 / 442 / 79 / 274 / 3,431) |
| 04 | W (strict leak, ex-AA ledger, non-deceased) | 1,765 accounts / $7,690,886 |
| 04 | addressable episodes / addressable work list | 29,114 / 1,863 |

One-line check pattern (run per layer, compare, then proceed):

```sql
SELECT count_if(in_ledger_exaa) AS ledger_exaa,      -- 189,146
       round(sum(CASE WHEN in_ledger_exaa THEN eom_bal END), 0) AS ledger_bal, -- 457,943,987
       count_if(touched_b1) AS touched_b1            -- 724,848
FROM populations
```

CO-window note: all CO8/10/12 layer columns are 31-Jan-2025-anchored
(CO8 [2025-01-31, 2025-09-30), CO10 [.., 2025-11-30), CO12 [.., 2026-01-31)).
The walkthrough's section-1 shares and Appendix 2.4 CO counts use the ORIGINAL
Jan-01 anchor; the blocks that reproduce them (5.3, 7.2 stage 6, 7.4's
`co_12m_accounts_orig`, 7.6) recompute those windows inline from
`co_dt_future`, so BOTH anchors are reachable. Old Jan-01-anchored CO counts
do NOT reproduce from the layer flags directly; that re-baseline is expected
and recorded.
