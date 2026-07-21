# Phase-4a cleanup notes: the lean B base

The lean base in `B_lean/` is the round-12 SAS-spine pipeline with the scaffolding
stripped out and every anchor/stop-rule moved into a `_checks` sibling. The cores
build tables and print a one-line summary; the checks re-read those tables and
assert the locked round-12 values. Analytical logic, SQL, regexes, derivations,
table names, and run order are unchanged. Cleanup removed scaffolding only.

Locked values are from the round-12 record
(`uc2-anchoring/records/bridge-round12-phase2-sas-spine-2026-07-16.md`) and each
original's own `EXPECTED` dict. Nothing was invented.

RUN ORDER (unchanged): `A_recon_lock_202501` once -> `B02` -> `B01` -> `B02b` -> `B03`.

## Per-file: original -> lean, what was stripped

Originals live in `databricks/` (A) and `databricks/B_sas_base/` (B00-B03).

### A_recon_lock_202501.py  (702 -> 174; core-only, build stays)
- Removed the `EXPECTED` dict and every `chk()` assert (57 chk/assert sites) -> moved to `A_recon_lock_checks.py`.
- Removed the `grid`/`kv`/`sec`/`record_block`/`flush_metrics`/`RESULTS`/`GRIDS` output plumbing and the `uc2_run_metrics` flush.
- Removed verbose per-block prints; kept a one-line "built ..." print per table.
- KEPT byte-identical: the CSV load (all-string, FAILFAST), the `uc2_sas_raw/wf/gap1942/sasflag/capture_delta` builds, the numeric-key rule, the `CSV_COLUMNS` 90-column census (`assert len == 90` kept as a cheap structural guard), the `call_type_INB` read (A is the one place it is read), the CQ-7 captured_sas definition.

### B00_setup.py  (323 -> 85; canonical SETUP)
- Removed the `chk`/`fmt`/`grid`/`kv`/`sec`/`record_block`/`flush_metrics` helper defs (14 sites) and the `EXPECTED` scaffolding -> anchor checks do not live in SETUP anymore; they are in the `_checks` siblings.
- Removed the verbose banner prints; kept the derived-window block, the numeric-key rule, and a cheap source preflight.
- KEPT byte-identical: all catalog/schema/path constants, the widget-override block, every derived date window and its 202501 comment.

### B01_sas_spine.py  (598 -> 214) + B01_checks.py (142)
- Removed the `EXPECTED` dict and all `chk()` asserts (30 sites) -> `B01_checks.py` (sign tripwire CQ-7, waterfall + native ladder, captured_sas, the perfectly-diagonal CSV-flag tie-out, ledger dollar sums).
- Removed `grid`/`kv`/`sec`/`record_block`/`flush_metrics` and their prints.
- KEPT byte-identical: the CSV load, the column-by-column `uc2_t16_01s_populations_<vintage>` SELECT, the LOOP CUT (no `call_type_*` read; `inb_native` rebuilt natively), the aws_ diagnostic joins by numeric key, the captured_sas definition.

### B02_keyfix_aws_layers.py  (1101 -> 505) + B02_checks.py (332)
- Removed the `EXPECTED` dict and all `chk()` asserts (52 sites) -> `B02_checks.py` (fmt key probe, population anchor sweep, call-table evidence ties, the re-anchor implication stop-rules, language partition, caller classes, W steps, the 202501 recovery reconciliation).
- Removed `grid`/`kv`/`sec`/`record_block`/`flush_metrics` and verbose prints.
- KEPT byte-identical: the 00n/01n/02n/03n/04n builds, THE KEY CHANGE (D2 numeric key) in every layer, the D8 bounded call scan `[EFFDT_SCAN_START, 2026-07-10)` + REFRESH, every transcript regex in 03n, the ex-AA gate and cpc_class map, the AWS day-grain gate as a diagnostic.

### B02b_outcomes_sas.py  (455 -> 109) + B02b_checks.py (77)
- Core was already cleaned by the prior agent; only its check sibling was missing.
- `B02b_checks.py` (NEW): re-reads `uc2_t16_04s_outcomes_<vintage>` and asserts the O3 summary - episodes 13,486 / callers 11,136 / captured_sas accounts 8,037 / leaked_sas accounts 1,801 / W_s accounts 1,646.
- Core KEPT byte-identical (untouched): the `uc2_t16_04s_outcomes_<vintage>` build, the account-grain month-grain captured_sas gate, `leaked_sas` / `w_s_flag` derivations, the aws_ diagnostics.

### B03_insights_sas.py  (590 -> 343) + B03_checks.py (79)
- Removed the `EXPECTED` dict and all `chk()` asserts (20 sites) -> `B03_checks.py` (the four funnel anchors).
- Removed `import time`, `RESULTS`/`GRIDS`, and the `sec`/`grid`/`kv`/`record_block`/`flush_metrics` helpers and their calls; the I0 precondition kept its table-exists asserts but dropped its ledger-count `chk`; the Block 7 trailing measured `chk` became a plain print.
- Replaced `grid(name, df)` / `kv(name, df)` calls with `spark.sql(<identical SQL>).show(50, truncate=False)` so every block still prints. The analytical SQL inside is byte-identical (all 7 blocks: population walk, funnel, language groups, W_s, addressable walk-down, ECL step by class, SAS x AWS continuity bridge).
- KEPT byte-identical: the standing dedup rules, the within-row-money-only rule, the 186,013 vs 186,412 side-by-side disclosure, the captured_sas class denominators.

## Core <-> checks split map

| core file | its checks sibling | locked values the sibling asserts |
|---|---|---|
| A_recon_lock_202501.py | A_recon_lock_checks.py | r10 anchors (204,323 / 189,146 / 724,848; string-key 9,389 callers / 11,262 episodes); CSV grain 610,183 + the 90-column census; export key probe; the 8 waterfall ladder numbers (610,183 / 202,479 / 186,848 / 186,013; INBOUND 34,234 / 12,615 / 11,289 / 11,136); gap decomposition (flagged 11,136 / ours 9,389 / shared 9,194 / flagged-only 1,942 / ours-only 195); numeric-rejoin evidence (1,942; 2,220 rows; 75,883 mismatches; 481,838 null / 1,267,227 digits / 0 other); the A11 PAYMT sign probe (6,029 / 5,767); the A12 capture-gate delta grid |
| B00_setup.py | (none) | SETUP has no anchors; it is inlined into each B core. The source preflight is a reachability guard, not an anchor. |
| B01_sas_spine.py | B01_checks.py | CQ-7 sign tripwire; waterfall 610,183 / 202,479 / 186,848 / 186,013; native ladder 34,234 / 12,615 / 11,289 / 11,136; captured_sas all 278,885 / ledger 125,275; ledger EOP_BAL_M1 $452,444,591 / ECL_M1 $93,543,576; the perfectly-diagonal CSV-flag tie-out (F/F 575,949, T/T 34,234, off-diagonal 0; ledger F/F 174,877, T/T 11,136, off-diagonal 0) |
| B02_keyfix_aws_layers.py | B02_checks.py | fmt key probe (0/0); ledger all 204,323 / AA 15,177 / ex-AA 189,146 / ex-AA balance $457,943,987 (+/-5); touched_b1 724,848 (a. 464,023 / b. 186,714 / c. 69,513 / d. 4,598); call-table ties (481,838 / 1,267,227 / 0 / 75,883); re-anchor (numeric-key ledger callers 11,347 / episodes 13,788; addressable 36,594 / work list 2,709 episodes / 2,543 accounts); language partition (7 groups); caller classes (177,799 / 7,101 / 2,451 / 1,795); W steps (2,451 / 172 / 2,279 / $9,277,926); 202501 recovery (gained 1,958 / recovered 1,942 / outside 16 / flagged overlap 11,136) |
| B02b_outcomes_sas.py | B02b_checks.py | 04s episodes 13,486 / callers 11,136 / captured_sas accounts 8,037 / leaked_sas accounts 1,801 / W_s accounts 1,646 |
| B03_insights_sas.py | B03_checks.py | funnel called (inb_native, ledger) 11,136 / callers with episodes 11,136 / intent accounts 7,459 / leaked accounts 1,801 |

`_checks_common.py` holds the shared `chk()`/`fmt()` helpers (the raising 5-arg
`chk()`; a miss STOPS). Each `_checks` sibling inlines the same two functions at
its top so it runs alone.

## One-time certification instruction for Ravi

Run this once in a fresh environment to certify the lean base equals the locked
originals. For each core, run the core file, then its `_checks` sibling.

1. `A_recon_lock_202501.py`  ->  `A_recon_lock_checks.py`
2. `B02_keyfix_aws_layers.py`  ->  `B02_checks.py`
3. `B01_sas_spine.py`  ->  `B01_checks.py`
4. `B02b_outcomes_sas.py`  ->  `B02b_checks.py`
5. `B03_insights_sas.py`  ->  `B03_checks.py`

(`B00_setup.py` has no checks sibling; it is the canonical SETUP inlined into
each B core.) Every `_checks` sibling ends with an "ALL PASS - certified
equivalent" print. If all five print it with no `AssertionError`, the lean base
is certified equivalent to the locked round-12 originals. A single miss STOPS its
run at the failing anchor - that stop rule is the point of the split.

## Kept, uncertain

Judgment calls made during the cleanup, flagged for review:

- **B03 `grid`/`kv` -> `.show(50, truncate=False)`.** Removing the helpers meant
  the block SQL had to print some other way. `.show()` keeps every block runnable
  and the SQL byte-identical, but its output is truncated/rounded on wide grids,
  not the lossless markdown the original `grid()` produced. If a future sitting
  needs to transcribe B03 output from screenshots, restore a lossless printer or
  read the values off `B03_checks.py` instead. The certified numbers come from the
  checks sibling, not from B03's console output, so this does not affect certification.
- **B00 source preflight kept.** The `spark.catalog.tableExists` reachability loop
  in B00 is a precondition guard, not an anchor. Kept because it is cheap and
  catches a mis-set catalog widget early. It asserts (STOPS) on an unreachable
  source; if that is unwanted noise in a checks-free SETUP, drop it.
- **`assert len(CSV_COLUMNS) == 90` kept in the A core.** It is technically an
  assert, but a structural self-check on a literal list, not a data anchor (the
  full 90-column header tie-out lives in `A_recon_lock_checks.py`). Kept as a
  cheap guard against an accidental edit to the pinned census.
- **Table-exists preconditions kept in every core** (`B01` S1, `B02` K1, `B02b`
  O1, `B03` I0, `A` A3). These are `assert`s but run-order guards, not value
  anchors. Kept so a core fails fast with a clear message when run out of order,
  rather than producing a wrong table.
