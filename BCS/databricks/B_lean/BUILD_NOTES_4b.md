# BUILD_NOTES_4b - the Story-B statement re-anchor + new B04 (2026-07-21)

Story B (statement-timing leakage) is now the CORE frame. The inbound analysis
is re-anchored from calendar-January to each account's STATEMENT CYCLE. Numbers
are expected to MOVE off the January values; the fresh statement-frame numbers
are the deliverable. All work is in `B_lean/`; `B_sas_base/` is untouched
(fallback).

## The statement cycle (domain meaning, encoded in labels/comments)

- `stmt_dt` = statement date (day 0, the bill lands). Source:
  `fmt_acct_c.stmt_last_dt` (NOT a SAS csv column). One per account =
  `max(stmt_last_dt)` over the bounded fmt window (sfx_nbr=0). Numeric-key join.
- Day 0 to ~25 = PRE-DUE (run-up to the payment due date; due date ~= day 25).
- Day 25 to ~56 = POST-DUE (past-due, until the NEXT statement lands ~31 days
  later).
- `days_since_stmt_dt = datediff(call_dt, stmt_dt)`. 5-day buckets
  `floor(days/5)*5` over 0-55, labelled with due-date meaning:
  `00-04 pre-due` .. `20-24 pre-due`, `25-29 post-due` .. `50-54 post-due`,
  plus an `outside 0-55 days` sentinel. `pre_due_f` (0-24) / `post_due_f`
  (25-55) attributes; due day = 25 marked in comments.

## What changed, file by file

### B02_keyfix_aws_layers.py (core)
- SETUP: added the statement-cycle constants `STMT_DUE_DAY = 25`,
  `STMT_WINDOW_DAYS = 56`, `STMT_BUCKET_WIDTH = 5`.
- K6 (02n episode build): added a `stmt_anchor` CTE (acct_key numeric,
  `max(stmt_last_dt)` as `stmt_dt`, bounded to `[MONTH_WIN_START,
  MONTH_WIN_END)`, sfx_nbr=0). Joined it to the calls in a `calls_anchored`
  CTE that computes `days_since_stmt_dt`, `in_stmt_window`, `pre_due_f`,
  `post_due_f`, `stmt_5day_bucket`, `stmt_5day_bucket_start`.
- **THE RE-ANCHOR**: `episodes_std` now requires `in_stmt_window = 1` in
  addition to first-inbound-per-day (`rn = 1`), `is_biz = 0`,
  `within_effdt_cap = 1`. Call-days outside `[stmt_dt, stmt_dt+56)` DROP from
  the episode/caller population. This is the cause of the number move.
- The 02n output carries the statement columns forward. Everything else in K6
  is byte-identical (numeric key, is_biz, within_effdt_cap, had_zero_pad,
  D8 bounded scan + EFFDT_HARD_END guard).
- K8 (04n build): carried `stmt_dt / days_since_stmt_dt / pre_due_f /
  post_due_f / stmt_5day_bucket / stmt_5day_bucket_start` through
  `episodes_exaa` -> `ep` -> `esig` -> `esig_acct` -> the final SELECT. The
  AWS 30-day capture gate, the callday bucket, and every diagnostic are
  unchanged.
- K3/K4 (00n/01n population logic) UNCHANGED: account-grain, frame-independent.
  The 03n regexes and the ex-AA gate are byte-identical.

### B02b_outcomes_sas.py (core)
- The `j` CTE now carries `stmt_dt / days_since_stmt_dt / pre_due_f /
  post_due_f / stmt_5day_bucket / stmt_5day_bucket_start` from 04n onto the
  04s outcomes table (they flow through `SELECT j.*`).
- The account-grain `captured_sas / leaked_sas / w_s_flag` derivations are
  BYTE-IDENTICAL (frame-independent). The caller/episode COUNTS they feed now
  reflect the in-window episode set.

### B03_insights_sas.py (core)
- UNTOUCHED. The analytical blocks report statement-frame numbers
  automatically now that 02n/04s changed. Nothing was added that asserts old
  values.

### B00_setup.py
- NOT edited. The statement-cycle constants are small and live inline in B02's
  SETUP (the file that uses them). Nothing else needs them, so keeping B00 lean.

## New columns and their meaning

| Column | Grain | Unit | Meaning |
|---|---|---|---|
| `stmt_dt` | account (carried to episode) | calendar date | statement date, day 0; from `fmt_acct_c.stmt_last_dt` |
| `days_since_stmt_dt` | episode | days | `datediff(call_dt, stmt_dt)` |
| `in_stmt_window` | call-day (02n only) | 0/1 | the re-anchor keep flag: call in `[stmt_dt, stmt_dt+56)` |
| `pre_due_f` | episode | 0/1 | days 0-24 (run-up to due day 25) |
| `post_due_f` | episode | 0/1 | days 25-55 (past-due) |
| `stmt_5day_bucket` | episode | label | `00-04 pre-due` .. `50-54 post-due`, else `outside 0-55 days` |
| `stmt_5day_bucket_start` | episode | days | the bucket's low edge, for ordering (NULL outside 0-55) |

## The episode re-anchor (one sentence)

Keep first-inbound-per-day BUT keep ONLY episodes whose `call_dt` falls in
`[stmt_dt, stmt_dt+56)`; drop every call-day outside all statement windows.

## Expected direction of the number move

The January-frame totals (episodes 13,486 / callers 11,136 / captured_sas
8,037 / leaked_sas 1,801 / W_s 1,646) are the PRE-re-anchor baseline. The
statement-frame counts are expected to fall (episodes and callers DROP by the
count of January call-days that fell outside `[stmt_dt, stmt_dt+56)`, plus any
account with no fmt statement anchor). The account-grain dollar/gate math is
unchanged, but the episode/caller counts move. This is by design; the
distribution file measures both frames side by side.

## B_stmt_distribution.py (run-once, NOT a core)

Descriptive (measure-mode; NO raising asserts). Reports:
- **(a) STORY-B PROOF**: leaked vs captured episode counts AND dollars
  (eop_bal_m1, gross_loss_12m_amt) across the 5-day statement buckets, with
  pre-due vs post-due subtotals - where leakage concentrates in the cycle.
- **(b) SHIFT EXPLANATION**: the OLD January-frame totals (round-12 record)
  printed side by side with the NEW statement-frame totals the re-anchored 04s
  produces, plus the count of episodes DROPPED for falling outside all
  statement windows (the cause of the move).

## B04_stmt_sampler.py (replaces the export sampler) + governance fixes

- Samples on the RE-ANCHORED 04s, bucket-balanced across the 5-day statement
  buckets (two-stage `xxhash64` deterministic pick: within-`(stratum, bucket)`
  rank, then round-robin across buckets before the quota cap; NO `random()`).
- Strata = the descriptive B05 pool names (`leaked_core`, `leaked_exec`,
  `leaked_promise`, `captured_contrast`, `captured_exec`, `captured_promise`,
  `silent_relaxed`, `reference`). NO reuse of A/B/C/D letters; NO
  `dlnqt_cd_m2` filter on the whole pool. Sub-strata tested before their core.
- Flags: payment = `captured_sas` (locked). spoken promise = transcript
  `promise_f` from 03n (locked, FULL regex including `pay (on|by|this|next)`).
- Waves: `WAVE` widget selects a disjoint `pick_rn` window per stratum;
  `QUOTAS` dict re-shapes the wave without a code edit (wave-1 shape:
  leaked_core 12 / leaked_exec 4 / leaked_promise 4 / captured_contrast 8).
  Anchor excerpts (wave > 1) via `ANCHOR_CONTACTIDS`, emitted at the top with
  `cb_is_anchor = 1`.

### Governance fixes (the Tier-1 must-not-ship items)
- Digit mask ON by default: `regexp_replace(t.content, '[0-9]{3,}', '###')`
  with the SINGLE `-- UNMASK EDIT POINT (owner-gated)` comment. That is the
  ONLY place masking can change; the keeper does not flip it.
- NO HTML writer. NO write to any personal Workspace path. NO DBFS
  `/FileStore` browser-reachable fallback. NO `_unmasked` filename. Output =
  the masked `display()` grid + optional CSV download to `output/csv/` only.
- RESTORED turn ordering:
  `array_join(transform(array_sort(collect_list(struct(beginmillis, line))),
  x -> x.line), char(10))` so turns render in speaking order.
- Correct catalog ids: `062108867742_glue_connectivity_catalog` (call/tx),
  `634153504162_glue_connection_catalog` (fmt), backtick-quoted in SQL.
- 8000-char cap; `text_state` = PARTIAL when the cap clips a call.
- Cheap guards only inside B04 (`assert non-anchor <= quota` per stratum,
  `assert sampled <= quota total`). NO EXPECTED dict, NO locked-value chk.

### The recorded-PTP seam (location)
- Section **X2b** (after the pool build, before the pick). A commented stub
  documents V_COLL_PRMS_DTL_TBL (schema SRC_COLL_DBA) and the three wiring
  steps, and a single well-named placeholder column `recorded_ptp_f`
  (`CAST(NULL AS int)`) is added in the `cb_pool_ptp` view. The export carries
  it as `cb_recorded_ptp_f` so the schema is stable and B06 sees it from wave
  1 on. The table is NOT joined now (unverified). Comment points to
  `ptp-table-probe-spec-2026-07-21.md`.

## B04_checks.py (run-once sibling)

- Uses the raising `chk()` from `_checks_common.py` (via `%run
  ./_checks_common`, with an inline fallback if pasted).
- Rebuilds the raw-predicate pool exactly as B05 measures it (promise_f from
  03n, tx_f from a bounded transcript-exists semi-join) and asserts the LOCKED
  B05 pool ties: `leaked_core = 1,857/1,646 (= W_s)`, `leaked_exec 95/88`,
  `leaked_promise 313/300`, `captured_contrast 6,496/5,658`, `captured_exec
  666/611`, `captured_promise 1,411/1,356`, `silent_relaxed 1,125/953`; the
  strict-G-with-M2-gate = 14/12 consistency tie; transcript coverage
  12,397/10,285; leaked_core GROSS_LOSS_12M = $5,719,683; and the untouched
  population gate (13,486 / 11,136 / 8,037 / 1,801 / 1,646). A miss STOPS the
  wave.

## Run order

`A_recon_lock_202501` (frozen) -> `B00_setup` -> `B01_sas_spine` ->
`B02_keyfix_aws_layers` (emits the statement anchor + re-anchor) ->
`B02b_outcomes_sas` (carries statement columns onto 04s) -> `B03_insights_sas`
-> `B_stmt_distribution` (the shift picture) -> `B04_stmt_sampler` (per wave)
-> `B04_checks` (certify once before a wave ships).

## Files written this build

| File | Lines | Kind |
|---|--:|---|
| B02_keyfix_aws_layers.py | 601 | core (edited) |
| B02b_outcomes_sas.py | 118 | core (edited) |
| B_stmt_distribution.py | 263 | run-once (new) |
| B04_stmt_sampler.py | 536 | core sampler (new) |
| B04_checks.py | 209 | run-once sibling (new) |
| BUILD_NOTES_4b.md | this | notes |

B03_insights_sas.py was left byte-identical on purpose.
