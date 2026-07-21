# B_lean run order (2026-07-21)

Only `B_lean/` files need to run. Nothing in `B_sas_base/` or the repo root.
Derived from each file's actual table reads/writes.

## The one non-obvious rule
`B02` builds the `00n-04n` layer; `B01` reads those and builds `01s`; `B02b`
reads `01s`. So the order is **B02 -> B01 -> B02b**, NOT alphabetical. Running
B01 before B02 fails (its input tables would not exist).

## Phase 1 - the build chain (STRICT sequence, one after another)

| # | Run | Builds | Reads (must exist first) |
|---|---|---|---|
| 1 | `B02_keyfix_aws_layers.py` | 00n,01n,02n,03n,04n (the RE-ANCHOR is in 02n) | round-10 tables + fmt/call/tx |
| 2 | `B01_sas_spine.py` | 01s populations | 00n-04n (from step 1) + SAS csv |
| 3 | `B02b_outcomes_sas.py` | **04s outcomes** (the key table) | 01s + 02n + 04n |
| 4 | `B03_insights_sas.py` | (nothing; prints insights) | 01s + 04s |

Do NOT reorder 1-4.
(`B00_setup.py` is not run standalone - its SETUP is inlined into each B file.)

## Runs in PARALLEL with Phase 1 (standalone)

- `A_recon_lock_202501.py` - reads the SAS csv + old round-10 tables, writes its
  own `uc2_sas_*` reconciliation tables, feeds nothing downstream. Run anytime.

## Phase 2 - after 04s exists (i.e. after step 3), ALL PARALLEL to each other

These only read `04s`/`02n`/`03n` and write no tables (B04 writes at most an
optional CSV). Run in any order, together:

- `B_stmt_distribution.py` - the Story-B proof: leaked vs captured across the
  5-day statement buckets (pre-due 0-24 / post-due 25-55) + the OLD January
  totals side by side with the NEW statement-frame totals and the dropped-episode
  count. THIS is where you see the fresh numbers and why they moved.
- `B04_stmt_sampler.py` - the transcript sampler (masked; produces the export
  grid + optional CSV to output/csv/). Set the WAVE / QUOTAS widgets as needed.
- `B04_checks.py` - certifies the locked B05 pool ties on the re-anchored 04s.

## Certification checks (optional, run-once; each AFTER its core)

Run each `_checks.py` right after the core that built its table. A miss STOPS.
- `A_recon_lock_checks.py`  after step A
- `B02_checks.py`           after step 1
- `B01_checks.py`           after step 2
- `B02b_checks.py`          after step 3
- `B03_checks.py`           after step 4

### How the checks behave under the re-anchor (updated 2026-07-21)

`B02_checks` and `B02b_checks` are now SPLIT so they do not false-STOP:

- **B02_checks** keeps the FRAME-INDEPENDENT anchors as raising asserts (a miss
  STOPS): the fmt key probe, the population sweep (204,323 / 189,146 / ex-AA
  balance / touched_b1 + class split), the call-table key probe, and the two
  partition-sum ties. The FRAME-DEPENDENT counts (ledger callers/episodes,
  addressable stream, work list, language partition, caller classes, W steps,
  the K9r recovery block) are now MEASURE MODE: each prints the fresh
  statement-frame value next to the January reference and the delta, and never
  STOPS. If a frame-INDEPENDENT anchor ever STOPS, that is a real defect to
  investigate; a moved frame-dependent count is the re-anchor working.
- **B02b_checks** is fully measure-mode: all five 04s counts (episodes, callers,
  captured_sas / leaked_sas / W_s accounts) are conditioned on the in-window
  caller set, so all five move; it prints each vs the January reference and
  asserts nothing.
- **B03_checks** is ALSO split: step 1 "called (inb_native, ledger)" = 11,136
  stays raising (reads the 01s spine, frame-independent); steps 2-4
  (callers-with-episodes, intent, leaked) are measure-mode (read the re-anchored
  04s, they move; step 2 drops below step 1, which is the informative part).
- **A_recon_lock_checks, B01_checks** are unchanged raising checks: A and the
  01s spine are pre-re-anchor layers, so their anchors hold. If either STOPS,
  that is a real defect (the lean cleanup broke something), not the re-anchor.

Net: after this split, a clean run STOPS on nothing unless a genuinely
frame-independent anchor moved. The frame-dependent counts print with their
January reference and delta so you can read the shift directly.

## Reading the result

`B_stmt_distribution.py` is the headline output: the statement-timing
distribution (Story B) and the January -> statement number shift explained. The
`_checks` that assert *counts* (B02, B02b, B03 caller/episode/funnel) are
expected to move under the re-anchor; the *account-grain dollar and gate*
anchors should hold. If a dollar/gate anchor moves, that is a real flag to
investigate.
```
