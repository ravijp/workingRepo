# UC2 Phase-2 Databricks package: the SAS-spine pipeline

Nine notebook-source files. Population, delinquency, and dollars come from
the client's SAS 003-program export; the AWS call and transcript tables stay
the sole conversation source; the call-join key is numeric everywhere. Every
file runs alone, pasted as ONE cell or imported as a notebook (no %run).

## The files and the run order

| # | File | Builds | Needs |
| --- | --- | --- | --- |
| 1 | `A_recon_lock_202501.py` | `uc2_sas_raw_202501`, `uc2_sas_wf_202501`, `uc2_gap1942_202501`, `uc2_sasflag_202501`, `uc2_capture_delta_202501` | round-10 tables, the export CSV, the call table |
| 2 | `B_sas_base/B02_keyfix_aws_layers.py` | `uc2_t16_00n_acct_monthly` .. `uc2_t16_04n_outcomes` (numeric key) | sources, round-10 tables, A's tables |
| 3 | `B_sas_base/B01_sas_spine.py` | `uc2_t16_01s_populations_<vintage>` | the n-layers, the CSV, A's wf table (tie-out only) |
| 4 | `B_sas_base/B02b_outcomes_sas.py` | `uc2_t16_04s_outcomes_<vintage>` | 01s + the n-layers |
| 5a | `B_sas_base/B03_insights_sas.py` | nothing (reads 01s/04s) | 01s + 04s |
| 5b | `B_sas_base/B04_copilot_export_sas.py` | nothing (temp views + the export grid) | 04s + transcript |
| 5c | `B_sas_base/B05_scale_pools_probe.py` | nothing (read-only probe: scale-campaign pools + the contact-center summary/category coverage) | 04s + 03n + transcript (+ summary/category if reachable) |
| 5d | `B_sas_base/B06_copilot_labels_ingest.py` | `uc2_copilot_excerpt_map`, `uc2_copilot_labels` (assistant responses become data) | 04s + transcript (map replay); then paste-per-response |

`B_sas_base/B00_setup.py` is the CANONICAL COPY of the SETUP block inlined at
the top of every B file. Edit it there, then re-paste into every B file.

**Parallelism: the chain 1 -> 2 -> 3 -> 4 is STRICTLY SEQUENTIAL** (each step
writes tables the next step reads, and each file's precondition cell
hard-fails if its inputs are missing, so a wrong order stops loud, never
silently). **B03, B04, and B05 can run in parallel** (steps 5a-5c): they
only read 01s/04s/03n; their concurrent appends to `uc2_run_metrics` are safe
(Delta). B06 (5d) also runs any time after B02b, but is a per-response-file
loop, not a chain step. Do not run two BUILDER notebooks at once even on
different clusters: they CREATE OR REPLACE the same table names (B06's two
campaign tables are CREATE IF NOT EXISTS + keyed DELETE/INSERT, so repeat
runs are safe by design).

## The rules the package enforces

- **The numeric key**: `acct_key = cast(try_cast(id AS bigint) AS string)`
  on every source. Why: the call table zero-pads some account ids; the old
  string key dropped 1,942 real January callers (2,220 calls). Every source
  gets a key-shape probe BEFORE its keys are used; a probe miss = STOP.
- **Assert-exact vs measure-then-lock**: population anchors (204,323 /
  15,177 / 189,146 / $457,943,987 within $5 / 724,848 + class split) and the
  SAS-native waterfall (610,183 / 202,479 / 186,848 / 186,013) are raising
  asserts, forever. Every caller-side number is MEASURED on the first run
  (EXPECTED entry = None), verified from screenshots, then written into
  EXPECTED in B00 (re-pasted into every B file); the second run asserts.
  The old string-key values (9,389 / 11,262 / 29,114) are historical
  references only and are never asserted.
- **The gates**: `captured_sas` (headline) = negative PAYMT_AMT in M1 or M2,
  ACCOUNT grain, month grain (the CQ-7 sign convention, confirmed by A's
  probe). The AWS day-grain 30-day gate survives only as `aws_captured`
  diagnostics and A's one-time delta table, never as a denominator.
- **The loop cut**: the pipeline never reads the CSV's call_type_* columns
  (they are AWS-origin). The caller flag is `inb_native` = the account has
  any January INBOUND id-resolved call row under the numeric key, tied out
  once against the CSV flag, after which the CSV flag lives only in frozen
  notebook A. Future vintage exports need SAS-native fields only.
- **Never merged**: the caller constructs (string-key 9,389 historical /
  numeric-key measured / CSV flag 11,136 / call-day stream / statement-window
  19,789) each keep their definition sentence. 186,013 (export replication)
  and 186,412 (SAS-recorded slice) are quoted side by side, never asserted
  equal. Per-group money is never added down a column; money collapses to
  one row per (group, account) first.

## What the new tables let you explore

- `uc2_t16_01s_populations_<ym>` (610,183 rows): the full export typed - the
  waterfall flags, DLNQT codes M1-M3, EOP balances, CR limits, monthly and
  windowed GROSS_LOSS / CHRGOFF / PLCY_LOSS, the ECL / ECL_12MO / ECL_LIFTM
  and STG_CD / WRITE_OFF families M0-M4, HRAM flags, `captured_sas`,
  `inb_native`, and the `aws_` diagnostic columns. Any slice of the SAS
  population priced in the client's own columns is one GROUP BY away.
- `uc2_t16_04s_outcomes_<ym>` (episode grain, SAS ledger): language groups,
  signal flags, call-day position, the captured_sas caller classes,
  leaked_sas, W_s, with the spine dollars repeated per row.
- `uc2_t16_00n..04n`: the fixed-key AWS layers - same shapes as the round-10
  tables plus `had_zero_pad` (which call rows the old key lost) and `has_tx`.
- `uc2_capture_delta_202501`: the ONLY place the two capture gates meet
  (construct x aws day-grain x sas month-grain cross-tab).
- `uc2_gap1942_202501` / `uc2_sasflag_202501`: the recovered-caller list and
  the CSV-flagged set, persisted for reconciliation.
- `uc2_copilot_excerpt_map` / `uc2_copilot_labels` (B06): the assistant
  campaign's data layer - which excerpt id is which contactid/account, and
  every parsed response line (wave, batch, prompt type, excerpt id, field,
  value, provenance, parse status). Blocker rates with intervals, second-read
  agreement, and dollar weighting are all SQL joins from here to 04s/01s; no
  number ever comes from a reading session.
- `uc2_run_metrics`: every chk and measured value from every run
  (run_ts, notebook, name, value, expected, status, vintage) - diff any two
  runs on-platform, e.g.
  `SELECT name, value FROM uc2_run_metrics WHERE notebook = 'B02_keyfix_aws_layers' ORDER BY run_ts DESC, name`.

## Output conventions (screenshots are the transfer channel)

Grids print as lossless markdown with a running index (40-row chunk banners);
wide one-row pulls print transposed (metric = value); every notebook ends
with a RECORD BLOCK (screenshot THIS - it is the transcription source) and a
`uc2_run_metrics` append; a failing chk prints its context grid before
raising; section banners carry elapsed seconds. B04's excerpt grid is the one
display() output: download it as CSV, never screenshot excerpt content.

## The lock and freeze procedure (after each verified sitting)

1. Screenshots verified against the pre-registered stop rules.
2. Measured values written into `EXPECTED["<vintage>"]` in `B00_setup.py`;
   the SETUP block re-pasted into every B file.
3. Notebook A only: fill every None in its own EXPECTED, pin the explicit
   all-string CSV schema from the census, set `FROZEN = True`, re-run clean.
4. Record + commit. The round-10 tables (`uc2_t16_00..04`, no suffix) and
   `waterfall_coll_call_enriched` are never overwritten: they are the dated
   record's substrate.
