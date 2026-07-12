# Lexicon v2 query set (tier 13) — run protocol and expected tie-outs

Sprint deliverable D1, 2026-07-12. Four console-pasteable queries. Drafted under
sprint/; copy into the kit (`sql/tier13/`) after keeper audit. Every number below
is copied from a verified record; nothing is estimated.

## What lexicon v2 is (and is not)

| Element | v2 treatment | Source |
|---|---|---|
| Deceased-estate group | TOP-priority routing category, m4's measured five phrases verbatim: `passed away`, `death certificate`, `executor`, `deceased`, `calling on behalf` | m2 mining + m4 validation (614 eps / 576 accts, partition exact) |
| Execution markers | New FLAG (column), customer utterances: `bank routing`, `routing number`, `check number`, `checkbook`, `a check for`, `that check`, `on the check` | m1 capture side, measured 5-11x in class b |
| Intent lexicons (promise / payment / plan / hardship / dispute) | UNCHANGED. Cleaning happens via the deceased routing (223 pay-talk + 56 promise episodes in the ledger were estate calls, not intent). No mined phrase passes the acceptance bar as an intent addition; new intent phrases come only through the D4 Copilot loop, validated per protocol §4 | b10 lexicons verbatim; m4 displacement |
| Capture/leak gate | FROZEN. The paid-30d clean gate is unchanged everywhere; execution markers never modify it | cohort picture §6 |
| Extension candidates | Measured as increments in lx3, adopted only past the acceptance bar (support >= 40 class-b episodes AND capture lift >= 2x or CO12 spread >= 10pts) | copilot protocol §4 |

Excluded from the execution flag by design (garble / ambiguous / claim-not-execution;
all stay Copilot-round candidates): `yes you do`, `the the late`, `lost in`,
`ok bank`, `received the bill`, `fee charge`, `january payment`, `the mail today`,
`in english`, `i mailed it`, `paper statement`, `paid it on`, `think my`,
`understand why i`, `january the`, `statement so i`, `i sent you`, `check or`,
`m calling on`, `away on`, `away in`, `report that`, `the routing` (bare; covered
in context by `bank routing`/`routing number`).

## Run protocol (rider-bound)

1. Run lx1 alone. Screenshot the full grid (~54 rows max: 18 block-A rows +
   <= 36 block-B rows; if scrolled, overlap a few rows per b14's stitch
   convention). Route through P12 transcription; orchestrator verifies every
   tie-out below BY HAND; keeper re-verifies.
2. STOP RULE (keeper rider 1): if block A's intent-group episode sums per class
   do not reproduce 45,408 / 5,537 / 2,462, or ANY b10 cell drifts, nothing else
   runs until diagnosed.
3. lx2 and lx3 run only after lx1 passes. Verify their tie-outs.
4. lx4 runs last, only after all January tie-outs are exact (keeper condition:
   calibration exact before the funnel window is touched).
5. Optional: CSV downloads into `output\csv\` for the standing import; not
   required for verification.

Troubleshooting note: lx3 uses `\b` word boundaries (so `he passed` does not fire
inside `she passed`). Athena's regexp_like (Joni/Java syntax) supports `\b`; if
the console errors on it, replace `\b` with `(^|[^a-z])` — same effect here.

## lx1_jan_lexicon_v2_partition — expected tie-outs

Population: b10's classes a/b/c, January 2025 episodes. Class labels: a =
month-MAX B1 entrant cured by EOM (month-max grain); b = entrant still DQ1 at
EOM; c = EOM bucket-1 stock non-entrant (month-end grain).

Block A (v1 marginal) must equal the recorded b10 table cell for cell
(episodes, accounts, pct paid-30d, pct CO12):

| v1 group | class a eps/accts/%cap/%co12 | class b | class c |
|---|---|---|---|
| a. future-dated promise | 9,273 / 8,978 / 94.7 / 2.0 | 1,094 / 1,055 / 67.4 / 15.9 | 480 / 461 / 84.2 / 25.8 |
| b. payment talk, no promise | 35,619 / 32,408 / 89.8 / 2.4 | 4,171 / 3,712 / 60.1 / 22.6 | 1,712 / 1,570 / 80.7 / 26.2 |
| c. plan or settlement talk | 516 / 502 / 91.1 / 8.2 | 272 / 260 / 38.6 / 48.5 | 270 / 255 / 78.5 / 41.6 |
| d. hardship talk | 147 / 146 / 83.7 / 10.3 | 73 / 69 / 34.2 / 44.9 | 21 / 20 / 47.6 / 65.0 |
| e. dispute or fraud talk | 879 / 814 / 72.8 / 4.2 | 287 / 252 / 36.9 / 29.4 | 48 / 42 / 39.6 / 35.7 |
| f. no payment-related language | 13,201 / 12,260 / 85.1 / 3.6 | 3,116 / 2,764 / 46.9 / 30.0 | 1,000 / 934 / 74.0 / 26.8 |

Block A summations (the b8 re-derivation; rider 1 is item 3):

| Check | class a | class b | class c |
|---|---:|---:|---:|
| 1. episodes sum over all 6 groups (b8 stage 2) | 59,635 | 9,013 | 3,531 |
| 2. no-transcript episodes sum (all inside group f); class total minus it = b8 stage 3 | 2,283 (-> 57,352) | 484 (-> 8,529) | 268 (-> 3,263) |
| 3. intent groups a+b+c episode sum (b8 stage 4) — STOP RULE | 45,408 | 5,537 | 2,462 |
| 4. eps_no_pay30d sum over groups a+b+c (b8 stage 5) | 4,179 | 2,187 | 465 |
| 5. eps_no_pay30d_co8m sum over groups a+b+c (b8 stage 6, episode cells) | 41 | 801 | 174 |

Block B (v2 x v1 displacement), classes b+c summed (the ledger = m4's
population; episode counts are additive, account counts are NOT — the m4
accounts figures 576/210 stay m4's, not re-derived here):

| Check | Expected |
|---|---|
| Deceased episodes pulled from v1 group f | 293 |
| ... from v1 group b (payment talk) | 223 |
| ... from v1 group a (promise) | 56 |
| ... from v1 group c (plan) | 28 |
| ... from v1 group e (dispute) | 9 |
| ... from v1 group d (hardship) | 5 |
| Deceased total (b+c) | 614 |
| v2 marginals (b+c): deceased / promise / pay-talk / plan / hardship / dispute / none | 614 / 1,518 / 5,660 / 514 / 89 / 326 / 3,823 |

Consistency arithmetic (holds in the records, must hold in the result):
b10 (b+c) minus displacement = m4: promise 1,574-56=1,518; pay-talk
5,883-223=5,660; plan 542-28=514; hardship 94-5=89; dispute 335-9=326;
none 4,116-293=3,823.

NEW MEASUREMENT (keeper rider 3): the class-a deceased row (the cured class's
deceased-language episodes, month-max grain, episode counts). m4 never scanned
class a. Whatever it reads, report it labeled as new with its grain; do not
fold it into the m4 story.

## lx2_jan_execution_markers — expected tie-outs

| Check | Expected |
|---|---|
| Episodes summed over exec_flag per (class, v2 group) | equals lx1 block B's v2 marginal for that (class, v2 group) exactly |
| Class-b flagged episodes, summed over v2 groups | >= 384 (every `bank routing` episode fires the flag; m1: 336 captured + 48 leaked). Lower bounds also hold per covered family: routing number >= 211-family cells, check number >= 103 |
| Direction | flagged episodes skew heavily captured (m1 lifts 5-11x); a flat or inverted read = stop and diagnose the regex |

No exact upper bound exists (the flag's substrings are broader than m1's exact
phrases); the class-b flagged count is a NEW measured cell.

## lx3_deceased_lexicon_increment — expected tie-outs

| Check | Expected |
|---|---|
| Base rows (label a) per class | equal lx1 block B's deceased episode totals per class exactly; classes b+c sum to 614 |
| Increment rows | NEW measurement; no recorded anchor. Per-candidate rows overlap; only the union row (label h) is additive with the base |
| Missing rows | a candidate with zero increment in a class produces no row; that is a result, not an error |

Acceptance bar for adopting any candidate into the v2 routing group: support
>= 40 class-b episodes AND (capture lift >= 2x either direction OR CO12 spread
>= 10 points vs the class base rate). Below bar = recorded, not used.

## lx4_funnel_v2_monthly — expected tie-outs

The first eight columns must reproduce the recorded f2 digest rows EXACTLY
(raw-payment gate; do not compare no_payment_30d to f1's clean-gate 142,982):

| call_month | episodes | matched | delinquent | with_transcript | pay_language | no_payment_30d | chargeoff_8m |
|---|---:|---:|---:|---:|---:|---:|---:|
| 2024-07-01 | 776,873 | 719,899 | 29,586 | 26,072 | 18,039 | 5,613 | 891 |
| 2024-08-01 | 883,039 | 823,541 | 84,191 | 80,123 | 61,804 | 11,174 | 3,044 |
| 2024-09-01 | 991,093 | 921,891 | 108,428 | 103,810 | 79,362 | 12,188 | 3,556 |
| 2024-10-01 | 1,009,253 | 929,188 | 120,854 | 115,407 | 86,763 | 13,659 | 3,832 |
| 2024-11-01 | 923,473 | 849,321 | 102,585 | 97,950 | 74,610 | 11,980 | 3,379 |
| 2024-12-01 | 982,490 | 902,372 | 111,147 | 102,686 | 77,267 | 12,389 | 3,359 |
| 2025-01-01 | 1,000,739 | 921,499 | 122,468 | 117,166 | 90,268 | 14,114 | 3,881 |
| 2025-02-01 | 884,416 | 815,974 | 105,897 | 101,606 | 78,398 | 11,544 | 3,355 |
| 2025-03-01 | 978,874 | 899,427 | 117,156 | 112,767 | 87,912 | 13,848 | 3,858 |
| 2025-04-01 | 909,122 | 838,662 | 89,225 | 86,051 | 66,367 | 11,111 | 3,470 |
| 2025-05-01 | 926,736 | 853,904 | 96,229 | 88,798 | 69,404 | 11,487 | 3,520 |
| 2025-06-01 | 928,296 | 857,219 | 99,131 | 93,374 | 72,020 | 11,002 | 3,384 |

Cross-check: no_payment_30d sums to 140,109 over the twelve rows (= x1's
raw-gate leaked, on record).

Internal consistency on the new columns, every month: pay_language minus
lx4_pay_language_net_dec <= lx4_deceased_eps; no_payment_30d minus
lx4_no_payment_30d_net_dec <= lx4_deceased_eps; lx4_leaked_with_exec <=
lx4_exec_eps. The six lx4_ columns are NEW measurement (delinquent-in-month
grain, month-max, verified-joined floors: ~28% January hole per b18, ~21%
recent months). The 2024-07 row is a recorded boundary artifact: quote nothing
from it.

Note: lx4's January no_payment_30d (14,114, delinquent-in-month, raw gate) is
NOT the cohort leak list (2,177/1,967, month-end, clean gate, exclusive rule).
Different population, gate, and grain; never mix them in one table.

## Kit binding snippets (for after audit)

manifest.json additions (tier 13):

```json
{
  "id": "lx1_jan_lexicon_v2_partition",
  "tier": 13,
  "file": "tier13/lx1_jan_lexicon_v2_partition.sql",
  "title": "Lexicon v2 calibration: v1-to-v2 displacement (Jan 2025)",
  "question": "Does the v2 partition reproduce b10/b8 exactly, and where does the deceased routing pull episodes from?",
  "render": "table",
  "columns": ["lx1_block", "lx1_v2_group", "lx1_v1_group", "lx1_bridge_class",
    "lx1_episodes", "lx1_accounts", "lx1_eps_paid30d", "lx1_eps_no_pay30d",
    "lx1_eps_no_pay30d_co8m", "lx1_eps_no_transcript", "lx1_co8m_accounts",
    "lx1_co12m_accounts", "lx1_pct_paid30d", "lx1_pct_co12m"],
  "story": "context"
},
{
  "id": "lx2_jan_execution_markers",
  "tier": 13,
  "file": "tier13/lx2_jan_execution_markers.sql",
  "title": "Execution-marker flag on the January cohort",
  "question": "How much in-call payment-mechanics language exists per class and v2 group, and how strongly does it mark capture?",
  "render": "table",
  "columns": ["lx2_bridge_class", "lx2_v2_group", "lx2_exec_flag",
    "lx2_episodes", "lx2_accounts", "lx2_eps_paid30d", "lx2_eps_no_pay30d",
    "lx2_co8m_accounts", "lx2_co12m_accounts", "lx2_pct_paid30d"],
  "story": "context"
},
{
  "id": "lx3_deceased_lexicon_increment",
  "tier": 13,
  "file": "tier13/lx3_deceased_lexicon_increment.sql",
  "title": "Deceased-lexicon extension candidates, measured as increments",
  "question": "How many episodes would each mined candidate phrase add beyond the m4 base list, and do they behave like estate calls?",
  "render": "table",
  "columns": ["lx3_candidate", "lx3_bridge_class", "lx3_episodes",
    "lx3_accounts", "lx3_eps_paid30d", "lx3_eps_no_pay30d",
    "lx3_co8m_accounts", "lx3_co12m_accounts", "lx3_pct_paid30d",
    "lx3_pct_co12m"],
  "story": "context"
},
{
  "id": "lx4_funnel_v2_monthly",
  "tier": 13,
  "file": "tier13/lx4_funnel_v2_monthly.sql",
  "title": "The monthly funnel with v2 columns (deceased routing, execution flag)",
  "question": "What does the funnel series look like with estate calls routed out and execution language counted?",
  "render": "table",
  "columns": ["call_month", "episodes", "matched", "delinquent",
    "with_transcript", "pay_language", "no_payment_30d", "chargeoff_8m",
    "lx4_deceased_eps", "lx4_pay_language_net_dec",
    "lx4_no_payment_30d_net_dec", "lx4_chargeoff_8m_net_dec",
    "lx4_exec_eps", "lx4_leaked_with_exec"],
  "story": "context"
}
```

explains.md additions:

```
## lx1_jan_lexicon_v2_partition

Window: January 2025, the DQ1 cohort's episodes (b10's classes a/b/c). Why: the
calibration gate for lexicon v2 - block A must reproduce b10 and b8 exactly
before any v2 number is used; block B shows exactly which episodes the deceased
routing pulls out of each v1 group. How to read: block A is the old partition
(if any cell drifts from the record, stop); block B rows where the v2 group is
deceased are the routing, everything else is identity. Caveat: account counts
do not sum across cells; episode counts do. The class-a deceased row is a new
measurement, not part of the m4 story.

## lx2_jan_execution_markers

Window: same population as lx1. Why: the mined capture markers (routing/check
mechanics) applied as a flag - the measured footprint of payments being
executed IN the call, the live-assist case. How to read: within each class and
v2 group, compare paid-30d rates with the flag present vs absent; class b is
where the m1 lifts (5-11x) were mined. Caveat: the flag never changes the
capture gate; flagged-but-leaked episodes are gate-false-negative candidates
(mailed checks), not proof of payment.

## lx3_deceased_lexicon_increment

Window: same population as lx1. Why: the mined phrases the m4 five-phrase base
list does not cover, each measured as the episodes it would ADD - the
acceptance-bar evidence for extending the routing lexicon. How to read: the
base row ties to lx1; increment rows show volume and behavior (paid-30d, CO12)
of the added episodes - estate-like behavior is low capture, high CO12.
Caveat: per-candidate rows overlap; only the union row is additive. Adoption
requires the protocol acceptance bar, not just a lift.

## lx4_funnel_v2_monthly

Window: W3 call months (2024-07..2025-06), delinquent-in-month episodes,
verified-joined only. Why: the funnel series with the v2 read attached -
how much of each month's leaked count is estate processing rather than
collectable leakage, and how much execution language the stream carries.
How to read: first eight columns are f2 verbatim (raw gate) and must match the
07-04 digest; the net-of-deceased columns are the routed funnel. Caveat:
2024-07 is a boundary artifact month; counts are floors (~21-28% unverified
inbound); this is not the cohort leak list - different gate, grain, population.
```

## Verification (design time)

Checked: all four queries derive their chains verbatim from verified kit SQL
(b10, m4, f2 read in full before writing; capture gate, cohort CTEs, windows,
and regexes copied character for character; the only additions are the
deceased/exec count_ifs, the v2 CASE, and output columns). The displacement
arithmetic was re-derived from the records and is internally consistent
(b10 minus m4 displacement equals m4's v2 marginals on all six groups). All
four are one-transcript-pass, aggregates-only builds (b11's lesson); lx1's
two-block UNION ALL follows b8's proven multi-reference pattern; lx3's UNNEST
fan-out runs over the small per-episode table only.

[OPEN] (run time, owner Ravi + orchestrator verification + keeper audit):
every tie-out above. Nothing from lx1-lx4 is quotable until the January
calibration reproduces the records exactly.
