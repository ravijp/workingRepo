# Validation tier (v-series): what each query settles

Two candidate sizing approaches exist for reading risk and call value from
these tables:

- **Backward frame:** start from charged-off accounts, look back at their
  delinquent-month inbound calls, and call the balance at DQ1 "at risk".
  Weakness: it conditions on the loss outcome. The decision population at call
  time is *all* delinquent callers, and nobody knows at that point who will
  charge off. It also overstates risk: most DQ1 balance cures.
- **Forward frame:** start from a vintage of accounts *entering* delinquency,
  overlay their calls, and follow outcomes forward (cure, roll, charge-off).
  Loss can then be expressed provision-style (stage migration) rather than
  waiting for charge-off, which lands months later.

The v-series measures, from the three tables alone, everything that choice
depends on. Run order does not matter; each query is independent.

| Query | Assumption it tests | Decision it informs |
| --- | --- | --- |
| `v1_dq1_call_concentration` | "Most inbound calls happen in DQ1" | Whether DQ1-only scoping loses material call volume (caller bucket joined in the month of the call) |
| `v2_vintage_roll` | Stage-migration multipliers (assumed 1x to 5x to 15x to 20x shape) | Replaces assumed multipliers with measured roll/cure rates per month-on-book |
| `v3_caller_vs_noncaller` | "Callers and non-callers differ in outcome" | The forward frame's headline split: charge-off and cure rates by caller group, on the full entrant population |
| `v4_balance_at_risk` | "Balance at DQ1 entry = balance at risk" | How entry dollars actually split between cure / still-delinquent / charge-off |
| `v5_payment_after_call` | "Call with no payment after it = candidate leakage" | Sizes the no-payment-within-30-days pool per bucket (payment proxied from next month-end snapshot) |
| `v6_stage_proxy` | Bucket ladder maps onto provision stages | The stage-1/2/3 proxy shape (accounts + balances) to pair with v2's roll rates for an expected-loss ladder |
| `v7_ib_ob_mix` | "Inbound-only misses most contact" | How much contact volume the inbound lens sees per bucket; keeps outbound separate but sized |
| `v8_reage_proxy` | "Re-age cases need a flag" | How often deep buckets reset straight to current; sizes the re-age question before chasing a flag |

## What these tables cannot answer (needs other sources)

| Gap | Why it matters | Where it lives |
| --- | --- | --- |
| Booked provision stage + expected-loss dollars | Turning the stage proxy into booked impairment impact | Finance / account-summary ledger tables, not these three |
| High-risk / program flags (stage overrides, stickiness window) | Stage-2 membership is not purely bucket-driven | Collections program / status tables |
| Bureau or internal PD scores at vintage month | Risk-adjusting the caller vs non-caller comparison | Score tables / bureau feeds |
| Promise-to-pay terms and kept/broken status | Separating "captured but broken" from "never captured" | Operations PTP dataset |
| Charge-off reversal amounts | Netting the loss number | Ledger tables |
| Deceased / bankruptcy exclusions beyond the status-text codes | Cleaning the special cases out of the population | Status-reason code mapping; `acct_status_rsn_txt` regex (FR/ST/BK) is a partial proxy |

## Reading guide

- v3 is an association, not a causal effect: callers self-select. Use it to
  size and segment, not to claim "calling causes cure".
- v5's payment read is a month-end proxy. Before quoting it, check the
  `paymt_last_dt` parse produced non-null dates (the query comment explains).
- v2/v3/v4 share one cohort definition (bucket 0 to 1 entrants, 11 months
  before the newest snapshot). Change the vintage by editing that one line in
  each file.
