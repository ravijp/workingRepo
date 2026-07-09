# BCS — Athena data story (accounts × calls × transcripts)

A tiered, self-contained analysis kit for three Athena tables:

| Table | Grain |
| --- | --- |
| `fmt_acct_dba.fmt_acct_c` | card account × monthly snapshot (`sfx_nbr = 0` filter) |
| `contactcenter_bdp_db.call` | one row per call leg |
| `contactcenter_bdp_db.transcript` | one row per spoken utterance |

It runs 86 aggregate queries in ten tiers and renders everything into **one
HTML file** (`output/bcs_story.html`) — an executive headline strip, charts,
KPI tiles, data tables, and every query verbatim in an appendix — plus **one
markdown digest** (`output/digest.md`) holding every result as compact
tables, so a whole run can leave the environment in a couple of screenshots.
The story tiers lead the report in walk order (5 funnel, 6 sizing, 9
operations, 10 follow-through, 7 conversation, 8 learned language); each
tier's context cards render collapsed. Two ways to get the results in:

- **API mode** — the kit runs the queries itself via boto3 (needs AWS network
  access from the machine).
- **Manual mode** — you run each query in the Athena console, download the
  CSV, and the kit imports and renders them. No AWS access needed from the
  machine running the kit. See "Manual mode" below.

## Layout (what changes vs what doesn't)

```text
BCS/
├─ run_all.py              orchestrator (stable)
├─ run.cmd                 cmd wrapper: creds + venv + run (stable)
├─ creds.local.cmd.example credential template -> copy to creds.local.cmd
├─ config/settings.py      region/workgroup/paths, env-overridable (stable)
├─ src/                    athena runner, check, fetch, csv import, report, digest (stable)
├─ sql/                    manifest.json + explains.md; ONE FILE PER QUERY  <- iterate here
│  └─ tier1/ .. tier10/    queries grouped by tier (manifest carries the tierN/ path)
└─ output/                 gitignored: JSONs + bcs_story.html + digest.md
   └─ csv/                 manual mode: drop console CSV downloads here
```

The point of the split: `config/` and `src/` are a one-time copy into the
environment. Day-to-day iteration happens in `sql/` — every query is a small
literal `.sql` file you can paste into the Athena console unchanged, and
`sql/manifest.json` binds it to a tier, a title, and a chart type. Swapping a
broken query = replacing one small file.

## The eight tiers

Story tiers (lead the report):

5. **The funnel (f-series)** — the headline waterfall: inbound episodes from
   delinquent accounts that showed payment intent, left without a capture,
   and charged off — episodes, accounts, and dollars at every stage — plus
   its stability (by month, by bucket), the dollar trend, the loss-ledger
   flip, the call-intensity read, two bias gates, and the full-history
   calibration (f0/f0b) that pins the analysis window.
6. **Sizing inputs (s-series)** — the measured numbers a dollar model needs:
   balance per bucket, the leading-edge roll matrix, payment size per
   capture, the caller vs non-caller payment baseline.
7. **Conversation deep-dive (h-series)** — does payment language predict
   payment, who raises payment first, language across the full ladder, call
   effort per bucket.
8. **Learned language (n-series)** — SQL-native NLP on outcomes: bigrams
   learned from paid-vs-leaked calls, the platform's sentiment scores tested
   against payment, a composed intent score, the agent-offer gap, early-intent
   latency.
9. **Where it leaks — operations (o-series)** — payment-after-call by queue,
   vendor × site, transfer paths, and abandons with recontact: leakage as a
   place someone owns.
10. **Follow-through — outcome curves (x-series)** — captured vs leaked three
    months later, repeat-leak chains, time from first leak to charge-off (the
    8-month-horizon check).

Foundation tiers:

1. **Shape of the data** — row counts, date ranges, snapshot freshness, key
   fill rates, who speaks in transcripts.
2. **Operational picture** — volume trends, queue/vendor/site splits, handle
   times, abandons/transfers, sentiment, conversation length, the delinquency
   ladder, the charge-off trend.
3. **Cross-table story** — call→account match rate, transcript coverage of
   inbound calls, repeat callers, caller delinquency profile, payment language
   in customer utterances, sentiment arc across the call.
4. **Approach validation (v-series)** — the checks that pin a sizing approach
   down (see [VALIDATION.md](VALIDATION.md)): where calls sit on the DQ ladder,
   vintage roll/cure curves, caller vs non-caller outcomes, balance-at-risk
   splits, payment-after-call, stage proxies, inbound/outbound mix, re-age
   signal.

Window policy: call-only queries anchor to the call table's newest data;
queries that JOIN calls to accounts anchor to the ACCOUNT table's newest
complete month (the account copy trails the calls, and joins past its edge
silently under-match); the funnel pins literal months chosen for outcome
runway and confirmed by f0. Each card's explainer states its window and why.

## One-time setup (BCS VDI, cmd)

```cmd
cd <where-you-placed-this>\BCS
setup.cmd
```

`setup.cmd` runs `dev activate` itself if python is missing, creates the venv
with `uv` (stdlib `venv` fallback), installs `requirements.txt`, and creates
`creds.local.cmd` from the template.

## Credentials (every session)

Your three AWS keys are session-based. When they refresh, copy the block from
the credentials page (it is already in `set AWS_...=...` form) and paste it
**over the three lines in `creds.local.cmd`**. That file is gitignored.
Pasting the block straight into the cmd window works too.

## Run

```cmd
cd ...\BCS
run.cmd                            :: dev activate (if needed) + creds + venv + all steps
```

or step by step:

```cmd
run.cmd --check                    :: connection check only
run.cmd --fetch                    :: run the queries
run.cmd --report                   :: rebuild the HTML from saved JSON
```

Then open `output\bcs_story.html` in a browser. Failed queries do not stop the
run; they appear in the report with the error and the SQL.

## Manual mode (Athena console + CSV downloads)

When the machine cannot reach AWS endpoints (VDI proxy restrictions), run the
queries in the Athena **query editor** instead and let the kit assemble the
story from the downloaded results:

1. Open a file from `sql/` (on GitHub or locally), paste it into the Athena
   editor, and run it.
2. Click **Download results (.csv)** and drop the file into `output\csv\`
   **as-is — no renaming needed**. Every query has a unique output header, so
   the importer recognizes each CSV automatically (a file renamed to
   `<id>.csv` also works and wins over header matching).
3. Import and render (works even without boto3 installed):

   ```cmd
   python run_all.py --import --report --digest
   ```

   `--digest` writes `output\digest.md` — every imported result as compact
   markdown tables, core queries first. Screenshot that one file to carry a
   whole run out of the environment; the HTML stays for walking the story.

   `--story` writes `output\uc2_story.html` — the act-ordered story report:
   the January 2025 story acts (from `sql/act_map.json`) with narrative
   interleaved and each act's query cards under it. Same imported results as
   the atlas; missing queries render as pending cards, so it fills in as CSVs
   land. The full `bcs_story.html` stays the 88-query atlas.

Imports **merge**: run two queries today and five tomorrow, re-import any
time — nothing already imported is lost, and a newer download of the same
query replaces the older one. Queries without a CSV yet appear in the report
as dashed "not run yet" cards naming the next file to paste. Unrecognized
CSVs are listed as unmatched, never guessed.

If a query errors in the console on a missing column (the deeper `past_due_*`
fields), paste its `_fallback.sql` variant instead — the importer assigns the
result to the right query either way.

Suggested paste order for a fresh run of the NEW tiers (calibration first —
check f0 before trusting the funnel's window):

```text
Calibrate: f0_period_calibration, f0b_method_history
Tier 5:    f1_funnel_waterfall, f2_funnel_by_month, f6_funnel_by_bucket,
           f7_leaked_dollars_by_month, f3_funnel_dollars,
           f5_calls_before_chargeoff, f4_match_by_auth, f4_coverage_by_bucket,
           f8_payment_window_sensitivity, f9_episode_chaining
Gates:     s6_payment_contamination (probes the autopay/NSF columns f1 needs),
           f4_scope_split, f4_bucket_drift
Tier 6:    s1_balance_by_bucket, s2_roll_matrix, s3_payment_size_by_bucket,
           s4_caller_payment_lift, s5_balance_roll_matrix
Tier 7:    h1_language_vs_payment, h2_first_payment_mention,
           h3_language_by_ladder, h4_call_effort_by_bucket,
           h5_triage_language, h6_promise_language
Tier 8:    n1_discriminative_bigrams, n2_sentiment_vs_payment,
           n3_intent_score, n4_agent_offer_vs_outcome, n5_early_intent
Tier 9:    o1_capture_by_queue, o2_capture_by_vendor_site,
           o3_delinquent_transfer_paths, o4_delinquent_abandons,
           o5_capture_by_auth, o6_coll_volume_monthly
Tier 10:   x1_leaked_vs_captured_roll, x2_repeat_leak_chains,
           x3_time_to_chargeoff_after_leak, x4_within_account_contrast,
           x5_payment_latency
Tier 3/4:  t3_unmatched_queues, t3_call_gap_days, v10_cure_durability,
           v11_multi_vintage_roll
Reruns:    t2_dpd_buckets (primary ladder), t3_match_rate, v1, v5, v7, v9
           (now anchored to the account table's clock)
```

The foundation tiers (t1/t2/t3/v-series) keep their prior results — imports
merge, so nothing already imported is lost.

Per-card explainer text (window, why, how to read, caveats) lives in
[sql/explains.md](sql/explains.md) — one `## <query-id>` section per query,
rendered on the card. Edit the prose there; no code or JSON involved.

If console time is short, the story core: `f0_period_calibration`,
`s6_payment_contamination` (the gate probe), `f1_funnel_waterfall`,
`f2_funnel_by_month`, `f8_payment_window_sensitivity`, `f3_funnel_dollars`,
`f4_match_by_auth`, `s2_roll_matrix`, `s4_caller_payment_lift`,
`h1_language_vs_payment`, `x4_within_account_contrast`, `n3_intent_score`.

## Config (env vars, all optional)

| Variable | Default | When to set |
| --- | --- | --- |
| `AWS_DEFAULT_REGION` | `us-east-1` | region differs |
| `ATHENA_WORKGROUP` | `primary` | your workgroup differs |
| `ATHENA_OUTPUT_S3` | (empty) | error mentions "output location": set `s3://bucket/prefix/` |
| `QUERY_TIMEOUT_S` | `600` | big tables, slow queries |
| `MAX_RESULT_ROWS` | `2000` | rarely |

## Safety rules

- **Never commit `output/`** (data extracts, the HTML report) or
  `creds.local.cmd` (keys). Both are gitignored — keep it that way.
- All queries are aggregates with bounded windows; the only row-level pulls
  are the 3-row samples in the connection check (values truncated).
- Windows anchor to each table's newest data, so a stale table copy shifts the
  window instead of returning empty.

## Adding a query

1. Drop `my_query.sql` into `sql/` (fully literal, console-runnable).
2. Add an entry to `sql/manifest.json`: id, tier, file, title, question,
   render (`kpis` | `bars` | `line` | `table`) with its column hints, and
   story (`core` | `context` — context cards render collapsed).
3. `run.cmd --fetch` then `run.cmd --report --digest`.
