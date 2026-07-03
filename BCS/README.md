# BCS — Athena data story (accounts × calls × transcripts)

A tiered, self-contained analysis kit for three Athena tables:

| Table | Grain |
| --- | --- |
| `fmt_acct_dba.fmt_acct_c` | card account × monthly snapshot (`sfx_nbr = 0` filter) |
| `contactcenter_bdp_db.call` | one row per call leg |
| `contactcenter_bdp_db.transcript` | one row per spoken utterance |

It runs 35 aggregate queries in four tiers and renders everything into **one
HTML file** (`output/bcs_story.html`) — charts, KPI tiles, data tables, and
every query verbatim in an appendix. Two ways to get the results in:

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
├─ src/                    athena runner, check, fetch, csv import, report (stable)
├─ sql/                    ONE FILE PER QUERY + manifest.json  <- iterate here
└─ output/                 gitignored: JSONs + bcs_story.html
   └─ csv/                 manual mode: drop console CSV downloads here
```

The point of the split: `config/` and `src/` are a one-time copy into the
environment. Day-to-day iteration happens in `sql/` — every query is a small
literal `.sql` file you can paste into the Athena console unchanged, and
`sql/manifest.json` binds it to a tier, a title, and a chart type. Swapping a
broken query = replacing one small file.

## The four tiers

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
   python run_all.py --import --report
   ```

Imports **merge**: run two queries today and five tomorrow, re-import any
time — nothing already imported is lost, and a newer download of the same
query replaces the older one. Queries without a CSV yet appear in the report
as dashed "not run yet" cards naming the next file to paste. Unrecognized
CSVs are listed as unmatched, never guessed.

If a query errors in the console on a missing column (the deeper `past_due_*`
fields), paste its `_fallback.sql` variant instead — the importer assigns the
result to the right query either way.

Suggested paste order (manifest order — tick them off):

```text
Tier 1: t1_acct_size, t1_call_size, t1_transcript_size,
        t1_initiation_mix, t1_participants, t1_sentiment_mix
Tier 2: t2_monthly_volume, t2_split_producttype, t2_split_vendor,
        t2_split_site, t2_split_queue, t2_split_auth, t2_split_transfer,
        t2_handle_time, t2_abandon_transfer_monthly, t2_sentiment_monthly,
        t2_utterances_per_call, t2_call_minutes, t2_dpd_buckets,
        t2_chargeoff_trend
Tier 3: t3_match_rate, t3_transcript_coverage, t3_repeat_callers,
        t3_caller_dpd, t3_payment_language, t3_conversation_arc,
        t3_first_last_speaker
Tier 4: v1_dq1_call_concentration, v2_vintage_roll, v3_caller_vs_noncaller,
        v4_balance_at_risk, v5_payment_after_call, v6_stage_proxy,
        v7_ib_ob_mix, v8_reage_proxy
```

If console time is short, the highest-value dozen: the three `t1_*_size`
queries, `t1_initiation_mix`, `t1_participants`, `t3_match_rate`,
`t3_transcript_coverage`, `t2_dpd_buckets`, `t2_chargeoff_trend`,
`v1_dq1_call_concentration`, `v2_vintage_roll`, `v5_payment_after_call`.

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
   render (`kpis` | `bars` | `line` | `table`) and its column hints.
3. `run.cmd --fetch` then `run.cmd --report`.
