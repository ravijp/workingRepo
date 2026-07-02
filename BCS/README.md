# BCS — Athena data story (accounts × calls × transcripts)

A tiered, self-contained analysis kit for three Athena tables:

| Table | Grain |
| --- | --- |
| `fmt_acct_dba.fmt_acct_c` | card account × monthly snapshot (`sfx_nbr = 0` filter) |
| `contactcenter_bdp_db.call` | one row per call leg |
| `contactcenter_bdp_db.transcript` | one row per spoken utterance |

It checks the AWS connection, runs ~34 aggregate queries in four tiers, and
renders everything into **one HTML file** (`output/bcs_story.html`) — charts,
KPI tiles, data tables, and every query verbatim in an appendix.

## Layout (what changes vs what doesn't)

```text
BCS/
├─ run_all.py              orchestrator (stable)
├─ run.cmd                 cmd wrapper: creds + venv + run (stable)
├─ creds.local.cmd.example credential template -> copy to creds.local.cmd
├─ config/settings.py      region/workgroup/paths, env-overridable (stable)
├─ src/                    athena runner, check, fetch, report (stable)
├─ sql/                    ONE FILE PER QUERY + manifest.json  <- iterate here
└─ output/                 gitignored: JSONs + bcs_story.html
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
