# SAS counter-query pack for Ishant — 2026-07-10

Split ownership per the 2026-07-10 decision: Athena owns the funnel + transcript
story; SAS owns the dollar layer (ECL, stages, HRAM, payments). These queries
run on Ishant's side against `zenon.waterfall_dtl_expd_coll_202501` (the wide
vintage table, fields as transcribed in
[records/ishant-sas-waterfall-202501-2026-07-09.md](records/ishant-sas-waterfall-202501-2026-07-09.md)).
We consume **aggregates only** — counts and sums, no account rows. Hand this file
over verbatim. CQ-1 to CQ-3 are the original pack; CQ-4 to CQ-8 were appended
2026-07-10 from the cohort-picture arbitration
([uc2-jan2025-cohort-picture.md](uc2-jan2025-cohort-picture.md)).

Baseline filters, same as every shared pivot (apply unless a query says otherwise):
`COHORT_ENTRY='DPD1PLUS'`, `cpc_M1='others'` (ex-AA),
`CHRGOFF_RSN_M1 in ('blank','PLY')`. Grain: everything here is month-END
(DLNQT_CD is the month-end position).

Note on flag coding: CO_8M/10M/12M_FLAG and call_type_INB are written below as
`='Y'` — adjust to however they are coded in the table (0/1 or Y/N).

## CQ-1 — the caller dollar layer (code-1 x inbound-caller x February outcome, with ECL)

Why: the `call_type_INB` flag is already joined on your side but filtered `(All)`
in every shared pivot. This is the one view your table can produce and Athena
cannot: what the code-1 CALLERS cost in reserve, by where they land at February
month-end. It prices the leak story in ECL.

```sas
proc sql;
  create table zenon.cq1_caller_dollar_layer as
  select DLNQT_CD_M1,
         call_type_INB,
         DLNQT_CD_M2,
         count(*)                                  as accounts,
         sum(ECL_M1)                               as ecl_m1,
         sum(ECL_M2)                               as ecl_m2,
         sum(ECL_M3)                               as ecl_m3,
         sum(EOP_BAL_M1)                           as eop_bal_m1,
         sum(EOP_BAL_M2)                           as eop_bal_m2,
         sum(case when CO_12M_FLAG='Y' then 1 else 0 end) as co_12m
  from zenon.waterfall_dtl_expd_coll_202501
  where COHORT_ENTRY='DPD1PLUS'
    and DLNQT_CD_M1='1'
    and cpc_M1='others'
    and CHRGOFF_RSN_M1 in ('blank','PLY')
  group by DLNQT_CD_M1, call_type_INB, DLNQT_CD_M2
  order by call_type_INB, DLNQT_CD_M2;
quit;
```

Return: the full result (it is ~16 rows). Please also state, one line each:

1. The **definition of `call_type_INB`**: which call table, what window (January
   only or ever?), inbound only or any contact, verified calls only or all? Our
   Athena caller overlay counts only verified inbound (acctid present); if your
   flag has the same verified-only limitation, both sides share the ~21% blind
   spot and we will say so; if not, this flag sees callers we cannot.
2. Whether the same cut can be rerun with filters open (`cpc_M1=(All)`,
   `CHRGOFF_RSN_M1=(All)`) — one screenshot, for the reconciliation margin.

What we do with it: Athena b7/b8/b9 give the caller/leak/charge-off story on
counts; this gives the same callers' ECL and balance. Numbers cross the boundary
only as these reconciled aggregates.

Optional extension (Zenon-internal, data stays in the client environment): we can
hand you the ~1,970-account leak list (still-DQ1 intent-no-payment accounts,
month-end grain) as an id list; `sum(ECL_M1), sum(ECL_M2), sum(EOP_BAL_M1)` on
that join prices the exact work list the story targets. Say the word and we drop
the list in the shared location.

Second optional ask, from the 2026-07-10 transcript mining: our phrase mining shows
the code-1 cohort's charge-off language is dominated by deceased-estate talk
(death certificate, executor, passed away). If the reason code at charge-off is
reachable for the 12-month window (the wide table carries CHRGOFF_RSN only to M3),
the **CHRGOFF_STATUS_RSN_CD distribution of the cohort's CO_12M accounts** — i.e.
the DEC share — would size that segment in your currency. One GROUP BY if the field
is there; skip if it needs a rebuild.

## CQ-2 — the HRAM / non-HRAM split of the rollers' ECL step

Why: the rollers' ECL step ($40.2M M1 → $69.2M M2 on the 65,585) is the story's
roll price, but it blends HRAM and non-HRAM accounts. An HRAM account is stage 2
before it rolls, so its 1→2 move carries no coverage step. Namit and Anupam agreed
this is the next cut (07-08). Without it we will not quote a prevented-roll dollar.

The HRAM flag is not in the waterfall table as transcribed. [OPEN] its source:
HRAM capture is keyed on the Apollo PD model (Anupam 07-08) — Anupam can point to
the table/field that marks HRAM capture per account-month. Once joined (call it
`HRAM_M1`, HRAM-captured as of January):

```sas
proc sql;
  create table zenon.cq2_hram_roll_split as
  select HRAM_M1,
         STG_CD_M1,
         count(*)          as accounts,
         sum(ECL_M1)       as ecl_m1,
         sum(ECL_M2)       as ecl_m2,
         sum(ECL_M3)       as ecl_m3,
         sum(EOP_BAL_M1)   as eop_bal_m1,
         sum(case when CO_12M_FLAG='Y' then 1 else 0 end) as co_12m
  from zenon.waterfall_dtl_expd_coll_202501   /* + HRAM join */
  where COHORT_ENTRY='DPD1PLUS'
    and DLNQT_CD_M1='1' and DLNQT_CD_M2='2'   /* the 65,585 rollers */
    and cpc_M1='others'
    and CHRGOFF_RSN_M1 in ('blank','PLY')
  group by HRAM_M1, STG_CD_M1;
quit;
```

Return: the grid (HRAM x stage, ~8 rows). The read we need: the non-HRAM share of
the +$29.0M step (that is the preventable coverage step), and whether the S2 rows
are mostly HRAM (which would explain stage 2 pre-roll).

Anupam's suggested extension, second priority: among January HRAM captures that
CURED, what share exits HRAM after the ~12-month monitoring period (that exit is
itself opportunity). Needs HRAM status ~202601; only if the source table reaches
that far.

## CQ-3 — the A-068 stage-timing check (decides the embargoed claim)

Why: 10,898 of the 28,014 stage-1 rollers are still S1 at M2 despite sitting at
bucket 2 — Anupam says that should be impossible and suspects the impairment
table's estimation used a different (late-cycle) bucket. Until this check closes,
the "stage blindness" claim stays out of the story. Direction from Anupam (07-08):
check the impairment table's own DQ variable; balance plays no role.

Step 1 — name the variable: list the columns of `imprmnt_monthly` and identify the
delinquency/bucket variable used at estimation time (the transcription of your
build only maps WGHTD_EXPTD_LOSS_AMT, ECL_12MO, ECL_LIFTM, STG_CD, WRITE_OFF).
[OPEN] its name; call it `IMP_DQ` below. Also note whether the table carries a
cycle date.

Step 2 — cross-tab the S1-stayers:

```sas
proc sql;
  create table zenon.cq3_s1_stayers_dq as
  select w.STG_CD_M2,
         i.IMP_DQ            as imp_dq_at_m2,
         w.DLNQT_CD_M2,
         count(*)            as accounts,
         sum(w.ECL_M2)       as ecl_m2,
         mean(w.EOP_BAL_M2)  as avg_bal_m2
  from zenon.waterfall_dtl_expd_coll_202501 w
  join imprmnt_monthly i
    on /* account key */ and i.month=202502
  where w.COHORT_ENTRY='DPD1PLUS'
    and w.DLNQT_CD_M1='1' and w.DLNQT_CD_M2='2'
    and w.STG_CD_M1='S1'
    and w.cpc_M1='others'
    and w.CHRGOFF_RSN_M1 in ('blank','PLY')
  group by w.STG_CD_M2, i.IMP_DQ, w.DLNQT_CD_M2;
quit;
```

The decisive read: for the S1→S1 rows (10,898 accounts), does `IMP_DQ` at 202502
show bucket 0/1 (estimation saw a pre-roll position → timing artifact, claim dies)
or bucket 2 (estimation saw the roll and stage still did not move → the claim
lives, subject to the customer-level-staging caveat)?

Step 3 — the 2,084 missing-ECL rollers: profile them in one cut — account open
month, EOP_BAL_M1, cpc_M1, and whether an imprmnt_monthly row exists at all for
202501 (Anupam speculated fresh accounts).

Return: the step-2 grid, the step-1 variable name (+ cycle-date yes/no), and the
step-3 profile. This closes A-068 items 1-2; CQ-2 closes item 3 (the HRAM
breakdown).

## CQ-4 — the four-row ledger priced at M2, HRAM-split

Why: extends CQ-2 from the rollers to all transitions. Your pivot's M2 rows sum to
105,715 + 15,002 + 65,585 = 186,302 of 186,412 (account grain, month-end); this
grid prices the cure row's deferred release, the stay row, and the exit row, and
names the 110-account remainder. HRAM_M1 as in CQ-2: the flag is not in the
waterfall table as transcribed, source [OPEN] (HRAM capture is keyed on the Apollo
PD model, Anupam 07-08).

```sas
proc sql;
  create table zenon.cq4_ledger_priced_m2 as
  select DLNQT_CD_M2,   /* every value, missing included: the missing rows name the 110 */
         HRAM_M1,
         count(*)                                  as accounts,
         sum(ECL_M1)                               as ecl_m1,
         sum(ECL_M2)                               as ecl_m2,
         sum(ECL_M3)                               as ecl_m3,
         sum(EOP_BAL_M1)                           as eop_bal_m1,
         sum(EOP_BAL_M2)                           as eop_bal_m2,
         sum(case when CO_12M_FLAG='Y' then 1 else 0 end) as co_12m
  from zenon.waterfall_dtl_expd_coll_202501       /* + HRAM join, same [OPEN] source as CQ-2 */
  where COHORT_ENTRY='DPD1PLUS'
    and DLNQT_CD_M1='1'
    and cpc_M1='others'
    and CHRGOFF_RSN_M1 in ('blank','PLY')
  group by DLNQT_CD_M2, HRAM_M1
  order by DLNQT_CD_M2, HRAM_M1;
quit;
```

Return: the grid. Tie-out: the account column must reproduce 105,715 (M2 code 0) /
15,002 (code 1) / 65,585 (code 2) and name the 110. Two reads we need, one line
each:

1. The per-account ECL difference at M2 between the cure row and the roll row:
   that is the unit value of a prevented roll in the client's currency.
2. The cure row's ECL_M2/M3: whether cure releases reserve inside a pilot's
   measurement window (your pivot shows cured ECL_M1 $34.8M and no M2 cell).
   Stage-mechanics interpretation waits on CQ-3.

## CQ-5 — price the leak list

Why: promotes CQ-1's optional id-list extension to a required item. The leak list,
about 2,394 accounts (month-end grain; intent-no-payment, pre-routing; the two
month-end classes' b8 cells, 1,970 + 424), is the work product, and it needs its
own ECL and balance before any value line is written.

Mechanics: Zenon drops the account-id list in the shared location; data stays in
the client environment.

```sas
proc sql;
  create table zenon.cq5_leak_list_priced as
  select count(*)                                  as accounts_joined,
         sum(w.ECL_M1)                             as ecl_m1,
         sum(w.ECL_M2)                             as ecl_m2,
         sum(w.EOP_BAL_M1)                         as eop_bal_m1,
         sum(case when w.CO_12M_FLAG='Y' then 1 else 0 end) as co_12m
  from zenon.waterfall_dtl_expd_coll_202501 w
  join zenon.leak_list_202501 l                    /* the dropped id list */
    on /* account key */
  where w.COHORT_ENTRY='DPD1PLUS'
    and w.DLNQT_CD_M1='1'
    and w.cpc_M1='others'
    and w.CHRGOFF_RSN_M1 in ('blank','PLY');
quit;
```

Return: the one aggregate row: sum(ECL_M1), sum(ECL_M2), sum(EOP_BAL_M1), CO_12M
count on the join. Run twice once the deceased-routing query (m4) lands on our
side: with and without the deceased-flagged accounts.

## CQ-6 — the customer key (one-liner)

Why: IFRS 9 staging is customer-level; the Athena side has no customer id.

Ask: does the waterfall table (or ASP) carry a customer identifier, and how many
code-1 accounts (month-end, baseline filters) share a customer with another
delinquent account? If the key exists (call it `CUST_ID`, name [OPEN]):

```sas
proc sql;
  create table zenon.cq6_customer_key as
  select count(*) as accts_sharing_customer
  from zenon.waterfall_dtl_expd_coll_202501 w
  where w.COHORT_ENTRY='DPD1PLUS'
    and w.DLNQT_CD_M1='1'
    and w.cpc_M1='others'
    and w.CHRGOFF_RSN_M1 in ('blank','PLY')
    and w.CUST_ID in (select CUST_ID
                      from zenon.waterfall_dtl_expd_coll_202501
                      group by CUST_ID
                      having count(*) > 1);
quit;
```

Return: yes/no on the key, and the one count. Decides whether the work list needs
customer-grain dedup and bounds the customer-level-staging disclosure.

## CQ-7 — true-payment check of the paid-30d gate

Why: every capture/leak label on the Athena side rests on a last-payment-date
proxy with autopay and bounced exclusions. PAYMT_AMT (negative = payment) is the
true payment record and can validate the gate itself.

Mechanics: Zenon drops the caller account-id list, each account carrying its
Athena paid-30d label (captured / leaked), in the shared location; data stays in
the client environment.

```sas
proc sql;
  create table zenon.cq7_paid30d_check as
  select l.paid_30d_label,
         count(*)                                  as accounts,
         sum(w.PAYMT_AMT_M1)                       as paymt_amt_m1,
         sum(w.PAYMT_AMT_M2)                       as paymt_amt_m2,
         sum(case when w.PAYMT_AMT_M1 < 0 or w.PAYMT_AMT_M2 < 0
                  then 1 else 0 end)               as accts_any_payment
  from zenon.waterfall_dtl_expd_coll_202501 w
  join zenon.caller_paid30d_list l                 /* the dropped id + label list */
    on /* account key */
  where w.COHORT_ENTRY='DPD1PLUS'
    and w.DLNQT_CD_M1='1'
    and w.cpc_M1='others'
    and w.CHRGOFF_RSN_M1 in ('blank','PLY')
  group by l.paid_30d_label;
quit;
```

Return: per label group: count, sum(PAYMT_AMT_M1), sum(PAYMT_AMT_M2), and
accounts with any negative PAYMT_AMT in M1/M2. The question: do the two sides
agree on who actually paid, and how large is the disagreement.

## CQ-8 — the unfiltered-perimeter ECL path

Why: every ECL figure quoted so far sits on the baseline filters (ex-AA, reasons
blank+PLY). The perimeter disclosure needs a dollar size, not just account counts
(15,120 AA accounts, 947 reason-excluded; both month-end grain).

```sas
proc sql;
  create table zenon.cq8_unfiltered_perimeter as
  select cpc_M1,
         count(*)                                  as accounts,
         sum(ECL_M1)                               as ecl_m1,
         sum(ECL_M2)                               as ecl_m2,
         sum(ECL_M3)                               as ecl_m3,
         sum(EOP_BAL_M1)                           as eop_bal_m1
  from zenon.waterfall_dtl_expd_coll_202501
  where COHORT_ENTRY='DPD1PLUS'
    and DLNQT_CD_M1='1'
    /* filters open on purpose: cpc_M1=(All), CHRGOFF_RSN_M1=(All) */
  group by cpc_M1;
quit;
```

Return: the split by cpc_M1; the AA row prices what the baseline perimeter
excludes.

## Standing counter-asks (carried from the bridge record, still open)

From [records/cohort-bridge-202501-2026-07-09.md](records/cohort-bridge-202501-2026-07-09.md):
(1) the NEW_ROLL_FLAG x NO_PRIOR_RECORD_FLAG split of DLNQT_CD_M1=1; (2) the
`zenon.cohort_coll_202501` build code (DPD1PLUS entry rule); (5) the DLNQT_CD code
dictionary. Refinements, not blockers — batch them with whichever CQ runs first.
