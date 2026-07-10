# SAS counter-query pack for Ishant — 2026-07-10

Split ownership per the 2026-07-10 decision: Athena owns the funnel + transcript
story; SAS owns the dollar layer (ECL, stages, HRAM, payments). All SQL below is
written against the columns your three shared files actually create (verbatim
transcription:
[records/ishant-sas-code-transcription-2026-07-10.md](records/ishant-sas-code-transcription-2026-07-10.md)).
We consume **aggregates only** — counts and sums, no account rows. Hand this file
over verbatim. CQ-1 to CQ-3 are the original pack (revised 2026-07-10 against the
code); CQ-4 to CQ-8 were appended 2026-07-10 from the cohort-picture arbitration
([uc2-jan2025-cohort-picture.md](uc2-jan2025-cohort-picture.md)).

Table targeting, per the build code (file 03 UPDATED version, 2026-07-10):

- CQ-1, CQ-2, CQ-4 run on **`zenon.waterfall_coll_call_202501`** (file 03's
  output: `_v1_` columns plus call_type_CLBCK/INB/TRSFR, the six HRAM flags
  hram_flag_refit_M1..M3 / hram_flag_apollo_M1..M3, and CPC_FLAG_NW).
- CQ-3, CQ-5 to CQ-8 run on **`zenon.waterfall_acct_coll_v1_202501`** (file
  02's account-wide pivot); they use no call or HRAM fields.
- The wide tables carry **no COHORT_ENTRY column** (the flag lives only on
  `zenon.cohort_coll_202501`). `DLNQT_CD_M1='1'` already implies the DPD1PLUS
  class, so the queries filter on it alone.

HRAM conventions, settled 2026-07-10 (Ravi): the operative flag is
`HRAM_FLAG_APOLLO_new6` (the SQL below uses it); HRAM is produced at end of
month, so the M1 alias correctly carries the 202412 (December) score; the HRAM
ids are 8-wide, so the `BEST8.` informat is correct.

One HRAM question the CQ-2/CQ-4 results opened (2026-07-11): 123,816 of the
186,412 code-1 accounts (66.4%) have NO row in HRAM_SIMULATION_202412 — the
flag comes back blank, not N. What does a missing row mean: not scored =
non-HRAM, or a coverage/population gap in the simulation table? No HRAM share
gets quoted until this is answered.

Run-state note (2026-07-10, first combined run): CQ-1, CQ-2, CQ-4, CQ-8
completed and produced their zenon.cq* tables — please share those four
results. CQ-3, CQ-5, CQ-6, CQ-7 errored for the reasons now fixed in their
sections below (CQ-3 step 1 must run first and name the DQ variable; CQ-5 and
CQ-7 wait on Zenon's id-list drops; CQ-6 starts with a column scan).

Baseline filters, same as every shared pivot (apply unless a query says
otherwise): `DLNQT_CD_M1='1'`, `cpc_M1='others'` (ex-AA),
`CHRGOFF_RSN_M1 in ('blank','PLY')`. Grain: everything here is month-END
(DLNQT_CD is the month-end position).

Codings confirmed from the build code (2026-07-10): CO_CURRENT/8M/10M/12M flags
are numeric 1/0; the account key is `EXTNL_ACCT_ID`; the impairment source view
is `pcds2.V_IFRS9_NON_PMA_IMPRMNT_OUTPUT` with month key
`IMPRMNT_RPTG_MNTH_CD` ('YYYYMMM', e.g. '2025M02').

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
         sum(case when CO_12M_FLAG=1 then 1 else 0 end) as co_12m
  from zenon.waterfall_coll_call_202501   /* file 03 (updated) output: call types + HRAM flags */
  where DLNQT_CD_M1='1'
    and cpc_M1='others'
    and CHRGOFF_RSN_M1 in ('blank','PLY')
  group by DLNQT_CD_M1, call_type_INB, DLNQT_CD_M2
  order by call_type_INB, DLNQT_CD_M2;
quit;
```

Return: the full result (it is ~16 rows). Please also state, one line each:

1. The **build of `work.call_accts_jan_mar_25`**. Your file 03 shows
   call_type_INB = max(initiationmethod='INBOUND') from that work table grouped
   by extnl_acct_id, but the code that builds the work table is not in the
   shared files. Two things we need: (a) its source and filter (AWS call table?
   only rows with an account id, i.e. verified-only like our side?); (b) confirm
   the window is January-March 2025 as the name suggests. Note for reading the
   result: your flag is then "called inbound at any point Jan-Mar", while our b7
   counts January callers only, so the counts will not match by construction;
   we reconcile on definitions, not on equality.
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

HRAM source found in file 03 (updated): `zenon.HRAM_SIMULATION_*` joined as
hram_flag_refit_M1..M3 / hram_flag_apollo_M1..M3. The SQL below uses the Apollo
flag; if refit_new2 is the operative definition, swap it (and answer the three
HRAM questions in the header).

```sas
proc sql;
  create table zenon.cq2_hram_roll_split as
  select hram_flag_apollo_M1,
         STG_CD_M1,
         count(*)          as accounts,
         sum(ECL_M1)       as ecl_m1,
         sum(ECL_M2)       as ecl_m2,
         sum(ECL_M3)       as ecl_m3,
         sum(EOP_BAL_M1)   as eop_bal_m1,
         sum(case when CO_12M_FLAG=1 then 1 else 0 end) as co_12m
  from zenon.waterfall_coll_call_202501
  where DLNQT_CD_M1='1' and DLNQT_CD_M2='2'   /* the 65,585 rollers */
    and cpc_M1='others'
    and CHRGOFF_RSN_M1 in ('blank','PLY')
  group by hram_flag_apollo_M1, STG_CD_M1;
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

Step 1 — name the variable (RUN THIS FIRST; step 2 will not compile without
it). List the view's columns and identify the delinquency/bucket variable the
estimation used (your build pulls only `WGHTD_EXPTD_LOSS_AMT`, `_12MO_AMT`,
`_LIFTM_AMT`, `WGHTD_STG_CD`, `WRITE_OFF_AMT`). Also note whether the view
carries a cycle date.

```sas
proc sql;
  select name, type, length
  from dictionary.columns
  where libname='PCDS2' and memname='V_IFRS9_NON_PMA_IMPRMNT_OUTPUT';
quit;
```

Then set the macro variable step 2 uses:

```sas
%let IMP_DQ = FILL_FROM_STEP_1;   /* the DQ/bucket column name found above */
```

Step 2 — cross-tab the S1-stayers (after step 1 only):

```sas
proc sql;
  create table zenon.cq3_s1_stayers_dq as
  select w.STG_CD_M2,
         i.&IMP_DQ.          as imp_dq_at_m2,
         w.DLNQT_CD_M2,
         count(*)            as accounts,
         sum(w.ECL_M2)       as ecl_m2,
         mean(w.EOP_BAL_M2)  as avg_bal_m2
  from zenon.waterfall_acct_coll_v1_202501 w
  join pcds2.V_IFRS9_NON_PMA_IMPRMNT_OUTPUT i
    on w.EXTNL_ACCT_ID = i.EXTNL_ACCT_ID
   and i.IMPRMNT_RPTG_MNTH_CD = '2025M02'
  where w.DLNQT_CD_M1='1' and w.DLNQT_CD_M2='2'
    and w.STG_CD_M1='S1'
    and w.cpc_M1='others'
    and w.CHRGOFF_RSN_M1 in ('blank','PLY')
  group by w.STG_CD_M2, i.&IMP_DQ., w.DLNQT_CD_M2;
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
names the 110-account remainder. HRAM flag as in CQ-2: the Apollo variant from
file 03 (updated); swap to refit_new2 if that is the operative definition.

```sas
proc sql;
  create table zenon.cq4_ledger_priced_m2 as
  select DLNQT_CD_M2,   /* every value, missing included: the missing rows name the 110 */
         hram_flag_apollo_M1,
         count(*)                                  as accounts,
         sum(ECL_M1)                               as ecl_m1,
         sum(ECL_M2)                               as ecl_m2,
         sum(ECL_M3)                               as ecl_m3,
         sum(EOP_BAL_M1)                           as eop_bal_m1,
         sum(EOP_BAL_M2)                           as eop_bal_m2,
         sum(case when CO_12M_FLAG=1 then 1 else 0 end) as co_12m
  from zenon.waterfall_coll_call_202501
  where DLNQT_CD_M1='1'
    and cpc_M1='others'
    and CHRGOFF_RSN_M1 in ('blank','PLY')
  group by DLNQT_CD_M2, hram_flag_apollo_M1
  order by DLNQT_CD_M2, hram_flag_apollo_M1;
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

PRECONDITION: do not run until Zenon confirms the drop of
`zenon.leak_list_202501` (account ids + a deceased flag column); running it
before produces "file does not exist". Mechanics: Zenon drops the account-id
list in the shared location; data stays in the client environment.

```sas
proc sql;
  create table zenon.cq5_leak_list_priced as
  select count(*)                                  as accounts_joined,
         sum(w.ECL_M1)                             as ecl_m1,
         sum(w.ECL_M2)                             as ecl_m2,
         sum(w.EOP_BAL_M1)                         as eop_bal_m1,
         sum(case when w.CO_12M_FLAG=1 then 1 else 0 end) as co_12m
  from zenon.waterfall_acct_coll_v1_202501 w
  join zenon.leak_list_202501 l                    /* the dropped id list */
    on w.EXTNL_ACCT_ID = l.EXTNL_ACCT_ID
  where w.DLNQT_CD_M1='1'
    and w.cpc_M1='others'
    and w.CHRGOFF_RSN_M1 in ('blank','PLY');
quit;
```

Return: the one aggregate row: sum(ECL_M1), sum(ECL_M2), sum(EOP_BAL_M1), CO_12M
count on the join. Run twice once the deceased-routing query (m4) lands on our
side: with and without the deceased-flagged accounts.

## CQ-6 — the customer key (one-liner)

Why: IFRS 9 staging is customer-level; the Athena side has no customer id.

Step 1 — scan for a customer-id-like column (the waterfall table has none named
CUST_ID; the first combined run proved that, which itself half-answers the ask):

```sas
proc sql;
  select libname, memname, name
  from dictionary.columns
  where ((libname='PCDS'  and memname='V_ASP_EOM_ACCT_SUM')
      or (libname='PCDS2' and memname='V_IFRS9_NON_PMA_IMPRMNT_OUTPUT'))
    and (upcase(name) like '%CUST%' or upcase(name) like '%PARTY%'
      or upcase(name) like '%HSHLD%' or upcase(name) like '%HOUSEHOLD%');
quit;
```

Step 2 — only if step 1 finds a key (fill its name):

```sas
%let CUST_ID = FILL_FROM_STEP_1;

proc sql;
  create table zenon.cq6_customer_key as
  select count(*) as accts_sharing_customer
  from zenon.waterfall_acct_coll_v1_202501 w
  join pcds.V_ASP_EOM_ACCT_SUM k
    on w.EXTNL_ACCT_ID = k.EXTNL_ACCT_ID and k.RPTG_PRD_MNTH_BID = 202501
  where w.DLNQT_CD_M1='1'
    and w.cpc_M1='others'
    and w.CHRGOFF_RSN_M1 in ('blank','PLY')
    and k.&CUST_ID. in (select &CUST_ID.
                        from pcds.V_ASP_EOM_ACCT_SUM
                        where RPTG_PRD_MNTH_BID = 202501
                          and DLNQT_CD is not null and DLNQT_CD ne '0'
                        group by &CUST_ID.
                        having count(*) > 1);
quit;
```

Return: step 1's hits (or "none", which is itself the answer: no customer key in
ASP), and the one count if step 2 runs. Decides whether the work list needs
customer-grain dedup and bounds the customer-level-staging disclosure. Side
note: on the Athena side, `contactcenter_bdp_db.call` carries a PartyID per the
table inventory; if ASP has no key, that is the fallback lead for callers.

## CQ-7 — true-payment check of the paid-30d gate

Why: every capture/leak label on the Athena side rests on a last-payment-date
proxy with autopay and bounced exclusions. PAYMT_AMT (negative = payment) is the
true payment record and can validate the gate itself.

PRECONDITION: do not run until Zenon confirms the drop of
`zenon.caller_paid30d_list`; running it before produces "file does not exist".
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
  from zenon.waterfall_acct_coll_v1_202501 w
  join zenon.caller_paid30d_list l                 /* the dropped id + label list */
    on w.EXTNL_ACCT_ID = l.EXTNL_ACCT_ID
  where w.DLNQT_CD_M1='1'
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
  from zenon.waterfall_acct_coll_v1_202501
  where DLNQT_CD_M1='1'
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
