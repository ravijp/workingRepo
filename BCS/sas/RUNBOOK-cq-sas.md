# SAS run sheet — import the two id lists, run CQ-5 / CQ-7 / CQ-10 / CQ-11

Follow top to bottom. Every code block is one paste-and-submit. The ONLY thing
you ever edit is the `%let csvdir` line in Step 2. Expected numbers are printed
next to each step; if a number does not match, stop and screenshot the log.
(Code is copied from the counter-query pack revision 6; the pack stays the
record, this is the run sheet.)

Before you start you need: the two CSVs from the Athena console exports
(e1 -> `leak_list_202501.csv`, e2 -> `caller_paid30d_list.csv`), saved
anywhere on this machine.

---

## Step 0 — new SAS session only: assign the library

Skip if `zenon` already shows in your library list.

```sas
libname zenon '/sasdata/ECM_CS/hram/Zenon';
```

A "library still in use" note in the log is fine. Any other error: screenshot
the log and stop.

## Step 1 — put the two CSVs where SAS can read them

SAS usually cannot read your Windows desktop directly; the file must sit on
the SAS server.

- **SAS Studio:** left panel, Files (Home) -> the upload button (arrow-up
  icon) -> pick both CSVs. They land in your home folder; the path is
  `/home/<your userid>`.
- **Enterprise Guide:** File menu -> Upload (or the "Copy Files" task,
  direction local-to-server) -> both CSVs -> note the target folder it shows.
- **If neither exists in your menus:** try Step 2 with the Windows path
  directly (e.g. `C:\Users\<you>\Downloads`) — on PC SAS that just works.

Test the path (edit the path in this one line, then submit):

```sas
data _null_; infile "/home/CHANGE-ME/leak_list_202501.csv" obs=3; input; put _infile_; run;
```

PASS = the log prints the header line `extnl_acct_id,deceased_flag` plus two
data rows. FAIL = "physical file does not exist" -> the path is wrong; fix the
path, not the code.

## Step 2 — import both CSVs (edit ONE line, then submit the whole block)

```sas
%let csvdir = /home/CHANGE-ME;   /* <-- EDIT: the folder that passed Step 1 */

data zenon.leak_list_202501;
  infile "&csvdir./leak_list_202501.csv" dsd firstobs=2 truncover;
  length EXTNL_ACCT_ID $20 DECEASED_FLAG 8;
  input EXTNL_ACCT_ID $ DECEASED_FLAG;
run;

data zenon.caller_paid30d_list;
  infile "&csvdir./caller_paid30d_list.csv" dsd firstobs=2 truncover;
  length EXTNL_ACCT_ID $20 PAID_30D_LABEL $12;
  input EXTNL_ACCT_ID $ PAID_30D_LABEL $;
run;
```

Log check: first data step "2177 observations", second "10417 observations".

## Step 3 — verify the round-trip. SCREENSHOT 1 (both little tables)

```sas
proc sql;
  select DECEASED_FLAG, count(*) as accounts
  from zenon.leak_list_202501 group by DECEASED_FLAG;
  select PAID_30D_LABEL, count(*) as accounts
  from zenon.caller_paid30d_list group by PAID_30D_LABEL;
quit;
```

Must show: deceased 0 = 1,967 and 1 = 210; captured 6,630 / leaked 2,177 /
other_caller 1,610. Match -> continue. No match -> stop, screenshot, done for
today.

## Step 4 — CQ-5: price the leak list. SCREENSHOT 2

```sas
proc sql;
  create table zenon.cq5_leak_list_priced as
  select l.DECEASED_FLAG,
         count(*)                                  as accounts_joined,
         sum(w.ECL_M1)                             as ecl_m1,
         sum(w.ECL_M2)                             as ecl_m2,
         sum(w.EOP_BAL_M1)                         as eop_bal_m1,
         sum(case when w.CO_12M_FLAG=1 then 1 else 0 end) as co_12m
  from zenon.waterfall_acct_coll_v1_202501 w
  join zenon.leak_list_202501 l
    on w.EXTNL_ACCT_ID = l.EXTNL_ACCT_ID
  where w.DLNQT_CD_M1='1'
    and w.cpc_M1='others'
    and w.CHRGOFF_RSN_M1 in ('blank','PLY')
  group by l.DECEASED_FLAG;

  select * from zenon.cq5_leak_list_priced;
quit;
```

Expect two rows; accounts_joined should total AT MOST 2,177 (a shortfall is
normal — the baseline filters drop some list accounts). If the total is
suspiciously tiny (say under ~1,500), run the rescue block at the bottom of
this sheet, then continue.

## Step 5 — CQ-7: the true-payment check. SCREENSHOT 3

```sas
proc sql;
  create table zenon.cq7_paid30d_check as
  select l.PAID_30D_LABEL,
         count(*)                                  as accounts,
         sum(w.PAYMT_AMT_M1)                       as paymt_amt_m1,
         sum(w.PAYMT_AMT_M2)                       as paymt_amt_m2,
         sum(case when w.PAYMT_AMT_M1 < 0 or w.PAYMT_AMT_M2 < 0
                  then 1 else 0 end)               as accts_any_payment
  from zenon.waterfall_acct_coll_v1_202501 w
  join zenon.caller_paid30d_list l
    on w.EXTNL_ACCT_ID = l.EXTNL_ACCT_ID
  where w.DLNQT_CD_M1='1'
    and w.cpc_M1='others'
    and w.CHRGOFF_RSN_M1 in ('blank','PLY')
  group by l.PAID_30D_LABEL;

  select * from zenon.cq7_paid30d_check;
quit;
```

Expect three rows (captured / leaked / other_caller); accounts total AT MOST
10,417. Same rescue rule as Step 4.

## Step 6 — CQ-10: what a blank HRAM row means. SCREENSHOT 4 (both grids)

```sas
proc sql;
  select count(*) as rows,
         count(distinct ACCT_ID) as accts,
         sum(case when HRAM_FLAG_APOLLO_new6='Y' then 1 else 0 end) as apollo_y,
         sum(case when HRAM_FLAG_APOLLO_new6='N' then 1 else 0 end) as apollo_n
  from zenon.HRAM_SIMULATION_202412;

  create table zenon.cq10_hram_blank_by_code as
  select w.DLNQT_CD_M1,
         count(*) as accounts,
         sum(case when h.ACCT_ID is not null then 1 else 0 end) as has_hram_row
  from zenon.waterfall_acct_coll_v1_202501 w
  left join zenon.HRAM_SIMULATION_202412 h
    on input(w.EXTNL_ACCT_ID, BEST8.) = h.ACCT_ID
  group by w.DLNQT_CD_M1
  order by w.DLNQT_CD_M1;

  select * from zenon.cq10_hram_blank_by_code;
quit;
```

No expected numbers here — this one ANSWERS a question. Just screenshot both
grids.

## Step 7 — CQ-11: the call table probe. SCREENSHOT 5 (contents) + 6 (grid)

```sas
proc contents data=zenon.aws_call_accts_jan_mar_25; run;

proc sql;
  select initiationmethod,
         count(*) as rows,
         count(distinct extnl_acct_id) as accts
  from zenon.aws_call_accts_jan_mar_25
  group by initiationmethod;
quit;
```

Screenshot the variable list from proc contents and the grid.

---

## Done. Screenshots to bring back: 1-6 above

1. Round-trip counts (Step 3)
2. CQ-5 grid
3. CQ-7 grid
4. CQ-10 both grids
5. CQ-11 proc contents variable list
6. CQ-11 grid

## Rescue block — ONLY if Step 4 or 5 joined suspiciously few accounts

The character keys may disagree on leading zeros. Re-run the failing step with
this join line instead, and say in your paste-back that the rescue version ran:

```sas
    on input(w.EXTNL_ACCT_ID, BEST12.) = input(l.EXTNL_ACCT_ID, BEST12.)
```
