Four blocks, in priority order. Each is aggregates-only and runs on what already exists.

A — CQ-3 rerun with the stage-2 reason code (names why 16,597 same-count accounts are S2 while 10,898 stay S1; likely settles the S1-stayer story and cross-checks the blank-HRAM question):


proc sql;
  create table zenon.cq3_s1_stayers_rsn as
  select w.STG_CD_M2,
         i.WGHTD_STG2_RSN_CD as stg2_rsn_at_m2,
         i.CYCL_DLNQT_CNT    as imp_dq_at_m2,
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
  group by w.STG_CD_M2, i.WGHTD_STG2_RSN_CD, i.CYCL_DLNQT_CNT;
quit;
B — the code-vs-count mapping probe (this can decide the A-068 reading from data, without waiting on a definition: cross-tab ASP's code against the impairment count for the whole cohort in the same month. If code k lines up cleanly with count k−1 everywhere, it's a convention offset and the estimation saw the roll; if the mass is shifted or mixed, it's a lag):


proc sql;
  create table zenon.cq9_code_count_map as
  select w.DLNQT_CD_M1,
         i.CYCL_DLNQT_CNT   as imp_dq_at_m1,
         count(*)           as accounts
  from zenon.waterfall_acct_coll_v1_202501 w
  join pcds2.V_IFRS9_NON_PMA_IMPRMNT_OUTPUT i
    on w.EXTNL_ACCT_ID = i.EXTNL_ACCT_ID
   and i.IMPRMNT_RPTG_MNTH_CD = '2025M01'
  where w.DLNQT_CD_M1 is not null
  group by w.DLNQT_CD_M1, i.CYCL_DLNQT_CNT
  order by w.DLNQT_CD_M1, i.CYCL_DLNQT_CNT;
quit;
(Deliberately no cpc/reason filters — the mapping is definitional, so wider is better. Expect roughly 7 codes × a few counts.)

C — the ASP schema scan (finishes CQ-6: look for any CUST/PARTY/HSHLD column; also gives us the full ASP variable list for the code dictionary work):


proc contents data=pcds.V_ASP_EOM_ACCT_SUM varnum;
run;
D — optional, one cheap closure: the full charge-off reason-code list (an [OPEN] since the waterfall record — the live workbook showed more codes than the screenshot):


proc sql;
  select CHRGOFF_STATUS_RSN_CD, count(*) as rows
  from pcds.V_ASP_EOM_ACCT_SUM
  where RPTG_PRD_MNTH_BID = 202501
    and CHRGOFF_STATUS_RSN_CD is not null
  group by CHRGOFF_STATUS_RSN_CD;
quit;
Share whatever lands, in any order — A and B are the ones that move the A-068 verdict and the gated cure-row/unit-value reads.
