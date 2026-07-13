Four blocks, in priority order. Each is aggregates-only and runs on what already exists.

A — CQ-3 rerun with the stage-2 reason code (names why 16,597 same-count accounts are S2 while 10,898 stay S1; likely settles the S1-stayer story and cross-checks the blank-HRAM question):

```
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
```
B — the code-vs-count mapping probe (this can decide the A-068 reading from data, without waiting on a definition: cross-tab ASP's code against the impairment count for the whole cohort in the same month. If code k lines up cleanly with count k−1 everywhere, it's a convention offset and the estimation saw the roll; if the mass is shifted or mixed, it's a lag):

```
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
```
(Deliberately no cpc/reason filters — the mapping is definitional, so wider is better. Expect roughly 7 codes × a few counts.)

C — the ASP schema scan (finishes CQ-6: look for any CUST/PARTY/HSHLD column; also gives us the full ASP variable list for the code dictionary work):

```
proc contents data=pcds.V_ASP_EOM_ACCT_SUM varnum;
run;
```
D — optional, one cheap closure: the full charge-off reason-code list (an [OPEN] since the waterfall record — the live workbook showed more codes than the screenshot):

```
proc sql;
  select CHRGOFF_STATUS_RSN_CD, count(*) as rows
  from pcds.V_ASP_EOM_ACCT_SUM
  where RPTG_PRD_MNTH_BID = 202501
    and CHRGOFF_STATUS_RSN_CD is not null
  group by CHRGOFF_STATUS_RSN_CD;
quit;
```
Share whatever lands, in any order — A and B are the ones that move the A-068 verdict and the gated cure-row/unit-value reads.

---

```
proc sql;
  create table zenon.cq5_leak_list_priced as
  select l.DECEASED_FLAG,
         count(*) as accounts_joined,
         sum(w.ECL_M1) as ecl_m1,
         sum(w.ECL_M2) as ecl_m2,
         sum(w.EOP_BAL_M1) as eop_bal_m1,
         sum(case when w.CO_12M_FLAG=1 then 1 else 0 end) as co_12m
  from zenon.waterfall_acct_coll_v1_202501 w
  join WORK.e1_leak_list_202501 l
    on input(w.EXTNL_ACCT_ID, BEST12.) = l.EXTNL_ACCT_ID
  where w.DLNQT_CD_M1='1'
    and w.cpc_M1='others'
    and w.CHRGOFF_RSN_M1 in ('blank','PLY')
  group by l.DECEASED_FLAG;
  select * from zenon.cq5_leak_list_priced;
quit;
```

```
proc sql;
  create table zenon.cq7_paid30d_check as
  select l.PAID_30D_LABEL,
         count(*) as accounts,
         sum(w.PAYMT_AMT_M1) as paymt_amt_m1,
         sum(w.PAYMT_AMT_M2) as paymt_amt_m2,
         sum(case when w.PAYMT_AMT_M1 < 0 or w.PAYMT_AMT_M2 < 0
                  then 1 else 0 end) as accts_any_payment
  from zenon.waterfall_acct_coll_v1_202501 w
  join WORK.e2_caller_paid30d_list l
    on input(w.EXTNL_ACCT_ID, BEST12.) = l.EXTNL_ACCT_ID
  where w.DLNQT_CD_M1='1'
    and w.cpc_M1='others'
    and w.CHRGOFF_RSN_M1 in ('blank','PLY')
  group by l.PAID_30D_LABEL;
  select * from zenon.cq7_paid30d_check;
quit;
```


---
```
WITH snap AS (
    SELECT extnl_acct_id,
           eff_dt,
           CASE
             WHEN past_due_271_up_amt  > 0 THEN 10
             WHEN past_due_241_270_amt > 0 THEN 9
             WHEN past_due_211_240_amt > 0 THEN 8
             WHEN past_due_181_210_amt > 0 THEN 7
             WHEN past_due_151_180_amt > 0 THEN 6
             WHEN past_due_121_150_amt > 0 THEN 5
             WHEN past_due_91_120_amt  > 0 THEN 4
             WHEN past_due_61_90_amt   > 0 THEN 3
             WHEN past_due_31_60_amt   > 0 THEN 2
             WHEN past_due_1_30_amt    > 0 THEN 1
             ELSE 0
           END AS bucket,
           try_cast(acct_bal_amt AS double) AS bal,
           try_cast(chrgoff_dt AS date) AS co_dt,
           clnt_prdct_cd
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250201'
),
monthly AS (
    SELECT extnl_acct_id,
           max_by(bucket, eff_dt) AS eom_bucket,
           max_by(bal, eff_dt) AS eom_bal,
           min(co_dt) AS co_dt,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc
    FROM snap GROUP BY 1
)
SELECT CASE
         WHEN eom_cpc IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2',
                          'AM2','AA3','AC3','AM3','AA4','AC4','AM4') THEN 'a. AA'
         WHEN eom_cpc IN ('BGC','BGM','CGM','GMR')                  THEN 'b. GM'
         WHEN eom_cpc IN ('FBS','IBS','U1C','U2C','U3C')            THEN 'c. Bronco'
         ELSE 'd. others'
       END AS pc_class,
       count(*) AS pc_accounts,
       round(sum(eom_bal), 0) AS pc_jan_eom_balance
FROM monthly
WHERE eom_bucket = 1
  AND (co_dt IS NULL OR co_dt >= DATE '2025-01-01')
GROUP BY 1 ORDER BY 1
```


---
```
WITH inb AS (
    SELECT trim(cast(acctid AS varchar)) AS acct_key, contactid,
           "date" AS call_dt, initiationtimestamp
    FROM "contactcenter_bdp_db"."call"
    WHERE initiationmethod = 'INBOUND'
      AND "date" >= DATE '2025-01-01' AND "date" < DATE '2025-02-01'
      AND effdt >= '2025-01-01' AND effdt < '2025-02-02'
      AND coalesce(cast(producttype AS varchar), '') <> 'BUSINESS_CARD'
),
episodes AS (
    SELECT acct_key, contactid, call_dt,
           cast(date_trunc('month', call_dt) AS date) AS call_month
    FROM (
        SELECT acct_key, contactid, call_dt,
               row_number() OVER (PARTITION BY acct_key, call_dt
                                  ORDER BY initiationtimestamp) AS rn
        FROM inb
        WHERE acct_key IS NOT NULL AND acct_key <> ''
    )
    WHERE rn = 1
),
cpc_monthly AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           max_by(clnt_prdct_cd, eff_dt) AS eom_cpc
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250201'
      AND trim(cast(extnl_acct_id AS varchar)) IN (SELECT acct_key FROM episodes)
    GROUP BY 1
),
episodes_exaa AS (
    SELECT e.acct_key, e.contactid, e.call_dt, e.call_month
    FROM episodes e
    LEFT JOIN cpc_monthly c ON c.acct_key = e.acct_key
    WHERE (c.eom_cpc IS NULL OR trim(c.eom_cpc) = ''
           OR c.eom_cpc NOT IN ('AA2','BC5','BA5','AA1','AC1','AM1','AC2','AM2',
                                 'AA3','AC3','AM3','AA4','AC4','AM4',
                                 'BGC','BGM','CGM','GMR',
                                 'FBS','IBS','U1C','U2C','U3C'))
),
snap AS (
    SELECT trim(cast(extnl_acct_id AS varchar)) AS acct_key,
           eff_dt,
           date(date_parse(eff_dt, '%Y%m%d')) AS snap_dt,
           CASE
             WHEN past_due_271_up_amt  > 0 THEN 10
             WHEN past_due_241_270_amt > 0 THEN 9
             WHEN past_due_211_240_amt > 0 THEN 8
             WHEN past_due_181_210_amt > 0 THEN 7
             WHEN past_due_151_180_amt > 0 THEN 6
             WHEN past_due_121_150_amt > 0 THEN 5
             WHEN past_due_91_120_amt  > 0 THEN 4
             WHEN past_due_61_90_amt   > 0 THEN 3
             WHEN past_due_31_60_amt   > 0 THEN 2
             WHEN past_due_1_30_amt    > 0 THEN 1
             ELSE 0
           END AS bucket,
           try_cast(chrgoff_dt AS date) AS co_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20240601' AND eff_dt < '20250201'
      AND trim(cast(extnl_acct_id AS varchar)) IN (SELECT acct_key FROM episodes_exaa)
),
pay_snap AS (
    SELECT extnl_acct_id, eff_dt,
           date_trunc('month', date(date_parse(eff_dt, '%Y%m%d'))) AS m,
           coalesce(try_cast(paymt_last_dt AS date),
                    try(cast(date_parse(try_cast(paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS pay_dt,
           coalesce(try_cast(atmtc_paymt_last_dt AS date),
                    try(cast(date_parse(try_cast(atmtc_paymt_last_dt AS varchar), '%d%b%Y') AS date))) AS auto_dt,
           coalesce(try_cast(nsf_last_paymt_dt AS date),
                    try(cast(date_parse(try_cast(nsf_last_paymt_dt AS varchar), '%d%b%Y') AS date))) AS nsf_dt
    FROM "fmt_acct_dba"."fmt_acct_c"
    WHERE sfx_nbr = 0
      AND eff_dt >= '20250101' AND eff_dt < '20250301'
      AND trim(cast(extnl_acct_id AS varchar)) IN (SELECT acct_key FROM episodes_exaa)
),
pay_monthly AS (
    SELECT extnl_acct_id, m,
           max(pay_dt) AS pay_dt,
           max(auto_dt) AS auto_dt,
           max(nsf_dt) AS nsf_dt
    FROM pay_snap GROUP BY 1, 2
),
pay_monthly2 AS (
    SELECT extnl_acct_id, m, pay_dt, auto_dt, nsf_dt,
           lead(pay_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_pay_dt,
           lead(auto_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_auto_dt,
           lead(nsf_dt) OVER (PARTITION BY extnl_acct_id ORDER BY m) AS next_nsf_dt
    FROM pay_monthly
),
callday AS (
    SELECT e.acct_key, e.call_dt,
           max_by(s.bucket, s.eff_dt) AS callday_bucket,
           max_by(s.co_dt, s.eff_dt) AS callday_co_dt
    FROM episodes_exaa e
    JOIN snap s
      ON s.acct_key = e.acct_key
     AND s.snap_dt <= e.call_dt
    GROUP BY 1, 2
),
kept AS (
    SELECT acct_key, call_dt
    FROM callday
    WHERE callday_bucket = 1
      AND (callday_co_dt IS NULL OR callday_co_dt >= DATE '2025-01-01')
),
ep AS (
    SELECT e.acct_key, e.contactid,
           CASE WHEN
                  (s.pay_dt IS NOT NULL
                   AND s.pay_dt >= e.call_dt
                   AND s.pay_dt <= date_add('day', 30, e.call_dt)
                   AND (s.auto_dt IS NULL OR s.auto_dt <> s.pay_dt)
                   AND (s.nsf_dt IS NULL OR s.nsf_dt <> s.pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.pay_dt))
                OR
                  (s.next_pay_dt IS NOT NULL
                   AND s.next_pay_dt >= e.call_dt
                   AND s.next_pay_dt <= date_add('day', 30, e.call_dt)
                   AND (s.next_auto_dt IS NULL OR s.next_auto_dt <> s.next_pay_dt)
                   AND (s.next_nsf_dt IS NULL OR s.next_nsf_dt <> s.next_pay_dt))
                THEN 1 ELSE 0 END AS captured
    FROM kept k
    JOIN episodes_exaa e
      ON e.acct_key = k.acct_key
     AND e.call_dt = k.call_dt
    LEFT JOIN pay_monthly2 s
      ON e.acct_key = trim(cast(s.extnl_acct_id AS varchar))
     AND e.call_month = cast(s.m AS date)
),
tx AS (
    SELECT t.contactid,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'passed away|death certificate|executor|deceased|calling on behalf'))
               AS deceased_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'pay|paid|payment|settle|payment plan|arrangement|work something out'))
               AS pay_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'settle|payment plan|arrangement|work something out'))
               AS plan_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'hardship|lost my job|laid off|unemploy|hospital|sick|struggl|can.t afford'))
               AS hard_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'dispute|not my charge|didn.t authorize|did not authorize|unauthorized|fraud|identity theft'))
               AS dispute_n,
           count_if(t.participantid = 'CUSTOMER'
                    AND regexp_like(lower(t.content),
                        'i.ll pay|i will pay|going to pay|gonna pay|pay (on|by|this|next)|when i get paid|payday|after my paycheck'))
               AS promise_n
    FROM "contactcenter_bdp_db"."transcript" t
    JOIN (SELECT DISTINCT contactid FROM ep) d
      ON t.contactid = d.contactid
     AND t.effdt >= '2025-01-01' AND t.effdt < '2025-02-02'
    WHERE t.content IS NOT NULL
    GROUP BY 1
),
ep2 AS (
    SELECT e.acct_key, e.captured,
           CASE
             WHEN coalesce(x.deceased_n, 0) > 0 THEN 'a. deceased or estate'
             WHEN coalesce(x.promise_n, 0) > 0 THEN 'b. future-dated promise'
             WHEN coalesce(x.pay_n, 0) > 0
                  AND coalesce(x.plan_n, 0) = 0 THEN 'c. payment talk, no promise'
             WHEN coalesce(x.plan_n, 0) > 0 THEN 'd. plan or settlement talk'
             WHEN coalesce(x.hard_n, 0) > 0 THEN 'e. hardship talk'
             WHEN coalesce(x.dispute_n, 0) > 0 THEN 'f. dispute or fraud talk'
             ELSE 'g. no payment-related language'
           END AS language_group
    FROM ep e
    LEFT JOIN tx x ON e.contactid = x.contactid
)
SELECT language_group AS b19_language_group,
       captured AS b19_captured,
       count(*) AS b19_episodes,
       count(DISTINCT acct_key) AS b19_accounts
FROM ep2
GROUP BY 1, 2
ORDER BY 1, 2
```
