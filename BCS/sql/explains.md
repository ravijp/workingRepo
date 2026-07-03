# Per-card explainers

One section per query id (`## <id>`). The text under a heading renders on
that query's card in the HTML report, under the question. Edit freely -
plain prose, no escaping. A query without a section simply gets no
explainer. Keep the pattern: Window / Why / How to read / Caveat.

## t1_acct_size

Window: the whole account table, current card only (sfx_nbr = 0). Why: sizes the account master and dates its copy - newest_snapshot is the freshness gate every account-side join below inherits. How to read: snapshot dates are yyyymmdd; distinct snapshot count far above the month count means the grain is denser than monthly. Caveat: if newest_snapshot trails the call table, every bucket join reads delinquency as of that older month.

## t1_call_size

Window: the whole call table. Why: establishes the covered period and how often a call carries an account id - the ceiling on every cross-table join. Caveat: the fill rate here is lifetime; see the acctid fill trend for how it moved by month.

## t1_transcript_size

Window: the whole transcript table. Why: history depth (how far back transcripts go) and text fill decide what periods a transcript read can cover. Caveat: one row per utterance; distinct calls is the comparable unit.

## t1_initiation_mix

Window: the whole call table. Why: one conversation can produce several legs (inbound + queue-transfer + transfer), so legs must not be counted as calls. How to read: TRANSFER / QUEUE_TRANSFER legs are downstream continuations of the same customer conversation.

## t1_participants

Window: the whole transcript table. Why: confirms who the speaker labels are and surfaces any third value. How to read: (blank) rows are unattributed utterances - check whether they cluster in specific periods or calls before ignoring them.

## t1_sentiment_mix

Window: the whole transcript table. Why: baselines the per-utterance sentiment label before anyone builds on it. How to read: NEUTRAL dominates for both speakers - normal for procedural talk. Use sentiment as a per-call feature or filter, not as a headline metric.

## t2_monthly_volume

Window: the last 12 complete months, anchored to the newest call date in the table (the in-progress month is excluded so the last point is not an artificial cliff). Why: volume trend and channel mix set the size of everything downstream.

## t2_split_producttype

Window: inbound legs in the last 6 months of call data. Why: which product lines drive inbound. Caveat: blanks are material, and literal 'PRODUCT_TYPE' rows are a loader artifact (header rows ingested as data) - a data-hygiene flag, not a product.

## t2_split_vendor

Window: inbound legs in the last 6 months of call data. Why: vendor split matters for handling-quality questions and for any agent-level read later.

## t2_split_site

Window: inbound legs in the last 6 months of call data. Why: onshore/offshore mix; pairs with the vendor split.

## t2_split_queue

Window: inbound legs in the last 6 months of call data, top 15 queues. Why: whether delinquent-relevant volume rides collections queues or care queues decides how a collections analysis should scope calls. See also the delinquent-caller queue split.

## t2_split_auth

Window: inbound legs in the last 6 months of call data. Why: failed authentication is a friction signal - callers who cannot get in cannot pay.

## t2_split_transfer

Window: inbound legs in the last 6 months of call data. Finding: transfertype on INBOUND legs is always 'Not A Transfer' - transfers live on separate TRANSFER / QUEUE_TRANSFER legs, so this field cannot measure transfer rates. Kept for completeness; use t2_transfer_episodes instead.

## t2_handle_time

Window: inbound handled legs (handled = 1) in the last 6 months of call data. Why: the effort/cost unit per call; median vs p90 shows the tail.

## t2_abandon_transfer_monthly

Window: the last 6 complete months of call data. Why: abandon trend. Caveat: pct_transferred is structurally 0 here (transfers are separate legs, see t2_transfer_episodes); read the abandon line only.

## t2_sentiment_monthly

Window: inbound legs in the last 6 complete months, on the subset with a sentiment score. Why: checks for drift. How to read: averages hover near zero with little variance - weak as a trend line; sentiment earns its keep as a per-call filter, not an aggregate.

## t2_utterances_per_call

Window: every transcribed call. Why: conversation length in turns approximates the reading/processing cost per call for any transcript-based analysis.

## t2_call_minutes

Window: every transcribed call, duration from the last utterance end-timestamp. Why: minutes per call, independent of the call table's handle-time fields - the two should broadly agree.

## t2_dpd_buckets

Window: the newest account snapshot only. Why: the stock ladder - how many accounts sit in each 30-day past-due bucket and with how much balance. How to read: run the PRIMARY query (buckets to 10); the fallback caps at 7 and lumps everything 181+ together. Bucket 10 includes long-charged-off stock still in the table.

## t2_chargeoff_trend

Window: the last 24 months of charge-off dates present in the table, one row per account (first charge-off date, max amount) to undo snapshot repetition. Caveat: the final month reflects the table copy's freshness, not necessarily reality.

## t3_match_rate

Window: inbound legs in the last 6 months of call data vs distinct account ids (sfx_nbr = 0, any snapshot). Why: THE gate - the share of inbound calls that resolve to the account master bounds every cross-table claim. Caveat: match is not correctness; check whether match rate differs for verified vs unverified callers before any agent- or segment-level finding.

## t3_transcript_coverage

Window: inbound legs in the last 6 complete months. Why: measures the known-partial transcript coverage month by month. How to read: a single-month dip is more likely load lag than true loss - re-pull before reading it as missing data.

## t3_repeat_callers

Window: inbound legs with an account id in the last 6 months of call data. Why: repeat calling signals friction or unresolved need and inflates leg counts vs customer counts. Caveat: acctid-less calls (about a quarter) are excluded.

## t3_caller_dpd

Window: inbound legs in the last 6 months of call data, bucket read at the NEWEST account snapshot. Caveat: if the account copy trails the calls, recent calls join to an old bucket and the delinquent share is diluted - v1 (same-month join) is the more honest read; use this one for volume, v1 for shares.

## t3_payment_language

Window: customer utterances on inbound calls in the last 1 month of call data. Why: a first lexicon read of caller intent. Caveat: unconditioned - payment words are common on ordinary service calls (autopay, balance, due date), so this number is an upper bound; v9 splits it by caller delinquency.

## t3_conversation_arc

Window: same 1-month slice, each call cut into thirds by utterance start time. Why: does customer tone recover by the end of the call? How to read: positive rises and negative falls in the final third - calls tend to end on resolution; a call that ends negative is the interesting exception.

## t3_first_last_speaker

Window: same 1-month slice. Why: conversation-shape sanity check (agent opens, who closes) and a place where blank-participant calls surface as '-'.

## v1_dq1_call_concentration

Window: inbound legs in the last 6 months of call data, caller bucket joined in the SAME month as the call. Why: tests 'most calls are DQ1' properly. Caveat: call months after the newest account snapshot cannot match and drop out, so totals trail t3_caller_dpd - read the shares, not the levels.

## v2_vintage_roll

Cohort: accounts entering DQ1 (bucket 0 to 1) 11 months before the newest account snapshot, tracked monthly. Why: the measured migration rates that replace assumed stage multipliers. Caveats: one vintage (rerun 2-3 for seasonality); the visible base shrinks as accounts leave the snapshots; charge-off read from chrgoff_dt.

## v3_caller_vs_noncaller

Cohort: same DQ1-entrant vintage; 'caller' = at least one inbound call in the first 3 months from entry. Why: the forward frame's headline - charge-off and cure rates for callers vs non-callers on the full entrant population. Caveat: callers self-select; this is association, not the effect of calling. Risk-adjust before quoting a lift.

## v4_balance_at_risk

Cohort: same vintage; entry-month balance split by eventual outcome. Why: tests 'balance at DQ1 = balance at risk' - most entry dollars cure. How to read: the charged-off slice is the realistic at-risk pool; per-account balances of eventual charge-offs run higher than cures.

## v5_payment_after_call

Window: inbound calls joined to the caller's same-month bucket (last 6 months), payment read from the NEXT month-end snapshot's last-payment date. Why: the capture/leakage proxy at scale - delinquent calls followed by no payment within 30 days. Caveats: month-end grain, paymt_last_dt parsing, and months after the newest account snapshot drop out.

## v6_stage_proxy

Window: the newest account snapshot, charged-off accounts excluded. Why: an IFRS9-shaped ladder (stage 1 = current+DQ1, 2 = DQ2-3, 3 = 90+) to pair with v2's roll rates for an evidence-based expected-loss shape. Caveat: real staging has overrides and stickiness; treat as the shape, not the booked number.

## v7_ib_ob_mix

Window: all call legs with an account id, same-month bucket join, last 6 months. Finding: OUTBOUND legs carry no account id in this table (outbound column ~0 everywhere) - the outbound side cannot be read here and needs the dialer/PCDS tables or a phone-number join. That absence is itself the result.

## v8_reage_proxy

Window: month-over-month bucket transitions in the last 6 account-snapshot months. Why: sizes deep-bucket resets to current (re-age or multi-cycle cure) before chasing a re-age flag. Caveats: cannot split re-age from genuine payment cure without payment/program fields; the bucket-10 base is dominated by charged-off stock that never resets.

## t1_acctid_fill_trend

Window: the last 24 complete months of call data. Why: the lifetime acctid fill hides a trend - old months with poor fill silently weaken any history-based cohort. How to read: if fill climbs over time, restrict historical joins to the well-filled era or state the bias.

## t2_transfer_episodes

Window: the last 6 complete months; transfer legs chained to their inbound origin via initialcontactid. Why: the correct transfer read (the transfertype field on inbound legs is structurally empty). Caveat: initialcontactid fill was low in profiling - treat as a lower bound on transfer intensity.

## t3_coverage_by_method

Window: all call legs in the last 6 months of call data. Why: transcribed distinct calls exceed total inbound calls, so transcripts must also cover outbound and queue-transfer legs - this sizes coverage per initiation method, which decides whether outbound conversations are readable too.

## t3_delinquent_queues

Window: inbound calls from delinquent accounts (bucket >= 1, same-month join) over the last 3 account-snapshot months. Why: whether delinquent callers ring collections queues or care queues decides the queue filter for any collections call analysis. Caveat: window pinned to months the account table actually covers.

## v9_payment_language_by_bucket

Window: inbound calls in the newest account-snapshot month, customer utterances only, split current vs delinquent caller. Why: conditions the raw payment-language rate on the population that matters - delinquent callers - and contrasts hardship/escalation language. Caveats: one month; lexicon proxy; same staleness pin as all same-month joins.
