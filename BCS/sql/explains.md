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

Window: inbound legs in the last 6 complete ACCOUNT months (anchored to the account table's clock, not the call table's) vs distinct account ids (sfx_nbr = 0, any snapshot). Why: THE gate - the share of inbound calls that resolve to the account master bounds every cross-table claim. The account copy trails the calls, so call months past its edge would under-match against a master that cannot know newly opened accounts; anchoring to the account clock removes that dilution and self-heals when the copy refreshes. Caveat: match is not correctness; f4_match_by_auth splits this rate by authentication outcome before any segment-level finding.

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

Window: inbound legs in the last 6 complete ACCOUNT months, caller bucket joined in the SAME month as the call. Why: tests 'most calls are DQ1' properly. The window now rides the account table's clock, so every call month has a complete same-month snapshot - levels and shares are both readable (the old call-clock window silently dropped its newest months). Self-heals when the account copy refreshes.

## v2_vintage_roll

Cohort: accounts entering DQ1 (bucket 0 to 1) 11 months before the newest account snapshot, tracked monthly. Why: the measured migration rates that replace assumed stage multipliers. Caveats: one vintage (rerun 2-3 for seasonality); the visible base shrinks as accounts leave the snapshots; charge-off read from chrgoff_dt.

## v3_caller_vs_noncaller

Cohort: same DQ1-entrant vintage; 'caller' = at least one inbound call in the first 3 months from entry. Why: the forward frame's headline - charge-off and cure rates for callers vs non-callers on the full entrant population. Caveat: callers self-select; this is association, not the effect of calling. Risk-adjust before quoting a lift.

## v4_balance_at_risk

Cohort: same vintage; entry-month balance split by eventual outcome. Why: tests 'balance at DQ1 = balance at risk' - most entry dollars cure. How to read: the charged-off slice is the realistic at-risk pool; per-account balances of eventual charge-offs run higher than cures.

## v5_payment_after_call

Window: inbound calls joined to the caller's same-month bucket, payment read from the NEXT month-end snapshot's last-payment date. Calls span the 5 complete account months BEFORE the account table's newest complete month, so every call keeps a full following-month snapshot for the 30-day check (the old call-clock window let its newest months drop out or truncate the payment runway). Why: the capture/leakage proxy at scale - delinquent calls followed by no payment within 30 days. Caveats: month-end grain and paymt_last_dt parsing still apply.

## v6_stage_proxy

Window: the newest account snapshot, charged-off accounts excluded. Why: an IFRS9-shaped ladder (stage 1 = current+DQ1, 2 = DQ2-3, 3 = 90+) to pair with v2's roll rates for an evidence-based expected-loss shape. Caveat: real staging has overrides and stickiness; treat as the shape, not the booked number.

## v7_ib_ob_mix

Window: all call legs with an account id, same-month bucket join, last 6 complete ACCOUNT months (anchored to the account table's clock so every window month joins completely; self-heals on refresh). Finding: OUTBOUND legs carry no account id in this table (outbound column ~0 everywhere) - the outbound side cannot be read here and needs the dialer/PCDS tables or a phone-number join. That absence is itself the result.

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

Window: inbound calls in the newest COMPLETE account month (one month before the account table's newest month, which is a partial position - days, not a month), customer utterances only, split current vs delinquent caller. Why: conditions the raw payment-language rate on the population that matters - delinquent callers - and contrasts hardship/escalation language. Caveats: one month; lexicon proxy; h-series and n-series extend this read.

## f1_funnel_waterfall

Window: PINNED to call months 2024-07 through 2025-06 (W3) so every episode keeps at least 8 account months of outcome runway before the account copy's edge; outcomes read through 2026-02. Run f0 first - if any W3 month shows broken coverage or fill, move the window before trusting this. Why: the headline. Episodes are first-inbound-per-account-per-day (the reference dedup); each row is a cumulative gate, so the last row is the leakage pool that actually charged off. Balance is the account's balance at its first matched episode, held constant down the funnel - the dollar column is a true waterfall. How to read: stage g (no payment within 30 days) is the leakage candidate pool; stage h is the slice of it that became realized loss. Caveats: payment is the next-snapshot proxy, not a ledger join; the language gate is the v9 lexicon; unmatched episodes (stage b minus c) are measured by f4_match_by_auth.

## f2_funnel_by_month

Window: same pinned W3 as f1, one row per call month. Why: kills the single-month objection - if the stage-to-stage drops hold vintage by vintage, the pooled waterfall is not an artifact of one odd month. How to read: columns are cumulative episode counts; divide any column by the previous one for the per-month gate rate. Caveat: later call months have shorter absolute follow-up beyond the guaranteed 8; chargeoff_8m stays comparable because the 8-month horizon is fixed per episode.

## f6_funnel_by_bucket

Window: same pinned W3 as f1, delinquent gates split by the caller's same-month bucket. Why: early buckets carry the episode volume, late buckets carry the loss probability - this shows where episodes x loss actually concentrates, which decides whether a capture play scopes to DQ1-3 or chases the deep ladder. How to read: divide chargeoff_8m by no_payment_30d per bucket for the leaked-to-loss conversion; compare against v5's payment rates.

## f7_leaked_dollars_by_month

Window: same pinned W3 as f1. Why: the dollar trend behind the funnel end - leaked balances and the charge-off dollars that followed, month by month. This is the chart-shaped version of the story for anyone who will not read a waterfall. Caveat: an account that leaks in several months appears in each; balances are month-end positions, not exposure-at-loss.

## f3_funnel_dollars

Window: charge-offs 2024-08 through 2026-02; calls observed 2024-07 through 2026-01 (each call needs a same-month snapshot and a 30-day payment runway inside the account edge). Why: flips the funnel around to the loss ledger - of realized charge-off dollars, how much sat on accounts the inbound channel touched and did not capture? The funnel-end share is the addressable slice of actual losses. How to read: compare avg_chargeoff across groups - do leaked accounts lose more per account than never-callers? Caveat: the last charge-off months see a shorter call lookback, so 'no inbound call observed' is an upper bound there.

## f5_calls_before_chargeoff

Window: DQ1 entrants 2024-07 through 2025-06 (3-month entry buffer) that charged off by 2026-02; inbound episodes counted from entry to charge-off. Why: the intervention surface on the loss path. '0 calls' losses are unreachable inbound; the 1-2 call bands are where a single missed capture is the whole story; the 11+ band is the repeat-friction narrative at volume. Caveat: episodes require an account id on the call, so the bands are floors.

## f4_match_by_auth

Window: last 5 complete account months (W2-equivalent), inbound calls vs the account master. Why: DPT-style bias gate - the funnel only sees matched calls, and if unmatched calls concentrate among FAILED-authentication callers, every funnel rate is biased toward the easy population. How to read: compare pct_matched_master across authentication outcomes; a large gap means quote funnel rates as 'of matched calls' and treat the unmatched slice as its own population. Caveat: blank authentication is its own signal, not noise.

## f4_coverage_by_bucket

Window: last 5 complete account months, same-month bucket join, inbound calls with an account id. Why: the second bias gate - if deep-bucket calls are transcribed less often, the transcript and language gates undercount exactly where the dollars are. How to read: flat coverage across buckets clears the gate; a slope means bucket-condition the language rates before quoting them.

## f0_period_calibration

Window: the full call history, inbound only, by month. Why: run FIRST. Every windowed claim sits on three constraints - reliable transcript coverage, stable account-id fill, and (for the funnel) 8 months of outcome runway before the account edge. This is the one table that shows all three eras at once and confirms or moves the pinned funnel window. How to read: find where pct_with_transcript stabilises high and pct_with_acctid enters its flat band; the funnel window must sit inside both. The in-progress final month is partial by construction.

## f0b_method_history

Window: the full call history, all initiation methods, by month. Why: the method mix shifted hard over time (recent outbound volume is a fraction of its lifetime share; queue-transfer legs have their own era). A window read across a mix break compares different processes. How to read: find the era boundaries; check no analysis window straddles one. Companion to f0, which is inbound-only.

## s1_balance_by_bucket

Window: the newest account snapshot, charged-off accounts excluded. Why: the per-account dollar input for any sizing arithmetic - a captured payment is one cycle's payment but the balance at risk is the account balance, and average DQ1 balance was an open input until now. How to read: median under average means a heavy tail; use the median for a conservative per-account figure. Caveat: t2_dpd_buckets keeps charged-off stock in, which is why its deep-bucket balances differ.

## s2_roll_matrix

Window: the last 12 complete account months, consecutive-month account-level transitions, already-charged-off accounts excluded from the base. Why: the measured migration rates - where each bucket actually goes next month - replacing pooled balance-flow percentages with account-level ones. This is the multiplier ladder a provision-shaped read needs. How to read: pct_to_current is the cure rate; chain (1 - pct_to_current - pct_improved) down the ladder for a survival-to-loss estimate per bucket. Caveat: month-end grain hides intra-month churn; program re-ages sit inside pct_to_current (see v8).

## s3_payment_size_by_bucket

Window: the last 6 complete account months; a delinquent account-month counts as paying when its month-end last-payment date falls inside that month. Why: the per-capture dollar value - what an incremental captured payment is actually worth per bucket. How to read: multiply an incremental capture count by the median payment for a floor-style dollar figure; the average is drawn up by large one-off payments. Caveat: the month-end snapshot keeps only the LAST payment - a proxy, one payment per month.

## s4_caller_payment_lift

Window: 6 complete account months ending two months before the account table's newest complete month (every observation keeps a full following month). Why: the self-cure baseline. Non-caller delinquent account-months that pay next month anyway are what any capture claim must beat; the caller column on top of it is raw association - callers self-select, so the gap is an upper bound on lift, not a causal effect. How to read: quote the non-caller column as the baseline; treat caller-minus-noncaller as directional only.

## h1_language_vs_payment

Window: one complete account month (anchored two months before the account table's newest month so the 30-day payment window is fully covered). Why: the lexicon validation - if calls WITH payment/plan language do not pay at a visibly higher rate than calls without it, the funnel's language gate is noise and the transcript read needs better features before anyone scales it. How to read: the no-transcript row doubles as a bias check; a payment-rate gap between 'a' and 'b' is the signal the whole language layer stands on.

## h2_first_payment_mention

Window: one complete account month (one before the account table's newest). Why: initiative. Customer-first payment mentions are arriving intent; agent-first mentions are pivots. The split sizes how much of capture depends on agent behaviour vs customer intent - which is the difference between a coaching play and a routing play. Caveat: transcribed delinquent calls only.

## h3_language_by_ladder

Window: one complete account month (one before the account table's newest). Why: extends v9's two-row split to the full ladder - where payment talk gives way to hardship and escalation talk as delinquency deepens. That tone shift marks the buckets a capture play can still reach vs the ones that need a hardship route. Caveat: lexicon proxy; rates are per transcribed call.

## h4_call_effort_by_bucket

Window: one complete account month (one before the account table's newest). Why: the cost side - handle seconds from the call log, turns and minutes from the transcripts, per bucket. Long deep-bucket calls that still leak are effort spent without capture; pair with f6 to see where effort and leakage stack. Caveat: minutes come from the last utterance timestamp; calls without transcripts show call-log time only.

## n1_discriminative_bigrams

Window: one complete account month of delinquent inbound calls, split into paid-within-30-days vs leaked. Why: learns the lexicon from outcomes instead of hand-writing it. Customer bigrams are ranked by how disproportionately they appear on leaked vs paid calls (and the reverse); the survivors are candidate rules for a live transcript read, each arriving with measured support and lift. How to read: leak-markers with high support are review-queue triggers; payment-markers sharpen the intent gate. Caveats: bigrams need 250+ calls of support; filler-word pairs are dropped; correlation with the outcome, not causation - vet each phrase before it becomes a rule.

## n2_sentiment_vs_payment

Window: one complete account month of delinquent inbound calls. Why: the transcripts already carry model-produced NLP - a call-level customer sentiment score. This tests whether that free signal predicts payment beyond the lexicon. How to read: flat payment rates across sentiment bands mean mood adds nothing to intent - worth knowing before anyone builds on sentiment; a slope earns it a slot as a funnel feature. The unscored band doubles as a coverage check.

## n3_intent_score

Window: one complete account month of delinquent, transcribed inbound calls. Why: the pre-production question for the live read - do deterministic signals COMBINED rank calls? The score adds payment language (+1), plan language (+2), customer-first initiative (+1), positive final third (+1), and subtracts hardship (-1) and escalation (-1). How to read: a monotonic climb of payment rate with score is the evidence that a rules engine can already rank calls, and that a proper model read starts from a proven floor. Caveat: the weights are asserted, not fitted - this is a floor, not a model.

## n4_agent_offer_vs_outcome

Window: one complete account month; delinquent calls where the CUSTOMER talks payment. Why: the first measured read of the recoverable slice - customer intent present, and the agent either put a plan/arrangement on the table or did not. The payment-rate gap between the two groups sizes how much leakage looks like a missed offer rather than a missing customer. Caveat: association, not causation - agents may offer more when capture already looks likely; treat the gap as an upper bound and audit a sample before quoting it.
