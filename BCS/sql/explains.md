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

Window: inbound calls joined to the caller's same-month bucket, payment read from the call month's OR the next month's last-payment date (a next-month-only check misclassifies same-month payers as leaked - the fix that moved every capture number in the kit). Calls span the 5 complete account months BEFORE the account table's newest complete month, so every call keeps a full following-month snapshot for the 30-day check. Why: the capture/leakage proxy at scale - delinquent calls followed by no payment within 30 days. Caveats: month-end grain and paymt_last_dt parsing still apply; leg grain (per call), while the f-series runs per episode - chatty accounts weight these rates differently.

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

Window: PINNED to call months 2024-07 through 2025-06 (W3) so every episode keeps at least 8 account months of outcome runway before the account copy's edge; outcomes read through 2026-02. Run f0 first - if any W3 month shows broken coverage or fill, move the window before trusting this. Why: the headline. Episodes are first-inbound-per-account-per-day (the reference dedup); each row is a cumulative gate, so the last row is the leakage pool that actually charged off. The capture gate is CLEAN and two-sided: a payment counts from the call month's or the next month's snapshot, and autopay-dated or NSF-marked payments do not count (s6 sizes both contaminations; the fallback file drops the autopay/NSF exclusion if those columns are absent - say so when quoting). episodes_strict reruns the language gate with the strict lexicon (plan / settlement / future-dated promise) so the pool reads as a band [strict, loose] - the loose payment-word rate on collections calls is high by construction. Balance is the account's month-end balance at its first matched episode, held constant down the funnel - the dollar column is a true waterfall. Scope: business-card legs excluded, blanks kept, no partner exclusions (f4_scope_split sizes the choice). How to read: stage h is the leakage candidate pool; stage i is the slice that became realized loss (x3 validates the 8-month horizon). Caveats: payment is a snapshot proxy, not a ledger join; row b makes the account-id gap explicit and f4_match_by_auth splits it by authentication; f4_bucket_drift bounds the month-grain bucket error.

## f2_funnel_by_month

Window: same pinned W3 as f1, one row per call month. Why: kills the single-month objection - if the stage-to-stage drops hold vintage by vintage, the pooled waterfall is not an artifact of one odd month. How to read: columns are cumulative episode counts; divide any column by the previous one for the per-month gate rate. Caveat: later call months have shorter absolute follow-up beyond the guaranteed 8; chargeoff_8m stays comparable because the 8-month horizon is fixed per episode.

## f6_funnel_by_bucket

Window: same pinned W3 as f1, delinquent gates split by the caller's same-month bucket. Why: early buckets carry the episode volume, late buckets carry the loss probability - this shows where episodes x loss actually concentrates, which decides whether a capture play scopes to DQ1-3 or chases the deep ladder. How to read: divide chargeoff_8m by no_payment_30d per bucket for the leaked-to-loss conversion; compare against v5's payment rates.

## f7_leaked_dollars_by_month

Window: same pinned W3 as f1. Why: the dollar trend behind the funnel end - leaked balances and the charge-off dollars that followed, month by month. This is the chart-shaped version of the story for anyone who will not read a waterfall. Caveat: an account that leaks in several months appears in each; balances are month-end positions, not exposure-at-loss.

## f3_funnel_dollars

BACKWARD / AUDIT FRAME - size and reconcile on this view; never read rates off it (it conditions on the loss outcome). Window: charge-offs 2024-08 through 2026-02; calls observed 2024-07 through 2026-01 (each call needs a same-month snapshot and a 30-day payment runway inside the account edge). Why: flips the funnel around to the loss ledger - of realized charge-off dollars, how much sat on accounts the inbound channel touched and did not capture? The funnel-end share is the addressable slice of actual losses. pct_dollars_bk_dcsd_fraud is the unsolvable slice (bankruptcy/deceased-like/fraud status text, a partial proxy) - quote funnel-end dollars net of it, because the client nets the same categories out of its own policy-loss number. How to read: compare avg_chargeoff across groups - do leaked accounts lose more per account than never-callers? Caveats: the last charge-off months see a shorter call lookback, and group c is 'no MATCHED inbound call' - accounts whose calls carried no account id land there, so it is an upper bound twice over.

## f5_calls_before_chargeoff

Window: DQ1 entrants 2024-07 through 2025-06 (3-month entry buffer) that charged off by 2026-02; inbound episodes counted from entry to charge-off. Why: the intervention surface on the loss path. '0 calls' losses are unreachable inbound; the 1-2 call bands are where a single missed capture is the whole story; the 11+ band is the repeat-friction narrative at volume. Caveat: episodes require an account id on the call, so the bands are floors.

## f4_match_by_auth

Window: last 5 complete account months (W2-equivalent), inbound calls vs the account master. Why: DPT-style bias gate - the funnel only sees matched calls, and if unmatched calls concentrate among FAILED-authentication callers, every funnel rate is biased toward the easy population. How to read: compare pct_matched_master across authentication outcomes; a large gap means quote funnel rates as 'of matched calls' and treat the unmatched slice as its own population. Caveats: blank authentication is its own signal, not noise; this gate runs on recent months while the funnel runs on W3 - the flat multi-year acctid-fill band (t1_acctid_fill_trend) is the argument that the split generalizes back.

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

Window: 6 complete account months ending two months before the account table's newest complete month (every observation keeps a full following month). Why: the self-cure baseline. A payment now counts if dated in the observation month or the next - symmetric for both groups, and it stops a call on the 3rd that collects on the 5th from reading as caller-did-not-pay. Non-caller delinquent account-months that pay anyway are what any capture claim must beat; the caller column on top is raw association - callers self-select, so the gap is an upper bound on lift, not a causal effect. How to read: quote the non-caller column as the baseline; treat caller-minus-noncaller as directional; x4 is the self-controlled version of this comparison.

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

## t3_unmatched_queues

Window: the last complete call month (bounds the text scan). Why: the quarter of calls without an account id vanish from every joined read - but transcripts join on contactid, so the invisible pool is directly readable. This shows which queues it rings AND how often the customer talks payment there. If collections-relevant queues run high and payment talk is rich, the funnel undercounts exactly its target population - the bias direction, measured, not assumed. How to read: pair with f4_match_by_auth - that says WHO fails to match, this says WHERE and WITH WHAT INTENT.

## t3_call_gap_days

Window: consecutive inbound calls per account, last 6 months of call data. Why: the cadence of repeat calling. Same-day and 1-3-day gaps read as friction (unresolved first attempt); month-scale gaps read as cycle-driven contact. Feeds the episode-window choice (see f9 for the funnel-window version) and sizes the callback burden.

## v10_cure_durability

Window: cures (bucket 1+ to current, consecutive months) observed over 4 account months ending 4 months before the copy's edge, so every cure keeps 3 complete follow-up months. Why: a cure that re-delinquents within a quarter is worth less than a durable one, and program re-ages masquerade as cures - this measures how much of the cure rate is real. How to read: high re-delinquency from deep buckets marks re-age-like resets (pair with v8); low re-delinquency from bucket 1 validates using cure rates in the value arithmetic.

## v11_multi_vintage_roll

Cohorts: DQ1 entrants in 2024-09, 2025-01, and 2025-05 (3-month entry buffer), tracked to month-on-book 9. Why: one vintage is one seasonal draw - if three curves spread across the funnel window agree, the migration rates are stable enough to quote; if they diverge, the divergence is the finding. How to read: compare pct_charged_off at the same month-on-book across vintages; v2 remains the single-vintage deep read with the 4-state split.

## f8_payment_window_sensitivity

Window: same pinned W3 chain as f1, through the language gate. Why: the leakage pool depends on one assumption - no payment within 30 days. Rerunning the gate at 7/14/30/60 days on the same episodes shows whether the pool is robust or full of slow payers; the 7-day read is the tight bracket the value arithmetic wants next to the 30-day one. How to read: the 30-to-60-day drop is the share of 'leaked' that is actually just late; quote the funnel with that band attached. Caveat: payment checks read the call month's and following months' snapshots (raw-payment gate, no autopay/NSF exclusion here - f1 carries the clean gate).

## f9_episode_chaining

Window: same pinned W3 episodes as f1. Why: the funnel's episode unit is one account-day; this measures how those days cluster. A large 1-3-day share means multi-day chains should collapse into one episode (and the funnel's episode counts overstate independent contacts); a long-gap profile validates the per-day unit. The open episode-window question, answered with data instead of an assumption.

## s5_balance_roll_matrix

Window: same 12 account months and transition logic as s2, dollar-weighted. Why: provision-shaped reads multiply balances, not account counts, and big-balance accounts do not roll like small ones. How to read: where pct_bal_charged_off exceeds s2's pct_charged_off, losses concentrate in larger accounts - the per-account average understates the dollar risk.

## h5_triage_language

Window: one complete account month (payment-runway anchored), delinquent transcribed calls. Why: not every delinquent call is a collections conversation - dispute/fraud and servicing-only calls should exit the funnel before any intent read. This sizes the triage gate and shows each group's payment rate. How to read: the servicing-only share bounds how much the language gate could over-drop; dispute calls' payment rate says whether they belong in the pool at all.

## h6_promise_language

Window: one complete account month (payment-runway anchored), delinquent transcribed calls. Why: the promise-to-pay table lives in another system, but promises are audible. Future-dated promise language followed by no payment is the broken-promise proxy - the closest thing to a kept-rate the three tables can produce. How to read: the promise row's no-payment share is the audible break rate; compare against the no-promise payment-talk row to see whether a promise means anything.

## n5_early_intent

Window: one complete account month (payment-runway anchored), delinquent transcribed calls, customer turns split into thirds. Why: a live transcript read must act mid-call. If first-third intent already separates payment rates, routing/prompting can happen while the customer is on the line; if the signal only forms late, the play is post-call follow-up. This is the latency evidence for the live-read design.

## o1_capture_by_queue

Window: account-clock (calls in the 5 complete account months before the newest complete month, payment from the following snapshot). Why: leakage as a place - queues with high delinquent volume AND high no-payment rates are where the funnel's stage-g episodes physically sit. How to read: volume times no-payment rate ranks the queues by leaked episodes; a queue with low volume but extreme rates is a process bug, not a priority.

## o2_capture_by_vendor_site

Window: same account-clock read as o1, split vendor x site, RESTRICTED to buckets 1-3 so the comparison is mix-controlled at least on delinquency depth. Why: if payment-after-call spreads across sites on a similar call mix, handling quality moves money (coaching headroom); if it does not, the leakage is process-shaped and coaching will not fix it. Caveat: queue mix still differs by site, and vendor comparisons get quoted politically - treat any gap as a lead to audit, never a scorecard. Context card for that reason.

## o3_delinquent_transfer_paths

Window: same account-clock read, transfers chained via initialcontactid (sparsely filled - lower bound). Why: the transfer-drop leakage cause, sized: are delinquent calls transferred more than current ones, and does a transferred delinquent call pay less afterwards? How to read: compare rows c vs d for the delinquent transfer penalty; rows a vs b give the current-caller baseline for the same comparison.

## o4_delinquent_abandons

Window: one complete account month; the 7-day recontact search runs on the call table alone so it can cross the month edge. Why: an abandoned delinquent call is intent that never reached a human - the purest friction leakage. How to read: pct_abandoned by bucket sizes the loss at the door; low recontact in deep buckets means the lost intent stays lost.

## x1_leaked_vs_captured_roll

Window: same pinned W3 chain as f1 through the language gate; outcome read three months after the call month. Why: the value story needs capture to be worth more than one payment - if captured episodes sit visibly better (more current, fewer charged off) at month three, capture is balance-shaped money. Caveats: association (payers differ from non-payers); 'not visible' rows are accounts that left the snapshots.

## x2_repeat_leak_chains

Window: same pinned W3 leaked episodes (stage-g definition); chains counted within 90 days of each account's first leak. Why: one leaked episode can be bad luck; a chain is a process failure. If charge-off climbs with the leak count, repeat leakage compounds - and the first leak is the intervention point, which is exactly what a live read would target. How to read: multiply band sizes by their charge-off rates for the chain-attributable loss share.

## x3_time_to_chargeoff_after_leak

Window: same pinned W3 leaked episodes; months from each account's FIRST leak to its charge-off date, for charge-offs observed inside the account window. Why: the funnel counts losses within 8 months of the call - this shows whether that horizon holds. How to read: the a-d bands are inside the horizon; a fat e band (9-12) means the funnel undercounts. Caveat: right-censored - later first-leaks have less observable runway, so read the shape, not the tail.

## f4_scope_split

Window: same pinned W3 delinquent episodes as f1, WITHOUT f1's business-card exclusion - that is the point. Why: the dollar frame this funnel reconciles against is consumer-card scoped, but the call table carries business-card legs, a large blank producttype share, and partner traffic that cannot be excluded yet (values unverified). This makes the scope error a measured number: the business-card row is exactly what f1's exclusion removed, and if the blank row behaves like the consumer row (similar language and leak rates), keeping blanks is safe. How to read: a blank row that behaves differently means the funnel needs a scope caveat before any reconciliation to policy-loss numbers.

## f4_bucket_drift

Window: one complete account month, snapshot-grain (the account table has ~14 positions a month; every other query uses the month-max bucket). Why: the kit's grain check - a call on the 2nd can be tagged with a bucket the account only reached on the 28th, after the call. This measures how often the as-of-call-date bucket disagrees with the month-max tag, per bucket. How to read: pct_asof_current on bucket-1 rows is the share of 'delinquent episodes' that were actually pre-delinquency calls - the funnel's stage-e overcount; small numbers validate the month-grain kit, large ones are its error band. Worst near the charge-off line, where one month moves an account across it.

## s6_payment_contamination

Window: last 6 complete account months, delinquent account-months with a payment dated in-month. Why: the capture gate's correction band. A payment that is just autopay firing is not a capture, and an NSF-marked payment is not money - this sizes both, by bucket, and doubles as the existence probe for the autopay/NSF columns (if it errors, they are absent in this copy and f1's fallback gate is the documented route). How to read: pct_payment_autopay_dated is the share of 'captures' the clean gate removes; if it is material, every raw-gate capture number in tiers 4/7/8/9 is overstated by roughly that share. Run BEFORE quoting the funnel.

## o5_capture_by_auth

Window: account-clock (calls in the 5 complete account months before the newest complete month). Why: the verification-block leakage cause - a delinquent caller who cannot get through authentication cannot pay on the call, and if failed-auth callers pay visibly less after, the block is costing money the customer wanted to hand over. That is process leakage, not willingness leakage: the most recoverable kind. How to read: pairs with f4_match_by_auth - that gate measures who is INVISIBLE to the funnel; this measures what happens to the visible-but-blocked.

## o6_coll_volume_monthly

Window: last 6 complete call months, inbound calls on queues whose name starts COLL. Why: the control total. The funnel's inbound universe is a candidate LOWER BOUND until this table's counts reconcile against the ops-reported call volumes; this produces the number to reconcile, with acctid fill beside it. How to read: compare offline against the workforce-management monthly volumes - a large gap means calls exist that this table does not carry, and every funnel count inherits that caveat (queue-name scoping is itself partial; the gap is directional).

## x4_within_account_contrast

Window: same pinned W3 chain as f1 through the language gate; only accounts with BOTH a captured and a leaked intent episode in the window. Why: the causality answer the caller/non-caller splits cannot give - within the same account, the who-calls selection effect cancels, so the after-capture month vs after-leak month difference is the closest these tables get to capture's real effect. How to read: if pct_current_next_month is visibly higher after captured episodes than after leaked ones FOR THE SAME ACCOUNTS, capture moves the account, not just one payment. Caveats: still not randomized (episode timing within the spell differs - captures may come earlier in the spell); episodes need a next-month snapshot.

## x5_payment_latency

Window: same pinned W3 chain as f1; captured intent episodes only. Why: validates the 30-day window from the inside and prices the live-assist case - days from call to payment splits 'paid on the call' (same-day to 3 days) from 'paid eventually' (the 15-30 tail, money that arrived but maybe not because of the call). How to read: the same-day-to-3-day share is the on-call capture rate; a fat 15-30 tail argues for tightening the funnel's gate toward f8's 7- or 14-day read. Caveat: the snapshot keeps one payment date per month, so latency is the LAST qualifying payment's - a floor on speed.

## b1_cohort_ladder

Window: 202412 and 202501 snapshots only. Why: the headline bridge - the Athena 493,139 (month-MAX DQ1 entrants, v11) and Ishant's SAS/ASP 186,412 (month-END DLNQT_CD=1) sit on different grains, and this ladder walks from one to the other one condition at a time, with accounts and month-end balance at every step. How to read: step d is our v11-comparable count (strict-December prior; b2 closes the residual to 493,139); steps g/h are the Ishant-comparable month-END reads; h1/h2/h3 split h into new roll vs no-prior-record vs already-delinquent; step i is the any-code entrant total to set against his 433,914 grand total. Caveats: grain is stated in every step label - month-MAX is the worst bucket on any snapshot in the month, month-END is the bucket on the last snapshot date in the month, which need not be the calendar end of month (b3 tests that). Ishant's side additionally filters cpc='others' and CHRGOFF_RSN in (blank, PLY); Athena has no CLNT_PRDCT_CD or charge-off-reason column, so those two exclusions cannot be mirrored here and sit with him as counter-asks.

## b2_entry_lookback

Window: 2024-06 through 202501 snapshots, entrant candidates at 2025-01 with month-MAX bucket 1. Why: b1 step d uses a strict 202412 prior, but v11's 493,139 uses lag over PRESENT months - an account last seen in September 2024 gets its September bucket as the prior. This query runs both definitions side by side and splits the difference. How to read: row a should reproduce v11 exactly; a minus b is accounts whose last sighting predates December; row d is the pure no-history slice within that. Caveat: the lookback floor is 2024-06, same as v11 - an account absent since before June counts as no prior row under both definitions.

## b3_snapshot_cadence

Window: 202501 snapshots only. Why: the whole month-END grain rests on max_by(bucket, eff_dt) - the LAST snapshot date per account in the month - and this measures whether that last date is one shared calendar month-end or a scatter of days. How to read: part 1 counts accounts by how many snapshot dates they carry in January; part 2 lists the top 15 last-dates by accounts. One dominant last-date near 20250131 means our month-END read is Ishant-comparable; a scatter means "month-end" is account-specific and part of the bridge gap can live here. Caveat: part 2 is capped at the top 15 last-dates; a long tail of rare dates is itself a cadence finding.

## b4_eom_bucket_by_entry

Window: 202412 and 202501 snapshots; 202501 accounts with month-END bucket >= 1. Why: Ishant's DLNQT_CD_M1 distribution has heavy deep entries (M1=5 at 42,257) - on a pure days-past-due ladder a new-roll entrant can only land at month-END bucket 1, so where our deep buckets come from tests whether his DLNQT_CD and our past-due ladder mean the same thing. How to read: within each month-END bucket, class a is new rolls (202412 EOM = 0), class b is already-delinquent carryover, class c is no December row. Deep buckets should be almost entirely class b; any material class-a rows at bucket 2+ mean the ladder and ASP DLNQT_CD disagree on semantics, and his deep entries must be class-c-like or a different code meaning. Caveats: grain is month-END on both sides of the split; the cpc and charge-off-reason filters on Ishant's pivot are not mirrored here.

## b5_co_scope

Window: 202501 snapshots only. Why: Ishant's 433,914 grand total includes a straight-to-CO class alongside the DQ-code entrants; these probes size the charge-off populations Athena can see in the same month. How to read: row a is the straight-to-CO analogue (charged off in January with month-MAX bucket 0 all month); row b is charge-off while delinquent; row c is pre-2025 charge-off stock still carrying January rows - a perimeter difference if his view drops it; row d is deep-bucket accounts with no charge-off date, the reverse mismatch. Caveats: chrgoff_dt is the only charge-off signal here - there is no charge-off reason column, so his CHRGOFF_RSN in (blank, PLY) filter cannot be applied; month-MAX and month-END grains are stated per row.

## b6_m2_roll_mirror

Window: 202412 through 202503 snapshots; the January month-END bucket-1 stock only. Why: mirrors Ishant's DLNQT_CD_M1=1 x M2 x M3 pivot - where his 186,412 sit in February and March (M2=0: 105,715 cured / M2=1: 15,002 / M2=2: 65,585, of which M3=3: 41,975 / deeper tiny). If our shares match his, the two ladders agree on roll dynamics even where the levels differ. How to read: one row per (Feb state, Mar state, entry class); 'co' means the charge-off date lands by the end of that month's window, 'gone' means no row that month; balance is the January month-end balance held constant. Compare SHARES not counts - his base is stock under ex-AA and CHRGOFF_RSN filters, ours is the unfiltered 207,006. Caveats: grain is month-END on all three months; a mid-month cure-and-relapse inside February is invisible to both sides here.

## b7_call_overlay

Window: account side 202412-202501 snapshots; call side January 2025 inbound legs, business-card excluded, episode = first inbound leg per account per day (f1's convention). Why: the month-max lens payoff - customers call on intra-month delinquency, and class a (delinquent during January, current again by month-end) is the population ASP month-end never sees. If its call rate is material, the month-max companion lens is justified for the call story. How to read: one row per bridge class with account base, distinct callers, episodes, and the percent of accounts that called; set class a's rate against classes b and c. Caveats: pre-2025 charge-off stock is excluded from all classes; the call join inherits the acctid fill gap (unlinked legs count nowhere); class d mixes deeper entrants and deeper stock - it is the residual, not a clean cohort.
