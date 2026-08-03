# UC2 transcript coding: the prompt and the full picture

Working document. Version 4 of the coding prompt, plus how its output becomes the numbers.

Date: 2026-08-03.

---

## 1. What this produces

Two things, from one pass over the same transcripts.

**A. The leakage number.** How many accounts rolled from DQ1 to DQ2 despite telling an agent they wanted to pay, and how many of those an AI assistant could plausibly have converted. This feeds the SI business case.

**B. The product spec.** What the assistant would actually put on the agent's screen, ranked by how often each type is needed. This is what Chris and Jay will engage with, and it costs nothing extra because it falls out of the same read.

---

## 2. The population, and why it is simpler than it looks

| Figure | Meaning |
|---|---|
| **3,292** | Inbound callers who rolled DQ1 to DQ2 |
| **2,917** | The clean single-cycle subset: DQ0 to DQ1 to DQ2 |

Two consequences, both good:

- **Every account in the frame already called.** There is no separate "did they contact us" gate to measure or estimate.
- **Every account rolled, so none of them paid.** Leakage does not need a payment-outcome join. **Rolling to DQ2 is the non-payment.** Intent-positive and rolled means leaked, by construction.

That removes the two weakest links from the chain. What is left to measure is: was the call readable, did they show intent, and could an assistant have changed it.

`[OPEN]` Decide which figure is the reporting denominator. 2,917 is the cleaner story (one cycle, one roll, no prior delinquency). 3,292 is the fuller population. Pick one and state it on every slide.

### The unit problem, and the rollup rule

The population counts **accounts**. The coding unit is **calls**. Observed on the 30-account sample: 30 accounts carried 41 contact IDs, so roughly **1.37 calls per account**.

So a 300-account sample is about **410 transcripts to code**.

**Rollup rule: an account is intent-positive if any of its calls shows intent.** That matches how an assistant works, since it only has to fire once. Apply the same rule to actionability.

### Sampling: stratify, do not draw at random

A random 300 gives a tight intent rate and a loose actionable rate, and actionable is the number the case runs on:

| Stage | n | Interval |
|---|---:|---|
| All sampled accounts | 300 | intent rate ±5.4 pts |
| Usable | ~225 | |
| Intent-positive | ~90 | **actionable rate ±10 pts** |

**Better: draw 150 from calls carrying payment language and 150 from the rest**, using the existing full-population lexicon to define the strata. Same effort, roughly double the effective sample on the actionable question, because the reads are concentrated where intent-positive calls live.

Two conditions, both mandatory:
1. Record the stratum on every row.
2. Weight back by each stratum's true population share, or the pooled rate is wrong.

---

## 3. The prompt, v4

Single message. Instructions, then the packet, then a short restate. The restate is what stops the format rules being attenuated behind a long paste.

````text
Read each collections call transcript below. Every customer was one payment behind
and later rolled to the next delinquency bucket. Turns appear as "CUSTOMER:" and
"AGENT:". Names appear as [NAME]. Some digits are masked. The audio was
machine-transcribed and is imperfect.

Some of these calls were placed by the bank's automated dialer and still record as
inbound. If the agent says why the bank is calling, the bank called.

Quote verbatim, or write exactly: none visible
Cap every quote at 15 words. Replace any tab or newline inside a quote with a space.

NEVER leave a cell empty. If a field does not apply, write: NONE
Every row must have exactly 13 tab-separated cells.

Read all calls in the packet before writing anything.

===============================================================
STEP 1. IS THERE ANYTHING TO READ
===============================================================

USABLE: YES / NO
  NO if the call disconnected, or there are fewer than 4 customer turns, or the
  customer never responds, or the money talk is garbled beyond following.
  A short call that is complete and followable is USABLE.

  IF USABLE IS NO: write NOT SCORED in every remaining column and stop.

WHY_NOT_USABLE: NONE if USABLE is YES. Otherwise pick one:
  DISCONNECTED MID-CALL / NO CUSTOMER RESPONSE / UNDER 4 CUSTOMER TURNS /
  GARBLED / OTHER: <3-5 words>

===============================================================
STEP 2. DID THEY SHOW INTENT
===============================================================

INTENT_QUOTE: a line spoken by the CUSTOMER. Never an agent line. If the agent
  raised it and the customer only agreed, quote the customer's agreement.
  If the customer says nothing about money, write: none visible

INTENT_WHY: why you chose the value below. Max 15 words. Write this before the value.

INTENT, pick one:
  COMMITTED            will pay AND named an amount or a day
                       e.g. "I'll pay the 50 on Friday"
  WILLING              said they can or want to pay, named neither
                       e.g. "I want to take care of this"
  ASKED HOW TO PAY     seeking a way to clear it: a programme, a plan, options,
                       what would resolve it. Never said they would pay
                       e.g. "are there any hardship programmes"
  CANNOT PAY           says they are unable to pay, now or at all
                       e.g. "I don't have any money right now"
  ASKED ABOUT ACCOUNT  balance, a statement, a charge. Not trying to resolve it
                       e.g. "what's my balance"
  NOT ABOUT MONEY      money never came up

  Boundary rules:
  - Said they can or want to pay? WILLING, even with no amount and no date.
  - Only asked for a route to clear the debt? ASKED HOW TO PAY.
  - Said they are unable? CANNOT PAY, even if they were asking about options.
  - Only wanted information, not resolution? ASKED ABOUT ACCOUNT.

EXCLUDE_FLAG: DISPUTE OR FRAUD / DECEASED OR ESTATE / THIRD PARTY / NONE
  Set whenever it applies, whatever the intent value.
  Consistency rule: if WHAT_STOPPED_IT is DISPUTING THE CHARGE then this is
  DISPUTE OR FRAUD. If WHAT_STOPPED_IT is NOT THE CARDHOLDER then this is
  THIRD PARTY or DECEASED OR ESTATE.

PAID_ON_CALL: YES / NO / NOT VISIBLE
  YES only if a card or bank detail was given, or a payment was taken or booked on
  this call. "I'll check online" or "I'll pay Friday" is NO.

===============================================================
STEP 3. WHAT STOPPED THE PAYMENT
===============================================================

WHAT_STOPPED_IT, the customer-side reason, pick one:
  NO WAY TO PAY ON FILE            no card or bank account saved, or it failed
  NO MONEY RIGHT NOW               says they cannot pay now
  WANTED AN OFFER WE DO NOT HAVE   plan, settlement or hardship not available yet
  LOCKED OUT ONLINE                login, password, app, or OTP to an old number
  NOT THE CARDHOLDER               third party, no authority, or estate
  DISPUTING THE CHARGE             fraud or dispute in progress
  AGENT NEVER ASKED                payment was relevant and nobody raised it
  PAYMENT NOT RELEVANT             this call was correctly not about collecting
  NOTHING STOPPED IT               they paid, or booked it
  OTHER: <3-5 words>

STOPPED_QUOTE: a CUSTOMER line showing it, or: none visible
  If only an agent line shows it, prefix the quote with AGENT:

===============================================================
STEP 4. COULD AN AI HAVE HELPED
===============================================================

An AI assistant listens to this call live and can put ONE short message on the
human agent's screen. The goal is to help the agent take a payment on this call.

SCREEN_TEXT: write the exact message that would have appeared on the agent's
  screen. Under 12 words. Be specific and use facts visible in this call.
  Good:  "Hardship programme opens 24 Jan. Offer to book a callback."
  Good:  "Checking account ending [MASK] is on file. Offer it now."
  Good:  "Offered amount is below the minimum due. Ask for the full amount."
  Bad:   "Provide hardship eligibility details."   (not a message, too vague)
  Bad:   "Help the customer pay."                  (says nothing)
  If you cannot write a specific, useful message from what this call shows,
  write exactly: nothing

  THE CONTROLLING RULE: if SCREEN_TEXT is "nothing", AI_COULD_HELP must be a NO.
  If you can write it, it is YES or MAYBE.

AI_COULD_HELP, pick one:
  YES               that message would likely have got the payment
  MAYBE             it might have helped, the text does not make it certain
  NO, ALREADY DONE  the agent already did the right thing, no assist needed
  NO, NOTHING       nothing on a screen would have changed this call

ASSIST_TYPE: choose the value matching the message you wrote in SCREEN_TEXT.
  Spell it EXACTLY as below, with spaces, never underscores.
  PROMPT TO ASK             the agent never asked for money anywhere
  PROMPT TO CLOSE           intent was shown and acknowledged, but the agent
                            never asked for a card or a date
  SURFACE A PAYMENT METHOD  a saved method existed and was not offered
  SURFACE ELIGIBILITY       a fee waiver, or the date a programme opens
  SURFACE THE NEXT STEP     the right web path, department, or document
  FLAG A WEAK PROMISE       a promise was made and something about it is weak
  NOTHING                   no assist applies
  OTHER: <3-5 words>

  If SCREEN_TEXT asks the agent to get a date or an amount, it is PROMPT TO CLOSE.
  FLAG A WEAK PROMISE requires a promise to exist in this call.
  Every other value requires its precondition to be visible in the text.

===============================================================
OUTPUT
===============================================================

Put your ENTIRE reply inside ONE fenced code block. Nothing before it, nothing
after it. No prose, no headings, no commentary. Tab-separated. Never use a tab
inside a cell. Never leave a cell empty.

Line 1 exactly:
RECEIVED nonce=K7QX n=6 contact_ids=<the contact ids, in the order given>

Line 2 is this header, exactly:
CONTACT_ID	USABLE	WHY_NOT_USABLE	INTENT_QUOTE	INTENT_WHY	INTENT	EXCLUDE_FLAG	PAID_ON_CALL	WHAT_STOPPED_IT	STOPPED_QUOTE	SCREEN_TEXT	AI_COULD_HELP	ASSIST_TYPE

Then one row per call, in the order given.

Last line exactly:
===END nonce=K7QX rows=6===

===============================================================
THE CALLS
===============================================================

=====BEGIN PACKET P07 nonce=K7QX n=6=====

--- SEQ 1/6 | contact_id=<id>
<transcript>

--- SEQ 2/6 | contact_id=<id>
<transcript>

=====END PACKET P07 nonce=K7QX n=6=====

RESTATE, do not skip:
- First line exactly: RECEIVED nonce=K7QX n=6 contact_ids=<contact ids in order>
- CONTACT_ID is the contact id, never the account id
- Entire reply inside one fenced code block, tab-separated, exactly 13 cells per row
- No cell is ever empty. Write NONE if a field does not apply
- USABLE=NO means NOT SCORED in every remaining column
- INTENT_QUOTE is a customer line only
- INTENT_WHY comes before INTENT. SCREEN_TEXT comes before AI_COULD_HELP
- SCREEN_TEXT is an actual message to an agent, under 12 words, or: nothing
- ASSIST_TYPE spelled with spaces, never underscores
- Use only the text between the BEGIN and END markers
- Last line exactly: ===END nonce=K7QX rows=6===
````

### What changed from v3

| Change | Defect it fixes |
|---|---|
| `NEVER leave a cell empty, write NONE`, and 13-cell rule | **Column shift.** v3 rows slid one column left when `WHY_NOT_USABLE` was blank, putting `ASKED HOW TO PAY` in `INTENT_WHY` and `DISPUTE OR FRAUD` in `INTENT` |
| `INTENT_QUOTE` customer-only stated in the field itself | v3 quoted an agent line, "Programs or options may be available" |
| `CANNOT PAY` added to `INTENT` | v3 had no home for "I don't have any money right now" and dumped it in `ASKED ABOUT ACCOUNT` |
| `EXCLUDE_FLAG` consistency rule | v3 had `DISPUTING THE CHARGE` with `EXCLUDE_FLAG = NONE` |
| `ASSIST_TYPE` derived from `SCREEN_TEXT`, exact spelling | v3 tagged "ask for payment date" as `FLAG A WEAK PROMISE`, and emitted both `SURFACE_ELIGIBILITY` and `SURFACE ELIGIBILITY` |

---

## 4. How to run it

| | |
|---|---|
| Packet size | Start at 6 calls. Ramp to 8, 10, 12 and stop one below where the echo count fails or the last rows degrade |
| Session | Fresh chat every packet, no exceptions |
| Nonce | Fresh 4-character code per packet, in both places |
| Save | Copy the raw reply verbatim to a file before parsing anything |
| Reject | Wrong nonce, wrong count, missing id, or a row without 13 cells. Discard before reading any labels |
| Gold check | Four known calls as the first packet of every session. More than two cells differ from the frozen answer, stop the session |

Log per session: date, packets, calls sent, rows back, rejects, gold cells wrong, and the running intent rate. If the running rate moves more than 10 points between consecutive sessions, stop and hand-read five rows.

---

## 5. From output to findings

### The funnel

```
3,292   inbound callers who rolled DQ1 to DQ2        [population]
  ×  usable_rate            transcript readable      [measured]
  ×  intent_rate            showed intent            [measured]
  =  LEAKED ACCOUNTS        intent shown, rolled anyway
  ×  actionable_rate        an assist could have helped  [measured]
  =  ADDRESSABLE ACCOUNTS
```

**Leaked needs no payment join.** They rolled, so they did not pay. That is the whole definition.

### The formulas

Compute on accounts after the rollup, not on calls.

```
usable_rate       = USABLE = YES                          / all sampled
intent_rate       = INTENT in (COMMITTED, WILLING, ASKED HOW TO PAY)
                                                          / USABLE = YES
leaked            = intent-positive accounts              (all rolled, so all leaked)
actionable_rate   = AI_COULD_HELP in (YES, MAYBE)         / leaked
firm_rate         = AI_COULD_HELP = YES                   / leaked
no_headroom_rate  = AI_COULD_HELP = NO, ALREADY DONE      / leaked
disconnect_rate   = USABLE = NO                           / all sampled
```

`CANNOT PAY` is **not** intent-positive. It is reported separately as the genuinely uncollectable share.

**Report every rate twice: all rows, and `EXCLUDE_FLAG = NONE` only.** The second reproduces Anupam's own exclusion, which is half the reason his figure reads higher than Zenon's.

**Report actionability as a range, not a point.** `firm_rate` is the floor, `actionable_rate` is the ceiling. `MAYBE` runs near half the helpable rows and will not go away, because the question genuinely is uncertain from a transcript. The range is still tighter than the band the SI model carries today.

### Scaling

```
weight              = 3,292 / n_accounts_sampled        e.g. 300 → 10.97
population estimate = sample count × weight
```

If stratified, weight each stratum separately by its own population share and sum. Do not pool.

### The value chain

```
addressable accounts
  ×  lift          0.30      [assumption, pilot only]
  ×  realization   0.244     [from the value model]
  ×  net NCL per prevented roll  = gross 12-month loss / cohort size × 0.79
  =  annual NCL avoided
```

Illustrative, using placeholder rates until the 300 lands:

| | |
|---|---:|
| Population | 3,292 |
| × usable 0.75 | 2,469 |
| × intent 0.40 | **988 leaked** |
| × actionable 0.60 | **593 addressable** |
| × 0.30 lift × 0.244 realization | 43 saves |
| × ~$2,085 net per save | **~$90K/year** |

At the ceiling (all leaked addressable) it is about $151K. **Against a run cost of $60-90K a year, this is thin, and that is the honest finding.** Every rate above is now measurable, so replace them as the sample lands rather than arguing about them.

`[VERIFY]` The $2,085 net-per-save figure comes from the Jan-2025 cohort of 2,879. Recompute gross 12-month loss for the 3,292 frame before using it.

### The two outputs that need no extra work

**`ASSIST_TYPE` counts × weight** is the product roadmap, ranked by volume. If `PROMPT TO CLOSE` is 40% of addressable, that is the first thing to build. This answers "what does the product actually do", which is the question Chris, Jay and Venkat all ask.

**`SCREEN_TEXT` grouped by `ASSIST_TYPE`** is the real message library. Ten verbatim examples on a slide are worth more than any rate.

### The finding Anupam asked for

`disconnect_rate` is his own request from 2026-08-03, after he saw 8 of 30 and said *"if almost close to one-third of the accounts are getting a call disconnected, that's concerning."* Report it whether or not UC2 can address it, split by `WHY_NOT_USABLE`, and note that the 4-turn rule changed the definition since he saw that number.

---

## 6. Before coding 300

| # | Do | Why |
|---|---|---|
| 1 | Run v4 on the same 20 | Confirm no column shift and no empty cell. Check `CANNOT PAY` picks up the right rows |
| 2 | Decide 2,917 or 3,292 as the denominator | It moves every population figure |
| 3 | Fix accounts as the reporting unit, with the any-call rollup rule | Otherwise the denominator changes between slides |
| 4 | Draw the stratified 300, not a random 300 | Roughly doubles precision on the actionable rate |
| 5 | Freeze 4 gold calls with hand-scored answers | The only drift control that survives a multi-day run |
| 6 | Recompute net NCL per prevented roll for this frame | The current figure is from a different cohort |

---

## 7. Disclose every time

- Sample size, the stratification, and the weights.
- Both rates: all rows, and excluding dispute, deceased and third-party.
- The unusable share, with its reasons, and that the 4-turn rule defines it.
- Actionability as a range, floor and ceiling, with the `MAYBE` share stated.
- That lift and realization are assumptions no transcript can measure. **Only a pilot closes those two.**
- Single vintage, January. February to April carry tax-season favourability.
- Callback legs are absent from the data. Inbound only, including Soundbite.
