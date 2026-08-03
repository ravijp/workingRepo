Read each collections call transcript below. Every customer was one payment behind
and later rolled to the next delinquency bucket. Turns appear as "CUSTOMER:" and
"AGENT:". Names appear as [NAME]. Some digits are masked. The audio was
machine-transcribed and is imperfect.

Some of these calls were placed by the bank's automated dialer and still record as
inbound. If the agent says why the bank is calling, the bank called.

Quote verbatim, or write exactly: none visible
For intent, quote CUSTOMER lines only. An agent's scripted "payment plan" is never
customer intent. Cap every quote at 15 words. Replace any tab or newline inside a
quote with a single space.

Read all calls in the packet before writing anything.

===============================================================
STEP 1. IS THERE ANYTHING TO READ
===============================================================

USABLE: YES / NO
  NO if the call disconnected, or there are fewer than 4 customer turns, or the
  customer never responds, or the money talk is garbled beyond following.
  A short call that is complete and followable is USABLE.

  IF USABLE IS NO: write NOT SCORED in every remaining column and stop.

WHY_NOT_USABLE (blank if USABLE):
  DISCONNECTED MID-CALL / NO CUSTOMER RESPONSE / UNDER 4 CUSTOMER TURNS /
  GARBLED / OTHER: <3-5 words>

===============================================================
STEP 2. DID THEY SHOW INTENT
===============================================================

INTENT_QUOTE: the CUSTOMER line that best shows what they wanted about the money.

INTENT_WHY: why you chose the value below. Max 15 words. Write this before the value.

INTENT, pick one:
  COMMITTED            will pay AND named an amount or a day
                       e.g. "I'll pay the 50 on Friday"
  WILLING              said they can or want to pay, named neither
                       e.g. "I want to take care of this"
  ASKED HOW TO PAY     seeking a way to clear it: a programme, a plan, options,
                       what would resolve it. Never said they would pay
                       e.g. "are there any hardship programmes"
  ASKED ABOUT ACCOUNT  balance, a statement, a charge. Not trying to resolve it
                       e.g. "what's my balance"
  NOT ABOUT MONEY      money never came up

  Boundary rules:
  - Said they can or want to pay? WILLING, even with no amount and no date.
  - Only asked for a route to clear the debt? ASKED HOW TO PAY.
  - Only wanted information, not resolution? ASKED ABOUT ACCOUNT.

EXCLUDE_FLAG: DISPUTE OR FRAUD / DECEASED OR ESTATE / THIRD PARTY / NONE
  Set whenever it applies, whatever the intent value. It changes no other field.

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

ASSIST_TYPE, pick one. Each has a precondition. If the precondition is not visible
in the call, you may not choose that value.
  PROMPT TO ASK             only if the agent never asked for money anywhere
  PROMPT TO CLOSE           only if intent was shown, the agent acknowledged it,
                            and never asked for a card or a date
  SURFACE A PAYMENT METHOD  only if the text shows a method exists or existed
  SURFACE ELIGIBILITY       only if a fee or a programme was discussed or relevant
  SURFACE THE NEXT STEP     only if the customer needed a route they did not get
  FLAG A WEAK PROMISE       only if a promise was actually made in this call
  NOTHING                   no assist applies
  OTHER: <3-5 words>

===============================================================
OUTPUT
===============================================================

Put your ENTIRE reply inside ONE fenced code block. Nothing before it, nothing
after it. No prose, no headings, no commentary. Tab-separated. Never use a tab
inside a cell.

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
- Entire reply inside one fenced code block, tab-separated, 13 columns
- USABLE=NO means NOT SCORED in every remaining column
- INTENT_WHY comes before INTENT. SCREEN_TEXT comes before AI_COULD_HELP
- SCREEN_TEXT is an actual message to an agent, under 12 words, or: nothing
- Use only the text between the BEGIN and END markers
- Last line exactly: ===END nonce=K7QX rows=6===
