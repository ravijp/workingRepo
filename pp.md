Read each call. Every customer was one payment behind and later rolled to the next
bucket. Turns are "CUSTOMER:" and "AGENT:". Names are [NAME]. Digits may be masked.
Machine-transcribed, imperfect.

Some of these calls were placed by the bank's automated dialer and still record as
inbound. If the agent says why the bank is calling, the bank called.

Quote CUSTOMER lines only for intent. An agent's scripted "payment plan" is never
customer intent. Quote verbatim or write: none visible

--- LENS 0: is there anything to read

USABLE: YES / NO
  NO if the call disconnected, or under 10 customer turns, or the customer never
  responds, or the money talk is garbled.
WHY_NOT_USABLE (only if NO): DISCONNECTED / NO CUSTOMER RESPONSE / TOO SHORT /
  GARBLED / OTHER: <3-5 words>

--- LENS 1: did they show intent

INTENT, pick one:
  COMMITTED        will pay AND named an amount or a day
  WILLING          said they can or want to pay, named neither
  ASKED ABOUT IT   balance, options, a programme or hardship discussed; never said
                   they would pay
  NOT ABOUT MONEY  money never came up
INTENT_QUOTE

--- LENS 2: could anything have been done about it on this call

At one cycle past due an agent may ONLY do these six things:
  M1 take a payment, if a method is already on file
  M2 take a promise to pay
  M3 waive a late or interest fee, if asked and eligible
  M4 verify the caller and discuss the account
  M5 transfer, or arrange a callback
  M6 tell them the amount due and the due date
NOT available at this stage: any payment programme, forbearance, settlement,
long-term arrangement, hardship programme, re-age, or acting for a third party
without documentation.

SOLVABLE, pick one:
  YES        one of M1-M6 was available and was not used. Name it.
  MAYBE      something might have helped but it is not certain from the text
             (for example the agent could have walked them to the web page)
  NO         they wanted something not available at this stage, or nothing
             would have changed the outcome
  CANNOT TELL
ACTION (if YES or MAYBE): M1 / M2 / M3 / M4 / M5 / M6 / GUIDED TO SELF-SERVE

WHAT_STOPPED_IT, pick one:
  NO WAY TO PAY ON FILE / NO MONEY RIGHT NOW / WANTED AN OFFER WE DO NOT HAVE /
  LOCKED OUT ONLINE / NOT THE CARDHOLDER / DISPUTING THE CHARGE / NOBODY ASKED /
  NOTHING STOPPED IT / OTHER: <3-5 words>
STOPPED_QUOTE

Put your ENTIRE reply in one fenced code block. Nothing outside it. Tab-separated.
No tabs inside cells. Replace any newline in a quote with a space.

Line 1 exactly: RECEIVED nonce=K7QX n=6 ids=<ids in order given>
Line 2 exactly:
ID	USABLE	WHY_NOT_USABLE	INTENT	INTENT_QUOTE	SOLVABLE	ACTION	WHAT_STOPPED_IT	STOPPED_QUOTE
Then one row per call.
Last line exactly: ===END nonce=K7QX rows=6===
