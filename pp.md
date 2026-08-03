Read each call. Every customer was one payment behind and later rolled to the next
bucket. Turns are "CUSTOMER:" and "AGENT:". Names are [NAME]. Digits may be masked.
Machine-transcribed, imperfect.

Some of these calls were placed by the bank's automated dialer and still record as
inbound. If the agent says why the bank is calling, the bank called.

Quote CUSTOMER lines only for intent. An agent's scripted "payment plan" is never
customer intent. Quote verbatim or write: none visible

USABLE: YES / NO
  NO only if the call disconnected, or there are fewer than 4 customer turns, or
  the customer never responds, or the money talk is garbled beyond following.
  A short call that is complete and followable is USABLE.
WHY_NOT_USABLE (blank if USABLE): DISCONNECTED MID-CALL / NO CUSTOMER RESPONSE /
  UNDER 4 CUSTOMER TURNS / GARBLED / OTHER: <3-5 words>

INTENT_QUOTE: the CUSTOMER line that best shows what they wanted about the money.
INTENT_WHY: why you chose the value below. Max 15 words. Write this before the value.
INTENT, pick one:
  COMMITTED        will pay AND named an amount or a day
  WILLING          said they can or want to pay, named neither
  ASKED ABOUT IT   balance, options, a programme or hardship discussed, never said
                   they would pay
  NOT ABOUT MONEY  money never came up

PAID_ON_CALL: YES / NO / NOT VISIBLE
  YES only if a card or bank detail was given, or a payment was taken or booked on
  this call. "I'll check online" or "I'll pay Friday" is NO.

WHAT_STOPPED_IT, the customer-side reason, pick one:
  NO WAY TO PAY ON FILE / NO MONEY RIGHT NOW / WANTED AN OFFER WE DO NOT HAVE /
  LOCKED OUT ONLINE / NOT THE CARDHOLDER / DISPUTING THE CHARGE / NOBODY ASKED /
  NOTHING STOPPED IT / OTHER: <3-5 words>
STOPPED_QUOTE: the line showing it, or: none visible

Now the assist question. An AI assistant listens to this call live and can put ONE
thing on the human agent's screen. The goal is to help the agent take a payment
from this customer on this call.

ASSIST_WHY: what the agent lacked, and whether putting it on screen would plausibly
  have got a payment. Max 15 words. Write this before the value.
AI_COULD_HELP, pick one:
  YES    a specific, nameable thing on screen would likely have got the payment
  MAYBE  it might have helped, but the text does not make it certain
  NO     nothing on a screen would have changed this call
ASSIST_TYPE, pick one:
  PROMPT TO ASK             the agent never asked for money
  PROMPT TO CLOSE           intent was shown and acknowledged, never asked for a
                            card or a date
  SURFACE A PAYMENT METHOD  a saved method existed and was not offered
  SURFACE ELIGIBILITY       a fee waiver, or the date a programme opens
  SURFACE THE NEXT STEP     the right web path, department, or document
  FLAG A WEAK PROMISE       conditional wording, date after the due date, or an
                            amount under the minimum
  NOTHING                   no assist applies
  OTHER: <3-5 words>

Put your ENTIRE reply in one fenced code block. Nothing outside it. Tab-separated.
No tabs inside cells. Replace any newline in a quote with a space.

Line 1 exactly: RECEIVED nonce=K7QX n=6 contact_ids=<contact ids, in the order given>
Line 2 exactly:
CONTACT_ID	USABLE	WHY_NOT_USABLE	INTENT_QUOTE	INTENT_WHY	INTENT	PAID_ON_CALL	WHAT_STOPPED_IT	STOPPED_QUOTE	ASSIST_WHY	AI_COULD_HELP	ASSIST_TYPE
Then one row per call.
Last line exactly: ===END nonce=K7QX rows=6===
