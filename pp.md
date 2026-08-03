Read each call. Every customer is one payment behind. Turns are "CUSTOMER:" and
"AGENT:". Names are [NAME]. Some digits are masked. Machine-transcribed, imperfect.

Some of these calls were started by the bank, not the customer. If the agent says
why the bank is calling, the bank called. "Why did you call me" is a real answer.

Quote CUSTOMER lines only for INTENT. An agent's scripted "payment plan" is never
customer intent. Quote verbatim or write: none visible

INTENT, pick one:
  COMMITTED        said they will pay AND named an amount or a day
  WILLING          said they can or want to pay, named neither
  MONEY CAME UP    balance, due date or options discussed, never said they'd pay
  NOT ABOUT MONEY  money never came up

WHAT STOPPED IT, pick one:
  NO WAY TO PAY ON FILE / NO MONEY RIGHT NOW / NOT THE CARDHOLDER /
  LOCKED OUT ONLINE / DISPUTING THE CHARGE / WANTED AN OFFER WE DO NOT HAVE /
  NOBODY ASKED / NOTHING STOPPED IT / CANNOT TELL /
  OTHER: <describe in 3-5 words>

TRIED TO PAY: YES if a card or bank detail was given or an arrangement booked,
  including failed attempts. "I'll check online" is NO.
AGENT ASKED: YES if the agent asked for money or asked when they will pay.
READABLE: NO if under 10 customer turns, text stops mid-turn, or the money talk
  is garbled.

Put your ENTIRE reply in one fenced code block. Nothing outside it.
Tab-separated. No tabs inside cells. Replace any newline in a quote with a space.

Line 1 exactly: RECEIVED nonce=K7QX n=6 ids=<ids in order given>
Line 2 exactly:
ID	INTENT	INTENT_QUOTE	TRIED_TO_PAY	AGENT_ASKED	WHAT_STOPPED_IT	STOPPED_QUOTE	READABLE
Then one row per call.
Last line exactly: ===END nonce=K7QX rows=6===
