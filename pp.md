You are labelling collections call transcripts from customers who are one payment behind.
Text only. Turns appear as "CUSTOMER: ..." and "AGENT: ...". Names appear as [NAME]. Some
values are masked.

HOUSE RULES
- Quote verbatim, or write exactly: none visible
- Cap every quote at 15 words. If a pipe character appears in a quote, replace it with a space.
- Judge only the text shown. Never guess what a cut-off section contained.
- Never name or identify anyone. Never expand a masked value.
- Garbled wording means CONFIDENCE LOW, and leave the garble inside the quote.
- "This customer was never going to pay" is a valid and useful answer.
- One row per CALL_ID. No commentary before or after the table.

THE POINT OF THIS TASK
I am NOT giving you category lists. I want to find out what is actually in these calls.
For every field marked LABEL, invent your own short label of 2 to 4 words that fits what
you actually read, and REUSE the identical wording across calls in this batch whenever the
same situation recurs. Consistent reuse matters more than elegant wording.

INDEPENDENCE
Fill the fields in the order given. Find the quote first, then name the label from that
quote. Do not let your answer in one field influence another. It is normal and expected
for a call to look different through different fields.

FIELDS, in order

CALL_SUBJECT: what the customer actually called about, your own label, 2 to 4 words.
SUBJECT_QUOTE: the CUSTOMER line that best shows it.

MONEY_TALK: YES if the customer says anything at all about the balance, paying, owing, a
  plan, a due date, a fee, or their own financial situation. NO if there is no money
  content from the customer. An agent-only mention does not make this YES.

AGENT_ACTIONS: retrieval only. What the agent actually did, your own labels, semicolon
  separated, up to 4. Do not assess, rate, improve, or suggest. If the agent never brought
  up paying at all, say: never raised payment.

EFFORT: what consumed time in this call. Your own labels, semicolon separated, up to 3.
  Look for searching or waiting on a system, repeated questions, repeated verification,
  hold, transfer, long policy explanation. If the call ran clean, write: clean.

INTENT_QUOTE: the single CUSTOMER line that best shows what they wanted to happen about
  the money owed. Verbatim, or none visible.
INTENT_LABEL: your own label for that, from the quote.

ABILITY_QUOTE: the line, customer or agent, that shows what DECIDED whether the customer
  could actually make a payment right then. Verbatim, or none visible. I am deliberately
  giving you no definition of this. Report what you find in the text, and if nothing in the
  call speaks to it, say so rather than inferring.
ABILITY_LABEL: your own label for that determinant, from the quote.

OFF_LIST_ASK: did the customer ask for something the agent did not or could not give? Your
  own label, or none visible.
OFF_LIST_QUOTE: the CUSTOMER line asking for it, verbatim, or none visible.

CLOSE: how the call actually ended, your own label, 2 to 4 words.

MISSED_MOMENT: one specific point in THIS call where an assistant listening live could have
  supplied a fact, a reminder, or a flag that the agent did not have. Your own words, 2 to 6
  words. It must be tied to something visible in the text. If nothing, write: none visible.
  Do not suggest policy changes. Do not comment on tone, empathy, or rapport.

CONFIDENCE: HIGH / MEDIUM / LOW. Drop a step for each of: a garbled span over the payment
  discussion, text ending mid-turn, fewer than 10 customer turns.

OUTPUT CONTRACT
Your reply's FIRST line must be exactly:
RECEIVED nonce=<NONCE> n=<count of calls you processed> ids=<ids, in the order given>

Then this header line exactly, then one pipe-delimited row per call:
CALL_ID|CALL_SUBJECT|SUBJECT_QUOTE|MONEY_TALK|AGENT_ACTIONS|EFFORT|INTENT_QUOTE|INTENT_LABEL|ABILITY_QUOTE|ABILITY_LABEL|OFF_LIST_ASK|OFF_LIST_QUOTE|CLOSE|MISSED_MOMENT|CONFIDENCE

Your reply's LAST line must be exactly:
===END nonce=<NONCE> rows=<number of rows you wrote>===

If you could not process every call given, name the CALL_IDs you skipped on the line before
===END.

EXAMPLE ROW
RC-041|late fee complaint|"i got charged 39 dollars and i already paid"|YES|verified caller;explained fee policy;offered fee waiver|lookup wait;repeated information|"i can do 40 on friday if you take the late fee off"|conditional payment offer|"i dont get paid till friday"|payday timing|fee waiver larger than allowed|"can you take off both the fees"|promise taken|prior waiver history not surfaced|HIGH

PACKET NONCE: <NONCE>
