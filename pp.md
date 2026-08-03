You are reading collections call transcripts. Reply with exactly: READY

WHAT THESE CALLS ARE

Recorded phone calls between a US credit card issuer's collections agents and its
cardholders, most of whom have missed one payment. Four things will mislead you if
you assume otherwise.

1. Not every call was started by the customer. An automated outbound dialer places
   many of these, verifies the customer, and hands the connected call to an agent.
   The system records it as inbound anyway. The tell is the agent's opening words.
   If the agent says why the bank is calling, the bank called. "Why did you call
   me" is then a real and complete customer goal, not a missing answer.
2. Agents work from scripts. Much of the text is boilerplate: recording notice,
   operator ID, product welcome, greeting. Scripted words about payments, plans or
   hardship are the script speaking, not the customer.
3. The audio was machine-transcribed badly. Amounts and proper nouns are often
   wrong. Names and card numbers appear as tags in square brackets, digit runs as
   hashes.
4. Turn labels are not always right. On a transfer the receiving agent can appear
   under CUSTOMER. Interpreters and IVR prompts appear under CUSTOMER. Use what the
   words say the speaker is doing, not only the label.

WHERE TO LOOK

The informative parts are rarely the wordiest. Read for these moments and quote them.
- The moment the reason for the call is first stated, and by whom.
- Any moment a specific object is named: an amount, a date, a card, a bank account,
  the app, a website, a letter, a statement, a document, an ID, another person,
  another company, another department. A call where a named object is missing,
  expired, wrong, locked, lost, unreceived, or held by someone else is a different
  call from one where it is present.
- The moment either side says something is not possible, not allowed, not available
  yet, or needs something the caller does not have. Quote those words.
- The moment the customer's own plan changes between the start and end of the call.
- Any moment work is repeated: a fact restated because the agent asked again,
  identity checked twice, a hold followed by a re-explanation.

WHY THIS PASS EXISTS

This is discovery, not scoring. I have no category list and will not give you one.
I am finding out what is actually in these calls so a fixed list can be built later.
Your labels ARE the deliverable and they will be counted. A label that is true but
tells me nothing is worse than no label, because it will look like a finding.

WHAT COUNTS AS A LABEL

2 to 4 words, lower case, no punctuation. A good label names an OBJECT, a STATE, or
a COUNT. A bad label names a TOPIC. Apply all four tests to your own label before
writing it, and rewrite until it passes.

1 SWAP TEST. Could this label sit unchanged on a call about a lost card, a disputed
  charge, or an address change? If yes it is too broad.
2 GUESS TEST. From the label alone, could I guess roughly what the customer said?
  If not, the label threw away the content.
3 OBJECT TEST. The label must contain a concrete thing from the customer's world,
  or a state of one: an amount, a date, a fee, a card, a check, a deposit, a payday,
  a spouse, the app, a bank, autopay, a letter. Abstract nouns only means it fails.
4 ACTION TEST. Two calls may share a label only if an assistant listening live to
  both should do the same thing about both. If the right move differs, split it.

BANNED, on sight: payment discussion, customer inquiry, account issue, billing
question, payment issue, financial difficulty, general inquiry, wants to pay, needs
help, unable to pay, hardship, verification, authentication issue, agent assisted,
account review, customer deferred, transfer, dispute, fee issue.
Banned as words anywhere in a label: discussion, inquiry, issue, concern, matter,
situation, general, various, related, regarding, appropriate, relevant.

BAD, THEN GOOD:
  payment discussion    ->  funds ready no rail on file
  unable to pay         ->  named payday three weeks out
  unable to pay         ->  bank balance negative today
  customer deferred     ->  deferred to app no amount
  authentication issue  ->  verbal password failed six attempts
  authentication issue  ->  third party no poa on file
  account access        ->  web login locked after attempts
  hardship              ->  asked for plan at fourteen days
  dispute               ->  merchant dispute withholding minimum
  customer inquiry      ->  returning missed ivr call
  transfer              ->  reverified from scratch after transfer
  deceased account      ->  estate waiting on mailed final bill
  fee issue             ->  fee refund demanded before paying
  agent did not close   ->  notated intent never asked card
  statement issue       ->  never received bill does not know amount
  payment failed        ->  scheduled ach stopped by customer bank

THE HONEST ESCAPE HATCH
Sometimes a call really is thin and any specific label would be invention. Do not
invent. Write it in this exact form: thin: <your best 2 to 4 words>
That lets me count thin calls rather than find them buried later. But note which
error is worse: using this on a call that had real content loses a finding
permanently. Use it for genuinely thin text, not merely difficult text.

REUSE, AND WHEN NOT TO
Reuse a label only when the ACTION TEST says two calls are the same. NEVER widen a
label so it covers more calls. If two calls are close but not the same, give them
two labels. I merge afterwards and merging is cheap for me. A label that was widened
to fit cannot be recovered without re-reading everything.

=====================================================================
REPLY IN EXACTLY THREE PARTS AND NOTHING ELSE
=====================================================================

PART A - READ EACH CALL. No labels here. Plain sentences only.
Read every call in the packet before writing anything.

### id=<id>
WHY THEY CALLED: one sentence from what is actually said, and say who raised it.
MONEY: one sentence on what the customer wanted to happen about the money owed. If
  the customer never touches the balance, paying, owing, a plan, a due date, a fee,
  or their own finances, write exactly: no money content
WHAT DECIDED IT: one sentence on what determined whether this customer could
  actually pay right then. I am deliberately giving you no definition of this and no
  list. Describe the mechanism you see, whatever kind of thing it turns out to be.
  If nothing in the text decided it, say so.
WHERE THE TIME WENT: one sentence on what consumed this call.
WHAT THE AGENT DID: short list, verb then object, from the text. No quality or tone.
THE ODD DETAIL: the one thing here you would not have predicted from the other calls
  in this packet. If genuinely nothing, write: nothing odd

PART B - NOW NAME THE LABELS.
Work from your Part A sentences. Do not go back and change Part A.
One block per call, in the order given, this shape exactly:

--- id=<id>
GOAL_LABEL: <<2-4 word label you invent. If MONEY is "no money content", write: none>>
GOAL_QUOTE: "<<verbatim CUSTOMER line only. An agent offering a plan is not customer
  intent. Or: none visible>>"
DETERMINANT_LABEL: <<your label for what decided whether they could pay>>
DETERMINANT_KIND: <<ONE word naming the family it belongs to. Invent families as you
  go and reuse them exactly>>
DETERMINANT_QUOTE: "<<the line showing it. May be an agent line. Or: none visible>>"
TIME_SINK_LABEL: <<your label for what consumed the call>>
AGENT_DID: <<verb then object, semicolon separated, up to 4>>
CONFIDENCE: <<HIGH, MEDIUM or LOW. Start HIGH, drop one step for each of: a garbled
  span over the money talk, text ending mid-turn, fewer than ten customer turns>>

PART C - LOOK ACROSS THE CALLS. This is the point of the exercise. Do not shorten it.

C1 ROSTER. Pipe-delimited, one row per distinct label used anywhere in Part B:
   DIMENSION|LABEL|COUNT|IDS|BEST_QUOTE
   DIMENSION is one of GOAL, DETERMINANT, KIND, TIME_SINK, AGENT_DID.
C2 WORST FIT. For each label used on more than one call, name the call that fits it
   least well and say in one line what is different. If no reuse, say so.
C3 NO LABEL FOR. The call, or part of a call, that did not sit cleanly under
   anything you wrote. One sentence. If everything fitted, say so and say why.
C4 UNASKED. The single thing you saw across these calls that none of my fields asked
   about. One sentence.
C5 THE SPLIT I EXPECT. Which of your labels will need splitting once more calls are
   read, and what would confirm it.
