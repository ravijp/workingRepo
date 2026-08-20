# The telemetry logging table

This document follows ASD-STE100 Simplified Technical English.

This document stands alone. You can walk a room through it without any other file.
The other files in this folder give the detail behind it.

---

# Part 1: The primer

Read this page before the table. It teaches the words the table uses.

## 1.1 A trace, a span and a span event

A **trace** is the full record of one unit of work.
For this application, one trace is one email, from the moment the message arrives to the moment the case is written.

A **span** is one timed piece of that work.
A span has a name, a start time and an end time.
A span therefore answers the question "how long did this take?".

A span can hold another span. This is called nesting.
The inner span is the **child**. The outer span is the **parent**.
Nesting makes a tree, and the tree shows you which piece of work is inside which.

A **span event** is one thing that happened at one moment inside a span.
A span event has no duration. It is a point in time, not a period.
A span event therefore answers the question "what happened?", not "how long?".

Use this rule to choose between them:

- If you want a duration, open a span.
- If you want a fact, emit a span event.

## 1.2 The correlation identifier

A trace only works if every record in it can be found.
The thing that ties the records together is the **correlation identifier**.

In a normal build, the tracing library does this for you.
It generates an identifier, puts it on every span, and copies it into an outbound HTTP header, so the next service joins the same trace.
This is called trace context propagation.

**This build does none of that.**
The application has no tracing library. It uses the Python standard library alone.
The team removed the OpenTelemetry distribution on 19 August, because it pulled 44 packages through a bank feed that vets every one.

So this build threads its own identifier by hand.
The value is `wab.correlation_id`.
It arrives on the Service Bus message, and each call site must pass it into each telemetry call.
Where a call site does not pass it, the field is simply absent.

You must know three consequences:

1. The identifier is **optional on every record**. It is not guaranteed.
2. It is **not sent to Foundry or to MuleSoft as a trace header**, so those services do not join the trace.
3. It is **not stored in the state store**, so a telemetry row cannot be joined to a durable row.

On 17 July the team decided that the CRM activity identifier is the tie-together key across systems.
That identifier is not yet on any telemetry record. The table marks that as PROPOSED.

## 1.3 A log record, an event and a metric

These three words are often used loosely. On this platform the difference decides what you can build.

A **log record** is a line of text with a level and a time.
An **event** is a structured fact with named fields.
A **metric** is a number that the platform adds up for you before you query it.

On this platform there is only one of the three.

Everything is a log record.
The Azure Functions host forwards standard-library log records to Application Insights, and each one becomes a row in the `traces` table.
There is no `customEvents` table and no `customMetrics` table for this application.

The application works around this in one way, and it is the single most important mechanical fact in this document:

> The Python worker does not forward a log record's `extra=` fields.
> So the application packs its structured fields into the log **message**, as one line of JSON.
> Every query must therefore call `parse_json(message)`.

A span is a log record whose JSON has `wab.telemetry` set to `span`.
A span event is a log record whose JSON has `wab.telemetry` set to `event`.
That one field is the whole difference.

Three results follow, and each one shapes the table:

- **There is no pre-aggregation.** Every number is a Kusto aggregate over `traces`.
- **There are no metric alerts.** Every alert is a scheduled query.
- **Sampling is on.** The host sends about 20 items per second and drops the rest. A simple count is too low under load. Use `sum(itemCount)`, which restores the estimate.

## 1.4 One email, followed all the way through

Take one email. This is what it produces today.

1. The message arrives at `func_orchestrate`. **Nothing is emitted.** There is no run-start record.
2. The application claims the email in the state store, reads it from MuleSoft, and turns the HTML into text. **Nothing is emitted** for any of these three steps.
3. The application redacts the personal data at step B.6. On success, nothing is emitted yet. On failure, the span `orchestration.step_failed` opens and closes, and inside it the span event `wab.case_intake.step_failed` fires with the step named as `redact`. The email then stops.
4. The application calls Foundry at step B.7. The span `classify_email` opens. Inside it the child span `model.responses_parse` opens, times the model call, records the input and output tokens, and closes. Then the span event `wab.case_intake.classified` fires, and `classify_email` closes.
5. The application now emits the redaction evidence, **after** the classification, not at step B.6. The span `guardrails.evidence` opens, the span event `wab.case_intake.guardrail_evidence` fires inside it, and the span closes.
6. The application builds the payload and writes the case. **Nothing is emitted** for either step.
7. If a case was created and a receipt time was known, the span event `wab.case_intake.latency` fires.
8. The function writes one plain text line with the outcome.

Note step 5 carefully.
The evidence record is emitted late, so the order of the records does not match the order of the steps.
The code does this at `src/orchestration/spine.py` line 686.
A reader who assumes that record order equals step order will read the log wrongly.

That is the whole trace. Five spans at most, five span events at most, and one text line.
Part 2 shows what is there, and what is not.

---

# Part 2: The master table

## 2.1 How to read the table

The step identifiers follow the design, not Pradeep's workbook.
Where the two differ, the "Log event" column names Pradeep's row, so you can line the two up.
Pradeep's step numbers are shifted by one from B.6 onward, because his workbook has no redaction step.
Section 4.4 of `08-pradeep-design-review.md` covers that.

The **Status** column is strict:

| Status | Meaning |
| --- | --- |
| **EMITTED TODAY** | The code emits this now. You can query it now. |
| **PARTIAL** | Something is emitted, but not in the form the row needs. The cell says what is short. |
| **PROPOSED** | The code does not emit this. It is a design proposal. Any field name marked PROPOSED does not exist yet. |

Every attribute name that is not in the current vocabulary is written as **PROPOSED**.
Do not implement a PROPOSED name as if it were settled.

The **Correlation key** column names the field that ties the record to the rest of the email's journey.
Where it says `wab.correlation_id (optional)`, the field is present only when the call site passed it.

## 2.2 The table

| Step | Component | Log event (Pradeep row) | Signal type | Span name | Parent span | Level | Key fields | Correlation key | Destination | Status |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| B | func_orchestrate | Function invoked (P row 2) | span event | `wab.case_intake.run_started` PROPOSED | none | INFO | `wab.email_id` PROPOSED, `wab.queueitem_id` PROPOSED, `wab.attempt` PROPOSED, `wab.correlation_id` | `wab.email_id` PROPOSED | App Insights `traces` | PROPOSED. Gap G1. Pradeep's row 2 also assumes OpenTelemetry extracts the trace context. It does not. See review 4.1. |
| B | func_orchestrate | Function completed with outcome (P row 3) | span event | `wab.case_intake.run_completed` PROPOSED | none | INFO | `wab.outcome` PROPOSED, `wab.email_id` PROPOSED, `wab.case_id` PROPOSED, `wab.duration_ms`, `wab.attempt` PROPOSED | `wab.email_id` PROPOSED | App Insights `traces` | PROPOSED. Gap G1. The outcome exists today only in a plain text line. |
| B | func_orchestrate | Per-email outcome line | log record | none | none | INFO | `status`, `source_id`, `case_id` as printf text | `source_id` in the text | App Insights `traces` | EMITTED TODAY at `src/functions/func_orchestrate/__init__.py`:104. A query must parse text, not JSON. |
| B | func_orchestrate | Unhandled exception (P row 4) | log record | none | none | ERROR | `queueitem_id`, `email_id`, `delivery_count`, full traceback | `email_id` in the text | App Insights `traces` | EMITTED TODAY at `:94`. The traceback can quote email content. See review 4.3. |
| B | func_orchestrate | Unparseable message, dropped as poison | log record | none | none | EXCEPTION | `message_id`, `delivery_count` | `message_id` in the text | App Insights `traces` | EMITTED TODAY at `:76`. Pradeep has no row for this. The email is lost and nobody is told. |
| B.1 | State store claim | Write attempt succeeded (P row 5) | span event | `wab.case_intake.step_completed` PROPOSED | none | INFO | `wab.step` = `claim`, `wab.email_id` PROPOSED, `wab.attempt` PROPOSED, `wab.duration_ms` | `wab.email_id` PROPOSED | App Insights + Dataverse | PROPOSED. Gap G4. Pradeep names "SQL (Email Instance Table)". The runtime uses Dataverse. See review 4.9. |
| B.1 | State store claim | Write failed (P row 6) | span event | `wab.case_intake.step_failed` | `orchestration.step_failed` | ERROR | `wab.step` = `claim`, `wab.error_type` | `wab.correlation_id` (optional) | App Insights `traces` | PROPOSED. The event exists, but no call site passes `claim`. Gap G4. |
| B.1 | State store claim | Claim degraded to the non-durable path | log record | none | none | WARNING | free text with the email identifier | text only | App Insights `traces` | EMITTED TODAY at `src/orchestration/state_store.py`:429, 435, 442. Pradeep has no row for this. The exactly-once guarantee is lost when it fires. |
| B.2 | Queue item write | Queue-item write result (P row 7) | span event | `wab.case_intake.step_completed` PROPOSED | none | INFO | `wab.step` = `queue_item`, `wab.queueitem_id` PROPOSED | `wab.email_id` PROPOSED | App Insights + Dataverse | PROPOSED. Gap G4. |
| B.2 | Queue item write | Queue-item write failed (P row 8) | span event | `wab.case_intake.step_failed` | `orchestration.step_failed` | ERROR | `wab.step` = `queue_item`, `wab.error_type` | `wab.correlation_id` (optional) | App Insights `traces` | PROPOSED. Gap G4. Today this failure is swallowed silently at `src/orchestration/spine.py`:333. |
| B.3 | Claim decision | Decision and status update (P rows 9, 10, 11) | none | none | none | none | none | none | none | **Not recommended.** B.3 is a branch on the B.1 outcome that B.1 already records. Use one `wab.outcome` value with an exit reason instead. See review 5.5. |
| B.4 | MuleSoft email read | Pull succeeded (P row 12) | span | `crm.read_email` PROPOSED | none | INFO | `wab.duration_ms`, `wab.http_status` PROPOSED, `wab.subject_present` PROPOSED, `wab.body_present` PROPOSED, `wab.sender_present` PROPOSED | `wab.email_id` PROPOSED | App Insights `traces` | PROPOSED. Gap G7. There is no dependency tracking, so an untimed call is invisible. Pradeep's presence booleans are correct and are kept. See review 3.2. |
| B.4 | MuleSoft email read | Pull failed (P row 13) | span event | `wab.case_intake.step_failed` | `crm.read_email` PROPOSED | ERROR | `wab.step` = `data_pull`, `wab.error_type`, `wab.http_status` PROPOSED | `wab.email_id` PROPOSED | App Insights `traces` | PROPOSED. Gap G4 and G7. Pradeep lists `error_message`. Do not log it. See review 4.3. |
| B.5 | Normalize HTML to text | Parse details and empty output (P rows 14, 15, 16) | span event | `wab.case_intake.step_completed` PROPOSED | none | INFO or WARNING | `wab.step` = `normalize`, `wab.input_length` PROPOSED, `wab.output_empty` PROPOSED | `wab.email_id` PROPOSED | App Insights `traces` | PROPOSED. Gap G4. Pradeep's three rows share one justification text and only one of the three matches it. See review 4.8. |
| **B.6** | **Azure AI Language redaction** | **Redaction evidence** | **span event** | `wab.case_intake.guardrail_evidence` | `guardrails.evidence` | INFO | `wab.redaction_count`, `wab.redaction_submitted_documents`, `wab.redaction_billable_text_records`, `wab.redaction_elapsed_ms`, `wab.redaction_model_version`, `wab.redaction_counts_by_category` | `wab.correlation_id` (optional) | App Insights `traces` | **EMITTED TODAY** at `src/orchestration/spine.py`:369. **Pradeep has no row for this step at all.** See review 4.4. Note it is emitted after B.7, not at B.6. |
| **B.6** | **Azure AI Language redaction** | **Redaction evidence span** | **span** | `guardrails.evidence` | none | INFO | `wab.duration_ms` only | none | App Insights `traces` | EMITTED TODAY at `:368`. The span wraps no work, so its duration is not the redaction time. Read `wab.redaction_elapsed_ms` for that. |
| **B.6** | **Azure AI Language redaction** | **Redaction failed, email stopped** | **span event** | `wab.case_intake.step_failed` | `orchestration.step_failed` | ERROR | `wab.step` = `redact`, `wab.error_type` | `wab.correlation_id` (optional) | App Insights `traces` | **EMITTED TODAY** at `:659`. This is the most important operational record in the system. Redaction fails closed, so a spike means nothing is processing. |
| B.6 | Azure AI Language redaction | Step-failure span | span | `orchestration.step_failed` | none | ERROR | `wab.duration_ms` only | none | App Insights `traces` | EMITTED TODAY at `:386`. |
| B.6 | Azure AI Language redaction | Redaction completed | log record | none | none | INFO | `redactions`, `documents`, `elapsed_ms`, `model` as printf text | none | App Insights `traces` | EMITTED TODAY at `src/guardrails/masking.py`:396. It duplicates the evidence event. Keep one, not both. |
| B.6 | Azure AI Language redaction | Out-of-policy category ignored | log record | none | none | WARNING | `category` | none | App Insights `traces` | EMITTED TODAY at `src/guardrails/masking.py`:508. |
| **B.7** | **Foundry classification** | **Agent invocation (P row 17)** | **span** | `classify_email` | none | INFO | `wab.model_deployment`, `wab.endpoint_resource`, `wab.agent_target`, `wab.agent_name`, `wab.foundry_via_gateway`, `wab.halted`, `wab.decision_gate`, `wab.final_confidence`, `wab.duration_ms`, `wab.status`, `wab.error_type` | `wab.correlation_id` (optional) | App Insights `traces` | **EMITTED TODAY** at `foundry/agent/classify.py`:131. Pradeep's row 17 claims the identifier is propagated to Foundry as trace context. It is not. See review 4.2. |
| **B.7** | **Foundry model call** | **Invocation succeeded (P row 18)** | **span** | `model.responses_parse` | `classify_email` | INFO | `gen_ai.system`, `gen_ai.request.model`, `gen_ai.usage.input_tokens`, `gen_ai.usage.output_tokens`, `wab.agent_version`, `wab.confidence_as_classified`, `wab.is_actual_work`, `wab.category`, `wab.duration_ms` | `wab.correlation_id` on the parent | App Insights `traces` | **EMITTED TODAY** at `:174`. The two token fields are the whole answer to the 15 July FinOps problem. Pradeep's single `tokens_used` must be two fields. See review 5.6. |
| B.7 | Foundry model call | Response identifier for the Foundry join | span attribute | `model.responses_parse` | `classify_email` | INFO | `gen_ai.response.id` PROPOSED | `gen_ai.response.id` PROPOSED | App Insights `traces` | PROPOSED. Gap G16. This is the achievable substitute for the trace propagation Pradeep's row 26 tests for. See review 5.7. |
| B.7 | Foundry model call | Agent response text (P row 18) | none | none | none | none | none | none | none | **Refused.** Pradeep lists `agent_response`. It is content. The design forbids content in telemetry and the compliance question is open. See review 4.3. |
| B.7 | Foundry classification | Invocation failed (P row 19) | span | `classify_email` and `model.responses_parse` | see above | ERROR | `wab.status` = `error`, `wab.error_type` | `wab.correlation_id` (optional) | App Insights `traces` | EMITTED TODAY. The span records the failure because `_Span.finish` takes the exception. No `classified` event fires. Pradeep's row sets ERROR while its own text argues for WARNING on a 429. See review 4.7. |
| B.7 | Decision gate | Classification decision (P row 20) | span event | `wab.case_intake.classified` | `classify_email` | INFO | `wab.decision_gate`, `wab.halted`, `wab.category`, `wab.final_confidence` | `wab.correlation_id` (optional) | App Insights `traces` | EMITTED TODAY at `foundry/agent/classify.py`:245. Success path only, so it is not an attempt counter. |
| B.7 | Decision gate | Version provenance on the decision | span attribute | `wab.case_intake.classified` | `classify_email` | INFO | `wab.agent_version`, `wab.prompt_version` PROPOSED, `wab.taxonomy_version` PROPOSED | `wab.correlation_id` (optional) | App Insights `traces` | PROPOSED. Gap G13. Without these, an accuracy change cannot be attributed to a prompt change. |
| B.7 | Decision gate | Decision outcome error (P row 21) | none | none | none | none | none | none | none | **Not recommended.** A decision outcome is not a failure. Pradeep's row 21 repeats row 20's justification and defines no error. See review 4.7. |
| B.8 | Build payload | Payload built (P row 22) | span event | `wab.case_intake.step_completed` PROPOSED | none | INFO | `wab.step` = `build_payload`, `wab.payload_kind` PROPOSED, `wab.payload_size_bytes` PROPOSED | `wab.email_id` PROPOSED | App Insights `traces` | PROPOSED. Gap G4. Pradeep's `payload_kind` and `payload_size_bytes` are correct and are kept. His `payload` field is refused as content. See review 4.3. |
| B.8 | Build payload | Payload build failed (P row 23) | span event | `wab.case_intake.step_failed` | `orchestration.step_failed` | ERROR | `wab.step` = `build_payload`, `wab.error_type` | `wab.email_id` PROPOSED | App Insights `traces` | PROPOSED. Gap G4. |
| B.9 | MuleSoft case write | POST succeeded (P row 24) | span | `crm.write_case` PROPOSED | none | INFO | `wab.duration_ms`, `wab.http_status` PROPOSED, `wab.case_id` PROPOSED | `wab.email_id` PROPOSED | App Insights + Dataverse | PROPOSED. Gap G7. This is the true end-to-end success signal and nothing times it today. |
| B.9 | MuleSoft case write | POST failed (P row 25) | span event | `wab.case_intake.step_failed` | `crm.write_case` PROPOSED | ERROR | `wab.step` = `write`, `wab.error_type`, `wab.http_status` PROPOSED | `wab.email_id` PROPOSED | App Insights `traces` | PROPOSED. Gap G4 and G7. Today it appears only as `status=write_failed` in the plain outcome line. |
| B.9 | Latency measure | Receipt to case created | span event | `wab.case_intake.latency` | none | INFO | `wab.latency_seconds`, `wab.latency_within_sla`, `wab.latency_sla_target_seconds` | `wab.correlation_id` (optional) | App Insights `traces` | **EMITTED TODAY** at `src/orchestration/spine.py`:772. Case-created path only, so it measures successes only. Gap G9. Pradeep has no row for latency at all. |
| C | Foundry platform | Foundry agent span | span | `invoke_agent <agent>:<version>` | Foundry's own tree | INFO | `gen_ai.agent.id`, `gen_ai.agent.version`, `gen_ai.response.id`, `gen_ai.operation.name`, `microsoft.foundry.project.id`, start and end time | `gen_ai.response.id` | Foundry diagnostic settings into Log Analytics | **Not emitted by this application.** Foundry emits it. Pradeep's row 26 is the only evidence anyone has gathered on this and it is good work. Two things must be confirmed: that the diagnostic setting targets the same workspace, and that the join key is the response identifier. See review 3.3 and 5.7. |
| D | func_recovery | Run summary (P row 27) | log record | none | none | INFO | `scanned`, `replayed`, `held`, `failed`, `missing`, `source` as printf text | none | App Insights `traces` | EMITTED TODAY at `src/functions/func_recovery/__init__.py`:65. A dashboard must parse the message text, which is fragile. Gap G10. Pradeep's row omits `held`. |
| D | func_recovery | Run summary as an event | span event | `wab.case_intake.recovery_summary` PROPOSED | none | INFO | `wab.recovery_scanned` PROPOSED, `wab.recovery_replayed` PROPOSED, `wab.recovery_held` PROPOSED, `wab.recovery_failed` PROPOSED, `wab.reconciliation_missing` PROPOSED | none | App Insights `traces` | PROPOSED. Gap G10. |
| D | func_recovery | Row retried (P row 28) | span event | `wab.case_intake.row_retried` PROPOSED | none | INFO | `wab.email_id` PROPOSED, `wab.attempt` PROPOSED | `wab.email_id` PROPOSED | App Insights + Dataverse | PROPOSED. Pradeep marks it "already implemented" and cites `recovery/service.py`, which does not exist in this repository. See review 4.10. |
| D | func_recovery | Unhandled exception (P row 29) | log record | none | none | ERROR | free text and traceback | none | App Insights `traces` | EMITTED TODAY at `:59`. |
| D | func_recovery | Max retries exceeded (P row 30) | span event | `wab.case_intake.retries_exhausted` PROPOSED | none | WARNING | `wab.email_id` PROPOSED, `wab.attempt` PROPOSED, `wab.max_retries` PROPOSED | `wab.email_id` PROPOSED | App Insights `traces` | PROPOSED. Pradeep's row 30 carries the only correct defect note in his workbook. See review 3.4. |
| D | Redaction-failed backlog | Rows the recovery scan never sees | none | none | none | none | state store query only | `zenon_uc1_originating_email_id` | Dataverse | **Not measurable from telemetry.** The scan filters on status `ai_complete` at `src/orchestration/recovery.py`:86. A redaction-failed row sits at `created` with recovery status `compensating_failed`, so nothing sweeps it. Gap G14. |
| E | func_feedback | Daily accuracy result | span event | `wab.case_intake.feedback_accuracy` | none | INFO | `wab.feedback_scored`, `wab.feedback_hits`, `wab.feedback_misses`, `wab.feedback_aggregate_accuracy`, `wab.feedback_corrected_fraction` | none | App Insights `traces` | **EMITTED TODAY** at `src/orchestration/feedback.py`:132. **Pradeep's workbook has no group E at all**, so the whole business accuracy report is unspecified in it. See review 4.5. |
| E | func_feedback | Accuracy by subject | span event | `wab.case_intake.feedback_accuracy` | none | INFO | `wab.feedback_by_subject` PROPOSED, `wab.feedback_population_accuracy` PROPOSED | none | App Insights `traces` | PROPOSED. Gap G11. The map is computed and then dropped by the emitter. |
| E | func_feedback | Precision, recall and F1 by subject | span event | `wab.case_intake.feedback_accuracy` | none | INFO | `wab.feedback_macro_f1` PROPOSED | none | App Insights `traces` | PROPOSED. Gap G12. Arun Bathula asked for these three on 30 July for automated drift detection. None exists in the repository. |
| E | func_feedback | Accuracy line | log record | none | none | INFO | `scored`, `hits`, `misses`, `aggregate_accuracy` as printf text | none | App Insights `traces` | EMITTED TODAY at `src/functions/func_feedback/__init__.py`:62. It duplicates the event. |

## 2.3 What the table shows at a glance

Count the rows by status:

- **EMITTED TODAY**: 16 rows. Almost all of them are in the redaction step, the classification step, and the two timer functions.
- **PARTIAL or text-only**: 6 rows. Each of these needs a query that parses a message string.
- **PROPOSED**: 19 rows. Most of the pipeline steps emit nothing at all.

Three observations for the room:

1. **The middle of the pipeline is well instrumented and the ends are not.** Redaction and classification emit rich records. The claim, the data pull, the normalise step, the payload build and the case write emit nothing.
2. **Pradeep's workbook is the mirror image.** It covers the ends in detail and omits the redaction step and the feedback function completely.
3. **The two together are close to a complete design.** Take his step coverage and identifier reasoning, take the code's vocabulary and content discipline, and the result is this table.

---

# Part 3: The span tree

The table lists rows. This section shows the shape.

## 3.1 What the code emits today

An indent means nesting. A record with no indent has no parent.

```
one email (there is no trace object; this is only the reader's view)
│
├─ (no record)                        message received by func_orchestrate
├─ (no record)                        B.1 claim in the state store
├─ (no record)                        B.2 queue item write
├─ (no record)                        B.4 MuleSoft email read
├─ (no record)                        B.5 normalize HTML to text
│
├─ span  orchestration.step_failed         ← only when B.6 redaction fails
│   └─ event wab.case_intake.step_failed        wab.step = "redact"
│                                              the email stops here
│
├─ span  classify_email                    ← B.7
│   ├─ span  model.responses_parse              gen_ai.usage.input_tokens
│   │                                           gen_ai.usage.output_tokens
│   │                                           wab.duration_ms
│   └─ event wab.case_intake.classified         wab.decision_gate
│                                               wab.final_confidence
│
├─ span  guardrails.evidence               ← B.6 evidence, emitted AFTER B.7
│   └─ event wab.case_intake.guardrail_evidence
│                                               wab.redaction_count
│                                               wab.redaction_billable_text_records
│
├─ (no record)                        B.8 build payload
├─ (no record)                        B.9 write the case to CRM
│
├─ event wab.case_intake.latency           ← only when a case was created
│                                               wab.latency_seconds
│                                               wab.latency_within_sla
│
└─ log record  "func_orchestrate: status=... source_id=... case_id=..."
```

Two things to say aloud when you show this:

- The tree is not a real tree. There is no trace object and no identity on a span. `wab.parent` carries only the parent's **name**, so two emails processed at the same time produce two `model.responses_parse` records that both say their parent is `classify_email`, and nothing says which one belongs to which email. Only `wab.correlation_id` separates them, and it is optional.
- The `guardrails.evidence` branch is drawn after `classify_email` on purpose. That is the real order.

## 3.2 What the tree looks like after the proposed gaps are closed

```
event wab.case_intake.run_started              wab.email_id, wab.attempt
│
├─ event step_completed  wab.step = "claim"
├─ event step_completed  wab.step = "queue_item"
│
├─ span  crm.read_email                        wab.duration_ms, wab.http_status
│   └─ event step_failed  wab.step = "data_pull"      (on failure)
│
├─ event step_completed  wab.step = "normalize"
│
├─ span  guardrails.evidence
│   ├─ event guardrail_evidence
│   └─ event step_failed  wab.step = "redact"         (on failure, email stops)
│
├─ span  classify_email                        wab.email_id
│   ├─ span  model.responses_parse             tokens, gen_ai.response.id
│   │                                          └─ joins to the Foundry span
│   └─ event classified                        + wab.prompt_version
│                                              + wab.taxonomy_version
│
├─ event step_completed  wab.step = "build_payload"
│
├─ span  crm.write_case                        wab.duration_ms, wab.case_id
│   └─ event step_failed  wab.step = "write"          (on failure)
│
├─ event latency                               on every terminal path
│                                              + wab.outcome
│
└─ event wab.case_intake.run_completed         wab.outcome, wab.case_id
```

Every name in this second tree that is not in the first tree is PROPOSED.
`05-gap-analysis.md` gives the change for each one.

## 3.3 The one field that makes the whole tree work

Look at what the second tree adds first, on the top line: `wab.email_id`.

One field, carried on every record, turns a set of unrelated log rows into one email's story.
It is the field the team already chose on 17 July, when it decided that the email identifier is the tie-together key across systems.
It is the field Pradeep's workbook puts on almost every row.
It is the field the application does not emit anywhere.

If only one change is made from this whole folder, make that one.
