# Rovo Queries — Iteration 1

**How to use this file.** Open Rovo chat on VDI. Run each query group as a single Rovo prompt. After each response, screenshot the full Rovo answer (including the source citations panel) and save the screenshot into the same folder with a filename matching the query group ID (e.g., `rovo_iter1_G1_screenshot.png`). Then come back here and either paste the key findings into `rovo_findings_iteration_1.md` or just let me know which screenshots to read.

Queries are grouped so each Rovo session produces multiple useful answers. **Six groups total.**

---

## G1 — UC1 hand-off contract to downstream agents

**Why we need this.** Iteration 1 starts at UC1's output boundary. We need to know whether UC1's hand-off to a downstream agent (like the HOA New Account agent) is purely an in-Dataverse state change, or whether UC1 emits any structured payload, log entry, or event. Affects claims C1.1, C1.4, C1.20.

**Query to paste into Rovo:**

> For the AAB Case Intake & Routing project (UC1), please share what is documented about how UC1 hands off a case to a downstream agent or workflow after it has classified the email and created the case. Specifically:
> (a) Does UC1 emit a structured payload, event, or log entry when classification completes, or does it simply set the Subject field on the Case and rely on downstream consumers to observe the state change?
> (b) Are confidence levels (High / Medium / Low) or rationale text captured anywhere on the Case record — as custom fields, as a Note (Annotation), or in the Case timeline?
> (c) Is there a documented "ready for downstream agent" Case state, or is the Case simply created with Subject set and Status = In Progress?
> (d) For cases where UC1's prediction is low-confidence, what is the documented downstream behavior?
> Please cite the Confluence pages directly and include screenshots of any sequence or architecture diagrams.

---

## G2 — D365 Case email-attachment storage at WAB

**Why we need this.** The agent reads attachments via Annotations. We claimed (C1.3, C1.10) that D365 stores attachments as Annotations with documentbody base64 and that the controlled File Document Type dropdown values include SOS, Management Ag..., Recert, KYC, Misc, Email Request fr.... Need to confirm against WAB's actual D365 configuration.

**Query to paste into Rovo:**

> For the AAB D365 CRM, please share documentation on how email attachments are stored on Case records, specifically:
> (a) Are attachments stored as Annotation records on the Case directly, on the Email Activity, or both? Is there a configuration setting that controls this?
> (b) The HOA New Account Review process documents a "File Document Type" dropdown that bankers select when uploading attachments. What is the complete controlled list of values for this dropdown — please provide all entries, not a partial list.
> (c) Is "CP575" or "IRS Notice" present in that controlled list today? If not, is there a documented process for proposing new File Document Type values?
> (d) Is there an automated process that classifies attachments at upload, or is the File Document Type entirely banker-selected today?
> Please cite the Confluence pages directly.

---

## G3 — Power Automate / Dataverse webhook surface for external agents

**Why we need this.** Claim C1.2 — the agent needs some mechanism to react to Case state changes. Need to know what WAB has actually deployed and approved for external integrations into D365 Case events.

**Query to paste into Rovo:**

> For WAB's Dynamics 365 CRM, please share documented patterns for an external service or agent to react to Case record events (creation, Subject field set, Status change). Specifically:
> (a) Has WAB deployed Power Automate flows that trigger on Case events? What is the standard pattern — flow runs in WAB tenant, calls an external HTTP endpoint?
> (b) Are Dataverse webhooks (the webhook subscription model) used at WAB today? Any documented approvals or restrictions for external service subscriptions?
> (c) Is there a documented integration pattern for an agent or RPA bot to be triggered by a Case state change in D365? Please cite any existing integrations that follow this pattern.
> (d) What WAB security or compliance team owns approvals for external services subscribing to Dataverse events?
> Please cite the Confluence pages directly.

---

## G4 — Banker-facing UI surface in D365 for AI-drafted content

**Why we need this.** Claim C1.15 — the agent drafts a clarification email and needs to surface it to the banker for one-click review-and-send. The mechanism could be an Annotation, an Adaptive Card, an Outlook Connector cue, or a custom Power App embedded in the Case form. Need to know what WAB has done before for similar UI surfaces.

**Query to paste into Rovo:**

> For WAB's Dynamics 365 CRM and Outlook Connector, please share what is documented about surfacing AI-drafted content (suggested replies, drafted emails, suggested values) to bankers within the Case form or Outlook. Specifically:
> (a) Has WAB deployed any custom Case form extensions, embedded canvas apps, or Adaptive Cards for AI suggestions or drafted content?
> (b) The Outlook Connector at WAB — does it support custom pane content beyond CRM lookups (e.g., a "Suggested draft" panel)?
> (c) Is there a documented pattern for "agent drafts an email → banker reviews → one-click send" in D365 — through the existing send-as-draft flow, through Adaptive Cards, or some other mechanism?
> (d) For the HOA AI New Account Review project specifically, what is the planned UI surface where bankers will see the bot's outputs and approve them?
> Please cite the Confluence pages directly.

---

## G5 — UiPath IXP as a callable extraction/classification service

**Why we need this.** Claim C1.12 — the agent positions UiPath IXP as a callable classifier for the 3 in-scope doc types. Need WAB's actual deployment shape — is IXP exposed as an HTTP API, an Orchestrator queue, or only inside RPA processes? Affects integration design for Iteration 2.

**Query to paste into Rovo:**

> For WAB's UiPath Automation Platform, specifically the Document Understanding / IXP capability, please share what is documented about programmatic invocation. Specifically:
> (a) Can UiPath Document Understanding / IXP be invoked synchronously from an external service via an HTTP API, or does invocation require an RPA process queued in Orchestrator?
> (b) For the existing IXP use cases at WAB (Corporate Trust Indenture, SOC Review Template, HOA Lockbox AI, Settlement Services Onboarding, HOA AI New Account Review), what is the documented invocation pattern — queue-based, API-based, both?
> (c) Is there a documented "shared service" pattern where one project's IDP extraction model can be called by another project (e.g., Zenon's agent calling the HOA AI New Account Review extractor for Management Agreement)?
> (d) What are the documented latency expectations for an IXP extraction call?
> Please cite the Confluence pages and any architecture diagrams.

---

## G6 — D365 Case state machine and "Waiting on Customer"

**Why we need this.** Claim C1.16 — the agent transitions the Case to "Waiting on Customer" on the Yellow path and waits for the client reply. Need to confirm this state exists in WAB's D365 case state machine and that there's a documented reply-routing that re-attaches replies to the same Case.

**Query to paste into Rovo:**

> For WAB's AAB Operations D365 CRM, please share documentation on the Case state machine and the "Waiting on Customer" state. Specifically:
> (a) Is "Waiting on Customer" a documented Status or Status Reason value on the Case entity? What are the transitions in and out?
> (b) When a banker sends an email from a Case to a client and the client replies, how is the reply re-attached to the original Case — through the original email's Conversation Index, through a tracking token in the subject line, through banker-manual association, or some other mechanism?
> (c) Is there a documented "Convert email to Case" rule that handles reply-routing automatically based on the original Case ID being in the email thread?
> (d) For the HOA New Account flow specifically, what is the documented case state when the banker is waiting on the client's Cert email reply?
> Please cite the Confluence pages directly.

---

## How to record findings

After running these six groups, please either:

1. **Drop screenshots** into this folder with names `rovo_iter1_G1.png`, `rovo_iter1_G2.png`, etc. (one or more per group is fine), and tell me they're in. I'll read them and distill into `rovo_findings_iteration_1.md`.
2. **Or paste the key text** from each Rovo response directly into `rovo_findings_iteration_1.md` under headings `## G1 Findings`, `## G2 Findings`, etc. — your choice.

Either way, what I need at minimum per group: the **bottom-line answer** Rovo gave, the **Confluence page names** it cited (verbatim), and any **screenshots of diagrams** if Rovo surfaced architectural pictures.
