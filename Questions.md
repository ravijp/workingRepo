# Rovo Queries — Iterations 3, 4, and 5 (consolidated)

**How to use this file.** Single VDI Rovo pass. Five query groups, each pasted into Rovo as one prompt. Screenshot the full Rovo answer (including source citations panel) and save to this folder as `rovo_iter345_G<N>_screenshot.png`. Distil findings into `rovo_findings_iterations_3_4_5.md` with sources next to each fact.

**Why one file across three iterations.** The validation needs across iterations 3, 4, and 5 are tightly coupled — the bb_opportunityservice schema (G-4.1) determines write-back syntax for both iterations 4 and 5; the MuleSoft DAO contract (G-5.1) is the back-half of the architecture; the smaller targeted lookups (G-3.3 FCM picklist, G-3.4 IBS Connector, G-3.5 ACH Tracker) feed the joins iteration 4 surfaces in its Outstanding Items panel and iteration 5 may need to reference. Running one consolidated Rovo pass costs one VDI session and produces all the Confluence-side evidence the three iterations need.

**Five groups total.** Aligned to the claim markers across `iteration_3_external_joins.md`, `iteration_4_arw_prefill.md`, and `iteration_5_submit_mulesoft.md`.

| Group | Targets | Iterations claim-mapped |
|---|---|---|
| G-3.3 | FCM document_type picklist for PMC Control Prong | Iteration 3 §5 (claim_3_1) |
| G-3.4 | IBS Connector contract from ops5/NSF for Controlling Individual lookup | Iteration 3 §6 (claim_3_2) |
| G-3.5 | ACH Tracker programmatic surface or confirmation of UI-only | Iteration 3 §7 (claim_3_3) |
| G-4.1 | bb_opportunityservice Dataverse entity schema (central unknown) | Iteration 4 §1, §6 (claim_4_1) |
| G-5.1 | wab-ent-digitalacntopen-eapi contract + AABOS Second-Review tooling | Iteration 5 §3, §6 (claim_5_1, claim_5_2, claim_5_3) |

---

## G-3.3 — FCM document_type picklist for PMC Control Prong records

**Why we need this.** Iteration 3 §5 (FCM Control Prong lookup) calls `POST /fcm/documents/query` with PMC bank account numbers and filters by document_type to retrieve the PMC's Control Prong individual reference. Rovo iteration-2 G6 confirmed the FCM System API exists (`wab-content-management-fis-sapi`), exposes the query endpoint, and accepts SSN/account number/CIF as query keys. What it did not return was the controlled picklist of `document_type` values applicable to PMC Control Prong / Authorization records — the codes the agent passes to filter the query. The Middesk-side codes are documented (`CU0304` for ID Verification/OFAC; `CU0101` for Formation Documents per Rovo Q3.7), but the PMC-Control-Prong-specific codes are the gap.

**Query to paste into Rovo:**

> For WAB's FIS Content Manager (FCM) and the MuleSoft `wab-content-management-fis-sapi` System API, please share the complete document_type picklist and the codes specifically associated with PMC Control Prong individual records. Specifically:
> (a) What is the complete controlled list of `document_type` values that `POST /fcm/documents/query` and `POST /fcm/document` accept? Is the picklist available via `GET /fcm/documentTypes`, and if so, share the response shape.
> (b) Which document_type values are used for PMC Control Prong individual / Authorization records — the documents §3.1.1 of the New Account procedure references when it says "Confirm Control Prong individual in D365 or Image Centre > Document Research > under the PMC TIN"?
> (c) When a PMC's Control Prong individual changes (e.g., the PMC names a new authorized individual), is a new Control Prong document filed under the PMC's existing account number, or under a new account number, or under a separate Control Prong-specific record key?
> (d) Are these FCM document_type codes the same as the D365 CRM "File Document Type" dropdown values bankers select when uploading attachments, or are they two separate controlled lists (as D-007 in `decisions_log.md` established for the broader case)?
> (e) Beyond Middesk's CU0304 / CU0101 codes already documented, what is the complete CU-prefixed code list, with semantics?
> Please cite the Confluence pages directly (FCM System API documentation, FIS - Content Manager / ImageCentre Runbook, FCM Business Customer/Account Document Service).

---

## G-3.4 — IBS Connector contract from ops5/NSF for Type A Controlling Individual lookup

**Why we need this.** Iteration 3 §6 (IBS Insight read for Type A entities) inherits the IBS Connector from ops5/NSF per the banker_assistant_overview's shared-substrate framing. The agent calls `ibs_connector.get_customer_relationship(cis_number, fields=[ControllingIndividual, CustomerStatus, RelationshipHistory, AuthorizedSigners])`. The exact contract shape — field names, types, authentication, error handling, latency profile — comes from the ops5/NSF integration design; Iteration 3 documents it as `claim_3_2`. If WAB Confluence carries the ops5/NSF integration design, that's the cleanest resolution; if not, this becomes a cross-Zenon-workstream lookup against ops5's iteration artifacts.

**Query to paste into Rovo:**

> For the ops5 NSF Decisioning project (or AAB NSF / AAB Non-Posted Decisioning, as it may be named in Confluence) and the IBS Insight integration that project consumes, please share what is documented about the IBS Connector contract. Specifically:
> (a) What is the integration pattern between Zenon's / WAB's NSF Decisioning agent and IBS Insight — is it a MuleSoft EAPI, a direct API call, an SDK, or an SQL-style read against IBS's database surface? What is the endpoint name?
> (b) What fields does the connector expose per customer record — specifically: CIS Number, Controlling Individual, Customer Status, Relationship History, Authorized Signers, and any related fields the HOA New Account Onboarding agent could consume?
> (c) What is the authentication pattern (service account credentials, OAuth client credentials, MuleSoft client_id + client_secret), and what is the latency profile per call?
> (d) Are there documented sample request/response payloads? Is there a sandbox / QA tier?
> (e) Which other projects at WAB consume IBS Insight programmatically today — and is the IBS Connector built as a shared enterprise service or as a project-specific integration?
> Please cite the Confluence pages directly (NSF Decisioning PDD/SDD, IBS Insight Runbook, MuleSoft API Inventory for IBS-related endpoints, AAB Decisioning project docs).

---

## G-3.5 — ACH Tracker programmatic surface

**Why we need this.** Procedure §3.3.5 documents the best-practice CoID uniqueness check via ACH Tracker before submitting accounts with ACH origination. The agent's Iteration 3 join (f) needs to know whether ACH Tracker has a programmatic API, a Dataverse-resident data model, or is strictly a UI-only tool. The pilot default per D-023 is to surface the check as a banker-acknowledgment Outstanding Item; a programmatic surface would let the agent automate it and would be a meaningful capability upgrade.

**Query to paste into Rovo:**

> For WAB's "ACH Tracker" — the tool the New Account Desktop Procedure §3.3.5 references when it advises bankers to "search in ACH Tracker to confirm the CoID number is not being used" before submitting accounts with ACH origination — please share what is documented about the tool. Specifically:
> (a) What is ACH Tracker? Is it a SharePoint list, a custom Power App, a separate web application, an Excel workbook on a shared drive, a D365 entity, a custom database — or some combination?
> (b) Does it expose a programmatic API, a Dataverse-resident schema, an ODBC-readable database, or any other surface a service-account-authenticated client could query?
> (c) What is the data model — what fields does each ACH Tracker entry carry (CoID, associated PMC, associated entity, status, dates, banker who registered it)?
> (d) Who owns ACH Tracker (business owner, technology owner, SME)? Where does the source-of-truth ACH origination registry live — in ACH Tracker, in IBS, in FIS, in Fiserv, elsewhere?
> (e) Is there a documented "ACH Tracker guide" the procedure §3.3.5 references? If so, what does it instruct bankers to check, and is there any documented automation around the check?
> (f) For the HOA AI New Account Review UiPath bot, is ACH Tracker in scope as an application the bot interacts with, or is it deliberately out of scope?
> Please cite the Confluence pages directly (ACH Tracker guide, ACH Origination Runbook, AAB Operations procedures, HOA AI New Account Review PDD).

---

## G-4.1 — Complete bb_opportunityservice Dataverse entity schema and ARW relationship structure

**Why we need this.** This is the single most consequential lookup for Iteration 4 (and downstream for Iteration 5's MuleSoft DAO call payload). Rovo Q3.10 named the entities involved (`AAB New Accounts Request`, `HOA case`, `PMC`, `Child Company`, `Control Prong`, `Accounts Requested`) but did not return field-level schema. Rovo Q3.12 surfaced `bb_opportunityservice` by name. Iteration 4's pre-fill writes ~30+ fields across the parent ARW entity, the Accounts Requested rows, and the Child Company entity; without the schema, the field-mapping table is logical, not literal.

**Query to paste into Rovo:**

> For the AAB Operations Automation D365 solution (owned by Janus Lund, Customer Service module, Silver tier) and specifically the `bb_opportunityservice` Dataverse entity that represents the AAB New Accounts Request Client Action, please share the complete entity field schema. Specifically:
> (a) What is the complete column list for `bb_opportunityservice` — logical names (e.g., `bb_request_notes`, `bb_mgmt_company`), display names, types (string, lookup, picklist, decimal, datetime, etc.), required-vs-optional flags, picklist values where applicable (e.g., the full `Type of Entity` dropdown referenced in §3.3.4 — Nonprofit-Corp, Nonprofit-Assoc/Org, etc.)?
> (b) What is the relationship structure: which related entities does `bb_opportunityservice` reference (PMC / `account`, Child Company / `account`, Control Prong / `contact`, Accounts Requested / what entity?), via what relationship attributes (`@odata.bind` lookup column names)?
> (c) What is the Accounts Requested child entity — is it `bb_aab_account_request`, `bb_accounts_requested`, or another logical name? Same schema breakdown: column list, types, picklists for Account Type, Approved Rate, Interest Plan, Lockbox, ACH.
> (d) What are the workflow / status_reason values on `bb_opportunityservice` for the documented workflow bar states (New / Submitted / Processed / CD Funding / Ready for Delivery / Completed/Cancelled per §3.3)?
> (e) Are there documented Dataverse plug-ins, custom workflows, or business rules attached to `bb_opportunityservice` that affect the write semantics (e.g., when the agent PATCHes a field, does a plug-in fire that recomputes a dependent field)?
> (f) What are the Case (`incident`) status_reason picklist values used by the AAB-specific lifecycle — specifically the values for "Waiting on Customer", "AABOS Review", "Partial Account Creation" if defined, "Ready for Delivery", "Resolved"? (This also resolves D-011 in `decisions_log.md`.)
> (g) For the File Document Type attachment picklist (D-012 in `decisions_log.md`), what is the complete controlled list of values, and on which entity does it live (Annotation, activitymimeattachment, custom)?
> Please cite the Confluence pages directly (AAB CRM solution documentation, D365 entity reference pages, Janus Lund's team's solution-config pages, AAB Operations Automation Project Intake) and where relevant include screenshots of the D365 solution explorer / customization views.

---

## G-5.1 — wab-ent-digitalacntopen-eapi MuleSoft contract + AABOS Second-Review tooling

**Why we need this.** Two related needs in one group because both center on the post-Submit handoff: the MuleSoft DAO call shape and the AABOS-side review surface that engages with the agent's outputs.

**For the MuleSoft DAO call (claim_5_2, claim_5_3 in Iteration 5).** The endpoint name `wab-ent-digitalacntopen-eapi` is established per `_STATUS.md` and Rovo Q3.9's documented target-state flow (Customer Search → Customer Creation → Account Number Generator → Orchestration EAPI → FCM API). The contract — request body schema, 200/206/400 response shapes, authentication, retry semantics, idempotency at row level — is the gap. The 206 partial-content handling is the non-trivial case Iteration 5 §3.2 designs around.

**For AABOS Second-Review (claim_5_1).** AABOS's existing review tooling shapes how the agent presents its outputs (Outstanding Items panel, TraceLog, external reference IDs). Pilot default is "D365 timeline + Smart Assist summary"; if AABOS uses a different surface (BAM+, Power BI, SharePoint queue, custom canvas app), the integration shape changes.

**Query to paste into Rovo:**

> For WAB's MuleSoft `wab-ent-digitalacntopen-eapi` endpoint and the AABOS Second-Review process at AAB Operations, please share what is documented. Specifically:
>
> **Part A — wab-ent-digitalacntopen-eapi contract:**
> (a) What is the endpoint URL and method? Is the RAML/OAS contract published in MuleSoft Anypoint Exchange or in Azure DevOps? Which user stories / change records cover its delivery (per `_STATUS.md` there is a US 527655 reference — confirm scope)?
> (b) What is the request body schema — top-level fields (customer / accountsRequested / controlProng / configuration / compliance reference IDs) and field types?
> (c) What are the documented response shapes for 200 (full success — account numbers, CIS number, FCM filings), 206 (partial — some accounts opened, some failed; what errorCodes are documented), 400 (validation failure), and 5xx (system error)?
> (d) Is the EAPI idempotent at the row level — when the agent re-submits a payload that contains a `rowId` for which an account has already been created, what happens?
> (e) What is the authentication pattern (OAuth client credentials via MuleSoft? Bearer token? client_id + client_secret)? What is the documented latency SLA and retry posture?
> (f) Which projects consume this EAPI today — AABOS only, or also DAO, BDAO, CDAO? Is the EAPI live in production, in QA, or planned?
>
> **Part B — AABOS Second-Review tooling and process:**
> (a) When a Case lands in AABOS's queue after the banker Submits, what tool(s) does AABOS use to perform the Second Review documented in Slide 19 #7 — D365 directly, BAM+, a SharePoint queue, a custom canvas app, a Power BI dashboard, something else?
> (b) Where does the AABOS review tracking live — is "Needs Correction" a status on the Case, an Annotation, an entity field, a separate workflow?
> (c) What does AABOS check during the Second Review — is there a documented Second-Review checklist?
> (d) After Second Review, the procedure §3.3.11 says "the AABOS team will send it back with 'Needs Correction' and the banker will need to fix the client action and re-submit it" — what is the technical implementation of the "send back" gesture? Status change? Notification? Email?
> (e) What is the Auto-Process click's technical implementation today — does it trigger the SSOT manual pipeline, or the new MuleSoft EAPI? Is there a feature-flag controlling which path runs?
>
> Please cite the Confluence pages directly (MuleSoft APIs for AAB Account Opening Business Case, AAB Account Opening MuleSoft Integration design, AABOS Operations procedures, Auto-Process documentation if any, BAM+ documentation if relevant) and include any screenshots of API definitions, AABOS review tools, or process flow diagrams.

---

## Status

- **Loop step:** 2 (Surface validation needs) — this file written.
- **Next action:** Ravi runs G-3.3 through G-5.1 on VDI Rovo. Screenshots saved as `rovo_iter345_G<X>_screenshot.png` in this folder. Findings distilled into `rovo_findings_iterations_3_4_5.md` with sources per fact. Then reconcile (loop step 4) — fold findings back into iterations 3, 4, 5; promote proposed decisions D-022 through D-027 to decided status where evidence supports; demote unresolved claims to explicit open questions in `decisions_log.md`.
- **Note on web queries:** the consolidated web queries file (`web_queries_iterations_3_4_5.md`) carries one fallback web query for the wab-ent-digitalacntopen-eapi contract if a publicly-accessible RAML exists. Most other claims resolve through Rovo or fall through to Chris-direct asks.
