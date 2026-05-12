# Rovo Queries — Iteration 2

**How to use this file.** Open Rovo chat on VDI. Run each query group as a single Rovo prompt. After each response, screenshot the full Rovo answer (including the source citations panel) and save the screenshot into this folder with a filename matching the query group ID (e.g., `rovo_iter2_G1_screenshot.png`). Then paste the key findings into `rovo_findings_iteration_2.md` or note which screenshots to read.

Queries are grouped so each Rovo session produces multiple useful answers. **Six groups total.** Aligned to the validation needs surfaced by `iteration_2_documents.md` §8 (Open questions Q-2.1 through Q-2.6).

---

## G1 — KYC handling: external vs internal artifact, and HOA AI New Account Review's documented stance

**Why we need this.** Q-2.1. `iteration_2_documents.md` §2.6 routes KYC as an internal banker-completed form (inverted flow — the agent generates rather than extracts) based on the Oak Hill case's Entity KYC Profile Form HOA. The procedure §3.1.1 states *"KYC: Required only for a New Non-HOA entity"*, which contradicts the Oak Hill case having one. The most likely reconciliation is two artifacts sharing a name — external client-completed KYC (Non-HOA only) and internal banker-completed KYC (both HOA and Non-HOA). Confirmation lets the agent's KYC routing close cleanly.

**Query to paste into Rovo:**

> For the HOA AI New Account Review project and the AAB New Account procedure, please share what is documented about KYC handling. Specifically:
> (a) Does the project distinguish between an *external* (client-completed) KYC form and an *internal* (banker-completed) KYC form? If so, what are the field sets for each, and which entity types (HOA vs Non-HOA) require which form?
> (b) For HOA entities specifically, is there an internal Entity KYC Profile Form HOA that the bank completes, or does the procedure's §3.1.1 "Required only for a New Non-HOA entity" rule mean HOAs do not have any KYC artifact at all?
> (c) For the to-be UiPath bot, what KYC handling is in scope and what is out of scope? Is KYC adjudication done by the bot, by the banker, or by a separate compliance / KYB process?
> (d) Are LexisNexis IDV, OFAC, and PEP screening results stored on the KYC form itself, on the Case, or in a separate compliance system?
> Please cite the Confluence pages directly and include the field-level documentation if available.

---

## G2 — Per-state SOS sample coverage and the HOA AI New Account Review IXP training corpus

**Why we need this.** Q-2.2. `iteration_2_documents.md` §2.1 plans a 5-state training-and-evaluation corpus (FL, OR, TX seen in discovery plus likely CA, NY, or AZ) for the SOS extractor. Rovo Q2.4 established that UiPath IXP requires 100+ samples per doc type — but per *jurisdiction* coverage is unknown. The HOA AI New Account Review project's SDD likely documents which states the IXP extractor was trained against; cross-referencing tells us whether Zenon's pilot corpus overlaps, complements, or fills a gap.

**Query to paste into Rovo:**

> For the HOA AI New Account Review project's IXP extractor for Secretary of State documents, please share what is documented about training corpus composition. Specifically:
> (a) Which states' SOS filings are included in the IXP training corpus today? Is the coverage equal per state, weighted by HOA volume, or limited to a few high-volume jurisdictions?
> (b) What is the documented minimum sample count per state for the SOS extractor to ship, given the broader 100+ samples per doc type rule for UiPath IXP?
> (c) Has the project documented per-state accuracy after training — are there states the extractor performs poorly on?
> (d) What is the process for adding a new state to the corpus once the extractor is in production — is it a model retrain, a corpus extension, or a separate document model?
> Please cite the Confluence pages directly (HOA AI New Account Review PDD / SDD).

---

## G3 — Documented extraction field lists for Management Agreement, SOS, and Recert email

**Why we need this.** Q-2.3. `iteration_2_documents.md` §2.1, §2.2, §2.4 specify per-doc-type schemas drawn from `onboarding_doc_review_deep_dive.md` (which derived from the three discovery cases plus the procedure). The HOA AI New Account Review project's SDD documents its own extraction field list for the three in-scope doc types (Management Agreement, SOS via Middesk API, Recert email per Rovo Q2.3); cross-referencing catches any field the agent is missing and any field whose canonical name should be aligned with WAB's vocabulary.

**Query to paste into Rovo:**

> For the HOA AI New Account Review project, please share the documented extraction field list for each in-scope document type. Specifically:
> (a) For Management Agreement, what fields does the IXP extractor produce (effective date, association legal name, manager name, signatures, term, recitals — what is the complete list, with field names, types, and required-vs-optional flags)?
> (b) For SOS (whether retrieved via Middesk API or extracted from an attachment), what fields does the project consume — entity legal name, filing number, formation date, FEIN, principal address, registered agent, entity type, status, history?
> (c) For Recert email (the certification response from the client), what fields does the bot extract, and how does it classify the three certification statements (complete-and-correct, seven-day-notice, no-25-percent-owner)? Is the response classification per-statement or a single whole-email verdict?
> (d) Is there a documented mapping from these extraction outputs to the ARW Dataverse field schema?
> Please cite the Confluence pages directly (PDD / SDD) and include any data dictionary or field-mapping table.

---

## G4 — LexisNexis IDV integration: pattern, contract, ownership

**Why we need this.** Q-2.4. `iteration_2_documents.md` §2.6 routes the KYC A5 field through a LexisNexis IDV API call that lives in Iteration 3 — but Iteration 2's KYC-generation routing surfaces the dependency. Need the integration shape (REST EAPI, MuleSoft-fronted, contract-shape sample) so Iteration 3 can plan against it rather than against a placeholder.

**Query to paste into Rovo:**

> For WAB's LexisNexis Identity Verification integration, please share what is documented about the integration pattern and contract. Specifically:
> (a) Is the LexisNexis IDV call wrapped behind a MuleSoft EAPI, a direct REST integration, or another pattern? What is the endpoint name (e.g., `wab-az-...-eapi`)?
> (b) What is the request payload — what fields does the bank send (entity name, principal address, individual name + DOB + SSN if applicable)?
> (c) What is the response payload — what fields come back, what does a "verified" response look like vs a "discrepancy" or "no match", and how are confidences expressed?
> (d) Which projects currently use LexisNexis IDV at WAB? Is the integration shared as an enterprise service or is each project's integration its own?
> (e) Are there documented latency SLAs and throughput limits?
> Please cite the Confluence pages directly (LexisNexis Runbook, EAPI documentation, Alloy Runbook if applicable).

---

## G5 — OFAC and PEP screening API surface at WAB

**Why we need this.** Q-2.5. Same shape as G4 — Iteration 2's KYC-generation routes D1 (OFAC) and D2 (PEP) through external API calls that Iteration 3 will run. Rovo Q1.5 + D-006 already established that Middesk's Watchlist/OFAC service was turned OFF at WAB on 2025-04-01 and that OFAC happens elsewhere; we need the elsewhere.

**Query to paste into Rovo:**

> For WAB's OFAC and PEP screening for business and individual customers, please share what is documented about the API surface used today. Specifically:
> (a) Which system performs OFAC screening at WAB today (now that Middesk's Watchlist service was turned off 2025-04-01)? Is it MuleSoft-fronted, MS Dynamics-native, or a separate compliance system?
> (b) What is the API contract shape — request and response — for an OFAC screen on a business name? On an individual?
> (c) Same question for PEP (Politically Exposed Person) screening — which system, what contract?
> (d) Are screening results persisted on the Case in D365, in a separate compliance application tier, or both?
> (e) Which projects currently consume the OFAC / PEP APIs? Is the integration shared or per-project?
> Please cite the Confluence pages directly (compliance Runbooks, EAPI documentation).

---

## G6 — Image Center programmatic access under PMC TIN

**Why we need this.** Q-2.6. The procedure §3.1.1 instructs bankers to *"Confirm Control Prong individual in D365 or Image Centre > Document Research > under the PMC TIN"*. Iteration 3 plans to perform this lookup programmatically. Need to confirm whether Image Center exposes a programmatic API or is strictly UI-only — and if UI-only, the implication is the agent must surface this as a banker action rather than executing it.

**Query to paste into Rovo:**

> For WAB's Image Center (the document research / archival system used in the New Account procedure §3.1.1), please share what is documented about programmatic access. Specifically:
> (a) Does Image Center expose a REST API, a SOAP service, a queue-based interface, or any other programmatic surface, or is it strictly a UI tool accessed by bankers through the standard application?
> (b) If programmatic access exists, what is the authentication pattern, the supported queries (search by TIN, search by document type, etc.), and the response shape?
> (c) For the HOA AI New Account Review project specifically, does the to-be UiPath bot read from Image Center, or is Image Center access deliberately left as a banker manual step?
> (d) Is there a known roadmap item to expose Image Center programmatically — for example, as part of a broader content-services modernization?
> Please cite the Confluence pages directly (Image Center Runbook, HOA AI New Account Review PDD).

---

## Status

- **Loop step:** 2 (Surface validation needs) — this file written.
- **Next action:** Ravi runs G1–G6 on VDI Rovo. Screenshots saved as `rovo_iter2_G<N>_screenshot.png` in this folder. Findings distilled into `rovo_findings_iteration_2.md` with sources per fact.
