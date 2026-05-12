Batch 1 — Settle the existing builds and their status
Before designing anything, we need to know what's already built, what's planned, and what's paused. These five questions reveal whether Zenon is building net-new, replacing a UiPath implementation, or layering on top.

"What is the current production / pilot / paused status of the HOA AI New Account Review project? Has BOT1 / BOT2 been deployed to production? What is the current monthly volume processed?"

"What were the Known Exceptions and pain points reported during UAT and post-deployment for the HOA AI New Account Review? Are there any logged exception reports or summary metrics from production runs?"

"For the Settlement Services Onboarding MVP and Phase 2 PDDs — what is the current status, what doc types are live, and what extraction accuracy has been observed in production? Are there any post-mortem or lessons-learned documents?"

"What is the relationship between the HOA AI New Account Review project and the broader HOA Operations AI strategy deck owned by Chris Crawford? Is the existing UiPath bot considered the long-term solution, or is it expected to be replaced?"

"Show me the Middesk Documentation Automation page in full — specifically the green-light / yellow-light / red-light decision rules, the field-level comparison logic between MidDesk and the Management Agreement, and any documented false-positive or false-negative rates."

Batch 2 — IDP and document-extraction surface
The slide says "Review documents for completeness and accuracy using IDP." We need to know exactly what IDP is, who owns it, and what its current document coverage is.

"What IDP platform does WAB use today — UiPath Document Understanding, Azure AI Document Intelligence, ABBYY, AWS Textract, or something else? Who is the platform owner? What is the licensing / capacity model?"

"List every IDP extractor or document model currently in production at WAB, with: doc type, owning project, accuracy in production, monthly volume processed."

"For each of these HOA New Account document types, does an IDP extractor exist today, is one in development, or is one not yet scoped: Management Agreement, Secretary of State filing, Articles of Incorporation, Bylaws, CP575 IRS letter, BOC, HOA Signature Card, KYC, Cert email response, Driver License / Passport / ID documents?"

"What is the IDP training-sample requirement per document type at WAB? Who is responsible for labeling training data? What is the typical timeline from start labeling to model in production?"

"How does IDP today hand off extracted fields to D365 CRM — direct API write, MuleSoft staging, RPA, manual review queue? Show me the integration pattern document or architecture diagram."

Batch 3 — MuleSoft, IBS, and ARW auto-processing
The slide calls out "MuleSoft-enabled auto-processing" and "MuleSoft / IBS — account creation." We need to know what API surface MuleSoft exposes today and what the contract for "Submit ARW → accounts created" looks like.

"Show me the MuleSoft API specification for the AAB New Accounts Request (ARW) Auto-Process flow. What fields does it require as input? What does it return? What are the failure modes and how are they surfaced back to the case?"

"What is the Dynamic Deposit API mentioned in the Account Automation Overview? Is it the same as MuleSoft, a layer on top, or a separate service? Who owns it? What is its production status?"

"What is the contract between AABOS Doc Review and MuleSoft auto-processing? Specifically: after AABOS clicks Auto-Process, what fields on the ARW must be populated, what attachments must be present, what document tags must be set, and what validation does MuleSoft run before invoking the Dynamic Deposit API?"

"For the AAB New Accounts Request entity in D365 — what is the full field schema, what fields are required vs. optional, what fields are populated by auto-pop logic, and what fields are calculated downstream by MuleSoft? Ideally I want the Dataverse entity metadata or solution export."

"What is the integration pattern between D365 CRM, MuleSoft, and IBS for account creation today? Walk me through the sequence: ARW Submit → AABOS Review → Auto-Process → CIS record creation → account number assignment → write-back to D365."

Batch 4 — Middesk specifically (since it's a hard external dependency)
Middesk is the single named external API and the agent must integrate with it directly.

"What is the current Middesk integration status at WAB? Is the API live, in pilot, or planned? Which projects use it today?"

"What Middesk API endpoints are licensed and accessible? Specifically: Business Search, Business Verification, SOS lookup, Officers, Beneficial Owners, Tax IDs, Watchlist. Is there a sandbox environment?"

"What is the agreed Middesk response format and field set used by the HOA AI New Account Review bot? Show me the BOT-to-MidDesk request/response sample if it exists."

"What states does Middesk currently cover for SOS lookup, and what is WAB's policy when a state is unsupported or returns a degraded response (e.g., yellow light)?"

Batch 5 — D365 CRM surface, IntraFi branch, and ConnectLive intake
The remaining named technologies — D365 case form, IntraFi, ConnectLive — round out the agent's environment.

"What D365 CRM solution / managed solution owns the AAB New Accounts Request, the HOA case, and the related entities (PMC, Child Company, Control Prong, Accounts Requested)? Who owns the solution, and what is the deployment model (cloud, on-prem, hybrid)?"

"What Power Platform / D365 components are licensed for the HOA Operations tenant: AI Builder, Copilot Studio, Power Automate premium connectors, Dataverse, custom Power Apps?"

"What is the IntraFi / IDS integration architecture? Specifically, for the Step 4 branch where AABOS hands off to Manager Review and then to Submit-to-IDS — is the IDS path manual or API-driven today? Where does it diverge from the WAB Account path?"

"What is the ConnectLive / Connect Portal new-account intake design? The Account Automation Overview deck shows this path as TBD — has any further design or scoping been done? Who owns Connect Portal?"

"What is the relationship between the email-to-D365 case-creation flow and the Outlook plug-in / Banker desktop? Is there an existing tracked-email mechanism, or do bankers manually convert each email into a case?"

How to use this list
A few notes on running these:

Batch 1 is non-negotiable as the first session. Until we know whether the UiPath bot is shipped/paused/dead, everything else is hypothetical. If the bot is in production and WAB is happy with it, our value-add is the upstream parts (banker-side review, ARW pre-fill, clarification email drafting) and the downstream parts (post-AABOS workflow), not the bot's core. If the bot is paused or struggling, Zenon's value-add is replacing it with an LLM-grounded design.

Volume discrepancy is worth flagging in Batch 1. Slide 17 says ~430 new HOA cases/month. The HOA PDD says ~20-30 accounts/month. Either the PDD is scoped narrower (maybe only specific entity types or only the SOS-validation step), or the deck and the PDD disagree. Resolving this changes the impact sizing materially.

For each Rovo answer, please share the screenshot + the source page title. I'll need page titles to cite in the final architecture doc the way UC1's walkthrough cites sources.

Don't run all 24 in one Rovo session. Rovo's response quality degrades when context gets crowded. One batch per session, sharing screenshots, then I'll synthesize between batches.

If a Rovo answer surfaces a page we should read in full, ask Rovo to summarize that specific page (the way you did with "Summarize HOA AI New Account Review PDD and relevant pages" — that's exactly the right pattern).

Want me to refine any of these questions before you run Batch 1, or shall I sit tight and wait for the results from the first batch?
