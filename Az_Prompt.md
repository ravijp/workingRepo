I need to document Azure infrastructure and access that was provisioned by hand.
There is no infrastructure-as-code and no environment definitions in this repo, so
read the LIVE Azure state, not files.

Run read-only `az` commands only — `show`, `list`, `get`. Nothing that creates,
modifies, or deletes. Never run `az login` with different credentials to widen
access. If a command fails, record the exact command and its error and move on.

## The three principals to check

I have three identities and I need the full picture for each. They are discovered
differently, so do not use one method for all three:

1. **A service principal** (app registration / deploy identity). Resolve it by
   appId or object id: `az ad sp show --id <appId-or-objectId>`. Capture its
   appId, object id, display name, and whether it has federated credentials
   (`az ad app federated-credential list --id <appId>`) or relies on a client
   secret (`az ad app credential list --id <appId>` — **report only the count,
   expiry dates, and key ids; never a secret value**).
2. **A user credential** (a human Entra account). Resolve by UPN:
   `az ad user show --id <upn>` to get the object id.
3. **A resource identity** (managed identity — user-assigned or system-assigned).
   For user-assigned: `az identity list -o table` then `az identity show`, which
   gives both clientId and principalId. For system-assigned: read it off the
   host resource, e.g. `az functionapp identity show`, `az webapp identity show`.
   **The principalId is the one role assignments use** — not the clientId. Do not
   confuse them.

**Ask me for any of the three identifiers you cannot discover.** Do not guess an
appId, UPN, or principal id, and do not proceed with a placeholder.

## Known Azure CLI trap — read this before enumerating roles

Azure CLI 2.83.x–2.86.x has an MSAL cache regression (azure-cli issue 32853) that
breaks Graph-backed calls even after a successful login. `az role assignment list
--assignee <UPN>` and `az ad user show` may fail or silently return nothing. Check
the version first with `az version`.

If Graph calls fail, use this workaround instead of retrying:

- Enumerate role assignments **by scope**, not by assignee:
  `az role assignment list --scope <scope> --include-inherited --query "[].{role:roleDefinitionName, principalId:principalId, principalType:principalType, scope:scope}"`
- Then **match `principalId` against the three object ids you resolved above**,
  rather than resolving names through Graph.
- Prefer typed `az` commands. Avoid `az rest`, which hits the broken generic
  token path.

Report which method you used, because it changes how much confidence to place in a
"no assignments found" result. A Graph failure is not evidence of an absent grant.

## What to collect

1. **Context** — `az account show`: subscription name, id, tenant id, and the
   signed-in identity. State plainly which identity the audit ran as, since that
   bounds everything below.
2. **Resource groups and regions** — `az group list -o table`.
3. **Every resource** — `az resource list -o table`: name, type, region, and SKU
   where available.
4. **Role assignments — the core of this audit.** For each of the three
   principals, every grant: role name, role definition id, full scope string, and
   scope level (subscription / resource group / individual resource). Also run a
   scope-wide sweep (`az role assignment list -g <rg> --include-inherited`) to
   catch grants held by principals I did not name — including any principal that
   cannot be resolved to a name.
5. **Authentication posture** — federated credential vs client secret for the
   service principal (count and expiry only). Any Key Vault in play. Whether
   local/key auth is disabled anywhere: `disableLocalAuth`,
   `publicNetworkAccess`, Entra-only SQL auth.
6. **Monitoring** — `az monitor log-analytics workspace show` (retention, sku, and
   `features.enableLogAccessUsingOnlyResourcePermissions`);
   `az monitor app-insights component show` (`WorkspaceResourceId`,
   `IngestionMode`, `DisableLocalAuth`, retention); `az monitor
   diagnostic-settings list --resource <id>` on the compute and AI resources;
   function app settings via `az functionapp config appsettings list`
   (**setting NAMES only, never values**) and `az functionapp identity show`.
7. **Networking, read-only** — `az network vnet list`, `az network nsg list`,
   `az network private-endpoint list`, plus whether `publicNetworkAccess` is
   disabled on the AI, data, and storage resources. Networking is controlled by
   another team, so expect gaps and mark them rather than inferring.

## Evidence rules

- **Report only what a command returned.** Never infer a value from an Azure
  default. This is the most important rule in this prompt.
- Mark anything unreadable as `NO ACCESS`, with the exact command that failed.
  Distinguish it from `NOT CONFIGURED` — permission denied and genuinely absent
  are different findings, and conflating them would mislead me.
- **Never print secret values, connection strings, or keys.** Setting names,
  credential counts, and expiry dates only.
- For managed identities, always state whether an id is a clientId or a
  principalId.

## Output

Write one self-contained `infra-live-audit.html`: inline CSS only, no external
stylesheets, scripts, fonts, or images. It will be screenshotted, so make it
light-background, print-friendly, readable at 100% zoom, with no horizontal
scroll on the page body — wide tables scroll inside their own container.

- Open with a summary table: Subscription / Tenant / Audited-as identity / Resource
  groups / Region(s) / Resource count / Role-assignment count / CLI method used
  (direct or scope-sweep workaround).
- Then a **principal identity card** for each of the three: kind, display name,
  object id, clientId/appId where applicable, credential type, and a count of its
  grants.
- Then the **role assignment table** — the centrepiece. One row per grant:
  Principal / Principal type / Role / Role definition id / Scope / Scope level /
  Which of my three principals it belongs to (or `UNRESOLVED`).
- Then one table per remaining area above.
- Badge styles for `NO ACCESS`, `NOT CONFIGURED`, and `UNRESOLVED` — distinct
  backgrounds **plus a text label**, never color alone.
- End with **"Observations"**: grants at subscription scope; use of Owner,
  Contributor, User Access Administrator, or Role Based Access Control
  Administrator; a service principal holding a client secret rather than a
  federated credential; secrets nearing expiry; public network access left
  enabled; telemetry sampling state; retention inconsistencies between a workspace
  and its Application Insights component. Each observation cites the command that
  evidenced it. Observations, not advice.

Tell me the file path when done.
