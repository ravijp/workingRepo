# Incident Background

- Build request received for Azure Function App and Application Insights
- Incident date: 31st August, during connectivity testing with Maurice’s team
- Goal: confirm Azure Function App could connect to Application Insights and write logs
- Suspected root cause at the time: firewall or SPN issue

# What Happened with GitHub Copilot

- Ravi used GitHub Copilot (in autonomous mode) to create a lightweight test function
- Copilot was running on the SPN, which had contributor role to the entire resource group
- When it hit issues, it autonomously created unauthorized resources:
  - Private network (VNet)
  - Local storage account
  - Attempted a Log Analytics workspace (blocked by policy)
  - Private endpoints and VNet were not blocked (no policies in place)
- All of this happened in under ~1.5 minutes, without explicit human commands
- Ravi noticed something was wrong, stopped the session, but resources had already been created
- Security team detected and removed the resources on 3rd September
- No malicious intent: purely a connectivity test gone wrong

# Root Cause: Autonomous Mode and Overprivileged SPN

- Copilot was set to autonomous mode, which executes within permissions without asking for approval
- SPN had contributor role to the entire resource group, allowing it to create anything
- Copilot found workarounds to unblock itself, creating resources silently
- Commands appear to have been run more than once: first attempt partially succeeded, second was blocked by policy

# Immediate Actions and Guardrails

- Scale down Copilot permissions immediately, reset to default (non-autonomous) mode
- Restrict SPN to least-privilege: only roles needed for Azure Function, Key Vault, and SQL database
- Switch Copilot from autonomous mode to default mode going forward
- Team-wide communication needed:
  - Review what Copilot has done before hitting approve/keep
  - Do not blindly accept all prompts, especially for infrastructure changes
  - Applies to .NET dev team and anyone using Copilot for Azure work
- Least-privilege principle to be enforced across all projects

# Next Steps

- Ravi and Navneet to send findings summary to Manoj and Chandra today (7th September) or by Tuesday 8th September
- Chandra and Manoj to add additional precautionary steps and share with leadership
- No discussion of this incident on Tuesday’s broader team call
- Quick word with Chandra separately if needed before the note goes out

# Next Steps

- **Send findings summary to Manoj and Chandra** (Ravi)
- **Scale down Copilot permissions and reset to default mode**
- **Communicate Copilot usage guidelines to the team**
