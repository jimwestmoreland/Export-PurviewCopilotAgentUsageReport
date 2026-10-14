# Export Purview Copilot Agent Usage Report

A PowerShell reporting tool that queries Microsoft Purview audit logs for Microsoft 365 Copilot agent usage and checks each user's current Microsoft 365 Copilot license assignment through Microsoft Graph.

The script runs directly against Microsoft services. It does not require a Microsoft 365 admin center usage-report export.

## What it does

- Searches Microsoft Purview for `CopilotInteraction` and `ConnectedAIAppInteraction` audit records.
- Filters interactions by a stable agent ID or exact agent name.
- Aggregates usage by user.
- Reports interaction totals and first and last activity timestamps.
- Reports the Copilot host applications used, such as Copilot Chat or Word.
- Queries Microsoft Graph for each user's current account and license status.
- Detects and rejects a mismatch between the Purview and Microsoft Graph tenants.
- Exports the results to a timestamped CSV file.
- Reports audit records that cannot be attributed to an agent.
- Processes audit records as a stream to support larger tenants and date ranges.

## Script

[`Export-PurviewCopilotAgentUsageReport.ps1`](./Export-PurviewCopilotAgentUsageReport.ps1)

## Requirements

### PowerShell

- Windows PowerShell 5.1 or PowerShell 7

### Modules

Install the required modules for the current user:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
```

### Required configuration and access

Microsoft Purview Audit must be enabled for the tenant. Use `-CheckPurviewAuditEnabled` to have the script verify this setting. Audit Standard normally retains records for 180 days; older data requires applicable Audit Premium licensing and retention policies.

Assign the following least-privileged access to the account that runs the script:

| Service | Role or permission | Why it is required |
| --- | --- | --- |
| Microsoft Purview | **Audit Reader** role group | Allows read-only access to search and export audit records. |
| Exchange Online | **View-Only Audit Logs** role | Allows the script to run `Search-UnifiedAuditLog`. |
| Microsoft Graph | `User.Read.All` delegated permission | Reads each user's account status and assigned licenses. |
| Microsoft Graph | `LicenseAssignment.Read.All` delegated permission | Reads the tenant's subscribed license SKUs. |

Administrator consent is required for the Microsoft Graph delegated permissions. The script does not modify users, licenses, or audit settings.

## Usage

Run the script from PowerShell:

```powershell
.\Export-PurviewCopilotAgentUsageReport.ps1 `
    -AgentName "Contoso HR Agent" `
    -StartDate 2026-10-07 `
    -EndDate 2026-10-13 `
    -CheckPurviewAuditEnabled
```

Dates are inclusive UTC calendar dates.

### Use the agent ID when available

Agent IDs are more reliable than names because names can change and multiple agents can use the same display name.

```powershell
.\Export-PurviewCopilotAgentUsageReport.ps1 `
    -AgentName "Contoso HR Agent" `
    -AgentId "CopilotStudio.Declarative.00000000-0000-0000-0000-000000000000" `
    -StartDate 2026-09-14 `
    -EndDate 2026-10-13 `
    -TenantId "contoso.onmicrosoft.com" `
    -ExchangeOrganization "contoso.onmicrosoft.com" `
    -OutputPath .\Contoso-HR-Agent-License-Report.csv
```

When `-AgentId` is supplied, it is the match key. `-AgentName` is used as a human-readable request label and to construct the default output filename.

## Parameters

| Parameter | Required | Description |
| --- | ---: | --- |
| `AgentName` | Yes | Exact agent name. Matching is case-insensitive when no agent ID is supplied. |
| `AgentId` | No | Exact, stable agent ID. Becomes the match key when supplied. |
| `StartDate` | Yes | First UTC calendar date to search, inclusive. |
| `EndDate` | Yes | Last UTC calendar date to search, inclusive. Cannot be in the future. |
| `AuditSliceMinutes` | No | Purview search interval size. Default: `360`. Valid range: `5`–`1440`. |
| `UseHighCompleteness` | No | Enables the preview Purview high-completeness search mode. |
| `CheckPurviewAuditEnabled` | No | Verifies `UnifiedAuditLogIngestionEnabled` before searching. |
| `OutputPath` | No | Output CSV path. Defaults to the current directory with a timestamped filename. |
| `TenantId` | No | Microsoft Entra tenant ID or verified domain for Microsoft Graph sign-in. |
| `ExchangeOrganization` | No | Primary `onmicrosoft.com` domain for Exchange Online sign-in. |

## Output

The script writes one row per unique user who interacted with the selected agent.

| Column | Description |
| --- | --- |
| `ReportStartDateUtc` | First date searched. |
| `ReportEndDateUtc` | Last date searched. |
| `ReportPeriodDays` | Number of inclusive calendar days searched. |
| `AgentName` | Agent name observed in Purview. |
| `AgentIds` | Agent IDs observed in matching records. |
| `AgentVersions` | Agent versions observed for the user. |
| `UserPrincipalName` | User principal name returned by Microsoft Graph or the audit record when lookup fails. |
| `DisplayName` | User display name from Microsoft Graph. |
| `AccountEnabled` | Current Microsoft Entra account status. |
| `M365CopilotLicenseAssigned` | Whether the Microsoft 365 Copilot SKU is currently assigned. |
| `M365CopilotSkuPartNumbers` | Matching Copilot SKU part numbers. |
| `CopilotInteractionCount` | Number of matching, unique Purview audit records. |
| `FirstInteractionUtc` | First matching interaction timestamp. |
| `LastInteractionUtc` | Last matching interaction timestamp. |
| `AppHosts` | Copilot host applications observed for the user. |
| `LicenseCheckedAtUtc` | Time the Microsoft Graph license lookup ran. |
| `LicenseLookupStatus` | `Success` or `Failed`. |
| `LicenseLookupError` | Explicit Microsoft Graph lookup error, when applicable. |

The default filename follows this pattern:

```text
CopilotAgentUsers_<AgentName>_<yyyyMMdd_HHmmss>.csv
```

## Search volume and paging

`Search-UnifiedAuditLog` returns at most 50,000 records per search session. The script divides the requested date range into smaller search intervals.

The default interval is six hours:

```powershell
-AuditSliceMinutes 360
```

If an interval reaches 50,000 records, the script stops rather than silently returning incomplete results. Rerun it with a smaller interval:

```powershell
-AuditSliceMinutes 60
```

For continuous collection at scale, consider the Office 365 Management Activity API with an `Audit.General` subscription and the `ActivityFeed.Read` application permission.

## High-completeness searches

Microsoft documents the `HighCompleteness` option as a preview feature that is unavailable in some organizations and subject to change. It is therefore opt-in:

```powershell
.\Export-PurviewCopilotAgentUsageReport.ps1 `
    -AgentName "Contoso HR Agent" `
    -StartDate 2026-10-07 `
    -EndDate 2026-10-13 `
    -UseHighCompleteness
```

## Important limitations

- Purview documents `AgentId` and `AgentName`, but these fields might not be populated for every historical event, agent type, or workload.
- Records without either field cannot be attributed to a specific agent. The script reports their count.
- License status is checked when the script runs, not historically as of the interaction date.
- Interaction count is based on unique matching audit records. It is not a Copilot Credits or cost measurement.
- Purview audit records do not expose monetary cost or Copilot Credits consumed.
- Actual metered cost must come from the Microsoft Copilot Credits report, Copilot Studio consumption analytics, or Azure Cost Management.
- The script uses interactive authentication and is not configured for unattended execution.

## Troubleshooting

### No matching records

- Confirm the date range is within your audit retention period.
- Prefer `-AgentId` when the ID is known.
- Confirm that Purview records contain `AgentId` or `AgentName`.
- Allow time for new audit records to become available.

### More than one agent ID uses the name

Run the script again with `-AgentId` to select the intended agent.

### Search interval reached 50,000 records

Reduce `-AuditSliceMinutes`.

### Tenant mismatch

The Purview audit records and Microsoft Graph session belong to different tenants. Reconnect both services to the same tenant and use `-TenantId` and `-ExchangeOrganization` to make the intended tenant explicit.

### User license lookup failed

Review `LicenseLookupStatus` and `LicenseLookupError` in the output. Confirm:

- Microsoft Graph admin consent was granted.
- The signed-in account has the required scopes.
- The user still exists in the tenant.

## Microsoft documentation

- [Audit logs for Copilot and AI applications](https://learn.microsoft.com/purview/audit-copilot)
- [Get started with Microsoft Purview auditing](https://learn.microsoft.com/purview/audit-get-started)
- [Permissions in the Microsoft Purview portal](https://learn.microsoft.com/purview/purview-permissions)
- [Manage role groups in Exchange Online](https://learn.microsoft.com/exchange/permissions-exo/role-groups)
- [Search-UnifiedAuditLog](https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog)
- [Office 365 Management Activity API](https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference)
- [List subscribed SKUs in Microsoft Graph](https://learn.microsoft.com/graph/api/subscribedsku-list)
- [Grant tenant-wide admin consent](https://learn.microsoft.com/entra/identity/enterprise-apps/grant-admin-consent)
- [Microsoft Copilot Credits report](https://learn.microsoft.com/microsoft-365/admin/activity-reports/microsoft-365-copilot-credits)

## Security and privacy

The output contains user identifiers and usage information. Store and share generated CSV files according to your organization's privacy, security, retention, and compliance requirements.
