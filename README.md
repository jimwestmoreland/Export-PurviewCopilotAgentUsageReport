# Export Purview Copilot Agent Usage Report

A PowerShell reporting tool that queries Microsoft Purview audit logs for Microsoft 365 Copilot agent usage and checks each user's current Microsoft 365 Copilot license assignment through Microsoft Graph.

The script runs directly against Microsoft services. It does not require a Microsoft 365 admin center usage-report export.

## What it does

- Searches Microsoft Purview for `CopilotInteraction` audit records.
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

### Microsoft Purview

Microsoft Purview Audit must be enabled for the tenant. Audit Standard is enabled by default for supported Microsoft 365 subscriptions.

You can ask the script to verify the setting by using `-CheckPurviewAuditEnabled`, or check it manually:

```powershell
Connect-ExchangeOnline
Get-AdminAuditLogConfig |
    Format-List UnifiedAuditLogIngestionEnabled
```

`UnifiedAuditLogIngestionEnabled` must be `True`.

### Assign the Purview audit role group

For least-privileged access to search and export audit data in the Microsoft Purview portal, assign the account to **Audit Reader**:

1. Sign in to the [Microsoft Purview portal](https://purview.microsoft.com).
2. Go to **Settings** > **Roles and scopes** > **Role groups**.
3. Select **Audit Reader**.
4. Select **Edit** or **Manage members**.
5. Add the account that will run the script.
6. Save the role-group membership.

The built-in audit role groups provide these capabilities:

| Purview role group | Included audit role | Capability |
|---|---|---|
| **Audit Reader** | **View-Only Audit Logs** | Search and export audit records. Recommended for this script. |
| **Audit Manager** | **View-Only Audit Logs** and **Audit Logs** | Search and export audit records and manage audit settings, including enabling or disabling auditing. |

Use **Audit Manager** only when the account must manage audit configuration. The report itself only needs read access.

### Assign the Exchange Online audit role

This script calls the Exchange Online `Search-UnifiedAuditLog` cmdlet. Microsoft currently requires Exchange admin center permissions for audit cmdlets, even when the user already has a Purview audit role group.

Assign the **View-Only Audit Logs** role through the Exchange admin center:

1. Sign in to the [Exchange admin center](https://admin.exchange.microsoft.com).
2. Go to **Roles** > **Admin roles**.
3. Select an appropriate existing role group or create a dedicated role group for this report.
4. In the role group's **Assigned** roles, add **View-Only Audit Logs**.
5. Add the account that will run the script as a member.
6. Save the role group.

For least privilege, create a dedicated role group containing only **View-Only Audit Logs** instead of adding the account to a broad administrative group such as Organization Management.

The administrator assigning Exchange role groups must have permission to manage roles, typically through the **Role Management** role. Permission changes can take time to propagate. After assignment, reconnect Exchange Online before testing.

Verify that the account can execute the audit search:

```powershell
Connect-ExchangeOnline

Search-UnifiedAuditLog `
    -StartDate (Get-Date).AddDays(-1) `
    -EndDate (Get-Date) `
    -RecordType CopilotInteraction `
    -ResultSize 1
```

The command is authorized if it runs without a permission or cmdlet-availability error. A zero-row result does not indicate a permissions failure; it can mean no matching activity occurred during the period.

Audit Standard normally retains audit records for 180 days. Older data requires applicable Audit Premium licensing and retention policies.

### Microsoft Graph

The script uses delegated Microsoft Graph access with these scopes:

- `User.Read.All`
- `Organization.Read.All`

Administrator consent is required for these delegated permissions. On first use, `Connect-MgGraph` requests the scopes interactively. If users aren't allowed to consent, an administrator must grant tenant-wide consent to the Microsoft Graph PowerShell enterprise application:

1. Attempt the script or run the following once so the Microsoft Graph PowerShell enterprise application is provisioned and the requested scopes are recorded:

   ```powershell
   Connect-MgGraph -Scopes "User.Read.All","Organization.Read.All"
   ```

2. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com) using an account permitted to grant tenant-wide consent.
3. Go to **Entra ID** > **Enterprise apps** > **All applications**.
4. Select the Microsoft Graph PowerShell enterprise application used by `Connect-MgGraph`.
5. Go to **Security** > **Permissions**.
6. Review the requested delegated permissions.
7. Select **Grant admin consent**.

Microsoft documents **Cloud Application Administrator**, **Application Administrator**, **AI Administrator**, or **Privileged Role Administrator** as roles that can grant these delegated permissions, subject to tenant policy. Use the least-privileged role approved by your organization.

These Graph permissions allow the script to read users, assigned licenses, and subscribed SKU information. They do not grant permission to modify users or licenses.

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
|---|---:|---|
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
|---|---|
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
