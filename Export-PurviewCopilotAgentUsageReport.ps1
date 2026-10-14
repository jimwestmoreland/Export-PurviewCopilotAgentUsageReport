<#
.SYNOPSIS
Reports users who interacted with a specific Microsoft 365 Copilot agent and
whether they currently have a Microsoft 365 Copilot license assigned.

.DESCRIPTION
This script queries Microsoft Purview audit data directly by using the
Search-UnifiedAuditLog Exchange Online PowerShell cmdlet. It retrieves
CopilotInteraction records for an inclusive UTC calendar-date range, filters by
agent ID when provided or otherwise by exact agent name, and aggregates the
matching interactions by user.

For every unique user, the script uses Microsoft Graph to check the user's current
Microsoft 365 Copilot license assignment. Results are written to a timestamped CSV
in the directory from which the script is run unless -OutputPath is specified.

No Microsoft 365 admin center report export is required.

.PURVIEW CONFIGURATION REQUIRED
1. Microsoft Purview Audit must be enabled for the tenant. Audit (Standard) is
   enabled by default for supported Microsoft 365 subscriptions. Use
   -CheckPurviewAuditEnabled to have the script verify this setting before searching,
   or verify it manually:

   Connect-ExchangeOnline
   Get-AdminAuditLogConfig | Format-List UnifiedAuditLogIngestionEnabled

   UnifiedAuditLogIngestionEnabled must be True. If it is False, an authorized
   administrator must enable auditing before the tenant starts collecting records.

2. The account running this script must be assigned a Purview role group that can
   search audit data. Use the least-privileged option supported by your tenant:
   - Audit Reader, or
   - Audit Manager

   In tenants using the classic permissions model, the equivalent role is
   View-Only Audit Logs. Role changes can take time to become effective.

3. CopilotInteraction records must exist within the tenant's retention window.
   Audit (Standard) normally retains audit records for 180 days. Audit (Premium)
   can provide longer retention when the required licenses and retention policies
   are configured.

4. The current Purview Copilot audit schema documents AgentId and AgentName, but
   these fields might not be populated for every historical event, agent type, or
   workload. The script reports the count of records with missing agent identity so
   remaining telemetry gaps are visible.

.PREREQUISITES
1. Windows PowerShell 5.1 or PowerShell 7.
2. Exchange Online PowerShell module:
   Install-Module ExchangeOnlineManagement -Scope CurrentUser
3. Microsoft Graph PowerShell SDK modules:
   Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
   Install-Module Microsoft.Graph.Users -Scope CurrentUser
   Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
4. Interactive sign-in is required for Exchange Online/Purview and Microsoft Graph.
5. Admin consent is required for these delegated Microsoft Graph scopes:
   - User.Read.All
   - Organization.Read.All

.ROLES AND PERMISSIONS REQUIRED
- Purview: Audit Reader is the recommended least-privileged role. Audit Manager or
  the classic View-Only Audit Logs role also provides audit-search access.
- Exchange Online PowerShell: the signed-in account must be allowed to run
  Search-UnifiedAuditLog.
- Microsoft Graph delegated permissions:
  - User.Read.All to read users and assigned licenses.
  - Organization.Read.All to read the tenant's subscribed SKUs.
- A Global Administrator is needed only to grant tenant-wide consent, enable audit
  collection when disabled, or assign the required administrative roles.

.PARAMETER AgentName
Exact agent name to report. Matching is case-insensitive.

.PARAMETER AgentId
Optional exact agent ID. Use this whenever it is known because names can change or
multiple agents can use the same display name. When supplied, AgentId is the match
key and AgentName is used only as a human-readable request label and file name.

.PARAMETER StartDate
First UTC calendar date to search. The entire date is included.

.PARAMETER EndDate
Last UTC calendar date to search. The entire date is included.

.PARAMETER AuditSliceMinutes
Size of each Purview audit-search interval. The default is 360 minutes. Each
Search-UnifiedAuditLog session is limited to 50,000 records. If a slice reaches
that limit, rerun with a smaller value.

.PARAMETER UseHighCompleteness
Uses the preview Search-UnifiedAuditLog HighCompleteness mode. Microsoft documents
this mode as unavailable in some organizations and subject to change, so it is
opt-in rather than enabled by default.

.PARAMETER CheckPurviewAuditEnabled
Checks UnifiedAuditLogIngestionEnabled after connecting to Exchange Online. The
script stops if auditing is disabled, the setting is unavailable, or the signed-in
account cannot read it. This check is optional because reading the tenant audit
configuration can require additional Exchange configuration permissions beyond
the permissions needed to search audit records.

.PARAMETER OutputPath
Optional output CSV path. By default, the script writes a timestamped CSV to the
current working directory.

.PARAMETER TenantId
Optional Microsoft Entra tenant ID or verified domain used during Graph sign-in.

.PARAMETER ExchangeOrganization
Optional primary onmicrosoft.com domain used during Exchange Online sign-in.

.EXAMPLE
.\Export-PurviewCopilotAgentUsageReport.ps1 `
    -AgentName "Contoso HR Agent" `
    -StartDate 2026-10-07 `
    -EndDate 2026-10-13 `
    -CheckPurviewAuditEnabled

.EXAMPLE
.\Export-PurviewCopilotAgentUsageReport.ps1 `
    -AgentName "Contoso HR Agent" `
    -AgentId "CopilotStudio.Declarative.00000000-0000-0000-0000-000000000000" `
    -StartDate 2026-09-14 `
    -EndDate 2026-10-13 `
    -TenantId "contoso.onmicrosoft.com" `
    -ExchangeOrganization "contoso.onmicrosoft.com" `
    -OutputPath .\Contoso-HR-Agent-License-Report.csv

.NOTES
- StartDate and EndDate are interpreted as inclusive UTC calendar dates.
- License status is checked at script runtime, not historically as of each event.
- Search-UnifiedAuditLog returns at most 50,000 records per search session. This
  script divides the requested range into smaller sessions and stops rather than
  silently accepting a potentially truncated interval.
- Use -UseHighCompleteness only when the tenant supports that preview search mode.
- The canonical Microsoft 365 Copilot SKU ID checked by this script is
  639dec6b-bb19-468b-871c-c5c441c4b0cb. The script also recognizes a subscribed
  SKU whose part number normalizes to MICROSOFT_365_COPILOT.
- For continuous collection at scale, use the Office 365 Management Activity API
  with an Audit.General subscription and ActivityFeed.Read application permission.
- Microsoft Learn:
  https://learn.microsoft.com/purview/audit-copilot
  https://learn.microsoft.com/powershell/module/exchangepowershell/search-unifiedauditlog
  https://learn.microsoft.com/office/office-365-management-api/office-365-management-activity-api-reference
  https://learn.microsoft.com/graph/api/subscribedsku-list
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$AgentName,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$AgentId,

    [Parameter(Mandatory)]
    [datetime]$StartDate,

    [Parameter(Mandatory)]
    [datetime]$EndDate,

    [Parameter()]
    [ValidateRange(5, 1440)]
    [int]$AuditSliceMinutes = 360,

    [Parameter()]
    [switch]$UseHighCompleteness,

    [Parameter()]
    [switch]$CheckPurviewAuditEnabled,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ExchangeOrganization
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-ObjectPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string[]]$Names
    )

    if ($null -eq $InputObject) {
        return $null
    }

    foreach ($name in $Names) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value) {
            return $property.Value
        }
    }

    return $null
}

function Get-CopilotAuditValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$AuditData,

        [Parameter(Mandatory)]
        [string[]]$Names
    )

    $value = Get-ObjectPropertyValue -InputObject $AuditData -Names $Names
    if ($null -ne $value) {
        return $value
    }

    foreach ($containerName in @('CopilotEventData', 'EventData')) {
        $container = Get-ObjectPropertyValue -InputObject $AuditData -Names @($containerName)
        $value = Get-ObjectPropertyValue -InputObject $container -Names $Names
        if ($null -ne $value) {
            return $value
        }
    }

    return $null
}

function Get-PurviewCopilotAuditRecords {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [datetime]$RangeStartUtc,

        [Parameter(Mandatory)]
        [datetime]$RangeEndExclusiveUtc,

        [Parameter(Mandatory)]
        [int]$SliceMinutes,

        [Parameter()]
        [switch]$HighCompleteness
    )

    $sliceStart = $RangeStartUtc

    while ($sliceStart -lt $RangeEndExclusiveUtc) {
        $sliceEnd = $sliceStart.AddMinutes($SliceMinutes)
        if ($sliceEnd -gt $RangeEndExclusiveUtc) {
            $sliceEnd = $RangeEndExclusiveUtc
        }

        $sessionId = "CopilotAgentUsage-$([guid]::NewGuid())"
        $sliceCount = 0

        Write-Verbose "Searching Purview from $($sliceStart.ToString('o')) through $($sliceEnd.ToString('o'))."

        while ($true) {
            $searchParameters = @{
                StartDate      = $sliceStart
                EndDate        = $sliceEnd
                RecordType     = 'CopilotInteraction'
                SessionId      = $sessionId
                SessionCommand = 'ReturnLargeSet'
                ResultSize     = 5000
                Formatted      = $true
            }
            if ($HighCompleteness) {
                $searchParameters.HighCompleteness = $true
            }

            $page = @(
                Search-UnifiedAuditLog @searchParameters
            )

            if ($page.Count -eq 0) {
                break
            }

            foreach ($record in $page) {
                Write-Output $record
            }
            $sliceCount += $page.Count

            if ($sliceCount -ge 50000) {
                throw "The Purview search interval beginning $($sliceStart.ToString('o')) reached the 50,000-record Search-UnifiedAuditLog session limit. Results could be truncated. Rerun with -AuditSliceMinutes set to a value smaller than $SliceMinutes."
            }
        }

        $sliceStart = $sliceEnd
    }
}

$startDateUtc = [datetime]::SpecifyKind($StartDate.Date, [DateTimeKind]::Utc)
$endDateUtc = [datetime]::SpecifyKind($EndDate.Date, [DateTimeKind]::Utc)
if ($endDateUtc -lt $startDateUtc) {
    throw 'EndDate must be on or after StartDate.'
}
if ($endDateUtc -gt [datetime]::UtcNow.Date) {
    throw 'EndDate cannot be later than the current UTC date.'
}
$endDateExclusiveUtc = $endDateUtc.AddDays(1)
$reportPeriodDays = ($endDateUtc - $startDateUtc).Days + 1

if ($startDateUtc -lt [datetime]::UtcNow.Date.AddDays(-180)) {
    Write-Warning 'StartDate is more than 180 days ago. Audit (Standard) normally retains records for 180 days; older records require applicable Audit (Premium) licensing and retention policies.'
}

$requiredModules = @(
    'ExchangeOnlineManagement'
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Users'
    'Microsoft.Graph.Identity.DirectoryManagement'
)

$missingModules = @(
    $requiredModules | Where-Object { -not (Get-Module -ListAvailable -Name $_) }
)
if ($missingModules.Count -gt 0) {
    $installCommands = $missingModules | ForEach-Object {
        "Install-Module $_ -Scope CurrentUser"
    }
    throw "Required module(s) are not installed: $($missingModules -join ', '). Install them with:`n$($installCommands -join [Environment]::NewLine)"
}

foreach ($moduleName in $requiredModules) {
    Import-Module $moduleName -ErrorAction Stop
}

$exchangeConnectParameters = @{
    ShowBanner = $false
}
if (-not [string]::IsNullOrWhiteSpace($ExchangeOrganization)) {
    $exchangeConnectParameters.Organization = $ExchangeOrganization
}
Connect-ExchangeOnline @exchangeConnectParameters

if ($CheckPurviewAuditEnabled) {
    if ($null -eq (Get-Command -Name Get-AdminAuditLogConfig -ErrorAction SilentlyContinue)) {
        throw 'Get-AdminAuditLogConfig is unavailable after connecting to Exchange Online. The Purview audit-enabled check could not be performed.'
    }

    try {
        $auditConfiguration = Get-AdminAuditLogConfig -ErrorAction Stop
    }
    catch {
        throw "Unable to read the Purview audit configuration. Verify that the signed-in account can run Get-AdminAuditLogConfig. $($_.Exception.Message)"
    }

    $auditEnabledProperty = $auditConfiguration.PSObject.Properties['UnifiedAuditLogIngestionEnabled']
    if ($null -eq $auditEnabledProperty) {
        throw 'Get-AdminAuditLogConfig did not return UnifiedAuditLogIngestionEnabled. The Purview audit-enabled check could not be completed.'
    }
    if (-not [bool]$auditEnabledProperty.Value) {
        throw 'Microsoft Purview auditing is disabled for this tenant (UnifiedAuditLogIngestionEnabled is False). Enable auditing and allow time for audit collection before running this report.'
    }

    Write-Host 'Microsoft Purview auditing is enabled (UnifiedAuditLogIngestionEnabled is True).'
}

if ($null -eq (Get-Command -Name Search-UnifiedAuditLog -ErrorAction SilentlyContinue)) {
    throw 'Search-UnifiedAuditLog is unavailable after connecting to Exchange Online. Verify the ExchangeOnlineManagement module and the account audit-search permissions.'
}

$matchingRecordIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$observedAgents = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$selectedAgentIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$selectedAgentNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$auditOrganizationIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$userAggregates = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
$matchingEventCount = 0
$invalidAuditRecordCount = 0
$missingAgentIdentityCount = 0

Get-PurviewCopilotAuditRecords `
    -RangeStartUtc $startDateUtc `
    -RangeEndExclusiveUtc $endDateExclusiveUtc `
    -SliceMinutes $AuditSliceMinutes `
    -HighCompleteness:$UseHighCompleteness |
ForEach-Object {
    $record = $_
    try {
        $auditData = [string]$record.AuditData | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        $invalidAuditRecordCount++
        Write-Warning "Unable to parse Purview audit record '$($record.Identity)': $($_.Exception.Message)"
        return
    }

    $eventAgentId = [string](Get-CopilotAuditValue -AuditData $auditData -Names @('AgentId', 'AgentID'))
    $eventAgentName = [string](Get-CopilotAuditValue -AuditData $auditData -Names @('AgentName'))
    if ([string]::IsNullOrWhiteSpace($eventAgentId) -and [string]::IsNullOrWhiteSpace($eventAgentName)) {
        $missingAgentIdentityCount++
        return
    }

    if ($observedAgents.Count -lt 50) {
        $observedAgents.Add("$eventAgentName [$eventAgentId]") | Out-Null
    }

    $isAgentMatch = if (-not [string]::IsNullOrWhiteSpace($AgentId)) {
        $eventAgentId -ieq $AgentId
    }
    else {
        $eventAgentName -ieq $AgentName
    }
    if (-not $isAgentMatch) {
        return
    }

    $recordId = [string](Get-CopilotAuditValue -AuditData $auditData -Names @('Id', 'RecordId'))
    if ([string]::IsNullOrWhiteSpace($recordId)) {
        $recordId = [string]$record.Identity
    }
    if ([string]::IsNullOrWhiteSpace($recordId)) {
        $recordId = '{0}|{1}|{2}' -f $record.CreationDate, $record.UserIds, $record.AuditData
    }
    if (-not $matchingRecordIds.Add($recordId)) {
        return
    }

    $eventUserId = [string](Get-CopilotAuditValue -AuditData $auditData -Names @('UserId', 'UserID'))
    if ([string]::IsNullOrWhiteSpace($eventUserId)) {
        $eventUserId = @($record.UserIds) |
            ForEach-Object { [string]$_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -First 1
    }
    if ([string]::IsNullOrWhiteSpace($eventUserId)) {
        Write-Warning "Purview audit record '$recordId' matched the Copilot schema but did not contain a user identifier."
        return
    }

    $creationValue = Get-CopilotAuditValue -AuditData $auditData -Names @('CreationTime', 'CreationDate')
    $creationDateUtc = [datetime]::MinValue
    if (-not [datetime]::TryParse(
        [string]$creationValue,
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal,
        [ref]$creationDateUtc
    )) {
        if (-not [datetime]::TryParse(
            [string]$record.CreationDate,
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal,
            [ref]$creationDateUtc
        )) {
            Write-Warning "Purview audit record '$recordId' did not contain a valid creation timestamp and was excluded."
            return
        }
    }

    $appHost = [string](Get-CopilotAuditValue -AuditData $auditData -Names @('AppHost'))
    $agentVersion = [string](Get-CopilotAuditValue -AuditData $auditData -Names @('AgentVersion'))
    $organizationId = [string](Get-CopilotAuditValue -AuditData $auditData -Names @('OrganizationId'))

    if (-not [string]::IsNullOrWhiteSpace($eventAgentId)) {
        $selectedAgentIds.Add($eventAgentId) | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($eventAgentName)) {
        $selectedAgentNames.Add($eventAgentName) | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($organizationId)) {
        $auditOrganizationIds.Add($organizationId) | Out-Null
    }

    if (-not $userAggregates.ContainsKey($eventUserId)) {
        $userAggregates[$eventUserId] = [pscustomobject]@{
            UserId          = $eventUserId
            InteractionCount = 0
            FirstActivityUtc = $creationDateUtc
            LastActivityUtc = $creationDateUtc
            AppHosts        = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            AgentVersions   = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        }
    }

    $userAggregate = $userAggregates[$eventUserId]
    $userAggregate.InteractionCount++
    if ($creationDateUtc -lt $userAggregate.FirstActivityUtc) {
        $userAggregate.FirstActivityUtc = $creationDateUtc
    }
    if ($creationDateUtc -gt $userAggregate.LastActivityUtc) {
        $userAggregate.LastActivityUtc = $creationDateUtc
    }
    if (-not [string]::IsNullOrWhiteSpace($appHost)) {
        $userAggregate.AppHosts.Add($appHost) | Out-Null
    }
    if (-not [string]::IsNullOrWhiteSpace($agentVersion)) {
        $userAggregate.AgentVersions.Add($agentVersion) | Out-Null
    }

    $matchingEventCount++
}

if ($invalidAuditRecordCount -gt 0) {
    Write-Warning "$invalidAuditRecordCount Purview record(s) contained invalid JSON and were excluded."
}
if ($missingAgentIdentityCount -gt 0) {
    Write-Warning "$missingAgentIdentityCount CopilotInteraction record(s) did not contain AgentId or AgentName and could not be attributed to an agent."
}

if ($matchingEventCount -eq 0) {
    $matchDescription = if ($AgentId) { "agent ID '$AgentId'" } else { "agent name '$AgentName'" }
    throw "No Purview CopilotInteraction records matched $matchDescription between $($startDateUtc.ToString('yyyy-MM-dd')) and $($endDateUtc.ToString('yyyy-MM-dd')) UTC. Observed agents (up to 50): $(@($observedAgents) -join ', '). Records without agent identity: $missingAgentIdentityCount."
}

if ([string]::IsNullOrWhiteSpace($AgentId) -and $selectedAgentIds.Count -gt 1) {
    throw "More than one agent ID uses the name '$AgentName': $(@($selectedAgentIds) -join ', '). Run the script again with -AgentId to select one agent."
}
if ($auditOrganizationIds.Count -gt 1) {
    throw "Matching Purview records contained more than one OrganizationId: $(@($auditOrganizationIds) -join ', '). The tenant cannot be determined safely."
}

$requiredScopes = @('User.Read.All', 'Organization.Read.All')
$graphContext = Get-MgContext
$hasRequiredScopes = $null -ne $graphContext -and @(
    $requiredScopes | Where-Object { $_ -notin $graphContext.Scopes }
).Count -eq 0
$tenantMatches = [string]::IsNullOrWhiteSpace($TenantId) -or
    ($null -ne $graphContext -and $graphContext.TenantId -ieq $TenantId)

if (-not $hasRequiredScopes -or -not $tenantMatches) {
    $graphConnectParameters = @{
        Scopes    = $requiredScopes
        NoWelcome = $true
    }
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $graphConnectParameters.TenantId = $TenantId
    }
    Connect-MgGraph @graphConnectParameters
    $graphContext = Get-MgContext
}

if ($null -eq $graphContext) {
    throw 'Microsoft Graph authentication did not produce an active context.'
}
if ($auditOrganizationIds.Count -eq 0) {
    Write-Warning 'Matching Purview records did not contain OrganizationId. The script could not verify that the Purview and Microsoft Graph sessions use the same tenant.'
}
else {
    $auditOrganizationId = @($auditOrganizationIds)[0]
    if ([string]::IsNullOrWhiteSpace([string]$graphContext.TenantId) -or
        [string]$graphContext.TenantId -ine $auditOrganizationId) {
        throw "Tenant mismatch: matching Purview records belong to tenant '$auditOrganizationId', but Microsoft Graph is connected to tenant '$($graphContext.TenantId)'. Reconnect both services to the same tenant."
    }
}

$canonicalCopilotSkuId = [guid]'639dec6b-bb19-468b-871c-c5c441c4b0cb'
$subscribedSkus = @(Get-MgSubscribedSku -All -Property @('SkuId', 'SkuPartNumber'))
$copilotSkus = @(
    $subscribedSkus | Where-Object {
        $_.SkuId -eq $canonicalCopilotSkuId -or
        ([string]$_.SkuPartNumber -replace '[^A-Za-z0-9]', '_').Trim('_') -ieq 'MICROSOFT_365_COPILOT'
    }
)
if ($copilotSkus.Count -eq 0) {
    Write-Warning 'No Microsoft 365 Copilot subscribed SKU was found in the tenant. Users will be reported as not licensed.'
}

$selectedAgentIdValues = @($selectedAgentIds | Sort-Object)
$selectedAgentNameValues = @($selectedAgentNames | Sort-Object)
$results = @(
    foreach ($userAggregate in $userAggregates.Values) {
        $userId = [string]$userAggregate.UserId
        $appHosts = @($userAggregate.AppHosts | Sort-Object)
        $agentVersions = @($userAggregate.AgentVersions | Sort-Object)

        try {
            $graphUser = Get-MgUser `
                -UserId $userId `
                -Property @('Id', 'UserPrincipalName', 'DisplayName', 'AccountEnabled', 'AssignedLicenses') `
                -ErrorAction Stop
            $assignedSkuIds = @($graphUser.AssignedLicenses | ForEach-Object { [guid]$_.SkuId })
            $assignedCopilotSkus = @(
                $copilotSkus | Where-Object { [guid]$_.SkuId -in $assignedSkuIds }
            )

            [pscustomobject][ordered]@{
                ReportStartDateUtc             = $startDateUtc.ToString('yyyy-MM-dd')
                ReportEndDateUtc               = $endDateUtc.ToString('yyyy-MM-dd')
                ReportPeriodDays               = $reportPeriodDays
                AgentName                      = ($selectedAgentNameValues -join ';')
                AgentIds                       = ($selectedAgentIdValues -join ';')
                AgentVersions                  = ($agentVersions -join ';')
                UserPrincipalName              = $graphUser.UserPrincipalName
                DisplayName                    = $graphUser.DisplayName
                AccountEnabled                 = $graphUser.AccountEnabled
                M365CopilotLicenseAssigned     = $assignedCopilotSkus.Count -gt 0
                M365CopilotSkuPartNumbers      = (($assignedCopilotSkus | ForEach-Object { $_.SkuPartNumber }) -join ';')
                CopilotInteractionCount        = $userAggregate.InteractionCount
                FirstInteractionUtc            = $userAggregate.FirstActivityUtc.ToString('o')
                LastInteractionUtc             = $userAggregate.LastActivityUtc.ToString('o')
                AppHosts                       = ($appHosts -join ';')
                LicenseCheckedAtUtc            = [datetime]::UtcNow.ToString('o')
                LicenseLookupStatus            = 'Success'
                LicenseLookupError             = $null
            }
        }
        catch {
            Write-Warning "Microsoft Graph lookup failed for '$userId': $($_.Exception.Message)"
            [pscustomobject][ordered]@{
                ReportStartDateUtc             = $startDateUtc.ToString('yyyy-MM-dd')
                ReportEndDateUtc               = $endDateUtc.ToString('yyyy-MM-dd')
                ReportPeriodDays               = $reportPeriodDays
                AgentName                      = ($selectedAgentNameValues -join ';')
                AgentIds                       = ($selectedAgentIdValues -join ';')
                AgentVersions                  = ($agentVersions -join ';')
                UserPrincipalName              = $userId
                DisplayName                    = $null
                AccountEnabled                 = $null
                M365CopilotLicenseAssigned     = $null
                M365CopilotSkuPartNumbers      = $null
                CopilotInteractionCount        = $userAggregate.InteractionCount
                FirstInteractionUtc            = $userAggregate.FirstActivityUtc.ToString('o')
                LastInteractionUtc             = $userAggregate.LastActivityUtc.ToString('o')
                AppHosts                       = ($appHosts -join ';')
                LicenseCheckedAtUtc            = [datetime]::UtcNow.ToString('o')
                LicenseLookupStatus            = 'Failed'
                LicenseLookupError             = $_.Exception.Message
            }
        }
    }
)

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $safeAgentName = ($AgentName -replace '[^A-Za-z0-9._-]', '_').Trim('_')
    if ([string]::IsNullOrWhiteSpace($safeAgentName)) {
        $safeAgentName = 'Agent'
    }
    $fileName = 'CopilotAgentUsers_{0}_{1:yyyyMMdd_HHmmss}.csv' -f $safeAgentName, (Get-Date)
    $OutputPath = Join-Path -Path (Get-Location).Path -ChildPath $fileName
}
elseif (-not [IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path -Path (Get-Location).Path -ChildPath $OutputPath
}

$outputDirectory = Split-Path -Parent $OutputPath
if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    throw "The output directory '$outputDirectory' does not exist."
}

$results |
    Sort-Object -Property UserPrincipalName |
    Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8

$failedLookups = @($results | Where-Object { $_.LicenseLookupStatus -eq 'Failed' }).Count
Write-Host "Exported $($results.Count) user record(s) from $matchingEventCount matching Copilot interaction(s) to: $OutputPath"
Write-Host "Purview records without agent identity: $missingAgentIdentityCount"
if ($failedLookups -gt 0) {
    Write-Warning "$failedLookups user license lookup(s) failed. Review the LicenseLookupStatus and LicenseLookupError columns in the output CSV."
}
