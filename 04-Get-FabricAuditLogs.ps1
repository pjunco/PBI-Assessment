<#
.SYNOPSIS
    Exports Power BI / Microsoft Fabric audit logs for the entire organization.

.DESCRIPTION
    Connects to the Microsoft 365 Unified Audit Log (via Exchange Online Management),
    retrieves all Power BI and Fabric audit records for the last N days, and produces:
        - .\Output\Fabric_AuditLog_Raw_<date>.json       — every raw audit record
        - .\Output\Fabric_AuditLog_Summary_<date>.json   — aggregated per-user and
          per-activity breakdown (views, exports, refreshes, shares, etc.)

    Prerequisites:
        • The Fabric Administrator role (or Global Administrator / Compliance role)
          is required to read audit logs.
        • Unified Audit Logging must be enabled in the Microsoft Purview compliance
          portal (Audit > Audit log search).
        • The ExchangeOnlineManagement module must be installed.

.PARAMETER DaysBack
    Number of days of history to retrieve (max 90 for E3; up to 1 year for E5/A5).
    Defaults to 90.

.PARAMETER UserPrincipalName
    Optional. Restrict audit records to a single user (UPN). Omit for org-wide.

.PARAMETER Activities
    Optional array of specific activity names to filter.
    Defaults to a comprehensive set of Power BI / Fabric activities.

.PARAMETER TenantId
    Azure AD Tenant ID. Optional — used to disambiguate when the account belongs
    to multiple tenants.

.PARAMETER OutputDir
    Folder for output files. Defaults to .\Output.

.EXAMPLE
    # Interactive — all users, last 90 days
    .\04-Get-FabricAuditLogs.ps1

.EXAMPLE
    # Specific user, last 30 days
    .\04-Get-FabricAuditLogs.ps1 -DaysBack 30 -UserPrincipalName "john@contoso.com"

.EXAMPLE
    # Service account with Tenant ID
    .\04-Get-FabricAuditLogs.ps1 -TenantId "xxx-yyy-zzz" -DaysBack 90
#>

[CmdletBinding()]
param (
    [ValidateRange(1, 365)]
    [int]$DaysBack = 90,

    [string]$UserPrincipalName,

    [string[]]$Activities = @(
        'ViewReport',
        'ExportReport',
        'ExportArtifact',
        'ExportDataflow',
        'AnalyzedByExternalApplication',
        'DatasetRefresh',
        'RefreshDataset',
        'ShareReport',
        'ShareDashboard',
        'CreateReport',
        'EditReport',
        'DeleteReport',
        'PublishToWebReport',
        'ViewDashboard',
        'ExportTile',
        'CreateDashboard',
        'DeleteDashboard',
        'ViewDataset',
        'CreateDataset',
        'DeleteDataset',
        'CreateDataflow',
        'DeleteDataflow',
        'RefreshDataflow',
        'CreateApp',
        'UpdateApp',
        'DeleteApp',
        'InstallApp',
        'CreateWorkspace',
        'UpdateWorkspace',
        'DeleteWorkspace',
        'AddGroupMembers',
        'DeleteGroupMembers'
    ),

    [string]$TenantId,

    [string]$OutputDir = ".\Output"
)

Set-StrictMode -Off
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ── 0. Pre-requisites notice ────────────────────────────────────────────────────
Write-Host ""
Write-Host "---------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host " PRE-REQUISITES" -ForegroundColor Yellow
Write-Host "---------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host " 1. Role required: Power BI Administrator in Entra ID." -ForegroundColor White
Write-Host "    (Fabric Admin alone is NOT sufficient for the Activity Events API.)" -ForegroundColor White
Write-Host " 2. Module required: MicrosoftPowerBIMgmt" -ForegroundColor White
Write-Host "    Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser" -ForegroundColor Cyan
Write-Host " NOTE: Activity Events API returns up to 90 days of history." -ForegroundColor Gray
Write-Host " NOTE: One API call is made per day; continuation tokens handle" -ForegroundColor Gray
Write-Host "       pagination within each day." -ForegroundColor Gray
Write-Host "---------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host ""

# ── 1. Verify MicrosoftPowerBIMgmt module ────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
    Write-Host "MicrosoftPowerBIMgmt module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
}

# ── 2. Authenticate to Power BI ───────────────────────────────────────────────────
Write-Host "Authenticating interactively (browser will open for MFA)..." -ForegroundColor Cyan
$connectSplat = @{}
if ($TenantId) { $connectSplat['TenantId'] = $TenantId }
Connect-PowerBIServiceAccount @connectSplat | Out-Null

# ── 3. Set date range ────────────────────────────────────────────────────────────
$endDate   = (Get-Date).ToUniversalTime()
$startDate = $endDate.AddDays(-$DaysBack)

Write-Host "Querying audit logs from $($startDate.ToString('yyyy-MM-dd')) to $($endDate.ToString('yyyy-MM-dd'))..." -ForegroundColor Cyan
if ($UserPrincipalName) {
    Write-Host "Filtering to user: $UserPrincipalName" -ForegroundColor Gray
}

# ── 4. Collect activity events via Power BI Activity Events API ──────────────────
$allRecords     = [System.Collections.Generic.List[object]]::new()
$script:apiError = $false
$baseApi        = 'https://api.powerbi.com/'

Write-Host "Collecting activity events ($DaysBack days)..." -ForegroundColor Cyan

for ($d = 0; $d -lt $DaysBack; $d++) {
    $dayStart = $startDate.Date.AddDays($d)
    $startStr = $dayStart.ToString('yyyy-MM-dd') + 'T00:00:00.000Z'
    $endStr   = $dayStart.ToString('yyyy-MM-dd') + 'T23:59:59.999Z'

    Write-Host "  $($dayStart.ToString('yyyy-MM-dd'))..." -ForegroundColor Gray -NoNewline

    $url      = "v1.0/myorg/admin/activityevents?startDateTime='$startStr'&endDateTime='$endStr'"
    $dayTotal = 0

    do {
        try {
            $response = Invoke-PowerBIRestMethod -Url $url -Method Get -ErrorAction Stop | ConvertFrom-Json
        } catch {
            Write-Host " ERROR: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host ""
            Write-Host "  Likely cause: your account is missing the 'Power BI Administrator'" -ForegroundColor Yellow
            Write-Host "  role in Entra ID (Fabric Admin alone is NOT sufficient)." -ForegroundColor Yellow
            $script:apiError = $true
            break
        }

        if ($response.activityEventEntities) {
            foreach ($ev in $response.activityEventEntities) {
                $allRecords.Add([ordered]@{
                    CreationDate       = $ev.CreationTime
                    Activity           = if ($ev.Activity)      { $ev.Activity }      else { $ev.Operation }
                    UserId             = $ev.UserId
                    WorkspaceName      = $ev.WorkSpaceName
                    WorkspaceId        = $ev.WorkspaceId
                    ArtifactId         = $ev.ArtifactId
                    ArtifactName       = $ev.ArtifactName
                    ArtifactKind       = $ev.ArtifactKind
                    CapacityId         = $ev.CapacityId
                    CapacityName       = $ev.CapacityName
                    ConsumptionMethod  = $ev.ConsumptionMethod
                    DistributionMethod = $ev.DistributionMethod
                    ReportPage         = $ev.ReportPage
                    RequestId          = $ev.RequestId
                    IsSuccess          = $ev.IsSuccess
                })
                $dayTotal++
            }
        }

        if ($response.continuationUri) {
            # Invoke-PowerBIRestMethod expects a relative URL; strip the base if absolute
            $url = $response.continuationUri -replace [regex]::Escape($baseApi), ''
        } else {
            break
        }
    } while ($true)

    Write-Host " $dayTotal event(s)" -ForegroundColor Gray
    if ($script:apiError) { break }
}

Write-Host "Total activity events retrieved: $($allRecords.Count)" -ForegroundColor Green

# Apply client-side filters (API does not support server-side user/activity filtering)
if ($UserPrincipalName -and $allRecords.Count -gt 0) {
    $before     = $allRecords.Count
    $allRecords = [System.Collections.Generic.List[object]]($allRecords | Where-Object { $_.UserId -eq $UserPrincipalName })
    Write-Host "  Filtered to user '$UserPrincipalName': $($allRecords.Count) of $before event(s)" -ForegroundColor Gray
}
if ($Activities.Count -gt 0 -and $allRecords.Count -gt 0) {
    $before     = $allRecords.Count
    $allRecords = [System.Collections.Generic.List[object]]($allRecords | Where-Object { $Activities -contains $_.Activity })
    Write-Host "  Filtered to $($Activities.Count) activity type(s): $($allRecords.Count) of $before event(s)" -ForegroundColor Gray
}

# ── 5. Build per-user summary ────────────────────────────────────────────────────
Write-Host "Aggregating per-user activity summary..." -ForegroundColor Cyan

$perUserSummary = $allRecords |
    Group-Object { $_.UserId } |
    ForEach-Object {
        $userRecords = $_.Group
        $activityBreakdown = $userRecords |
            Group-Object { $_.Activity } |
            ForEach-Object { [ordered]@{ Activity = $_.Name; Count = $_.Count } } |
            Sort-Object { $_.Count } -Descending

        [ordered]@{
            UserId            = $_.Name
            TotalEvents       = $userRecords.Count
            FirstActivity     = ($userRecords | Sort-Object CreationDate | Select-Object -First 1).CreationDate
            LastActivity      = ($userRecords | Sort-Object CreationDate -Descending | Select-Object -First 1).CreationDate
            ActivityBreakdown = @($activityBreakdown)
        }
    } | Sort-Object { $_.TotalEvents } -Descending

# ── 6. Build per-activity summary ───────────────────────────────────────────────
$perActivitySummary = $allRecords |
    Group-Object { $_.Activity } |
    ForEach-Object {
        [ordered]@{
            Activity     = $_.Name
            TotalCount   = $_.Count
            UniqueUsers  = ($_.Group | Select-Object -ExpandProperty UserId -Unique | Measure-Object).Count
        }
    } | Sort-Object { $_.TotalCount } -Descending

# ── 7. Build per-artifact summary (report/dataset/workspace) ─────────────────────
$perArtifactSummary = $allRecords |
    Where-Object { $_.ArtifactId } |
    Group-Object { $_.ArtifactId } |
    ForEach-Object {
        $first = $_.Group | Select-Object -First 1
        [ordered]@{
            ArtifactId    = $_.Name
            ArtifactName  = $first.ArtifactName
            ArtifactKind  = $first.ArtifactKind
            WorkspaceName = $first.WorkspaceName
            TotalEvents   = $_.Count
            UniqueUsers   = ($_.Group | Select-Object -ExpandProperty UserId -Unique | Measure-Object).Count
            Activities    = @(
                $_.Group | Group-Object { $_.Activity } |
                ForEach-Object { [ordered]@{ Activity = $_.Name; Count = $_.Count } } |
                Sort-Object { $_.Count } -Descending
            )
        }
    } | Sort-Object { $_.TotalEvents } -Descending

# ── 8. Build daily trend ─────────────────────────────────────────────────────────
$dailyTrend = $allRecords |
    Group-Object { ([datetime]$_.CreationDate).ToString('yyyy-MM-dd') } |
    ForEach-Object {
        [ordered]@{
            Date         = $_.Name
            TotalEvents  = $_.Count
            UniqueUsers  = ($_.Group | Select-Object -ExpandProperty UserId -Unique | Measure-Object).Count
        }
    } | Sort-Object { $_.Date }

# ── 9. Build output objects ──────────────────────────────────────────────────────
$dateTag = (Get-Date -Format 'yyyyMMdd')

$summaryOutput = [ordered]@{
    GeneratedAt         = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ')
    PeriodStartDate     = $startDate.ToString('yyyy-MM-dd')
    PeriodEndDate       = $endDate.ToString('yyyy-MM-dd')
    PeriodDays          = $DaysBack
    FilteredUser        = if ($UserPrincipalName) { $UserPrincipalName } else { 'All Users (org-wide)' }
    TotalAuditRecords   = $allRecords.Count
    UniqueUsers         = ($allRecords | Select-Object -ExpandProperty UserId -Unique | Measure-Object).Count
    UniqueArtifacts     = ($allRecords | Where-Object { $_.ArtifactId } | Select-Object -ExpandProperty ArtifactId -Unique | Measure-Object).Count
    DailyTrend          = @($dailyTrend)
    PerActivitySummary  = @($perActivitySummary)
    PerUserSummary      = @($perUserSummary)
    PerArtifactSummary  = @($perArtifactSummary)
}

# ── 10. Ensure output directory exists ───────────────────────────────────────────
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# ── 11. Export raw records ───────────────────────────────────────────────────────
$rawPath = "$OutputDir\Fabric_AuditLog_Raw_$dateTag.json"
$allRecords | ConvertTo-Json -Depth 10 | Out-File -FilePath $rawPath -Encoding UTF8
Write-Host "Raw audit records exported to: $rawPath" -ForegroundColor Green

# ── 12. Export summary ───────────────────────────────────────────────────────────
$summaryPath = "$OutputDir\Fabric_AuditLog_Summary_$dateTag.json"
$summaryOutput | ConvertTo-Json -Depth 10 | Out-File -FilePath $summaryPath -Encoding UTF8
Write-Host "Audit summary exported to: $summaryPath" -ForegroundColor Green

# Warn if data is incomplete due to API error
if ($script:apiError) {
    Write-Host ""
    Write-Host "---------------------------------------------------------------------" -ForegroundColor Red
    Write-Host "  WARNING: Exported files contain 0 records (API access denied)." -ForegroundColor Red
    Write-Host ""
    Write-Host "  To get real audit data:" -ForegroundColor Yellow
    Write-Host "  1. Go to admin.microsoft.com -> Users -> $($env:USERNAME)" -ForegroundColor White
    Write-Host "     -> Manage roles -> check 'Power BI administrator'" -ForegroundColor White
    Write-Host "  2. Re-run: .\04-Get-FabricAuditLogs.ps1" -ForegroundColor Cyan
    Write-Host "---------------------------------------------------------------------" -ForegroundColor Red
}

# ── 13. Print quick console table ────────────────────────────────────────────────
Write-Host ""
Write-Host "-- Activity Breakdown -----------------------------------------------" -ForegroundColor Cyan
$perActivitySummary |
    ForEach-Object { [PSCustomObject]@{ Activity = $_.Activity; TotalCount = $_.TotalCount; UniqueUsers = $_.UniqueUsers } } |
    Format-Table -AutoSize

Write-Host "-- Top 10 Users by Activity -----------------------------------------" -ForegroundColor Cyan
$perUserSummary | Select-Object -First 10 |
    ForEach-Object { [PSCustomObject]@{ UserId = $_.UserId; TotalEvents = $_.TotalEvents; FirstActivity = $_.FirstActivity; LastActivity = $_.LastActivity } } |
    Format-Table -AutoSize

# ── 14. Disconnect ───────────────────────────────────────────────────────────────
Disconnect-PowerBIServiceAccount
Write-Host "Disconnected from Power BI." -ForegroundColor Gray
