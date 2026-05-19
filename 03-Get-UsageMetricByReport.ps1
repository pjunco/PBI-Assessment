#Requires -Modules MicrosoftPowerBIMgmt

<#
.SYNOPSIS
    Retrieves usage & adoption metrics for all reports produced by 02-Get-PowerBIReportsByWorkspace.ps1.

.DESCRIPTION
    Reads PowerBI_Reports_All_Workspaces.json (produced by script 02), collects
    Power BI Activity Events for the last N days with a single pass of API calls,
    then computes per-report metrics for every workspace:
        - Total views / unique viewers / shares
        - Views per day / unique viewers per day / shares per day
        - Views per user breakdown
        - Most viewed pages
        - Access method breakdown (Web vs. Mobile vs. other)
        - Rank of each report by total views relative to all reports in the org

    Output:
        .\Output\PowerBI_UsageMetrics_<WorkspaceName>.json  (one per workspace)
        .\Output\PowerBI_UsageMetrics_All.json              (combined)

.PARAMETER InputJson
    Path to the combined JSON produced by script 02.
    Defaults to .\Output\PowerBI_Reports_All_Workspaces.json.

.PARAMETER DaysBack
    Number of days of activity history to retrieve. Maximum 28 (API hard limit). Defaults to 28.

.PARAMETER TenantId
    Azure AD Tenant ID. Required for service principal authentication.

.PARAMETER ClientId
    Service principal (App Registration) Client ID.

.PARAMETER ClientSecret
    Service principal Client Secret (as a SecureString).

.EXAMPLE #1
    # Interactive login (uses default input path)
    .\03-Get-UsageMetricByReport.ps1

.EXAMPLE #2
    # Service principal login
    $secret = ConvertTo-SecureString "your-secret" -AsPlainText -Force
    .\03-Get-UsageMetricByReport.ps1 -DaysBack 28 -TenantId "xxx" -ClientId "yyy" -ClientSecret $secret
#>

[CmdletBinding()]
param (
    [string]$InputJson = ".\Output\PowerBI_Reports_All_Workspaces.json",

    [int]$DaysBack = 28,

    [string]$TenantId,
    [string]$ClientId,
    [System.Security.SecureString]$ClientSecret
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ── 0. Pre-requisites notice ────────────────────────────────────────────────────
Write-Host ""
Write-Host "---------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host " PRE-REQUISITES" -ForegroundColor Yellow
Write-Host "---------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host " The account / service principal needs:" -ForegroundColor White
Write-Host "  1. Power BI Service API permission: Tenant.Read.All" -ForegroundColor Cyan
Write-Host "  2. Power BI Admin role in the Power BI Admin portal." -ForegroundColor White
Write-Host "  3. Admin portal > Tenant settings > Admin API settings" -ForegroundColor White
Write-Host "       > Allow service principals to use read-only Power BI admin APIs" -ForegroundColor Cyan
Write-Host "  NOTE: Activity Events API supports a maximum of 28 days of history." -ForegroundColor Gray
Write-Host "  NOTE: For 28 days this script makes ~28 API calls (1 per day)." -ForegroundColor Gray
Write-Host "---------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host ""

# ── 1. Install module if missing ────────────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
    Write-Host "MicrosoftPowerBIMgmt module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
}

# ── 2. Authenticate ─────────────────────────────────────────────────────────────
if ($ClientId -and $ClientSecret -and $TenantId) {
    Write-Host "Authenticating with service principal..." -ForegroundColor Cyan
    $credential = New-Object System.Management.Automation.PSCredential($ClientId, $ClientSecret)
    Connect-PowerBIServiceAccount -ServicePrincipal -Credential $credential -TenantId $TenantId
} else {
    Write-Host "Authenticating interactively (browser will open for MFA)..." -ForegroundColor Cyan
    Connect-PowerBIServiceAccount
}

# Capture the Bearer token for direct REST calls (preserves single-quoted datetime params)
$script:pbiHeaders = Get-PowerBIAccessToken

# ── Helper: fetch all activity events for a single UTC day ──────────────────────
$script:activityApiError = $false
$script:daySkipped       = $false

function Get-ActivityEventsForDay {
    param([string]$DateStr)   # yyyy-MM-dd

    # Single quotes around datetimes are REQUIRED by the API.
    # Invoke-PowerBIRestMethod URL-encodes them, so we call the REST endpoint
    # directly with Invoke-RestMethod to preserve the exact URL format.
    $startDt   = "${DateStr}T00:00:00.000Z"
    $endDt     = "${DateStr}T23:59:59.999Z"
    $uri       = "https://api.powerbi.com/v1.0/myorg/admin/activityevents?startDateTime='$startDt'&endDateTime='$endDt'"
    $allEvents = [System.Collections.Generic.List[object]]::new()

    try {
        do {
            $response = Invoke-RestMethod -Uri $uri -Headers $script:pbiHeaders -Method Get -ErrorAction Stop
            if ($response.activityEventEntities) {
                $allEvents.AddRange([object[]]$response.activityEventEntities)
            }
            $uri = $response.continuationUri   # null/empty when there are no more pages
        } while ($uri)
    }
    catch {
        # Extract HTTP status code (works on PS 5.1 and PS 7+)
        $statusCode = try { $_.Exception.Response.StatusCode.value__ } catch { 0 }
        if ($statusCode -eq 400) {
            # 400 on a single day = date is at/beyond the API's supported window.
            # Skip this day and let the loop continue to valid dates.
            $script:daySkipped = $true
            Write-Host " (skipped: outside API window)" -ForegroundColor DarkYellow
        } else {
            $script:activityApiError = $true
            $errMsg       = $_.Exception.Message
            $responseBody = $_.ErrorDetails.Message
            Write-Host "" # newline after -NoNewline
            Write-Host ""
            Write-Host "  ERROR: Activity Events API call failed$(if ($statusCode) { " (HTTP $statusCode)" })." -ForegroundColor Red
            Write-Host "  Raw error  : $errMsg" -ForegroundColor DarkRed
            if ($responseBody) {
                Write-Host "  API response: $responseBody" -ForegroundColor DarkRed
            }
            Write-Host ""
            Write-Host "  Likely causes and fixes:" -ForegroundColor Yellow
            Write-Host "  1. Missing role: your account needs the 'Power BI Administrator'" -ForegroundColor White
            Write-Host "     role in Entra ID (Fabric Admin alone is NOT sufficient)." -ForegroundColor White
            Write-Host "     -> Go to admin.microsoft.com > Users > [your account]" -ForegroundColor Cyan
            Write-Host "        > Manage roles > check 'Power BI administrator'" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "  2. Activity Events API is disabled in tenant settings:" -ForegroundColor White
            Write-Host "     -> Go to app.powerbi.com > Admin portal > Tenant settings" -ForegroundColor Cyan
            Write-Host "        > Audit and usage settings" -ForegroundColor Cyan
            Write-Host "        > Enable 'Usage metrics for content creators'" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "  Stopping event collection. Metrics will show 0 views/viewers." -ForegroundColor Yellow
        }
    }

    return $allEvents
}

try {
    # ── 3. Load combined workspace+report list from script 02 output ─────────────
    if (-not (Test-Path $InputJson)) {
        throw "Input file not found: $InputJson. Run 02-Get-PowerBIReportsByWorkspace.ps1 first."
    }
    $allWorkspaces = Get-Content $InputJson -Raw | ConvertFrom-Json
    if (-not $allWorkspaces -or $allWorkspaces.Count -eq 0) {
        throw "Input file contains no workspaces."
    }

    # Only process workspaces that have at least one report
    $workspacesWithReports = $allWorkspaces | Where-Object { $_.ReportCount -gt 0 }
    $totalReportCount = ($allWorkspaces | ForEach-Object { $_.ReportCount } | Measure-Object -Sum).Sum
    Write-Host "Loaded $($allWorkspaces.Count) workspace(s), $totalReportCount total report(s) from $InputJson" -ForegroundColor Green
    Write-Host "$($workspacesWithReports.Count) workspace(s) have reports and will be processed." -ForegroundColor Gray

    $outputDir = ".\Output"
    if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }

    # ── 4. Collect activity events for the last N days (single pass) ─────────────
    # The Activity Events API only supports the last 28 days (hard API limit).
    if ($DaysBack -gt 28) {
        Write-Warning "DaysBack ($DaysBack) exceeds the 28-day limit of the Activity Events API. Capping at 28."
        $DaysBack = 28
    }

    $endDate   = (Get-Date).Date
    $startDate = $endDate.AddDays(-$DaysBack)

    Write-Host "Collecting activity events from $($startDate.ToString('yyyy-MM-dd')) to $($endDate.AddDays(-1).ToString('yyyy-MM-dd'))..." -ForegroundColor Cyan

    $allEvents = [System.Collections.Generic.List[object]]::new()
    $current   = $startDate

    while ($current -lt $endDate) {
        if ($script:activityApiError) { break }   # stop on first API failure
        $dayStr = $current.ToString('yyyy-MM-dd')
        Write-Host "  Fetching $dayStr..." -ForegroundColor Gray -NoNewline
        $script:daySkipped = $false
        $dayEvents = @(Get-ActivityEventsForDay -DateStr $dayStr)
        if ($dayEvents.Count -gt 0) { $allEvents.AddRange([object[]]$dayEvents) }
        if (-not $script:activityApiError -and -not $script:daySkipped) {
            Write-Host " $($dayEvents.Count) event(s)" -ForegroundColor Gray
        }
        $current = $current.AddDays(1)
    }

    Write-Host "Total events retrieved: $($allEvents.Count)" -ForegroundColor Green

    # ── 5. Separate ViewReport and ShareReport events ─────────────────────────────
    $viewEvents  = $allEvents | Where-Object { $_.Activity -eq 'ViewReport' }
    $shareEvents = $allEvents | Where-Object { $_.Activity -eq 'ShareReport' }

    # ── 6. Org-wide rank map ──────────────────────────────────────────────────────
    $orgViewTotals = $viewEvents |
        Group-Object { $_.ArtifactId } |
        ForEach-Object { [PSCustomObject]@{ ReportId = $_.Name.ToLower(); TotalViews = $_.Count } } |
        Sort-Object TotalViews -Descending

    $rankMap = @{}
    $rank = 1
    foreach ($entry in $orgViewTotals) { $rankMap[$entry.ReportId] = $rank; $rank++ }
    $orgReportCount = $orgViewTotals.Count
    Write-Host "Org-wide: $($viewEvents.Count) view event(s) across $orgReportCount unique report(s)." -ForegroundColor Green

    # ── 7. Process each workspace ─────────────────────────────────────────────────
    $allOutputs = [System.Collections.Generic.List[object]]::new()

    foreach ($workspace in $workspacesWithReports) {
        Write-Host ""
        Write-Host "Computing metrics: $($workspace.WorkspaceName)" -ForegroundColor Cyan

        $reportMetrics = foreach ($rpt in $workspace.Reports) {
            $rid = $rpt.ReportId.ToString().ToLower()

            $rViews  = $viewEvents  | Where-Object { $_.ArtifactId -and $_.ArtifactId.ToString().ToLower() -eq $rid }
            $rShares = $shareEvents | Where-Object { $_.ArtifactId -and $_.ArtifactId.ToString().ToLower() -eq $rid }

            $totalViews   = @($rViews).Count
            $totalViewers = ($rViews | Select-Object -ExpandProperty UserId -Unique | Measure-Object).Count
            $totalShares  = @($rShares).Count

            $viewsPerDay = $rViews |
                Group-Object { ([datetime]$_.CreationTime).ToString('yyyy-MM-dd') } |
                ForEach-Object { [ordered]@{ Date = $_.Name; Views = $_.Count } } |
                Sort-Object { $_.Date }

            $uniqueViewersPerDay = $rViews |
                Group-Object { ([datetime]$_.CreationTime).ToString('yyyy-MM-dd') } |
                ForEach-Object {
                    [ordered]@{
                        Date          = $_.Name
                        UniqueViewers = ($_.Group | Select-Object -ExpandProperty UserId -Unique | Measure-Object).Count
                    }
                } | Sort-Object { $_.Date }

            $viewsPerUser = $rViews |
                Group-Object UserId |
                ForEach-Object { [ordered]@{ UserId = $_.Name; Views = $_.Count } } |
                Sort-Object { $_.Views } -Descending

            $sharesPerDay = $rShares |
                Group-Object { ([datetime]$_.CreationTime).ToString('yyyy-MM-dd') } |
                ForEach-Object { [ordered]@{ Date = $_.Name; Shares = $_.Count } } |
                Sort-Object { $_.Date }

            $pageViews = $rViews |
                Where-Object { $_.ReportPage } |
                Group-Object ReportPage |
                ForEach-Object { [ordered]@{ Page = $_.Name; Views = $_.Count } } |
                Sort-Object { $_.Views } -Descending

            $accessBreakdown = $rViews |
                Group-Object { if ($_.ConsumptionMethod) { $_.ConsumptionMethod } else { 'Unknown' } } |
                ForEach-Object { [ordered]@{ Method = $_.Name; Views = $_.Count } } |
                Sort-Object { $_.Views } -Descending

            $browserBreakdown = $rViews |
                Group-Object {
                    $ua = $_.UserAgent
                    if     (-not $ua)                             { 'Unknown'           }
                    elseif ($ua -match 'PowerBIDesktop')          { 'Power BI Desktop'  }
                    elseif ($ua -match 'Edg/|EdgA/|Edge/')        { 'Edge'              }
                    elseif ($ua -match 'Firefox/')                { 'Firefox'           }
                    elseif ($ua -match 'OPR/|Opera/')             { 'Opera'             }
                    elseif ($ua -match 'Chrome/')                 { 'Chrome'            }
                    elseif ($ua -match 'Safari/')                 { 'Safari'            }
                    else                                          { 'Other'             }
                } |
                ForEach-Object { [ordered]@{ Browser = $_.Name; Views = $_.Count } } |
                Sort-Object { $_.Views } -Descending

            $orgRank = if ($rankMap.ContainsKey($rid)) { $rankMap[$rid] } else { $null }

            Write-Host "  $($rpt.ReportName): $totalViews views, $totalViewers viewers" -ForegroundColor Gray

            [ordered]@{
                ReportId              = $rpt.ReportId
                ReportName            = $rpt.ReportName
                WebUrl                = $rpt.WebUrl
                PeriodDays            = $DaysBack
                TotalViews            = $totalViews
                TotalUniqueViewers    = $totalViewers
                TotalShares           = $totalShares
                RankByTotalViewsInOrg = $orgRank
                TotalReportsInOrg     = $orgReportCount
                ViewsPerDay           = @($viewsPerDay)
                UniqueViewersPerDay   = @($uniqueViewersPerDay)
                ViewsPerUser          = @($viewsPerUser)
                SharesPerDay          = @($sharesPerDay)
                MostViewedPages       = @($pageViews)
                AccessMethodBreakdown = @($accessBreakdown)
                BrowserBreakdown      = @($browserBreakdown)
            }
        }

        $wsOutput = [ordered]@{
            GeneratedAt       = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ')
            WorkspaceId       = $workspace.WorkspaceId
            WorkspaceName     = $workspace.WorkspaceName
            CapacityType      = $workspace.CapacityType
            PeriodStartDate   = $startDate.ToString('yyyy-MM-dd')
            PeriodEndDate     = $endDate.AddDays(-1).ToString('yyyy-MM-dd')
            PeriodDays        = $DaysBack
            ReportCount       = $workspace.ReportCount
            Reports           = $reportMetrics
        }

        # Per-workspace file
        $safeWsName = $workspace.WorkspaceName -replace '[\\/:*?"<>|]', '_'
        $outputPath = "$outputDir\PowerBI_UsageMetrics_$safeWsName.json"
        $wsOutput | ConvertTo-Json -Depth 10 | Out-File -LiteralPath $outputPath -Encoding UTF8
        Write-Host "  Exported: $outputPath" -ForegroundColor Gray

        $allOutputs.Add($wsOutput)
    }

    # ── 8. Export combined file ───────────────────────────────────────────────────
    $combinedPath = "$outputDir\PowerBI_UsageMetrics_All.json"
    $allOutputs | ConvertTo-Json -Depth 10 | Out-File -LiteralPath $combinedPath -Encoding UTF8
    Write-Host ""
    Write-Host "Combined usage metrics exported to: $combinedPath" -ForegroundColor Green

    # ── 9. Final console summary (workspace-level totals only) ───────────────────
    Write-Host ""
    Write-Host "===== USAGE SUMMARY =====" -ForegroundColor Yellow
    Write-Host ""
    $grandTotalViews   = 0
    $grandTotalReports = 0
    foreach ($ws in $allOutputs) {
        $wsViews   = ($ws.Reports | ForEach-Object { $_.TotalViews }   | Measure-Object -Sum).Sum
        $wsViewers = ($ws.Reports | ForEach-Object { $_.TotalUniqueViewers } | Measure-Object -Sum).Sum
        $grandTotalViews   += $wsViews
        $grandTotalReports += $ws.ReportCount
        $viewsLabel = if ($script:activityApiError) { "N/A (API error)" } else { "$wsViews views, $wsViewers viewers" }
        Write-Host ("  {0,-45} [{1}]  Reports: {2,3}   {3}" -f $ws.WorkspaceName, $ws.CapacityType, $ws.ReportCount, $viewsLabel) -ForegroundColor White
    }
    Write-Host ""
    Write-Host ("  Total: {0} report(s) across {1} workspace(s)" -f $grandTotalReports, $allOutputs.Count) -ForegroundColor Green

    # ── 10. Warn if data is incomplete due to API error ───────────────────────────
    if ($script:activityApiError) {
        Write-Host ""
        Write-Host "---------------------------------------------------------------------" -ForegroundColor Red
        Write-Host "  WARNING: All exported files contain 0 values (no usage data)." -ForegroundColor Red
        Write-Host "  The Activity Events API could not be accessed." -ForegroundColor Red
        Write-Host ""
        Write-Host "  To get real usage data:" -ForegroundColor Yellow
        Write-Host "  1. Go to admin.microsoft.com -> Users -> $($env:USERNAME)" -ForegroundColor White
        Write-Host "     -> Manage roles -> check 'Power BI administrator'" -ForegroundColor White
        Write-Host "  2. Re-run: .\03-Get-UsageMetricByReport.ps1" -ForegroundColor Cyan
        Write-Host "---------------------------------------------------------------------" -ForegroundColor Red
    }
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    Disconnect-PowerBIServiceAccount
    Write-Host "Disconnected from Power BI." -ForegroundColor Gray
}
