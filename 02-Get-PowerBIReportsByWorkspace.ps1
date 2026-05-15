###prerequisite:
###powershell.exe -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser"
###Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber




#Requires -Modules MicrosoftPowerBIMgmt


<#
.SYNOPSIS
    Retrieves reports for all workspaces produced by 01-Get-PowerBIWorkspaces.ps1.

.DESCRIPTION
    Reads the workspace list from the script 01 JSON output, authenticates to
    Power BI, issues a single batched Scanner API call for all workspaces, then
    exports per-workspace JSON files and one combined summary file.

.PARAMETER InputJson
    Path to the JSON produced by script 01. Defaults to .\Output\PowerBI_Workspaces.json.

.PARAMETER TenantId
    Azure AD Tenant ID. Required for service principal authentication.

.PARAMETER ClientId
    Service principal (App Registration) Client ID.

.PARAMETER ClientSecret
    Service principal Client Secret (as a SecureString).

.EXAMPLE
    # Interactive login (uses default input path)
    .\02-Get-PowerBIReportsByWorkspace.ps1

.EXAMPLE
    # Custom input path
    .\02-Get-PowerBIReportsByWorkspace.ps1 -InputJson ".\Output\PowerBI_Workspaces.json"

.EXAMPLE
    # Service principal login
    $secret = ConvertTo-SecureString "your-secret" -AsPlainText -Force
    .\02-Get-PowerBIReportsByWorkspace.ps1 -TenantId "xxx" -ClientId "yyy" -ClientSecret $secret
#>

[CmdletBinding()]
param (
    [string]$InputJson   = ".\Output\PowerBI_Workspaces.json",
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
Write-Host " The account needs:" -ForegroundColor White
Write-Host "  1. Power BI Service API permission: Tenant.Read.All" -ForegroundColor Cyan
Write-Host "  2. Power BI Admin portal > Tenant settings > Admin API settings" -ForegroundColor Cyan
    Write-Host "       > Allow service principals to use read-only Power BI admin APIs" -ForegroundColor Cyan
    Write-Host "  3. Power BI Admin portal > Tenant settings > Admin API settings" -ForegroundColor Cyan
    Write-Host "       > Enhance admin APIs responses with detailed metadata" -ForegroundColor Cyan
    Write-Host "       (Required for TableCount/MeasureCount via datasetSchema=true)" -ForegroundColor DarkCyan
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
    Write-Host "Authenticating interactively..." -ForegroundColor Cyan
    Connect-PowerBIServiceAccount
}

try {
    # ── 3. Load workspaces from script 01 output ────────────────────────────────
    if (-not (Test-Path $InputJson)) {
        throw "Input file not found: $InputJson. Run 01-Get-PowerBIWorkspaces.ps1 first."
    }
    $workspaces = Get-Content $InputJson -Raw | ConvertFrom-Json
    Write-Host "Loaded $($workspaces.Count) workspace(s) from $InputJson" -ForegroundColor Green

    $outputDir = ".\Output"
    if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }

    # ── 4. Single batched Scanner API call for all workspaces ────────────────────
    Write-Host "Initiating batched Scanner API scan for $($workspaces.Count) workspace(s)..." -ForegroundColor Cyan
    $wsIds = $workspaces | ForEach-Object { $_.Id.ToString() }
    $scanRequestBody = @{ workspaces = $wsIds } | ConvertTo-Json
    $scanResponse = Invoke-PowerBIRestMethod `
        -Url "admin/workspaces/getInfo?datasetExpressions=false&datasetSchema=true&datasourceDetails=true&getArtifactUsers=false&lineage=false" `
        -Method Post `
        -Body $scanRequestBody
    $scanId = ($scanResponse | ConvertFrom-Json).id
    Write-Host "Scan initiated. ScanId: $scanId" -ForegroundColor Gray

    # Poll until scan completes
    $maxWaitSeconds = 120
    $waited = 0
    do {
        Start-Sleep -Seconds 3
        $waited += 3
        $statusResponse = Invoke-PowerBIRestMethod `
            -Url "admin/workspaces/scanStatus/$scanId" `
            -Method Get | ConvertFrom-Json
        Write-Host "  Scan status: $($statusResponse.status)" -ForegroundColor Gray
    } while ($statusResponse.status -ne 'Succeeded' -and $waited -lt $maxWaitSeconds)

    # Build dataset metadata map: DatasetId -> { DataSources }
    $datasetMap = @{}
    if ($statusResponse.status -eq 'Succeeded') {
        $scanResult = Invoke-PowerBIRestMethod `
            -Url "admin/workspaces/scanResult/$scanId" `
            -Method Get | ConvertFrom-Json

        foreach ($scannedWs in $scanResult.workspaces) {
            foreach ($ds in $scannedWs.datasets) {
                $dataSources = @()
                if ($ds.PSObject.Properties['datasourceUsages']) {
                    foreach ($usage in $ds.datasourceUsages) {
                        $instance = $scanResult.datasourceInstances |
                            Where-Object { $_.datasourceId -eq $usage.datasourceInstanceId } |
                            Select-Object -First 1
                        if ($instance) {
                            $dataSources += [ordered]@{
                                datasourceName     = $instance.datasourceType
                                connectionDetails  = $instance.connectionDetails
                            }
                        }
                    }
                }
                # Build table summary from schema
                $tables = @()
                if ($ds.PSObject.Properties['tables']) {
                    $tables = @($ds.tables | ForEach-Object {
                        [ordered]@{
                            TableName    = $_.name
                            MeasureCount = if ($_.PSObject.Properties['measures']) { $_.measures.Count } else { 0 }
                        }
                    })
                }

                $datasetMap[$ds.id] = [ordered]@{
                    TableCount   = $tables.Count
                    MeasureCount = [int]($tables | Measure-Object -Property MeasureCount -Sum).Sum
                    Tables       = $tables
                    DataSources  = $dataSources
                }
            }
        }
        Write-Host "Dataset metadata retrieved for $($datasetMap.Count) dataset(s)." -ForegroundColor Green
    } else {
        Write-Warning "Scan did not complete in time. Dataset metadata will be unavailable."
    }

    # ── 5. Process each workspace ────────────────────────────────────────────────
    $allSummaries = [System.Collections.Generic.List[object]]::new()

    foreach ($workspace in $workspaces) {
        Write-Host ""
        Write-Host "Processing: $($workspace.Name)" -ForegroundColor Cyan

        $reports = Get-PowerBIReport -WorkspaceId $workspace.Id -Scope Organization
        Write-Host "  Found $($reports.Count) report(s)." -ForegroundColor Gray

        # Get isRefreshable per dataset to infer connection mode (Import vs DirectQuery/LiveConnect)
        $refreshableMap   = @{}
        try {
            $dsListResponse = Invoke-PowerBIRestMethod -Url "admin/groups/$($workspace.Id)/datasets" -Method Get | ConvertFrom-Json
            foreach ($ds in $dsListResponse.value) {
                $refreshableMap[$ds.id] = $ds.isRefreshable
            }
            Write-Host "  Retrieved isRefreshable for $($refreshableMap.Count) dataset(s)." -ForegroundColor Gray
        } catch {
            Write-Warning "  Could not retrieve dataset list for $($workspace.Name): $_"
        }

        # Fetch datasources per unique dataset via Admin API
        # GET admin/groups/{wsId}/datasets/{datasetId}/datasources is more reliable than the Scanner API
        $datasourcesMap = @{}
        $uniqueDatasetIds = $reports | Select-Object -ExpandProperty DatasetId -Unique
        foreach ($dsId in $uniqueDatasetIds) {
            try {
                $dsResp = Invoke-PowerBIRestMethod `
                    -Url "admin/datasets/$dsId/datasources" `
                    -Method Get | ConvertFrom-Json
                $datasourcesMap[$dsId.ToString()] = @(
                    $dsResp.value | ForEach-Object {
                        [ordered]@{
                            datasourceName    = $_.datasourceType
                            connectionDetails = $_.connectionDetails
                        }
                    }
                )
            } catch {
                $datasourcesMap[$dsId.ToString()] = @()
            }
        }
        Write-Host "  Retrieved datasources for $($datasourcesMap.Count) dataset(s)." -ForegroundColor Gray

        $enrichedReports = foreach ($report in $reports) {
            $dsMeta         = $datasetMap[$report.DatasetId.ToString()]
            $isRefreshable  = $refreshableMap[$report.DatasetId.ToString()]
            $connectionMode = if ($null -ne $isRefreshable) {
                if ($isRefreshable) { 'Import' } else { 'DirectQuery_or_LiveConnect' }
            } else { 'Unknown' }
            [ordered]@{
                ReportId     = $report.Id
                ReportName   = $report.Name
                WebUrl       = $report.WebUrl
                EmbedUrl     = $report.EmbedUrl
                Dataset      = [ordered]@{
                    DatasetId      = $report.DatasetId
                    ConnectionMode = $connectionMode
                }
                DataSources  = if ($datasourcesMap.ContainsKey($report.DatasetId.ToString())) { $datasourcesMap[$report.DatasetId.ToString()] } else { @() }
            }
        }

        $output = [ordered]@{
            WorkspaceName                = $workspace.Name
            WorkspaceId                  = $workspace.Id
            WorkspaceType                = $workspace.Type
            WorkspaceState               = $workspace.State
            IsOnDedicatedCapacity        = $workspace.IsOnDedicatedCapacity
            CapacityId                   = $workspace.CapacityId
            CapacityName                 = if ($workspace.PSObject.Properties['CapacityName']) { $workspace.CapacityName } else { 'Unknown' }
            CapacitySku                  = if ($workspace.PSObject.Properties['CapacitySku'])  { $workspace.CapacitySku  } else { 'Unknown' }
            CapacityType                 = $workspace.CapacityType
            DefaultDatasetStorageFormat  = if ($workspace.PSObject.Properties['DefaultDatasetStorageFormat']) { $workspace.DefaultDatasetStorageFormat } else { 'Small' }
            ReportCount                  = $reports.Count
            Reports                      = $enrichedReports
        }

        # Per-workspace file
        $safeWsName = $workspace.Name -replace '[\\/:*?"<>|]', '_'
        $outputPath = "$outputDir\PowerBI_Reports_By_WS_$safeWsName.json"
        $output | ConvertTo-Json -Depth 5 | Out-File -LiteralPath $outputPath -Encoding UTF8
        Write-Host "  Exported: $outputPath" -ForegroundColor Gray

        $allSummaries.Add([ordered]@{
            WorkspaceId                  = $workspace.Id
            WorkspaceName                = $workspace.Name
            CapacityId                   = $workspace.CapacityId
            CapacityName                 = if ($workspace.PSObject.Properties['CapacityName']) { $workspace.CapacityName } else { 'Unknown' }
            CapacitySku                  = if ($workspace.PSObject.Properties['CapacitySku'])  { $workspace.CapacitySku  } else { 'Unknown' }
            CapacityType                 = $workspace.CapacityType
            DefaultDatasetStorageFormat  = if ($workspace.PSObject.Properties['DefaultDatasetStorageFormat']) { $workspace.DefaultDatasetStorageFormat } else { 'Small' }
            ReportCount                  = $reports.Count
            Reports                      = @(
                $enrichedReports | ForEach-Object {
                    [ordered]@{
                        ReportId       = $_.ReportId
                        ReportName     = $_.ReportName
                        WebUrl         = $_.WebUrl
                        DatasetId      = $_.Dataset.DatasetId
                        ConnectionMode = $_.Dataset.ConnectionMode
                        DataSources    = $_.DataSources
                    }
                }
            )
        })
    }

    # ── 6. Export combined summary ───────────────────────────────────────────────
    $combinedPath = "$outputDir\PowerBI_Reports_All_Workspaces.json"
    $allSummaries | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $combinedPath -Encoding UTF8
    Write-Host ""
    Write-Host "Combined summary exported to: $combinedPath" -ForegroundColor Green

    # ── 7. Final console summary ─────────────────────────────────────────────────
    Write-Host ""
    Write-Host "===== SUMMARY =====" -ForegroundColor Yellow
    $allSummaries | ForEach-Object {
        Write-Host "  $($_.WorkspaceName) [$($_.CapacityType)]: $($_.ReportCount) report(s)" -ForegroundColor White
    }
    $totalReports = ($allSummaries | ForEach-Object { $_.ReportCount } | Measure-Object -Sum).Sum
    Write-Host "  Total: $totalReports report(s) across $($workspaces.Count) workspace(s)" -ForegroundColor Green
}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    Disconnect-PowerBIServiceAccount
    Write-Host "Disconnected from Power BI." -ForegroundColor Gray
}
