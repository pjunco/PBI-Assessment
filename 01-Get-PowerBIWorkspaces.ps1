#Requires -Modules MicrosoftPowerBIMgmt

<#
.SYNOPSIS
    Retrieves only Premium or Fabric capacity workspaces from a Power BI tenant.

.DESCRIPTION
    Authenticates to Power BI using a service principal or interactive login,
    fetches all tenant capacities to determine their SKU type (Premium P-SKU vs
    Fabric F-SKU), then exports only dedicated-capacity workspaces with a
    CapacityType property ("PBI Premium Capacity" or "Fabric Capacity") to JSON.

.PARAMETER TenantId
    Azure AD Tenant ID. Required for service principal authentication.

.PARAMETER ClientId
    Service principal (App Registration) Client ID.

.PARAMETER ClientSecret
    Service principal Client Secret (as a SecureString).

.PARAMETER OutputJson
    Path to export results as JSON. Defaults to .\Output\PowerBI_Workspaces.json.

.EXAMPLE
    # Interactive login (browser / MFA)
    .\01-Get-PowerBIWorkspaces.ps1

.EXAMPLE
    # Service principal login
    $secret = ConvertTo-SecureString "your-secret" -AsPlainText -Force
    .\01-Get-PowerBIWorkspaces.ps1 -TenantId "xxx" -ClientId "yyy" -ClientSecret $secret
#>

[CmdletBinding()]
param (
    [string]$TenantId,
    [string]$ClientId,
    [System.Security.SecureString]$ClientSecret,
    [string]$OutputJson = ".\Output\PowerBI_Workspaces.json"
)

# ── 0. Pre-requisites notice ────────────────────────────────────────────────────
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Host ""
Write-Host "---------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host " PRE-REQUISITES" -ForegroundColor Yellow
Write-Host "---------------------------------------------------------------------" -ForegroundColor Yellow
Write-Host " The service principal needs:" -ForegroundColor White
Write-Host "  1. Power BI Service API permission:" -ForegroundColor White
Write-Host "       Tenant.Read.All  (or Tenant.ReadWrite.All)" -ForegroundColor Cyan
Write-Host "     granted by an Azure AD admin (admin consent required)." -ForegroundColor White
Write-Host "  2. Added in the Power BI Admin portal:" -ForegroundColor White
Write-Host "       Admin portal > Tenant settings > Developer settings" -ForegroundColor Cyan
Write-Host "       > Allow service principals to use read-only Power BI admin APIs" -ForegroundColor Cyan
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

# ── 3. Retrieve all workspaces and capacities ──────────────────────────────────
Write-Host "Fetching all workspaces from tenant..." -ForegroundColor Cyan

try {
    # Get-PowerBIWorkspace -Scope Organization requires Power BI Admin role
    $allWorkspaces = Get-PowerBIWorkspace -Scope Organization -All -Include All

    # Filter to dedicated-capacity workspaces only (exclude personal workspaces)
    $workspaces = $allWorkspaces | Where-Object { $_.IsOnDedicatedCapacity -eq $true -and $_.Type -ne 'PersonalGroup' }
    Write-Host "Found $($workspaces.Count) dedicated-capacity workspace(s) (out of $($allWorkspaces.Count) total, personal workspaces excluded)." -ForegroundColor Green

    # ── 4. Fetch capacities to determine SKU type ───────────────────────────────
    Write-Host "Fetching capacity SKUs..." -ForegroundColor Cyan
    $capacities = Get-PowerBICapacity -Scope Organization
    # Build lookups: CapacityId -> CapacityType label, CapacityName and CapacitySku
    $capacityTypeMap = @{}
    $capacityNameMap = @{}
    $capacitySkuMap  = @{}
    foreach ($cap in $capacities) {
        $sku = $cap.Sku
        $key = $cap.Id.ToString().ToUpper()
        $capacityNameMap[$key] = $cap.DisplayName
        $capacitySkuMap[$key]  = $sku
        if ($sku -match '^F') {
            $capacityTypeMap[$key] = 'Fabric Capacity'
        } else {
            $capacityTypeMap[$key] = 'PBI Premium Capacity'
        }
    }

    # ── 5. Enrich workspaces with CapacityType, ReportCount and DefaultDatasetStorageFormat ─
    Write-Host "Fetching report counts and storage format per workspace..." -ForegroundColor Cyan
    $enriched = foreach ($ws in $workspaces) {
        $capKey  = $ws.CapacityId.ToString().ToUpper()
        $capType = if ($capacityTypeMap.ContainsKey($capKey)) { $capacityTypeMap[$capKey] } else { 'PBI Premium Capacity' }
        $capName = if ($capacityNameMap.ContainsKey($capKey)) { $capacityNameMap[$capKey] } else { 'Unknown' }
        $capSku  = if ($capacitySkuMap.ContainsKey($capKey))  { $capacitySkuMap[$capKey]  } else { 'Unknown' }

        # Retry Get-PowerBIReport on 429
        $reportCount = 0
        $maxRetries = 5
        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                $reportCount = (Get-PowerBIReport -WorkspaceId $ws.Id -Scope Organization -ErrorAction Stop).Count
                break
            } catch {
                if ($_ -match '429' -and $attempt -lt $maxRetries) {
                    $wait = $attempt * 5
                    Write-Warning "  429 on '$($ws.Name)' (attempt $attempt). Waiting ${wait}s..."
                    Start-Sleep -Seconds $wait
                } else {
                    Write-Warning "  Could not fetch reports for '$($ws.Name)': $_"
                    break
                }
            }
        }
        Start-Sleep -Milliseconds 1000

        # Fetch workspace-level Large Semantic Model storage format setting
        # defaultDatasetStorageFormat: "Small" = standard (default), "Large" = LSM enabled
        $defaultStorageFormat = 'Small'
        try {
            $wsDetail = Invoke-PowerBIRestMethod -Url "admin/groups/$($ws.Id)" -Method Get | ConvertFrom-Json
            if ($wsDetail.PSObject.Properties['defaultDatasetStorageFormat']) {
                $defaultStorageFormat = $wsDetail.defaultDatasetStorageFormat
            }
        } catch {
            if ($_ -match '429') {
                Start-Sleep -Seconds 10
                try {
                    $wsDetail = Invoke-PowerBIRestMethod -Url "admin/groups/$($ws.Id)" -Method Get | ConvertFrom-Json
                    if ($wsDetail.PSObject.Properties['defaultDatasetStorageFormat']) {
                        $defaultStorageFormat = $wsDetail.defaultDatasetStorageFormat
                    }
                } catch {
                    Write-Warning "  Could not fetch storage format for '$($ws.Name)': $_"
                }
            } else {
                Write-Warning "  Could not fetch storage format for '$($ws.Name)': $_"
            }
        }

        Write-Host "  $($ws.Name): $reportCount report(s) | $capType '$capName' ($capSku) | StorageFormat: $defaultStorageFormat" -ForegroundColor Gray
        $ws | Select-Object Id, Name, Type, State, IsOnDedicatedCapacity, CapacityId,
            @{N='CapacityName';               E={ $capName }},
            @{N='CapacitySku';                E={ $capSku }},
            @{N='CapacityType';               E={ $capType }},
            @{N='ReportCount';                E={ $reportCount }},
            @{N='DefaultDatasetStorageFormat'; E={ $defaultStorageFormat }},
            IsReadOnly
    }

    # ── 6. Display summary ──────────────────────────────────────────────────────
    $enriched | Format-Table Id, Name, CapacityName, CapacitySku, CapacityType, ReportCount, CapacityId -AutoSize

    # ── 7. Export to JSON ────────────────────────────────────────────────────────
    $outputDir = Split-Path $OutputJson -Parent
    if ($outputDir -and -not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    $enriched | ForEach-Object {
        $ws = $_
        $users = ($allWorkspaces | Where-Object { $_.Id -eq $ws.Id }).Users |
                 Select-Object UserPrincipalName, GroupUserAccessRight
        $ws | Select-Object *, @{N='Users'; E={ $users }}
    } |
        ConvertTo-Json -Depth 5 |
        Out-File -FilePath $OutputJson -Encoding UTF8

    Write-Host "Results exported to: $OutputJson" -ForegroundColor Green
}
catch {
    Write-Error "Failed to retrieve workspaces: $_"
}
finally {
    Disconnect-PowerBIServiceAccount
    Write-Host "Disconnected from Power BI." -ForegroundColor Gray
}
