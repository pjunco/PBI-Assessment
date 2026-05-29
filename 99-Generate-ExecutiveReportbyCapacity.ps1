<#
.SYNOPSIS
    Generates a self-contained HTML executive summary by capacity.

.DESCRIPTION
    Reads workspace, report inventory, and usage metrics outputs from the PBI Assessment
    pipeline and builds a capacity-level executive summary indicating:
      - Number of workspaces per capacity
      - Number of reports per capacity
      - Top 5 reports per capacity by unique viewers

.PARAMETER OutputFolder
    Path to the folder containing assessment JSON output files.
    Defaults to .\Output-BRA

.PARAMETER CustomerName
    Display name shown in the report header.
    Defaults to folder name.

.PARAMETER ReportPath
    Output path for the generated HTML file.
    Defaults to <OutputFolder>\Executive_Report_ByCapacity_<CustomerName>.html

.EXAMPLE
    .\99-Generate-ExecutiveReportbyCapacity.ps1 -OutputFolder .\Output -CustomerName "My Org"
#>

[CmdletBinding()]
param (
    [string]$OutputFolder = ".\Output",
    [string]$CustomerName = "My_Org",
    [string]$ReportPath   = ""
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Convert-ToInt {
    param([object]$Value)
    if ($null -eq $Value) { return 0 }
    $result = 0
    if ([int]::TryParse($Value.ToString(), [ref]$result)) { return $result }
    return 0
}

function New-CapacityKey {
    param([object]$Workspace)

    $capId   = if ($Workspace.PSObject.Properties['CapacityId']   -and $Workspace.CapacityId)   { $Workspace.CapacityId }   else { '' }
    $capName = if ($Workspace.PSObject.Properties['CapacityName'] -and $Workspace.CapacityName) { $Workspace.CapacityName } else { 'Unknown' }
    $capSku  = if ($Workspace.PSObject.Properties['CapacitySku']  -and $Workspace.CapacitySku)  { $Workspace.CapacitySku }  else { 'Unknown' }
    $capType = if ($Workspace.PSObject.Properties['CapacityType'] -and $Workspace.CapacityType) { $Workspace.CapacityType } else { 'Unknown' }

    if ($capId) {
        return $capId
    }

    return "$capName|$capSku|$capType"
}

function HtmlEncode {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Get-FirstPropertyValue {
  param(
    [object]$Object,
    [string[]]$PropertyNames
  )

  if ($null -eq $Object) { return $null }

  foreach ($prop in $PropertyNames) {
    $p = $Object.PSObject.Properties[$prop]
    if ($p -and $null -ne $p.Value) {
      $value = $p.Value.ToString().Trim()
      if ($value) {
        return $value
      }
    }
  }

  return $null
}

function Is-UnknownValue {
  param([string]$Value)
  if (-not $Value) { return $true }
  return $Value -match '^(unknown|n/a|null|none|-)$'
}

function Get-TargetFabricSku {
  param(
    [string]$CurrentSku,
    [string]$CurrentCapacityType
  )

  $sku = if ($CurrentSku) { $CurrentSku.Trim().ToUpperInvariant() } else { '' }
  $capType = if ($CurrentCapacityType) { $CurrentCapacityType.Trim() } else { '' }

  if ($capType -like '*Fabric*' -or $sku -match '^F\s*\d+') {
    return $sku
  }

  switch ($sku) {
    'P1' { return 'F64' }
    'P2' { return 'F128' }
    'P3' { return 'F256' }
    'P4' { return 'F512' }
    'P5' { return 'F1024' }
    default {
      if ($sku -match '^PP\d+$') {
        return 'Review required (PPU SKU)'
      }
      return 'Review required'
    }
  }
}

# 1) Resolve paths
$OutputFolder = Resolve-Path $OutputFolder -ErrorAction Stop

if (-not $CustomerName) {
    $CustomerName = Split-Path $OutputFolder -Leaf
}

if (-not $ReportPath) {
    $safeName = $CustomerName -replace '[\\/:*?"<>| ]', '_'
    $ReportPath = Join-Path $OutputFolder "Executive_Report_ByCapacity_$safeName.html"
}

$wsFile     = Join-Path $OutputFolder "PowerBI_Workspaces.json"
$reportsFile = Join-Path $OutputFolder "PowerBI_Reports_All_Workspaces.json"
$usageFile   = Join-Path $OutputFolder "PowerBI_UsageMetrics_All.json"

if (-not (Test-Path $wsFile))      { throw "File not found: $wsFile" }
if (-not (Test-Path $reportsFile)) { throw "File not found: $reportsFile" }

Write-Host ""
Write-Host "Generating executive capacity report for: $CustomerName" -ForegroundColor Cyan
Write-Host "Input folder : $OutputFolder" -ForegroundColor Gray

# 2) Parse data
$workspaces = @(Get-Content $wsFile -Raw | ConvertFrom-Json)
$reportsByWorkspace = @(Get-Content $reportsFile -Raw | ConvertFrom-Json)

# Build workspace lookup for reliable capacity metadata fallback.
$workspaceLookupById = @{}
$workspaceLookupByName = @{}
$capacityLookupById = @{}
foreach ($wsInv in $workspaces) {
  $invId = Get-FirstPropertyValue -Object $wsInv -PropertyNames @('Id', 'WorkspaceId', 'id', 'workspaceId')
  if ($invId) {
    $workspaceLookupById[$invId.ToLowerInvariant()] = $wsInv
  }

  $invName = Get-FirstPropertyValue -Object $wsInv -PropertyNames @('Name', 'WorkspaceName', 'name', 'workspaceName')
  if ($invName) {
    $workspaceLookupByName[$invName.ToLowerInvariant()] = $wsInv
  }

  $invCapId = Get-FirstPropertyValue -Object $wsInv -PropertyNames @('CapacityId', 'capacityId')
  if ($invCapId -and -not $capacityLookupById.ContainsKey($invCapId.ToLowerInvariant())) {
    $capacityLookupById[$invCapId.ToLowerInvariant()] = [PSCustomObject]@{
      CapacityId   = $invCapId
      CapacityName = Get-FirstPropertyValue -Object $wsInv -PropertyNames @('CapacityName', 'capacityName')
      CapacitySku  = Get-FirstPropertyValue -Object $wsInv -PropertyNames @('CapacitySku', 'capacitySku')
      CapacityType = Get-FirstPropertyValue -Object $wsInv -PropertyNames @('CapacityType', 'capacityType')
    }
  }
}

$usageByWorkspace = @()
$usageAvailable = Test-Path $usageFile
if ($usageAvailable) {
    $usageByWorkspace = @(Get-Content $usageFile -Raw | ConvertFrom-Json)
} else {
    Write-Warning "Usage file not found: $usageFile"
    Write-Warning "Top 5 reports will be generated with 0 unique viewers. Run 03-Get-UsageMetricByReport.ps1 first for real usage values."
}

# 3) Build usage index by ReportId
$usageByReportId = @{}
foreach ($wsUsage in $usageByWorkspace) {
    foreach ($r in @($wsUsage.Reports)) {
        if (-not $r.ReportId) { continue }

        $rid = $r.ReportId.ToString().ToLowerInvariant()
        $uniqueViewers = Convert-ToInt $r.TotalUniqueViewers

        if ($usageByReportId.ContainsKey($rid)) {
            # Keep the highest value in case of duplicate report entries.
            if ($uniqueViewers -gt $usageByReportId[$rid]) {
                $usageByReportId[$rid] = $uniqueViewers
            }
        } else {
            $usageByReportId[$rid] = $uniqueViewers
        }
    }
}

# 4) Aggregate by capacity
$capacityMap = @{}

foreach ($ws in $reportsByWorkspace) {
  $wsIdRaw = Get-FirstPropertyValue -Object $ws -PropertyNames @('WorkspaceId', 'Id', 'workspaceId', 'id')
  $wsNameRaw = Get-FirstPropertyValue -Object $ws -PropertyNames @('WorkspaceName', 'Name', 'workspaceName', 'name')
  $wsIdKey = if ($wsIdRaw) { $wsIdRaw.ToLowerInvariant() } else { $null }
  $wsNameKey = if ($wsNameRaw) { $wsNameRaw.ToLowerInvariant() } else { $null }

  $wsInventory = $null
  if ($wsIdKey -and $workspaceLookupById.ContainsKey($wsIdKey)) {
    $wsInventory = $workspaceLookupById[$wsIdKey]
  } elseif ($wsNameKey -and $workspaceLookupByName.ContainsKey($wsNameKey)) {
    $wsInventory = $workspaceLookupByName[$wsNameKey]
  }

  $capId = Get-FirstPropertyValue -Object $ws -PropertyNames @('CapacityId', 'capacityId')
  if ((-not $capId) -and $wsInventory) {
    $capId = Get-FirstPropertyValue -Object $wsInventory -PropertyNames @('CapacityId', 'capacityId')
  }

  $capInventory = $null
  if ($capId) {
    $capIdKey = $capId.ToLowerInvariant()
    if ($capacityLookupById.ContainsKey($capIdKey)) {
      $capInventory = $capacityLookupById[$capIdKey]
    }
  }

  $capName = Get-FirstPropertyValue -Object $ws -PropertyNames @('CapacityName', 'capacityName')
  if ((Is-UnknownValue $capName) -and $wsInventory) {
    $capName = Get-FirstPropertyValue -Object $wsInventory -PropertyNames @('CapacityName', 'capacityName')
  }
  if ((Is-UnknownValue $capName) -and $capInventory) {
    $capName = Get-FirstPropertyValue -Object $capInventory -PropertyNames @('CapacityName', 'capacityName')
  }

  $capSku = Get-FirstPropertyValue -Object $ws -PropertyNames @('CapacitySku', 'capacitySku')
  if ((Is-UnknownValue $capSku) -and $wsInventory) {
    $capSku = Get-FirstPropertyValue -Object $wsInventory -PropertyNames @('CapacitySku', 'capacitySku')
  }
  if ((Is-UnknownValue $capSku) -and $capInventory) {
    $capSku = Get-FirstPropertyValue -Object $capInventory -PropertyNames @('CapacitySku', 'capacitySku')
  }

  $capType = Get-FirstPropertyValue -Object $ws -PropertyNames @('CapacityType', 'capacityType')
  if ((Is-UnknownValue $capType) -and $wsInventory) {
    $capType = Get-FirstPropertyValue -Object $wsInventory -PropertyNames @('CapacityType', 'capacityType')
  }
  if ((Is-UnknownValue $capType) -and $capInventory) {
    $capType = Get-FirstPropertyValue -Object $capInventory -PropertyNames @('CapacityType', 'capacityType')
  }

  if (Is-UnknownValue $capName) {
    $capName = if ($capId) { "Capacity $($capId.Substring(0, [Math]::Min(8, $capId.Length)))" } else { 'Unknown' }
  }
  if (Is-UnknownValue $capSku) {
    $capSku = 'Unknown'
  }
  if (Is-UnknownValue $capType) {
    $capType = 'Unknown'
  }

  $capacityRef = [PSCustomObject]@{
    CapacityId   = $capId
    CapacityName = $capName
    CapacitySku  = $capSku
    CapacityType = $capType
  }

  $capacityKey = New-CapacityKey $capacityRef

    if (-not $capacityMap.ContainsKey($capacityKey)) {
        $capacityMap[$capacityKey] = [ordered]@{
            CapacityName = $capName
            CapacitySku  = $capSku
            CapacityType = $capType
            Workspaces   = @{}
            Reports      = @{}
        }
    }

    $cap = $capacityMap[$capacityKey]

    $wsId = if ($wsIdRaw) { $wsIdRaw } else { $wsNameRaw }
    if ($wsId) {
      $cap.Workspaces[$wsId] = if ($wsNameRaw) { $wsNameRaw } else { 'Unknown Workspace' }
    }

    foreach ($report in @($ws.Reports)) {
        if (-not $report.ReportId) { continue }

        $reportId = $report.ReportId.ToString().ToLowerInvariant()
        $reportName = if ($report.ReportName) { $report.ReportName } else { 'Unnamed Report' }
        $connectionMode = if ($report.PSObject.Properties['ConnectionMode'] -and $report.ConnectionMode) { $report.ConnectionMode } else { 'Unknown' }
        $uniqueViews = if ($usageByReportId.ContainsKey($reportId)) { $usageByReportId[$reportId] } else { 0 }

        if ($cap.Reports.ContainsKey($reportId)) {
            if ($uniqueViews -gt $cap.Reports[$reportId].UniqueViews) {
                $cap.Reports[$reportId].UniqueViews = $uniqueViews
            }
          if ((-not $cap.Reports[$reportId].ConnectionMode -or $cap.Reports[$reportId].ConnectionMode -eq 'Unknown') -and $connectionMode -ne 'Unknown') {
            $cap.Reports[$reportId].ConnectionMode = $connectionMode
          }
        } else {
            $cap.Reports[$reportId] = [ordered]@{
                ReportId     = $reportId
                ReportName   = $reportName
                WorkspaceName = if ($wsNameRaw) { $wsNameRaw } else { 'Unknown Workspace' }
            ConnectionMode = $connectionMode
                UniqueViews  = $uniqueViews
            }
        }
    }
}

$capacitySummary = @(
    $capacityMap.Values |
        ForEach-Object {
            $reports = @($_.Reports.Values)
          $totalUniqueViews = if ($reports.Count -gt 0) {
            ($reports | ForEach-Object { Convert-ToInt $_.UniqueViews } | Measure-Object -Sum).Sum
          } else {
            0
          }

          $importCount = @(
            $reports | Where-Object {
              $_.ConnectionMode -and $_.ConnectionMode.ToString().Trim().ToLowerInvariant() -eq 'import'
            }
          ).Count

          $directQueryCount = @(
            $reports | Where-Object {
              $mode = if ($_.ConnectionMode) { $_.ConnectionMode.ToString().Trim().ToLowerInvariant() } else { '' }
              $mode -like 'directquery*' -or $mode -like 'direct*' -or $mode -like '*live*'
            }
          ).Count

          $totalReportsInCap = $reports.Count
          $importPct = if ($totalReportsInCap -gt 0) { [Math]::Round(($importCount / $totalReportsInCap) * 100, 1) } else { 0 }
          $directQueryPct = if ($totalReportsInCap -gt 0) { [Math]::Round(($directQueryCount / $totalReportsInCap) * 100, 1) } else { 0 }

          $importText = "{0:N1}% ({1:N0})" -f $importPct, $importCount
          $directQueryText = "{0:N1}% ({1:N0})" -f $directQueryPct, $directQueryCount

            $top5 = @(
                $reports |
                    Sort-Object -Property @{Expression = { $_.UniqueViews }; Descending = $true}, @{Expression = { $_.ReportName }; Descending = $false} |
                    Select-Object -First 5
            )

            [PSCustomObject]@{
                CapacityName   = $_.CapacityName
                CapacitySku    = $_.CapacitySku
                CapacityType   = $_.CapacityType
                WorkspaceCount = $_.Workspaces.Count
                ReportCount    = $_.Reports.Count
                TopReports     = $top5
                TotalUniqueViews = $totalUniqueViews
                ImportModeSummary = $importText
                DirectQuerySummary = $directQueryText
            }
        } |
        Sort-Object -Property @{Expression = { $_.ReportCount }; Descending = $true}, @{Expression = { $_.WorkspaceCount }; Descending = $true}, @{Expression = { $_.CapacityName }; Descending = $false}
)

$totalCapacities = $capacitySummary.Count
$totalWorkspaces = @($workspaces).Count
$totalReports = @($reportsByWorkspace | ForEach-Object { @($_.Reports).Count } | Measure-Object -Sum).Sum

$zeroReportCapacities = @($capacitySummary | Where-Object { (Convert-ToInt $_.ReportCount) -eq 0 })
$warningsHtml = ''
if ($zeroReportCapacities.Count -gt 0) {
  $capNames = $zeroReportCapacities |
    ForEach-Object {
      if ($_.CapacityName) {
        HtmlEncode $_.CapacityName
      } else {
        'Unknown'
      }
    }

  $warningsHtml = @"
  <section class="warning-box" role="status" aria-live="polite">
    <strong>Warning:</strong> $(('{0:N0}' -f $zeroReportCapacities.Count)) capacit$(if ($zeroReportCapacities.Count -eq 1) { 'y has' } else { 'ies have' }) 0 reports in the current inventory.<br />
    <span class="warning-detail">Affected: $($capNames -join ', ')</span>
  </section>
"@
}

$generatedOn = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# 5) Build HTML rows
$capacityRows = foreach ($cap in $capacitySummary) {
    $top5Html = if (@($cap.TopReports).Count -gt 0) {
        $items = $cap.TopReports | ForEach-Object {
            "<li><span class='report-name'>$(HtmlEncode $_.ReportName)</span> <span class='views'>$(('{0:N0}' -f $_.UniqueViews)) unique views</span></li>"
        }
        "<ol class='top-reports'>" + ($items -join "") + "</ol>"
    } else {
        "<span class='muted'>No reports</span>"
    }

    @"
    <tr>
      <td>$(HtmlEncode $cap.CapacityName)</td>
      <td>$(HtmlEncode $cap.CapacitySku)</td>
      <td>$(HtmlEncode $cap.CapacityType)</td>
      <td class='num'>$(('{0:N0}' -f $cap.WorkspaceCount))</td>
      <td class='num'>$(('{0:N0}' -f $cap.ReportCount))</td>
      <td class='num'>$(HtmlEncode $cap.ImportModeSummary)</td>
      <td class='num'>$(HtmlEncode $cap.DirectQuerySummary)</td>
      <td class='top5-cell'>$top5Html</td>
    </tr>
"@
}

$capacityRowsHtml = if (@($capacityRows).Count -gt 0) {
    $capacityRows -join "`n"
} else {
    @"
    <tr>
      <td colspan='8' class='muted'>No capacity data found in input files.</td>
    </tr>
"@
}

$translationRows = foreach ($cap in $capacitySummary) {
  $targetSku = Get-TargetFabricSku -CurrentSku $cap.CapacitySku -CurrentCapacityType $cap.CapacityType

  @"
    <tr>
      <td>$(HtmlEncode $cap.CapacityName)</td>
      <td>$(HtmlEncode $cap.CapacitySku)</td>
      <td>$(HtmlEncode $cap.CapacityType)</td>
      <td><strong>$(HtmlEncode $targetSku)</strong></td>
    </tr>
"@
}

$translationRowsHtml = if (@($translationRows).Count -gt 0) {
  $translationRows -join "`n"
} else {
  @"
    <tr>
      <td colspan='4' class='muted'>No capacity data found in input files.</td>
    </tr>
"@
}

# 6) Generate HTML output
$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>$([System.Net.WebUtility]::HtmlEncode($CustomerName)) - Executive Report by Capacity</title>
  <style>
    :root {
      --bg: #f3f6fb;
      --surface: #ffffff;
      --text: #1f2937;
      --muted: #6b7280;
      --border: #dbe4ee;
      --brand: #005ea8;
      --accent: #0ea5e9;
      --chip: #eef6ff;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
      color: var(--text);
      background:
        radial-gradient(circle at top left, #dcecff 0%, transparent 35%),
        radial-gradient(circle at top right, #e6fff7 0%, transparent 30%),
        var(--bg);
    }
    .container {
      max-width: 1280px;
      margin: 24px auto;
      padding: 0 16px 32px;
    }
    .header {
      background: linear-gradient(120deg, var(--brand), #0b7cc4);
      color: #fff;
      border-radius: 14px;
      padding: 22px;
      box-shadow: 0 10px 24px rgba(0, 42, 90, 0.22);
    }
    .header h1 {
      margin: 0 0 8px;
      font-size: 1.55rem;
    }
    .meta {
      margin: 0;
      font-size: 0.95rem;
      opacity: 0.95;
    }
    .kpis {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
      gap: 12px;
      margin-top: 16px;
    }
    .kpi {
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 14px;
    }
    .kpi .label {
      color: var(--muted);
      font-size: 0.83rem;
      text-transform: uppercase;
      letter-spacing: 0.04em;
    }
    .kpi .value {
      margin-top: 6px;
      font-size: 1.5rem;
      font-weight: 700;
      color: var(--brand);
    }
    .card {
      margin-top: 16px;
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 12px;
      overflow: hidden;
    }
    .warning-box {
      margin-top: 16px;
      padding: 12px 14px;
      border: 1px solid #f5d08a;
      border-radius: 10px;
      background: #fff8e7;
      color: #7a4f01;
      font-size: 0.92rem;
      line-height: 1.4;
    }
    .warning-detail {
      color: #8a5b07;
      font-size: 0.88rem;
    }
    .card h2 {
      margin: 0;
      padding: 14px 16px;
      border-bottom: 1px solid var(--border);
      font-size: 1.05rem;
      background: #f9fbfe;
    }
    .table-wrap {
      overflow-x: auto;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      min-width: 980px;
    }
    thead th {
      text-align: left;
      font-size: 0.82rem;
      text-transform: uppercase;
      letter-spacing: 0.03em;
      color: #4b5563;
      background: #f7fbff;
      border-bottom: 1px solid var(--border);
      padding: 12px;
    }
    tbody td {
      padding: 12px;
      border-bottom: 1px solid #edf2f7;
      vertical-align: top;
    }
    tbody tr:hover {
      background: #fbfdff;
    }
    .num {
      text-align: right;
      font-variant-numeric: tabular-nums;
      font-weight: 600;
      color: #0f172a;
      white-space: nowrap;
    }
    .top5-cell {
      min-width: 360px;
    }
    .top-reports {
      margin: 0;
      padding-left: 18px;
    }
    .top-reports li {
      margin: 0 0 6px;
      line-height: 1.3;
    }
    .report-name {
      font-weight: 600;
    }
    .views {
      margin-left: 6px;
      padding: 2px 8px;
      background: var(--chip);
      border: 1px solid #d8ecff;
      border-radius: 999px;
      color: #075985;
      font-size: 0.8rem;
      white-space: nowrap;
    }
    .muted {
      color: var(--muted);
    }
    .section-text {
      margin: 14px 16px;
      color: #334155;
      line-height: 1.5;
      font-size: 0.94rem;
    }
    .note-box {
      margin: 12px 16px 16px;
      padding: 12px 14px;
      border: 1px solid #cfe6ff;
      border-radius: 10px;
      background: #f3f9ff;
      color: #0b3f6f;
      font-size: 0.92rem;
      line-height: 1.45;
    }
    .footer {
      margin-top: 12px;
      font-size: 0.82rem;
      color: var(--muted);
    }
    @media (max-width: 760px) {
      .container { margin: 14px auto; }
      .header { padding: 16px; }
      .header h1 { font-size: 1.2rem; }
      .kpi .value { font-size: 1.2rem; }
    }
  </style>
</head>
<body>
  <div class="container">
    <section class="header">
      <h1>Power BI Executive Summary by Capacity</h1>
      <p class="meta"><strong>Customer:</strong> $(HtmlEncode $CustomerName)</p>
      <p class="meta"><strong>Generated on:</strong> $generatedOn</p>
    </section>

    <section class="kpis">
      <article class="kpi">
        <div class="label">Capacities</div>
        <div class="value">$(('{0:N0}' -f $totalCapacities))</div>
      </article>
      <article class="kpi">
        <div class="label">Workspaces</div>
        <div class="value">$(('{0:N0}' -f $totalWorkspaces))</div>
      </article>
      <article class="kpi">
        <div class="label">Reports</div>
        <div class="value">$(('{0:N0}' -f $totalReports))</div>
      </article>
      <article class="kpi">
        <div class="label">Usage Source</div>
        <div class="value">$(if ($usageAvailable) { 'Available' } else { 'Missing' })</div>
      </article>
    </section>

$warningsHtml

    <section class="card">
      <h2>Capacity Summary (Workspaces, Reports, Top 5 by Unique Views)</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Capacity</th>
              <th>SKU</th>
              <th>Type</th>
              <th>Workspaces</th>
              <th>Reports</th>
              <th>Import Mode</th>
              <th>DirectQuery/Live</th>
              <th>Top 5 Reports (Unique Views)</th>
            </tr>
          </thead>
          <tbody>
$capacityRowsHtml
          </tbody>
        </table>
      </div>
    </section>

    <section class="card">
      <h2>Translating Power BI Premium to Microsoft Fabric SKUs</h2>
      <p class="section-text">
        Power BI Premium per capacity SKUs (P-SKUs) are being retired and organizations like yours should transition to a Microsoft Fabric SKU at the time of your next renewal. The Power BI Premium product capabilities will not change and there is no immediate action required. When Microsoft Fabric was launched, existing Power BI Premium customers were enabled to turn on Fabric through the Power BI admin portal and use their existing capacity to power Fabric capabilities.
      </p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Capacity</th>
              <th>Current SKU</th>
              <th>Current Capacity Type</th>
              <th>Target F-SKU</th>
            </tr>
          </thead>
          <tbody>
$translationRowsHtml
          </tbody>
        </table>
      </div>
      <p class="note-box">
        If you are familiar with PBI vCores, 8 CUs provide the same compute power as 1 Power BI Premium v-core for Power BI workloads. That is, F64 provides the equivalent compute power of Power BI Premium P1.
      </p>
    </section>

    <p class="footer">Inputs: PowerBI_Workspaces.json, PowerBI_Reports_All_Workspaces.json$(if ($usageAvailable) { ', PowerBI_UsageMetrics_All.json' } else { '' }).</p>
  </div>
</body>
</html>
"@

$html | Out-File -LiteralPath $ReportPath -Encoding UTF8

Write-Host ""
Write-Host "Report generated successfully:" -ForegroundColor Green
Write-Host "  $ReportPath" -ForegroundColor White
