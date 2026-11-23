<#
.SYNOPSIS
Creates a draft access-package workbook from an access-review discovery export.
.DESCRIPTION
This is a read-only preview. Complete package names and approval stages before
using the package creation script. Shared or uncertain resources require manual
review and are not placed on the first worksheet.
#>
[CmdletBinding()]
param(
	[Parameter(Mandatory)][string]$DiscoveryPath,
	[Parameter(Mandatory)][string]$WorkbookPath
)

$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Groups
Import-Module ImportExcel -WarningAction SilentlyContinue
. (Join-Path $PSScriptRoot 'AccessPackage.Group.ps1')
. (Join-Path $PSScriptRoot 'AccessPackage.Coverage.ps1')
. (Join-Path $PSScriptRoot 'AccessPackage.Discovery.ps1')

$inputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DiscoveryPath)
$outputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($WorkbookPath)
if ([IO.Path]::GetExtension($outputPath) -ne '.xlsx') { throw 'WorkbookPath must end in .xlsx.' }
if ([IO.File]::Exists($outputPath) -or [IO.Directory]::Exists($outputPath)) { throw "Workbook output already exists: '$outputPath'." }
if (-not [IO.Directory]::Exists([IO.Path]::GetDirectoryName($outputPath))) { throw 'Workbook output directory does not exist.' }
$data = Get-Content -LiteralPath $inputPath -Raw | ConvertFrom-Json
$discovery = Get-PackageDiscoveryCandidates -Data $data
$applicationIds = @($discovery.ServicePrincipalIds)
$workbookVersion = if ($applicationIds.Count -gt 1) { 2 } else { 1 }
$groups = $discovery.Groups
$candidates = $discovery.Candidates
try { $inventory = @(Get-PackageCoverageInventory); $inventoryError = $null }
catch { $inventory = @(); $inventoryError = $_.Exception.Message }
$eligibility = @{}
foreach ($candidate in $candidates) {
	$reasons = [System.Collections.Generic.List[string]]::new()
	if ($candidate.Reason) { $reasons.Add($candidate.Reason) }
	if ($inventoryError) { $reasons.Add("Cannot establish package coverage: $inventoryError") }
	else {
		try {
			$coverage = Get-PackageCoverageResult -GroupIds $candidate.GroupIds -Inventory $inventory
			if ($coverage.Status -ne 'Ready') { $candidate.Status = $coverage.Status; $reasons.Add($coverage.Reason) }
		}
		catch { $reasons.Add("Cannot establish package activity: $($_.Exception.Message)") }
	}
	foreach ($id in $candidate.GroupIds) {
		if (-not $eligibility.ContainsKey($id)) {
			try { $null = Get-PackageGroup -GroupId $id; $eligibility[$id] = '' }
			catch { $eligibility[$id] = $_.Exception.Message }
		}
		if ($eligibility[$id]) { $reasons.Add($eligibility[$id]) }
	}
	if ($reasons.Count -and $candidate.Status -ne 'Skipped') { $candidate.Status = 'Manual review' }
	$candidate.Reason = $reasons -join ' '
}

$temporaryPath = Join-Path ([IO.Path]::GetDirectoryName($outputPath)) ('.discovery-' + [guid]::NewGuid().ToString() + '.xlsx')
$excel = $null
try {
	$excel = Open-ExcelPackage -Path $temporaryPath -Create
	$sheet = $excel.Workbook.Worksheets.Add('Packages')
	$blocked = $excel.Workbook.Worksheets.Add('Manual review')
	$topology = $excel.Workbook.Worksheets.Add('Topology')
	$metadata = $excel.Workbook.Worksheets.Add('Discovery metadata')
	$metadata.Cells[1, 1].Value = 'Discovery workbook version'
	$metadata.Cells[1, 2].Value = "$workbookVersion"
	for ($i = 0; $i -lt $applicationIds.Count; $i++) {
		$metadata.Cells[($i + 2), 1].Value = 'Service principal ID'
		$metadata.Cells[($i + 2), 2].Value = "$($applicationIds[$i])"
	}
	$columns = @('Role', 'Brand', 'Region', 'Potential Teams', 'GroupId', 'Description', 'Access Package Name',
		'Approver 1 List', 'Approver 2', 'Approver 3', 'Discovery ServicePrincipalId', 'Discovery RootGroupId', 'Action required')
	for ($i = 0; $i -lt $columns.Count; $i++) { $sheet.Cells[1, ($i + 1)].Value = $columns[$i] }
	$blockedColumns = @('Status', 'Reason', 'Root group', 'Root group ID', 'Group IDs', 'Group names')
	for ($i = 0; $i -lt $blockedColumns.Count; $i++) { $blocked.Cells[1, ($i + 1)].Value = $blockedColumns[$i] }
	$topologyColumns = @('Root group ID', 'Root group', 'Group ID', 'Group', 'Parent group IDs', 'Direct user count')
	for ($i = 0; $i -lt $topologyColumns.Count; $i++) { $topology.Cells[1, ($i + 1)].Value = $topologyColumns[$i] }
	$sheetRow = 2
	$blockedRow = 2
	$topologyRow = 2
	foreach ($candidate in $candidates) {
		$rootName = [string]$groups[$candidate.Root].DisplayName
		$appName = [string]$candidate.ServicePrincipalName
		$name = "$appName - $rootName".Trim(' ', '-')
		$description = "Application: $appName. $(Get-DiscoveryPackageWarning)"
		if ($candidate.Status -eq 'Ready') {
			$action = 'Complete approvers and verify business information.'
			if (-not $name -or $name -match '(?i)(?<![a-z0-9])(?:access[\s_-]*package|ap|sg)(?![a-z0-9])') {
				$action = "Replace the suggested package name; it violates naming rules. $action"
			}
			$values = @('', '', '', '', ($candidate.GroupIds -join ';'), $description, $name, '', '', '', $candidate.ServicePrincipalId, $candidate.Root, $action)
			for ($i = 0; $i -lt $values.Count; $i++) { $sheet.Cells[$sheetRow, ($i + 1)].Value = [string]$values[$i] }
			$sheetRow++
		}
		else {
			$values = @($candidate.Status, $candidate.Reason, $rootName, $candidate.Root,
				($candidate.GroupIds -join ';'), (($candidate.GroupIds | ForEach-Object { $groups[$_].DisplayName }) -join '; '))
			for ($i = 0; $i -lt $values.Count; $i++) { $blocked.Cells[$blockedRow, ($i + 1)].Value = [string]$values[$i] }
			$blockedRow++
		}
		foreach ($id in $candidate.GroupIds) {
			$parents = @($candidate.GroupIds | Where-Object { $groups[$_].ChildIds -contains $id })
			$values = @($candidate.Root, $rootName, $id, $groups[$id].DisplayName, ($parents -join ';'), $groups[$id].DirectUserCount)
			for ($i = 0; $i -lt $values.Count; $i++) { $topology.Cells[$topologyRow, ($i + 1)].Value = [string]$values[$i] }
			$topologyRow++
		}
	}
	Close-ExcelPackage -ExcelPackage $excel
	$excel = $null
	[IO.File]::Move($temporaryPath, $outputPath)
}
finally {
	if ($excel) { Close-ExcelPackage -ExcelPackage $excel -NoSave }
	if ([IO.File]::Exists($temporaryPath)) { [IO.File]::Delete($temporaryPath) }
}
Write-Host "Draft workbook: $outputPath. Complete names and approvals; blocked rows are on the Manual review sheet."
