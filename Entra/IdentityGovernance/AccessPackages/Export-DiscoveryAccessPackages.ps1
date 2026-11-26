<#
.SYNOPSIS
Creates a draft access-package workbook from an access-review discovery export.
.DESCRIPTION
This is a read-only preview. Complete package names and approval stages before
using the package creation script. Uncertain resources require manual review
and are not placed on the first worksheet.
.PARAMETER OnePackagePerGroup
Propose one package per unique discovered group instead of one per root and
its descendants. Does not read or list the groups' application assignments.
#>
[CmdletBinding()]
param(
	[Parameter(Mandatory)][string]$DiscoveryPath,
	[Parameter(Mandatory)][string]$WorkbookPath,
	[switch]$OnePackagePerGroup
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
$discovery = Get-PackageDiscoveryCandidates -Data $data -AllowSharedGroups
$applicationIds = @($discovery.ServicePrincipalIds)
$workbookVersion = if ($OnePackagePerGroup) { 4 } else { 5 }
$groups = $discovery.Groups
$candidates = if ($OnePackagePerGroup) {
	$seenGroups = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
	@(
		foreach ($source in $discovery.Candidates) {
			foreach ($id in $source.GroupIds) {
				if ($seenGroups.Add($id)) {
					[pscustomobject]@{
						ServicePrincipalId = $source.ServicePrincipalId; ServicePrincipalName = $source.ServicePrincipalName
						Root = $id; GroupIds = @($id); Status = 'Ready'; Reason = ''; Description = ''
					}
				}
			}
		}
	)
}
else { $discovery.Candidates }
try { $inventory = @(Get-PackageCoverageInventory); $inventoryError = $null }
catch { $inventory = @(); $inventoryError = $_.Exception.Message }
$eligibility = @{}
$applicationCache = @{}
foreach ($candidate in $candidates) {
	$reasons = [System.Collections.Generic.List[string]]::new()
	if ($candidate.Reason) { $reasons.Add($candidate.Reason) }
	$resourceGroupIds = if ($OnePackagePerGroup) { $candidate.GroupIds } else { @($candidate.Root) }
	if ($inventoryError) { $reasons.Add("Cannot establish package coverage: $inventoryError") }
	else {
		try {
			$coverage = Get-PackageCoverageResult -GroupIds $resourceGroupIds -Inventory $inventory -AllowSharedGroups
			if ($coverage.Status -ne 'Ready') { $candidate.Status = $coverage.Status; $reasons.Add($coverage.Reason) }
		}
		catch { $reasons.Add("Cannot establish package activity: $($_.Exception.Message)") }
	}
	foreach ($id in $resourceGroupIds) {
		if (-not $eligibility.ContainsKey($id)) {
			try { $null = Get-PackageGroup -GroupId $id; $eligibility[$id] = '' }
			catch { $eligibility[$id] = $_.Exception.Message }
		}
		if ($eligibility[$id]) { $reasons.Add($eligibility[$id]) }
	}
	try {
		if ($OnePackagePerGroup) {
			$candidate.Description = Get-DiscoverySingleGroupDescription -GroupId $candidate.Root -Groups $groups
		}
		else {
			$candidate.Description = Get-DiscoveryPackageDescription -GroupIds $candidate.GroupIds `
				-Groups $groups -ApplicationCache $applicationCache -RootGroupId $candidate.Root
		}
	}
	catch {
		$reason = if ($OnePackagePerGroup) { 'Cannot establish group description' } else { 'Cannot establish group application assignments' }
		$reasons.Add("${reason}: $($_.Exception.Message)")
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
	$topologyColumns = @('Root group ID', 'Root group', 'Group ID', 'Group', 'Parent group IDs', 'Direct user count', 'Application', 'Service principal ID')
	for ($i = 0; $i -lt $topologyColumns.Count; $i++) { $topology.Cells[1, ($i + 1)].Value = $topologyColumns[$i] }
	$sheetRow = 2
	$blockedRow = 2
	$topologyRow = 2
	$privilegedRolePattern = '(?:(?<![\p{L}\p{N}])|(?<=[a-z])(?=[A-Z]))(?i:admin(?:s|istrators?)?|super[\s_-]*users?)(?:(?![\p{L}\p{N}])|(?=[A-Z][a-z]))'
	foreach ($candidate in $candidates) {
		$resourceGroupIds = if ($OnePackagePerGroup) { $candidate.GroupIds } else { @($candidate.Root) }
		$rootName = [string]$groups[$candidate.Root].DisplayName
		$appName = [string]$candidate.ServicePrincipalName
		$name = "$appName - $rootName".Trim(' ', '-')
		$description = $candidate.Description
		if ($candidate.Status -eq 'Ready') {
			$matches = @(
				foreach ($id in $candidate.GroupIds) {
					$groupName = [string]$groups[$id].DisplayName
					[regex]::Matches($groupName, $privilegedRolePattern) | ForEach-Object { $_.Value }
				}
			)
			$role = if ($matches.Count) { "Potential privileged access (group name matches: $(@($matches | Sort-Object -Unique) -join '; '))" } else { '' }
			$action = 'Complete approvers and verify business information.'
			if ($role) { $action = "Verify the suggested privileged role. $action" }
			if (-not $name -or $name -match '(?i)(?<![a-z0-9])(?:access[\s_-]*package|ap|sg)(?![a-z0-9])') {
				$action = "Replace the suggested package name; it violates naming rules. $action"
			}
			$values = @($role, '', '', '', ($resourceGroupIds -join ';'), $description, $name, '', '', '', $candidate.ServicePrincipalId, $candidate.Root, $action)
			for ($i = 0; $i -lt $values.Count; $i++) { $sheet.Cells[$sheetRow, ($i + 1)].Value = [string]$values[$i] }
			$sheetRow++
		}
		else {
			$values = @($candidate.Status, $candidate.Reason, $rootName, $candidate.Root,
				($candidate.GroupIds -join ';'), (($candidate.GroupIds | ForEach-Object { $groups[$_].DisplayName }) -join '; '))
			for ($i = 0; $i -lt $values.Count; $i++) { $blocked.Cells[$blockedRow, ($i + 1)].Value = [string]$values[$i] }
			$blockedRow++
		}
	}
	foreach ($candidate in $discovery.Candidates) {
		$rootName = [string]$groups[$candidate.Root].DisplayName
		$appName = [string]$candidate.ServicePrincipalName
		foreach ($id in $candidate.GroupIds) {
			$parents = @($candidate.GroupIds | Where-Object { $groups[$_].ChildIds -contains $id })
			$values = @($candidate.Root, $rootName, $id, $groups[$id].DisplayName, ($parents -join ';'),
				$groups[$id].DirectUserCount, $appName, $candidate.ServicePrincipalId)
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
