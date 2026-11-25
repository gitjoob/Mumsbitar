<#
.SYNOPSIS
List on-premises synced groups that have in-progress or completed access reviews,
with denied users that must be removed manually in Active Directory.
.DESCRIPTION
Read-only: uses Graph GET requests only and changes nothing in the tenant.
Reviews of synced groups are created without auto-apply because Entra cannot
change on-premises memberships. This script finds those reviews afterwards and
writes an Excel workbook with these sheets:
  Groups     - one row per review instance and synced group, with decision counts
  Removals   - Deny decisions (manual AD removal worklist), with current membership
  Decisions  - every decision for the synced groups (audit evidence)
  Issues     - anything that could not be read
Group IDs are taken from the instance scope (falling back to the definition
scope), so group membership reviews and application reviews scoped to groups
are both covered. Requires Microsoft.Graph.Authentication and ImportExcel.
Permissions: AccessReview.Read.All, GroupMember.Read.All and User.Read.All
(or Directory.Read.All).
.PARAMETER Prefixes
Optional literal, case-insensitive review display-name prefixes. Empty means
every access review definition.
.PARAMETER InstanceStatus
Instance statuses to include. Defaults to running and completed states.
.PARAMETER OutputPath
New .xlsx file. Existing files are never overwritten.
.EXAMPLE
.\Get-SyncedGroupReviewWorklist.ps1 -Prefixes @('AR ')
#>
[CmdletBinding()]
param(
	[string[]]$Prefixes = @(),
	[ValidateNotNullOrEmpty()]
	[string[]]$InstanceStatus = @('InProgress', 'Completed', 'AutoReviewing', 'AutoReviewed', 'Applying', 'Applied'),
	[string]$OutputPath = (Join-Path $PSScriptRoot "reports/SyncedGroupReviews_$([datetime]::UtcNow.ToString('yyyyMMddTHHmmss')).xlsx")
)

$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module ImportExcel -WarningAction SilentlyContinue -ErrorAction Stop

$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
if (Test-Path -LiteralPath $OutputPath) { throw "Output file already exists: $OutputPath" }
$context = Get-MgContext
if (-not $context) { throw 'Connect to Microsoft Graph first.' }
$environment = Get-MgEnvironment -Name $context.Environment
$base = "$($environment.GraphEndpoint.TrimEnd('/'))/v1.0"

$issues = [System.Collections.Generic.List[object]]::new()
function Add-WorklistIssue {
	param([string]$Review, [string]$Source, [string]$Message)
	$issues.Add([pscustomobject]@{ Review = $Review; Source = $Source; Message = $Message })
	Write-Warning "$Source`: $Message"
}

function Get-WorklistCollection {
	param([string]$Uri)
	$seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
	while ($Uri) {
		if (-not $Uri.StartsWith("$base/", [System.StringComparison]::OrdinalIgnoreCase) -or -not $seen.Add($Uri)) {
			throw 'Invalid or repeated Graph paging URL.'
		}
		$response = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
		foreach ($item in @($response.value)) { if ($null -ne $item) { $item } }
		$Uri = $response.'@odata.nextLink'
	}
}

function Get-WorklistGroupIds {
	param([object]$Scope)
	if ($null -eq $Scope) { return @() }
	$json = $Scope | ConvertTo-Json -Depth 15 -Compress
	@([regex]::Matches($json, '/groups/([0-9a-fA-F-]{36})') | ForEach-Object { $_.Groups[1].Value.ToLowerInvariant() } | Select-Object -Unique)
}

function Format-WorklistTime {
	param([object]$Value)
	if ($null -eq $Value -or '' -eq [string]$Value) { return '' }
	if ($Value -is [datetime]) {
		$utc = if ($Value.Kind -eq [DateTimeKind]::Local) { $Value.ToUniversalTime() } else { $Value }
		return $utc.ToString('yyyy-MM-dd HH:mm')
	}
	$parsed = [datetimeoffset]::MinValue
	if ([datetimeoffset]::TryParse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture,
		[System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) {
		return $parsed.UtcDateTime.ToString('yyyy-MM-dd HH:mm')
	}
	[string]$Value
}

$groupCache = @{}
function Get-WorklistGroup {
	param([string]$GroupId, [string]$Review)
	if ($groupCache.ContainsKey($GroupId)) { return $groupCache[$GroupId] }
	$group = $null
	try {
		$group = Invoke-MgGraphRequest -Method GET -OutputType PSObject -ErrorAction Stop `
			-Uri "$base/groups/$GroupId`?`$select=id,displayName,onPremisesSyncEnabled,onPremisesSamAccountName,onPremisesDomainName,onPremisesNetBiosName,onPremisesSecurityIdentifier"
	}
	catch { Add-WorklistIssue $Review "Group $GroupId" $_.Exception.Message }
	$groupCache[$GroupId] = $group
	$group
}

$memberCache = @{}
function Get-WorklistMemberSet {
	param([string]$GroupId, [string]$Review)
	if ($memberCache.ContainsKey($GroupId)) { return , $memberCache[$GroupId] }
	$set = $null
	try {
		$set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
		foreach ($member in Get-WorklistCollection "$base/groups/$GroupId/members?`$select=id&`$top=999") { $null = $set.Add([string]$member.id) }
	}
	catch {
		$set = $null
		Add-WorklistIssue $Review "Members of group $GroupId" $_.Exception.Message
	}
	$memberCache[$GroupId] = $set
	, $set
}

$userCache = @{}
function Get-WorklistUser {
	param([string]$UserId)
	if (-not $UserId) { return $null }
	if ($userCache.ContainsKey($UserId)) { return $userCache[$UserId] }
	$user = $null
	try {
		$user = Invoke-MgGraphRequest -Method GET -OutputType PSObject -ErrorAction Stop `
			-Uri "$base/users/$UserId`?`$select=id,displayName,userPrincipalName,mail,onPremisesSamAccountName,onPremisesSyncEnabled,accountEnabled"
	}
	catch { $user = $null }
	$userCache[$UserId] = $user
	$user
}

$groupRows = [System.Collections.Generic.List[object]]::new()
$removalRows = [System.Collections.Generic.List[object]]::new()
$decisionRows = [System.Collections.Generic.List[object]]::new()
$definitionsSeen = 0

foreach ($definition in Get-WorklistCollection "$base/identityGovernance/accessReviews/definitions") {
	$name = [string]$definition.displayName
	if ($Prefixes.Count -and -not @($Prefixes | Where-Object { $_ -and $name.StartsWith($_, [System.StringComparison]::OrdinalIgnoreCase) }).Count) { continue }
	$definitionsSeen++
	$definitionUri = "$base/identityGovernance/accessReviews/definitions/$([uri]::EscapeDataString($definition.id))"
	try { $instances = @(Get-WorklistCollection "$definitionUri/instances") }
	catch { Add-WorklistIssue $name 'Instance listing' $_.Exception.Message; continue }
	foreach ($instance in $instances) {
		if ($InstanceStatus -notcontains [string]$instance.status) { continue }
		$groupIds = @(Get-WorklistGroupIds $instance.scope)
		if (-not $groupIds.Count) { $groupIds = @(Get-WorklistGroupIds $definition.scope) }
		$syncedGroups = @($groupIds | ForEach-Object { Get-WorklistGroup $_ $name } | Where-Object { $_ -and $_.onPremisesSyncEnabled -eq $true })
		if (-not $syncedGroups.Count) { continue }
		Write-Host "Synced group review: $name ($($instance.status))"
		$instanceUri = "$definitionUri/instances/$([uri]::EscapeDataString($instance.id))"
		try {
			$decisions = @(Get-WorklistCollection "$instanceUri/decisions?`$select=id,decision,justification,reviewedBy,reviewedDateTime,applyResult,appliedDateTime,principal,resource,resourceLink")
		}
		catch {
			Add-WorklistIssue $name 'Decisions' $_.Exception.Message
			$decisions = $null
		}
		foreach ($group in $syncedGroups) {
			$groupDecisions = @($decisions | Where-Object {
				$_ -and ($syncedGroups.Count -eq 1 -or [string]$_.resourceLink -match [regex]::Escape($group.id) -or [string]$_.resource.id -eq $group.id)
			})
			$count = @{}
			foreach ($d in $groupDecisions) { $count[[string]$d.decision] = 1 + [int]$count[[string]$d.decision] }
			$members = $null
			if (@($groupDecisions | Where-Object decision -EQ 'Deny').Count) { $members = Get-WorklistMemberSet $group.id $name }
			$common = [ordered]@{
				Review = $name; InstanceStatus = $instance.status
				InstanceStart = Format-WorklistTime $instance.startDateTime; InstanceEnd = Format-WorklistTime $instance.endDateTime
				Group = $group.displayName; GroupSamAccountName = $group.onPremisesSamAccountName
				Domain = $group.onPremisesDomainName; GroupId = $group.id; InstanceId = $instance.id
			}
			foreach ($d in $groupDecisions) {
				$principalId = [string]$d.principal.id
				$user = Get-WorklistUser $principalId
				$row = [ordered]@{}
				foreach ($key in $common.Keys) { $row[$key] = $common[$key] }
				$row.Decision = $d.decision
				$row.UserDisplayName = if ($user) { $user.displayName } else { $d.principal.displayName }
				$row.UserPrincipalName = if ($user) { $user.userPrincipalName } else { $d.principal.userPrincipalName }
				$row.UserSamAccountName = if ($user) { $user.onPremisesSamAccountName } else { '' }
				$row.UserId = $principalId
				$row.ReviewedBy = $d.reviewedBy.displayName
				$row.ReviewedByUpn = $d.reviewedBy.userPrincipalName
				$row.ReviewedDate = Format-WorklistTime $d.reviewedDateTime
				$row.Justification = $d.justification
				$row.ApplyResult = $d.applyResult
				$decisionRows.Add([pscustomobject]$row)
				if ($d.decision -eq 'Deny') {
					$row.StillMemberInEntra = if ($null -eq $members) { 'Unknown' } elseif ($members.Contains($principalId)) { 'Yes' } else { 'No (already removed)' }
					$removalRows.Add([pscustomobject]$row)
				}
			}
			$row = [ordered]@{}
			foreach ($key in $common.Keys) { $row[$key] = $common[$key] }
			$row.DecisionsReadable = if ($null -eq $decisions) { 'No (see Issues)' } else { 'Yes' }
			$row.Total = $groupDecisions.Count
			foreach ($key in @('Approve', 'Deny', 'DontKnow', 'NotReviewed')) { $row[$key] = [int]$count[$key] }
			$row.DenyStillMember = @($removalRows | Where-Object { $_.GroupId -eq $group.id -and $_.InstanceId -eq $instance.id -and $_.StillMemberInEntra -eq 'Yes' }).Count
			$row.ManualActionNeeded = if ($row.DenyStillMember -gt 0) { 'Yes - remove denied users in AD' }
				elseif ($instance.status -eq 'InProgress') { 'Review still running' } else { 'No' }
			$row.SharedReviewWithGroups = $syncedGroups.Count
			$groupRows.Add([pscustomobject]$row)
		}
	}
}

$directory = Split-Path -Parent $OutputPath
if (-not (Test-Path -LiteralPath $directory)) { $null = New-Item -ItemType Directory -Path $directory }
$excel = @{ Path = $OutputPath; AutoSize = $true; FreezeTopRow = $true; BoldTopRow = $true; AutoFilter = $true }
$empty = { param($Text) [pscustomobject]@{ Info = $Text } }
if ($groupRows.Count) { $groupRows | Export-Excel @excel -WorksheetName 'Groups' -TableStyle Medium2 }
else { & $empty 'No synced groups with matching review instances found.' | Export-Excel @excel -WorksheetName 'Groups' }
if ($removalRows.Count) { $removalRows | Export-Excel @excel -WorksheetName 'Removals' }
else { & $empty 'No Deny decisions on synced groups.' | Export-Excel @excel -WorksheetName 'Removals' }
if ($decisionRows.Count) { $decisionRows | Export-Excel @excel -WorksheetName 'Decisions' }
else { & $empty 'No decisions returned.' | Export-Excel @excel -WorksheetName 'Decisions' }
if ($issues.Count) { $issues | Export-Excel @excel -WorksheetName 'Issues' }
else { & $empty 'No collection issues.' | Export-Excel @excel -WorksheetName 'Issues' }

Write-Host "Definitions checked: $definitionsSeen | Synced group rows: $($groupRows.Count) | Deny decisions: $($removalRows.Count) | Still members: $(@($removalRows | Where-Object StillMemberInEntra -EQ 'Yes').Count) | Issues: $($issues.Count)"
Write-Host "Workbook: $OutputPath"
