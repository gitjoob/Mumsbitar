<#
.SYNOPSIS
Report possible conflicts between access review decisions and package assignments.
.DESCRIPTION
This standalone script is read-only. It does not remove assignments or submit
requests. The report contains sensitive identities and must be kept private.
All time windows are inclusive of the start and exclusive of the end.
.EXAMPLE
.\Get-ReviewAssignmentReconciliation.ps1 -CatalogId '<catalog ID>' `
    -PackageCreatedFrom '2026-01-01T00:00:00Z' -PackageCreatedBefore '2026-01-03T00:00:00Z' `
    -ReviewPrefix 'AR ' -ReviewStartedFrom '2025-12-01T00:00:00Z' `
    -ReviewStartedBefore '2026-01-04T00:00:00Z' `
    -RequestCreatedFrom '2026-01-01T00:00:00Z' -RequestCreatedBefore '2026-01-03T00:00:00Z' `
    -ReportPath '.\private\reconciliation.json'
#>
[CmdletBinding()]
param(
	[Parameter(Mandatory)][guid]$CatalogId,
	[Parameter(Mandatory)][datetimeoffset]$PackageCreatedFrom,
	[Parameter(Mandatory)][datetimeoffset]$PackageCreatedBefore,
	[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ReviewPrefix,
	[Parameter(Mandatory)][datetimeoffset]$ReviewStartedFrom,
	[Parameter(Mandatory)][datetimeoffset]$ReviewStartedBefore,
	[Parameter(Mandatory)][datetimeoffset]$RequestCreatedFrom,
	[Parameter(Mandatory)][datetimeoffset]$RequestCreatedBefore,
	[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ReportPath
)

$ErrorActionPreference = 'Stop'
if ($CatalogId -eq [guid]::Empty -or $PackageCreatedFrom -ge $PackageCreatedBefore -or
	$ReviewStartedFrom -ge $ReviewStartedBefore -or $RequestCreatedFrom -ge $RequestCreatedBefore) {
	throw 'Supply a non-empty catalog ID and valid start/end windows.'
}
if (Test-Path -LiteralPath $ReportPath) { throw 'The report path already exists; choose a new file.' }

function Get-ReconciliationCollection {
	param([string]$Uri)
	$visited = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
	while ($Uri) {
		if (-not $Uri.StartsWith("$script:graphBase/", [StringComparison]::OrdinalIgnoreCase) -or -not $visited.Add($Uri)) {
			throw 'Graph returned an invalid or repeated paging URL.'
		}
		$page = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Hashtable -ErrorAction Stop
		if ($page -isnot [System.Collections.IDictionary] -or -not $page.Contains('value') -or $null -eq $page.value) {
			throw 'Graph returned an incomplete collection; no report was produced.'
		}
		foreach ($item in $page.value) {
			if ($null -eq $item) { throw 'Graph returned an empty collection item; no report was produced.' }
			$item
		}
		$Uri = [string]$page.'@odata.nextLink'
	}
}

function Get-ReconciliationTime {
	param($Value, [string]$Description)
	$result = [datetimeoffset]::MinValue
	if (-not [datetimeoffset]::TryParse([string]$Value, [ref]$result)) {
		throw "Graph returned an invalid $Description; no report was produced."
	}
	$result
}

function Get-ReconciliationLinkId {
	param([string]$Link, [string]$Kind)
	if ($Link -notmatch "/v1\.0/$Kind/([0-9a-fA-F-]{36})(?:/?(?:\?.*)?)$") { return $null }
	$id = [guid]::Empty
	if (-not [guid]::TryParse($Matches[1], [ref]$id) -or $id -eq [guid]::Empty) { return $null }
	$id.ToString()
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$context = Get-MgContext
if (-not $context -or -not $context.TenantId) { throw 'Connect to Microsoft Graph first.' }
$environment = Get-MgEnvironment -Name $context.Environment
if (-not $environment.GraphEndpoint) { throw 'Unable to determine the connected Microsoft Graph endpoint.' }
$script:graphBase = "$($environment.GraphEndpoint.TrimEnd('/'))/v1.0"
$em = "$script:graphBase/identityGovernance/entitlementManagement"
$reviews = "$script:graphBase/identityGovernance/accessReviews/definitions"
$started = [datetimeoffset]::UtcNow
$issues = [System.Collections.Generic.List[object]]::new()
$packageIssues = [System.Collections.Generic.List[object]]::new()
$decisionsByGroup = @{}
$matchedGroupIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$definitions = @(Get-ReconciliationCollection $reviews)
$matchingDefinitions = 0
foreach ($definition in $definitions) {
	if (-not $definition.id -or -not $definition.displayName) { throw 'A review definition has no ID or name; no report was produced.' }
	if (-not $definition.displayName.StartsWith($ReviewPrefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
	$matchingDefinitions++
	$definitionUri = "$reviews/$([uri]::EscapeDataString([string]$definition.id))"
	foreach ($instance in @(Get-ReconciliationCollection "$definitionUri/instances")) {
		if (-not $instance.id) { throw 'A review instance has no ID; no report was produced.' }
		$instanceStart = Get-ReconciliationTime $instance.startDateTime 'review instance start'
		if ($instanceStart -lt $ReviewStartedFrom -or $instanceStart -ge $ReviewStartedBefore) { continue }
		$instanceUri = "$definitionUri/instances/$([uri]::EscapeDataString([string]$instance.id))"
		foreach ($decision in @(Get-ReconciliationCollection "$instanceUri/decisions")) {
			if (-not $decision.id -or -not $decision.decision) { throw 'A review decision has no ID or decision; no report was produced.' }
			if ($decision.decision -notin @('Deny', 'NotReviewed', 'DontKnow')) { continue }
			$groupId = Get-ReconciliationLinkId ([string]$decision.resourceLink) 'groups'
			$userId = Get-ReconciliationLinkId ([string]$decision.principalLink) 'users'
			if ($userId -and $decision.principal.id -and [string]$decision.principal.id -ne $userId) { $userId = $null }
			if (-not $groupId -or -not $userId) {
				$issues.Add([pscustomobject]@{
					ReviewId = $definition.id; InstanceId = $instance.id; DecisionId = $decision.id
					Reason = 'Cannot identify both a group resource and a user from the decision links; inspect manually.'
					Decision = $decision.decision; ResourceLink = $decision.resourceLink; PrincipalLink = $decision.principalLink
				})
				continue
			}
			if (-not $decisionsByGroup.ContainsKey($groupId)) { $decisionsByGroup[$groupId] = [System.Collections.Generic.List[object]]::new() }
			$decisionsByGroup[$groupId].Add([pscustomobject]@{
				ReviewId = [string]$definition.id; ReviewName = [string]$definition.displayName
				InstanceId = [string]$instance.id; InstanceStatus = [string]$instance.status
				InstanceStarted = $instanceStart.ToString('o'); InstanceScope = $instance.scope
				DecisionId = [string]$decision.id; Decision = [string]$decision.decision
				ApplyResult = [string]$decision.applyResult; ReviewedAt = $decision.reviewedDateTime
				AppliedAt = $decision.appliedDateTime; UserId = $userId
				UserName = [string]$decision.principal.displayName; GroupId = $groupId
			})
		}
	}
}

$rows = [System.Collections.Generic.List[object]]::new()
$packages = @(Get-ReconciliationCollection "$em/accessPackages?`$expand=catalog")
$packagesInWindow = 0
foreach ($package in $packages) {
	$packageCatalogId = if ($package.catalog.id) { [string]$package.catalog.id } else { [string]$package.catalogId }
	if (-not $packageCatalogId) { throw 'A package has no catalog ID; package scope cannot be verified.' }
	if ($packageCatalogId -ne $CatalogId.ToString()) { continue }
	if (-not $package.id) { throw 'A package has no ID; no report was produced.' }
	$created = Get-ReconciliationTime $package.createdDateTime 'package creation time'
	if ($created -lt $PackageCreatedFrom -or $created -ge $PackageCreatedBefore) { continue }
	$packagesInWindow++
	$packageId = [string]$package.id
	$detail = Invoke-MgGraphRequest -Method GET -Uri "$em/accessPackages/$([uri]::EscapeDataString($packageId))?`$expand=resourceRoleScopes(`$expand=role,scope)" -OutputType Hashtable -ErrorAction Stop
	if ($detail -isnot [System.Collections.IDictionary] -or -not $detail.Contains('resourceRoleScopes') -or $null -eq $detail.resourceRoleScopes) {
		throw "Package '$packageId' has no resource role data; no report was produced."
	}
	$roles = @($detail.resourceRoleScopes)
	if ($detail.'resourceRoleScopes@odata.nextLink') { $roles += @(Get-ReconciliationCollection $detail.'resourceRoleScopes@odata.nextLink') }
	if (-not $roles.Count) {
		$packageIssues.Add([pscustomobject]@{ PackageId = $packageId; Reason = 'No resource roles were returned; group coverage is unknown.' })
	}
	$groupIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
	foreach ($role in $roles) {
		if (-not $role.role.originSystem -or -not $role.scope.originSystem -or -not $role.scope.originId) {
			$packageIssues.Add([pscustomobject]@{ PackageId = $packageId; Reason = 'A resource role or scope is incomplete; group coverage is unknown.' })
			continue
		}
		$group = [guid]::Empty
		if ($role.scope.originSystem -eq 'AadGroup') {
			if ([guid]::TryParse([string]$role.scope.originId, [ref]$group) -and $group -ne [guid]::Empty) {
				$null = $groupIds.Add($group.ToString())
			}
			else {
				$packageIssues.Add([pscustomobject]@{ PackageId = $packageId; Reason = 'A group scope has an invalid group ID; group coverage is unknown.' })
			}
		}
	}
	$matched = @(foreach ($groupId in $groupIds) {
		if ($decisionsByGroup.ContainsKey($groupId)) {
			$null = $matchedGroupIds.Add($groupId)
			$decisionsByGroup[$groupId]
		}
	})
	if (-not $matched.Count) { continue }
	$filter = [uri]::EscapeDataString("accessPackage/id eq '$packageId'")
	$assignments = @(Get-ReconciliationCollection "$em/assignments?`$filter=$filter&`$expand=target")
	$requests = @(Get-ReconciliationCollection "$em/assignmentRequests?`$filter=$filter&`$expand=assignment")
	if (@($assignments | Where-Object { -not $_.target.objectId }).Count) {
		$packageIssues.Add([pscustomobject]@{ PackageId = $packageId; Reason = 'An assignment target object ID was not returned; assignment matching is incomplete.' })
	}
	foreach ($decision in $matched) {
		$matchingAssignments = @($assignments | Where-Object { $_.target.objectId -eq $decision.UserId -and $_.state -ne 'expired' })
		if (-not $matchingAssignments.Count) {
			$rows.Add([pscustomobject]@{ PackageId = $packageId; PackageName = $package.displayName; PackageCreatedAt = $created.ToString('o')
				ResourceRoles = $roles; Review = $decision; Assignment = $null; IncidentAdminAddRequests = @()
				Reason = 'No matching non-expired assignment was returned. Verify visibility and pending requests manually.' })
			continue
		}
		foreach ($assignment in $matchingAssignments) {
			if (-not $assignment.id) { throw "An assignment in package '$packageId' has no ID; no report was produced." }
			$matchingRequests = @(
				foreach ($request in $requests) {
					if ($request.requestType -ne 'adminAdd' -or $request.assignment.id -ne $assignment.id) { continue }
					$requestTime = Get-ReconciliationTime $request.createdDateTime 'request creation time'
					if ($requestTime -ge $RequestCreatedFrom -and $requestTime -lt $RequestCreatedBefore) {
						[pscustomobject]@{ Id = $request.id; State = $request.state; CreatedAt = $requestTime.ToString('o') }
					}
				}
			)
			$rows.Add([pscustomobject]@{ PackageId = $packageId; PackageName = $package.displayName; PackageCreatedAt = $created.ToString('o')
				ResourceRoles = $roles; Review = $decision
				Assignment = [pscustomobject]@{ Id = $assignment.id; State = $assignment.state; Target = $assignment.target; Schedule = $assignment.schedule }
				IncidentAdminAddRequests = $matchingRequests
				Reason = 'Manual review required: this report does not verify later approvals, pending requests, or current resource changes.' })
		}
	}
}

$noPackageMatches = @(foreach ($groupId in $decisionsByGroup.Keys) {
	if (-not $matchedGroupIds.Contains($groupId)) { $decisionsByGroup[$groupId] }
})
$report = [ordered]@{
	SchemaVersion = 1; TenantId = [string]$context.TenantId; StartedAt = $started.ToString('o'); FinishedAt = [datetimeoffset]::UtcNow.ToString('o')
	CatalogId = $CatalogId.ToString()
	PackageCreatedFrom = $PackageCreatedFrom.ToString('o'); PackageCreatedBefore = $PackageCreatedBefore.ToString('o')
	ReviewPrefix = $ReviewPrefix; ReviewStartedFrom = $ReviewStartedFrom.ToString('o'); ReviewStartedBefore = $ReviewStartedBefore.ToString('o')
	RequestCreatedFrom = $RequestCreatedFrom.ToString('o'); RequestCreatedBefore = $RequestCreatedBefore.ToString('o')
	MatchingDefinitions = $matchingDefinitions; PackagesInWindow = $packagesInWindow
	Limitations = @('Read-only inventory; no row is approved for removal.', 'A matching adminAdd request does not prove it came from the backfill script.', 'Other reviews and subsequent approvals are not checked.', 'Results depend on Graph permissions and visibility.')
	UnmatchedDecisions = @($issues.ToArray()); PackageIssues = @($packageIssues.ToArray())
	DecisionsWithoutMatchingPackage = $noPackageMatches; Results = @($rows.ToArray())
}
$json = $report | ConvertTo-Json -Depth 40
$fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
$stream = [System.IO.File]::Open($fullPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
try {
	$writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
	try { $writer.Write($json) } finally { $writer.Dispose() }
} finally { $stream.Dispose() }
Write-Host "Read-only report saved to $fullPath. $($rows.Count) result(s); $($issues.Count) unidentified decision(s); $($packageIssues.Count) package issue(s)."
if (-not $matchingDefinitions -or -not $packagesInWindow) {
	Write-Warning "Scope returned $matchingDefinitions matching review definition(s) and $packagesInWindow package(s). Check the scope and Graph visibility; an empty report does not prove there are no affected assignments."
}
