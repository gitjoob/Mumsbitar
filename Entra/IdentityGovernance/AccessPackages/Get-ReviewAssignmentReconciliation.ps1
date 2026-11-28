<#
.SYNOPSIS
Investigate and optionally remediate access review and package assignment conflicts.
.DESCRIPTION
Preview finds successfully applied Deny decisions for users with assignments
to packages granting the reviewed group. Every package in the specified
catalog is checked, regardless of creation date. By default the script only
reports matches. Add -Remediate to find and submit eligible removals in the
same run, or -Remediate -WhatIf to simulate the entire operation. The report
and candidate CSV are outputs; no approval file is needed. Packages with
multiple resources and uncertain or conflicting cases are reported for manual
review and not removed automatically.

Pause backfill, reviews and other changes until remediation finishes, and run
only one remediation at a time. Graph cannot lock these objects against
concurrent changes. Review and package data is reused during a run to avoid
repeated scans; each decision and assignment is re-read before removal. The
script cannot detect every out-of-band change. A submitted request is not
confirmation of removal; verify requests and group membership afterwards.
Never retry an uncertain request without checking Entra and the audit first.

Reports, selection files and audit logs contain sensitive identities; store
them privately. -LookbackDays defaults to 30 and scopes review instances whose
scheduled periods overlap that window. Reviews applied late may be missed.
Explicit time windows include the start and exclude the end. For delegated
processing, connect with AccessReview.Read.All,
EntitlementManagement.ReadWrite.All and RoleManagement.Read.Directory and
activate a directly assigned tenant-wide Identity Governance Administrator
or Global Administrator role. Application authentication requires tenant-wide
AccessReview.Read.All and EntitlementManagement.ReadWrite.All application
permissions for remediation. Preview accepts EntitlementManagement.Read.All.
.EXAMPLE
Connect-MgGraph -Scopes 'AccessReview.Read.All','EntitlementManagement.ReadWrite.All','RoleManagement.Read.Directory'
.\Get-ReviewAssignmentReconciliation.ps1 -CatalogId '<catalog ID>' -Remediate
# Add -WhatIf to preview without submitting removals.
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
	[Parameter(Mandatory)][guid]$CatalogId,
	[ValidateRange(1, 36500)][int]$LookbackDays = 30,
	[string]$ReviewPrefix = '',
	[datetimeoffset]$ReviewStartedFrom,
	[datetimeoffset]$ReviewStartedBefore,
	[ValidateNotNullOrEmpty()][string]$ReportPath = (Join-Path (Get-Location).Path "reconciliation-$([datetime]::UtcNow.ToString('yyyyMMddTHHmmssfff'))-$([guid]::NewGuid().ToString('N')).json"),
	[string]$AuditPath,
	[switch]$Remediate
)

$ErrorActionPreference = 'Stop'
$windowEnd = [datetimeoffset]::UtcNow
$windowStart = $windowEnd.AddDays(-$LookbackDays)
$explicitReviewWindow = $PSBoundParameters.ContainsKey('ReviewStartedFrom') -or $PSBoundParameters.ContainsKey('ReviewStartedBefore')
if ($explicitReviewWindow -and
	(-not $PSBoundParameters.ContainsKey('ReviewStartedFrom') -or -not $PSBoundParameters.ContainsKey('ReviewStartedBefore'))) {
	throw 'Supply both ReviewStartedFrom and ReviewStartedBefore when overriding the review window.'
}
if (-not $explicitReviewWindow) { $ReviewStartedFrom = $windowStart; $ReviewStartedBefore = $windowEnd }
if ($CatalogId -eq [guid]::Empty -or $ReviewStartedFrom -ge $ReviewStartedBefore) {
	throw 'Supply a non-empty catalog ID and valid start/end windows.'
}
if (Test-Path -LiteralPath $ReportPath) { throw 'The report path already exists; choose a new file.' }
if ($Remediate -and -not $AuditPath) {
	$AuditPath = Join-Path (Get-Location).Path "reconciliation-audit-$([datetime]::UtcNow.ToString('yyyyMMddTHHmmssfff'))-$([guid]::NewGuid().ToString('N')).jsonl"
}
if ($AuditPath -and (Test-Path -LiteralPath $AuditPath)) { throw 'The audit path already exists; choose a new file.' }
foreach ($path in @($ReportPath, $AuditPath)) {
	if (-not $path) { continue }
	$resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path)
	if (-not [System.IO.Directory]::Exists([System.IO.Path]::GetDirectoryName($resolved))) {
		throw "The output directory does not exist for '$path'."
	}
}
if ($AuditPath -and
	$ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath) -eq
	$ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($AuditPath)) {
	throw 'The report and audit must use different output paths.'
}
$progress = @{ LastMessage = [datetime]::MinValue; Phase = 'Starting'; Current = 0; Total = 0 }
function Show-ReconciliationProgress {
	param([string]$Phase, [int]$Current = 0, [int]$Total = 0, [switch]$Force)
	$progress.Phase = $Phase
	$progress.Current = $Current
	$progress.Total = $Total
	$status = if ($Total -gt 0) { "$Phase ($Current of $Total). Last update: $(Get-Date -Format 'HH:mm:ss')" }
		else { "$Phase. Last update: $(Get-Date -Format 'HH:mm:ss')" }
	$parameters = @{ Activity = 'Reconciling access reviews and assignments'; Status = $status }
	if ($Total -gt 0) { $parameters.PercentComplete = [int][math]::Min(100, [math]::Floor(100.0 * $Current / $Total)) }
	Write-Progress @parameters
	if ($Force -or ([datetime]::UtcNow - $progress.LastMessage).TotalSeconds -ge 15) {
		Write-Host $status
		$progress.LastMessage = [datetime]::UtcNow
	}
}

function Get-ReconciliationCollection {
	param([string]$Uri)
	$visited = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
	while ($Uri) {
		if (-not $Uri.StartsWith("$script:graphBase/", [StringComparison]::OrdinalIgnoreCase) -or -not $visited.Add($Uri)) {
			throw 'Graph returned an invalid or repeated paging URL.'
		}
		Show-ReconciliationProgress $progress.Phase $progress.Current $progress.Total
		$page = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Hashtable -ErrorAction Stop
		if ($page -isnot [System.Collections.IDictionary] -or -not $page.Contains('value') -or $null -eq $page.value) {
			throw 'Graph returned an incomplete collection; the operation stopped.'
		}
		foreach ($item in $page.value) {
			if ($null -eq $item) { throw 'Graph returned an empty collection item; the operation stopped.' }
			$item
		}
		$Uri = [string]$page.'@odata.nextLink'
	}
}

$reviewSnapshot = @{}
$verificationDetails = @{}
$verificationCollections = @{}
function Get-ReconciliationReviewSnapshot {
	param([string]$Uri)
	if (-not $reviewSnapshot.ContainsKey($Uri)) { $reviewSnapshot[$Uri] = @(Get-ReconciliationCollection $Uri) }
	$reviewSnapshot[$Uri]
}

function Get-ReconciliationVerificationDetail {
	param([string]$Uri)
	if (-not $verificationDetails.ContainsKey($Uri)) {
		$verificationDetails[$Uri] = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Hashtable -ErrorAction Stop
	}
	$verificationDetails[$Uri]
}

function Get-ReconciliationVerificationCollection {
	param([string]$Uri)
	if (-not $verificationCollections.ContainsKey($Uri)) { $verificationCollections[$Uri] = @(Get-ReconciliationCollection $Uri) }
	$verificationCollections[$Uri]
}

function Get-ReconciliationTime {
	param($Value, [string]$Description)
	$result = [datetimeoffset]::MinValue
	if (-not [datetimeoffset]::TryParse([string]$Value, [ref]$result)) {
		throw "Graph returned an invalid $Description; the operation stopped."
	}
	$result
}

function Get-ReconciliationLinkId {
	param([string]$Link, [string]$Kind)
	if ($Link -notmatch "(?:^|/v1\.0)/$Kind/([0-9a-fA-F-]{36})(?:/?(?:\?.*)?)$") { return $null }
	$id = [guid]::Empty
	if (-not [guid]::TryParse($Matches[1], [ref]$id) -or $id -eq [guid]::Empty) { return $null }
	$id.ToString()
}

function Assert-ReconciliationVisibility {
	param($Context)
	$scopes = @($Context.Scopes)
	if (-not @($scopes | Where-Object { $_ -in @('AccessReview.Read.All', 'AccessReview.ReadWrite.All') }).Count) {
		throw 'Approval processing requires tenant-wide AccessReview.Read.All or AccessReview.ReadWrite.All in the Graph connection.'
	}
	$entitlementScopes = if ($Remediate) { @('EntitlementManagement.ReadWrite.All') }
		else { @('EntitlementManagement.Read.All', 'EntitlementManagement.ReadWrite.All') }
	if (-not @($scopes | Where-Object { $_ -in $entitlementScopes }).Count) {
		throw "Approval processing requires one of these Graph permissions: $($entitlementScopes -join ', ')."
	}
	if ($Context.AuthType -eq 'AppOnly') { return }
	if ($Context.AuthType -ne 'Delegated' -or
		-not @($scopes | Where-Object { $_ -in @('RoleManagement.Read.Directory', 'RoleManagement.ReadWrite.Directory') }).Count) {
		throw 'Delegated approval processing requires RoleManagement.Read.Directory and a directly assigned, active tenant-wide Identity Governance Administrator or Global Administrator role.'
	}
	$me = Invoke-MgGraphRequest -Method GET -Uri "$script:graphBase/me?`$select=id" -OutputType Hashtable -ErrorAction Stop
	$id = [guid]::Empty
	if (-not [guid]::TryParse([string]$me.id, [ref]$id) -or $id -eq [guid]::Empty) { throw 'Cannot verify the signed-in administrator identity.' }
	$filter = [uri]::EscapeDataString("principalId eq '$id'")
	$roles = @(Get-ReconciliationCollection "$script:graphBase/roleManagement/directory/roleAssignments?`$filter=$filter")
	$allowed = @('62e90394-69f5-4237-9190-012177145e10', '45d8d3c5-c802-45c6-b32a-1d70b5e1e86e')
	if (-not @($roles | Where-Object { $_.directoryScopeId -eq '/' -and $_.roleDefinitionId -in $allowed }).Count) {
		throw 'Approval processing requires a directly assigned, active tenant-wide Identity Governance Administrator or Global Administrator role.'
	}
}

function Test-ReconciliationDirectScope {
	param($Scope, [string]$GroupId)
	$query = [string]$Scope.query
	$scopeType = [string]$Scope.queryType
	$scopeType -eq 'MicrosoftGraph' -and $query -match "^/(?:v1\.0/)?groups/$([regex]::Escape($GroupId))/members(?:/microsoft\.graph\.user)?/?$"
}

function Get-ReconciliationDecisionIds {
	param($Decision, $Scope)
	$groupId = Get-ReconciliationLinkId ([string]$Decision.resourceLink) 'groups'
	if (-not $Decision.resourceLink -and [string]$Scope.query -match '^/(?:v1\.0/)?groups/([0-9a-fA-F-]{36})/') {
		$scopeId = [guid]::Empty
		if ([guid]::TryParse($Matches[1], [ref]$scopeId) -and $scopeId -ne [guid]::Empty -and
			(Test-ReconciliationDirectScope $Scope $scopeId.ToString())) { $groupId = $scopeId.ToString() }
	}
	$userId = Get-ReconciliationLinkId ([string]$Decision.principalLink) 'users'
	if (-not $Decision.principalLink -and $Decision.principal.'@odata.type' -match '^#?microsoft\.graph\.userIdentity$') {
		$principalId = [guid]::Empty
		if ([guid]::TryParse([string]$Decision.principal.id, [ref]$principalId) -and $principalId -ne [guid]::Empty) {
			$userId = $principalId.ToString()
		}
	}
	if ($userId -and $Decision.principal.id -and [string]$Decision.principal.id -ne $userId) { $userId = $null }
	if ($userId -and $Decision.principal.'@odata.type' -and
		$Decision.principal.'@odata.type' -notmatch '^#?microsoft\.graph\.userIdentity$') { $userId = $null }
	[pscustomobject]@{ GroupId = $groupId; UserId = $userId }
}

function Get-ReconciliationReviewConflicts {
	param([object[]]$Candidates)
	$groups = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
	foreach ($candidate in $Candidates) { $null = $groups.Add([string]$candidate.Review.GroupId) }
	$records = @{}
	$uncertainGroups = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
	Show-ReconciliationProgress 'Scanning all reviews for conflicting decisions' -Force
	$allDefinitions = @(Get-ReconciliationReviewSnapshot $reviews)
	$index = 0
	foreach ($definition in $allDefinitions) {
		$index++
		Show-ReconciliationProgress 'Scanning all reviews for conflicting decisions' $index $allDefinitions.Count
		if (-not $definition.id) { throw 'A review definition has no ID; conflict checking is incomplete.' }
		$definitionUri = "$reviews/$([uri]::EscapeDataString([string]$definition.id))"
		foreach ($instance in @(Get-ReconciliationReviewSnapshot "$definitionUri/instances")) {
			if (-not $instance.id -or -not $instance.status) { throw 'A review instance has no ID or status; conflict checking is incomplete.' }
			$scopeQuery = [string]$instance.scope.query
			$scopeGroup = $null
			if ($scopeQuery -match '^/(?:v1\.0/)?groups/([0-9a-fA-F-]{36})(?:/|$|\?)') { $scopeGroup = $Matches[1] }
			if ($scopeGroup -and (Test-ReconciliationDirectScope $instance.scope $scopeGroup)) {
				if (-not $groups.Contains($scopeGroup)) { continue }
				if ($instance.status -ne 'Completed') {
					$start = [datetimeoffset]::MinValue
					if ($instance.status -ne 'NotStarted' -or
						-not [datetimeoffset]::TryParse([string]$instance.startDateTime, [ref]$start) -or $start -le [datetimeoffset]::UtcNow) {
						$null = $uncertainGroups.Add($scopeGroup)
					}
				}
			}
			$instanceUri = "$definitionUri/instances/$([uri]::EscapeDataString([string]$instance.id))"
			foreach ($decision in @(Get-ReconciliationReviewSnapshot "$instanceUri/decisions")) {
				if (-not $decision.id) { throw 'A review decision has no ID; conflict checking is incomplete.' }
				$ids = Get-ReconciliationDecisionIds $decision $instance.scope
				$groupId = $ids.GroupId
				if ($groupId -and $scopeGroup -and $groupId -ne $scopeGroup -and $groups.Contains($scopeGroup)) {
					$null = $uncertainGroups.Add($scopeGroup)
				}
				if (-not $groupId) {
					if ($scopeGroup) {
						if ($groups.Contains($scopeGroup)) { $null = $uncertainGroups.Add($scopeGroup) }
						continue
					}
					if ($decision.resourceLink -match '/v1\.0/servicePrincipals/[0-9a-fA-F-]{36}(?:/|$|\?)' -or
						$scopeQuery -match '^/(?:v1\.0/)?servicePrincipals/[0-9a-fA-F-]{36}(?:/|$|\?)') { continue }
					throw 'A review decision has an unidentifiable resource; tenant-wide conflict checking is incomplete.'
				}
				if (-not $groups.Contains($groupId)) { continue }
				$userId = $ids.UserId
				if (-not $userId) {
					if ($decision.principal.'@odata.type' -match 'servicePrincipalIdentity') { continue }
					$null = $uncertainGroups.Add($groupId)
					continue
				}
				$key = "$groupId/$userId"
				if (-not $records.ContainsKey($key)) { $records[$key] = [System.Collections.Generic.List[object]]::new() }
				$records[$key].Add([pscustomobject]@{
					GroupId = $groupId; UserId = $userId; ReviewId = [string]$definition.id
					InstanceId = [string]$instance.id; DecisionId = [string]$decision.id
					InstanceStatus = [string]$instance.status; Decision = [string]$decision.decision
					ApplyResult = [string]$decision.applyResult; ReviewedAt = [string]$decision.reviewedDateTime
					AppliedAt = [string]$decision.appliedDateTime
				})
			}
		}
	}
	[pscustomobject]@{ Records = $records; UncertainGroups = $uncertainGroups }
}

function Get-ReconciliationEvidenceIssue {
	param($Row)
	$review = $Row.Review
	if ($review.Decision -ne 'Deny' -or $review.ApplyResult -ne 'AppliedSuccessfully' -or $review.InstanceStatus -ne 'Completed') {
		return 'The instance is not completed with an applied Deny.'
	}
	if (-not (Test-ReconciliationDirectScope $review.InstanceScope $review.GroupId)) {
		return 'The review is not an identifiable direct group membership review.'
	}
	$roles = @($Row.ResourceRoles)
	if ($roles.Count -ne 1 -or $roles[0].role.originSystem -ne 'AadGroup' -or
		$roles[0].scope.originSystem -ne 'AadGroup' -or $roles[0].scope.isRootScope -ne $true -or
		$roles[0].scope.originId -ne $review.GroupId -or $roles[0].role.originId -ne "Member_$($review.GroupId)") {
		return 'The package is not exactly one matching group Member role; manual review is required.'
	}
	if (-not $Row.Assignment.Id -or $Row.Assignment.State -ne 'delivered' -or
		$Row.Assignment.Target.objectId -ne $review.UserId) {
		return 'The delivered assignment and its directory user could not be verified.'
	}
}

function Get-ReconciliationConflictIssue {
	param($Row, $Conflicts)
	$review = $Row.Review
	if ($Conflicts.UncertainGroups.Contains([string]$review.GroupId)) {
		return 'A review decision for this group could not be identified; conflict checking is incomplete.'
	}
	$reviewed = [datetimeoffset]::MinValue
	$hasReviewedTime = [datetimeoffset]::TryParse([string]$review.ReviewedAt, [ref]$reviewed)
	$applied = [datetimeoffset]::MinValue
	$hasAppliedTime = [datetimeoffset]::TryParse([string]$review.AppliedAt, [ref]$applied)
	foreach ($other in @($Conflicts.Records["$($review.GroupId)/$($review.UserId)"])) {
		if (-not $other) { continue }
		if ($other.ReviewId -eq $review.ReviewId -and $other.InstanceId -eq $review.InstanceId -and $other.DecisionId -eq $review.DecisionId) { continue }
		if ($other.InstanceStatus -eq 'Completed' -and $other.Decision -eq 'Deny' -and $other.ApplyResult -eq 'AppliedSuccessfully') { continue }
		$otherTime = [datetimeoffset]::MinValue
		$otherApplied = [datetimeoffset]::MinValue
		if ($other.InstanceStatus -ne 'Completed' -or -not $hasReviewedTime -or
			-not [datetimeoffset]::TryParse($other.ReviewedAt, [ref]$otherTime) -or $otherTime -ge $reviewed -or
			-not $hasAppliedTime -or -not [datetimeoffset]::TryParse($other.AppliedAt, [ref]$otherApplied) -or $otherApplied -ge $applied) {
			return 'Another review has a conflicting, later or unresolved decision for this user and group.'
		}
	}
}

function Get-ReconciliationLiveIssue {
	param($Row, $Conflicts)
	$review = $Row.Review
	$issue = Get-ReconciliationConflictIssue $Row $Conflicts
	if ($issue) { return $issue }
	$reviewed = [datetimeoffset]::MinValue
	$hasReviewedTime = [datetimeoffset]::TryParse([string]$review.ReviewedAt, [ref]$reviewed)
	$applied = [datetimeoffset]::MinValue
	$hasAppliedTime = [datetimeoffset]::TryParse([string]$review.AppliedAt, [ref]$applied)
	$definitionUri = "$reviews/$([uri]::EscapeDataString([string]$review.ReviewId))"
	$instanceUri = "$definitionUri/instances/$([uri]::EscapeDataString([string]$review.InstanceId))"
	$definition = Get-ReconciliationVerificationDetail $definitionUri
	$instance = Get-ReconciliationVerificationDetail $instanceUri
	$decision = Invoke-MgGraphRequest -Method GET -Uri "$instanceUri/decisions/$([uri]::EscapeDataString([string]$review.DecisionId))" -OutputType Hashtable -ErrorAction Stop
	if ($definition.id -ne $review.ReviewId -or -not $definition.displayName -or
		-not $definition.displayName.StartsWith($ReviewPrefix, [StringComparison]::OrdinalIgnoreCase) -or
		$instance.id -ne $review.InstanceId -or $decision.id -ne $review.DecisionId) {
		return 'Review identity or prefix changed after the inventory.'
	}
	$instanceStart = Get-ReconciliationTime $instance.startDateTime 'review instance start'
	if ($instanceStart -ge $ReviewStartedBefore) { return 'Review instance is outside the chosen window.' }
	if ($explicitReviewWindow) {
		if ($instanceStart -lt $ReviewStartedFrom) { return 'Review instance is outside the chosen window.' }
	}
	else {
		$instanceEnd = Get-ReconciliationTime $instance.endDateTime 'review instance end'
		if ($instanceEnd -lt $ReviewStartedFrom) { return 'Review instance is outside the chosen window.' }
	}
	$liveIds = Get-ReconciliationDecisionIds $decision $instance.scope
	$liveUser = $liveIds.UserId
	$liveGroup = $liveIds.GroupId
	$liveReviewed = [datetimeoffset]::MinValue
	$liveApplied = [datetimeoffset]::MinValue
	$hasLiveReviewedTime = [datetimeoffset]::TryParse([string]$decision.reviewedDateTime, [ref]$liveReviewed)
	$hasLiveAppliedTime = [datetimeoffset]::TryParse([string]$decision.appliedDateTime, [ref]$liveApplied)
	if ($liveUser -ne $review.UserId -or $liveGroup -ne $review.GroupId -or
		($decision.principal.id -and [string]$decision.principal.id -ne $liveUser) -or
		$hasReviewedTime -ne $hasLiveReviewedTime -or ($hasReviewedTime -and $liveReviewed -ne $reviewed) -or
		(-not $hasReviewedTime -and [string]$review.ReviewedAt -ne [string]$decision.reviewedDateTime) -or
		$hasAppliedTime -ne $hasLiveAppliedTime -or ($hasAppliedTime -and $applied -ne $liveApplied) -or
		(-not $hasAppliedTime -and [string]$review.AppliedAt -ne [string]$decision.appliedDateTime)) {
		return 'The review decision or its user, group or review time changed.'
	}
	$packageId = [string]$Row.PackageId
	$packageUri = "$em/accessPackages/$([uri]::EscapeDataString($packageId))"
	$package = Get-ReconciliationVerificationDetail "$packageUri`?`$expand=catalog,resourceRoleScopes(`$expand=role,scope)"
	$catalog = if ($package.catalog.id) { $package.catalog.id } else { $package.catalogId }
	if ($package.id -ne $packageId -or $catalog -ne $CatalogId.ToString() -or
		$package -isnot [System.Collections.IDictionary] -or -not $package.Contains('resourceRoleScopes') -or
		$null -eq $package.resourceRoleScopes) { return 'The package identity, catalog or resource roles could not be verified.' }
	$roles = @($package.resourceRoleScopes)
	if ($package.'resourceRoleScopes@odata.nextLink') { $roles += @(Get-ReconciliationVerificationCollection $package.'resourceRoleScopes@odata.nextLink') }
	$filter = [uri]::EscapeDataString("accessPackage/id eq '$($packageId.Replace("'", "''"))'")
	$assignments = @(Get-ReconciliationVerificationCollection "$em/assignments?`$filter=$filter&`$expand=target")
	$assignmentMatches = @($assignments | Where-Object { $_.id -eq $Row.Assignment.Id })
	if ($assignmentMatches.Count -ne 1) { return 'The assignment was not found uniquely in the package.' }
	$assignment = Invoke-MgGraphRequest -Method GET -Uri "$em/assignments/$([uri]::EscapeDataString([string]$Row.Assignment.Id))?`$expand=target,accessPackage" -OutputType Hashtable -ErrorAction Stop
	if ($assignment.id -ne $Row.Assignment.Id -or $assignment.accessPackage.id -ne $packageId) {
		return 'The assignment identity or package changed.'
	}
	if (@($assignments | Where-Object { $_.state -ne 'expired' -and -not $_.target.objectId }).Count) {
		return 'An active assignment has an unknown target; duplicate assignment checking is incomplete.'
	}
	if (@($assignments | Where-Object { $_.id -ne $assignment.id -and $_.state -ne 'expired' -and $_.target.objectId -eq $review.UserId }).Count) {
		return 'Another active assignment exists for this user in the package.'
	}
	$requests = @(Get-ReconciliationVerificationCollection "$em/assignmentRequests?`$filter=$filter&`$expand=assignment(`$expand=target)")
	$live = [pscustomobject]@{
		Review = [pscustomobject]@{ Decision = $decision.decision; ApplyResult = $decision.applyResult; InstanceStatus = $instance.status
			InstanceScope = $instance.scope; ReviewedAt = $decision.reviewedDateTime; GroupId = $liveGroup; UserId = $liveUser }
		ResourceRoles = $roles; Assignment = $assignment
	}
	$issue = Get-ReconciliationEvidenceIssue $live
	if ($issue) { return $issue }
	foreach ($request in $requests) {
		if (-not $request.id -or -not $request.state -or -not $request.requestType) {
			return 'Request history contains an incomplete record.'
		}
		$target = [string]$request.assignment.target.objectId
		$sameAssignment = $request.assignment.id -eq $assignment.id
		$sameUser = $target -eq $review.UserId
		$terminal = $request.state -in @('delivered', 'denied', 'canceled', 'deliveryFailed')
		if (-not $terminal -and (-not $target -or $sameAssignment -or $sameUser)) {
			return 'A pending or unresolved request may affect this assignment or user.'
		}
		if ($request.state -in @('denied', 'canceled', 'deliveryFailed')) { continue }
		if ($sameAssignment -and $request.requestType -in @('adminAdd', 'userAdd')) { continue }
		if ($target -and -not $sameAssignment -and -not $sameUser) { continue }
		if (-not $hasReviewedTime) { return 'A request may affect this user or assignment, but the decision review time is unavailable.' }
		$requestTime = Get-ReconciliationTime $request.createdDateTime 'request creation time'
		if ($requestTime -ge $reviewed) {
			if (-not $target -and -not $sameAssignment) { return 'A later assignment request has an unknown target; request checking is incomplete.' }
			if ($sameAssignment -or $sameUser) { return 'A later assignment request may have changed the user or assignment.' }
		}
	}
}

function Write-ReconciliationAudit {
	param($Writer, $Stream, $Record)
	$Writer.WriteLine(($Record | ConvertTo-Json -Compress -Depth 15))
	$Writer.Flush()
	$Stream.Flush($true)
}

try {
Show-ReconciliationProgress 'Connecting and reading review definitions' -Force
Write-Warning 'Reports, candidate selections and audit logs contain sensitive user and access data; store them privately.'
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$context = Get-MgContext
if (-not $context -or -not $context.TenantId) { throw 'Connect to Microsoft Graph first.' }
$environment = Get-MgEnvironment -Name $context.Environment
if (-not $environment.GraphEndpoint) { throw 'Unable to determine the connected Microsoft Graph endpoint.' }
$script:graphBase = "$($environment.GraphEndpoint.TrimEnd('/'))/v1.0"
$em = "$script:graphBase/identityGovernance/entitlementManagement"
$reviews = "$script:graphBase/identityGovernance/accessReviews/definitions"
Show-ReconciliationProgress 'Checking tenant-wide visibility' -Force
Assert-ReconciliationVisibility $context
Write-Host "Catalog: $CatalogId (all packages). Reviews overlapping $($ReviewStartedFrom.ToString('o')) through $($ReviewStartedBefore.ToString('o'))."
Write-Warning 'Pause backfill, review decisions and other assignment or package changes until remediation finishes.'
$started = [datetimeoffset]::UtcNow
$issues = [System.Collections.Generic.List[object]]::new()
$packageIssues = [System.Collections.Generic.List[object]]::new()
$decisionsByGroup = @{}
$matchedGroupIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$definitions = @(Get-ReconciliationReviewSnapshot $reviews)
$matchingDefinitions = 0
$definitionIndex = 0
foreach ($definition in $definitions) {
	$definitionIndex++
	Show-ReconciliationProgress 'Checking review definitions' $definitionIndex $definitions.Count
	if (-not $definition.id -or -not $definition.displayName) { throw 'A review definition has no ID or name; no report was produced.' }
	if (-not $definition.displayName.StartsWith($ReviewPrefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
	$matchingDefinitions++
	$definitionUri = "$reviews/$([uri]::EscapeDataString([string]$definition.id))"
	foreach ($instance in @(Get-ReconciliationReviewSnapshot "$definitionUri/instances")) {
		if (-not $instance.id) { throw 'A review instance has no ID; no report was produced.' }
		$instanceStart = Get-ReconciliationTime $instance.startDateTime 'review instance start'
		if ($instanceStart -ge $ReviewStartedBefore) { continue }
		if ($explicitReviewWindow) {
			if ($instanceStart -lt $ReviewStartedFrom) { continue }
		}
		else {
			$instanceEnd = Get-ReconciliationTime $instance.endDateTime 'review instance end'
			if ($instanceEnd -lt $ReviewStartedFrom) { continue }
		}
		$instanceUri = "$definitionUri/instances/$([uri]::EscapeDataString([string]$instance.id))"
		foreach ($decision in @(Get-ReconciliationReviewSnapshot "$instanceUri/decisions")) {
			if (-not $decision.id -or -not $decision.decision) { throw 'A review decision has no ID or decision; no report was produced.' }
			if ($decision.decision -notin @('Deny', 'NotReviewed', 'DontKnow')) { continue }
			$ids = Get-ReconciliationDecisionIds $decision $instance.scope
			$groupId = $ids.GroupId
			$userId = $ids.UserId
			if (-not $groupId -or -not $userId) {
				$issues.Add([pscustomobject]@{
					ReviewId = $definition.id; InstanceId = $instance.id; DecisionId = $decision.id
					Reason = 'Cannot identify both a group and a user from the decision or direct review scope; inspect manually.'
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
$catalogFilter = [uri]::EscapeDataString("catalog/id eq '$CatalogId'")
$packages = @(Get-ReconciliationCollection "$em/accessPackages?`$filter=$catalogFilter&`$expand=catalog")
$packagesInWindow = 0
$packageIndex = 0
foreach ($package in $packages) {
	$packageIndex++
	Show-ReconciliationProgress 'Checking access packages and assignments' $packageIndex $packages.Count
	$packageCatalogId = if ($package.catalog.id) { [string]$package.catalog.id } else { [string]$package.catalogId }
	if (-not $packageCatalogId) { throw 'A package has no catalog ID; package scope cannot be verified.' }
	if ($packageCatalogId -ne $CatalogId.ToString()) { continue }
	if (-not $package.id) { throw 'A package has no ID; no report was produced.' }
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
	if (@($assignments | Where-Object { -not $_.target.objectId }).Count) {
		$packageIssues.Add([pscustomobject]@{ PackageId = $packageId; Reason = 'An assignment target object ID was not returned; assignment matching is incomplete.' })
	}
	foreach ($decision in $matched) {
		$matchingAssignments = @($assignments | Where-Object { $_.target.objectId -eq $decision.UserId -and $_.state -ne 'expired' })
		if (-not $matchingAssignments.Count) {
			$rows.Add([pscustomobject]@{ PackageId = $packageId; PackageName = $package.displayName; PackageCreatedAt = $package.createdDateTime
				ResourceRoles = $roles; Review = $decision; Assignment = $null
				Reason = 'No matching non-expired assignment was returned. Verify visibility and pending requests manually.' })
			continue
		}
		foreach ($assignment in $matchingAssignments) {
			if (-not $assignment.id) { throw "An assignment in package '$packageId' has no ID; no report was produced." }
			$rows.Add([pscustomobject]@{ PackageId = $packageId; PackageName = $package.displayName; PackageCreatedAt = $package.createdDateTime
				ResourceRoles = $roles; Review = $decision
				Assignment = [pscustomobject]@{ Id = $assignment.id; State = $assignment.state; Target = $assignment.target; Schedule = $assignment.schedule }
				Reason = 'Not yet checked for conflicting reviews or pending requests.' })
		}
	}
}

$noPackageMatches = @(foreach ($groupId in $decisionsByGroup.Keys) {
	if (-not $matchedGroupIds.Contains($groupId)) { $decisionsByGroup[$groupId] }
})
$rowsByAssignment = @{}
foreach ($row in $rows) {
	if (-not $row.Assignment.Id) { continue }
	$id = [string]$row.Assignment.Id
	if (-not $rowsByAssignment.ContainsKey($id)) { $rowsByAssignment[$id] = [System.Collections.Generic.List[object]]::new() }
	$rowsByAssignment[$id].Add($row)
}
$preliminary = [System.Collections.Generic.List[object]]::new()
$candidates = [System.Collections.Generic.List[object]]::new()
$eligibleRows = [System.Collections.Generic.List[object]]::new()
foreach ($id in @($rowsByAssignment.Keys | Sort-Object)) {
	$matchingRows = @($rowsByAssignment[$id])
	$validRows = @(foreach ($row in $matchingRows) {
		$issue = Get-ReconciliationEvidenceIssue $row
		if ($issue) { $row.Reason = $issue } else { $row }
	})
	if ($validRows.Count) {
		$keys = @($validRows | ForEach-Object { "$($_.Review.GroupId)/$($_.Review.UserId)" } | Select-Object -Unique)
		if ($keys.Count -ne 1) {
			foreach ($row in $validRows) { $row.Reason = 'Multiple different review matches require manual review.' }
			continue
		}
		$row = @($validRows | Sort-Object {
			$time = [datetimeoffset]::MinValue
			$null = [datetimeoffset]::TryParse([string]$_.Review.ReviewedAt, [ref]$time)
			$time
		} -Descending)[0]
		$eligibleRows.Add($row)
	}
}
if ($eligibleRows.Count) {
	$conflicts = Get-ReconciliationReviewConflicts $eligibleRows.ToArray()
	$index = 0
	foreach ($row in $eligibleRows) {
		$index++
		Show-ReconciliationProgress 'Checking candidates for conflicts' $index $eligibleRows.Count
		$reason = if ($Remediate) { Get-ReconciliationConflictIssue $row $conflicts }
			else { Get-ReconciliationLiveIssue $row $conflicts }
		if ($reason) {
			foreach ($match in $rowsByAssignment[[string]$row.Assignment.Id]) {
				if ($match.Reason -eq 'Not yet checked for conflicting reviews or pending requests.') { $match.Reason = $reason }
			}
			continue
		}
		$row.Reason = if ($Remediate) { 'Preliminary candidate; rechecked before removal.' }
			else { 'Candidate for removal; review before submitting.' }
		$packageName = [string]$row.PackageName
		$userName = [string]$row.Review.UserName
		if ($packageName -match '^\s*[=+\-@]') { $packageName = "'$packageName" }
		if ($userName -match '^\s*[=+\-@]') { $userName = "'$userName" }
		$item = [pscustomobject]@{
			TenantId = [string]$context.TenantId; AssignmentId = [string]$row.Assignment.Id
			PackageName = $packageName; UserName = $userName; PackageId = [string]$row.PackageId; UserId = [string]$row.Review.UserId
			GroupId = [string]$row.Review.GroupId; ReviewId = [string]$row.Review.ReviewId
			InstanceId = [string]$row.Review.InstanceId; DecisionId = [string]$row.Review.DecisionId
			ReviewedAt = [string]$row.Review.ReviewedAt; AppliedAt = [string]$row.Review.AppliedAt
		}
		foreach ($property in $item.PSObject.Properties) {
			if ([string]$property.Value -match '^\s*[=+\-@]') { $property.Value = "'$($property.Value)" }
		}
		$preliminary.Add($item)
		$candidates.Add($row)
		foreach ($match in $rowsByAssignment[[string]$row.Assignment.Id]) {
			if ($match.Reason -eq 'Not yet checked for conflicting reviews or pending requests.') { $match.Reason = $row.Reason }
		}
	}
}
$fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
$selectionPath = "$fullPath.candidates.csv"
if (Test-Path -LiteralPath $selectionPath) { throw 'The candidates CSV path already exists; choose a different report path.' }
$report = [ordered]@{
	SchemaVersion = 1; TenantId = [string]$context.TenantId; StartedAt = $started.ToString('o'); FinishedAt = [datetimeoffset]::UtcNow.ToString('o')
	CatalogId = $CatalogId.ToString()
	ReviewPrefix = $ReviewPrefix; ReviewStartedFrom = $ReviewStartedFrom.ToString('o'); ReviewStartedBefore = $ReviewStartedBefore.ToString('o')
	ReviewWindowBasis = if ($explicitReviewWindow) { 'InstanceStart' } else { 'ScheduledPeriodOverlap' }
	MatchingDefinitions = $matchingDefinitions; PackagesInWindow = $packagesInWindow
	PreliminaryCandidateCount = $preliminary.Count; PreliminarySelectionPath = $selectionPath
	Limitations = @('All packages in the specified catalog are checked, regardless of creation date.', 'Review instances outside the selected review window may be missed.', 'Concurrent or out-of-band changes may not be detected.', 'Results depend on Graph permissions and visibility.', 'For remediation, candidates are preliminary; consult the audit for revalidation and request outcomes.')
	UnmatchedDecisions = @($issues.ToArray()); PackageIssues = @($packageIssues.ToArray())
	DecisionsWithoutMatchingPackage = $noPackageMatches; Results = @($rows.ToArray())
}
$json = $report | ConvertTo-Json -Depth 40
$stream = [System.IO.File]::Open($fullPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
try {
	$writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
	try { $writer.Write($json) } finally { $writer.Dispose() }
} finally { $stream.Dispose() }
Write-Host "Inventory report saved to $fullPath. $($rows.Count) result(s); $($issues.Count) unidentified decision(s); $($packageIssues.Count) package issue(s)."
if ($selectionPath) {
	$csvLines = if ($preliminary.Count) { @($preliminary | ConvertTo-Csv -NoTypeInformation) } else { @('"TenantId","AssignmentId","PackageName","UserName","PackageId","UserId","GroupId","ReviewId","InstanceId","DecisionId","ReviewedAt","AppliedAt"') }
	$selectionStream = [System.IO.File]::Open($selectionPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
	try {
		$selectionWriter = [System.IO.StreamWriter]::new($selectionStream, [System.Text.UTF8Encoding]::new($false))
		try { foreach ($line in $csvLines) { $selectionWriter.WriteLine($line) } } finally { $selectionWriter.Dispose() }
	}
	finally { $selectionStream.Dispose() }
	Write-Host "$($preliminary.Count) candidate(s) saved to $selectionPath."
	if ($Remediate -and $preliminary.Count) {
		Write-Host 'These candidates are preliminary; the audit records the final checks and request outcomes.'
	}
	$candidateIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
	foreach ($candidate in $candidates) { $null = $candidateIds.Add([string]$candidate.Assignment.Id) }
	$manualRows = @($rows | Where-Object {
		$_.Assignment.Id -and $_.Review.Decision -eq 'Deny' -and $_.Review.ApplyResult -eq 'AppliedSuccessfully' -and
			-not $candidateIds.Contains([string]$_.Assignment.Id)
	})
	$manualCount = @($manualRows | ForEach-Object { $_.Assignment.Id } | Select-Object -Unique).Count
	if ($manualCount) {
		Write-Warning "$manualCount assignment(s) with an applied Deny were not eligible for automatic removal."
		foreach ($group in @($manualRows | Group-Object Reason | Sort-Object Count -Descending)) {
			$reasonCount = @($group.Group | ForEach-Object { $_.Assignment.Id } | Select-Object -Unique).Count
			Write-Host "  ${reasonCount}: $($group.Name)"
		}
		Write-Host 'Counts by reason can overlap; see the JSON report for details.'
	}
	$otherCount = @($rows | Where-Object {
		$_.Assignment.Id -and ($_.Review.Decision -ne 'Deny' -or $_.Review.ApplyResult -ne 'AppliedSuccessfully') -and
			-not $candidateIds.Contains([string]$_.Assignment.Id)
	} | ForEach-Object { $_.Assignment.Id } | Select-Object -Unique).Count
	if ($otherCount) {
		Write-Host "$otherCount assignment(s) also matched decisions that were not confirmed applied Deny; these are not automatic removal candidates."
	}
}
if (-not $preliminary.Count) {
	Write-Warning 'No automatic candidates were found. Check the report for manual-review issues and scope limitations; an empty CSV does not prove there are no affected assignments.'
}
if (-not $matchingDefinitions -or -not $packagesInWindow) {
	Write-Warning "Scope returned $matchingDefinitions matching review definition(s) and $packagesInWindow package(s). Check the scope and Graph visibility; an empty report does not prove there are no affected assignments."
}
if ($Remediate) {
	$auditFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($AuditPath)
	$auditStream = [System.IO.File]::Open($auditFullPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
	try {
		Write-Host "Audit log: $auditFullPath"
		$auditWriter = [System.IO.StreamWriter]::new($auditStream, [System.Text.UTF8Encoding]::new($false))
		try {
			$baseAudit = @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Started'
				TenantId = [string]$context.TenantId; ReportPath = $fullPath; Mode = if ($Remediate -and -not $WhatIfPreference) { 'Remediate' } else { 'Preview' }
				CandidateCount = $candidates.Count }
			Write-ReconciliationAudit $auditWriter $auditStream $baseAudit
			$skipped = 0
			$skipReasons = @{}
			foreach ($id in @($rowsByAssignment.Keys | Sort-Object)) {
				if ($candidateIds.Contains($id)) { continue }
				$denials = @($rowsByAssignment[$id] | Where-Object { $_.Review.Decision -eq 'Deny' } |
					Sort-Object { $_.Review.ApplyResult -ne 'AppliedSuccessfully' })
				if (-not $denials.Count) { continue }
				Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Skipped'; AssignmentId = $id; Reason = $denials[0].Reason }
				$skipReasons[$denials[0].Reason]++
				$skipped++
			}
			$submitted = 0
			$previewed = 0
			if ($candidates.Count) {
				$proceed = $PSCmdlet.ShouldProcess("up to $($candidates.Count) assignments in catalog $CatalogId", 'Submit administrative removal requests')
				$index = 0
				foreach ($candidate in @($candidates | Sort-Object PackageId)) {
					$index++
					Show-ReconciliationProgress 'Checking and processing assignments' $index $candidates.Count
					$id = [string]$candidate.Assignment.Id
					try { $reason = Get-ReconciliationLiveIssue $candidate $conflicts }
					catch {
						Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Error'; AssignmentId = $id; Reason = $_.Exception.Message }
						throw
					}
					if ($reason) {
						Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Skipped'; AssignmentId = $id; Reason = $reason }
						$skipReasons[$reason]++
						$skipped++
						continue
					}
					if (-not $proceed) {
						$type = if ($WhatIfPreference) { 'Preview' } else { 'Declined' }
						Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = $type; AssignmentId = $id; PackageId = $candidate.PackageId; UserId = $candidate.Review.UserId; GroupId = $candidate.Review.GroupId }
						if ($WhatIfPreference) { $previewed++ } else { $skipped++ }
						continue
					}
					Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Submitting'; AssignmentId = $id; PackageId = $candidate.PackageId; UserId = $candidate.Review.UserId; GroupId = $candidate.Review.GroupId }
					$body = @{ requestType = 'adminRemove'; assignment = @{ id = $id } } | ConvertTo-Json -Compress -Depth 5
					try {
						$response = Invoke-MgGraphRequest -Method POST -Uri "$em/assignmentRequests" -Body $body -ContentType 'application/json' -OutputType Hashtable -ErrorAction Stop
					}
					catch {
						Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Unknown'; AssignmentId = $id; Reason = $_.Exception.Message }
						throw "Removal request outcome is unknown for assignment '$id'. Check Entra and the audit before retrying. $($_.Exception.Message)"
					}
					if (-not $response.id -or -not $response.state) {
						Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Unknown'; AssignmentId = $id; Reason = 'Graph did not return a request ID and state.' }
						throw "Removal request outcome is unknown for assignment '$id'. Check Entra and the audit before retrying."
					}
					if ($response.state -in @('denied', 'canceled', 'deliveryFailed')) {
						Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Rejected'; AssignmentId = $id; RequestId = $response.id; RequestState = $response.state }
						Write-Warning "Removal request '$($response.id)' was not accepted (state: $($response.state)); inspect the audit."
						$skipped++
						continue
					}
					if ($response.state -notin @('submitted', 'pendingApproval', 'delivering', 'delivered', 'partiallyDelivered', 'scheduled')) {
						Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Unknown'; AssignmentId = $id; RequestId = $response.id; RequestState = $response.state }
						throw "Removal request has an unrecognized state for assignment '$id'. Check Entra and the audit before retrying."
					}
					Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Submitted'; AssignmentId = $id; RequestId = $response.id; RequestState = $response.state }
					$submitted++
				}
			}
			Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Finished'; Submitted = $submitted; Previewed = $previewed; Skipped = $skipped }
			Write-Host "Processing complete: $submitted request(s) submitted, $previewed previewed, $skipped skipped. Audit: $auditFullPath"
			foreach ($reason in @($skipReasons.Keys | Sort-Object)) { Write-Host "  $($skipReasons[$reason]): $reason" }
			if ($submitted) { Write-Warning 'Submitted requests are not confirmation of removal. Verify each request and group membership in Entra.' }
		}
		finally { $auditWriter.Dispose() }
	}
	finally { $auditStream.Dispose() }
}
}
finally {
	Write-Progress -Activity 'Reconciling access reviews and assignments' -Completed
}
