<#
.SYNOPSIS
Investigate and optionally remediate access review and package assignment conflicts.
.DESCRIPTION
Without -Remediate, this script makes no Graph writes. The initial inventory
creates a preliminary candidates CSV beside the JSON report. Review the report
and CSV and remove any rows that are not approved, keeping all columns intact.
Run again with -ApprovalCsvPath to preview the selection. Add -PreviewAuditPath
pointing to the completed preview audit when remediating. Add
-Remediate and -ConfirmChangeFreeze only after review; -WhatIf previews even
with -Remediate. Run only one remediation at a time, and pause backfill, review
decisions and other assignment or package changes until it finishes.

Bulk removal is restricted to delivered assignments for one group Member role,
an applied Deny in a completed direct group review, and a delivered matching
adminAdd request in the incident window. A matching adminAdd does not prove
which script created it. Conflicting or incomplete evidence is not removed.
Groups used by more than one package are left for manual review.
Other reviews of the same user and group are left for manual review unless
their decisions and applications can both be shown to precede the denial.
Each assignment and its evidence are re-read before submission. A separate
tenant-wide review scan checks for conflicting decisions. Other actors can
still change state between the checks and the request; Graph offers no
transactional lock. The script cannot prove which script made an adminAdd or
detect every out-of-band approval or membership change. Review those risks
before approving the selection. Inspect the audit, verify request delivery and
group membership afterward, and never retry an uncertain request without checking it.

Reports, selection files and audit logs contain sensitive identities. Store them
privately. -LookbackDays defaults to 30. It scopes package and request creation
times and review instances whose scheduled periods overlap that window.
An instance applied late after its scheduled period may fall outside this window.
Advanced explicit time windows include the start and exclude the end. For
delegated approval processing, connect with AccessReview.Read.All,
EntitlementManagement.ReadWrite.All and RoleManagement.Read.Directory and
activate a directly assigned tenant-wide Identity Governance Administrator
or Global Administrator role. Application authentication requires tenant-wide
AccessReview.Read.All and EntitlementManagement.ReadWrite.All application
permissions for remediation. Preview accepts EntitlementManagement.Read.All.
.EXAMPLE
Connect-MgGraph -Scopes 'AccessReview.Read.All','EntitlementManagement.ReadWrite.All','RoleManagement.Read.Directory'
$scope = @{ CatalogId = '<catalog ID>'; ReviewPrefix = 'AR '; LookbackDays = 30 }
.\Get-ReviewAssignmentReconciliation.ps1 @scope -ReportPath '.\private\inventory.json'
# Review the inventory and its inventory.json.candidates.csv file. Edit the
# candidate CSV to remove rows that should not be included, preserving columns.
.\Get-ReviewAssignmentReconciliation.ps1 @scope -ApprovalCsvPath '.\private\inventory.json.candidates.csv' `
    -AuditPath '.\private\preview.jsonl'
# Inspect preview and skip reasons before proceeding:
Get-Content '.\private\preview.jsonl' | ForEach-Object { $_ | ConvertFrom-Json } |
    Format-Table Type, AssignmentId, Reason, RequestState
# Review the preview audit. After pausing other changes, submit approved removals:
.\Get-ReviewAssignmentReconciliation.ps1 @scope -ApprovalCsvPath '.\private\inventory.json.candidates.csv' `
    -PreviewAuditPath '.\private\preview.jsonl' -AuditPath '.\private\remediation.jsonl' -Remediate -ConfirmChangeFreeze
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
	[Parameter(Mandatory)][guid]$CatalogId,
	[ValidateRange(1, 36500)][int]$LookbackDays = 30,
	[datetimeoffset]$PackageCreatedFrom,
	[datetimeoffset]$PackageCreatedBefore,
	[Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$ReviewPrefix,
	[datetimeoffset]$ReviewStartedFrom,
	[datetimeoffset]$ReviewStartedBefore,
	[datetimeoffset]$RequestCreatedFrom,
	[datetimeoffset]$RequestCreatedBefore,
	[ValidateNotNullOrEmpty()][string]$ReportPath = (Join-Path (Get-Location).Path "reconciliation-$([datetime]::UtcNow.ToString('yyyyMMddTHHmmssfff'))-$([guid]::NewGuid().ToString('N')).json"),
	[string]$ApprovalCsvPath,
	[string]$AuditPath,
	[string]$PreviewAuditPath,
	[switch]$Remediate,
	[switch]$ConfirmChangeFreeze
)

$ErrorActionPreference = 'Stop'
$windowEnd = [datetimeoffset]::UtcNow
$windowStart = $windowEnd.AddDays(-$LookbackDays)
$explicitReviewWindow = $PSBoundParameters.ContainsKey('ReviewStartedFrom') -or $PSBoundParameters.ContainsKey('ReviewStartedBefore')
if ($explicitReviewWindow -and
	(-not $PSBoundParameters.ContainsKey('ReviewStartedFrom') -or -not $PSBoundParameters.ContainsKey('ReviewStartedBefore'))) {
	throw 'Supply both ReviewStartedFrom and ReviewStartedBefore when overriding the review window.'
}
if (-not $PSBoundParameters.ContainsKey('PackageCreatedFrom')) { $PackageCreatedFrom = $windowStart }
if (-not $PSBoundParameters.ContainsKey('PackageCreatedBefore')) { $PackageCreatedBefore = $windowEnd }
if (-not $explicitReviewWindow) { $ReviewStartedFrom = $windowStart; $ReviewStartedBefore = $windowEnd }
if (-not $PSBoundParameters.ContainsKey('RequestCreatedFrom')) { $RequestCreatedFrom = $windowStart }
if (-not $PSBoundParameters.ContainsKey('RequestCreatedBefore')) { $RequestCreatedBefore = $windowEnd }
if ($CatalogId -eq [guid]::Empty -or $PackageCreatedFrom -ge $PackageCreatedBefore -or
	$ReviewStartedFrom -ge $ReviewStartedBefore -or $RequestCreatedFrom -ge $RequestCreatedBefore) {
	throw 'Supply a non-empty catalog ID and valid start/end windows.'
}
if (Test-Path -LiteralPath $ReportPath) { throw 'The report path already exists; choose a new file.' }
if ($Remediate -and (-not $ApprovalCsvPath -or -not $ConfirmChangeFreeze)) {
	throw 'Remediation requires an approved assignment CSV and -ConfirmChangeFreeze. Pause concurrent reviews, backfill and assignment or resource changes first.'
}
if ($Remediate -and -not $WhatIfPreference -and -not $PreviewAuditPath) {
	throw 'Remediation requires the completed no-write preview audit via -PreviewAuditPath.'
}
if ($ApprovalCsvPath -and -not $AuditPath) {
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
$approvedIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$approvedEvidence = @{}
$selectionHash = $null
if ($ApprovalCsvPath) {
	$selectionFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ApprovalCsvPath)
	$selectionBytes = [System.IO.File]::ReadAllBytes($selectionFullPath)
	$sha = [System.Security.Cryptography.SHA256]::Create()
	try { $selectionHash = [BitConverter]::ToString($sha.ComputeHash($selectionBytes)).Replace('-', '') }
	finally { $sha.Dispose() }
	$selectionText = [System.Text.Encoding]::UTF8.GetString($selectionBytes).TrimStart([char]0xFEFF)
	$selection = @($selectionText | ConvertFrom-Csv)
	if (-not $selection.Count) { throw 'The approval CSV is empty.' }
	foreach ($entry in $selection) {
		$id = [string]$entry.AssignmentId
		if (-not $id -or $id -ne $id.Trim()) { throw 'Every approval CSV row must have a non-empty AssignmentId without surrounding whitespace.' }
		if (-not $approvedIds.Add($id)) { throw "Duplicate assignment ID in the approval CSV: '$id'." }
		foreach ($field in @('PackageId', 'UserId', 'GroupId', 'ReviewId', 'InstanceId', 'DecisionId', 'AdminAddRequestId', 'ReviewedAt')) {
			$value = [string]$entry.$field
			if (-not $value -or $value -ne $value.Trim()) { throw "Every approval CSV row must retain a non-empty $field without surrounding whitespace." }
		}
		if ($entry.PSObject.Properties.Name -notcontains 'AppliedAt') { throw 'Every approval CSV row must retain the AppliedAt column, even when it is empty.' }
		$approvedEvidence[$id] = $entry
	}
}
$scopeDescriptor = [ordered]@{ CatalogId = $CatalogId.ToString(); ReviewPrefix = $ReviewPrefix; LookbackDays = $LookbackDays }
foreach ($key in @('PackageCreatedFrom', 'PackageCreatedBefore', 'ReviewStartedFrom', 'ReviewStartedBefore', 'RequestCreatedFrom', 'RequestCreatedBefore')) {
	$scopeDescriptor[$key] = if ($PSBoundParameters.ContainsKey($key)) { ([datetimeoffset]$PSBoundParameters[$key]).ToString('o') } else { $null }
}
$scopeKey = $scopeDescriptor | ConvertTo-Json -Compress
$previewedIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

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
	if ($Link -notmatch "/v1\.0/$Kind/([0-9a-fA-F-]{36})(?:/?(?:\?.*)?)$") { return $null }
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

function Get-ReconciliationReviewConflicts {
	param([object[]]$Candidates)
	$groups = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
	foreach ($candidate in $Candidates) { $null = $groups.Add([string]$candidate.Review.GroupId) }
	$records = @{}
	$uncertainGroups = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
	Show-ReconciliationProgress 'Scanning all reviews for conflicting decisions' -Force
	$allDefinitions = @(Get-ReconciliationCollection $reviews)
	$index = 0
	foreach ($definition in $allDefinitions) {
		$index++
		Show-ReconciliationProgress 'Scanning all reviews for conflicting decisions' $index $allDefinitions.Count
		if (-not $definition.id) { throw 'A review definition has no ID; conflict checking is incomplete.' }
		$definitionUri = "$reviews/$([uri]::EscapeDataString([string]$definition.id))"
		foreach ($instance in @(Get-ReconciliationCollection "$definitionUri/instances")) {
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
			foreach ($decision in @(Get-ReconciliationCollection "$instanceUri/decisions")) {
				if (-not $decision.id) { throw 'A review decision has no ID; conflict checking is incomplete.' }
				$groupId = Get-ReconciliationLinkId ([string]$decision.resourceLink) 'groups'
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
				$userId = Get-ReconciliationLinkId ([string]$decision.principalLink) 'users'
				if (-not $userId -or ($decision.principal.id -and [string]$decision.principal.id -ne $userId)) {
					if ($decision.principal.'@odata.type' -match 'servicePrincipalIdentity') { continue }
					$null = $uncertainGroups.Add($groupId)
					continue
				}
				$key = "$groupId/$userId"
				if (-not $records.ContainsKey($key)) { $records[$key] = [System.Collections.Generic.List[object]]::new() }
				$records[$key].Add([pscustomobject]@{
					GroupId = $groupId; UserId = $userId; ReviewId = [string]$definition.id
					InstanceId = [string]$instance.id; DecisionId = [string]$decision.id
					InstanceStatus = [string]$instance.status; ReviewedAt = [string]$decision.reviewedDateTime
					AppliedAt = [string]$decision.appliedDateTime
				})
			}
		}
	}
	[pscustomobject]@{ Records = $records; UncertainGroups = $uncertainGroups }
}

function Get-ReconciliationSharedGroups {
	param([object[]]$Candidates, [object[]]$Packages)
	$groupPackages = @{}
	foreach ($candidate in $Candidates) {
		$groupId = [string]$candidate.Review.GroupId
		if (-not $groupPackages.ContainsKey($groupId)) {
			$groupPackages[$groupId] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
		}
	}
	Show-ReconciliationProgress 'Checking for groups shared with other packages' -Force
	$index = 0
	foreach ($package in $Packages) {
		$index++
		Show-ReconciliationProgress 'Checking for groups shared with other packages' $index $Packages.Count
		if (-not $package.id) { throw 'A package has no ID; shared group checking is incomplete.' }
		$packageId = [string]$package.id
		$detail = Invoke-MgGraphRequest -Method GET -Uri "$em/accessPackages/$([uri]::EscapeDataString($packageId))?`$expand=resourceRoleScopes(`$expand=role,scope)" -OutputType Hashtable -ErrorAction Stop
		if ($detail -isnot [System.Collections.IDictionary] -or $detail.id -ne $packageId -or
			-not $detail.Contains('resourceRoleScopes') -or $null -eq $detail.resourceRoleScopes) {
			throw "Package '$packageId' has incomplete resource data; shared group checking is incomplete."
		}
		$roles = @($detail.resourceRoleScopes)
		if ($detail.'resourceRoleScopes@odata.nextLink') { $roles += @(Get-ReconciliationCollection $detail.'resourceRoleScopes@odata.nextLink') }
		foreach ($role in $roles) {
			if (-not $role.role.originSystem -or -not $role.role.originId -or -not $role.scope.originSystem -or -not $role.scope.originId) {
				throw "Package '$packageId' has an incomplete resource role or scope; shared group checking is incomplete."
			}
			$ids = [System.Collections.Generic.List[string]]::new()
			if ($role.scope.originSystem -eq 'AadGroup') { $ids.Add([string]$role.scope.originId) }
			if ($role.role.originSystem -eq 'AadGroup') {
				if ($role.role.originId -notmatch '^[^_]+_([0-9a-fA-F-]{36})$') {
					throw "Package '$packageId' has an unidentifiable group role; shared group checking is incomplete."
				}
				$ids.Add($Matches[1])
			}
			foreach ($value in $ids) {
				$group = [guid]::Empty
				if (-not [guid]::TryParse($value, [ref]$group) -or $group -eq [guid]::Empty) {
					throw "Package '$packageId' has an unidentifiable group ID; shared group checking is incomplete."
				}
				$key = $group.ToString()
				if ($groupPackages.ContainsKey($key)) { $null = $groupPackages[$key].Add($packageId) }
			}
		}
	}
	$shared = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
	foreach ($key in $groupPackages.Keys) {
		if ($groupPackages[$key].Count -gt 1) { $null = $shared.Add($key) }
	}
	return ,$shared
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
	$reviewed = [datetimeoffset]::MinValue
	if (-not [datetimeoffset]::TryParse([string]$review.ReviewedAt, [ref]$reviewed)) {
		return 'The decision has no usable review time for conflict checking.'
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
	$origins = @($Row.IncidentAdminAddRequests | Where-Object { $_.Id -and $_.State -eq 'delivered' })
	if ($origins.Count -ne 1) { return 'Exactly one delivered incident-window adminAdd linked to this assignment is required.' }
}

function Get-ReconciliationLiveIssue {
	param($Row, $Conflicts, $SharedGroups)
	$review = $Row.Review
	if ($SharedGroups.Contains([string]$review.GroupId)) { return 'The group is used by another access package; manual review is required.' }
	if ($Conflicts.UncertainGroups.Contains([string]$review.GroupId)) {
		return 'A review decision for this group could not be identified; conflict checking is incomplete.'
	}
	$reviewed = [datetimeoffset]::MinValue
	if (-not [datetimeoffset]::TryParse([string]$review.ReviewedAt, [ref]$reviewed)) { return 'Decision review time is unavailable.' }
	$applied = [datetimeoffset]::MinValue
	$hasAppliedTime = [datetimeoffset]::TryParse([string]$review.AppliedAt, [ref]$applied)
	foreach ($other in @($Conflicts.Records["$($review.GroupId)/$($review.UserId)"])) {
		if (-not $other) { continue }
		if ($other.ReviewId -eq $review.ReviewId -and $other.InstanceId -eq $review.InstanceId -and $other.DecisionId -eq $review.DecisionId) { continue }
		$otherTime = [datetimeoffset]::MinValue
		$otherApplied = [datetimeoffset]::MinValue
		if ($other.InstanceStatus -ne 'Completed' -or
			-not [datetimeoffset]::TryParse($other.ReviewedAt, [ref]$otherTime) -or $otherTime -ge $reviewed -or
			-not $hasAppliedTime -or -not [datetimeoffset]::TryParse($other.AppliedAt, [ref]$otherApplied) -or $otherApplied -ge $applied) {
			return 'Another review has a conflicting, later or unresolved decision for this user and group.'
		}
	}
	$definitionUri = "$reviews/$([uri]::EscapeDataString([string]$review.ReviewId))"
	$instanceUri = "$definitionUri/instances/$([uri]::EscapeDataString([string]$review.InstanceId))"
	$definition = Invoke-MgGraphRequest -Method GET -Uri $definitionUri -OutputType Hashtable -ErrorAction Stop
	$instance = Invoke-MgGraphRequest -Method GET -Uri $instanceUri -OutputType Hashtable -ErrorAction Stop
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
	$liveUser = Get-ReconciliationLinkId ([string]$decision.principalLink) 'users'
	$liveGroup = Get-ReconciliationLinkId ([string]$decision.resourceLink) 'groups'
	$liveReviewed = [datetimeoffset]::MinValue
	$liveApplied = [datetimeoffset]::MinValue
	$hasLiveAppliedTime = [datetimeoffset]::TryParse([string]$decision.appliedDateTime, [ref]$liveApplied)
	if ($liveUser -ne $review.UserId -or $liveGroup -ne $review.GroupId -or
		($decision.principal.id -and [string]$decision.principal.id -ne $liveUser) -or
		-not [datetimeoffset]::TryParse([string]$decision.reviewedDateTime, [ref]$liveReviewed) -or $liveReviewed -ne $reviewed -or
		$hasAppliedTime -ne $hasLiveAppliedTime -or ($hasAppliedTime -and $applied -ne $liveApplied) -or
		(-not $hasAppliedTime -and [string]$review.AppliedAt -ne [string]$decision.appliedDateTime)) {
		return 'The review decision or its user, group or review time changed.'
	}
	$packageId = [string]$Row.PackageId
	$packageUri = "$em/accessPackages/$([uri]::EscapeDataString($packageId))"
	$package = Invoke-MgGraphRequest -Method GET -Uri "$packageUri`?`$expand=catalog,resourceRoleScopes(`$expand=role,scope)" -OutputType Hashtable -ErrorAction Stop
	$catalog = if ($package.catalog.id) { $package.catalog.id } else { $package.catalogId }
	if ($package.id -ne $packageId -or $catalog -ne $CatalogId.ToString() -or
		$package -isnot [System.Collections.IDictionary] -or -not $package.Contains('resourceRoleScopes') -or
		$null -eq $package.resourceRoleScopes) { return 'The package identity, catalog or resource roles could not be verified.' }
	$created = Get-ReconciliationTime $package.createdDateTime 'package creation time'
	if ($created -lt $PackageCreatedFrom -or $created -ge $PackageCreatedBefore) { return 'The package is outside the chosen creation window.' }
	$roles = @($package.resourceRoleScopes)
	if ($package.'resourceRoleScopes@odata.nextLink') { $roles += @(Get-ReconciliationCollection $package.'resourceRoleScopes@odata.nextLink') }
	$filter = [uri]::EscapeDataString("accessPackage/id eq '$($packageId.Replace("'", "''"))'")
	$assignments = @(Get-ReconciliationCollection "$em/assignments?`$filter=$filter&`$expand=target")
	$assignmentMatches = @($assignments | Where-Object { $_.id -eq $Row.Assignment.Id })
	if ($assignmentMatches.Count -ne 1) { return 'The assignment was not found uniquely in the package.' }
	$assignment = $assignmentMatches[0]
	if (@($assignments | Where-Object { $_.state -ne 'expired' -and -not $_.target.objectId }).Count) {
		return 'An active assignment has an unknown target; duplicate assignment checking is incomplete.'
	}
	if (@($assignments | Where-Object { $_.id -ne $assignment.id -and $_.state -ne 'expired' -and $_.target.objectId -eq $review.UserId }).Count) {
		return 'Another active assignment exists for this user in the package.'
	}
	$requests = @(Get-ReconciliationCollection "$em/assignmentRequests?`$filter=$filter&`$expand=assignment(`$expand=target)")
	$origins = @(
		foreach ($request in $requests) {
			if ($request.requestType -ne 'adminAdd' -or $request.assignment.id -ne $assignment.id -or $request.state -ne 'delivered') { continue }
			$requestTime = Get-ReconciliationTime $request.createdDateTime 'request creation time'
			if ($requestTime -ge $RequestCreatedFrom -and $requestTime -lt $RequestCreatedBefore) {
				[pscustomobject]@{ Id = $request.id; State = $request.state; CreatedAt = $requestTime.ToString('o') }
			}
		}
	)
	$live = [pscustomobject]@{
		Review = [pscustomobject]@{ Decision = $decision.decision; ApplyResult = $decision.applyResult; InstanceStatus = $instance.status
			InstanceScope = $instance.scope; ReviewedAt = $decision.reviewedDateTime; GroupId = $liveGroup; UserId = $liveUser }
		ResourceRoles = $roles; Assignment = $assignment; IncidentAdminAddRequests = $origins
	}
	$issue = Get-ReconciliationEvidenceIssue $live
	if ($issue) { return $issue }
	$origin = $origins[0]
	$originalOrigin = @($Row.IncidentAdminAddRequests | Where-Object { $_.Id -and $_.State -eq 'delivered' })[0]
	if ($origin.Id -ne $originalOrigin.Id) { return 'The incident adminAdd evidence changed.' }
	$originTime = Get-ReconciliationTime $origin.CreatedAt 'incident request creation time'
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
		if ($request.id -eq $origin.Id -or $request.state -in @('denied', 'canceled', 'deliveryFailed')) { continue }
		$requestTime = Get-ReconciliationTime $request.createdDateTime 'request creation time'
		if ($requestTime -ge $originTime) {
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
if ($ApprovalCsvPath) {
	Show-ReconciliationProgress 'Checking tenant-wide visibility' -Force
	Assert-ReconciliationVisibility $context
}
if ($Remediate -and -not $WhatIfPreference) {
	$previewRecords = @(Get-Content -LiteralPath $PreviewAuditPath | ForEach-Object { $_ | ConvertFrom-Json })
	if ($previewRecords.Count -lt 2 -or $previewRecords[0].Type -ne 'Started' -or
		$previewRecords[0].Mode -ne 'Preview' -or $previewRecords[-1].Type -ne 'Finished' -or
		$previewRecords[0].TenantId -ne $context.TenantId -or
		$previewRecords[0].SelectionHash -ne $selectionHash -or $previewRecords[0].ScopeKey -ne $scopeKey) {
		throw 'The preview audit is incomplete or does not match the tenant, scope and unchanged approval CSV. Run a fresh preview.'
	}
	foreach ($record in @($previewRecords | Where-Object Type -EQ 'Preview')) {
		if (-not $record.AssignmentId -or -not $approvedIds.Contains([string]$record.AssignmentId) -or
			-not $previewedIds.Add([string]$record.AssignmentId)) {
			throw 'The preview audit contains an invalid or duplicate assignment. Run a fresh preview.'
		}
	}
	if ($previewedIds.Count -ne [int]$previewRecords[-1].Previewed) { throw 'The preview audit counts do not match. Run a fresh preview.' }
}
$started = [datetimeoffset]::UtcNow
$issues = [System.Collections.Generic.List[object]]::new()
$packageIssues = [System.Collections.Generic.List[object]]::new()
$decisionsByGroup = @{}
$matchedGroupIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$definitions = @(Get-ReconciliationCollection $reviews)
$matchingDefinitions = 0
$definitionIndex = 0
foreach ($definition in $definitions) {
	$definitionIndex++
	Show-ReconciliationProgress 'Checking review definitions' $definitionIndex $definitions.Count
	if (-not $definition.id -or -not $definition.displayName) { throw 'A review definition has no ID or name; no report was produced.' }
	if (-not $definition.displayName.StartsWith($ReviewPrefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
	$matchingDefinitions++
	$definitionUri = "$reviews/$([uri]::EscapeDataString([string]$definition.id))"
	foreach ($instance in @(Get-ReconciliationCollection "$definitionUri/instances")) {
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
$packageIndex = 0
foreach ($package in $packages) {
	$packageIndex++
	Show-ReconciliationProgress 'Checking access packages and assignments' $packageIndex $packages.Count
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
$rowsByAssignment = @{}
foreach ($row in $rows) {
	if (-not $row.Assignment.Id) { continue }
	$id = [string]$row.Assignment.Id
	if (-not $rowsByAssignment.ContainsKey($id)) { $rowsByAssignment[$id] = [System.Collections.Generic.List[object]]::new() }
	$rowsByAssignment[$id].Add($row)
}
$preliminary = [System.Collections.Generic.List[object]]::new()
foreach ($id in @($rowsByAssignment.Keys | Sort-Object)) {
	$matchingRows = @($rowsByAssignment[$id])
	if ($matchingRows.Count -eq 1 -and -not (Get-ReconciliationEvidenceIssue $matchingRows[0])) {
		$row = $matchingRows[0]
		$origin = @($row.IncidentAdminAddRequests | Where-Object { $_.Id -and $_.State -eq 'delivered' })[0]
		$item = [pscustomobject]@{
			AssignmentId = $id; PackageId = [string]$row.PackageId; UserId = [string]$row.Review.UserId
			GroupId = [string]$row.Review.GroupId; ReviewId = [string]$row.Review.ReviewId
			InstanceId = [string]$row.Review.InstanceId; DecisionId = [string]$row.Review.DecisionId
			AdminAddRequestId = [string]$origin.Id; ReviewedAt = [string]$row.Review.ReviewedAt; AppliedAt = [string]$row.Review.AppliedAt
		}
		if (@($item.PSObject.Properties | Where-Object { [string]$_.Value -match '^[\s]*[=+\-@]' }).Count) {
			Write-Warning 'An identifier is unsafe for a spreadsheet CSV and was excluded from the preliminary selection; inspect the JSON report.'
			continue
		}
		$preliminary.Add($item)
	}
}
$fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)
$selectionPath = if (-not $ApprovalCsvPath) { "$fullPath.candidates.csv" } else { $null }
if ($selectionPath -and (Test-Path -LiteralPath $selectionPath)) { throw 'The preliminary candidates CSV path already exists; choose a different report path.' }
$report = [ordered]@{
	SchemaVersion = 1; TenantId = [string]$context.TenantId; StartedAt = $started.ToString('o'); FinishedAt = [datetimeoffset]::UtcNow.ToString('o')
	CatalogId = $CatalogId.ToString()
	PackageCreatedFrom = $PackageCreatedFrom.ToString('o'); PackageCreatedBefore = $PackageCreatedBefore.ToString('o')
	ReviewPrefix = $ReviewPrefix; ReviewStartedFrom = $ReviewStartedFrom.ToString('o'); ReviewStartedBefore = $ReviewStartedBefore.ToString('o')
	ReviewWindowBasis = if ($explicitReviewWindow) { 'InstanceStart' } else { 'ScheduledPeriodOverlap' }
	RequestCreatedFrom = $RequestCreatedFrom.ToString('o'); RequestCreatedBefore = $RequestCreatedBefore.ToString('o')
	MatchingDefinitions = $matchingDefinitions; PackagesInWindow = $packagesInWindow
	PreliminaryCandidateCount = $preliminary.Count; PreliminarySelectionPath = $selectionPath
	Limitations = @('Read-only inventory; no row is approved for removal.', 'A matching adminAdd request does not prove it came from the backfill script.', 'Other reviews and subsequent approvals are not checked.', 'Results depend on Graph permissions and visibility.')
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
	$csvLines = if ($preliminary.Count) { @($preliminary | ConvertTo-Csv -NoTypeInformation) } else { @('"AssignmentId","PackageId","UserId","GroupId","ReviewId","InstanceId","DecisionId","AdminAddRequestId","ReviewedAt","AppliedAt"') }
	$selectionStream = [System.IO.File]::Open($selectionPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
	try {
		$selectionWriter = [System.IO.StreamWriter]::new($selectionStream, [System.Text.UTF8Encoding]::new($false))
		try { foreach ($line in $csvLines) { $selectionWriter.WriteLine($line) } } finally { $selectionWriter.Dispose() }
	}
	finally { $selectionStream.Dispose() }
	Write-Host "$($preliminary.Count) preliminary candidate(s) saved to $selectionPath. Review them and run an approval preview before remediation."
}
if (-not $matchingDefinitions -or -not $packagesInWindow) {
	Write-Warning "Scope returned $matchingDefinitions matching review definition(s) and $packagesInWindow package(s). Check the scope and Graph visibility; an empty report does not prove there are no affected assignments."
}
if ($ApprovalCsvPath) {
	$auditFullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($AuditPath)
	$auditStream = [System.IO.File]::Open($auditFullPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
	try {
		Write-Host "Audit log: $auditFullPath"
		$auditWriter = [System.IO.StreamWriter]::new($auditStream, [System.Text.UTF8Encoding]::new($false))
		try {
			$baseAudit = @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Started'
				TenantId = [string]$context.TenantId; ReportPath = $fullPath; Mode = if ($Remediate -and -not $WhatIfPreference) { 'Remediate' } else { 'Preview' }
				SelectionCount = $approvedIds.Count; SelectionHash = $selectionHash; ScopeKey = $scopeKey }
			Write-ReconciliationAudit $auditWriter $auditStream $baseAudit
			$candidates = [System.Collections.Generic.List[object]]::new()
			$skipped = 0
			foreach ($id in @($approvedIds | Sort-Object)) {
				$matches = if ($rowsByAssignment.ContainsKey($id)) { @($rowsByAssignment[$id]) } else { @() }
				$reason = if ($matches.Count -ne 1) { 'The assignment has no unique matching review and package in the current inventory.' }
					else { Get-ReconciliationEvidenceIssue $matches[0] }
				if (-not $reason) {
					$row = $matches[0]
					$approval = $approvedEvidence[$id]
					$origin = @($row.IncidentAdminAddRequests | Where-Object { $_.Id -and $_.State -eq 'delivered' })[0]
					$approvedTime = [datetimeoffset]::MinValue
					$currentTime = [datetimeoffset]::MinValue
					$approvedAppliedTime = [datetimeoffset]::MinValue
					$currentAppliedTime = [datetimeoffset]::MinValue
					$hasApprovedAppliedTime = [datetimeoffset]::TryParse([string]$approval.AppliedAt, [ref]$approvedAppliedTime)
					$hasCurrentAppliedTime = [datetimeoffset]::TryParse([string]$row.Review.AppliedAt, [ref]$currentAppliedTime)
					if ($approval.PackageId -ne $row.PackageId -or $approval.UserId -ne $row.Review.UserId -or
						$approval.GroupId -ne $row.Review.GroupId -or $approval.ReviewId -ne $row.Review.ReviewId -or
						$approval.InstanceId -ne $row.Review.InstanceId -or $approval.DecisionId -ne $row.Review.DecisionId -or
						$approval.AdminAddRequestId -ne $origin.Id -or
						-not [datetimeoffset]::TryParse([string]$approval.ReviewedAt, [ref]$approvedTime) -or
						-not [datetimeoffset]::TryParse([string]$row.Review.ReviewedAt, [ref]$currentTime) -or $approvedTime -ne $currentTime -or
						$hasApprovedAppliedTime -ne $hasCurrentAppliedTime -or
						($hasApprovedAppliedTime -and $approvedAppliedTime -ne $currentAppliedTime) -or
						(-not $hasApprovedAppliedTime -and [string]$approval.AppliedAt -ne [string]$row.Review.AppliedAt)) {
						$reason = 'The current assignment or review evidence differs from the approved selection.'
					}
				}
				if ($reason) {
					Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Skipped'; AssignmentId = $id; Reason = $reason }
					$skipped++
				}
				else { $candidates.Add($matches[0]) }
			}
			$submitted = 0
			$previewed = 0
			if ($candidates.Count) {
				try {
					$conflicts = Get-ReconciliationReviewConflicts $candidates.ToArray()
					$sharedGroups = Get-ReconciliationSharedGroups $candidates.ToArray() $packages
				}
				catch {
					Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Error'; Reason = $_.Exception.Message }
					throw
				}
				$index = 0
				foreach ($candidate in $candidates) {
					$index++
					Show-ReconciliationProgress 'Revalidating approved assignments' $index $candidates.Count
					$id = [string]$candidate.Assignment.Id
					try { $reason = Get-ReconciliationLiveIssue $candidate $conflicts $sharedGroups }
					catch {
						Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Error'; AssignmentId = $id; Reason = $_.Exception.Message }
						throw
					}
					if ($reason) {
						Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Skipped'; AssignmentId = $id; Reason = $reason }
						$skipped++
						continue
					}
					if ($Remediate -and -not $WhatIfPreference -and -not $previewedIds.Contains($id)) {
						Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Skipped'; AssignmentId = $id; Reason = 'This assignment did not pass the approved preview.' }
						$skipped++
						continue
					}
					if (-not $Remediate) {
						Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = 'Preview'; AssignmentId = $id; PackageId = $candidate.PackageId; UserId = $candidate.Review.UserId; GroupId = $candidate.Review.GroupId }
						$previewed++
						continue
					}
					if (-not $PSCmdlet.ShouldProcess("assignment $id", 'Submit administrative removal request')) {
						$type = if ($WhatIfPreference) { 'Preview' } else { 'Declined' }
						Write-ReconciliationAudit $auditWriter $auditStream @{ Timestamp = [datetimeoffset]::UtcNow.ToString('o'); Type = $type; AssignmentId = $id }
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
			Write-Host "Approval processing complete: $submitted request(s) submitted, $previewed previewed, $skipped skipped. Audit: $auditFullPath"
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
