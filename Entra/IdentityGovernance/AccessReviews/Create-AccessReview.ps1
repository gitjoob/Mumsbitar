<#
.SYNOPSIS
Creates group-membership and direct application-user access reviews.
.DESCRIPTION
Use GroupId for one group, CsvPath for a batch, or ServicePrincipalId for an
enterprise application. Nested groups get separate reviews of their direct users.
Application mode also reviews direct user assignments; owners and non-user
assignments are listed for manual follow-up.

Stages depend on direct-user count: 1-150 owner; 151-499 managers then owner;
500+ self-review, managers, then owner. ForceTwoStage adds managers for 1-150.
ReviewerType defaults to ApplicationOwner (requires AppOwnerIds).
ResourceOwner uses the group's or application's human owners.

Denied and unanswered access is removed automatically, except for synchronized
groups, which require changes in on-premises AD. Shared-group removal can affect
multiple resources. Existing reviews are never changed; name and resource/scope
checks prevent duplicates, with uncertain overlaps flagged for manual assessment.
Start with WhatIf: it reads Graph but creates no reviews or logs.
.NOTES
Connect to Graph first with AccessReview.ReadWrite.All and resource-read access.
Keep AccessReview.Logging.ps1 and AccessReview.Reporting.ps1 beside this script.
See README.md for CSV columns, schedules, fallback reviewers and report usage.
.PARAMETER HtmlReportPath
Optional new HTML report of planned settings and creation results, not live status.
With WhatIf, only this explicitly requested file is written. Never overwrites files.
.PARAMETER DisableReminders
Disable automatic reminder emails for newly created reviews. Initial notification
emails remain enabled. Existing reviews are never modified.
.PARAMETER DiscoveryOutputPath
Optional new JSON file containing complete application group discovery. Only
available for an application WhatIf run; it contains no user membership lists.
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '<group Object ID>' -ReviewerType ResourceOwner -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -ReviewerType ResourceOwner -WhatIf -HtmlReportPath '.\reports\preview.html'
.EXAMPLE
.\Create-AccessReview.ps1 -ServicePrincipalId '<enterprise-application Object ID>' -AppOwnerIds '<owner user Object ID>' -WhatIf
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Single')]
param(
	[Parameter(Mandatory, ParameterSetName = 'Single')]
	[guid]$GroupId,
	[Parameter(ParameterSetName = 'Single')]
	[ValidateNotNullOrEmpty()]
	[string]$ReviewName,
	[AllowEmptyString()]
	[string]$ReviewNamePrefix = 'AR',
	[Parameter(Mandatory, ParameterSetName = 'Csv')]
	[ValidateNotNullOrEmpty()]
	[string]$CsvPath,
	[ValidateNotNullOrEmpty()]
	[string]$HtmlReportPath,
	[Parameter(ParameterSetName = 'Application')]
	[ValidateNotNullOrEmpty()]
	[string]$DiscoveryOutputPath,
	[Parameter(Mandatory, ParameterSetName = 'Application')]
	[guid]$ServicePrincipalId,
	[Alias('ReviewType')]
	[ValidateSet('ApplicationOwner', 'LineManager', 'ResourceOwner')]
	[string]$ReviewerType = 'ApplicationOwner',
	[switch]$ForceTwoStage,
	[switch]$DisableReminders,
	[ValidateRange(1, [int]::MaxValue)]
	[int]$SelfReviewDurationInDays = 5,
	[ValidateRange(1, [int]::MaxValue)]
	[int]$SelectedReviewerDurationInDays = 5,
	[ValidateRange(1, [int]::MaxValue)]
	[int]$ApplicationOwnerDurationInDays = 5,
	[guid[]]$AppOwnerIds = @(),
	[guid[]]$FallbackReviewerUserId = @(),
	[ValidateSet('OneTime', 'Weekly', 'Monthly', 'Quarterly', 'SemiAnnual', 'Annual')]
	[string]$Frequency = 'OneTime',
	[datetime]$StartDate = (Get-Date).Date,
	[datetime]$EndDate = (Get-Date).Date.AddDays(1 - (Get-Date).Day).AddMonths(1).AddDays(-1)
)

$SelfReviewStageThreshold = 500

# Optional extra text for the reviewer email (instructions, contact details or a link).
# It is appended to the default reviewer text; the email supports up to 256 characters in total.
$ReviewerMessage = ''

# Optional recipients of the completed-review notification. Leave empty to disable it.
$AdditionalNotificationRecipients = @(
	# @{ Type = 'User'; Id = '00000000-0000-0000-0000-000000000000' }
	# @{ Type = 'Group'; Id = '00000000-0000-0000-0000-000000000000' }
)

# Decisions passed on from one stage to the next in multi-stage reviews.
$DecisionsThatMoveToNextStage = @('Approve', 'NotReviewed')

$ErrorActionPreference = 'Stop'
if ($DiscoveryOutputPath -and -not $WhatIfPreference) {
	throw 'DiscoveryOutputPath is only supported with -WhatIf; no reviews or logs may be created by a discovery export.'
}
if ($DiscoveryOutputPath) {
	$DiscoveryOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DiscoveryOutputPath)
	if ([IO.Path]::GetExtension($DiscoveryOutputPath) -ne '.json') { throw 'DiscoveryOutputPath must end in .json.' }
	if ([IO.File]::Exists($DiscoveryOutputPath) -or [IO.Directory]::Exists($DiscoveryOutputPath)) {
		throw "Discovery output already exists: '$DiscoveryOutputPath'. Choose a new filename."
	}
	if (-not [IO.Directory]::Exists([IO.Path]::GetDirectoryName($DiscoveryOutputPath))) {
		throw "Discovery output directory does not exist: '$([IO.Path]::GetDirectoryName($DiscoveryOutputPath))'."
	}
}
function ConvertFrom-CsvUserIds {
	param([string]$Value, [string]$Column)

	foreach ($part in ($Value -split '[,;]')) {
		$id = [guid]::Empty
		if (-not [guid]::TryParse($part.Trim(), [ref]$id) -or $id -eq [guid]::Empty) {
			throw "CSV $Column contains an invalid user object GUID: '$part'."
		}
		$id
	}
}

function Get-OwnerPolicyIds {
	param([object[]]$Values)

	@($Values | ForEach-Object {
		foreach ($part in ("$_" -split '[,;]')) {
			$id = [guid]::Empty
			if ([guid]::TryParse($part.Trim(), [ref]$id)) { $id.ToString() }
			elseif ($part.Trim()) { $part.Trim().ToLowerInvariant() }
		}
	} | Sort-Object -Unique)
}

function Get-ReviewPolicy {
	param([object]$Source)

	$effectiveType = if ($Source.ReviewerType) { $Source.ReviewerType } else { $ReviewerType }
	$effectiveOwnerIds = if ([string]::IsNullOrWhiteSpace($Source.AppOwnerIds)) { @($AppOwnerIds) } else { @($Source.AppOwnerIds) }
	[ordered]@{
		ReviewerType = $effectiveType.ToLowerInvariant()
		AppOwnerIds = @(Get-OwnerPolicyIds -Values $effectiveOwnerIds)
		Role = if ($Source.Role) { $Source.Role.Trim() } else { '' }
	} | ConvertTo-Json -Compress
}

function Get-ReviewProperty {
	param([AllowNull()][object]$Object, [string]$Name)
	if ($null -eq $Object) { return }
	if ($null -ne $Object.$Name) { return $Object.$Name }
	if ($Object.AdditionalProperties -and $null -ne $Object.AdditionalProperties[$Name]) { $Object.AdditionalProperties[$Name] }
}

function Get-ReviewQueryPath {
	param([AllowEmptyString()][string]$Query)
	$Query -replace '^https://[^/]+', '' -replace '^/(?:v1\.0|beta)(?=/)', ''
}

function Get-ReviewResourceCoverage {
	param([object]$Definition)
	$scope = Get-ReviewProperty $Definition 'scope'
	$scopeQuery = [string](Get-ReviewProperty $scope 'query')
	$resourceScopes = @(Get-ReviewProperty $scope 'resourceScopes')
	$enumeration = Get-ReviewProperty $Definition 'instanceEnumerationScope'
	$query = [string](Get-ReviewProperty $enumeration 'query')
	# Enumeration supplies a base only for a relative scope, not a second resource target.
	if ($query -and -not $resourceScopes.Count -and $scopeQuery -like './*') {
		$path = Get-ReviewQueryPath $query
		if ($path -match '^/(groups|servicePrincipals)\?\$filter=(.+)$') {
			$collection = $Matches[1]
			$filter = [uri]::UnescapeDataString($Matches[2]).Trim()
			$clause = "id\s+eq\s+'[0-9a-f-]{36}'"
			if ($filter -match "^(?:$clause(?:\s+or\s+$clause)*|\(\s*$clause(?:\s+or\s+$clause)*\s*\))$") {
				foreach ($idClause in [regex]::Matches($filter, $clause, 'IgnoreCase')) {
					Get-ReviewResourceCoverage ([pscustomobject]@{
						Id = $Definition.Id
						Scope = $scope
						InstanceEnumerationScope = @{ query = "/$collection/$($idClause.Value.Split("'")[1])" }
					})
				}
				return
			}
		}
		if ($path -match '^/(groups|servicePrincipals)/([0-9a-fA-F-]{36})/?(?:\?.*)?$') {
			$kind = $Matches[1]
			$id = [guid]::Empty
			if (-not [guid]::TryParse($Matches[2], [ref]$id) -or $id -eq [guid]::Empty) {
				throw "Review '$($Definition.Id)' contains an invalid enumerated resource ID."
			}
			$resolvedScope = @{ query = "/$kind/$id" + $scopeQuery.Substring(1) }
			foreach ($property in @('queryRoot', 'inactiveDuration', '@odata.type')) {
				$value = Get-ReviewProperty $scope $property
				if ($null -ne $value) { $resolvedScope[$property] = $value }
			}
			foreach ($entry in @(Get-ReviewResourceCoverage ([pscustomobject]@{ Id = $Definition.Id; Scope = $resolvedScope }))) {
				if ($path.Contains('?')) { $entry.Coverage = 'Partial' }
				$entry.Query = "instanceEnumerationScope: $query; resolved scope: $($entry.Query)"
				$entry
			}
			return
		}
		# A broad enumeration does not identify a target. Still inspect any explicit resource scopes.
	}
	if (-not $scope) { throw "Review '$($Definition.Id)' did not return its scope." }
	$hasExplicitResources = $resourceScopes.Count -gt 0
	$allUsers = $true
	if ($resourceScopes.Count) {
		$principalScopes = @(Get-ReviewProperty $scope 'principalScopes')
		$allUsers = @($principalScopes | Where-Object {
			(Get-ReviewQueryPath ([string](Get-ReviewProperty $_ 'query'))) -match '^/users/?$' -and
				-not (Get-ReviewProperty $_ 'queryRoot') -and -not (Get-ReviewProperty $_ 'inactiveDuration')
		}).Count -gt 0
	}
	else { $resourceScopes = @($scope) }
	foreach ($resourceScope in $resourceScopes) {
		$query = [string](Get-ReviewProperty $resourceScope 'query')
		if (-not $query) { throw "Review '$($Definition.Id)' returned a resource scope without a query." }
		$path = Get-ReviewQueryPath $query
		$key = $null
		if ($path -match '^/(groups|servicePrincipals)/([0-9a-fA-F-]{36})(?=[/?]|$)(.*)$') {
			$kind = $Matches[1]
			$id = [guid]::Empty
			$suffix = $Matches[3]
			if (-not [guid]::TryParse($Matches[2], [ref]$id) -or $id -eq [guid]::Empty) {
				throw "Review '$($Definition.Id)' contains an invalid resource ID."
			}
			$key = if ($kind -eq 'groups') { "GroupMembership:$id" } else { "ApplicationDirectUsers:$id" }
			$membershipPath = if ($kind -eq 'groups') { '^/(?:transitiveMembers|members)(?:/microsoft\.graph\.user)?/?$' }
				else { '^/appRoleAssignedTo/?$' }
			$resourcePath = ($suffix -split '\?', 2)[0]
			$resourceRoot = $hasExplicitResources -and $resourcePath -in @('', '/')
			if (-not $resourceRoot -and $resourcePath -notmatch $membershipPath) { continue }
			$restricted = (Get-ReviewProperty $scope 'inactiveDuration') -or
				(Get-ReviewProperty $scope '@odata.type') -match 'InactiveUsers' -or
				(Get-ReviewProperty $resourceScope 'inactiveDuration') -or
				(Get-ReviewProperty $resourceScope 'queryRoot')
			$full = $allUsers -and -not $restricted -and ($suffix -match $membershipPath -or
				($resourceRoot -and $suffix -in @('', '/')))
			[pscustomobject]@{ Key = $key; Coverage = if ($full) { 'Full' } else { 'Partial' }; Query = $query }
		}
	}
}

function Test-ReviewCurrentCoverage {
	param([object]$Definition)
	$status = [string]$Definition.Status
	if ($status -in @('Initializing', 'NotStarted', 'Starting', 'InProgress', 'Completing', 'AutoReviewing', 'AutoReviewed', 'Applying')) { return $true }
	if ($status -notin @('Completed', 'Applied')) { throw "Review '$($Definition.Id)' returned an unknown or missing status '$status'; duplicate protection cannot determine whether it is active." }
	$settings = Get-ReviewProperty $Definition 'settings'
	$recurrence = Get-ReviewProperty $settings 'recurrence'
	if (-not (Get-ReviewProperty $recurrence 'pattern')) { return $false }
	$range = Get-ReviewProperty $recurrence 'range'
	switch (Get-ReviewProperty $range 'type') {
		'noEnd' { return $true }
		'endDate' {
			$end = Get-ReviewProperty $range 'endDate'
			if (-not $end) { throw "Review '$($Definition.Id)' has no recurrence end date." }
			return ([datetime]$end).Date -gt (Get-Date).Date
		}
		'numbered' {
			$count = Get-ReviewProperty $range 'numberOfOccurrences'
			if (-not $count -or [int]$count -lt 1) { throw "Review '$($Definition.Id)' has an invalid recurrence occurrence count." }
			$instances = @(Get-MgIdentityGovernanceAccessReviewDefinitionInstance -AccessReviewScheduleDefinitionId $Definition.Id -All -Property id,status -ErrorAction Stop)
			return @($instances | Where-Object { $_.Status -notin @('Completed', 'Applied') }).Count -gt 0 -or @($instances.Id | Sort-Object -Unique).Count -lt [int]$count
		}
		default { throw "Review '$($Definition.Id)' has an unsupported recurrence range; duplicate protection cannot determine future coverage." }
	}
}

function Add-DiscoveredGroup {
	param(
		[Parameter(Mandatory)][object]$Source,
		[Parameter(Mandatory)][string]$Id,
		[string[]]$ParentPath = @(),
		[Parameter(Mandatory)][hashtable]$State
	)

	$isCycleClosure = $ParentPath -contains $Id
	$path = @($ParentPath) + $Id
	if (-not $State.Snapshots.ContainsKey($Id)) {
		try {
			$group = Get-MgGroup -GroupId $Id -Property id, displayName, onPremisesSyncEnabled, groupTypes -ErrorAction Stop
			if (-not $group) { throw 'Graph returned no group.' }
			$members = @(Get-MgGroupMember -GroupId $Id -All -Property id -ErrorAction Stop)
			$userIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
			$childIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
			foreach ($member in $members) {
				$memberType = if ($member.AdditionalProperties) { $member.AdditionalProperties['@odata.type'] } else { $null }
				if (-not $member.Id -or -not $memberType) {
					throw 'Graph returned a member without an ID or directory-object type; membership cannot be classified safely.'
				}
				switch ($memberType) {
					'#microsoft.graph.user' { $null = $userIds.Add("$($member.Id)") }
					'#microsoft.graph.group' { $null = $childIds.Add("$($member.Id)") }
				}
			}
			$State.Snapshots[$Id] = @{
				Group = $group
				DirectUserCount = $userIds.Count
				ChildIds = @($childIds | Sort-Object)
			}
		}
		catch {
			throw "Unable to discover group '$Id' and its direct members. Check group read permissions, including hidden membership where applicable. $($_.Exception.Message)"
		}
	}
	$snapshot = $State.Snapshots[$Id]
	$policy = Get-ReviewPolicy -Source $Source
	if (-not $State.Groups.ContainsKey($Id)) {
		$State.Groups[$Id] = @{
			Policy = $policy
			Sources = [System.Collections.Generic.List[object]]::new()
			Aliases = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
		}
		$State.Order.Add($Id)
	}
	$entry = $State.Groups[$Id]
	if ($entry.Policy -cne $policy) {
		$previous = $entry.Sources[0]
		throw "Conflicting inherited review settings for group '$Id' from root '$($previous.GroupIds[0])' (application '$($previous.Source.ServicePrincipalId)') and root '$($path[0])' (application '$($Source.ServicePrincipalId)'). Specify consistent ReviewerType, AppOwnerIds and Role before rerunning."
	}
	$isRoot = $ParentPath.Count -eq 0
	$baseName = "$ReviewNamePrefix $($snapshot.Group.DisplayName)".Trim()
	$name = if ($isRoot -and $Source.ReviewName) { $Source.ReviewName.Trim() }
		elseif ($isRoot -and $Source.ServicePrincipalId) { "$ReviewNamePrefix $($Source.ServicePrincipal.DisplayName) - $($snapshot.Group.DisplayName)".Trim() }
		elseif ($isRoot) { $baseName }
		else { $baseName }
	$entry.Sources.Add([pscustomobject]@{
		Source = $Source
		GroupIds = $path
		Name = $name
		Priority = if (-not $isRoot) { 3 } elseif ($Source.ReviewName) { 0 } elseif (-not $Source.ServicePrincipalId) { 1 } else { 2 }
	})
	$null = $entry.Aliases.Add($name)
	$null = $entry.Aliases.Add("$baseName [Nested]")
	if (-not $isRoot) {
		$null = $entry.Aliases.Add($baseName)
		if ($Source.ServicePrincipalId) {
			$null = $entry.Aliases.Add("$ReviewNamePrefix $($Source.ServicePrincipal.DisplayName) - $($snapshot.Group.DisplayName)".Trim())
		}
	}
	if ($isCycleClosure) {
		$groupNames = @($path | ForEach-Object { $State.Snapshots[$_].Group.DisplayName })
		$description = ($path | ForEach-Object { "$($State.Snapshots[$_].Group.DisplayName) ($_)" }) -join ' -> '
		$message = "Circular nested group membership detected: $description. The connection is retained but will not be followed again. Discovery continues; review targets cover direct users only. Investigate this circular membership."
		Write-AccessReviewLog -Log $creationLog -Entry @{
			Event = 'GroupCycleDetected'
			GroupIds = @($path)
			GroupNames = $groupNames
			ServicePrincipalId = if ($Source.ServicePrincipalId) { "$($Source.ServicePrincipalId)" } else { $null }
			Reason = $message
		}
		Write-Warning $message
		return
	}
	foreach ($childId in $snapshot.ChildIds) {
		Add-DiscoveredGroup -Source $Source -Id $childId -ParentPath $path -State $State
	}
}

if ($ReviewerType -eq 'LineManager') {
	Write-Warning 'ReviewerType LineManager is deprecated; using ApplicationOwner for the owner stage. Managers are selected based on target size or ForceTwoStage.'
	$ReviewerType = 'ApplicationOwner'
}
$invocationParameters = @{} + $PSBoundParameters
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Identity.Governance
if ($PSCmdlet.ParameterSetName -eq 'Application') {
	Import-Module Microsoft.Graph.Applications
}

$requiredScope = @('AccessReview.ReadWrite.All')
if ($PSCmdlet.ParameterSetName -eq 'Application') {
	$requiredScope += 'Application.Read.All'
}
$context = Get-MgContext

if (-not $context) {
	throw 'Connect to Microsoft Graph first.'
}

$missingScopes = @($requiredScope | Where-Object { $context.Scopes -notcontains $_ -and $context.AuthType -ne 'AppOnly' })
if ($missingScopes.Count -gt 0) {
	throw "The required Microsoft Graph permission(s) not granted: $($missingScopes -join ', ')."
}

if ($AppOwnerIds -contains [guid]::Empty) {
	throw 'AppOwnerIds must contain nonempty user object GUIDs.'
}
if ($PSBoundParameters.ContainsKey('FallbackReviewerUserId') -and
	(-not $FallbackReviewerUserId.Count -or $FallbackReviewerUserId -contains [guid]::Empty)) {
	throw 'FallbackReviewerUserId must be a nonempty list of nonempty user object GUIDs.'
}
$FallbackReviewerUserId = @($FallbackReviewerUserId | Select-Object -Unique)
if ($PSCmdlet.ParameterSetName -ne 'Csv' -and $ReviewerType -eq 'ApplicationOwner' -and -not $AppOwnerIds.Count) {
	throw 'Application-owner sign-off is required. Supply -AppOwnerIds.'
}

$additionalNotificationRecipients = foreach ($recipient in $AdditionalNotificationRecipients) {
	$recipientId = [guid]::Empty
	if (-not [guid]::TryParse($recipient.Id, [ref]$recipientId)) {
		throw "Notification recipient Id '$($recipient.Id)' is not a valid GUID."
	}

	$query = switch ($recipient.Type) {
		'User' { "/users/$recipientId" }
		'Group' { "/groups/$recipientId/members" }
		default { throw "Notification recipient Type '$($recipient.Type)' must be 'User' or 'Group'." }
	}

	@{
		notificationRecipientScope = @{
			"@odata.type" = '#microsoft.graph.accessReviewNotificationRecipientQueryScope'
			query = $query
			queryType = 'MicrosoftGraph'
		}
		notificationTemplateType = 'CompletedAdditionalRecipients'
	}
}

$useDateWindow = $Frequency -eq 'OneTime' -and $PSBoundParameters.ContainsKey('EndDate')
if ($useDateWindow -and @('SelfReviewDurationInDays', 'SelectedReviewerDurationInDays', 'ApplicationOwnerDurationInDays' |
	Where-Object { $PSBoundParameters.ContainsKey($_) }).Count) {
	Write-Warning 'EndDate takes priority over stage-duration parameters for one-time reviews. The date window will be divided evenly across the included stages.'
}

if ($PSCmdlet.ParameterSetName -eq 'Csv') {
	if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
		throw "CSV file '$CsvPath' was not found."
	}

	$header = Get-Content -LiteralPath $CsvPath -TotalCount 1
	if ([string]::IsNullOrWhiteSpace($header)) {
		throw 'The CSV file is empty or has a blank header.'
	}
	$csvDelimiter = $null
	foreach ($candidate in @(',', ';', "`t")) {
		$probe = ConvertFrom-Csv -InputObject "$header`nheader-detection" -Delimiter $candidate
		if ($probe.PSObject.Properties['GroupId'] -or $probe.PSObject.Properties['ServicePrincipalId']) {
			$csvDelimiter = $candidate
			break
		}
	}
	if ($null -eq $csvDelimiter) {
		throw "The CSV file must contain a 'GroupId' or 'ServicePrincipalId' column. Supported column delimiters are comma, semicolon and tab. Detected header: $header"
	}

	$inputTargets = @(foreach ($row in @(Import-Csv -LiteralPath $CsvPath -Delimiter $csvDelimiter)) {
		$groupValue = if ($row.PSObject.Properties['GroupId']) { $row.GroupId.Trim() } else { '' }
		$applicationValue = if ($row.PSObject.Properties['ServicePrincipalId']) { $row.ServicePrincipalId.Trim() } else { '' }
		if ([bool]$groupValue -eq [bool]$applicationValue) {
			throw 'Each CSV row must contain exactly one GroupId or ServicePrincipalId.'
		}
		$idValue = if ($groupValue) { $groupValue } else { $applicationValue }
		$idColumn = if ($groupValue) { 'GroupId' } else { 'ServicePrincipalId' }
		$rowId = [guid]::Empty
		if (-not [guid]::TryParse($idValue, [ref]$rowId) -or $rowId -eq [guid]::Empty) {
			throw "CSV $idColumn '$idValue' is not a valid nonempty GUID."
		}
		$rowReviewerType = if ($row.PSObject.Properties['ReviewerType']) { $row.ReviewerType.Trim() } else { '' }
		if ($rowReviewerType -and $rowReviewerType -notin @('ApplicationOwner', 'ResourceOwner')) {
			throw "CSV ReviewerType '$rowReviewerType' must be ApplicationOwner or ResourceOwner."
		}
		$rowReviewName = if ($row.PSObject.Properties['ReviewName']) { $row.ReviewName.Trim() } else { '' }
		if ($applicationValue -and $rowReviewName) {
			throw 'CSV ReviewName is only supported for GroupId rows; an application can produce multiple reviews.'
		}
		[pscustomobject]@{
			GroupId = if ($groupValue) { $rowId } else { $null }
			ServicePrincipalId = if ($applicationValue) { $rowId } else { $null }
			ReviewName = $rowReviewName
			AppOwnerIds = if ($row.PSObject.Properties['AppOwnerIds']) { $row.AppOwnerIds.Trim() } else { $null }
			Owner = if ($row.PSObject.Properties['Owner']) { $row.Owner.Trim() } else { $null }
			ReviewerType = $rowReviewerType
			Role = if ($row.PSObject.Properties['Role']) { $row.Role.Trim() } else { $null }
		}
	})

	$seenTargets = @{}
	$inputTargets = @($inputTargets | Where-Object {
		$key = if ($_.GroupId) { "group:$($_.GroupId)" } else { "application:$($_.ServicePrincipalId)" }
		if ($seenTargets.ContainsKey($key)) {
			$previous = $seenTargets[$key]
			$owners = (Get-OwnerPolicyIds -Values @($_.Owner)) -join ';'
			$previousOwners = (Get-OwnerPolicyIds -Values @($previous.Owner)) -join ';'
			if ($_.ReviewName -ne $previous.ReviewName -or $owners -ne $previousOwners -or
				(Get-ReviewPolicy -Source $_) -cne (Get-ReviewPolicy -Source $previous)) {
				throw "CSV contains conflicting settings for duplicate target '$key'."
			}
			return
		}
		$seenTargets[$key] = $_
		$true
	})
	if ($inputTargets.Count -eq 0) {
		throw 'The CSV file contains no targets.'
	}
	if (@($inputTargets | Where-Object ServicePrincipalId).Count) {
		Import-Module Microsoft.Graph.Applications
		if ($context.AuthType -ne 'AppOnly' -and $context.Scopes -notcontains 'Application.Read.All') {
			throw 'The required Microsoft Graph permission(s) not granted: Application.Read.All.'
		}
	}
}
elseif ($PSCmdlet.ParameterSetName -eq 'Single') {
	$inputTargets = @([pscustomobject]@{
		GroupId = $GroupId
		ServicePrincipalId = $null
		ReviewName = if ($PSBoundParameters.ContainsKey('ReviewName')) { $ReviewName } else { $null }
		AppOwnerIds = $null
		Owner = $null
		ReviewerType = $null
		Role = $null
	})
}
else {
	$inputTargets = @([pscustomobject]@{
		GroupId = $null
		ServicePrincipalId = $ServicePrincipalId
		ReviewName = $null
		AppOwnerIds = $null
		Owner = $null
		ReviewerType = $null
		Role = $null
	})
}

$results = [System.Collections.Generic.List[object]]::new()
. (Join-Path $PSScriptRoot 'AccessReview.Logging.ps1')
. (Join-Path $PSScriptRoot 'AccessReview.Reporting.ps1')
if ($HtmlReportPath) {
	$HtmlReportPath = Resolve-AccessReviewReportPath -Path $HtmlReportPath
}
$reportRun = @{
	Mode = if ($WhatIfPreference) { 'Preview (WhatIf)' } else { 'Provisioning run' }
	RunId = $null
	Completed = $false
	Failure = $null
	Applications = [System.Collections.Generic.List[object]]::new()
}
$runStarted = $false
try {
$creationLog = Start-AccessReviewLog -ScriptPath $PSCommandPath -Parameters $invocationParameters -Enabled (-not $WhatIfPreference) -Settings @{
	ReportSchemaVersion = 1
	GroupCountBasis = 'DirectUsers'
	ReviewerType = $ReviewerType
	ForceTwoStage = [bool]$ForceTwoStage
	SelfReviewDurationInDays = $SelfReviewDurationInDays
	SelectedReviewerDurationInDays = $SelectedReviewerDurationInDays
	ApplicationOwnerDurationInDays = $ApplicationOwnerDurationInDays
	FallbackReviewerUserId = if ($FallbackReviewerUserId.Count -eq 1) { "$($FallbackReviewerUserId[0])" }
		elseif ($FallbackReviewerUserId.Count -gt 1) { @($FallbackReviewerUserId | ForEach-Object { "$_" }) }
		else { $null }
	StartDate = $StartDate.ToString('o')
	EndDate = $EndDate.ToString('o')
	Frequency = $Frequency
	SelfReviewStageThreshold = $SelfReviewStageThreshold
}
$runStarted = $true
$reportRun.RunId = $creationLog.RunId
if ($creationLog.Enabled) {
	Write-Host "Creation log: $($creationLog.Path)"
}
$targets = [System.Collections.Generic.List[object]]::new()
foreach ($inputTarget in $inputTargets) {
	if (-not $inputTarget.ServicePrincipalId) {
		$targets.Add($inputTarget)
		continue
	}
	$ServicePrincipalId = $inputTarget.ServicePrincipalId
	try {
		$servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $ServicePrincipalId -Property id, displayName
		$assignments = @(Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $ServicePrincipalId -All -Property id, principalId, principalType, principalDisplayName, appRoleId, resourceId)
		$applicationOwners = @(Get-MgServicePrincipalOwner -ServicePrincipalId $ServicePrincipalId -All)
		Write-Host "Application: $($servicePrincipal.DisplayName) (service-principal Object ID: $ServicePrincipalId)"
		$assignedGroupCount = @($assignments | Where-Object PrincipalType -eq 'Group' | Select-Object -ExpandProperty PrincipalId -Unique).Count
		$assignedUserCount = @($assignments | Where-Object PrincipalType -eq 'User' | Select-Object -ExpandProperty PrincipalId -Unique).Count
		$otherAssignmentCount = @($assignments | Where-Object { $_.PrincipalType -notin @('Group', 'User') }).Count
		Write-Host "Discovered app-role assignments: $assignedGroupCount distinct group(s), $assignedUserCount distinct direct user(s), $otherAssignmentCount other assignment(s)."
		if ($assignedUserCount -eq 0) {
			Write-Warning "Graph returned no direct user app-role assignments for service principal '$ServicePrincipalId'. Verify the enterprise application's Object ID and its Users and groups assignments. Ownership and group membership are not direct user assignments; recently changed assignments can take time to appear."
		}
		$assignedGroups = @($assignments | Where-Object PrincipalType -eq 'Group' |
			Select-Object -ExpandProperty PrincipalId -Unique)
		$directAssignments = @($assignments | Where-Object PrincipalType -eq 'User')
		$applicationRecord = [pscustomobject]@{
			ServicePrincipalId = "$ServicePrincipalId"
			ServicePrincipalName = $servicePrincipal.DisplayName
			GroupIds = @($assignedGroups)
			DirectUserCount = $assignedUserCount
		}
		$reportRun.Applications.Add($applicationRecord)
		Write-AccessReviewLog -Log $creationLog -Entry @{ Event = 'ApplicationDiscovered'; Application = $applicationRecord }
		foreach ($assignedGroup in $assignedGroups) {
			$targets.Add([pscustomobject]@{
				GroupId = $assignedGroup; ServicePrincipalId = $ServicePrincipalId
				ServicePrincipal = $servicePrincipal; ApplicationOwners = $applicationOwners
				DirectAssignments = $directAssignments; ReviewName = $null
				AppOwnerIds = $inputTarget.AppOwnerIds; Owner = $inputTarget.Owner
				ReviewerType = $inputTarget.ReviewerType; Role = $inputTarget.Role
			})
		}
		if ($directAssignments.Count) {
			$targets.Add([pscustomobject]@{
				TargetType = 'ApplicationDirectUsers'
				GroupId = $null
				ServicePrincipalId = $ServicePrincipalId
				ServicePrincipal = $servicePrincipal
				ApplicationOwners = $applicationOwners
				DirectAssignments = $directAssignments
				ReviewName = "$ReviewNamePrefix $($servicePrincipal.DisplayName)".Trim()
				AppOwnerIds = $inputTarget.AppOwnerIds
				Owner = $inputTarget.Owner
				ReviewerType = $inputTarget.ReviewerType
				Role = $inputTarget.Role
			})
		}
		if (-not $assignedGroups.Count -and -not $directAssignments.Count) {
			Write-Warning "Application '$($servicePrincipal.DisplayName)' has no assigned groups or users."
		}
		$manualItems = @(
			foreach ($owner in $applicationOwners) {
				[pscustomobject]@{
					TargetType = 'ServicePrincipalOwner'
					PrincipalId = $owner.Id
					PrincipalType = $owner.AdditionalProperties['@odata.type'] -replace '^#microsoft\.graph\.', ''
					PrincipalDisplayName = $owner.AdditionalProperties['displayName']
					AssignmentId = $null
					AppRoleId = $null
					Reason = 'Service-principal ownership is not covered by the application user-access review. Review ownership manually; no ownership changes are applied.'
				}
			}
			foreach ($assignment in @($assignments | Where-Object { $_.PrincipalType -notin @('Group', 'User') })) {
				[pscustomobject]@{
					TargetType = 'ApplicationNonUserAssignment'
					PrincipalId = $assignment.PrincipalId
					PrincipalType = $assignment.PrincipalType
					PrincipalDisplayName = $assignment.PrincipalDisplayName
					AssignmentId = $assignment.Id
					AppRoleId = $assignment.AppRoleId
					Reason = 'Non-user application-role assignment is not covered by the application user-access review. Review this assignment manually; no access changes are applied.'
				}
			}
		)
		foreach ($item in $manualItems) {
			$manualResult = [pscustomobject]@{
				TargetType = $item.TargetType
				ServicePrincipalId = "$ServicePrincipalId"
				ServicePrincipalName = $servicePrincipal.DisplayName
				Target = $servicePrincipal.DisplayName
				Group = $null
				GroupId = $null
				MemberCount = $null
				ReportSchemaVersion = 1
				GroupCountBasis = $null
				IsSyncedGroup = $null
				ReviewConfiguration = $null
				IsNested = $false
				RootGroupIds = @()
				ParentGroupIds = @()
				GroupPaths = @()
				SourceServicePrincipalIds = @()
				MatchedReviewName = $null
				MatchedReviewId = $null
				MatchedReviewScope = $null
				DuplicateMatchType = $null
				SubjectCount = $null
				PrincipalIds = @()
				AssignmentIds = @()
				ReviewName = $null
				ReviewId = $null
				ReviewerType = $null
				Role = $inputTarget.Role
				Stages = $null
				StageNames = $null
				AppOwnerIds = @()
				CsvOwnerIds = @()
				PrincipalId = $item.PrincipalId
				PrincipalType = $item.PrincipalType
				PrincipalDisplayName = $item.PrincipalDisplayName
				AssignmentId = $item.AssignmentId
				AppRoleId = $item.AppRoleId
				Status = 'Manual review required'
				Reason = $item.Reason
			}
			Write-AccessReviewLog -Log $creationLog -Entry @{ Event = 'ManualReviewRequired'; Result = $manualResult }
			$results.Add($manualResult)
		}
	}
	catch {
		Write-AccessReviewLog -Log $creationLog -Entry @{ Event = 'RunFailed'; Reason = $_.Exception.Message }
		throw
	}
}

$discovery = @{
	Snapshots = @{}
	Groups = @{}
	Order = [System.Collections.Generic.List[string]]::new()
}
try {
	foreach ($seed in @($targets | Where-Object { $_.TargetType -ne 'ApplicationDirectUsers' })) {
		Add-DiscoveredGroup -Source $seed -Id "$($seed.GroupId)" -State $discovery
	}
	$expandedTargets = [System.Collections.Generic.List[object]]::new()
	foreach ($id in $discovery.Order) {
		$entry = $discovery.Groups[$id]
		$origins = @($entry.Sources | Sort-Object Priority, Name, @{ Expression = { "$($_.Source.ServicePrincipalId)" } })
		$chosen = $origins[0]
		$rootOrigins = @($origins | Where-Object { $_.GroupIds.Count -eq 1 })
		$rootOwnerSettings = @($rootOrigins | ForEach-Object {
			(Get-OwnerPolicyIds -Values @($_.Source.Owner)) -join ';'
		} | Select-Object -Unique)
		if ($rootOwnerSettings.Count -gt 1) {
			throw "Conflicting CSV Owner settings for shared group '$id'. Supply consistent group-specific owners before rerunning."
		}
		$properties = @{}
		foreach ($property in $chosen.Source.PSObject.Properties) { $properties[$property.Name] = $property.Value }
		$properties.GroupId = $id
		$properties.ReviewName = $chosen.Name
		$properties.HasExplicitReviewName = [bool]@($rootOrigins | Where-Object { $_.Source.ReviewName }).Count
		$properties.Owner = if ($chosen.GroupIds.Count -eq 1) { $chosen.Source.Owner } else { $null }
		$properties.GroupSnapshot = $discovery.Snapshots[$id]
		$properties.KnownReviewNames = @($entry.Aliases | Sort-Object)
		$properties.IsNested = @($origins | Where-Object { $_.GroupIds.Count -gt 1 }).Count -gt 0
		$properties.IsNestedOnly = $rootOrigins.Count -eq 0
		$properties.RootGroupIds = @($origins | ForEach-Object { $_.GroupIds[0] } | Sort-Object -Unique)
		$properties.ParentGroupIds = @($origins | Where-Object { $_.GroupIds.Count -gt 1 } |
			ForEach-Object { $_.GroupIds[-2] } | Sort-Object -Unique)
		$properties.SourceServicePrincipalIds = @($origins | ForEach-Object { $_.Source.ServicePrincipalId } |
			Where-Object { $_ } | Sort-Object -Unique)
		$properties.GroupPaths = @($origins | ForEach-Object {
			[pscustomobject]@{
				ServicePrincipalId = if ($_.Source.ServicePrincipalId) { "$($_.Source.ServicePrincipalId)" } else { $null }
				ServicePrincipalName = if ($_.Source.ServicePrincipalId) { $_.Source.ServicePrincipal.DisplayName } else { $null }
				GroupIds = @($_.GroupIds)
				GroupNames = @($_.GroupIds | ForEach-Object { $discovery.Snapshots[$_].Group.DisplayName })
			}
		})
		$expandedTargets.Add([pscustomobject]$properties)
	}
	$orderedTargets = [System.Collections.Generic.List[object]]::new()
	$addedGroups = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
	foreach ($seed in $targets) {
		if ($seed.TargetType -eq 'ApplicationDirectUsers') {
			$orderedTargets.Add($seed)
			continue
		}
		foreach ($groupTarget in $expandedTargets) {
			if ($groupTarget.RootGroupIds -contains "$($seed.GroupId)" -and $addedGroups.Add("$($groupTarget.GroupId)")) {
				$orderedTargets.Add($groupTarget)
			}
		}
	}
	$targets = $orderedTargets
}
catch {
	Write-AccessReviewLog -Log $creationLog -Entry @{ Event = 'RunFailed'; Reason = $_.Exception.Message }
	throw "Group discovery failed; no reviews were created. $($_.Exception.Message)"
}

$existingResourceReviews = [System.Collections.Generic.List[object]]::new()
if ($targets.Count) {
	try {
		$existingReviews = @(Get-MgIdentityGovernanceAccessReviewDefinition -All -Property id,displayName,status,scope,instanceEnumerationScope,settings -ErrorAction Stop)
		$targetKeys = @(foreach ($target in $targets) {
			if ($target.TargetType -eq 'ApplicationDirectUsers') { "ApplicationDirectUsers:$($target.ServicePrincipalId)" }
			else { "GroupMembership:$($target.GroupId)" }
		})
		foreach ($existing in $existingReviews) {
			$coverageEntries = @(Get-ReviewResourceCoverage $existing | Where-Object { $_.Key -in $targetKeys })
			if (-not $coverageEntries.Count) { continue }
			$isCurrent = Test-ReviewCurrentCoverage $existing
			foreach ($coverage in $coverageEntries) {
				$existingResourceReviews.Add([pscustomobject]@{
					Key = $coverage.Key; Coverage = $coverage.Coverage; Query = $coverage.Query
					Id = $existing.Id; DisplayName = $existing.DisplayName; IsCurrent = $isCurrent
				})
			}
		}
	}
	catch {
		Write-AccessReviewLog -Log $creationLog -Entry @{ Event = 'RunFailed'; Reason = $_.Exception.Message }
		throw "Unable to check existing access review names and resource coverage; no reviews were created. $($_.Exception.Message)"
	}
}

$relatedGroups = @{}
$relatedCounts = @{}
foreach ($groupTarget in $targets) {
	if ($groupTarget.TargetType -eq 'ApplicationDirectUsers' -or $groupTarget.HasExplicitReviewName) { continue }
	$rootIds = @($groupTarget.RootGroupIds)
	$appIds = @($groupTarget.SourceServicePrincipalIds | Where-Object { $_ })
	$key = $null
	$label = $null
	if ($appIds.Count -eq 1 -and $groupTarget.ServicePrincipalId -eq $appIds[0]) {
		$key = "app:$($appIds[0])"
		$label = $groupTarget.ServicePrincipal.DisplayName
	}
	elseif ($appIds.Count -eq 0 -and $rootIds.Count -eq 1) {
		$key = "root:$($rootIds[0])"
		$label = $discovery.Snapshots[$rootIds[0]].Group.DisplayName
	}
	if ($key -and $label) {
		$relatedGroups["$($groupTarget.GroupId)"] = @{ Key = $key; Label = $label }
		$relatedCounts[$key]++
	}
}
$groupPlans = [System.Collections.Generic.List[object]]::new()
foreach ($target in $targets) {
	$isApplicationReview = $target.TargetType -eq 'ApplicationDirectUsers'
	$isApplicationTarget = $null -ne $target.ServicePrincipalId
	if ($isApplicationTarget) {
		$ServicePrincipalId = $target.ServicePrincipalId
		$servicePrincipal = $target.ServicePrincipal
		$applicationOwners = @($target.ApplicationOwners)
		$directAssignments = @($target.DirectAssignments)
	}
	$targetReviewerType = if ($target.ReviewerType) { $target.ReviewerType } else { $ReviewerType }
	if (-not $isApplicationReview) { $GroupId = $target.GroupId }
	$result = [ordered]@{
		TargetType = if ($isApplicationReview) { 'ApplicationDirectUsers' } else { 'GroupMembership' }
		Target = "$GroupId"
		ServicePrincipalId = if ($isApplicationTarget) { "$ServicePrincipalId" } else { $null }
		ServicePrincipalName = if ($isApplicationTarget) { $servicePrincipal.DisplayName } else { $null }
		Group = "$GroupId"
		GroupId = "$GroupId"
		MemberCount = $null
		ReportSchemaVersion = 1
		GroupCountBasis = if ($isApplicationReview) { $null } else { 'DirectUsers' }
		IsSyncedGroup = $null
		ReviewConfiguration = $null
		IsNested = [bool]$target.IsNested
		RootGroupIds = @($target.RootGroupIds | Where-Object { $_ })
		ParentGroupIds = @($target.ParentGroupIds | Where-Object { $_ })
		GroupPaths = @($target.GroupPaths | Where-Object { $_ })
		SourceServicePrincipalIds = @($target.SourceServicePrincipalIds | Where-Object { $_ })
		MatchedReviewName = $null
		MatchedReviewId = $null
		MatchedReviewScope = $null
		DuplicateMatchType = $null
		SubjectCount = $null
		PrincipalIds = @()
		AssignmentIds = @()
		PrincipalId = $null
		PrincipalType = $null
		PrincipalDisplayName = $null
		AssignmentId = $null
		AppRoleId = $null
		ReviewName = $null
		ReviewId = $null
		ReviewerType = $targetReviewerType
		Role = $target.Role
		Stages = $null
		StageNames = $null
		AppOwnerIds = @()
		CsvOwnerIds = @()
		Status = $null
		Reason = $null
	}

	try {
		$targetOwnerIds = @($AppOwnerIds)
		if (-not [string]::IsNullOrWhiteSpace($target.AppOwnerIds)) {
			$targetOwnerIds = @(ConvertFrom-CsvUserIds -Value $target.AppOwnerIds -Column 'AppOwnerIds')
		}
		$targetOwnerIds = @($targetOwnerIds | Select-Object -Unique)
		$result.AppOwnerIds = @($targetOwnerIds | ForEach-Object { "$_" })
		$csvOwnerIds = @()
		if (-not [string]::IsNullOrWhiteSpace($target.Owner)) {
			$csvOwnerIds = @(ConvertFrom-CsvUserIds -Value $target.Owner -Column 'Owner' | Select-Object -Unique)
		}
		$result.CsvOwnerIds = @($csvOwnerIds | ForEach-Object { "$_" })
		$applicationOwnerReviewers = @(foreach ($ownerId in $targetOwnerIds) {
			@{ query = "/users/$ownerId"; queryType = 'MicrosoftGraph' }
		})
		$fallbackReviewers = @(if ($FallbackReviewerUserId.Count) {
			foreach ($fallbackId in $FallbackReviewerUserId) {
				@{ query = "/users/$fallbackId"; queryType = 'MicrosoftGraph' }
			}
		}
		else { $applicationOwnerReviewers })
		if ($isApplicationReview) {
			$memberCount = @($directAssignments.PrincipalId | Select-Object -Unique).Count
			$result.SubjectCount = $memberCount
			$result.PrincipalIds = @($directAssignments.PrincipalId | Select-Object -Unique)
			$result.AssignmentIds = @($directAssignments | ForEach-Object { $_.Id } | Where-Object { $_ } | Select-Object -Unique)
			$result.Group = $null
			$result.GroupId = $null
			$result.Target = $servicePrincipal.DisplayName
			$isSyncedGroup = $false
			$reviewDisplayName = $target.ReviewName
			$resourceDescription = "application $($servicePrincipal.DisplayName) (direct user assignments)"
			$reviewScope = @{
				'@odata.type' = '#microsoft.graph.principalResourceMembershipsScope'
				principalScopes = @(@{
					'@odata.type' = '#microsoft.graph.accessReviewQueryScope'
					query = '/v1.0/users'
					queryType = 'MicrosoftGraph'
				})
				resourceScopes = @(@{
					'@odata.type' = '#microsoft.graph.accessReviewQueryScope'
					query = "/v1.0/servicePrincipals/$ServicePrincipalId"
					queryType = 'MicrosoftGraph'
				})
			}
		}
		else {
			$group = $target.GroupSnapshot.Group
			$memberCount = $target.GroupSnapshot.DirectUserCount
			$result.MemberCount = $memberCount

			$isSyncedGroup = $group.OnPremisesSyncEnabled -eq $true
			$result.Group = $group.displayName
			$result.Target = $group.DisplayName
			$reviewDisplayName = if ($target.ReviewName) { $target.ReviewName.Trim() }
				elseif ($isApplicationTarget) { "$ReviewNamePrefix $($servicePrincipal.DisplayName) - $($group.DisplayName)".Trim() }
				else { "$ReviewNamePrefix $($group.DisplayName)".Trim() }
			$resourceDescription = "group $($group.DisplayName)"
			$reviewScope = @{
				'@odata.type' = '#microsoft.graph.accessReviewQueryScope'
				query = "/groups/$GroupId/members"
				queryType = 'MicrosoftGraph'
			}
		}
		$result.IsSyncedGroup = $isSyncedGroup
		$result.ReviewName = $reviewDisplayName
		if ($memberCount -eq 0) {
			$result.Status = 'Skipped'
			$result.Reason = 'No direct users; no access review needed'
			Write-Host "The $resourceDescription has no direct users; no access review needed."
			continue
		}
		if ($isApplicationTarget -or $target.IsNestedOnly) {
			if (-not $isApplicationReview -and $group.GroupTypes -contains 'DynamicMembership') {
				$result.Status = 'Skipped'
				$result.Reason = 'Dynamic membership group'
				continue
			}
		}
		$resourceKey = if ($isApplicationReview) { "ApplicationDirectUsers:$ServicePrincipalId" } else { "GroupMembership:$GroupId" }
		$matchesForResource = @($existingResourceReviews | Where-Object Key -EQ $resourceKey)
		$matchingNames = @($reviewDisplayName) + @($target.KnownReviewNames)
		$family = if (-not $isApplicationReview) { $relatedGroups["$GroupId"] }
		if ($family) {
			$sharedName = "$ReviewNamePrefix $($family.Label) - Group membership".Trim()
			$matchingNames += $sharedName
		}
		$nameMatch = $matchesForResource | Where-Object { $_.DisplayName -and $_.DisplayName -in $matchingNames } | Select-Object -First 1
		if ($nameMatch) {
			$result.MatchedReviewName = $nameMatch.DisplayName
			$result.MatchedReviewId = $nameMatch.Id
			$result.MatchedReviewScope = $nameMatch.Query
			$result.DuplicateMatchType = 'Name'
			$result.Status = 'Skipped'
			$result.Reason = 'Access review with the same name already exists for this resource'
			continue
		}
		$match = $matchesForResource | Where-Object IsCurrent | Sort-Object @{ Expression = { if ($_.Coverage -eq 'Full') { 0 } else { 1 } } }, DisplayName, Id | Select-Object -First 1
		if ($match) {
			$result.MatchedReviewName = $match.DisplayName
			$result.MatchedReviewId = $match.Id
			$result.MatchedReviewScope = $match.Query
			if ($match.Coverage -eq 'Full') {
				$result.DuplicateMatchType = 'Resource'
				$result.Status = 'Skipped'
				$result.Reason = "Existing resource coverage (active/scheduled or potentially future recurring) in '$($match.DisplayName)' ($($match.Id)); skipped regardless of reviewer/stage policy. Existing review was not changed."
			}
			else {
				$result.DuplicateMatchType = 'UncertainResourceScope'
				$result.Status = 'Manual review required'
				$result.Reason = "Existing review '$($match.DisplayName)' ($($match.Id)) explicitly targets this resource but its scope is partial or unsupported. No review created; manually assess coverage rather than assume this resource is fully reviewed."
			}
			continue
		}
		if ($isSyncedGroup) {
			Write-Warning "Group '$($group.displayName)' is synchronized from on-premises Active Directory. Decisions will be recorded but not applied; remove denied access in Active Directory."
		}

		if ($targetReviewerType -eq 'ApplicationOwner' -and -not $targetOwnerIds.Count) {
			throw 'Application-owner sign-off is required. Supply -AppOwnerIds or an AppOwnerIds CSV value.'
		}
		$selfReviewStage = @{ name = 'Self-review'; durationInDays = $SelfReviewDurationInDays; reviewers = @(); fallbackReviewers = @() }
		$ownerStage = @{
			name = 'Application owners'
			durationInDays = $ApplicationOwnerDurationInDays
			reviewers = @($applicationOwnerReviewers)
			fallbackReviewers = $fallbackReviewers
		}
		if ($targetReviewerType -eq 'ResourceOwner') {
			if ($isSyncedGroup) {
				if ($csvOwnerIds.Count) {
					$ownerReviewers = @(foreach ($ownerId in $csvOwnerIds) {
						@{ query = "/users/$ownerId"; queryType = 'MicrosoftGraph' }
					})
					$ownerStageName = 'CSV owners (synced group)'
				}
				elseif ($FallbackReviewerUserId.Count) {
					Write-Warning "The $resourceDescription is synchronized from on-premises. Using the specified fallback reviewers instead of group owners."
					$ownerReviewers = @($fallbackReviewers)
					$ownerStageName = 'Fallback reviewers (synced group)'
				}
				else {
					throw "The $resourceDescription is synchronized from on-premises; group owners are not available for ResourceOwner routing. Supply -FallbackReviewerUserId or CSV Owner."
				}
			}
			else {
				$ownerStageName = if ($isApplicationReview) { 'Resource owners (service-principal owners)' } else { 'Resource owners (group owners)' }
				$userOwners = @(if ($isApplicationReview) {
					$applicationOwners | Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' }
				}
				else {
					Get-MgGroupOwner -GroupId $GroupId -All -Property id |
						Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' }
				})
				$ownerReviewers = @(foreach ($owner in $userOwners) {
					@{ query = "/users/$($owner.Id)"; queryType = 'MicrosoftGraph' }
				})
				if (-not $isApplicationReview -and -not $ownerReviewers.Count -and $FallbackReviewerUserId.Count) {
					Write-Warning "The $resourceDescription has no user owners. Using the specified fallback reviewers instead of group owners."
					$ownerReviewers = @($fallbackReviewers)
					$ownerStageName = 'Fallback reviewers (no group user owners)'
				}
			}
			if (-not $ownerReviewers.Count) {
				throw "The $resourceDescription has no user owners for ResourceOwner routing."
			}
			if (-not $fallbackReviewers.Count) { $fallbackReviewers = @($ownerReviewers) }
			$ownerStage = @{
				name = $ownerStageName
				durationInDays = $ApplicationOwnerDurationInDays
				reviewers = $ownerReviewers
				fallbackReviewers = $fallbackReviewers
			}
		}
		$managerStage = @{
			name = 'Line managers'
			durationInDays = $SelectedReviewerDurationInDays
			reviewers = @(
				@{
					query = './manager'
					queryType = 'MicrosoftGraph'
					queryRoot = 'decisions'
				}
			)
			fallbackReviewers = $fallbackReviewers
		}

		$stages = @(
			if ($memberCount -ge $SelfReviewStageThreshold) { $selfReviewStage }
			if ($memberCount -gt 150 -or $ForceTwoStage) { $managerStage }
			$ownerStage
		)
		$stageCount = $stages.Count
		$result.Stages = $stageCount
		if ($useDateWindow) {
			$windowDurationInDays = ($EndDate.Date - $StartDate.Date).Days
			if ($windowDurationInDays -lt $stageCount) {
				throw "EndDate must allow at least $stageCount days for $stageCount stage(s), with at least one day per stage."
			}
			$daysPerStage = [int][math]::Floor($windowDurationInDays / $stageCount)
			$extraDays = $windowDurationInDays % $stageCount
			for ($i = 0; $i -lt $stageCount; $i++) {
				$stages[$i].durationInDays = $daysPerStage + $(if ($i -lt $extraDays) { 1 } else { 0 })
			}
		}
		$countDescription = if ($isApplicationReview) { "$memberCount directly assigned users" } else { "$memberCount direct users" }
		Write-Verbose "The $resourceDescription has $countDescription. Using $stageCount stage(s)."

		$totalDurationInDays = 0
		foreach ($stage in $stages) {
			$totalDurationInDays += $stage.durationInDays
		}

		if ($StartDate.Date -lt (Get-Date).Date) {
			throw 'StartDate must be today or later.'
		}

		if ($Frequency -ne 'OneTime' -and $EndDate.Date -lt $StartDate.Date.AddDays($totalDurationInDays)) {
			throw "EndDate must allow at least $totalDurationInDays days for the review to complete."
		}

		$descriptionForReviewers = "Confirm whether each user still needs access to the $resourceDescription."
		if ($ReviewerMessage.Trim()) {
			$descriptionForReviewers = "$descriptionForReviewers $($ReviewerMessage.Trim())"
		}
		if (-not [string]::IsNullOrWhiteSpace($target.Role)) {
			$descriptionForReviewers = "$descriptionForReviewers Role: $($target.Role.Trim())."
		}
		if (($ReviewerMessage.Trim() -or $target.Role) -and $descriptionForReviewers.Length -gt 256) {
			throw "descriptionForReviewers is $($descriptionForReviewers.Length) characters; reviewer emails support up to 256. Shorten ReviewerMessage or Role."
		}
		$reviewerMessageSuffix = $descriptionForReviewers.Substring(("Confirm whether each user still needs access to the $resourceDescription.").Length)

		$relationshipPaths = @()
		if (-not $isApplicationReview) {
			$parentNames = @($target.GroupPaths | Where-Object { $_.GroupIds.Count -gt 1 } |
				ForEach-Object { $_.GroupNames[-2] } | Sort-Object -Unique)
			$applicationNames = @($target.GroupPaths | Where-Object ServicePrincipalId |
				ForEach-Object { if ($_.ServicePrincipalName) { $_.ServicePrincipalName } else { $_.ServicePrincipalId } } |
				Sort-Object -Unique)
			$contextParts = @(
				if ($parentNames.Count) { "Member of '$($parentNames -join "', '")'." }
				if ($applicationNames.Count) { "Related application assignments: '$($applicationNames -join "', '")'." }
			)
			if ($contextParts.Count) {
				$relationshipText = $contextParts -join ' '
				if ("$descriptionForReviewers $relationshipText".Length -gt 256) {
					$relationshipText = @(
						if ($parentNames.Count) { "Parent groups: $(@($target.ParentGroupIds).Count)." }
						if ($applicationNames.Count) { "Related app assignments: $(@($target.SourceServicePrincipalIds).Count)." }
					) -join ' '
					if ("$descriptionForReviewers $relationshipText".Length -gt 256) {
						throw 'descriptionForReviewers cannot fit membership context within 256 characters. Shorten ReviewerMessage or Role, or the group display name.'
					}
					Write-Warning "Using a relationship summary in the reviewer message for group '$($group.DisplayName)' to stay within 256 characters. Full membership paths are retained in the administrator description, results/logs and requested HTML report."
				}
				$descriptionForReviewers = "$descriptionForReviewers $relationshipText"
				$relationshipPaths = @($target.GroupPaths | ForEach-Object {
					$pathDescription = @($_.GroupNames) -join ' -> '
					if ($_.ServicePrincipalId) {
						$pathDescription += " [root assigned to application '$($_.ServicePrincipalName)' ($($_.ServicePrincipalId))]"
					}
					$pathDescription
				})
			}
		}
		$descriptionForAdmins = if ($isApplicationReview) {
			"Review direct user assignments to the application $($servicePrincipal.DisplayName) ($stageCount stage(s), $countDescription). Denied access removes the user's application assignment, not group membership or ownership."
		}
		elseif ($isSyncedGroup) {
			"Review direct membership of the group $($group.displayName) ($stageCount stage(s), $memberCount direct users). The group is synchronized from on-premises Active Directory; apply denied access in Active Directory."
		}
		else {
			"Review direct membership of the group $($group.displayName) ($stageCount stage(s), $memberCount direct users). Removal affects every resource granted by this group."
		}
		if ($relationshipPaths.Count) {
			$descriptionForAdmins += " Recorded membership paths (parent -> child): $($relationshipPaths -join '; '). Membership paths alone do not prove effective application access."
		}

		$params = @{
			displayName = $reviewDisplayName
			descriptionForAdmins = $descriptionForAdmins
			descriptionForReviewers = $descriptionForReviewers
			scope = $reviewScope
			settings = @{
				mailNotificationsEnabled = $true
				reminderNotificationsEnabled = -not $DisableReminders
				justificationRequiredOnApproval = $true
				defaultDecisionEnabled = $true
				defaultDecision = "Deny"
				instanceDurationInDays = $totalDurationInDays
				autoApplyDecisionsEnabled = -not $isSyncedGroup
				recommendationsEnabled = $true
				recommendationLookBackDuration = "P30D"
				decisionHistoriesForReviewersEnabled = $stageCount -gt 1
			}
		}

		if (-not $isSyncedGroup) {
			$params.settings.applyActions = @(
				@{
					"@odata.type" = "#microsoft.graph.removeAccessApplyAction"
				}
			)
		}

		if ($additionalNotificationRecipients) {
			$params.additionalNotificationRecipients = @($additionalNotificationRecipients)
		}

		$recurrenceRangeEndDate = if ($Frequency -eq 'OneTime') { $StartDate.AddDays($totalDurationInDays) } else { $EndDate }
		$recurrence = @{
			range = @{
				type = 'endDate'
				startDate = $StartDate.ToString('yyyy-MM-dd')
				endDate = $recurrenceRangeEndDate.ToString('yyyy-MM-dd')
			}
		}

		if ($Frequency -ne 'OneTime') {
			$recurrence.pattern = switch ($Frequency) {
				'Weekly' { @{ type = 'weekly'; interval = 1 } }
				'Monthly' { @{ type = 'absoluteMonthly'; interval = 1; dayOfMonth = $StartDate.Day } }
				'Quarterly' { @{ type = 'absoluteMonthly'; interval = 3; dayOfMonth = $StartDate.Day } }
				'SemiAnnual' { @{ type = 'absoluteMonthly'; interval = 6; dayOfMonth = $StartDate.Day } }
				'Annual' { @{ type = 'absoluteMonthly'; interval = 12; dayOfMonth = $StartDate.Day } }
			}
		}
		$params.settings.recurrence = $recurrence

		if ($stageCount -eq 1) {
			$params.reviewers = $stages[0].reviewers
			$params.fallbackReviewers = $stages[0].fallbackReviewers
		}
		else {
			$params.stageSettings = @(
				for ($i = 0; $i -lt $stages.Count; $i++) {
					$stageSetting = @{
						stageId = [string]($i + 1)
						durationInDays = $stages[$i].durationInDays
						recommendationsEnabled = $true
						reviewers = @($stages[$i].reviewers)
					}
					if ($stages[$i].fallbackReviewers) {
						$stageSetting.fallbackReviewers = @($stages[$i].fallbackReviewers)
					}
					if ($i -gt 0) {
						$stageSetting.dependsOn = @([string]$i)
					}
					if ($i -lt $stages.Count - 1) {
						$stageSetting.decisionsThatWillMoveToNextStage = $DecisionsThatMoveToNextStage
					}
					$stageSetting
				}
			)
		}

		$stageDescription = ($stages | ForEach-Object { $_.name }) -join ' -> '
		$stageStart = $StartDate.Date
		$result.ReviewConfiguration = [pscustomobject]@{
			Frequency = $Frequency
			StartDate = $StartDate.ToString('yyyy-MM-dd')
			InstanceDurationInDays = $totalDurationInDays
			SeriesEndDate = if ($Frequency -ne 'OneTime') { $EndDate.ToString('yyyy-MM-dd') } else { $null }
			Scope = $reviewScope
			AutoApplyDecisionsEnabled = $params.settings.autoApplyDecisionsEnabled
			ReminderNotificationsEnabled = $params.settings.reminderNotificationsEnabled
			DefaultDecision = $params.settings.defaultDecision
			DescriptionForReviewers = $params.descriptionForReviewers
			DescriptionForAdmins = $params.descriptionForAdmins
			Stages = @(
				for ($i = 0; $i -lt $stages.Count; $i++) {
					$stageEnd = $stageStart.AddDays($stages[$i].durationInDays)
					[pscustomobject]@{
						Name = $stages[$i].name
						DurationInDays = $stages[$i].durationInDays
						StartDate = $stageStart.ToString('yyyy-MM-dd')
						EndDate = $stageEnd.ToString('yyyy-MM-dd')
						Reviewers = @($stages[$i].reviewers)
						FallbackReviewers = @($stages[$i].fallbackReviewers)
						DecisionsThatWillMoveToNextStage = if ($i -lt $stages.Count - 1) { @($DecisionsThatMoveToNextStage) } else { @() }
					}
					$stageStart = $stageEnd
				}
			)
		}
		$result.StageNames = $stageDescription
		$result.ReviewName = $reviewDisplayName
		$decisionHandling = if ($isSyncedGroup) { 'recorded decisions only (group synchronized from on-premises)' } else { 'automatic removal of denied access' }

		if ($family -and $relatedCounts[$family.Key] -gt 1) {
			$sharedMessage = "Confirm whether each user still needs access to the group shown. Related resource: '$($family.Label)'. Groups can be linked through nested membership.$reviewerMessageSuffix"
			if ($sharedMessage.Length -le 256) {
				$policy = ConvertTo-Json -Depth 10 -Compress -InputObject ([ordered]@{
					GroupFamily = $family.Key
					Stages = @($result.ReviewConfiguration.Stages | ForEach-Object {
						[ordered]@{
							Name = $_.Name; Duration = $_.DurationInDays
							Reviewers = @($_.Reviewers | ForEach-Object { $_.query } | Sort-Object -Unique)
							Fallback = @($_.FallbackReviewers | ForEach-Object { $_.query } | Sort-Object -Unique)
						}
					})
					Synced = $isSyncedGroup
					Message = $sharedMessage
				})
				$resultObject = [pscustomobject]$result
				$results.Add($resultObject)
				$groupPlans.Add([pscustomobject]@{
					Policy = $policy
					Order = $groupPlans.Count
					SharedName = $sharedName
					Result = $resultObject
					Body = $params
					SharedMessage = $sharedMessage
					ResourceDescription = "$resourceDescription ($GroupId, $countDescription)"
				})
				continue
			}
		}
		if ($PSCmdlet.ShouldProcess(
			"$resourceDescription ($(if ($isApplicationReview) { $ServicePrincipalId } else { $GroupId }), $countDescription)",
			"Create $stageCount-stage access review '$reviewDisplayName' ($stageDescription) with notifications and $decisionHandling"
		)) {
			$createdReview = New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $params -ErrorAction Stop
			$result.ReviewId = $createdReview.Id
			$existingResourceReviews.Add([pscustomobject]@{
				Key = $resourceKey; Coverage = 'Full'; Query = ($reviewScope | ConvertTo-Json -Depth 6 -Compress)
				Id = $createdReview.Id; DisplayName = $reviewDisplayName; IsCurrent = $true
			})
			$result.Status = 'Created'
		}
		else {
			$result.Status = if ($WhatIfPreference) { 'Would be created' } else { 'Not created' }
			if ($WhatIfPreference) {
				$existingResourceReviews.Add([pscustomobject]@{
					Key = $resourceKey; Coverage = 'Full'; Query = ($reviewScope | ConvertTo-Json -Depth 6 -Compress)
					Id = $null; DisplayName = $reviewDisplayName; IsCurrent = $true
				})
			}
		}

		$result.ReviewName = $reviewDisplayName
		if ($isSyncedGroup) {
			$result.Reason = 'Synced from on-premises AD; decisions recorded only'
		}
	}
	catch {
		$result.Status = 'Failed'
		$result.Reason = $_.Exception.Message
		if ($PSCmdlet.ParameterSetName -eq 'Single') {
			throw
		}
	}
	finally {
		if ($result.Status) {
			$resultObject = [pscustomobject]$result
			$results.Add($resultObject)
			Write-AccessReviewLog -Log $creationLog -Entry @{
				Event = if ($isApplicationReview) { 'ApplicationResult' } else { 'GroupResult' }
				Result = $resultObject
			}
		}
	}
}

foreach ($batch in @($groupPlans | Group-Object -Property Policy -CaseSensitive | Sort-Object { $_.Group[0].Order })) {
	$plans = @($batch.Group)
	$first = $plans[0]
	$body = $first.Body
	$createdId = $null
	$reason = $null
	$status = 'Failed'
	try {
		$description = $first.ResourceDescription
		if ($plans.Count -gt 1) {
			$body.displayName = $first.SharedName
			$description = "$($plans.Count) related groups: $($plans.Result.Group -join ', ')"
			$body.scope = @{
				'@odata.type' = '#microsoft.graph.principalResourceMembershipsScope'
				principalScopes = @(@{
					'@odata.type' = '#microsoft.graph.accessReviewQueryScope'
					query = '/v1.0/users'; queryType = 'MicrosoftGraph'
				})
				resourceScopes = @($plans | ForEach-Object { $_.Body.scope })
			}
			$body.descriptionForReviewers = $first.SharedMessage
			$body.descriptionForAdmins = ($plans | ForEach-Object { $_.Result.ReviewConfiguration.DescriptionForAdmins }) -join ' '
		}
		foreach ($plan in $plans) {
			$plan.Result.ReviewName = $body.displayName
			$plan.Result.ReviewConfiguration.Scope = $body.scope
			$plan.Result.ReviewConfiguration.DescriptionForReviewers = $body.descriptionForReviewers
			$plan.Result.ReviewConfiguration.DescriptionForAdmins = $body.descriptionForAdmins
		}
		$decisionHandling = if ($first.Result.IsSyncedGroup) { 'recorded decisions only (group synchronized from on-premises)' } else { 'automatic removal of denied access' }
		if ($PSCmdlet.ShouldProcess($description, "Create $($first.Result.Stages)-stage access review '$($body.displayName)' ($($first.Result.StageNames)) with notifications and $decisionHandling")) {
			$createdId = (New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $body -ErrorAction Stop).Id
			$status = 'Created'
		}
		else { $status = if ($WhatIfPreference) { 'Would be created' } else { 'Not created' } }
	}
	catch { $reason = $_.Exception.Message }
	foreach ($plan in $plans) {
		$plan.Result.ReviewId = $createdId
		$plan.Result.Status = $status
		$plan.Result.Reason = if ($reason) { $reason } elseif ($plan.Result.IsSyncedGroup) { 'Synced from on-premises AD; decisions recorded only' }
		Write-AccessReviewLog -Log $creationLog -Entry @{ Event = 'GroupResult'; Result = $plan.Result }
	}
	if ($reason -and $PSCmdlet.ParameterSetName -eq 'Single') { throw $reason }
}

foreach ($matchGroup in @($results | Where-Object DuplicateMatchType -EQ 'UncertainResourceScope' | Group-Object MatchedReviewId)) {
	$match = $matchGroup.Group[0]
	Write-Warning "$($matchGroup.Count) target(s) need manual duplicate-coverage assessment for '$($match.MatchedReviewName)' ($($match.MatchedReviewId)); no reviews were created for those targets. See MatchedReviewScope in the results/report."
}
Write-AccessReviewLog -Log $creationLog -Entry @{
	Event = 'RunCompleted'
	Total = $results.Count
	Created = @($results | Where-Object Status -eq 'Created').Count
	Previewed = @($results | Where-Object Status -eq 'Would be created').Count
	Failed = @($results | Where-Object Status -eq 'Failed').Count
	Skipped = @($results | Where-Object Status -eq 'Skipped').Count
	ManualReviewRequired = @($results | Where-Object Status -eq 'Manual review required').Count
}

$createdStatus = if ($WhatIfPreference) { 'Would be created' } else { 'Created' }
$createdResults = @($results | Where-Object { $_.Status -eq $createdStatus })
if ($PSCmdlet.ParameterSetName -ne 'Single') {
	$createdCount = $createdResults.Count
	$failedCount = @($results | Where-Object { $_.Status -eq 'Failed' }).Count
	$skippedCount = @($results | Where-Object Status -eq 'Skipped').Count
	$reviewTargetCount = @($results | Where-Object { $_.TargetType -in @('GroupMembership', 'ApplicationDirectUsers') }).Count
	$targetLabel = if (@($inputTargets | Where-Object ServicePrincipalId).Count) { 'review target(s)' } else { 'group(s)' }
	$manualCoverageCount = @($results | Where-Object DuplicateMatchType -EQ 'UncertainResourceScope').Count
	$manualLabel = if ($manualCoverageCount) { " $manualCoverageCount require manual duplicate assessment." } else { '' }
	Write-Host "$reviewTargetCount ${targetLabel}: $createdCount $($createdStatus.ToLower()), $skippedCount skipped, $failedCount failed.$manualLabel"
}
$totalMemberCount = [long](($createdResults | Measure-Object -Property MemberCount -Sum).Sum)
$reviewSummary = if ($WhatIfPreference) { 'reviews that would be created' } else { 'created reviews' }
Write-Host "Total direct users in group ${reviewSummary}: $totalMemberCount"
if (@($inputTargets | Where-Object ServicePrincipalId).Count) {
	foreach ($directUserResult in @($results | Where-Object TargetType -eq 'ApplicationDirectUsers')) {
		Write-Host "Direct-user review: $($directUserResult.Status); discovered users: $($directUserResult.SubjectCount); review name: $($directUserResult.ReviewName); reason: $($directUserResult.Reason)"
	}
	$directUserCount = [long](($createdResults | Where-Object TargetType -eq 'ApplicationDirectUsers' | Measure-Object SubjectCount -Sum).Sum)
	$manualReviewCount = @($results | Where-Object { $_.TargetType -in @('ServicePrincipalOwner', 'ApplicationNonUserAssignment') }).Count
	Write-Host "Directly assigned users in ${reviewSummary}: $directUserCount"
	Write-Host "Manual review required: $manualReviewCount ownership/assignment item(s). These are not covered by the created reviews."
}
$reportRun.Completed = $true
}
catch {
	$reportRun.Failure = $_.Exception.Message
	throw
}
finally {
	try {
		if ($HtmlReportPath -and $runStarted) {
			try {
				Write-AccessReviewReport -Path $HtmlReportPath -Results @($results.ToArray()) -Context $reportRun
			}
			catch {
				$message = "HTML report export failed: $($_.Exception.Message) Reviews may already have been created; check returned IDs and the JSONL log before rerunning."
				if ($reportRun.Failure) {
					# WarningAction Stop must not replace the original provisioning error.
					Write-Warning $message -WarningAction Continue
				}
				else { throw $message }
			}
		}
		if ($DiscoveryOutputPath -and $reportRun.Completed -and -not $reportRun.Failure) {
			$discoveryDocument = [ordered]@{
				SchemaVersion = 1
				ServicePrincipalId = "$ServicePrincipalId"
				Applications = @($reportRun.Applications)
				Groups = @($discovery.Order | ForEach-Object {
					$snapshot = $discovery.Snapshots[$_]
					[ordered]@{
						Id = "$_"
						DisplayName = "$($snapshot.Group.DisplayName)"
						ChildIds = @($snapshot.ChildIds)
						DirectUserCount = $snapshot.DirectUserCount
					}
				})
			}
			$json = $discoveryDocument | ConvertTo-Json -Depth 20
			$stream = [IO.File]::Open($DiscoveryOutputPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
			try {
				$writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
				try { $writer.Write($json) } finally { $writer.Dispose() }
			}
			finally { $stream.Dispose() }
			Write-Host "Discovery output: $DiscoveryOutputPath"
		}
	}
	finally {
		$results.ToArray()
	}
}
