<#
.SYNOPSIS
Creates group-membership and direct application-user access reviews.
.DESCRIPTION
Choose GroupId, CsvPath, or ServicePrincipalId. ServicePrincipalId discovers
assigned groups and directly assigned users of an enterprise application.
It retains group-membership reviews and creates a separate application-access
review for direct users. It also reports service-principal owners and non-user
assignments (including service principals/managed identities) as Manual review
required; these are not silently skipped or automatically revoked.
Existing review names are skipped in all modes. Application mode and nested-only
targets skip dynamic groups; groups with no direct users are skipped in all modes.
All group inputs recursively discover nested groups before any review creation.
Each group is reviewed once per run, including shared children. Nested-only
reviews use '<prefix> <group> [Nested]'; root review names remain unchanged.
Known legacy unmarked child names also prevent duplicates. Existing parent
reviews do not prevent discovering or creating missing child reviews.
All modes use the same stage policy.
Connect to Microsoft Graph first.
Requires AccessReview.ReadWrite.All, access to read the group and its member count,
and an appropriate access review license.

ReviewerType selects the owner stage: ApplicationOwner (default) or ResourceOwner.
For 1-150 members the owner is the only stage; for 151-499, managers precede
the owner; for 500+ the stages are self-review, managers and the selected owner.
Group stages count distinct direct users only, excluding nested groups and
other directory objects. Direct-user application reviews count distinct directly
assigned users. Zero-direct-user groups are skipped after discovering children.
ResourceOwner uses the group's user owners, or service-principal user owners for
direct-user application reviews at every size. For synchronized groups,
CSV Owner users review instead of group owners when supplied; otherwise explicit
FallbackReviewerUserId users review. Without either, the target fails.
Non-synchronized groups without user owners also use explicit fallback users,
including groups owned only by service principals. Without fallbacks they fail.
Direct-user application reviews without user owners still fail.
ApplicationOwner uses AppOwnerIds for the final stage at every size.
Missing managers use FallbackReviewerUserId, or the effective application owners,
or the selected resource owners when no application owners are supplied.
FallbackReviewerUserId accepts one or more user object IDs.
AppOwnerIds supplies application-owner user object IDs and is required whenever
application owners review. CSV AppOwnerIds and ReviewerType override their
run-wide parameters per row; blank cells use the parameters. Children inherit
ReviewerType, AppOwnerIds and Role, but not a parent's CSV Owner or ReviewName.
ResourceOwner resolves each child's own human owners. Synced children need their
own explicit CSV Owner row or explicit fallback users. Conflicting inherited
settings for a shared group stop the run before creation. Multiple owner IDs
are comma- or semicolon-separated. Quote cells when needed for the CSV delimiter.
ForceTwoStage adds a manager stage before the selected owner for targets with
1-150 members. Targets with 151-499 already use two stages; 500+ retain all three.
It applies in single-group, CSV, and application modes; empty targets remain skipped.
The old LineManager ReviewerType maps to ApplicationOwner with a warning.

Each validated real run writes its invocation syntax and per-group results to
a separate logs/Create-AccessReview_<UTC timestamp>_<RunId>.jsonl file next to
the script and prints the log path. Logs are retained even if no reviews are
created. This directory is gitignored. WhatIf writes no logs and makes no Graph
writes. Existing combined logs are left untouched; no logs are deleted automatically.
Keep AccessReview.Logging.ps1 beside this script.

$ReviewerMessage optionally adds text (for example contact details) to the reviewer email.
Application owners are supplied through AppOwnerIds or the CSV. Optional completion
notification recipients are configured in $AdditionalNotificationRecipients;
leave it empty to send no additional notifications. Notifications are enabled.
Unanswered reviews default to Deny and denied access is automatically removed.
Groups synchronized from on-premises Active Directory cannot be changed by the
access review. For those groups, decisions are only recorded and must be
applied in Active Directory.
For one-time reviews, an explicitly supplied EndDate sets the total duration
from StartDate to EndDate. Days are divided evenly across the included stages,
with any extra days assigned to earlier stages. Each stage needs at least one day.
EndDate takes priority over the stage-duration parameters.
Otherwise, stage durations default to five days each. Override SelfReviewDurationInDays,
SelectedReviewerDurationInDays (managers), and ApplicationOwnerDurationInDays
(the final owner stage, including resource owners) independently.
Frequency defaults to OneTime; use
Frequency to select Weekly, Monthly, Quarterly, SemiAnnual, or Annual recurrence.
For recurring reviews, the series ends at EndDate, which defaults to the end of
the current month; provide a later date to allow subsequent occurrences.
StartDate defaults to today. For recurring reviews, EndDate must allow enough
days for all stages to complete.

CsvPath accepts GroupId or ServicePrincipalId rows and optional AppOwnerIds,
Owner, ReviewerType, and Role columns; ReviewName is supported for group rows only.
Owner contains one or more user object IDs and applies to the row's synchronized
group, or directly assigned synced groups for an application row, using
ResourceOwner routing at any size. Nested children do not inherit Owner.
Role is appended to the reviewer message for each review generated by the row;
it does not filter assignments. A table of skipped or failed targets and a summary
are shown at the end. All modes also show the total direct-user count across
created group reviews (or would-be-created reviews with WhatIf), counting the
same user separately if directly present in multiple groups. Group reviews cover
direct membership; denied child memberships affect all resources using that group.
Application mode reports direct-user totals and manual-review item counts
separately. Ownership is not application assignment; reporting an owner does not
create an ownership review. Application registration ownership, consent grants,
and other access paths are outside this script's scope.
If ReviewName is omitted, the name is
derived from the group's display name with the ReviewNamePrefix parameter (default AR), for
example 'AR Demo Group'.
Results and JSONL retain root/parent group IDs and all discovery paths.
Hidden-membership discovery requires Member.Read.Hidden and appropriate roles.
Discovery failures stop the run before creation; no existing reviews are changed.
.PARAMETER HtmlReportPath
Optional new .html/.htm file for a private, offline configuration/outcome report.
Supports red/white and system-aware dark themes. With WhatIf, explicitly supplying
this parameter permits writing the HTML report only, never reviews or JSONL logs.
Existing output files are rejected. The report is not live certification state.
.EXAMPLE
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -WhatIf -HtmlReportPath '.\reports\preview.html'
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '11111111-1111-1111-1111-111111111111' -AppOwnerIds '22222222-2222-2222-2222-222222222222' -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '11111111-1111-1111-1111-111111111111' -AppOwnerIds '22222222-2222-2222-2222-222222222222' -ReviewerType ApplicationOwner
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '11111111-1111-1111-1111-111111111111' -AppOwnerIds '22222222-2222-2222-2222-222222222222' -ReviewName 'Quarterly review of Demo Group' -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -ReviewerType ResourceOwner -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '11111111-1111-1111-1111-111111111111' -ReviewerType ResourceOwner -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -AppOwnerIds '22222222-2222-2222-2222-222222222222' -Frequency Quarterly -EndDate '2027-10-31' -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -ServicePrincipalId '33333333-3333-3333-3333-333333333333' -AppOwnerIds '22222222-2222-2222-2222-222222222222' -ReviewerType ResourceOwner -WhatIf
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
	[Parameter(Mandatory, ParameterSetName = 'Application')]
	[guid]$ServicePrincipalId,
	[Alias('ReviewType')]
	[ValidateSet('ApplicationOwner', 'LineManager', 'ResourceOwner')]
	[string]$ReviewerType = 'ApplicationOwner',
	[switch]$ForceTwoStage,
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

function Add-DiscoveredGroup {
	param(
		[Parameter(Mandatory)][object]$Source,
		[Parameter(Mandatory)][string]$Id,
		[string[]]$ParentPath = @(),
		[Parameter(Mandatory)][hashtable]$State
	)

	if ($ParentPath -contains $Id) {
		throw "Cycle detected in nested group membership: $((@($ParentPath) + $Id) -join ' -> ')."
	}
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
		else { "$baseName [Nested]" }
	$entry.Sources.Add([pscustomobject]@{
		Source = $Source
		GroupIds = $path
		Name = $name
		Priority = if (-not $isRoot) { 3 } elseif ($Source.ReviewName) { 0 } elseif (-not $Source.ServicePrincipalId) { 1 } else { 2 }
	})
	$null = $entry.Aliases.Add($name)
	if (-not $isRoot) {
		$null = $entry.Aliases.Add($baseName)
		if ($Source.ServicePrincipalId) {
			$null = $entry.Aliases.Add("$ReviewNamePrefix $($Source.ServicePrincipal.DisplayName) - $($snapshot.Group.DisplayName)".Trim())
		}
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
$existingReviewNames = @()
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
			Write-Warning "$($item.TargetType) '$($item.PrincipalId)' on '$($servicePrincipal.DisplayName)': $($item.Reason)"
			Write-AccessReviewLog -Log $creationLog -Entry @{ Event = 'ManualReviewRequired'; Result = $manualResult }
			$results.Add($manualResult)
			$manualResult
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

if ($targets.Count) {
	try {
		$existingReviewNames = @(Get-MgIdentityGovernanceAccessReviewDefinition -All -Property id, displayName -ErrorAction Stop |
			ForEach-Object { $_.DisplayName })
	}
	catch {
		Write-AccessReviewLog -Log $creationLog -Entry @{ Event = 'RunFailed'; Reason = $_.Exception.Message }
		throw "Unable to check existing access review names; no reviews were created. $($_.Exception.Message)"
	}
}

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
		$matchingNames = @(@($reviewDisplayName) + @($target.KnownReviewNames) |
			Where-Object { $_ -and $existingReviewNames -contains $_ })
		if ($matchingNames.Count) {
			$result.MatchedReviewName = $matchingNames[0]
			$result.Status = 'Skipped'
			$result.Reason = 'Access review with the same name already exists'
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

		$params = @{
			displayName = $reviewDisplayName
			descriptionForAdmins = if ($isApplicationReview) {
				"Review direct user assignments to the application $($servicePrincipal.DisplayName) ($stageCount stage(s), $countDescription). Denied access removes the user's application assignment, not group membership or ownership."
			}
			elseif ($isSyncedGroup) {
				"Review direct membership of the group $($group.displayName) ($stageCount stage(s), $memberCount direct users). The group is synchronized from on-premises Active Directory; apply denied access in Active Directory."
			}
			else {
				"Review direct membership of the group $($group.displayName) ($stageCount stage(s), $memberCount direct users). Removal affects every resource granted by this group."
			}
			descriptionForReviewers = $descriptionForReviewers
			scope = $reviewScope
			settings = @{
				mailNotificationsEnabled = $true
				reminderNotificationsEnabled = $true
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
			DefaultDecision = $params.settings.defaultDecision
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

		if ($PSCmdlet.ShouldProcess(
			"$resourceDescription ($(if ($isApplicationReview) { $ServicePrincipalId } else { $GroupId }), $countDescription)",
			"Create $stageCount-stage access review '$reviewDisplayName' ($stageDescription) with notifications and $decisionHandling"
		)) {
			$createdReview = New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $params -ErrorAction Stop
			$result.ReviewId = $createdReview.Id
			$existingReviewNames += $reviewDisplayName
			$existingReviewNames += @($target.KnownReviewNames)
			$result.Status = 'Created'
		}
		else {
			$result.Status = if ($WhatIfPreference) { 'Would be created' } else { 'Not created' }
			if ($WhatIfPreference) { $existingReviewNames += @($reviewDisplayName) + @($target.KnownReviewNames) }
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
		$resultObject = [pscustomobject]$result
		$results.Add($resultObject)
		Write-AccessReviewLog -Log $creationLog -Entry @{
			Event = if ($isApplicationReview) { 'ApplicationResult' } else { 'GroupResult' }
			Result = $resultObject
		}
		$resultObject
	}
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
	$notCreated = @($results | Where-Object { $_.Status -in 'Failed', 'Skipped' })
	if ($notCreated) {
		Write-Host ($notCreated | Select-Object Target, Status, Reason | Format-Table -AutoSize -Wrap | Out-String -Width 200)
	}

	$reviewTargetCount = @($results | Where-Object Status -ne 'Manual review required').Count
	$targetLabel = if (@($inputTargets | Where-Object ServicePrincipalId).Count) { 'review target(s)' } else { 'group(s)' }
	Write-Host "$reviewTargetCount ${targetLabel}: $createdCount $($createdStatus.ToLower()), $skippedCount skipped, $failedCount failed."
}
$totalMemberCount = [long](($createdResults | Measure-Object -Property MemberCount -Sum).Sum)
$reviewSummary = if ($WhatIfPreference) { 'reviews that would be created' } else { 'created reviews' }
Write-Host "Total direct users in group ${reviewSummary}: $totalMemberCount"
if (@($inputTargets | Where-Object ServicePrincipalId).Count) {
	foreach ($directUserResult in @($results | Where-Object TargetType -eq 'ApplicationDirectUsers')) {
		Write-Host "Direct-user review: $($directUserResult.Status); discovered users: $($directUserResult.SubjectCount); review name: $($directUserResult.ReviewName); reason: $($directUserResult.Reason)"
	}
	$directUserCount = [long](($createdResults | Where-Object TargetType -eq 'ApplicationDirectUsers' | Measure-Object SubjectCount -Sum).Sum)
	$manualReviewCount = @($results | Where-Object Status -eq 'Manual review required').Count
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
}
