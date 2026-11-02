<#
.SYNOPSIS
Creates a single- or multi-stage access review of a group's direct members.
.DESCRIPTION
Choose GroupId, CsvPath, or ServicePrincipalId. ServicePrincipalId discovers
groups assigned to an enterprise application, skips dynamic/empty-ownerless
groups and existing review names, and requires explicit AppOwnerIds.
All modes use the same stage policy and application-owner sign-off.
Connect to Microsoft Graph first.
Requires AccessReview.ReadWrite.All, access to read the group and its member count,
and an appropriate access review license.

ReviewerType selects ApplicationOwner (one stage), LineManager (default), or
ResourceOwner. LineManager and ResourceOwner always precede application-owner
sign-off, even for small groups. At 500 or more transitive members, an initial
self-review stage is added to those two-stage reviews. ResourceOwner uses
the group's user owners; groups without user owners use the fallback reviewer.
AppOwnerIds supplies application-owner user object IDs. The optional CSV
AppOwnerIds column overrides that parameter per row, using semicolon-separated
user IDs. Blank cells use the parameter. Every group requires application owners.

Each validated run appends its invocation syntax and per-group results to
logs/Create-AccessReview.jsonl next to the script. This directory is
gitignored. Local logging also occurs with WhatIf; no Graph writes occur.
Keep AccessReview.Logging.ps1 beside this script.

$ReviewerMessage optionally adds text (for example contact details) to the reviewer email.
Application owners are supplied through AppOwnerIds or the CSV. Optional completion
notification recipients are configured in $AdditionalNotificationRecipients;
leave it empty to send no additional notifications. Notifications are enabled.
Unanswered reviews default to Deny and denied access is automatically removed.
Groups synchronized from on-premises Active Directory cannot be changed by the
access review. For those groups, decisions are only recorded and must be
applied in Active Directory.
Each stage lasts $StageDurationInDays days. Frequency defaults to OneTime; use
Frequency to select Weekly, Monthly, Quarterly, SemiAnnual, or Annual recurrence.
For recurring reviews, the series ends at EndDate, which defaults to the end of
the current month; provide a later date to allow subsequent occurrences.
StartDate defaults to today. For recurring reviews, EndDate must allow enough
days for all stages to complete.

The access review is named with ReviewName. CsvPath creates one review per row of a CSV file with a GroupId column and an
optional ReviewName and AppOwnerIds columns. A table of skipped or failed groups and a summary
are shown at the end. If ReviewName is omitted, the name is
derived from the group's display name with the $ReviewNamePrefix prefix, for
example 'AR Demo Group'.
.EXAMPLE
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '11111111-1111-1111-1111-111111111111' -AppOwnerIds '22222222-2222-2222-2222-222222222222' -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '11111111-1111-1111-1111-111111111111' -AppOwnerIds '22222222-2222-2222-2222-222222222222' -ReviewerType ApplicationOwner
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '11111111-1111-1111-1111-111111111111' -AppOwnerIds '22222222-2222-2222-2222-222222222222' -ReviewName 'Quarterly review of Demo Group' -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -ReviewerType ResourceOwner -WhatIf
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
	[Parameter(Mandatory, ParameterSetName = 'Csv')]
	[ValidateNotNullOrEmpty()]
	[string]$CsvPath,
	[Parameter(Mandatory, ParameterSetName = 'Application')]
	[guid]$ServicePrincipalId,
	[Alias('ReviewType')]
	[ValidateSet('ApplicationOwner', 'LineManager', 'ResourceOwner')]
	[string]$ReviewerType = 'LineManager',
	[guid[]]$AppOwnerIds = @(),
	[ValidateSet('OneTime', 'Weekly', 'Monthly', 'Quarterly', 'SemiAnnual', 'Annual')]
	[string]$Frequency = 'OneTime',
	[datetime]$StartDate = (Get-Date).Date,
	[datetime]$EndDate = (Get-Date).Date.AddDays(1 - (Get-Date).Day).AddMonths(1).AddDays(-1)
)

# Prefix for the access review name when ReviewName is not provided.
$ReviewNamePrefix = 'AR'

$SelfReviewStageThreshold = 500

# Optional extra text for the reviewer email (instructions, contact details or a link).
# It is appended to the default reviewer text; the email supports up to 256 characters in total.
$ReviewerMessage = ''

# Optional recipients of the completed-review notification. Leave empty to disable it.
$AdditionalNotificationRecipients = @(
	# @{ Type = 'User'; Id = '00000000-0000-0000-0000-000000000000' }
	# @{ Type = 'Group'; Id = '00000000-0000-0000-0000-000000000000' }
)

# Reviewer used when primary reviewers (for example a line manager) are missing.
$FallbackReviewerUserId = '00000000-0000-0000-0000-000000000000'

$StageDurationInDays = 5

# Decisions passed on from one stage to the next in multi-stage reviews.
$DecisionsThatMoveToNextStage = @('Approve', 'NotReviewed')

$ErrorActionPreference = 'Stop'
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
if ($PSCmdlet.ParameterSetName -ne 'Csv' -and -not $AppOwnerIds.Count) {
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

$fallbackReviewers = @(
	@{
		query = "/users/$FallbackReviewerUserId"
		queryType = 'MicrosoftGraph'
	}
)

if ($Frequency -eq 'OneTime' -and $PSBoundParameters.ContainsKey('EndDate')) {
	Write-Warning 'EndDate is ignored for one-time reviews. The review period is calculated from StartDate and the number of stages; use -Frequency for recurring reviews.'
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
		if ($probe.PSObject.Properties['GroupId']) {
			$csvDelimiter = $candidate
			break
		}
	}
	if ($null -eq $csvDelimiter) {
		throw "The CSV file must contain a 'GroupId' column. Supported column delimiters are comma, semicolon and tab. Detected header: $header"
	}

	$targets = foreach ($row in @(Import-Csv -LiteralPath $CsvPath -Delimiter $csvDelimiter)) {
		$rowGroupId = [guid]::Empty
		if (-not [guid]::TryParse($row.GroupId.Trim(), [ref]$rowGroupId)) {
			throw "CSV GroupId '$($row.GroupId)' is not a valid GUID."
		}

		[pscustomobject]@{
			GroupId = $rowGroupId
			ReviewName = if ($row.PSObject.Properties['ReviewName']) { $row.ReviewName.Trim() } else { $null }
			AppOwnerIds = if ($row.PSObject.Properties['AppOwnerIds']) { $row.AppOwnerIds.Trim() } else { $null }
		}
	}

	$seenGroupIds = @{}
	$targets = @($targets | Where-Object { -not $seenGroupIds.ContainsKey($_.GroupId) -and ($seenGroupIds[$_.GroupId] = $true) })
	if ($targets.Count -eq 0) {
		throw 'The CSV file contains no groups.'
	}
}
elseif ($PSCmdlet.ParameterSetName -eq 'Single') {
	$targets = @([pscustomobject]@{
		GroupId = $GroupId
		ReviewName = if ($PSBoundParameters.ContainsKey('ReviewName')) { $ReviewName } else { $null }
		AppOwnerIds = $null
	})
}

else {
	$targets = @()
}

$results = [System.Collections.Generic.List[object]]::new()
. (Join-Path $PSScriptRoot 'AccessReview.Logging.ps1')
$creationLog = Start-AccessReviewLog -ScriptPath $PSCommandPath -Parameters $invocationParameters -Settings @{
	ReviewerType = $ReviewerType
	StartDate = $StartDate.ToString('o')
	EndDate = $EndDate.ToString('o')
	Frequency = $Frequency
	SelfReviewStageThreshold = $SelfReviewStageThreshold
}
$existingReviewNames = @()
if ($PSCmdlet.ParameterSetName -eq 'Application') {
	try {
		$servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $ServicePrincipalId -Property id, displayName
		$assignments = @(Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $ServicePrincipalId -All)
		$targets = @($assignments | Where-Object PrincipalType -eq 'Group' |
			Select-Object -ExpandProperty PrincipalId -Unique | ForEach-Object {
				[pscustomobject]@{ GroupId = $_; ReviewName = $null; AppOwnerIds = $null }
			})
		if ($targets.Count) {
			$existingReviewNames = @(Get-MgIdentityGovernanceAccessReviewDefinition -All -Property id, displayName |
				ForEach-Object { $_.DisplayName })
		}
		else {
			Write-Warning "Application '$($servicePrincipal.DisplayName)' has no assigned groups."
		}
	}
	catch {
		Write-AccessReviewLog -Log $creationLog -Entry @{ Event = 'RunFailed'; Reason = $_.Exception.Message }
		throw
	}
}

foreach ($target in $targets) {
	$GroupId = $target.GroupId
	$result = [ordered]@{
		Group = "$GroupId"
		GroupId = "$GroupId"
		ReviewName = $null
		ReviewId = $null
		ReviewerType = $ReviewerType
		Stages = $null
		StageNames = $null
		AppOwnerIds = @()
		Status = $null
		Reason = $null
	}

	try {
		$targetOwnerIds = @($AppOwnerIds)
		if (-not [string]::IsNullOrWhiteSpace($target.AppOwnerIds)) {
			$targetOwnerIds = @(foreach ($value in $target.AppOwnerIds.Split(';')) {
				$ownerId = [guid]::Empty
				if (-not [guid]::TryParse($value.Trim(), [ref]$ownerId) -or $ownerId -eq [guid]::Empty) {
					throw "CSV AppOwnerIds contains an invalid user object GUID: '$value'."
				}
				$ownerId
			})
		}
		if (-not $targetOwnerIds.Count) {
			throw 'Application-owner sign-off is required. Supply -AppOwnerIds or an AppOwnerIds CSV value for this group.'
		}
		$targetOwnerIds = @($targetOwnerIds | Select-Object -Unique)
		$result.AppOwnerIds = @($targetOwnerIds | ForEach-Object { "$_" })
		$applicationOwnerReviewers = @(foreach ($ownerId in $targetOwnerIds) {
			@{ query = "/users/$ownerId"; queryType = 'MicrosoftGraph' }
		})
		try {
			$group = Get-MgGroup -GroupId $GroupId -Property id, displayName, onPremisesSyncEnabled, groupTypes
			$memberCount = [int](Get-MgGroupTransitiveMemberCount -GroupId $GroupId -ConsistencyLevel eventual)
		}
		catch {
			throw "Unable to read group '$GroupId' or its member count. Check that the group exists and that the signed-in account can read it. $($_.Exception.Message)"
		}

		$isSyncedGroup = $group.OnPremisesSyncEnabled -eq $true
		$result.Group = $group.displayName
		$reviewDisplayName = if ($target.ReviewName) { $target.ReviewName.Trim() }
			elseif ($PSCmdlet.ParameterSetName -eq 'Application') { "$ReviewNamePrefix $($servicePrincipal.DisplayName) - $($group.DisplayName)" }
			else { "$ReviewNamePrefix $($group.DisplayName)".Trim() }
		$result.ReviewName = $reviewDisplayName
		if ($PSCmdlet.ParameterSetName -eq 'Application') {
			if ($group.GroupTypes -contains 'DynamicMembership') {
				$result.Status = 'Skipped'
				$result.Reason = 'Dynamic membership group'
				continue
			}
			if ($existingReviewNames -contains $reviewDisplayName) {
				$result.Status = 'Skipped'
				$result.Reason = 'Access review with the same name already exists'
				continue
			}
			if ($memberCount -eq 0) {
				$userOwnerCount = @(Get-MgGroupOwner -GroupId $GroupId -All -Property id |
					Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' }).Count
				if ($userOwnerCount -eq 0) {
					$result.Status = 'Skipped'
					$result.Reason = 'No members and no owners'
					continue
				}
			}
		}
		if ($isSyncedGroup) {
			Write-Warning "Group '$($group.displayName)' is synchronized from on-premises Active Directory. Decisions will be recorded but not applied; remove denied access in Active Directory."
		}

		$stageCount = if ($ReviewerType -eq 'ApplicationOwner') { 1 }
			elseif ($memberCount -ge $SelfReviewStageThreshold) { 3 }
			else { 2 }
		$result.Stages = $stageCount

		Write-Verbose "Group '$($group.displayName)' has $memberCount transitive members. Using $stageCount stage(s)."

		$selfReviewStage = @{ name = 'Self-review'; reviewers = @(); fallbackReviewers = @() }
		$selectedReviewerStage = @{
			name = 'Line managers'
			reviewers = @(
				@{
					query = './manager'
					queryType = 'MicrosoftGraph'
					queryRoot = 'decisions'
				}
			)
			fallbackReviewers = $fallbackReviewers
		}
		if ($ReviewerType -eq 'ResourceOwner' -and $stageCount -gt 1) {
			$userOwners = @(Get-MgGroupOwner -GroupId $GroupId -All -Property id |
				Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' })
			$ownerReviewers = @(foreach ($owner in $userOwners) {
				@{ query = "/users/$($owner.Id)"; queryType = 'MicrosoftGraph' }
			})
			if (-not $ownerReviewers.Count) {
				Write-Warning "Group '$($group.displayName)' has no user owners. Using the fallback reviewer for the ResourceOwner stage."
				$ownerReviewers = $fallbackReviewers
			}
			$selectedReviewerStage = @{
				name = 'Resource owners (group owners)'
				reviewers = $ownerReviewers
				fallbackReviewers = $fallbackReviewers
			}
		}
		$applicationOwnerStage = @{
			name = 'Application owners'
			reviewers = @($applicationOwnerReviewers)
			fallbackReviewers = $fallbackReviewers
		}

		$stages = @(switch ($stageCount) {
			1 { $applicationOwnerStage }
			2 { $selectedReviewerStage, $applicationOwnerStage }
			3 { $selfReviewStage, $selectedReviewerStage, $applicationOwnerStage }
		})

		$totalDurationInDays = $stageCount * $StageDurationInDays

		if ($StartDate.Date -lt (Get-Date).Date) {
			throw 'StartDate must be today or later.'
		}

		if ($Frequency -ne 'OneTime' -and $EndDate.Date -lt $StartDate.Date.AddDays($totalDurationInDays)) {
			throw "EndDate must allow at least $totalDurationInDays days for the review to complete."
		}

		$descriptionForReviewers = "Confirm whether each user still needs access to the group $($group.displayName)."
		if ($ReviewerMessage.Trim()) {
			$descriptionForReviewers = "$descriptionForReviewers $($ReviewerMessage.Trim())"
		}
		if ($ReviewerMessage.Trim() -and $descriptionForReviewers.Length -gt 256) {
			throw "descriptionForReviewers is $($descriptionForReviewers.Length) characters; reviewer emails support up to 256. Shorten ReviewerMessage."
		}

		$params = @{
			displayName = $reviewDisplayName
			descriptionForAdmins = if ($isSyncedGroup) {
				"Review direct membership of the group $($group.displayName) ($stageCount stage(s), $memberCount transitive members). The group is synchronized from on-premises Active Directory; apply denied access in Active Directory."
			}
			else {
				"Review direct membership of the group $($group.displayName) ($stageCount stage(s), $memberCount transitive members)."
			}
			descriptionForReviewers = $descriptionForReviewers
			scope = @{
				"@odata.type" = "#microsoft.graph.accessReviewQueryScope"
				query = "/groups/$GroupId/members"
				queryType = "MicrosoftGraph"
			}
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
			$params.reviewers = $applicationOwnerStage.reviewers
			$params.fallbackReviewers = $applicationOwnerStage.fallbackReviewers
		}
		else {
			$params.stageSettings = @(
				for ($i = 0; $i -lt $stages.Count; $i++) {
					$stageSetting = @{
						stageId = [string]($i + 1)
						durationInDays = $StageDurationInDays
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
		$result.StageNames = $stageDescription
		$result.ReviewName = $reviewDisplayName
		$decisionHandling = if ($isSyncedGroup) { 'recorded decisions only (group synchronized from on-premises)' } else { 'automatic removal of denied access' }

		if ($PSCmdlet.ShouldProcess(
			"Group $($group.displayName) ($GroupId, $memberCount transitive members)",
			"Create $stageCount-stage access review '$reviewDisplayName' ($stageDescription) with notifications and $decisionHandling"
		)) {
			$createdReview = New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $params -ErrorAction Stop
			$result.ReviewId = $createdReview.Id
			$existingReviewNames += $reviewDisplayName
			$result.Status = 'Created'
		}
		else {
			$result.Status = if ($WhatIfPreference) { 'Would be created' } else { 'Not created' }
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
		Write-AccessReviewLog -Log $creationLog -Entry @{ Event = 'GroupResult'; Result = [pscustomobject]$result }
		$resultObject = [pscustomobject]$result
		$results.Add($resultObject)
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
}

if ($PSCmdlet.ParameterSetName -ne 'Single') {
	$createdStatus = if ($WhatIfPreference) { 'Would be created' } else { 'Created' }
	$createdCount = @($results | Where-Object { $_.Status -eq $createdStatus }).Count
	$failedCount = @($results | Where-Object { $_.Status -eq 'Failed' }).Count
	$skippedCount = @($results | Where-Object Status -eq 'Skipped').Count
	$notCreated = @($results | Where-Object { $_.Status -in 'Failed', 'Skipped' })
	if ($notCreated) {
		Write-Host ($notCreated | Select-Object Group, Status, Reason | Format-Table -AutoSize -Wrap | Out-String -Width 200)
	}

	Write-Host "$($results.Count) group(s): $createdCount $($createdStatus.ToLower()), $skippedCount skipped, $failedCount failed."
}
