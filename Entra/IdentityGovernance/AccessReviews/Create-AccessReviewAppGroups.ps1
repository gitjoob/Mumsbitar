<#
.SYNOPSIS
Creates one access review per group assigned to an enterprise application, reviewed by the group owners.
.DESCRIPTION
Connect to Microsoft Graph first.
Requires AccessReview.ReadWrite.All, access to read the service principal, its
app role assignments, the assigned groups, their owners and member counts, and an
appropriate access review license.

An access review scoped to an application cannot use group owners as reviewers,
and users who get access through a group are not removed from that group when
denied. This script therefore reads every group assigned to the application
(ServicePrincipalId) and creates a separate review of each group's direct
members, so that the owners of each group review their own group.

The number of transitive members in each group decides how many stages are used:
  - Up to $SelfReviewStageThreshold: 1 stage  - group owners.
  - More than $SelfReviewStageThreshold: 2 stages - self-review, then group owners.
$FallbackReviewerUserId reviews when a group has no owner.
Each stage lasts $StageDurationInDays days.

Optional completion notification recipients are configured in
$AdditionalNotificationRecipients; leave it empty to send no additional
notifications. Notifications are enabled. Unanswered reviews default to Deny and
denied access is automatically removed. Groups synchronized from on-premises
Active Directory cannot be changed by the access review. For those groups,
decisions are only recorded and must be applied in Active Directory. Groups with
dynamic membership, and groups with no members and no owners, are skipped.

Frequency defaults to OneTime; use Frequency to select Weekly, Monthly,
Quarterly, SemiAnnual, or Annual recurrence. For recurring reviews, the series
ends at EndDate, which defaults to the end of the current month. StartDate
defaults to today.

Each review is named '$ReviewNamePrefix <application> - <group>'. Groups that
already have an access review with that name are skipped, so the script can be
run again. A result object is returned for each group. At the end, a table lists
only the skipped or failed groups with the reason, followed by the number of
reviews created (or, with -WhatIf, that would be created).
.EXAMPLE
.\Create-AccessReviewAppGroups.ps1 -ServicePrincipalId '00000000-0000-0000-0000-000000000000' -WhatIf
.EXAMPLE
.\Create-AccessReviewAppGroups.ps1 -ServicePrincipalId '00000000-0000-0000-0000-000000000000'
.EXAMPLE
.\Create-AccessReviewAppGroups.ps1 -ServicePrincipalId '00000000-0000-0000-0000-000000000000' -Frequency Quarterly -StartDate '2026-11-02' -EndDate '2027-12-31'
#>
[CmdletBinding(SupportsShouldProcess)]
param(
	[Parameter(Mandatory)]
	[guid]$ServicePrincipalId,
	[ValidateSet('OneTime', 'Weekly', 'Monthly', 'Quarterly', 'SemiAnnual', 'Annual')]
	[string]$Frequency = 'OneTime',
	[datetime]$StartDate = (Get-Date).Date,
	[datetime]$EndDate = (Get-Date).Date.AddDays(1 - (Get-Date).Day).AddMonths(1).AddDays(-1)
)

# Prefix for the access review names.
$ReviewNamePrefix = 'AR'

# Groups with more transitive members than this get a self-review stage before the group owners.
$SelfReviewStageThreshold = 150

# Optional recipients of the completed-review notification. Leave empty to disable it.
$AdditionalNotificationRecipients = @(
	# @{ Type = 'User'; Id = '00000000-0000-0000-0000-000000000000' }
	# @{ Type = 'Group'; Id = '00000000-0000-0000-0000-000000000000' }
)

# Reviewer used when a group has no owner.
$FallbackReviewerUserId = '00000000-0000-0000-0000-000000000000'

$StageDurationInDays = 5

# Decisions passed on from the self-review stage to the group owner stage.
$DecisionsThatMoveToNextStage = @('Approve', 'NotReviewed')

$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Applications
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Identity.Governance

$requiredScope = 'AccessReview.ReadWrite.All'
$context = Get-MgContext

if (-not $context) {
	throw 'Connect to Microsoft Graph first.'
}

if ($context.Scopes -notcontains $requiredScope) {
	throw "The required Microsoft Graph permission '$requiredScope' is not granted."
}

if ($StartDate.Date -lt (Get-Date).Date) {
	throw 'StartDate must be today or later.'
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

$recurrenceRangeEndDate = if ($Frequency -eq 'OneTime') { $StartDate } else { $EndDate }
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

try {
	$servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $ServicePrincipalId -Property id, displayName
	$assignments = @(Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $ServicePrincipalId -All)
}
catch {
	throw "Unable to read application '$ServicePrincipalId' or its assignments. Check that the service principal exists and that the signed-in account can read it. $($_.Exception.Message)"
}

# A group can be assigned to several app roles; review it once.
$groupIds = @($assignments | Where-Object { $_.PrincipalType -eq 'Group' } | ForEach-Object { $_.PrincipalId } | Sort-Object -Unique)

if (-not $groupIds) {
	Write-Warning "Application '$($servicePrincipal.DisplayName)' has no assigned groups. No access reviews were created."
	return
}

Write-Verbose "Application '$($servicePrincipal.DisplayName)' has $($groupIds.Count) assigned group(s)."

$existingReviewNames = @(Get-MgIdentityGovernanceAccessReviewDefinition -All -Property id, displayName | ForEach-Object { $_.DisplayName })

$results = [System.Collections.Generic.List[object]]::new()

foreach ($groupId in $groupIds) {
	$result = [ordered]@{
		Group = $groupId
		GroupId = $groupId
		Members = $null
		Owners = $null
		Stages = $null
		ReportOnly = $false
		ReviewName = $null
		Status = $null
		Reason = $null
	}
	$step = 'Reading group'

	try {
		$group = Get-MgGroup -GroupId $groupId -Property id, displayName, onPremisesSyncEnabled, groupTypes
		$memberCount = [int](Get-MgGroupTransitiveMemberCount -GroupId $groupId -ConsistencyLevel eventual)
		$userOwnerCount = @(Get-MgGroupOwner -GroupId $groupId -All -Property id |
			Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' }).Count

		$reviewDisplayName = "$ReviewNamePrefix $($servicePrincipal.DisplayName) - $($group.DisplayName)".Trim()
		$stageCount = if ($memberCount -gt $SelfReviewStageThreshold) { 2 } else { 1 }
		$isSyncedGroup = $group.OnPremisesSyncEnabled -eq $true
		$isDynamicGroup = $group.GroupTypes -contains 'DynamicMembership'
		$isReportOnly = $isSyncedGroup

		$result.Group = $group.DisplayName
		$result.Members = $memberCount
		$result.Owners = $userOwnerCount
		$result.Stages = $stageCount
		$result.ReportOnly = $isReportOnly
		$result.ReviewName = $reviewDisplayName

		if ($isDynamicGroup) {
			$result.Status = 'Skipped'
			$result.Reason = 'Dynamic membership group'
			continue
		}

		if ($memberCount -eq 0 -and $userOwnerCount -eq 0) {
			$result.Status = 'Skipped'
			$result.Reason = 'No members and no owners'
			continue
		}

		if ($existingReviewNames -contains $reviewDisplayName) {
			$result.Status = 'Skipped'
			$result.Reason = 'Access review with the same name already exists'
			continue
		}

		if ($userOwnerCount -eq 0) {
			Write-Warning "Group '$($group.DisplayName)' has no user owners. The fallback reviewer reviews this group."
		}

		$sourceNote = if ($isSyncedGroup) {
			'The group is synchronized from on-premises Active Directory; apply denied access in Active Directory.'
		}

		if ($sourceNote) {
			Write-Warning "Group '$($group.DisplayName)': $sourceNote Decisions will be recorded but not applied."
		}

		$step = 'Validating schedule'
		$totalDurationInDays = $stageCount * $StageDurationInDays

		if ($Frequency -ne 'OneTime' -and $EndDate.Date -lt $StartDate.Date.AddDays($totalDurationInDays)) {
			throw "EndDate must allow at least $totalDurationInDays days for the review to complete."
		}

		$groupOwnerReviewers = @(
			@{
				query = "/groups/$groupId/owners"
				queryType = 'MicrosoftGraph'
			}
		)

		$params = @{
			displayName = $reviewDisplayName
			descriptionForAdmins = "Review direct membership of the group $($group.DisplayName), assigned to the application $($servicePrincipal.DisplayName) ($stageCount stage(s), $memberCount transitive members). $sourceNote".Trim()
			descriptionForReviewers = "Confirm whether each user still needs access to the application $($servicePrincipal.DisplayName) through the group $($group.DisplayName)."
			scope = @{
				"@odata.type" = '#microsoft.graph.accessReviewQueryScope'
				query = "/groups/$groupId/members"
				queryType = 'MicrosoftGraph'
			}
			settings = @{
				mailNotificationsEnabled = $true
				reminderNotificationsEnabled = $true
				justificationRequiredOnApproval = $true
				defaultDecisionEnabled = $true
				defaultDecision = 'Deny'
				instanceDurationInDays = $totalDurationInDays
				autoApplyDecisionsEnabled = -not $isReportOnly
				recommendationsEnabled = $true
				recommendationLookBackDuration = 'P30D'
				decisionHistoriesForReviewersEnabled = $stageCount -gt 1
				recurrence = $recurrence
			}
		}

		if (-not $isReportOnly) {
			$params.settings.applyActions = @(
				@{
					"@odata.type" = '#microsoft.graph.removeAccessApplyAction'
				}
			)
		}

		if ($additionalNotificationRecipients) {
			$params.additionalNotificationRecipients = @($additionalNotificationRecipients)
		}

		if ($stageCount -eq 1) {
			$params.reviewers = $groupOwnerReviewers
			$params.fallbackReviewers = $fallbackReviewers
		}
		else {
			$params.stageSettings = @(
				@{
					stageId = '1'
					durationInDays = $StageDurationInDays
					recommendationsEnabled = $true
					reviewers = @()
					decisionsThatWillMoveToNextStage = $DecisionsThatMoveToNextStage
				}
				@{
					stageId = '2'
					dependsOn = @('1')
					durationInDays = $StageDurationInDays
					recommendationsEnabled = $true
					reviewers = $groupOwnerReviewers
					fallbackReviewers = $fallbackReviewers
				}
			)
		}

		$stageDescription = if ($stageCount -eq 1) { 'Group owners' } else { 'Self-review -> Group owners' }
		$decisionHandling = if ($isReportOnly) { 'recorded decisions only' } else { 'automatic removal of denied access' }

		$step = 'Creating access review'

		if ($PSCmdlet.ShouldProcess(
			"Group $($group.DisplayName) ($groupId, $memberCount transitive members)",
			"Create $stageCount-stage access review '$reviewDisplayName' ($stageDescription) with notifications and $decisionHandling"
		)) {
			$null = New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $params -ErrorAction Stop
			$existingReviewNames += $reviewDisplayName
			$result.Status = 'Created'
		}
		else {
			$result.Status = if ($WhatIfPreference) { 'Would be created' } else { 'Not created' }
		}

		if ($userOwnerCount -eq 0) {
			$result.Reason = 'No user owners; fallback reviewer used'
		}
		elseif ($isSyncedGroup) {
			$result.Reason = 'Synced from on-premises AD; decisions recorded only'
		}
	}
	catch {
		$result.Status = 'Failed'
		$result.Reason = "$step`: $($_.Exception.Message)"
	}
	finally {
		$resultObject = [pscustomobject]$result
		$results.Add($resultObject)
		$resultObject
	}
}

$createdStatus = if ($WhatIfPreference) { 'Would be created' } else { 'Created' }
$createdCount = @($results | Where-Object { $_.Status -eq $createdStatus }).Count
$skippedCount = @($results | Where-Object { $_.Status -eq 'Skipped' }).Count
$failedCount = @($results | Where-Object { $_.Status -eq 'Failed' }).Count

$notCreated = @($results | Where-Object { $_.Status -in 'Skipped', 'Failed' })
if ($notCreated) {
	''
	'Skipped or failed groups:'
	$notCreated | ForEach-Object {
		[pscustomobject]@{
			PSTypeName = 'AccessReviewNotCreated'
			Group = $_.Group
			Status = $_.Status
			Reason = $_.Reason
		}
	}
}

''
"$($groupIds.Count) assigned group(s): $createdCount $($createdStatus.ToLower()), $skippedCount skipped, $failedCount failed."
