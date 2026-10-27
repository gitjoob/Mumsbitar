<#
.SYNOPSIS
Creates a single- or multi-stage access review of a group's direct members.
.DESCRIPTION
Connect to Microsoft Graph first.
Requires AccessReview.ReadWrite.All, access to read the group and its member count,
and an appropriate access review license.

The number of transitive members in the group (direct members plus members of
nested groups, as counted by Get-MgGroupTransitiveMemberCount) decides how many
stages are used,
based on the thresholds in $StageThresholds:
  - Fewer than threshold 2: 1 stage  - application owners.
  - Threshold 2 or more:    2 stages - line managers, then application owners.
  - Threshold 3 or more:    3 stages - self-review, line managers, then application owners.
Groups below threshold 1 still get the 1-stage review, but a warning is shown.

Application owners are configured in $ApplicationOwners. Optional completion
notification recipients are configured in $AdditionalNotificationRecipients;
leave it empty to send no additional notifications. Notifications are enabled.
Unanswered reviews default to Deny and denied access is automatically removed.
Each stage lasts $StageDurationInDays days. Frequency defaults to OneTime; use
Frequency to select Weekly, Monthly, Quarterly, SemiAnnual, or Annual recurrence.
For recurring reviews, the series ends at EndDate, which defaults to the end of
the current month; provide a later date to allow subsequent occurrences.
StartDate defaults to today. For recurring reviews, EndDate must allow enough
days for all stages to complete.

The access review is named with ReviewName. If ReviewName is omitted, the name is
derived from the group's display name with the $ReviewNamePrefix prefix, for
example 'AR Demo Group'.
.EXAMPLE
.\Create-AccessReviewStages.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -WhatIf
.EXAMPLE
.\Create-AccessReviewStages.ps1 -GroupId '00000000-0000-0000-0000-000000000000'
.EXAMPLE
.\Create-AccessReviewStages.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -ReviewName 'Quarterly review of Demo Group' -WhatIf
.EXAMPLE
.\Create-AccessReviewStages.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -StartDate '2026-10-06' -EndDate '2027-10-31' -WhatIf
.EXAMPLE
.\Create-AccessReviewStages.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -Frequency Quarterly -StartDate '2026-10-06' -EndDate '2027-10-31' -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
	[Parameter(Mandatory)]
	[guid]$GroupId,
	[ValidateNotNullOrEmpty()]
	[string]$ReviewName,
	[ValidateSet('OneTime', 'Weekly', 'Monthly', 'Quarterly', 'SemiAnnual', 'Annual')]
	[string]$Frequency = 'OneTime',
	[datetime]$StartDate = (Get-Date).Date,
	[datetime]$EndDate = (Get-Date).Date.AddDays(1 - (Get-Date).Day).AddMonths(1).AddDays(-1)
)

# Prefix for the access review name when ReviewName is not provided.
$ReviewNamePrefix = 'AR'

# Member count thresholds for stage 1, 2 and 3 (ascending).
$StageThresholds = @(50, 150, 500)

# Application owners who review in the final stage. Type is 'User' or 'Group'.
# Members of a group are resolved transitively.
$ApplicationOwners = @(
	@{ Type = 'User'; Id = '00000000-0000-0000-0000-000000000000' }
	# @{ Type = 'Group'; Id = '00000000-0000-0000-0000-000000000000' }
)

# Optional recipients of the completed-review notification. Leave empty to disable it.
$AdditionalNotificationRecipients = @(
	# @{ Type = 'User'; Id = '00000000-0000-0000-0000-000000000000' }
	# @{ Type = 'Group'; Id = '00000000-0000-0000-0000-000000000000' }
)

# Reviewer used when primary reviewers (for example a line manager) are missing.
$FallbackReviewerUserId = '00000000-0000-0000-0000-000000000000'

$StageDurationInDays = 3

# Decisions passed on from one stage to the next in multi-stage reviews.
$DecisionsThatMoveToNextStage = @('Approve', 'NotReviewed')

$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Groups
Import-Module Microsoft.Graph.Identity.Governance

$requiredScope = 'AccessReview.ReadWrite.All'
$context = Get-MgContext

if (-not $context) {
	throw 'Connect to Microsoft Graph first, for example with Connect-DemoOnevinn.'
}

if ($context.Scopes -notcontains $requiredScope) {
	throw "The required Microsoft Graph permission '$requiredScope' is not granted."
}

if ($StageThresholds.Count -ne 3 -or ($StageThresholds | Where-Object { $_ -isnot [int] -or $_ -lt 1 })) {
	throw 'StageThresholds must contain exactly three positive integers.'
}

if (-not ($StageThresholds[0] -lt $StageThresholds[1] -and $StageThresholds[1] -lt $StageThresholds[2])) {
	throw 'StageThresholds must be in ascending order.'
}

if (-not $ApplicationOwners) {
	throw 'At least one application owner must be configured in $ApplicationOwners.'
}

$applicationOwnerReviewers = foreach ($owner in $ApplicationOwners) {
	$ownerId = [guid]::Empty
	if (-not [guid]::TryParse($owner.Id, [ref]$ownerId)) {
		throw "Application owner Id '$($owner.Id)' is not a valid GUID."
	}

	switch ($owner.Type) {
		'User' { @{ query = "/users/$ownerId"; queryType = 'MicrosoftGraph' } }
		'Group' { @{ query = "/groups/$ownerId/transitiveMembers"; queryType = 'MicrosoftGraph' } }
		default { throw "Application owner Type '$($owner.Type)' must be 'User' or 'Group'." }
	}
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

try {
	$group = Get-MgGroup -GroupId $GroupId -Property id, displayName
	$memberCount = [int](Get-MgGroupTransitiveMemberCount -GroupId $GroupId -ConsistencyLevel eventual)
}
catch {
	throw "Unable to read group '$GroupId' or its member count. Check that the group exists and that the signed-in account can read it. $($_.Exception.Message)"
}

$stageCount = if ($memberCount -ge $StageThresholds[2]) { 3 }
	elseif ($memberCount -ge $StageThresholds[1]) { 2 }
	else { 1 }

if ($memberCount -lt $StageThresholds[0]) {
	Write-Warning "Group '$($group.displayName)' has $memberCount transitive members, below threshold 1 ($($StageThresholds[0])). Creating a 1-stage review with application owners."
}

Write-Verbose "Group '$($group.displayName)' has $memberCount transitive members. Using $stageCount stage(s)."

$reviewDisplayName = if ($PSBoundParameters.ContainsKey('ReviewName')) {
	$ReviewName.Trim()
}
else {
	"$ReviewNamePrefix $($group.displayName)".Trim()
}

$selfReviewStage = @{ name = 'Self-review'; reviewers = @(); fallbackReviewers = @() }
$managerStage = @{
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
$applicationOwnerStage = @{
	name = 'Application owners'
	reviewers = @($applicationOwnerReviewers)
	fallbackReviewers = $fallbackReviewers
}

$stages = @(switch ($stageCount) {
	1 { $applicationOwnerStage }
	2 { $managerStage, $applicationOwnerStage }
	3 { $selfReviewStage, $managerStage, $applicationOwnerStage }
})

$totalDurationInDays = $stageCount * $StageDurationInDays

if ($StartDate.Date -lt (Get-Date).Date) {
	throw 'StartDate must be today or later.'
}

if ($Frequency -ne 'OneTime' -and $EndDate.Date -lt $StartDate.Date.AddDays($totalDurationInDays)) {
	throw "EndDate must allow at least $totalDurationInDays days for the review to complete."
}

$params = @{
	displayName = $reviewDisplayName
	descriptionForAdmins = "Review direct membership of the group $($group.displayName) ($stageCount stage(s), $memberCount transitive members)."
	descriptionForReviewers = "Confirm whether each user still needs access to the group $($group.displayName)."
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
		autoApplyDecisionsEnabled = $true
		recommendationsEnabled = $true
		recommendationLookBackDuration = "P30D"
		decisionHistoriesForReviewersEnabled = $stageCount -gt 1
		applyActions = @(
			@{
				"@odata.type" = "#microsoft.graph.removeAccessApplyAction"
			}
		)
	}
}

if ($additionalNotificationRecipients) {
	$params.additionalNotificationRecipients = @($additionalNotificationRecipients)
}

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

if ($PSCmdlet.ShouldProcess(
	"Group $($group.displayName) ($GroupId, $memberCount transitive members)",
	"Create $stageCount-stage access review '$reviewDisplayName' ($stageDescription) with notifications and automatic removal of denied access"
)) {
	New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $params -ErrorAction Stop
}
