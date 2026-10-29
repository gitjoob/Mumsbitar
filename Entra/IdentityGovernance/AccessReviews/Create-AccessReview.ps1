<#
.SYNOPSIS
Creates an access review of a group's direct members.
.DESCRIPTION
Connect to Microsoft Graph first.
Requires AccessReview.ReadWrite.All and an appropriate access review license.
Reviewers are configured in $ApplicationOwners, with $FallbackReviewerUserId as fallback.
Optional completion-notification recipients are configured in
$ReviewerMessage optionally adds text (for example contact details) to the reviewer email.
$AdditionalNotificationRecipients; leave it empty to send no additional notifications.
Notifications are enabled. After three days, unanswered reviews default to Deny
and denied access is automatically removed.
Groups synchronized from on-premises Active Directory cannot be changed by the
access review. For those groups, decisions are only recorded and must be
applied in Active Directory.
Frequency defaults to OneTime; use Frequency to select Weekly, Monthly,
Quarterly, SemiAnnual, or Annual recurrence. StartDate defaults to today.
For recurring reviews, EndDate defaults to the end of the current month; provide
a later date to allow subsequent occurrences. EndDate must allow at least three
days for the review to complete.
If ReviewName is omitted, the name is derived from the group's display name.
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '00000000-0000-0000-0000-000000000000'
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -ReviewName 'Quarterly access review' -StartDate '2026-10-06' -EndDate '2027-10-31' -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -Frequency Quarterly -StartDate '2026-10-06' -EndDate '2027-10-31' -WhatIf
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

$ReviewNamePrefix = 'AR'

# Reviewers. Type is 'User' or 'Group'. Members of a group are resolved transitively.
$ApplicationOwners = @(
	@{ Type = 'User'; Id = '00000000-0000-0000-0000-000000000000' }
	# @{ Type = 'Group'; Id = '00000000-0000-0000-0000-000000000000' }
)

# Optional extra text for the reviewer email (instructions, contact details or a link).
# It is appended to the default reviewer text; the email supports up to 256 characters in total.
$ReviewerMessage = ''

# Optional recipients of the completed-review notification. Leave empty to disable it.
$AdditionalNotificationRecipients = @(
	# @{ Type = 'User'; Id = '00000000-0000-0000-0000-000000000000' }
	# @{ Type = 'Group'; Id = '00000000-0000-0000-0000-000000000000' }
)

$FallbackReviewerUserId = '00000000-0000-0000-0000-000000000000'

$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication
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

if (-not $ApplicationOwners) {
	throw 'At least one reviewer must be configured in $ApplicationOwners.'
}

$applicationOwnerReviewers = foreach ($owner in $ApplicationOwners) {
	$ownerId = [guid]::Empty
	if (-not [guid]::TryParse($owner.Id, [ref]$ownerId)) {
		throw "Reviewer Id '$($owner.Id)' is not a valid GUID."
	}

	switch ($owner.Type) {
		'User' { @{ query = "/users/$ownerId"; queryType = 'MicrosoftGraph' } }
		'Group' { @{ query = "/groups/$ownerId/transitiveMembers"; queryType = 'MicrosoftGraph' } }
		default { throw "Reviewer Type '$($owner.Type)' must be 'User' or 'Group'." }
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

if ($StartDate.Date -lt (Get-Date).Date) {
	throw 'StartDate must be today or later.'
}

if ($Frequency -ne 'OneTime' -and $EndDate.Date -lt $StartDate.Date.AddDays(3)) {
	throw 'EndDate must allow at least three days for the review to complete.'
}

try {
	$group = Get-MgGroup -GroupId $GroupId -Property id, displayName, onPremisesSyncEnabled
}
catch {
	throw "Unable to read group '$GroupId'. Check that the group exists and that the signed-in account can read it. $($_.Exception.Message)"
}

$isSyncedGroup = $group.OnPremisesSyncEnabled -eq $true
if ($isSyncedGroup) {
	Write-Warning "Group '$($group.displayName)' is synchronized from on-premises Active Directory. Decisions will be recorded but not applied; remove denied access in Active Directory."
}

$reviewDisplayName = if ($PSBoundParameters.ContainsKey('ReviewName')) {
	$ReviewName.Trim()
}
else {
	"$ReviewNamePrefix $($group.displayName)".Trim()
}

if ([string]::IsNullOrWhiteSpace($reviewDisplayName)) {
	throw 'ReviewName must not be empty or whitespace.'
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
		"Review direct membership of the group $($group.displayName). The group is synchronized from on-premises Active Directory; apply denied access in Active Directory."
	}
	else {
		"Review direct membership of the group $($group.displayName)."
	}
	descriptionForReviewers = $descriptionForReviewers
	scope = @{
		"@odata.type" = "#microsoft.graph.accessReviewQueryScope"
		query = "/groups/$GroupId/members"
		queryType = "MicrosoftGraph"
	}
	reviewers = @($applicationOwnerReviewers)
	fallbackReviewers = @(
		@{
			query = "/users/$FallbackReviewerUserId"
			queryType = "MicrosoftGraph"
		}
	)
	settings = @{
		mailNotificationsEnabled = $true
		reminderNotificationsEnabled = $true
		justificationRequiredOnApproval = $true
		defaultDecisionEnabled = $true
		defaultDecision = "Deny"
		instanceDurationInDays = 3
		autoApplyDecisionsEnabled = -not $isSyncedGroup
		recommendationsEnabled = $true
		recommendationLookBackDuration = "P30D"
		decisionHistoriesForReviewersEnabled = $false
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

$decisionHandling = if ($isSyncedGroup) { 'recorded decisions only (group synchronized from on-premises)' } else { 'automatic removal of denied access' }

if ($PSCmdlet.ShouldProcess(
	"Group $($group.displayName) ($GroupId)",
	"Create access review '$reviewDisplayName' with notifications and $decisionHandling"
)) {
	New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $params -ErrorAction Stop
}
