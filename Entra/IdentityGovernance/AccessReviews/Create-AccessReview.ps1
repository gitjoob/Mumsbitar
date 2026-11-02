<#
.SYNOPSIS
Creates an access review of a group's direct members.
.DESCRIPTION
Connect to Microsoft Graph first.
Requires AccessReview.ReadWrite.All and an appropriate access review license.
Reviewers are configured in $ApplicationOwners, with $FallbackReviewerUserId as fallback.
ReviewerType can select ResourceOwner (group user owners) or LineManager
instead. ApplicationOwner is the default to preserve existing behavior.
Invocation syntax and results are appended to logs/Create-AccessReview.jsonl
(gitignored), including with WhatIf. Keep AccessReview.Logging.ps1 beside this script.
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
CsvPath creates one review per row of a CSV file with a GroupId column and an
optional ReviewName column. A table of skipped or failed groups and a summary
are shown at the end. If ReviewName is omitted, the name is derived from the group's display name.
.EXAMPLE
.\Create-AccessReview.ps1 -CsvPath '.\groups.csv' -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '00000000-0000-0000-0000-000000000000'
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -ReviewName 'Quarterly access review' -StartDate '2026-10-06' -EndDate '2027-10-31' -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1 -GroupId '00000000-0000-0000-0000-000000000000' -Frequency Quarterly -StartDate '2026-10-06' -EndDate '2027-10-31' -WhatIf
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
	[ValidateSet('ApplicationOwner', 'ResourceOwner', 'LineManager')]
	[string]$ReviewerType = 'ApplicationOwner',
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
$invocationParameters = @{} + $PSBoundParameters
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

if ($ReviewerType -eq 'ApplicationOwner' -and -not $ApplicationOwners) {
	throw 'At least one reviewer must be configured in $ApplicationOwners.'
}

$applicationOwnerReviewers = foreach ($owner in $(if ($ReviewerType -eq 'ApplicationOwner') { $ApplicationOwners })) {
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

if ($Frequency -eq 'OneTime' -and $PSBoundParameters.ContainsKey('EndDate')) {
	Write-Warning 'EndDate is ignored for one-time reviews. The review period is calculated from StartDate and the number of stages; use -Frequency for recurring reviews.'
}

if ($PSCmdlet.ParameterSetName -eq 'Csv') {
	if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
		throw "CSV file '$CsvPath' was not found."
	}

	$targets = foreach ($row in @(Import-Csv -LiteralPath $CsvPath)) {
		if (-not $row.PSObject.Properties['GroupId']) {
			throw "The CSV file must contain a 'GroupId' column."
		}

		$rowGroupId = [guid]::Empty
		if (-not [guid]::TryParse($row.GroupId.Trim(), [ref]$rowGroupId)) {
			throw "CSV GroupId '$($row.GroupId)' is not a valid GUID."
		}

		[pscustomobject]@{
			GroupId = $rowGroupId
			ReviewName = if ($row.PSObject.Properties['ReviewName']) { $row.ReviewName.Trim() } else { $null }
		}
	}

	$seenGroupIds = @{}
	$targets = @($targets | Where-Object { -not $seenGroupIds.ContainsKey($_.GroupId) -and ($seenGroupIds[$_.GroupId] = $true) })
	if ($targets.Count -eq 0) {
		throw 'The CSV file contains no groups.'
	}
}
else {
	$targets = @([pscustomobject]@{
		GroupId = $GroupId
		ReviewName = if ($PSBoundParameters.ContainsKey('ReviewName')) { $ReviewName } else { $null }
	})
}

$results = [System.Collections.Generic.List[object]]::new()
. (Join-Path $PSScriptRoot 'AccessReview.Logging.ps1')
$creationLog = Start-AccessReviewLog -ScriptPath $PSCommandPath -Parameters $invocationParameters -Settings @{
	ReviewerType = $ReviewerType
	StartDate = $StartDate.ToString('o')
	EndDate = $EndDate.ToString('o')
	Frequency = $Frequency
}

foreach ($target in $targets) {
	$GroupId = $target.GroupId
	$result = [ordered]@{
		Group = "$GroupId"
		GroupId = "$GroupId"
		ReviewName = $null
		ReviewId = $null
		ReviewerType = $ReviewerType
		Status = $null
		Reason = $null
	}

	try {
		try {
			$group = Get-MgGroup -GroupId $GroupId -Property id, displayName, onPremisesSyncEnabled
		}
		catch {
			throw "Unable to read group '$GroupId'. Check that the group exists and that the signed-in account can read it. $($_.Exception.Message)"
		}

		$isSyncedGroup = $group.OnPremisesSyncEnabled -eq $true
		$result.Group = $group.displayName
		if ($isSyncedGroup) {
			Write-Warning "Group '$($group.displayName)' is synchronized from on-premises Active Directory. Decisions will be recorded but not applied; remove denied access in Active Directory."
		}

		$reviewDisplayName = if ($target.ReviewName) {
			$target.ReviewName.Trim()
		}
		else {
			"$ReviewNamePrefix $($group.displayName)".Trim()
		}

		if ([string]::IsNullOrWhiteSpace($reviewDisplayName)) {
			throw 'ReviewName must not be empty or whitespace.'
		}
		$result.ReviewName = $reviewDisplayName
		$reviewers = @($applicationOwnerReviewers)
		if ($ReviewerType -eq 'LineManager') {
			$reviewers = @(@{ query = './manager'; queryType = 'MicrosoftGraph'; queryRoot = 'decisions' })
		}
		elseif ($ReviewerType -eq 'ResourceOwner') {
			$userOwners = @(Get-MgGroupOwner -GroupId $GroupId -All -Property id |
				Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' })
			$reviewers = @(foreach ($owner in $userOwners) {
				@{ query = "/users/$($owner.Id)"; queryType = 'MicrosoftGraph' }
			})
			if (-not $reviewers.Count) {
				Write-Warning "Group '$($group.displayName)' has no user owners. Using the fallback reviewer."
				$reviewers = @(@{ query = "/users/$FallbackReviewerUserId"; queryType = 'MicrosoftGraph' })
			}
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
			reviewers = $reviewers
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

		$recurrenceRangeEndDate = if ($Frequency -eq 'OneTime') { $StartDate.AddDays($params.settings.instanceDurationInDays) } else { $EndDate }
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
			$createdReview = New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $params -ErrorAction Stop
			$result.ReviewId = $createdReview.Id
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
	}

	$resultObject = [pscustomobject]$result
	$results.Add($resultObject)
	$resultObject
}

Write-AccessReviewLog -Log $creationLog -Entry @{
	Event = 'RunCompleted'
	Total = $results.Count
	Created = @($results | Where-Object Status -eq 'Created').Count
	Previewed = @($results | Where-Object Status -eq 'Would be created').Count
	Failed = @($results | Where-Object Status -eq 'Failed').Count
}

if ($PSCmdlet.ParameterSetName -eq 'Csv') {
	$createdStatus = if ($WhatIfPreference) { 'Would be created' } else { 'Created' }
	$createdCount = @($results | Where-Object { $_.Status -eq $createdStatus }).Count
	$failedCount = @($results | Where-Object { $_.Status -eq 'Failed' }).Count
	$notCreated = @($results | Where-Object { $_.Status -eq 'Failed' })
	if ($notCreated) {
		''
		'Failed groups:'
		$notCreated | Select-Object Group, Status, Reason | Format-Table -AutoSize -Wrap | Out-String -Width 200 -Stream | Where-Object { $_.Trim() }
	}

	''
	"$($results.Count) group(s): $createdCount $($createdStatus.ToLower()), $failedCount failed."
}
