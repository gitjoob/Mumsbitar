<#
.SYNOPSIS
Creates an access review of the demo group's direct members.
.DESCRIPTION
Connect to Microsoft Graph first.
Requires AccessReview.ReadWrite.All and an appropriate access review license.
Group owners review access, with the configured user as fallback.
Notifications are enabled. After three days, unanswered reviews default to Deny
and denied access is automatically removed.
The recurrence interval is three months. By default, the series ends at the end
of the current month; provide a later EndDate to allow subsequent occurrences.
StartDate defaults to today. EndDate must allow at least three days for completion.
.EXAMPLE
.\Create-AccessReview.ps1 -WhatIf
.EXAMPLE
.\Create-AccessReview.ps1
.EXAMPLE
.\Create-AccessReview.ps1 -StartDate '2026-10-06' -EndDate '2027-10-31' -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
	[datetime]$StartDate = (Get-Date).Date,
	[datetime]$EndDate = (Get-Date).Date.AddDays(1 - (Get-Date).Day).AddMonths(1).AddDays(-1)
)

$ErrorActionPreference = 'Stop'
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Identity.Governance

$requiredScope = 'AccessReview.ReadWrite.All'
$context = Get-MgContext

if (-not $context) {
	throw 'Connect to Microsoft Graph first, for example with Connect-DemoOnevinn.'
}

if ($context.Scopes -notcontains $requiredScope) {
	throw "The required Microsoft Graph permission '$requiredScope' is not granted."
}

if ($StartDate.Date -lt (Get-Date).Date) {
	throw 'StartDate must be today or later.'
}

if ($EndDate.Date -lt $StartDate.Date.AddDays(3)) {
	throw 'EndDate must allow at least three days for the review to complete.'
}

$params = @{
	displayName = "Access to Demo Group"
	descriptionForAdmins = "Review direct membership of the demo group."
	descriptionForReviewers = "Confirm whether each user still needs access to the demo group."
	scope = @{
		"@odata.type" = "#microsoft.graph.accessReviewQueryScope"
		query = "/groups/655c5c03-1fe9-4aff-989c-34c334cdbcdc/members"
		queryType = "MicrosoftGraph"
	}
	reviewers = @(
		@{
			query = "/groups/655c5c03-1fe9-4aff-989c-34c334cdbcdc/owners"
			queryType = "MicrosoftGraph"
			queryRoot = $null
		}
	)
	fallbackReviewers = @(
		@{
			query = "/users/1f3fbc1f-e167-4e40-82d0-191f3f936ca9"
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
		autoApplyDecisionsEnabled = $true
		recommendationsEnabled = $true
		recommendationLookBackDuration = "P30D"
		decisionHistoriesForReviewersEnabled = $false
		recurrence = @{
			pattern = @{
				type = "absoluteMonthly"
				interval = 3
				dayOfMonth = $StartDate.Day
			}
			range = @{
				type = "endDate"
				startDate = $StartDate.ToString('yyyy-MM-dd')
				endDate = $EndDate.ToString('yyyy-MM-dd')
			}
		}
		applyActions = @(
			@{
				"@odata.type" = "#microsoft.graph.removeAccessApplyAction"
			}
		)
	}
}

if ($PSCmdlet.ShouldProcess(
	'Group 655c5c03-1fe9-4aff-989c-34c334cdbcdc',
	'Create access review with notifications and automatic removal of denied access'
)) {
	New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $params -ErrorAction Stop
}