<#
.SYNOPSIS
Create a read-only HTML status report for all instances of prefix-matched reviews.
.DESCRIPTION
Requires an existing Microsoft Graph connection. Uses GET requests only.
Keep AccessReview.Reporting.ps1 beside this script. Current configuration,
runtime reviewer evidence and current resource owners are reported separately.
Includes every historical instance returned by Graph. The report and returned
object contain private user identities. No tenant data is changed.
.PARAMETER Prefixes
Literal, case-insensitive display-name prefixes. The first matching prefix wins.
.PARAMETER HtmlReportPath
New HTML output file. Existing files are never overwritten.
.PARAMETER OpenReport
Open the completed local HTML report in the default browser.
.EXAMPLE
.\Get-AccessReviewStatusReport.ps1 -Prefixes @('AR ', 'Quarterly ') -OpenReport
#>
[CmdletBinding()]
param(
	[ValidateNotNullOrEmpty()]
	[string[]]$Prefixes = @('AR '),
	[string]$HtmlReportPath = (Join-Path $PSScriptRoot "reports/AccessReviewStatus_$([datetime]::UtcNow.ToString('yyyyMMddTHHmmssfff'))_$([guid]::NewGuid().ToString('N')).html"),
	[switch]$OpenReport
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'AccessReview.Reporting.ps1')

function Add-AccessReviewStatusIssue {
	param([hashtable]$State, [string]$Source, [string]$Message)
	$State.Report.Issues.Add([pscustomobject]@{
		DefinitionId = $State.DefinitionId
		InstanceId = $State.InstanceId
		Source = $Source
		Message = $Message
	})
	Write-Warning "$Source`: $Message"
}

function Get-AccessReviewStatusCollection {
	param([hashtable]$State, [string]$Uri)
	$items = [System.Collections.Generic.List[object]]::new()
	$seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
	while ($Uri) {
		if (-not $Uri.StartsWith("$($State.Base)/", [System.StringComparison]::OrdinalIgnoreCase)) {
			throw 'Graph returned a paging URL outside the connected Graph API endpoint.'
		}
		if (-not $seen.Add($Uri)) { throw 'Graph returned a repeated paging URL.' }
		$response = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
		if ($null -eq $response -or $response.PSObject.Properties.Name -notcontains 'value') {
			throw 'Graph returned a collection response without a value property.'
		}
		foreach ($item in @($response.value)) {
			if ($null -ne $item) { $items.Add($item) }
		}
		$Uri = $response.'@odata.nextLink'
	}
	$items.ToArray()
}

function Get-AccessReviewStatusUser {
	param([hashtable]$State, [string]$Id, [string]$DisplayName, [string]$UserPrincipalName, [string]$RecordSource = 'Graph review record')
	if ($DisplayName -and $UserPrincipalName) {
		return [pscustomobject]@{ Id = $Id; DisplayName = $DisplayName; UserPrincipalName = $UserPrincipalName; ProfileSource = $RecordSource; Error = $null }
	}
	if (-not $State.Users.ContainsKey($Id)) {
		$user = $null
		$errorText = $null
		try {
			$guid = [guid]::Empty
			if (-not [guid]::TryParse($Id, [ref]$guid) -or $guid -eq [guid]::Empty) { throw "Invalid user ID '$Id'." }
			$user = Invoke-MgGraphRequest -Method GET -Uri "$($State.Base)/users/$guid`?`$select=id,displayName,userPrincipalName" -OutputType PSObject -ErrorAction Stop
			if (-not $user.id -or -not $user.userPrincipalName) { throw 'The directory did not return a user ID and sign-in address; check user-read permissions.' }
		}
		catch {
			$errorText = $_.Exception.Message
			Add-AccessReviewStatusIssue $State "User $Id" $errorText
		}
		$State.Users[$Id] = @{ User = $user; Error = $errorText }
	}
	$lookup = $State.Users[$Id]
	[pscustomobject]@{
		Id = $Id
		DisplayName = if ($DisplayName) { $DisplayName } elseif ($lookup.User.displayName) { $lookup.User.displayName } else { 'Name unavailable' }
		UserPrincipalName = if ($UserPrincipalName) { $UserPrincipalName } else { $lookup.User.userPrincipalName }
		ProfileSource = if ($lookup.Error) { "Directory lookup unavailable; supplied name source: $RecordSource" }
			elseif ($DisplayName) { "$RecordSource name; current directory sign-in address" }
			else { 'Current directory profile' }
		Error = $lookup.Error
	}
}

function Get-AccessReviewStatusResource {
	param([hashtable]$State, [string]$Kind, [string]$Id)
	$key = "$Kind/$Id"
	if (-not $State.Resources.ContainsKey($key)) {
		$resource = [pscustomobject]@{
			Kind = $Kind; Id = $Id; DisplayName = $Id
			Owners = @(); OtherOwners = @(); MetadataError = $null; OwnersError = $null
		}
		$State.Resources[$key] = $resource
		try {
			$metadata = Invoke-MgGraphRequest -Method GET -Uri "$($State.Base)/$key`?`$select=id,displayName" -OutputType PSObject -ErrorAction Stop
			if (-not $metadata.displayName) { throw 'Resource display name unavailable.' }
			$resource.DisplayName = $metadata.displayName
		}
		catch {
			$resource.MetadataError = $_.Exception.Message
			Add-AccessReviewStatusIssue $State "Resource $key" $resource.MetadataError
		}
		try {
			$owners = @(Get-AccessReviewStatusCollection $State "$($State.Base)/$key/owners?`$select=id,displayName")
			$resource.Owners = @(
				foreach ($owner in $owners) {
					if ($owner.'@odata.type' -eq '#microsoft.graph.user') {
						Get-AccessReviewStatusUser $State $owner.id $owner.displayName '' 'Current directory owner record'
					}
				}
			)
			$resource.OtherOwners = @($owners | Where-Object { $_.'@odata.type' -ne '#microsoft.graph.user' })
			if (@($resource.OtherOwners | Where-Object { -not $_.'@odata.type' }).Count) {
				throw 'Some owner object types were unavailable; their IDs are retained as non-user/unclassified owners.'
			}
		}
		catch {
			$resource.OwnersError = $_.Exception.Message
			Add-AccessReviewStatusIssue $State "Current owners of $key" $resource.OwnersError
		}
	}
	$State.Resources[$key]
}

function Get-AccessReviewStatusResources {
	param([hashtable]$State, [object[]]$Queries)
	$seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
	foreach ($query in $Queries) {
		if ([string]$query -match '(?:^/|/v1\.0/)(groups|servicePrincipals)/([0-9a-fA-F-]{36})(?:/|$|\?)') {
			$kind = $Matches[1]
			$id = [guid]::Empty
			if ([guid]::TryParse($Matches[2], [ref]$id) -and $id -ne [guid]::Empty -and $seen.Add("$kind/$id")) {
				Get-AccessReviewStatusResource $State $kind "$id"
			}
		}
	}
}

function Get-AccessReviewStatusReviewers {
	param([hashtable]$State, [object[]]$Scopes, [object[]]$Resources, [bool]$Available = $true)
	$Scopes = @($Scopes | Where-Object { $null -ne $_ })
	if (-not $Available) {
		return [pscustomobject]@{ Mode = 'Reviewer scopes unavailable'; Query = ''; Users = @(); Note = 'No reviewer type or users inferred.' }
	}
	if (-not $Scopes.Count) {
		return [pscustomobject]@{ Mode = 'Self-review / default reviewer behavior'; Query = '[]'; Users = @(); Note = 'Normally each subject reviews their own access; shared-channel reviews can use team owners. See recorded instance reviewers.' }
	}
	foreach ($scope in $Scopes) {
		$query = [string]$scope.query
		$mode = 'Unexpanded reviewer scope'
		$users = @()
		$note = 'Scope retained without inferring reviewer identities.'
		if ($scope.queryType -and $scope.queryType -ne 'MicrosoftGraph') {
			$note = "Query type '$($scope.queryType)' is not expanded."
		}
		elseif ($query -match '^/users/([0-9a-fA-F-]{36})/?$') {
			$mode = 'Named user'
			$users = @(Get-AccessReviewStatusUser $State $Matches[1])
			$note = 'Explicit user ID in this scope. The original owner/fallback policy label cannot be recovered from a user ID alone.'
		}
		elseif ($query -eq './manager') {
			$mode = 'Line managers'
			$note = 'Per-subject manager rule. No historical manager assignment inferred from current directory managers.'
		}
		elseif ($query -match '^/groups/([0-9a-fA-F-]{36})/(transitiveMembers|members)$') {
			$mode = 'Reviewer group members'
			if (-not $State.ReviewerGroups.ContainsKey($query)) {
				try {
					$members = @(Get-AccessReviewStatusCollection $State "$($State.Base)$query/microsoft.graph.user?`$select=id,displayName,userPrincipalName")
					$State.ReviewerGroups[$query] = @($members | ForEach-Object {
						Get-AccessReviewStatusUser $State $_.id $_.displayName $_.userPrincipalName 'Current reviewer-group membership'
					})
				}
				catch {
					$State.ReviewerGroups[$query] = @()
					Add-AccessReviewStatusIssue $State "Reviewer group $query" $_.Exception.Message
				}
			}
			$users = @($State.ReviewerGroups[$query])
			$note = 'Current human members of this reviewer group, not proof of historical stage assignments. Consult collection issues if identities are unavailable.'
		}
		elseif ($query -eq './owners' -or $query -match '^/(groups|servicePrincipals)/([0-9a-fA-F-]{36})/owners$') {
			$mode = 'Resource owners'
			$ownerResources = if ($query -eq './owners') { @($Resources) }
				else { @(Get-AccessReviewStatusResource $State $Matches[1] $Matches[2]) }
			$users = @($ownerResources | ForEach-Object { $_.Owners })
			$note = 'Dynamic owner rule: users below are current human owners, not proof of past stage assignment. See recorded reviewers and decision actors.'
			if (-not @($ownerResources | Where-Object { $_ }).Count -or @($ownerResources | Where-Object OwnersError).Count) {
				$note += ' Owner resolution is unavailable or partial.'
			}
		}
		elseif ($query -match '/owners\?') {
			$mode = 'Filtered owner rule'
			$note = 'Filter retained verbatim; eligibility is not inferred from the unfiltered current owner list shown separately.'
		}
		[pscustomobject]@{ Mode = $mode; Query = $query; Users = @($users); Note = $note; RawScope = $scope }
	}
}

function Get-AccessReviewStatusDecisions {
	param([hashtable]$State, [string]$Uri)
	$counts = @{ Approve = 0; Deny = 0; DontKnow = 0; NotReviewed = 0; Other = 0 }
	$raw = @{}
	$apply = @{}
	$actors = @{}
	$resourceLinks = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
	$seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
	$pages = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
	while ($Uri) {
		if (-not $Uri.StartsWith("$($State.Base)/", [System.StringComparison]::OrdinalIgnoreCase) -or -not $pages.Add($Uri)) {
			throw 'Invalid or repeated decision paging URL.'
		}
		$response = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
		if ($null -eq $response -or $response.PSObject.Properties.Name -notcontains 'value') { throw 'Invalid decision collection response.' }
		foreach ($decision in @($response.value)) {
			if (-not $decision.id) { throw 'Graph returned a decision without an ID.' }
			if (-not $seen.Add([string]$decision.id)) { continue }
			$value = if ($decision.decision) { [string]$decision.decision } else { '(missing)' }
			if (-not $raw.ContainsKey($value)) { $raw[$value] = 0 }
			$raw[$value]++
			if ($value -in @('Approve', 'Deny', 'DontKnow', 'NotReviewed')) { $counts[$value]++ } else { $counts.Other++ }
			if ($decision.applyResult) {
				$result = [string]$decision.applyResult
				if (-not $apply.ContainsKey($result)) { $apply[$result] = 0 }
				$apply[$result]++
			}
			$actor = $decision.reviewedBy
			if ($actor.id -and $actor.id -ne [guid]::Empty.ToString()) { $actors[[string]$actor.id] = $actor }
			if ($decision.resourceLink) { $null = $resourceLinks.Add([string]$decision.resourceLink) }
		}
		$Uri = $response.'@odata.nextLink'
	}
	$recorded = $counts.Approve + $counts.Deny + $counts.DontKnow
	[pscustomobject]@{
		Total = $seen.Count; Recorded = $recorded; NotReviewed = $counts.NotReviewed; Other = $counts.Other
		Approve = $counts.Approve; Deny = $counts.Deny; DontKnow = $counts.DontKnow
		Completion = if ($seen.Count) { [math]::Round(100 * $recorded / $seen.Count, 1) } else { $null }
		DecisionCounts = $raw; ApplyResults = $apply; ResourceLinks = @($resourceLinks)
		Actors = @($actors.Values | ForEach-Object { Get-AccessReviewStatusUser $State $_.id $_.displayName })
	}
}

foreach ($prefix in $Prefixes) {
	if ([string]::IsNullOrWhiteSpace($prefix)) { throw 'Prefixes must contain nonblank literal strings.' }
}
$HtmlReportPath = Resolve-AccessReviewReportPath $HtmlReportPath
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$context = Get-MgContext
if (-not $context) { throw 'Connect to Microsoft Graph first.' }
if ($context.AuthType -ne 'AppOnly' -and -not @($context.Scopes | Where-Object { $_ -in @('AccessReview.Read.All', 'AccessReview.ReadWrite.All') }).Count) {
	throw 'AccessReview.Read.All (or AccessReview.ReadWrite.All) is required.'
}
$environment = Get-MgEnvironment -Name $context.Environment
if (-not $environment.GraphEndpoint) { throw 'Unable to determine the connected Microsoft Graph endpoint.' }
$report = [pscustomobject]@{
	StartedAt = [datetime]::UtcNow.ToString('o'); FinishedAt = $null; TenantId = $context.TenantId
	Prefixes = @($Prefixes); Completed = $false
	Definitions = [System.Collections.Generic.List[object]]::new()
	Issues = [System.Collections.Generic.List[object]]::new()
}
$state = @{
	Base = "$($environment.GraphEndpoint.TrimEnd('/'))/v1.0"; Report = $report
	Users = @{}; Resources = @{}; ReviewerGroups = @{}; DefinitionId = $null; InstanceId = $null
}
try {
	$definitions = @(Get-AccessReviewStatusCollection $state "$($state.Base)/identityGovernance/accessReviews/definitions")
	foreach ($definition in $definitions) {
		$matched = @($Prefixes | Where-Object { $definition.displayName -and $definition.displayName.StartsWith($_, [System.StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1)
		if (-not $matched.Count) { continue }
		$state.DefinitionId = $definition.id
		$state.InstanceId = $null
		Write-Host "Reading all instances: $($definition.displayName)"
		$item = [pscustomobject]@{
			Prefix = $matched[0]; Id = $definition.id; DisplayName = $definition.displayName; Status = $definition.status
			Description = $definition.descriptionForReviewers; Configuration = @(); Scope = $definition.scope
			CurrentResources = @(Get-AccessReviewStatusResources $state @($definition.scope.query))
			Instances = [System.Collections.Generic.List[object]]::new(); Error = $null
		}
		$report.Definitions.Add($item)
		$settings = @($definition.stageSettings | Where-Object { $_ } | Sort-Object stageId)
		if (-not $settings.Count) {
			$settings = @([pscustomobject]@{
				stageId = 'Single-stage configuration'; reviewers = @($definition.reviewers)
				fallbackReviewers = @($definition.fallbackReviewers); durationInDays = $definition.settings.instanceDurationInDays
			})
		}
		$item.Configuration = @(
			foreach ($setting in $settings) {
				[pscustomobject]@{
					Id = $setting.stageId; DurationInDays = $setting.durationInDays; DependsOn = @($setting.dependsOn)
					Reviewers = @(Get-AccessReviewStatusReviewers $state @($setting.reviewers) $item.CurrentResources)
					Fallbacks = if (@($setting.fallbackReviewers | Where-Object { $_ }).Count) { @(Get-AccessReviewStatusReviewers $state @($setting.fallbackReviewers) $item.CurrentResources) } else { @() }
				}
			}
		)
		$definitionUri = "$($state.Base)/identityGovernance/accessReviews/definitions/$([uri]::EscapeDataString($definition.id))"
		try { $instances = @(Get-AccessReviewStatusCollection $state "$definitionUri/instances") }
		catch {
			$item.Error = $_.Exception.Message
			Add-AccessReviewStatusIssue $state 'Instance listing' $item.Error
			continue
		}
		foreach ($instance in @($instances | Sort-Object startDateTime -Descending)) {
			$state.InstanceId = $instance.id
			$instanceUri = "$definitionUri/instances/$([uri]::EscapeDataString($instance.id))"
			$row = [pscustomobject]@{
				Id = $instance.id; Status = $instance.status; Start = $instance.startDateTime; End = $instance.endDateTime; Scope = $instance.scope
				Summary = $null; DecisionError = $null; StageError = $null; ContactError = $null
				Resources = @(); Reviewers = @(); Fallbacks = @(); ContactedReviewers = @()
				Stages = [System.Collections.Generic.List[object]]::new()
			}
			$item.Instances.Add($row)
			try { $row.Summary = Get-AccessReviewStatusDecisions $state "$instanceUri/decisions?`$select=id,decision,applyResult,reviewedBy,resourceLink" }
			catch {
				$row.DecisionError = $_.Exception.Message
				Add-AccessReviewStatusIssue $state 'Instance decisions' $row.DecisionError
			}
			$row.Resources = @(Get-AccessReviewStatusResources $state (@($instance.scope.query) + @($row.Summary.ResourceLinks)))
			$row.Reviewers = @(Get-AccessReviewStatusReviewers $state @($instance.reviewers) $row.Resources ($instance.PSObject.Properties.Name -contains 'reviewers'))
			if (@($instance.fallbackReviewers | Where-Object { $_ }).Count) { $row.Fallbacks = @(Get-AccessReviewStatusReviewers $state @($instance.fallbackReviewers) $row.Resources) }
			try {
				$row.ContactedReviewers = @(
					Get-AccessReviewStatusCollection $state "$instanceUri/contactedReviewers" |
						ForEach-Object { Get-AccessReviewStatusUser $state $_.id $_.displayName $_.userPrincipalName }
				)
			}
			catch {
				$row.ContactError = $_.Exception.Message
				Add-AccessReviewStatusIssue $state 'Recorded instance reviewers' $row.ContactError
			}
			try { $runtimeStages = @(Get-AccessReviewStatusCollection $state "$instanceUri/stages") }
			catch {
				$row.StageError = $_.Exception.Message
				Add-AccessReviewStatusIssue $state 'Runtime stage listing' $row.StageError
				continue
			}
			foreach ($stage in @($runtimeStages | Sort-Object startDateTime, id)) {
				$stageRow = [pscustomobject]@{
					Id = $stage.id; Status = $stage.status; Start = $stage.startDateTime; End = $stage.endDateTime
					Reviewers = @(Get-AccessReviewStatusReviewers $state @($stage.reviewers) $row.Resources ($stage.PSObject.Properties.Name -contains 'reviewers'))
					Fallbacks = if (@($stage.fallbackReviewers | Where-Object { $_ }).Count) { @(Get-AccessReviewStatusReviewers $state @($stage.fallbackReviewers) $row.Resources) } else { @() }
					Summary = $null; Error = $null
				}
				$row.Stages.Add($stageRow)
				try {
					$stageRow.Summary = Get-AccessReviewStatusDecisions $state "$instanceUri/stages/$([uri]::EscapeDataString($stage.id))/decisions?`$select=id,decision,applyResult,reviewedBy,resourceLink"
				}
				catch {
					$stageRow.Error = $_.Exception.Message
					Add-AccessReviewStatusIssue $state "Stage $($stage.id) decisions" $stageRow.Error
				}
			}
		}
	}
	$report.Completed = $true
}
catch {
	Add-AccessReviewStatusIssue $state 'Report collection stopped' $_.Exception.Message
	throw
}
finally {
	$report.FinishedAt = [datetime]::UtcNow.ToString('o')
	try {
		Write-AccessReviewHtmlFile -Path $HtmlReportPath -Html (New-AccessReviewStatusHtml -Report $report)
	}
	catch {
		if (-not $report.Completed) { Write-Warning "Partial report export failed: $($_.Exception.Message)" -WarningAction Continue }
		else { throw }
	}
}
if ($OpenReport) { Start-Process -FilePath $HtmlReportPath -ErrorAction Stop }
$report
