BeforeAll {
	Import-Module Microsoft.Graph.Authentication
	Import-Module Microsoft.Graph.Groups
	Import-Module Microsoft.Graph.Identity.Governance
	Import-Module Microsoft.Graph.Applications
	$script:sourcePath = Join-Path $PSScriptRoot '../Create-AccessReview.ps1'
	$script:groupId = '11111111-1111-1111-1111-111111111111'
	$script:ownerId = '22222222-2222-2222-2222-222222222222'
	$script:groupOwnerId = '66666666-6666-6666-6666-666666666666'
	$script:reviewId = '33333333-3333-3333-3333-333333333333'
	function Get-RunLogs {
		Get-ChildItem -LiteralPath (Join-Path $script:testScriptRoot 'logs') -Filter 'Create-AccessReview_*.jsonl' -File
	}
	function Get-LogEntries {
		param([string]$Path)
		if (-not $Path) {
			$files = @(Get-RunLogs)
			if ($files.Count -ne 1) { throw "Expected one run log, found $($files.Count). Supply a path." }
			$Path = $files[0].FullName
		}
		@(Get-Content -LiteralPath $Path |
			ForEach-Object { $_ | ConvertFrom-Json })
	}
	function Set-NestedGroupFixture {
		$child = '44444444-4444-4444-4444-444444444444'
		$script:fixture.Children[$script:groupId] = @($child)
		$script:fixture.MemberCounts[$script:groupId] = 50
		$script:fixture.MemberCounts[$child] = 470
		foreach ($item in @(
			@{ Id = $script:groupId; Name = 'Group 1' },
			@{ Id = $child; Name = 'Group 2' },
			@{ Id = '55555555-5555-5555-5555-555555555555'; Name = 'Group 3' },
			@{ Id = '77777777-7777-7777-7777-777777777777'; Name = 'Other root' }
		)) {
			$script:fixture.Groups[$item.Id] = [pscustomobject]@{
				Id = $item.Id; DisplayName = $item.Name; OnPremisesSyncEnabled = $false; GroupTypes = @()
			}
		}
		$script:fixture.GroupOwners[$child] = @(
			[pscustomobject]@{ Id = $script:reviewId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
		)
	}
	function New-ExistingResourceReview {
		param(
			[string]$Query = "/groups/$script:groupId/members",
			[string]$Status = 'InProgress',
			[object]$Recurrence = $null
		)
		[pscustomobject]@{
			Id = '99999999-9999-9999-9999-999999999999'; DisplayName = 'Unrelated prefix - original campaign'
			Status = $Status
			Scope = @{ query = $Query; queryType = 'MicrosoftGraph' }
			Settings = @{ recurrence = $Recurrence }
		}
	}
}

Describe 'Unified access review creation and local logging' {
	BeforeEach {
		$script:testScriptRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString())
		$null = New-Item -ItemType Directory -Path $script:testScriptRoot
		$script:testScript = Join-Path $script:testScriptRoot 'Create-AccessReview.ps1'
		Copy-Item $script:sourcePath $script:testScript
		Copy-Item (Join-Path $PSScriptRoot '../AccessReview.Logging.ps1') $script:testScriptRoot
		Copy-Item (Join-Path $PSScriptRoot '../AccessReview.Reporting.ps1') $script:testScriptRoot
		$PSDefaultParameterValues['Create-AccessReview.ps1:AppOwnerIds'] = @($script:ownerId)
		$script:fixture = @{
			Count = 150
			MemberCounts = @{}
			Children = @{}
			ExtraMembers = @{}
			Groups = @{}
			GroupOwners = @{}
			Synced = $false
			Owners = @(
				[pscustomobject]@{ Id = $script:groupOwnerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } },
				[pscustomobject]@{ Id = 'service-principal'; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.servicePrincipal' } }
			)
			OwnerFailure = $false
			CreateFailure = $false
			Bodies = [System.Collections.Generic.List[object]]::new()
			ReviewId = $script:reviewId
			Dynamic = $false
			Assignments = @([pscustomobject]@{ PrincipalType = 'Group'; PrincipalId = $script:groupId })
			ApplicationOwners = @()
			ExistingNames = @()
			ExistingReviews = @()
			ExistingInstances = @()
			HostMessages = [System.Collections.Generic.List[string]]::new()
		}
		$fixture = $script:fixture
		Mock Write-Host ({
			param($Object)
			$fixture.HostMessages.Add([string]$Object)
		}.GetNewClosure())
		Mock Get-MgContext { @{ Scopes = @('AccessReview.ReadWrite.All', 'Application.Read.All') } }
		Mock Get-MgGroup ({
			param($GroupId)
			if ($fixture.Groups.ContainsKey("$GroupId")) { return $fixture.Groups["$GroupId"] }
			[pscustomobject]@{
				Id = "$GroupId"; DisplayName = 'Demo group'; OnPremisesSyncEnabled = $fixture.Synced
				GroupTypes = $(if ($fixture.Dynamic) { @('DynamicMembership') } else { @() })
			}
		}.GetNewClosure())
		Mock Get-MgGroupMember ({
			param($GroupId)
			$count = if ($fixture.MemberCounts.ContainsKey("$GroupId")) { $fixture.MemberCounts["$GroupId"] } else { $fixture.Count }
			for ($i = 1; $i -le $count; $i++) {
				[pscustomobject]@{ Id = ('00000000-0000-0000-0001-{0:d12}' -f $i); AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
			}
			foreach ($childId in $fixture.Children["$GroupId"]) {
				[pscustomobject]@{ Id = $childId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.group' } }
			}
			if ($fixture.ExtraMembers.ContainsKey("$GroupId")) { $fixture.ExtraMembers["$GroupId"] }
		}.GetNewClosure())
		Mock Get-MgGroupOwner ({
			param($GroupId)
			if ($fixture.OwnerFailure) { throw 'Owner read forbidden.' }
			if ($fixture.GroupOwners.ContainsKey("$GroupId")) { return $fixture.GroupOwners["$GroupId"] }
			$fixture.Owners
		}.GetNewClosure())
		Mock Get-MgServicePrincipal { [pscustomobject]@{ DisplayName = 'Demo app' } }
		Mock Get-MgServicePrincipalAppRoleAssignedTo ({ $fixture.Assignments }.GetNewClosure())
		Mock Get-MgServicePrincipalOwner ({ $fixture.ApplicationOwners }.GetNewClosure())
		Mock Get-MgIdentityGovernanceAccessReviewDefinition ({
			foreach ($name in $fixture.ExistingNames) { [pscustomobject]@{ Id = $fixture.ReviewId; DisplayName = $name; Status = 'Completed' } }
			$fixture.ExistingReviews
		}.GetNewClosure())
		Mock Get-MgIdentityGovernanceAccessReviewDefinitionInstance ({ $fixture.ExistingInstances }.GetNewClosure())
		Mock New-MgIdentityGovernanceAccessReviewDefinition ({
			param($BodyParameter)
			if ($fixture.CreateFailure) { throw 'Create review forbidden.' }
			$fixture.Bodies.Add($BodyParameter)
			[pscustomobject]@{ Id = $fixture.ReviewId }
		}.GetNewClosure())
	}

	It 'exports an explicitly requested HTML report without adding pipeline items (preview: <Preview>, inherited: <Inherited>)' -TestCases @(
		@{ Preview = $false; Inherited = $false }, @{ Preview = $true; Inherited = $false }, @{ Preview = $true; Inherited = $true }
	) {
		param($Preview, $Inherited)
		Set-NestedGroupFixture
		$path = Join-Path $script:testScriptRoot 'reports/campaign.html'
		$start = (Get-Date).Date.AddDays(1)
		$previewOptions = @{}
		if ($Inherited) { $WhatIfPreference = $true }
		else { $previewOptions.WhatIf = $Preview }
		$results = @(& $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner `
			-StartDate $start -EndDate $start.AddDays(11) -HtmlReportPath $path @previewOptions -Confirm:$false)
		$results | Should -HaveCount 2
		$results.GroupCountBasis | Should -Be @('DirectUsers', 'DirectUsers')
		$results[0].ReviewConfiguration.Stages.DurationInDays | Should -Be @(11)
		$results[1].ReviewConfiguration.Stages.DurationInDays | Should -Be @(6, 5)
		$results[1].ReviewConfiguration.Stages[1].Reviewers.query | Should -Be @("/users/$script:reviewId")
		$results[1].ReviewConfiguration.Stages[1].StartDate | Should -Be $start.AddDays(6).ToString('yyyy-MM-dd')
		$html = Get-Content -LiteralPath $path -Raw
		$html | Should -Match 'AR Group 2'
		$html | Should -Not -Match 'AR Group 2 \[Nested\]'
		$html | Should -Match 'Group 1 -&gt; Group 2'
		$html | Should -Match '50 direct users'
		$html | Should -Match '470 direct users'
		$html | Should -Match $(if ($Preview) { 'Preview \(WhatIf\)' } else { 'Submitted configuration' })
		if ($Preview) {
			Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
		}
		else {
			$entry = @(Get-LogEntries | Where-Object Event -eq 'GroupResult')[1]
			$entry.Result.ReviewConfiguration.Stages.DurationInDays | Should -Be @(6, 5)
			(Get-LogEntries)[0].ReportSchemaVersion | Should -Be 1
		}
	}

	It 'records the exact <Durations> window with ForceTwoStage <Forced> at <Count> users' -TestCases @(
		@{ Count = 50; Forced = $false; Durations = @(11) },
		@{ Count = 50; Forced = $true; Durations = @(6, 5) },
		@{ Count = 650; Forced = $true; Durations = @(4, 4, 3) }
	) {
		param($Count, $Forced, $Durations)
		$script:fixture.Count = $Count
		$start = (Get-Date).Date.AddDays(1)
		$result = & $script:testScript -GroupId $script:groupId -ForceTwoStage:$Forced -StartDate $start -EndDate $start.AddDays(11) -Confirm:$false
		$result.ReviewConfiguration.Stages.DurationInDays | Should -Be $Durations
		$result.ReviewConfiguration.Stages[-1].EndDate | Should -Be $start.AddDays(11).ToString('yyyy-MM-dd')
		$result.ReviewConfiguration.AutoApplyDecisionsEnabled | Should -BeTrue
		if ($Durations.Count -gt 1) {
			$result.ReviewConfiguration.Stages.DurationInDays | Should -Be $script:fixture.Bodies[0].stageSettings.durationInDays
			$result.ReviewConfiguration.Stages[0].DecisionsThatWillMoveToNextStage | Should -Be @('Approve', 'NotReviewed')
		}
	}

	It 'rejects a report destination blocked by <Blocker> before creation' -TestCases @(
		@{ Blocker = 'an existing file' }, @{ Blocker = 'a directory' }, @{ Blocker = 'a parent file' }
	) {
		param($Blocker)
		$path = Join-Path $script:testScriptRoot 'campaign.html'
		switch ($Blocker) {
			'an existing file' { Set-Content -LiteralPath $path -Value 'Keep me' }
			'a directory' { $null = New-Item -ItemType Directory -Path $path }
			'a parent file' {
				Set-Content -LiteralPath $path -Value 'Keep me'
				$path = Join-Path $path 'child.html'
			}
		}
		{ & $script:testScript -GroupId $script:groupId -HtmlReportPath $path -Confirm:$false } | Should -Throw '*Report output*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
		Should -Invoke Get-MgGroup -Times 0 -Exactly
		Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse
	}

	It 'writes a failed discovery report without pretending any reviews were created' {
		Mock Get-MgGroupMember { throw 'Discovery forbidden.' }
		$path = Join-Path $script:testScriptRoot 'failed.html'
		{ & $script:testScript -GroupId $script:groupId -HtmlReportPath $path -Confirm:$false } | Should -Throw '*Discovery forbidden*'
		$html = Get-Content -LiteralPath $path -Raw
		$html | Should -Match 'Partial or failed run'
		$html | Should -Match 'Discovery forbidden'
		$html | Should -Not -Match '<tr data-status="Created"'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'preserves original creation errors when export fails and warnings normally terminate' {
		$path = Join-Path $script:testScriptRoot 'race.html'
		Mock New-MgIdentityGovernanceAccessReviewDefinition ({
			[System.IO.File]::WriteAllText($path, 'Do not overwrite')
			throw 'Original Graph failure.'
		}.GetNewClosure())
		{ & $script:testScript -GroupId $script:groupId -HtmlReportPath $path -Confirm:$false -WarningAction Stop } |
			Should -Throw '*Original Graph failure*'
		Get-Content -LiteralPath $path -Raw | Should -Be 'Do not overwrite'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 1 -Exactly
	}

	It 'reports export failure after successful creation without retrying Graph' {
		$path = Join-Path $script:testScriptRoot 'late-collision.html'
		Mock New-MgIdentityGovernanceAccessReviewDefinition ({
			[System.IO.File]::WriteAllText($path, 'Another writer')
			[pscustomobject]@{ Id = '33333333-3333-3333-3333-333333333333' }
		}.GetNewClosure())
		$captured = [System.Collections.Generic.List[object]]::new()
		{ & $script:testScript -GroupId $script:groupId -HtmlReportPath $path -Confirm:$false | ForEach-Object { $captured.Add($_) } } |
			Should -Throw '*Reviews may already have been created*'
		$captured[0].Status | Should -Be 'Created'
		(Get-LogEntries)[1].Result.Status | Should -Be 'Created'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 1 -Exactly
	}

	It 'reviews 50 parent users and 470 child users independently in <Mode> mode (preview: <Preview>)' -TestCases @(
		@{ Mode = 'Single'; Preview = $false },
		@{ Mode = 'Single'; Preview = $true },
		@{ Mode = 'Csv'; Preview = $false },
		@{ Mode = 'CsvApplication'; Preview = $false },
		@{ Mode = 'Application'; Preview = $false },
		@{ Mode = 'Application'; Preview = $true }
	) {
		param($Mode, $Preview)
		Set-NestedGroupFixture
		$child = '44444444-4444-4444-4444-444444444444'
		$parameters = switch ($Mode) {
			'Single' { @{ GroupId = $script:groupId } }
			'Application' { @{ ServicePrincipalId = $script:groupId } }
			default {
				$csv = Join-Path $script:testScriptRoot 'nested.csv'
				$column = if ($Mode -eq 'Csv') { 'GroupId' } else { 'ServicePrincipalId' }
				"$column`n$script:groupId" | Set-Content $csv
				@{ CsvPath = $csv }
			}
		}
		$start = (Get-Date).Date.AddDays(1)
		$results = @(& $script:testScript @parameters -ReviewerType ResourceOwner -AppOwnerIds @() `
			-StartDate $start -EndDate $start.AddDays(11) -WhatIf:$Preview -Confirm:$false -WarningAction SilentlyContinue)
		$results.MemberCount | Should -Be @(50, 470)
		$results.Stages | Should -Be @(1, 2)
		$results[0].ReviewName | Should -Be $(if ($Mode -in @('Application', 'CsvApplication')) { 'AR Demo app - Group 1' } else { 'AR Group 1' })
		$results[1].ReviewName | Should -Be 'AR Group 2'
		$results[1].IsNested | Should -BeTrue
		$results[1].RootGroupIds | Should -Be @($script:groupId)
		$results[1].ParentGroupIds | Should -Be @($script:groupId)
		$results[1].GroupPaths[0].GroupIds | Should -Be @($script:groupId, $child)
		$results[1].GroupPaths[0].GroupNames | Should -Be @('Group 1', 'Group 2')
		$results[1].ReviewConfiguration.DescriptionForReviewers | Should -Match "Member of 'Group 1'\."
		$results[1].ReviewConfiguration.DescriptionForAdmins | Should -Match 'Group 1 -> Group 2'
		if ($Mode -in @('Application', 'CsvApplication')) {
			$results[1].GroupPaths[0].ServicePrincipalName | Should -Be 'Demo app'
			$results[1].ReviewConfiguration.DescriptionForReviewers | Should -Match "Related application assignments: 'Demo app'\."
			$results[1].ReviewConfiguration.DescriptionForAdmins | Should -Match "root assigned to application 'Demo app'"
		}
		else {
			$results[1].ReviewConfiguration.DescriptionForReviewers | Should -Not -Match 'application assignments'
		}
		Should -Invoke Get-MgGroupMember -Times 2 -Exactly -ParameterFilter { $All }
		if ($Preview) {
			$results.Status | Should -Be @('Would be created', 'Would be created')
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
			Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse
		}
		else {
			$script:fixture.Bodies | Should -HaveCount 2
			$script:fixture.Bodies[0].scope.AdditionalProperties.query | Should -Be "/groups/$script:groupId/members"
			$script:fixture.Bodies[1].scope.AdditionalProperties.query | Should -Be "/groups/$child/members"
			$script:fixture.Bodies[0].reviewers.query | Should -Be @("/users/$script:groupOwnerId")
			$script:fixture.Bodies[1].stageSettings[1].reviewers.query | Should -Be @("/users/$script:reviewId")
			$script:fixture.Bodies[0].settings.instanceDurationInDays | Should -Be 11
			$script:fixture.Bodies[1].stageSettings.durationInDays | Should -Be @(6, 5)
			$entries = @(Get-LogEntries | Where-Object Event -eq 'GroupResult')
			$entries[1].Result.GroupPaths[0].GroupIds | Should -Be @($script:groupId, $child)
		}
	}

	It 'discovers children even when the parent is <State>' -TestCases @(
		@{ State = 'already reviewed' }, @{ State = 'without direct users' }
	) {
		param($State)
		Set-NestedGroupFixture
		if ($State -eq 'already reviewed') { $script:fixture.ExistingNames = @('ar group 1') }
		else { $script:fixture.MemberCounts[$script:groupId] = 0 }
		$results = @(& $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false)
		$results.Status | Should -Be @('Skipped', 'Created')
		$script:fixture.Bodies | Should -HaveCount 1
		$script:fixture.Bodies[0].displayName | Should -Be 'AR Group 2'
	}

	It 'skips a nested review matching <Name> without changing that definition' -TestCases @(
		@{ Name = 'ar group 2 [nested]' },
		@{ Name = 'ar group 2' },
		@{ Name = 'AR Demo app - Group 2' }
	) {
		param($Name)
		Set-NestedGroupFixture
		$script:fixture.ExistingNames = @($Name)
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -ReviewerType ResourceOwner -Confirm:$false -WarningAction SilentlyContinue)
		$results.Status | Should -Be @('Created', 'Skipped')
		$results[1].MatchedReviewName | Should -Be $Name
		$script:fixture.Bodies | Should -HaveCount 1
	}

	It 'does not recreate root or nested reviews on a rerun' {
		Set-NestedGroupFixture
		$null = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false
		$script:fixture.ExistingNames = @($script:fixture.Bodies.displayName)
		$results = @(& $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false)
		$results.Status | Should -Be @('Skipped', 'Skipped')
		$script:fixture.Bodies | Should -HaveCount 2
	}

	It 'recognizes an old nested name when the child is now a standalone root' {
		$script:fixture.ExistingNames = @('ar demo group [nested]')
		$result = & $script:testScript -GroupId $script:groupId -Confirm:$false
		$result.Status | Should -Be 'Skipped'
		$result.MatchedReviewName | Should -Be 'ar demo group [nested]'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'retains every deep shared path while fetching and reviewing each group once' {
		Set-NestedGroupFixture
		$child = '44444444-4444-4444-4444-444444444444'
		$leaf = '55555555-5555-5555-5555-555555555555'
		$other = '77777777-7777-7777-7777-777777777777'
		$script:fixture.Children[$script:groupId] = @($child, $other)
		$script:fixture.Children[$child] = @($leaf)
		$script:fixture.Children[$other] = @($leaf)
		$script:fixture.MemberCounts[$other] = 0
		$script:fixture.MemberCounts[$leaf] = 500
		$results = @(& $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -ForceTwoStage -Confirm:$false)
		$results | Should -HaveCount 4
		$results[0].Stages | Should -Be 2
		$leafResult = @($results | Where-Object GroupId -eq $leaf)[0]
		$leafResult.Stages | Should -Be 3
		$leafResult.ParentGroupIds | Should -Be @($child, $other)
		$leafResult.GroupPaths | Should -HaveCount 2
		$pathStrings = @($leafResult.GroupPaths | ForEach-Object { $_.GroupIds -join '/' })
		$pathStrings | Should -Contain "$script:groupId/$child/$leaf"
		$pathStrings | Should -Contain "$script:groupId/$other/$leaf"
		$leafResult.ReviewConfiguration.DescriptionForReviewers | Should -Match "Member of 'Group 2', 'Other root'"
		$leafResult.ReviewConfiguration.DescriptionForAdmins | Should -Match 'Group 1 -> Group 2 -> Group 3'
		$leafResult.ReviewConfiguration.DescriptionForAdmins | Should -Match 'Group 1 -> Other root -> Group 3'
		$script:fixture.Bodies | Should -HaveCount 3
		Should -Invoke Get-MgGroup -Times 4 -Exactly
		Should -Invoke Get-MgGroupMember -Times 4 -Exactly -ParameterFilter { $All }
	}

	It 'deduplicates users and excludes non-user objects from the direct-user count' {
		Set-NestedGroupFixture
		$script:fixture.ExtraMembers[$script:groupId] = @(
			[pscustomobject]@{ Id = '00000000-0000-0000-0001-000000000001'; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } },
			[pscustomobject]@{ Id = $script:reviewId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.device' } },
			[pscustomobject]@{ Id = $script:ownerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.servicePrincipal' } }
		)
		$results = @(& $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false)
		$results.MemberCount | Should -Be @(50, 470)
	}

	It 'merges shared children with equivalent effective settings regardless of input order (reverse: <Reverse>)' -TestCases @(
		@{ Reverse = $false }, @{ Reverse = $true }
	) {
		param($Reverse)
		Set-NestedGroupFixture
		$other = '77777777-7777-7777-7777-777777777777'
		$child = '44444444-4444-4444-4444-444444444444'
		$script:fixture.Children[$other] = @($child)
		$rows = @("$script:groupId,,$script:ownerId;$script:reviewId", "$other,ResourceOwner,$script:reviewId;$script:ownerId")
		if ($Reverse) { [array]::Reverse($rows) }
		$csv = Join-Path $script:testScriptRoot 'shared.csv'
		(@('GroupId,ReviewerType,AppOwnerIds') + $rows) | Set-Content $csv
		$results = @(& $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -Confirm:$false)
		$nested = @($results | Where-Object GroupId -eq $child)
		$nested | Should -HaveCount 1
		$nested[0].ReviewName | Should -Be 'AR Group 2'
		$nested[0].RootGroupIds | Should -Be @($script:groupId, $other)
		$nested[0].GroupPaths | Should -HaveCount 2
		Should -Invoke Get-MgGroupMember -Times 1 -Exactly -ParameterFilter { "$GroupId" -eq '44444444-4444-4444-4444-444444444444' }
		$script:fixture.Bodies | Should -HaveCount 3
	}

	It 'fails before any writes when shared-child <Field> settings conflict' -TestCases @(
		@{ Field = 'ReviewerType' }, @{ Field = 'AppOwnerIds' }, @{ Field = 'Role' }
	) {
		param($Field)
		Set-NestedGroupFixture
		$other = '77777777-7777-7777-7777-777777777777'
		$script:fixture.Children[$other] = @('44444444-4444-4444-4444-444444444444')
		$type = if ($Field -eq 'ReviewerType') { 'ApplicationOwner' } else { 'ResourceOwner' }
		$owners = if ($Field -eq 'AppOwnerIds') { $script:reviewId } else { $script:ownerId }
		$role = if ($Field -eq 'Role') { 'Different role' } else { '' }
		$csv = Join-Path $script:testScriptRoot 'conflicting-parents.csv'
		"GroupId,ReviewerType,AppOwnerIds,Role`n$script:groupId,ResourceOwner,$script:ownerId,`n$other,$type,$owners,$role" | Set-Content $csv
		{ & $script:testScript -CsvPath $csv -Confirm:$false } | Should -Throw '*Conflicting inherited review settings*44444444*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
		(Get-LogEntries)[-1].Event | Should -Be 'RunFailed'
	}

	It 'does not inherit a synced parent CSV Owner and respects an explicit child Owner (explicit child: <Explicit>)' -TestCases @(
		@{ Explicit = $false }, @{ Explicit = $true }
	) {
		param($Explicit)
		Set-NestedGroupFixture
		$child = '44444444-4444-4444-4444-444444444444'
		$script:fixture.Groups[$script:groupId].OnPremisesSyncEnabled = $true
		$script:fixture.Groups[$child].OnPremisesSyncEnabled = $true
		$csv = Join-Path $script:testScriptRoot 'specific-owners.csv'
		$rows = @("GroupId,Owner,ReviewName", "$script:groupId,$script:ownerId,Keep root name")
		if ($Explicit) { $rows += "$child,$script:groupOwnerId,Keep child name" }
		$rows | Set-Content $csv
		$results = @(& $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -AppOwnerIds @() `
			-FallbackReviewerUserId $script:reviewId -Confirm:$false -WarningAction SilentlyContinue)
		$results | Should -HaveCount 2
		$results[0].ReviewName | Should -Be 'Keep root name'
		$results[1].ReviewName | Should -Be $(if ($Explicit) { 'Keep child name' } else { 'AR Group 2' })
		$results[1].CsvOwnerIds | Should -Be $(if ($Explicit) { @($script:groupOwnerId) } else { @() })
		$script:fixture.Bodies[0].reviewers.query | Should -Be @("/users/$script:ownerId")
		$script:fixture.Bodies[1].stageSettings[-1].reviewers.query | Should -Be $(if ($Explicit) { @("/users/$script:groupOwnerId") } else { @("/users/$script:reviewId") })
		$script:fixture.Bodies.settings.autoApplyDecisionsEnabled | Should -Be @($false, $false)
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
	}

	It 'fails a synced child with no group-specific Owner or explicit fallback' {
		Set-NestedGroupFixture
		$child = '44444444-4444-4444-4444-444444444444'
		$script:fixture.Groups[$script:groupId].OnPremisesSyncEnabled = $true
		$script:fixture.Groups[$child].OnPremisesSyncEnabled = $true
		$csv = Join-Path $script:testScriptRoot 'missing-child-owner.csv'
		"GroupId,Owner`n$script:groupId,$script:ownerId" | Set-Content $csv
		$results = @(& $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -Confirm:$false -WarningAction SilentlyContinue)
		$results.Status | Should -Be @('Created', 'Failed')
		$results[1].Reason | Should -Match 'Supply -FallbackReviewerUserId or CSV Owner'
	}

	It 'reports child owner lookup failures against that child' {
		Set-NestedGroupFixture
		Mock Get-MgGroupOwner { throw 'Child owners forbidden.' } -ParameterFilter { "$GroupId" -eq '44444444-4444-4444-4444-444444444444' }
		$csv = Join-Path $script:testScriptRoot 'child-owner-failure.csv'
		"GroupId`n$script:groupId" | Set-Content $csv
		$results = @(& $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -Confirm:$false)
		$results.Status | Should -Be @('Created', 'Failed')
		$results[1].Reason | Should -Match 'Child owners forbidden'
	}

	It 'retains circular connections and reviews every distinct group in <Mode> mode (preview: <Preview>)' -TestCases @(
		@{ Mode = 'Single'; Preview = $false },
		@{ Mode = 'Single'; Preview = $true },
		@{ Mode = 'Csv'; Preview = $false },
		@{ Mode = 'Csv'; Preview = $true },
		@{ Mode = 'Application'; Preview = $false },
		@{ Mode = 'Application'; Preview = $true },
		@{ Mode = 'CsvApplication'; Preview = $false },
		@{ Mode = 'CsvApplication'; Preview = $true }
	) {
		param($Mode, $Preview)
		Set-NestedGroupFixture
		$child = '44444444-4444-4444-4444-444444444444'
		$leaf = '55555555-5555-5555-5555-555555555555'
		$other = '77777777-7777-7777-7777-777777777777'
		$script:fixture.Children[$script:groupId] = @($child, $other)
		$script:fixture.Children[$child] = @($script:groupId, $leaf)
		$script:fixture.MemberCounts[$leaf] = 25
		$script:fixture.MemberCounts[$other] = 30
		$parameters = switch ($Mode) {
			'Single' { @{ GroupId = $script:groupId } }
			'Application' { @{ ServicePrincipalId = $script:groupId } }
			default {
				$csv = Join-Path $script:testScriptRoot 'circular.csv'
				$column = if ($Mode -eq 'Csv') { 'GroupId' } else { 'ServicePrincipalId' }
				"$column`n$script:groupId" | Set-Content $csv
				@{ CsvPath = $csv }
			}
		}
		$reportPath = Join-Path $script:testScriptRoot 'circular.html'
		$warnings = @()
		$results = @(& $script:testScript @parameters -ReviewerType ResourceOwner -HtmlReportPath $reportPath `
			-WhatIf:$Preview -Confirm:$false -WarningAction SilentlyContinue -WarningVariable warnings)
		$results | Should -HaveCount 4
		@($results.GroupId | Sort-Object -Unique) | Should -HaveCount 4
		$results.MemberCount | Should -Be @(50, 470, 25, 30)
		$results.Stages | Should -Be @(1, 2, 1, 1)
		$results[0].ReviewName | Should -Be $(if ($Mode -in @('Application', 'CsvApplication')) { 'AR Demo app - Group 1' } else { 'AR Group 1' })
		$results[0].ParentGroupIds | Should -Be @($child)
		$closure = @($results[0].GroupPaths | Where-Object { $_.GroupIds.Count -eq 3 })
		$closure | Should -HaveCount 1
		$closure[0].GroupIds | Should -Be @($script:groupId, $child, $script:groupId)
		$closure[0].GroupNames | Should -Be @('Group 1', 'Group 2', 'Group 1')
		$results[0].ReviewConfiguration.DescriptionForReviewers | Should -Match "Member of 'Group 2'"
		$results[0].ReviewConfiguration.DescriptionForAdmins | Should -Match 'Group 1 -> Group 2 -> Group 1'
		$cycleWarnings = @($warnings | Where-Object { "$_" -like '*Circular nested group membership detected*' })
		$cycleWarnings | Should -HaveCount 1
		"$($cycleWarnings[0])" | Should -Match "Group 1 \($script:groupId\) -> Group 2 \($child\) -> Group 1 \($script:groupId\)"
		Should -Invoke Get-MgGroup -Times 4 -Exactly
		Should -Invoke Get-MgGroupMember -Times 4 -Exactly -ParameterFilter { $All }
		$html = Get-Content -LiteralPath $reportPath -Raw
		[regex]::Matches($html, 'class="topology-node"').Count | Should -Be 4
		[regex]::Matches($html, 'class="topology-edge"').Count | Should -Be 4
		$html | Should -Match 'Group 1 -&gt; Group 2 -&gt; Group 1'
		if ($Preview) {
			$results.Status | Should -Be @('Would be created', 'Would be created', 'Would be created', 'Would be created')
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
			Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse
		}
		else {
			$results.Status | Should -Be @('Created', 'Created', 'Created', 'Created')
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 4 -Exactly
			$cycleEvents = @(Get-LogEntries | Where-Object Event -eq 'GroupCycleDetected')
			$cycleEvents | Should -HaveCount 1
			$cycleEvents[0].GroupIds | Should -Be @($script:groupId, $child, $script:groupId)
			$cycleEvents[0].GroupNames | Should -Be @('Group 1', 'Group 2', 'Group 1')
			$cycleEvents[0].ServicePrincipalId | Should -Be $(if ($Mode -in @('Application', 'CsvApplication')) { $script:groupId } else { $null })
			(Get-LogEntries)[-1].Event | Should -Be 'RunCompleted'
		}
	}

	It 'handles direct self-membership as a circular connection without duplicate reads or reviews' {
		$script:fixture.Children[$script:groupId] = @($script:groupId)
		$results = @(& $script:testScript -GroupId $script:groupId -Confirm:$false -WarningAction SilentlyContinue)
		$results | Should -HaveCount 1
		$results[0].Status | Should -Be 'Created'
		$results[0].ParentGroupIds | Should -Be @($script:groupId)
		$results[0].GroupPaths[1].GroupIds | Should -Be @($script:groupId, $script:groupId)
		Should -Invoke Get-MgGroupMember -Times 1 -Exactly
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 1 -Exactly
		@(Get-LogEntries | Where-Object Event -eq 'GroupCycleDetected') | Should -HaveCount 1
	}

	It 'does not recreate reviews when a circular component is rerun' {
		Set-NestedGroupFixture
		$script:fixture.Children['44444444-4444-4444-4444-444444444444'] = @($script:groupId)
		$null = & $script:testScript -GroupId $script:groupId -Confirm:$false -WarningAction SilentlyContinue
		$script:fixture.ExistingNames = @($script:fixture.Bodies.displayName)
		$results = @(& $script:testScript -GroupId $script:groupId -Confirm:$false -WarningAction SilentlyContinue)
		$results.Status | Should -Be @('Skipped', 'Skipped')
		$results[0].ParentGroupIds | Should -Be @('44444444-4444-4444-4444-444444444444')
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 2 -Exactly
	}

	It 'still fails closed when policies conflict across a circular component' {
		Set-NestedGroupFixture
		$child = '44444444-4444-4444-4444-444444444444'
		$script:fixture.Children[$child] = @($script:groupId)
		$csv = Join-Path $script:testScriptRoot 'circular-conflict.csv'
		"GroupId,Role`n$script:groupId,First role`n$child,Conflicting role" | Set-Content $csv
		{ & $script:testScript -CsvPath $csv -Confirm:$false -WarningAction SilentlyContinue } |
			Should -Throw '*Conflicting inherited review settings*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
		(Get-LogEntries)[-1].Event | Should -Be 'RunFailed'
	}

	It 'still stops before writes when a descendant is unreadable after a circular connection' {
		Set-NestedGroupFixture
		$script:fixture.Children['44444444-4444-4444-4444-444444444444'] = @(
			$script:groupId, '55555555-5555-5555-5555-555555555555'
		)
		Mock Get-MgGroupMember { throw 'Descendant membership denied.' } -ParameterFilter { "$GroupId" -eq '55555555-5555-5555-5555-555555555555' }
		{ & $script:testScript -GroupId $script:groupId -Confirm:$false -WarningAction SilentlyContinue } |
			Should -Throw '*Group discovery failed; no reviews were created*Descendant membership denied*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
		@(Get-LogEntries | Where-Object Event -eq 'GroupCycleDetected') | Should -HaveCount 1
		(Get-LogEntries)[-1].Event | Should -Be 'RunFailed'
	}

	It 'fails all discovery before writes for <Failure>' -TestCases @(
		@{ Failure = 'unreadable child' }, @{ Failure = 'unclassified member' }
	) {
		param($Failure)
		Set-NestedGroupFixture
		switch ($Failure) {
			'unreadable child' {
				Mock Get-MgGroupMember { throw 'Hidden membership denied.' } -ParameterFilter { "$GroupId" -eq '44444444-4444-4444-4444-444444444444' }
			}
			'unclassified member' {
				$script:fixture.ExtraMembers[$script:groupId] = @([pscustomobject]@{ Id = $script:reviewId; AdditionalProperties = @{} })
			}
		}
		{ & $script:testScript -GroupId $script:groupId -Confirm:$false } | Should -Throw '*Group discovery failed; no reviews were created*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
		(Get-LogEntries)[-1].Event | Should -Be 'RunFailed'
	}

	It 'skips dynamic nested-group reviews explicitly' {
		Set-NestedGroupFixture
		$script:fixture.Groups['44444444-4444-4444-4444-444444444444'].GroupTypes = @('DynamicMembership')
		$results = @(& $script:testScript -GroupId $script:groupId -Confirm:$false)
		$results.Status | Should -Be @('Created', 'Skipped')
		$results[1].Reason | Should -Be 'Dynamic membership group'
	}

	It 'inherits application-owner sign-off and Role without inheriting the parent name' {
		Set-NestedGroupFixture
		$csv = Join-Path $script:testScriptRoot 'inherited-role.csv'
		"GroupId,AppOwnerIds,Role,ReviewName`n$script:groupId,$script:reviewId,Payroll reader,Payroll parent" | Set-Content $csv
		$results = @(& $script:testScript -CsvPath $csv -Confirm:$false)
		$results.ReviewName | Should -Be @('Payroll parent', 'AR Group 2')
		$results[1].Role | Should -Be 'Payroll reader'
		$script:fixture.Bodies[1].descriptionForReviewers | Should -Match 'Payroll reader'
		$script:fixture.Bodies[1].stageSettings[-1].reviewers.query | Should -Be @("/users/$script:reviewId")
	}

	It 'uses multiple explicit fallback users for a child with <Kind> owners' -TestCases @(
		@{ Kind = 'no' }, @{ Kind = 'service-principal-only' }
	) {
		param($Kind)
		Set-NestedGroupFixture
		$script:fixture.GroupOwners['44444444-4444-4444-4444-444444444444'] = if ($Kind -eq 'no') { @() }
			else { @([pscustomobject]@{ Id = 'workload'; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.servicePrincipal' } }) }
		$results = @(& $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner `
			-FallbackReviewerUserId $script:ownerId, $script:groupOwnerId -Confirm:$false -WarningAction SilentlyContinue)
		$results.Status | Should -Be @('Created', 'Created')
		$script:fixture.Bodies[1].stageSettings[-1].reviewers.query | Should -Be @("/users/$script:ownerId", "/users/$script:groupOwnerId")
	}

	It 'merges equivalent duplicate CSV roots using effective policy and owner ID sets' {
		Set-NestedGroupFixture
		$csv = Join-Path $script:testScriptRoot 'equivalent-roots.csv'
		"GroupId,ReviewerType,AppOwnerIds,Owner`n$script:groupId,,$script:ownerId;$script:reviewId,$script:ownerId;$script:reviewId`n$script:groupId,ApplicationOwner,$script:reviewId;$script:ownerId,$script:reviewId;$script:ownerId" | Set-Content $csv
		$results = @(& $script:testScript -CsvPath $csv -Confirm:$false)
		$results.Status | Should -Be @('Created', 'Created')
		Should -Invoke Get-MgGroupMember -Times 2 -Exactly
	}

	It 'deduplicates shared application children and records both application origins' {
		Set-NestedGroupFixture
		$secondApp = '88888888-8888-8888-8888-888888888888'
		$other = '77777777-7777-7777-7777-777777777777'
		$child = '44444444-4444-4444-4444-444444444444'
		$script:fixture.Children[$other] = @($child)
		$script:fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId; Id = 'direct'; AppRoleId = 'role' }
		Mock Get-MgServicePrincipal { [pscustomobject]@{ DisplayName = 'Second app' } } -ParameterFilter { "$ServicePrincipalId" -eq '88888888-8888-8888-8888-888888888888' }
		Mock Get-MgServicePrincipalAppRoleAssignedTo {
			[pscustomobject]@{ PrincipalType = 'Group'; PrincipalId = '77777777-7777-7777-7777-777777777777' }
		} -ParameterFilter { "$ServicePrincipalId" -eq '88888888-8888-8888-8888-888888888888' }
		$csv = Join-Path $script:testScriptRoot 'shared-apps.csv'
		"ServicePrincipalId`n$script:groupId`n$secondApp" | Set-Content $csv
		$results = @(& $script:testScript -CsvPath $csv -Confirm:$false -WarningAction SilentlyContinue)
		$results.ReviewName | Should -Be @('AR Demo app - Group 1', 'AR Group 2', 'AR Demo app', 'AR Second app - Other root')
		$nested = @($results | Where-Object GroupId -eq $child)[0]
		$nested.SourceServicePrincipalIds | Should -Be @($script:groupId, $secondApp)
		$nested.GroupPaths.ServicePrincipalId | Should -Be @($script:groupId, $secondApp)
		$nested.GroupPaths.ServicePrincipalName | Should -Be @('Demo app', 'Second app')
		$nested.GroupPaths | Should -HaveCount 2
		$nested.ReviewConfiguration.DescriptionForReviewers | Should -Match "Member of 'Group 1', 'Other root'"
		$nested.ReviewConfiguration.DescriptionForReviewers | Should -Match "'Demo app', 'Second app'"
		$nested.ReviewConfiguration.DescriptionForAdmins | Should -Match "Group 1 -> Group 2 \[root assigned to application 'Demo app'"
		$nested.ReviewConfiguration.DescriptionForAdmins | Should -Match "Other root -> Group 2 \[root assigned to application 'Second app'"
		Should -Invoke Get-MgGroupMember -Times 3 -Exactly
		$script:fixture.Bodies | Should -HaveCount 4
	}

	It 'uses the requested reviewers at <Count> members with <Type> owner routing' -TestCases @(
		@{ Count = 1; Type = 'ApplicationOwner' },
		@{ Count = 150; Type = 'ApplicationOwner' },
		@{ Count = 150; Type = 'ResourceOwner' },
		@{ Count = 151; Type = 'ApplicationOwner' },
		@{ Count = 151; Type = 'ResourceOwner' },
		@{ Count = 499; Type = 'ResourceOwner' },
		@{ Count = 500; Type = 'ApplicationOwner' },
		@{ Count = 500; Type = 'ResourceOwner' }
	) {
		param($Count, $Type)
		$script:fixture.Count = $Count
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType $Type `
			-SelfReviewDurationInDays 3 -SelectedReviewerDurationInDays 7 -ApplicationOwnerDurationInDays 4 -Confirm:$false
		$body = $script:fixture.Bodies[0]
		$expectedOwner = if ($Type -eq 'ResourceOwner') { $script:groupOwnerId } else { $script:ownerId }
		if ($Count -le 150) {
			$result.Stages | Should -Be 1
			$body.reviewers.query | Should -Be @("/users/$expectedOwner")
			$body.settings.instanceDurationInDays | Should -Be 4
		}
		else {
			$result.Stages | Should -Be $(if ($Count -ge 500) { 3 } else { 2 })
			$body.stageSettings[-2].reviewers[0].query | Should -Be './manager'
			$body.stageSettings[-1].reviewers.query | Should -Be @("/users/$expectedOwner")
			$body.stageSettings[-1].dependsOn | Should -Be @([string]($result.Stages - 1))
			$body.settings.instanceDurationInDays | Should -Be $(if ($Count -ge 500) { 14 } else { 11 })
			if ($Count -ge 500) { $body.stageSettings[0].reviewers | Should -HaveCount 0 }
		}
	}

	It 'fails a selected ownerless group without explicit fallback users at <Count> members' -TestCases @(
		@{ Count = 150 }, @{ Count = 500 }
	) {
		param($Count)
		$script:fixture.Count = $Count
		$script:fixture.Owners = @()
		{ & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false } |
			Should -Throw '*no user owners*'
		(Get-LogEntries)[1].Result.Status | Should -Be 'Failed'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'uses fallback users for groups without human owners (SP only: <SpOnly>, count: <Count>, preview: <Preview>)' -TestCases @(
		@{ SpOnly = $true; Count = 150; Preview = $false },
		@{ SpOnly = $true; Count = 151; Preview = $false },
		@{ SpOnly = $true; Count = 151; Preview = $true },
		@{ SpOnly = $true; Count = 500; Preview = $false },
		@{ SpOnly = $false; Count = 150; Preview = $false }
	) {
		param($SpOnly, $Count, $Preview)
		$script:fixture.Count = $Count
		$script:fixture.Owners = @($script:fixture.Owners | Where-Object {
			$SpOnly -and $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.servicePrincipal'
		})
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -AppOwnerIds @() `
			-FallbackReviewerUserId @($script:reviewId, $script:ownerId) -WhatIf:$Preview -Confirm:$false -WarningAction SilentlyContinue
		$result.Status | Should -Be $(if ($Preview) { 'Would be created' } else { 'Created' })
		$result.StageNames | Should -Be $(if ($Count -ge 500) { 'Self-review -> Line managers -> Fallback reviewers (no group user owners)' }
			elseif ($Count -gt 150) { 'Line managers -> Fallback reviewers (no group user owners)' }
			else { 'Fallback reviewers (no group user owners)' })
		if ($Preview) {
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
		}
		else {
			$body = $script:fixture.Bodies[0]
			$expected = @("/users/$script:reviewId", "/users/$script:ownerId")
			if ($Count -le 150) { $body.reviewers.query | Should -Be $expected }
			else {
				$body.stageSettings[-2].reviewers[0].query | Should -Be './manager'
				$body.stageSettings[-2].fallbackReviewers.query | Should -Be $expected
				$body.stageSettings[-1].reviewers.query | Should -Be $expected
			}
			$body.settings.autoApplyDecisionsEnabled | Should -BeTrue
			(Get-LogEntries)[1].Result.StageNames | Should -Be $result.StageNames
		}
	}

	It 'allows a group-owner review without application owners and uses group owners as manager fallback' {
		$script:fixture.Count = 151
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -AppOwnerIds @() -Confirm:$false
		$result.Status | Should -Be 'Created'
		$script:fixture.Bodies[0].stageSettings[0].fallbackReviewers.query | Should -Be @("/users/$script:groupOwnerId")
		$script:fixture.Bodies[0].stageSettings[1].reviewers.query | Should -Be @("/users/$script:groupOwnerId")
	}

	It 'uses group owners as the third stage without application owners and splits the date window' {
		$script:fixture.Count = 500
		$start = (Get-Date).Date.AddDays(1)
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -AppOwnerIds @() `
			-StartDate $start -EndDate $start.AddDays(11) -Confirm:$false
		$result.Status | Should -Be 'Created'
		$result.StageNames | Should -Be 'Self-review -> Line managers -> Resource owners (group owners)'
		$body = $script:fixture.Bodies[0]
		$body.stageSettings[0].reviewers | Should -HaveCount 0
		$body.stageSettings[1].reviewers[0].query | Should -Be './manager'
		$body.stageSettings[1].fallbackReviewers.query | Should -Be @("/users/$script:groupOwnerId")
		$body.stageSettings[2].reviewers.query | Should -Be @("/users/$script:groupOwnerId")
		$body.stageSettings.durationInDays | Should -Be @(4, 4, 3)
	}

	It 'fails ResourceOwner routing for synchronized groups without silently substituting reviewers' {
		$script:fixture.Synced = $true
		{ & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false } |
			Should -Throw '*Supply -FallbackReviewerUserId*'
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'uses explicit fallback users for synced groups in <Mode> mode at <Count> members (preview: <Preview>)' -TestCases @(
		@{ Mode = 'Single'; Count = 150; Preview = $false },
		@{ Mode = 'Single'; Count = 151; Preview = $false },
		@{ Mode = 'Single'; Count = 500; Preview = $false },
		@{ Mode = 'Csv'; Count = 150; Preview = $false },
		@{ Mode = 'Application'; Count = 151; Preview = $false },
		@{ Mode = 'Csv'; Count = 151; Preview = $true }
	) {
		param($Mode, $Count, $Preview)
		$script:fixture.Synced = $true
		$script:fixture.Count = $Count
		$parameters = switch ($Mode) {
			'Single' { @{ GroupId = $script:groupId } }
			'Csv' {
				$csv = Join-Path $script:testScriptRoot 'synced.csv'
				"GroupId`n$script:groupId" | Set-Content $csv
				@{ CsvPath = $csv }
			}
			'Application' { @{ ServicePrincipalId = $script:groupId } }
		}
		$fallbackIds = @($script:reviewId, $script:groupOwnerId, $script:reviewId)
		$result = & $script:testScript @parameters -ReviewerType ResourceOwner -FallbackReviewerUserId $fallbackIds `
			-WhatIf:$Preview -Confirm:$false -WarningAction SilentlyContinue
		$result.Status | Should -Be $(if ($Preview) { 'Would be created' } else { 'Created' })
		$result.StageNames | Should -Be $(if ($Count -ge 500) { 'Self-review -> Line managers -> Fallback reviewers (synced group)' }
			elseif ($Count -gt 150) { 'Line managers -> Fallback reviewers (synced group)' }
			else { 'Fallback reviewers (synced group)' })
		$result.Reason | Should -Match 'decisions recorded only'
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
		if ($Preview) {
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
		}
		else {
			$body = $script:fixture.Bodies[0]
			$expectedFallback = @("/users/$script:reviewId", "/users/$script:groupOwnerId")
			$body.settings.autoApplyDecisionsEnabled | Should -BeFalse
			$body.settings.ContainsKey('applyActions') | Should -BeFalse
			if ($Count -le 150) {
				$body.reviewers.query | Should -Be $expectedFallback
			}
			else {
				$body.stageSettings[-2].reviewers[0].query | Should -Be './manager'
				$body.stageSettings[-2].fallbackReviewers.query | Should -Be $expectedFallback
				$body.stageSettings[-1].reviewers.query | Should -Be $expectedFallback
			}
			(Get-LogEntries)[0].FallbackReviewerUserId | Should -Be @($script:reviewId, $script:groupOwnerId)
			(Get-LogEntries)[0].Syntax | Should -Match '-FallbackReviewerUserId @\('
		}
	}

	It 'accepts a single fallback user for a synced group without application owners' {
		$script:fixture.Synced = $true
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -AppOwnerIds @() `
			-FallbackReviewerUserId $script:reviewId -Confirm:$false -WarningAction SilentlyContinue
		$result.Status | Should -Be 'Created'
		$script:fixture.Bodies[0].reviewers.query | Should -Be @("/users/$script:reviewId")
		(Get-LogEntries)[0].FallbackReviewerUserId | Should -Be $script:reviewId
	}

	It 'uses CSV owners before fallbacks for synced groups at <Count> members (preview: <Preview>)' -TestCases @(
		@{ Count = 150; Preview = $false },
		@{ Count = 151; Preview = $false },
		@{ Count = 500; Preview = $false },
		@{ Count = 151; Preview = $true }
	) {
		param($Count, $Preview)
		$script:fixture.Synced = $true
		$script:fixture.Count = $Count
		$csv = Join-Path $script:testScriptRoot 'synced-owners.csv'
		"GroupId,Owner`n$script:groupId,`"$script:groupOwnerId,$script:reviewId,$script:groupOwnerId`"" | Set-Content $csv
		$result = & $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -FallbackReviewerUserId $script:ownerId `
			-WhatIf:$Preview -Confirm:$false -WarningAction SilentlyContinue
		$result.Status | Should -Be $(if ($Preview) { 'Would be created' } else { 'Created' })
		$result.CsvOwnerIds | Should -Be @($script:groupOwnerId, $script:reviewId)
		$result.StageNames | Should -Be $(if ($Count -ge 500) { 'Self-review -> Line managers -> CSV owners (synced group)' }
			elseif ($Count -gt 150) { 'Line managers -> CSV owners (synced group)' }
			else { 'CSV owners (synced group)' })
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
		if ($Preview) {
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
			Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse
		}
		else {
			$body = $script:fixture.Bodies[0]
			$body.settings.autoApplyDecisionsEnabled | Should -BeFalse
			$body.settings.ContainsKey('applyActions') | Should -BeFalse
			if ($Count -le 150) {
				$body.reviewers.query | Should -Be @("/users/$script:groupOwnerId", "/users/$script:reviewId")
			}
			else {
				$body.stageSettings[-2].reviewers[0].query | Should -Be './manager'
				$body.stageSettings[-2].fallbackReviewers.query | Should -Be @("/users/$script:ownerId")
				$expected = @("/users/$script:groupOwnerId", "/users/$script:reviewId")
				$body.stageSettings[-1].reviewers.query | Should -Be $expected
			}
			(Get-LogEntries)[1].Result.CsvOwnerIds | Should -Be @($script:groupOwnerId, $script:reviewId)
		}
	}

	It 'uses CSV owners as manager fallback if no other fallback or application owner is supplied' {
		$script:fixture.Synced = $true
		$script:fixture.Count = 151
		$csv = Join-Path $script:testScriptRoot 'synced-owner-only.csv'
		"GroupId;Owner`n$script:groupId;$script:groupOwnerId,$script:reviewId" | Set-Content $csv
		$result = & $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -AppOwnerIds @() -Confirm:$false -WarningAction SilentlyContinue
		$result.Status | Should -Be 'Created'
		$body = $script:fixture.Bodies[0]
		$body.stageSettings[0].fallbackReviewers.query | Should -Be @("/users/$script:groupOwnerId", "/users/$script:reviewId")
		$body.stageSettings[1].reviewers.query | Should -Be @("/users/$script:groupOwnerId", "/users/$script:reviewId")
	}

	It 'uses the explicit fallback when the CSV Owner cell is blank' {
		$script:fixture.Synced = $true
		$csv = Join-Path $script:testScriptRoot 'blank-owner.csv'
		"GroupId,Owner`n$script:groupId," | Set-Content $csv
		$result = & $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -FallbackReviewerUserId $script:reviewId `
			-Confirm:$false -WarningAction SilentlyContinue
		$result.StageNames | Should -Be 'Fallback reviewers (synced group)'
		$script:fixture.Bodies[0].reviewers.query | Should -Be @("/users/$script:reviewId")
	}

	It 'keeps cloud group owners when a CSV Owner is supplied' {
		$csv = Join-Path $script:testScriptRoot 'cloud-owner.csv'
		"GroupId,Owner`n$script:groupId,$script:reviewId" | Set-Content $csv
		$result = & $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -Confirm:$false
		$result.StageNames | Should -Be 'Resource owners (group owners)'
		$script:fixture.Bodies[0].reviewers.query | Should -Be @("/users/$script:groupOwnerId")
		Should -Invoke Get-MgGroupOwner -Times 1 -Exactly
	}

	It 'applies application-row CSV owners only to discovered synced groups' {
		$script:fixture.Synced = $true
		$script:fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
		$script:fixture.ApplicationOwners = @(
			[pscustomobject]@{ Id = $script:ownerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
		)
		$csv = Join-Path $script:testScriptRoot 'application-owner.csv'
		"ServicePrincipalId,ReviewerType,Owner`n$script:groupId,ResourceOwner,$script:reviewId" | Set-Content $csv
		$results = @(& $script:testScript -CsvPath $csv -Confirm:$false -WarningAction SilentlyContinue)
		$reviews = @($results | Where-Object TargetType -in @('GroupMembership', 'ApplicationDirectUsers'))
		$reviews.Status | Should -Be @('Created', 'Created')
		$script:fixture.Bodies[0].reviewers.query | Should -Be @("/users/$script:reviewId")
		$script:fixture.Bodies[1].reviewers.query | Should -Be @("/users/$script:ownerId")
	}

	It 'fails an invalid CSV Owner instead of using a fallback (value: <Value>)' -TestCases @(
		@{ Value = 'not-a-guid' },
		@{ Value = '00000000-0000-0000-0000-000000000000' },
		@{ Value = '66666666-6666-6666-6666-666666666666;bad-id' }
	) {
		param($Value)
		$script:fixture.Synced = $true
		$csv = Join-Path $script:testScriptRoot 'invalid-owner.csv'
		"GroupId,Owner`n$script:groupId,$Value" | Set-Content $csv
		$result = & $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -FallbackReviewerUserId $script:ownerId -Confirm:$false
		$result.Status | Should -Be 'Failed'
		$result.Reason | Should -Match 'CSV Owner contains an invalid user object GUID'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'rejects duplicate CSV targets with conflicting Owner values' {
		$csv = Join-Path $script:testScriptRoot 'conflicting-owners.csv'
		"GroupId,Owner`n$script:groupId,$script:groupOwnerId`n$script:groupId,$script:reviewId" | Set-Content $csv
		{ & $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -Confirm:$false } |
			Should -Throw '*conflicting settings*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'isolates per-row CSV owners and continues after an invalid owner' {
		$second = '44444444-4444-4444-4444-444444444444'
		$third = '55555555-5555-5555-5555-555555555555'
		$script:fixture.Synced = $true
		$csv = Join-Path $script:testScriptRoot 'owner-batch.csv'
		"GroupId,ReviewName,Owner`n$script:groupId,Invalid,bad-id`n$second,Second,$script:groupOwnerId`n$third,Third,$script:reviewId" |
			Set-Content $csv
		$results = @(& $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -Confirm:$false -WarningAction SilentlyContinue)
		$results.Status | Should -Be @('Failed', 'Created', 'Created')
		$results[0].Reason | Should -Match 'CSV Owner contains an invalid user object GUID'
		$script:fixture.Bodies[0].reviewers.query | Should -Be @("/users/$script:groupOwnerId")
		$script:fixture.Bodies[1].reviewers.query | Should -Be @("/users/$script:reviewId")
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 2 -Exactly
	}

	It 'uses multiple manager fallbacks without replacing owners of a cloud group' {
		$script:fixture.Count = 151
		$null = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner `
			-FallbackReviewerUserId @($script:reviewId, $script:ownerId) -Confirm:$false
		$body = $script:fixture.Bodies[0]
		$body.stageSettings[0].fallbackReviewers.query | Should -Be @("/users/$script:reviewId", "/users/$script:ownerId")
		$body.stageSettings[1].reviewers.query | Should -Be @("/users/$script:groupOwnerId")
	}

	It 'rejects invalid entries in a multi-user fallback list before creation' -TestCases @(
		@{ Fallbacks = @() }, @{ Fallbacks = @('33333333-3333-3333-3333-333333333333', '00000000-0000-0000-0000-000000000000') }
	) {
		param($Fallbacks)
		{ & $script:testScript -GroupId $script:groupId -FallbackReviewerUserId $Fallbacks -Confirm:$false } |
			Should -Throw '*FallbackReviewerUserId must be a nonempty*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'forces a manager stage at <Count> members with <Type> routing without reducing large reviews' -TestCases @(
		@{ Count = 1; Type = 'ApplicationOwner' },
		@{ Count = 150; Type = 'ApplicationOwner' },
		@{ Count = 150; Type = 'ResourceOwner' },
		@{ Count = 151; Type = 'ResourceOwner' },
		@{ Count = 500; Type = 'ResourceOwner' }
	) {
		param($Count, $Type)
		$script:fixture.Count = $Count
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType $Type -ForceTwoStage `
			-SelfReviewDurationInDays 3 -SelectedReviewerDurationInDays 7 -ApplicationOwnerDurationInDays 4 -Confirm:$false
		$body = $script:fixture.Bodies[0]
		$result.Stages | Should -Be $(if ($Count -ge 500) { 3 } else { 2 })
		$body.stageSettings.durationInDays | Should -Be $(if ($Count -ge 500) { @(3, 7, 4) } else { @(7, 4) })
		$body.stageSettings[-2].reviewers[0].query | Should -Be './manager'
		$body.stageSettings[-1].dependsOn | Should -Be @([string]($result.Stages - 1))
		$expectedOwner = if ($Type -eq 'ResourceOwner') { $script:groupOwnerId } else { $script:ownerId }
		$body.stageSettings[-1].reviewers.query | Should -Be @("/users/$expectedOwner")
		$body.settings.instanceDurationInDays | Should -Be $(if ($Count -ge 500) { 14 } else { 11 })
	}

	It 'preserves multiple fallbacks for a forced small resource-owner review (synced: <Synced>)' -TestCases @(
		@{ Synced = $false }, @{ Synced = $true }
	) {
		param($Synced)
		$script:fixture.Synced = $Synced
		$script:fixture.Owners = @()
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -AppOwnerIds @() `
			-ForceTwoStage -FallbackReviewerUserId @($script:ownerId, $script:reviewId) -Confirm:$false
		$result.Stages | Should -Be 2
		$body = $script:fixture.Bodies[0]
		$body.stageSettings[0].fallbackReviewers.query | Should -Be @("/users/$script:ownerId", "/users/$script:reviewId")
		$body.stageSettings[1].reviewers.query | Should -Be @("/users/$script:ownerId", "/users/$script:reviewId")
		$body.settings.autoApplyDecisionsEnabled | Should -Be (-not $Synced)
	}

	It 'forces two stages and splits the date window for every target in <Mode> mode' -TestCases @(
		@{ Mode = 'Csv' }, @{ Mode = 'Application' }
	) {
		param($Mode)
		$script:fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
		$parameters = @{ ServicePrincipalId = $script:groupId }
		if ($Mode -eq 'Csv') {
			$csv = Join-Path $script:testScriptRoot 'forced-stages.csv'
			"ServicePrincipalId`n$script:groupId" | Set-Content -LiteralPath $csv
			$parameters = @{ CsvPath = $csv }
		}
		$start = (Get-Date).Date.AddDays(1)
		$results = @(& $script:testScript @parameters -ForceTwoStage -StartDate $start -EndDate $start.AddDays(11) -Confirm:$false)
		$results.Status | Should -Be @('Created', 'Created')
		$results.Stages | Should -Be @(2, 2)
		foreach ($body in $script:fixture.Bodies) {
			$body.stageSettings.durationInDays | Should -Be @(6, 5)
			$body.stageSettings[0].reviewers[0].query | Should -Be './manager'
			$body.settings.instanceDurationInDays | Should -Be 11
			([datetime]$body.settings.recurrence.range.endDate).Date | Should -Be $start.AddDays(11)
		}
	}

	It 'rejects a one-day window for a forced two-stage review' {
		$start = (Get-Date).Date.AddDays(1)
		{ & $script:testScript -GroupId $script:groupId -ForceTwoStage -StartDate $start -EndDate $start.AddDays(1) } |
			Should -Throw '*at least 2 days*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'still skips empty groups when forcing two stages' {
		$script:fixture.Count = 0
		$result = & $script:testScript -GroupId $script:groupId -ForceTwoStage -Confirm:$false
		$result.Status | Should -Be 'Skipped'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'uses the application-owner duration for a single-stage review' {
		$script:fixture.Count = 22
		$null = & $script:testScript -GroupId $script:groupId -ApplicationOwnerDurationInDays 9 -Confirm:$false
		$script:fixture.Bodies[0].settings.instanceDurationInDays | Should -Be 9
	}

	It 'uses the one-time date window instead of durations at <Count> members over <Days> days' -TestCases @(
		@{ Count = 150; Days = 11; Durations = @(11) },
		@{ Count = 151; Days = 11; Durations = @(6, 5) },
		@{ Count = 500; Days = 11; Durations = @(4, 4, 3) },
		@{ Count = 500; Days = 12; Durations = @(4, 4, 4) },
		@{ Count = 500; Days = 3; Durations = @(1, 1, 1) }
	) {
		param($Count, $Days, $Durations)
		$script:fixture.Count = $Count
		$start = (Get-Date).Date.AddDays(1)
		$end = $start.AddDays($Days)
		$null = & $script:testScript -GroupId $script:groupId -StartDate $start -EndDate $end `
			-SelfReviewDurationInDays 3 -SelectedReviewerDurationInDays 7 -ApplicationOwnerDurationInDays 4 `
			-WarningAction SilentlyContinue -Confirm:$false
		$body = $script:fixture.Bodies[0]
		$body.settings.instanceDurationInDays | Should -Be $Days
		if ($Count -gt 150) {
			$body.stageSettings.durationInDays | Should -Be $Durations
		}
		$body.settings.recurrence.pattern.type | Should -BeNullOrEmpty
		([datetime]$body.settings.recurrence.range.startDate).Date | Should -Be $start
		([datetime]$body.settings.recurrence.range.endDate).Date | Should -Be $end
	}

	It 'rejects a one-time window too short for <Count> members over <Days> days' -TestCases @(
		@{ Count = 150; Days = 0; Minimum = 1 },
		@{ Count = 150; Days = -1; Minimum = 1 },
		@{ Count = 500; Days = 2; Minimum = 3 }
	) {
		param($Count, $Days, $Minimum)
		$script:fixture.Count = $Count
		$start = (Get-Date).Date.AddDays(1)
		{ & $script:testScript -GroupId $script:groupId -StartDate $start -EndDate $start.AddDays($Days) -Confirm:$false } |
			Should -Throw "*at least $Minimum days*"
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
		(Get-LogEntries)[1].Result.Status | Should -Be 'Failed'
	}

	It 'splits the date window independently for group and direct-user targets in <Mode> mode' -TestCases @(
		@{ Mode = 'Csv' }, @{ Mode = 'Application' }
	) {
		param($Mode)
		$script:fixture.Count = 500
		$script:fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
		$parameters = @{ ServicePrincipalId = $script:groupId }
		if ($Mode -eq 'Csv') {
			$csv = Join-Path $script:testScriptRoot 'date-window.csv'
			"ServicePrincipalId`n$script:groupId" | Set-Content -LiteralPath $csv
			$parameters = @{ CsvPath = $csv }
		}
		$start = (Get-Date).Date.AddDays(1)
		$results = @(& $script:testScript @parameters -StartDate $start -EndDate $start.AddDays(11) -Confirm:$false)
		$results | Should -HaveCount 2
		$results.Status | Should -Be @('Created', 'Created')
		$groupBody = @($script:fixture.Bodies | Where-Object { $_.stageSettings.Count -eq 3 })[0]
		$groupBody.stageSettings.durationInDays | Should -Be @(4, 4, 3)
		$script:fixture.Bodies.settings.instanceDurationInDays | Should -Be @(11, 11)
		foreach ($body in $script:fixture.Bodies) {
			([datetime]$body.settings.recurrence.range.endDate).Date | Should -Be $start.AddDays(11)
		}
	}

	It 'validates recurring schedules against the actual duration sum (days: <Days>)' -TestCases @(
		@{ Days = 11; Valid = $false }, @{ Days = 12; Valid = $true }
	) {
		param($Days, $Valid)
		$script:fixture.Count = 151
		$parameters = @{
			GroupId = $script:groupId
			SelectedReviewerDurationInDays = 7
			ApplicationOwnerDurationInDays = 5
			Frequency = 'Monthly'
			EndDate = (Get-Date).Date.AddDays($Days)
			Confirm = $false
		}
		if ($Valid) {
			$null = & $script:testScript @parameters
			$script:fixture.Bodies[0].settings.instanceDurationInDays | Should -Be 12
		}
		else {
			{ & $script:testScript @parameters } | Should -Throw '*at least 12 days*'
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
		}
	}

	It 'rejects nonpositive <Parameter> durations before creating reviews' -TestCases @(
		@{ Parameter = 'SelfReviewDurationInDays'; Value = 0 },
		@{ Parameter = 'SelectedReviewerDurationInDays'; Value = -1 },
		@{ Parameter = 'ApplicationOwnerDurationInDays'; Value = 0 }
	) {
		param($Parameter, $Value)
		$parameters = @{ GroupId = $script:groupId; $Parameter = $Value }
		{ & $script:testScript @parameters } | Should -Throw
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'accepts a review-name prefix in <Mode> mode' -TestCases @(
		@{ Mode = 'Single' }, @{ Mode = 'Csv' }, @{ Mode = 'Application' }
	) {
		param($Mode)
		$parameters = switch ($Mode) {
			'Single' { @{ GroupId = $script:groupId } }
			'Csv' {
				$csv = Join-Path $script:testScriptRoot 'prefix.csv'
				"GroupId`n$script:groupId" | Set-Content $csv
				@{ CsvPath = $csv }
			}
			'Application' {
				$script:fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
				@{ ServicePrincipalId = $script:groupId }
			}
		}
		$results = @(& $script:testScript @parameters -ReviewNamePrefix 'Custom' -Confirm:$false)
		if ($Mode -eq 'Application') {
			$results.ReviewName | Should -Be @('Custom Demo app - Demo group', 'Custom Demo app')
		}
		else { $results[0].ReviewName | Should -Be 'Custom Demo group' }
		(Get-LogEntries)[0].Syntax | Should -Match "-ReviewNamePrefix 'Custom'"
	}

	It 'allows an empty prefix without a leading space' {
		$result = & $script:testScript -GroupId $script:groupId -ReviewNamePrefix '' -Confirm:$false
		$result.ReviewName | Should -Be 'Demo group'
	}

	It 'preserves explicit review names when a prefix is supplied' {
		$result = & $script:testScript -GroupId $script:groupId -ReviewName 'Explicit review' -ReviewNamePrefix 'Custom' -Confirm:$false
		$result.ReviewName | Should -Be 'Explicit review'
	}

	It 'preserves default manager routing at the two-stage threshold' {
		$script:fixture.Count = 151
		$result = & $script:testScript -GroupId $script:groupId -Confirm:$false
		$result.Status | Should -Be 'Created'
		$result.ReviewId | Should -Be $script:reviewId
		$body = $script:fixture.Bodies[0]
		$body.stageSettings.Count | Should -Be 2
		$body.stageSettings[0].reviewers[0].query | Should -Be './manager'
		$body.stageSettings[0].reviewers[0].queryRoot | Should -Be 'decisions'
		$body.stageSettings[1].reviewers[0].query | Should -Be "/users/$script:ownerId"
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
	}

	It 'reports the existing single-group count of <Count> and logs only real runs (preview: <Preview>)' -TestCases @(
		@{ Count = 150; Preview = $false },
		@{ Count = 150; Preview = $true },
		@{ Count = 1; Preview = $false }
	) {
		param($Count, $Preview)
		$script:fixture.Count = $Count
		$result = & $script:testScript -GroupId $script:groupId -WhatIf:$Preview -Confirm:$false
		$result.MemberCount | Should -Be $Count
		if ($Preview) {
			Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse
		}
		else {
			(Get-LogEntries)[1].Result.MemberCount | Should -Be $Count
		}
		$label = if ($Preview) { 'reviews that would be created' } else { 'created reviews' }
		$script:fixture.HostMessages[-1] | Should -Be "Total direct users in group ${label}: $Count"
		Should -Invoke Get-MgGroupMember -Times 1 -Exactly -ParameterFilter {
			"$GroupId" -eq $script:groupId -and $All
		}
		if ($Preview) {
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
		}
	}

	It 'sums different group counts below the CSV summary without counting duplicate targets (preview: <Preview>)' -TestCases @(
		@{ Preview = $false }, @{ Preview = $true }
	) {
		param($Preview)
		$second = '44444444-4444-4444-4444-444444444444'
		$script:fixture.MemberCounts[$second] = 75
		Mock Get-MgGroup {
			[pscustomobject]@{ DisplayName = 'Second group'; OnPremisesSyncEnabled = $false; GroupTypes = @() }
		} -ParameterFilter { "$GroupId" -eq '44444444-4444-4444-4444-444444444444' }
		$csv = Join-Path $script:testScriptRoot 'member-counts.csv'
		"GroupId`n$script:groupId`n$second`n$script:groupId" | Set-Content $csv
		$results = @(& $script:testScript -CsvPath $csv -WhatIf:$Preview -Confirm:$false)
		$results | Should -HaveCount 2
		$results.MemberCount | Should -Be @(150, 75)
		$label = if ($Preview) { 'would be created' } else { 'created' }
		$script:fixture.HostMessages[-2] | Should -Be "2 group(s): 2 $label, 0 skipped, 0 failed."
		$script:fixture.HostMessages[-1] | Should -Match ': 225$'
		Should -Invoke Get-MgGroupMember -Times 2 -Exactly
	}

	It 'skips an existing review in <Mode> mode (preview: <Preview>)' -TestCases @(
		@{ Mode = 'Single'; Preview = $false }, @{ Mode = 'Single'; Preview = $true },
		@{ Mode = 'Csv'; Preview = $false }, @{ Mode = 'Csv'; Preview = $true }
	) {
		param($Mode, $Preview)
		$script:fixture.ExistingNames = @('ar demo group')
		if ($Mode -eq 'Single') {
			$parameters = @{ GroupId = $script:groupId }
		}
		else {
			$csv = Join-Path $script:testScriptRoot 'existing.csv'
			"GroupId`n$script:groupId" | Set-Content $csv
			$parameters = @{ CsvPath = $csv }
		}
		$result = & $script:testScript @parameters -WhatIf:$Preview -Confirm:$false
		$result.Status | Should -Be 'Skipped'
		$result.Reason | Should -Match 'already exists'
		Should -Invoke Get-MgIdentityGovernanceAccessReviewDefinition -Times 1 -Exactly
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'skips the same resource under another name in <Mode> mode regardless of reviewer policy (preview: <Preview>)' -TestCases @(
		@{ Mode = 'Single'; Preview = $false }, @{ Mode = 'Single'; Preview = $true },
		@{ Mode = 'Csv'; Preview = $false }, @{ Mode = 'Csv'; Preview = $true },
		@{ Mode = 'Application'; Preview = $false }, @{ Mode = 'Application'; Preview = $true }
	) {
		param($Mode, $Preview)
		$fixture.ExistingReviews = @(New-ExistingResourceReview)
		$parameters = switch ($Mode) {
			'Single' { @{ GroupId = $script:groupId; ReviewName = 'New campaign name' } }
			'Csv' {
				$csv = Join-Path $script:testScriptRoot 'resource.csv'
				"GroupId,ReviewName`n$script:groupId,New campaign name" | Set-Content $csv
				@{ CsvPath = $csv }
			}
			'Application' { @{ ServicePrincipalId = $script:reviewId } }
		}
		$result = & $script:testScript @parameters -ReviewerType ResourceOwner -ForceTwoStage -WhatIf:$Preview -Confirm:$false
		$result.Status | Should -Be 'Skipped'
		$result.DuplicateMatchType | Should -Be 'Resource'
		$result.MatchedReviewId | Should -Be $fixture.ExistingReviews[0].Id
		$result.MatchedReviewName | Should -Be 'Unrelated prefix - original campaign'
		$result.Reason | Should -Match 'regardless of reviewer/stage policy'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0
		Should -Invoke Get-MgGroupOwner -Times 0
		if ($Preview) { Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse }
		else { (Get-LogEntries | Where-Object Event -EQ GroupResult).Result.DuplicateMatchType | Should -Be 'Resource' }
	}

	It 'catches a shared group across separate application runs with different prefixes' {
		Mock New-MgIdentityGovernanceAccessReviewDefinition ({
			param($BodyParameter)
			$fixture.Bodies.Add($BodyParameter)
			$fixture.ExistingReviews += [pscustomobject]@{
				Id = $fixture.ReviewId; DisplayName = $BodyParameter.displayName
				Status = 'InProgress'; Scope = $BodyParameter.scope; Settings = $BodyParameter.settings
			}
			[pscustomobject]@{ Id = $fixture.ReviewId }
		}.GetNewClosure())
		$first = & $script:testScript -ServicePrincipalId $script:reviewId -ReviewNamePrefix First -Confirm:$false
		$second = & $script:testScript -ServicePrincipalId 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' -ReviewNamePrefix Second -Confirm:$false
		$first.Status | Should -Be 'Created'
		$second.Status | Should -Be 'Skipped'
		$second.DuplicateMatchType | Should -Be 'Resource'
		$second.MatchedReviewId | Should -Be $first.ReviewId
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 1 -Exactly
	}

	It 'allows renamed historical coverage but keeps future recurring coverage (status: <Status>, kind: <Kind>)' -TestCases @(
		foreach ($status in @('Completed', 'Applied')) {
			@{ Status = $status; Kind = 'completed one-time'; Block = $false }
			@{ Status = $status; Kind = 'expired recurring'; Block = $false }
			@{ Status = $status; Kind = 'future recurring'; Block = $true }
			@{ Status = $status; Kind = 'unbounded recurring'; Block = $true }
			@{ Status = $status; Kind = 'numbered remaining'; Block = $true }
			@{ Status = $status; Kind = 'numbered finished'; Block = $false }
		}
	) {
		param($Status, $Kind, $Block)
		$recurrence = switch ($Kind) {
			'expired recurring' { @{ pattern = @{ type = 'weekly'; interval = 1 }; range = @{ type = 'endDate'; endDate = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd') } } }
			'future recurring' { @{ pattern = @{ type = 'weekly'; interval = 1 }; range = @{ type = 'endDate'; endDate = (Get-Date).AddDays(30).ToString('yyyy-MM-dd') } } }
			'unbounded recurring' { @{ pattern = @{ type = 'weekly'; interval = 1 }; range = @{ type = 'noEnd' } } }
			'numbered remaining' { @{ pattern = @{ type = 'weekly'; interval = 1 }; range = @{ type = 'numbered'; numberOfOccurrences = 2 } } }
			'numbered finished' { @{ pattern = @{ type = 'weekly'; interval = 1 }; range = @{ type = 'numbered'; numberOfOccurrences = 1 } } }
		}
		$fixture.ExistingInstances = @([pscustomobject]@{ Id = 'past'; Status = $Status })
		$fixture.ExistingReviews = @(New-ExistingResourceReview -Status $Status -Recurrence $recurrence)
		$result = & $script:testScript -GroupId $script:groupId -Confirm:$false
		$result.Status | Should -Be $(if ($Block) { 'Skipped' } else { 'Created' })
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times $(if ($Block) { 0 } else { 1 }) -Exactly
	}

	It 'blocks scheduled or applying resource coverage (status: <Status>)' -TestCases @(
		@{ Status = 'NotStarted' }, @{ Status = 'Applying' }
	) {
		param($Status)
		$fixture.ExistingReviews = @(New-ExistingResourceReview -Status $Status)
		$result = & $script:testScript -GroupId $script:groupId -WhatIf
		$result.DuplicateMatchType | Should -Be 'Resource'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0
	}

	It 'keeps numbered recurring coverage while an instance is applying' {
		$recurrence = @{ pattern = @{ type = 'weekly'; interval = 1 }; range = @{ type = 'numbered'; numberOfOccurrences = 1 } }
		$fixture.ExistingReviews = @(New-ExistingResourceReview -Status Applied -Recurrence $recurrence)
		$fixture.ExistingInstances = @([pscustomobject]@{ Id = 'past'; Status = 'Applying' })
		$result = & $script:testScript -GroupId $script:groupId -Confirm:$false
		$result.Status | Should -Be 'Skipped'
		$result.DuplicateMatchType | Should -Be 'Resource'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0
	}

	It 'preserves name protection for applied historical reviews' {
		$existing = New-ExistingResourceReview -Status Applied
		$existing.DisplayName = 'AR Demo group'
		$fixture.ExistingReviews = @($existing)
		$result = & $script:testScript -GroupId $script:groupId -Confirm:$false
		$result.Status | Should -Be 'Skipped'
		$result.DuplicateMatchType | Should -Be 'Name'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0
	}

	It 'does not let a parent review suppress missing nested child reviews' {
		Set-NestedGroupFixture
		$fixture.ExistingReviews = @(New-ExistingResourceReview -Query "/groups/$script:groupId/transitiveMembers")
		$results = @(& $script:testScript -GroupId $script:groupId -Confirm:$false)
		$results[0].Status | Should -Be 'Skipped'
		$results[0].DuplicateMatchType | Should -Be 'Resource'
		$results[1].Status | Should -Be 'Created'
		$results[1].GroupId | Should -Be '44444444-4444-4444-4444-444444444444'
	}

	It 'does not treat partial or enumerated scopes as full coverage (kind: <Kind>)' -TestCases @(
		@{ Kind = 'guests only' }, @{ Kind = 'enumeration' }, @{ Kind = 'restricted principals' }, @{ Kind = 'inactive users' }
	) {
		param($Kind)
		$existing = New-ExistingResourceReview -Query "/groups/$script:groupId/members?`$filter=userType eq 'Guest'"
		if ($Kind -eq 'enumeration') {
			$existing | Add-Member instanceEnumerationScope @{ query = "/groups?`$filter=groupTypes/any(c:c eq 'Unified')" }
			$existing.Scope.query = './members/microsoft.graph.user'
		}
		if ($Kind -eq 'restricted principals') {
			$existing.Scope = @{
				principalScopes = @(@{ query = "/users/$script:ownerId" })
				resourceScopes = @(@{ query = "/groups/$script:groupId" })
			}
		}
		if ($Kind -eq 'inactive users') {
			$existing.Scope = @{ query = "/groups/$script:groupId/members"; inactiveDuration = 'P30D'; '@odata.type' = '#microsoft.graph.accessReviewInactiveUsersQueryScope' }
		}
		$fixture.ExistingReviews = @($existing)
		$result = & $script:testScript -GroupId $script:groupId -Confirm:$false -WarningAction SilentlyContinue
		$result.Status | Should -Be 'Manual review required'
		$result.DuplicateMatchType | Should -Be 'UncertainResourceScope'
		$result.Reason | Should -Match 'rather than assume'
		if ($Kind -eq 'enumeration') {
			$result.MatchedReviewScope | Should -Match 'instanceEnumerationScope: /groups'
		}
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0
	}

	It 'does not match access-package coverage to application groups or direct users (kind: <Kind>)' -TestCases @(
		@{ Kind = 'enumerated packages'; Prefix = '' }
		@{ Kind = 'versioned enumeration'; Prefix = '/beta' }
		@{ Kind = 'absolute enumeration'; Prefix = 'https://graph.microsoft.com/v1.0' }
		@{ Kind = 'nested assignment query'; Prefix = '/v1.0' }
		@{ Kind = 'empty SDK enumeration'; Prefix = '' }
	) {
		param($Kind, $Prefix)
		$fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
		$existing = New-ExistingResourceReview
		if ($Kind -eq 'nested assignment query') {
			$existing.Scope.query = "$Prefix/identityGovernance/entitlementManagement/accessPackages/$script:groupId/assignments"
		}
		elseif ($Kind -eq 'empty SDK enumeration') {
			$existing.Scope.query = "/identityGovernance/entitlementManagement/accessPackageAssignments?`$filter=accessPackageId eq '$script:groupId'"
			$existing | Add-Member instanceEnumerationScope ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphAccessReviewScope]::new())
		}
		else {
			$existing.Scope.query = './assignments'
			$enumeration = [Microsoft.Graph.PowerShell.Models.MicrosoftGraphAccessReviewScope]::new()
			$enumeration.AdditionalProperties['query'] = "$Prefix/identityGovernance/entitlementManagement/accessPackages/$script:groupId"
			$existing | Add-Member instanceEnumerationScope $enumeration
		}
		$fixture.ExistingReviews = @($existing)
		$results = @(& $script:testScript -ServicePrincipalId $script:reviewId -Confirm:$false)
		$results.TargetType | Should -Be @('GroupMembership', 'ApplicationDirectUsers')
		$results.Status | Should -Be @('Created', 'Created')
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 2 -Exactly
	}

	It 'limits unresolved enumeration coverage to its resource family (kind: <Kind>)' -TestCases @(
		@{ Kind = 'groups'; Query = '/groups'; GroupStatus = 'Manual review required'; AppStatus = 'Created' }
		@{ Kind = 'applications'; Query = '/v1.0/servicePrincipals'; GroupStatus = 'Created'; AppStatus = 'Manual review required' }
		@{ Kind = 'unknown'; Query = '/unsupportedResources'; GroupStatus = 'Manual review required'; AppStatus = 'Manual review required' }
	) {
		param($Kind, $Query, $GroupStatus, $AppStatus)
		$fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
		$existing = New-ExistingResourceReview
		$existing.Scope.query = './members'
		$existing | Add-Member instanceEnumerationScope @{ query = $Query }
		$fixture.ExistingReviews = @($existing)
		$results = @(& $script:testScript -ServicePrincipalId $script:reviewId -Confirm:$false -WarningAction SilentlyContinue)
		$results.Status | Should -Be @($GroupStatus, $AppStatus)
		$manual = @($results | Where-Object Status -EQ 'Manual review required')
		$manual[0].MatchedReviewScope | Should -Be "instanceEnumerationScope: $Query"
		$created = @($results | Where-Object Status -EQ 'Created').Count
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times $created -Exactly
	}

	It 'retains a real group resource match in a mixed group and access-package scope' {
		$existing = New-ExistingResourceReview
		$existing.Scope = @{
			principalScopes = @(@{ query = '/users' })
			resourceScopes = @(
				@{ query = '/identityGovernance/entitlementManagement/accessPackages/aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa/assignments' }
				@{ query = "/groups/$script:groupId/members" }
			)
		}
		$fixture.ExistingReviews = @($existing)
		$result = & $script:testScript -GroupId $script:groupId -Confirm:$false
		$result.DuplicateMatchType | Should -Be 'Resource'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'limits a specific enumerated group to its own ID (same: <Same>, filtered: <Filtered>)' -TestCases @(
		@{ Same = $false; Filtered = $false }, @{ Same = $true; Filtered = $false }
		@{ Same = $false; Filtered = $true }, @{ Same = $true; Filtered = $true }
	) {
		param($Same, $Filtered)
		$id = if ($Same) { $script:groupId } else { 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' }
		$existing = New-ExistingResourceReview -Query './members/microsoft.graph.user'
		if ($Filtered) { $existing.Scope.query += "?`$filter=userType eq 'Guest'" }
		$existing | Add-Member instanceEnumerationScope @{ query = "/v1.0/groups/$id" }
		$fixture.ExistingReviews = @($existing)
		$result = & $script:testScript -GroupId $script:groupId -Confirm:$false -WarningAction SilentlyContinue
		$result.Status | Should -Be $(if (-not $Same) { 'Created' } elseif ($Filtered) { 'Manual review required' } else { 'Skipped' })
		if ($Same) {
			$result.MatchedReviewScope | Should -Match "/v1.0/groups/$id"
			$result.DuplicateMatchType | Should -Be $(if ($Filtered) { 'UncertainResourceScope' } else { 'Resource' })
		}
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times $(if ($Same) { 0 } else { 1 }) -Exactly
	}

	It 'creates CSV direct-user reviews despite unrelated exact enumeration and empty SDK scopes (preview: <Preview>, existing app: <Same>)' -TestCases @(
		@{ Preview = $false; Same = $false }, @{ Preview = $true; Same = $false }
		@{ Preview = $false; Same = $true }, @{ Preview = $true; Same = $true }
	) {
		param($Preview, $Same)
		$first = $script:reviewId
		$second = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
		Mock Get-MgServicePrincipal {
			param($ServicePrincipalId)
			[pscustomobject]@{ Id = "$ServicePrincipalId"; DisplayName = if ("$ServicePrincipalId" -eq '33333333-3333-3333-3333-333333333333') { 'First app' } else { 'Second app' } }
		}
		$fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId; Id = 'direct-role-one' }
		$fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:groupOwnerId; Id = 'direct-role-two' }
		$groupReview = New-ExistingResourceReview -Query './members'
		$groupReview | Add-Member instanceEnumerationScope @{ query = '/v1.0/groups/aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa' }
		$appId = if ($Same) { $first } else { 'cccccccc-cccc-cccc-cccc-cccccccccccc' }
		$appReview = New-ExistingResourceReview
		$appReview.Id = 'dddddddd-dddd-dddd-dddd-dddddddddddd'
		$appReview.DisplayName = 'Unrelated application campaign'
		$appReview.Scope = @{
			principalScopes = @(@{ query = '/v1.0/users' })
			resourceScopes = @(@{ query = "/v1.0/servicePrincipals/$appId" })
		}
		$appReview | Add-Member instanceEnumerationScope ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphAccessReviewScope]::new())
		$fixture.ExistingReviews = @($groupReview, $appReview)
		$csv = Join-Path $script:testScriptRoot 'applications.csv'
		"ServicePrincipalId`n$first`n$second" | Set-Content $csv
		$path = Join-Path $script:testScriptRoot 'applications.html'
		$results = @(& $script:testScript -CsvPath $csv -WhatIf:$Preview -HtmlReportPath $path -Confirm:$false)
		$direct = @($results | Where-Object TargetType -EQ 'ApplicationDirectUsers')
		$created = if ($Preview) { 'Would be created' } else { 'Created' }
		$direct | Should -HaveCount 2
		$direct.SubjectCount | Should -Be @(2, 2)
		$direct.Status | Should -Be @($(if ($Same) { 'Skipped' } else { $created }), $created)
		if ($Same) { $direct[0].DuplicateMatchType | Should -Be 'Resource' }
		@($results | Where-Object TargetType -EQ 'GroupMembership').Status | Should -Be $created
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times $(if ($Preview) { 0 } elseif ($Same) { 2 } else { 3 }) -Exactly
		$html = Get-Content $path -Raw
		[regex]::Matches($html, 'class="topology-node application-node"').Count | Should -Be 2
		[regex]::Matches($html, 'class="topology-edge application-edge"').Count | Should -Be 2
		$html | Should -Match '2 directly assigned users'
	}

	It 'writes the requested HTML before emitting the full result batch' {
		$fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
		$fixture.ApplicationOwners = @(
			[pscustomobject]@{ Id = $script:ownerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
		)
		$path = Join-Path $script:testScriptRoot 'batch.html'
		$observed = [System.Collections.Generic.List[int]]::new()
		$results = @(& $script:testScript -ServicePrincipalId $script:reviewId -HtmlReportPath $path -Confirm:$false -ErrorAction Stop -WarningVariable warnings |
			ForEach-Object {
				$path | Should -Exist
				$observed.Add($fixture.Bodies.Count)
				$_
			})
		$results | Should -HaveCount 3
		@($observed | Sort-Object -Unique) | Should -Be @(2)
		@($warnings).Count | Should -Be 0
		($results | ConvertTo-Csv -NoTypeInformation)[0] | Should -Match 'ReportSchemaVersion'
	}

	It 'groups uncertain duplicate warnings and distinguishes coverage from ownership follow-up' {
		Set-NestedGroupFixture
		$fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
		$fixture.ApplicationOwners = @(
			[pscustomobject]@{ Id = $script:ownerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
		)
		$existing = New-ExistingResourceReview -Query './members'
		$existing | Add-Member instanceEnumerationScope @{ query = "/groups?`$filter=groupTypes/any(c:c eq 'Unified')" }
		$fixture.ExistingReviews = @($existing)
		$results = @(& $script:testScript -ServicePrincipalId $script:reviewId -Confirm:$false -WarningAction SilentlyContinue -WarningVariable warnings)
		$results | Should -HaveCount 4
		@($warnings) | Should -HaveCount 1
		$warnings[0].Message | Should -Match '^2 target\(s\) need manual duplicate-coverage assessment'
		$fixture.HostMessages | Should -Contain '3 review target(s): 1 created, 0 skipped, 0 failed. 2 require manual duplicate assessment.'
		$fixture.HostMessages | Should -Contain 'Manual review required: 1 ownership/assignment item(s). These are not covered by the created reviews.'
		@($results | Where-Object TargetType -EQ 'ApplicationDirectUsers').Status | Should -Be 'Created'
	}

	It 'reads SDK AdditionalProperties and versioned group scopes' {
		$existing = New-ExistingResourceReview
		$existing.Scope = [Microsoft.Graph.PowerShell.Models.MicrosoftGraphAccessReviewScope]::new()
		$existing.Scope.AdditionalProperties['query'] = "/v1.0/groups/$script:groupId/members/microsoft.graph.user"
		$fixture.ExistingReviews = @($existing)
		$result = & $script:testScript -GroupId $script:groupId -Confirm:$false
		$result.DuplicateMatchType | Should -Be 'Resource'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0
	}

	It 'does not confuse different groups with the same users or application access with membership' {
		$fixture.ExistingReviews = @(
			(New-ExistingResourceReview -Query '/groups/aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa/members'),
			(New-ExistingResourceReview -Query "/servicePrincipals/$script:groupId/appRoleAssignedTo")
		)
		$result = & $script:testScript -GroupId $script:groupId -Confirm:$false
		$result.Status | Should -Be 'Created'
	}

	It 'protects direct application-user reviews independently of their names' {
		$fixture.Assignments = @([pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId })
		$existing = New-ExistingResourceReview
		$existing.Scope = @{
			principalScopes = @(@{ query = '/v1.0/users' })
			resourceScopes = @(@{ query = "/v1.0/servicePrincipals/$script:groupId" })
		}
		$fixture.ExistingReviews = @($existing)
		$result = & $script:testScript -ServicePrincipalId $script:groupId -Confirm:$false
		$result.TargetType | Should -Be 'ApplicationDirectUsers'
		$result.DuplicateMatchType | Should -Be 'Resource'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0
	}

	It 'fails closed before any writes when existing coverage cannot be read (kind: <Kind>)' -TestCases @(
		@{ Kind = 'scope missing' }, @{ Kind = 'status missing' }, @{ Kind = 'status unknown' }, @{ Kind = 'numbered instances denied' }
	) {
		param($Kind)
		$existing = New-ExistingResourceReview
		switch ($Kind) {
			'scope missing' { $existing.Scope = $null }
			'status missing' { $existing.Status = $null }
			'status unknown' { $existing.Status = 'UnexpectedStatus' }
			'numbered instances denied' {
				$existing.Status = 'Completed'
				$existing.Settings.recurrence = @{ pattern = @{ type = 'weekly' }; range = @{ type = 'numbered'; numberOfOccurrences = 2 } }
				Mock Get-MgIdentityGovernanceAccessReviewDefinitionInstance { throw 'Instances forbidden' }
			}
		}
		$fixture.ExistingReviews = @($existing)
		{ & $script:testScript -GroupId $script:groupId -Confirm:$false } | Should -Throw '*resource coverage; no reviews were created*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0
		(Get-LogEntries)[-1].Event | Should -Be 'RunFailed'
	}

	It 'skips duplicate names within a CSV batch (preview: <Preview>)' -TestCases @(
		@{ Preview = $false }, @{ Preview = $true }
	) {
		param($Preview)
		$csv = Join-Path $script:testScriptRoot 'same-name.csv'
		"GroupId`n$script:groupId`n44444444-4444-4444-4444-444444444444" | Set-Content $csv
		$results = @(& $script:testScript -CsvPath $csv -WhatIf:$Preview -Confirm:$false)
		$results[0].Status | Should -Be $(if ($Preview) { 'Would be created' } else { 'Created' })
		$results[1].Status | Should -Be 'Skipped'
		$results[1].Reason | Should -Match 'already exists'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times $(if ($Preview) { 0 } else { 1 }) -Exactly
	}

	It 'does not create another review when a <Mode> run is repeated' -TestCases @(
		@{ Mode = 'Single' }, @{ Mode = 'Csv' }, @{ Mode = 'Application' }
	) {
		param($Mode)
		$fixture = $script:fixture
		Mock New-MgIdentityGovernanceAccessReviewDefinition ({
			param($BodyParameter)
			$fixture.ExistingNames += $BodyParameter.displayName
			[pscustomobject]@{ Id = $fixture.ReviewId }
		}.GetNewClosure())
		$parameters = switch ($Mode) {
			'Single' { @{ GroupId = $script:groupId } }
			'Csv' {
				$csv = Join-Path $script:testScriptRoot 'rerun.csv'
				"GroupId`n$script:groupId" | Set-Content $csv
				@{ CsvPath = $csv }
			}
			'Application' { @{ ServicePrincipalId = $script:groupId } }
		}
		$first = & $script:testScript @parameters -Confirm:$false -WarningAction SilentlyContinue
		$second = & $script:testScript @parameters -Confirm:$false -WarningAction SilentlyContinue
		$first.Status | Should -Be 'Created'
		$second.Status | Should -Be 'Skipped'
		$second.Reason | Should -Match 'already exists'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 1 -Exactly
	}

	It 'stops without creating reviews if the existing review lookup fails (mode: <Mode>)' -TestCases @(
		@{ Mode = 'Single' }, @{ Mode = 'Csv' }, @{ Mode = 'Application' }
	) {
		param($Mode)
		Mock Get-MgIdentityGovernanceAccessReviewDefinition { throw 'Review lookup forbidden.' }
		$parameters = switch ($Mode) {
			'Single' { @{ GroupId = $script:groupId } }
			'Csv' {
				$csv = Join-Path $script:testScriptRoot 'lookup-failure.csv'
				"GroupId`n$script:groupId" | Set-Content $csv
				@{ CsvPath = $csv }
			}
			'Application' { @{ ServicePrincipalId = $script:groupId } }
		}
		{ & $script:testScript @parameters -Confirm:$false } | Should -Throw '*Unable to check existing access review names*'
		(Get-LogEntries)[-1].Event | Should -Be 'RunFailed'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'stops all creation if any application group membership discovery fails' {
		$second = '44444444-4444-4444-4444-444444444444'
		$third = '55555555-5555-5555-5555-555555555555'
		$script:fixture.Assignments += @(
			[pscustomobject]@{ PrincipalType = 'Group'; PrincipalId = $second },
			[pscustomobject]@{ PrincipalType = 'Group'; PrincipalId = $third }
		)
		Mock Get-MgGroup {
			[pscustomobject]@{ DisplayName = 'Dynamic group'; GroupTypes = @('DynamicMembership') }
		} -ParameterFilter { "$GroupId" -eq '44444444-4444-4444-4444-444444444444' }
		Mock Get-MgGroupMember { throw 'Count unavailable.' } -ParameterFilter {
			"$GroupId" -eq '55555555-5555-5555-5555-555555555555'
		}
		{ & $script:testScript -ServicePrincipalId $script:groupId -Confirm:$false } |
			Should -Throw '*Group discovery failed; no reviews were created*Count unavailable*'
		(Get-LogEntries)[-1].Event | Should -Be 'RunFailed'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'excludes groups whose review creation failed after retrieving a count' {
		$script:fixture.CreateFailure = $true
		$csv = Join-Path $script:testScriptRoot 'failed-counts.csv'
		"GroupId`n$script:groupId" | Set-Content $csv
		$result = & $script:testScript -CsvPath $csv -Confirm:$false
		$result.Status | Should -Be 'Failed'
		$result.MemberCount | Should -Be 150
		$script:fixture.HostMessages[-1] | Should -Be 'Total direct users in group created reviews: 0'
	}

	It 'excludes a declined single-group review from the member total' {
		# Simulate a declined prompt only in the isolated test copy.
		$source = Get-Content $script:testScript -Raw
		$ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
		$prompt = $ast.Find({
			param($node)
			$node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
				$node.Member.Value -eq 'ShouldProcess'
		}, $true).Extent
		$source.Remove($prompt.StartOffset, $prompt.EndOffset - $prompt.StartOffset).Insert($prompt.StartOffset, '$false') |
			Set-Content $script:testScript
		$result = & $script:testScript -GroupId $script:groupId -Confirm:$false
		$result.Status | Should -Be 'Not created'
		$result.MemberCount | Should -Be 150
		$script:fixture.HostMessages[-1] | Should -Be 'Total direct users in group created reviews: 0'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'skips empty groups even with owners in <Mode> mode (preview: <Preview>)' -TestCases @(
		@{ Mode = 'Single'; Preview = $false },
		@{ Mode = 'Single'; Preview = $true },
		@{ Mode = 'Csv'; Preview = $false },
		@{ Mode = 'Csv'; Preview = $true },
		@{ Mode = 'Application'; Preview = $false },
		@{ Mode = 'Application'; Preview = $true }
	) {
		param($Mode, $Preview)
		$script:fixture.Count = 0
		$parameters = switch ($Mode) {
			'Single' { @{ GroupId = $script:groupId } }
			'Csv' {
				$csv = Join-Path $script:testScriptRoot 'empty-group.csv'
				"GroupId`n$script:groupId" | Set-Content $csv
				@{ CsvPath = $csv }
			}
			'Application' { @{ ServicePrincipalId = $script:groupId } }
		}
		$result = & $script:testScript @parameters -ReviewerType ResourceOwner -WhatIf:$Preview -Confirm:$false -WarningAction SilentlyContinue
		$result.Status | Should -Be 'Skipped'
		$result.Reason | Should -Be 'No direct users; no access review needed'
		$result.MemberCount | Should -Be 0
		$result.Stages | Should -BeNullOrEmpty
		$script:fixture.HostMessages | Should -Contain 'The group Demo group has no direct users; no access review needed.'
		if (-not $Preview) {
			((Get-LogEntries) | Where-Object Event -EQ 'GroupResult').Result.Status | Should -Be 'Skipped'
			(Get-LogEntries)[-1].Skipped | Should -Be 1
		}
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'applies stage thresholds to <Count> members in <Mode> mode' -TestCases @(
		@{ Mode = 'Single'; Count = 10; Stages = 1 },
		@{ Mode = 'Single'; Count = 149; Stages = 1 },
		@{ Mode = 'Single'; Count = 150; Stages = 1 },
		@{ Mode = 'Single'; Count = 151; Stages = 2 },
		@{ Mode = 'Single'; Count = 499; Stages = 2 },
		@{ Mode = 'Single'; Count = 500; Stages = 3 },
		@{ Mode = 'Application'; Count = 10; Stages = 1 },
		@{ Mode = 'Application'; Count = 149; Stages = 1 },
		@{ Mode = 'Application'; Count = 150; Stages = 1 },
		@{ Mode = 'Application'; Count = 151; Stages = 2 },
		@{ Mode = 'Application'; Count = 499; Stages = 2 },
		@{ Mode = 'Application'; Count = 500; Stages = 3 }
	) {
		param($Mode, $Count, $Stages)
		$script:fixture.Count = $Count
		$parameters = @{ GroupId = $script:groupId }
		if ($Mode -eq 'Application') {
			$parameters = @{ ServicePrincipalId = $script:groupId }
			$script:fixture.Assignments = @(1..$Count | ForEach-Object {
				[pscustomobject]@{ PrincipalType = 'User'; PrincipalId = ([guid]::NewGuid().ToString()) }
			})
		}
		$result = & $script:testScript @parameters -Confirm:$false
		$result.Stages | Should -Be $Stages
		$body = $script:fixture.Bodies[0]
		$body.settings.instanceDurationInDays | Should -Be ($Stages * 5)
		if ($Stages -eq 1) {
			$body.ContainsKey('stageSettings') | Should -BeFalse
			$body.reviewers[0].query | Should -Be "/users/$script:ownerId"
		}
		else {
			$body.stageSettings | Should -HaveCount $Stages
			$body.stageSettings[-1].reviewers[0].query | Should -Be "/users/$script:ownerId"
		}
	}

	It 'uses per-row reviewer selection and role text without changing group scope at <Count> members' -TestCases @(
		@{ Count = 150 }, @{ Count = 500 }
	) {
		param($Count)
		$script:fixture.Count = $Count
		$second = '44444444-4444-4444-4444-444444444444'
		$csv = Join-Path $script:testScriptRoot 'roles.csv'
		"GroupId,ReviewName,ReviewerType,Role`n$script:groupId,First,ResourceOwner,Finance approver`n$second,Second,,Help desk" |
			Set-Content $csv
		$results = @(& $script:testScript -CsvPath $csv -ReviewerType ApplicationOwner -Confirm:$false)
		$results.Status | Should -Be @('Created', 'Created')
		$results.ReviewerType | Should -Be @('ResourceOwner', 'ApplicationOwner')
		if ($Count -ge 500) {
			$results.Stages | Should -Be @(3, 3)
			$script:fixture.Bodies[0].stageSettings[-1].reviewers.query | Should -Be @("/users/$script:groupOwnerId")
			$script:fixture.Bodies[1].stageSettings[-1].reviewers.query | Should -Be @("/users/$script:ownerId")
		}
		else {
			$script:fixture.Bodies[0].reviewers.query | Should -Be @("/users/$script:groupOwnerId")
			$script:fixture.Bodies[1].reviewers.query | Should -Be @("/users/$script:ownerId")
		}
		$script:fixture.Bodies[0].descriptionForReviewers | Should -Match 'Role: Finance approver\.'
		$script:fixture.Bodies[1].descriptionForReviewers | Should -Match 'Role: Help desk\.'
		$script:fixture.Bodies[0].scope.AdditionalProperties.query | Should -Be "/groups/$script:groupId/members"
	}

	It 'rejects an oversized role message without creating a review' {
		$csv = Join-Path $script:testScriptRoot 'long-role.csv'
		"GroupId,Role`n$script:groupId,$('x' * 256)" | Set-Content $csv
		$result = & $script:testScript -CsvPath $csv -Confirm:$false
		$result.Status | Should -Be 'Failed'
		$result.Reason | Should -Match 'Shorten ReviewerMessage or Role'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'summarizes long membership context within 256 characters without losing explicit text or full paths' {
		Set-NestedGroupFixture
		$parentName = 'Department' * 8
		$applicationName = 'Payroll' * 20
		$role = 'Approver' * 8
		(Get-Content -LiteralPath $script:testScript -Raw).Replace(
			'$ReviewerMessage = ''''', '$ReviewerMessage = ''Contact IAM.'''
		) | Set-Content -LiteralPath $script:testScript
		$script:fixture.Groups[$script:groupId].DisplayName = $parentName
		Mock Get-MgServicePrincipal ({ [pscustomobject]@{ DisplayName = $applicationName } }.GetNewClosure())
		$csv = Join-Path $script:testScriptRoot 'long-context.csv'
		"ServicePrincipalId,Role`n$script:groupId,$role" | Set-Content $csv
		$warnings = @()
		$results = @(& $script:testScript -CsvPath $csv -Confirm:$false -WarningAction SilentlyContinue -WarningVariable warnings)
		$results.Status | Should -Be @('Created', 'Created')
		foreach ($result in $results) {
			$result.ReviewConfiguration.DescriptionForReviewers.Length | Should -BeLessOrEqual 256
			$result.ReviewConfiguration.DescriptionForReviewers | Should -Match "Role: $role\."
			$result.ReviewConfiguration.DescriptionForReviewers | Should -Match 'Contact IAM\.'
		}
		$results[1].ReviewConfiguration.DescriptionForReviewers | Should -Match 'Parent groups: 1\. Related app assignments: 1\.'
		$results[1].ReviewConfiguration.DescriptionForAdmins | Should -Match "$parentName -> Group 2"
		$results[1].ReviewConfiguration.DescriptionForAdmins | Should -Match $applicationName
		@($warnings | Where-Object { "$_" -like '*relationship summary*Group 2*' }) | Should -HaveCount 1
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 2 -Exactly
	}

	It 'honors the 256-character membership context boundary without dropping Role (length: <RoleLength>)' -TestCases @(
		@{ RoleLength = 164; Fits = $true },
		@{ RoleLength = 170; Fits = $false }
	) {
		param($RoleLength, $Fits)
		Set-NestedGroupFixture
		$csv = Join-Path $script:testScriptRoot 'context-cannot-fit.csv'
		"GroupId,Role`n$script:groupId,$('x' * $RoleLength)" | Set-Content $csv
		$results = @(& $script:testScript -CsvPath $csv -Confirm:$false -WarningAction SilentlyContinue)
		if ($Fits) {
			$results.Status | Should -Be @('Created', 'Created')
			$results[1].ReviewConfiguration.DescriptionForReviewers.Length | Should -Be 256
			$results[1].ReviewConfiguration.DescriptionForReviewers | Should -Match "Role: $('x' * $RoleLength)\."
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 2 -Exactly
		}
		else {
			$results.Status | Should -Be @('Created', 'Failed')
			$results[1].Reason | Should -Match 'cannot fit membership context within 256 characters.*Shorten ReviewerMessage or Role'
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 1 -Exactly
		}
	}

	It 'accepts mixed group and application CSV rows without leaking their role settings' {
		$second = '44444444-4444-4444-4444-444444444444'
		$csv = Join-Path $script:testScriptRoot 'mixed.csv'
		"GroupId,ServicePrincipalId,Role,ReviewName`n$script:groupId,,Group role,Group review`n,$second,Application role," | Set-Content $csv
		$script:fixture.Assignments = @([pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId })
		$results = @(& $script:testScript -CsvPath $csv -Confirm:$false)
		$results.Status | Should -Be @('Created', 'Created')
		$results.ReviewName | Should -Be @('Group review', 'AR Demo app')
		$results.Role | Should -Be @('Group role', 'Application role')
		$script:fixture.Bodies[0].scope.AdditionalProperties.query | Should -Be "/groups/$script:groupId/members"
		$script:fixture.Bodies[1].scope.AdditionalProperties.resourceScopes[0].query | Should -Be "/v1.0/servicePrincipals/$second"
	}

	It 'rejects a CSV row with both target IDs' {
		$csv = Join-Path $script:testScriptRoot 'ambiguous.csv'
		"GroupId,ServicePrincipalId`n$script:groupId,$script:reviewId" | Set-Content $csv
		{ & $script:testScript -CsvPath $csv -Confirm:$false } | Should -Throw '*exactly one GroupId or ServicePrincipalId*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'discovers multiple CSV applications with isolated role, owner and scope settings (preview: <Preview>)' -TestCases @(
		@{ Preview = $false }, @{ Preview = $true }
	) {
		param($Preview)
		$second = '44444444-4444-4444-4444-444444444444'
		$csv = Join-Path $script:testScriptRoot 'applications.csv'
		"ServicePrincipalId,AppOwnerIds,ReviewerType,Role`n$script:groupId,$script:ownerId,ResourceOwner,Approver`n$second,$script:reviewId,ApplicationOwner,Reader" |
			Set-Content $csv
		Mock Get-MgServicePrincipal {
			param($ServicePrincipalId)
			[pscustomobject]@{ DisplayName = if ("$ServicePrincipalId" -eq '11111111-1111-1111-1111-111111111111') { 'First app' } else { 'Second app' } }
		}
		Mock Get-MgServicePrincipalOwner {
			[pscustomobject]@{ Id = '66666666-6666-6666-6666-666666666666'; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
		} -ParameterFilter { "$ServicePrincipalId" -eq '11111111-1111-1111-1111-111111111111' }
		Mock Get-MgServicePrincipalOwner { } -ParameterFilter { "$ServicePrincipalId" -eq '44444444-4444-4444-4444-444444444444' }
		Mock Get-MgServicePrincipalAppRoleAssignedTo {
			[pscustomobject]@{ PrincipalType = 'Group'; PrincipalId = '55555555-5555-5555-5555-555555555555' }
			[pscustomobject]@{ PrincipalType = 'User'; PrincipalId = '22222222-2222-2222-2222-222222222222'; Id = 'first-assignment' }
		} -ParameterFilter { "$ServicePrincipalId" -eq '11111111-1111-1111-1111-111111111111' }
		Mock Get-MgServicePrincipalAppRoleAssignedTo {
			[pscustomobject]@{ PrincipalType = 'User'; PrincipalId = '33333333-3333-3333-3333-333333333333'; Id = 'second-assignment' }
		} -ParameterFilter { "$ServicePrincipalId" -eq '44444444-4444-4444-4444-444444444444' }
		$results = @(& $script:testScript -CsvPath $csv -WhatIf:$Preview -Confirm:$false -WarningAction SilentlyContinue)
		$reviews = @($results | Where-Object TargetType -in @('GroupMembership', 'ApplicationDirectUsers'))
		$reviews | Should -HaveCount 3
		$reviews.Status | Should -Be @($(if ($Preview) { 'Would be created' } else { 'Created' }), $(if ($Preview) { 'Would be created' } else { 'Created' }), $(if ($Preview) { 'Would be created' } else { 'Created' }))
		$reviews.Role | Should -Be @('Approver', 'Approver', 'Reader')
		$reviews.ServicePrincipalId | Should -Be @($script:groupId, $script:groupId, $second)
		$reviews[1].PrincipalIds | Should -Be @($script:ownerId)
		$reviews[2].PrincipalIds | Should -Be @($script:reviewId)
		if ($Preview) {
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
			Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse
		}
		else {
			$script:fixture.Bodies[0].descriptionForReviewers | Should -Match 'Role: Approver\.'
			$script:fixture.Bodies[1].descriptionForReviewers | Should -Match 'Role: Approver\.'
			$script:fixture.Bodies[2].descriptionForReviewers | Should -Match 'Role: Reader\.'
			$script:fixture.Bodies[1].reviewers.query | Should -Be @("/users/$script:groupOwnerId")
			$script:fixture.Bodies[2].reviewers.query | Should -Be @("/users/$script:reviewId")
			$script:fixture.Bodies[1].scope.AdditionalProperties.resourceScopes[0].query | Should -Be "/v1.0/servicePrincipals/$script:groupId"
			$script:fixture.Bodies[2].scope.AdditionalProperties.resourceScopes[0].query | Should -Be "/v1.0/servicePrincipals/$second"
		}
	}

	It 'stops all CSV application creation if a later discovery fails' {
		$second = '44444444-4444-4444-4444-444444444444'
		$csv = Join-Path $script:testScriptRoot 'discovery-error.csv'
		"ServicePrincipalId`n$script:groupId`n$second" | Set-Content $csv
		Mock Get-MgServicePrincipal { throw 'Application read forbidden.' } -ParameterFilter { "$ServicePrincipalId" -eq '44444444-4444-4444-4444-444444444444' }
		{ & $script:testScript -CsvPath $csv -Confirm:$false } | Should -Throw '*Application read forbidden*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'rejects conflicting duplicate CSV target settings before creating reviews' {
		$csv = Join-Path $script:testScriptRoot 'conflicting.csv'
		"GroupId,Role`n$script:groupId,Reader`n$script:groupId,Writer" | Set-Content $csv
		{ & $script:testScript -CsvPath $csv -Confirm:$false } | Should -Throw '*conflicting settings*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'requires application read scope for service-principal CSV rows' {
		$csv = Join-Path $script:testScriptRoot 'application-scope.csv'
		"ServicePrincipalId`n$script:groupId" | Set-Content $csv
		Mock Get-MgContext { @{ Scopes = @('AccessReview.ReadWrite.All') } }
		{ & $script:testScript -CsvPath $csv -WhatIf } | Should -Throw '*Application.Read.All*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'skips an existing review discovered from a service-principal CSV row' {
		$csv = Join-Path $script:testScriptRoot 'application-existing.csv'
		"ServicePrincipalId`n$script:groupId" | Set-Content $csv
		$script:fixture.Assignments = @([pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId })
		$script:fixture.ExistingNames = @('AR Demo app')
		$result = & $script:testScript -CsvPath $csv -Confirm:$false
		$result.Status | Should -Be 'Skipped'
		$result.Reason | Should -Match 'already exists'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'uses per-row CSV owners and falls back to the parameter for blank cells' {
		$csv = Join-Path $script:testScriptRoot 'owners.csv'
		$second = '44444444-4444-4444-4444-444444444444'
		$third = '55555555-5555-5555-5555-555555555555'
		"GroupId,AppOwnerIds,ReviewName`n$script:groupId,$second;$third;$second,First`n$second,,Second" | Set-Content $csv
		$null = & $script:testScript -CsvPath $csv -ReviewerType ApplicationOwner -Confirm:$false
		$script:fixture.Bodies[0].reviewers.query | Should -Be @("/users/$second", "/users/$third")
		$script:fixture.Bodies[1].reviewers[0].query | Should -Be "/users/$script:ownerId"
		$entries = Get-LogEntries
		$entries[1].Result.AppOwnerIds | Should -Be @($second, $third)
	}

	It 'rejects missing application owners without creating a review' {
		{ & $script:testScript -GroupId $script:groupId -AppOwnerIds @() } | Should -Throw '*sign-off is required*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'supports CSV-only owners without a parameter and arbitrary column order' {
		$csv = Join-Path $script:testScriptRoot 'owners.csv'
		"AppOwnerIds,GroupId`n$script:ownerId,$script:groupId" | Set-Content $csv
		$null = & $script:testScript -CsvPath $csv -AppOwnerIds @() -ReviewerType ApplicationOwner -Confirm:$false
		$script:fixture.Bodies[0].reviewers[0].query | Should -Be "/users/$script:ownerId"
	}

	It 'rejects the empty GUID parameter before creating any reviews' {
		{ & $script:testScript -GroupId $script:groupId -AppOwnerIds ([guid]::Empty) } | Should -Throw '*nonempty user object GUIDs*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'fails malformed CSV owners instead of silently using the default' {
		$csv = Join-Path $script:testScriptRoot 'owners.csv'
		"GroupId,AppOwnerIds`n$script:groupId,not-a-guid" | Set-Content $csv
		$null = & $script:testScript -CsvPath $csv -Confirm:$false
		(Get-LogEntries)[1].Result.Reason | Should -Match 'invalid user object GUID'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'logs replayable array parameter syntax' {
		$ids = @($script:ownerId, $script:reviewId)
		$null = & $script:testScript -GroupId $script:groupId -AppOwnerIds $ids -Confirm:$false
		(Get-LogEntries)[0].Syntax | Should -Match "-AppOwnerIds @\('$script:ownerId', '$script:reviewId'\)"
	}

	It 'uses per-row application owners as manager fallback instead of the zero GUID' {
		$script:fixture.Count = 151
		$csv = Join-Path $script:testScriptRoot 'fallback-owners.csv'
		"GroupId;AppOwnerIds`n$script:groupId;$script:ownerId,$script:reviewId" | Set-Content -LiteralPath $csv
		$null = & $script:testScript -CsvPath $csv -ReviewerType LineManager -Confirm:$false
		$script:fixture.Bodies[0].stageSettings[0].fallbackReviewers.query | Should -Be @("/users/$script:ownerId", "/users/$script:reviewId")
	}

	It 'rejects an explicitly empty fallback before creating a review' {
		{ & $script:testScript -GroupId $script:groupId -FallbackReviewerUserId ([guid]::Empty) } | Should -Throw '*FallbackReviewerUserId must be a nonempty*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'logs an owner lookup failure before rethrowing a single-group error' {
		$script:fixture.OwnerFailure = $true
		{ & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner } | Should -Throw '*Owner read forbidden*'
		$entries = Get-LogEntries
		$entries.Count | Should -Be 2
		$entries[1].Event | Should -Be 'GroupResult'
		$entries[1].Result.Status | Should -Be 'Failed'
		$entries[1].Result.Reason | Should -Match 'Owner read forbidden'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'logs real-run syntax with escaped quotes and created review IDs' {
		$result = & $script:testScript -GroupId $script:groupId -ReviewName "Review O'Brien" -ReviewerType ResourceOwner -Confirm:$false
		$result.Status | Should -Be 'Created'
		$entries = Get-LogEntries
		$entries.Count | Should -Be 3
		$entries[0].Syntax | Should -Match "-ReviewName 'Review O''Brien'"
		$entries[0].Syntax | Should -Match "-ReviewerType 'ResourceOwner'"
		$entries[1].Result.ReviewId | Should -Be $script:reviewId
		$entries[1].Result.Status | Should -Be 'Created'
		$entries[2].Previewed | Should -Be 0
		$entries[2].Created | Should -Be 1
	}

	It 'writes independent run files with timestamps, invocation and created review IDs' {
		$null = & $script:testScript -GroupId $script:groupId -Confirm:$false
		$firstFile = @(Get-RunLogs)[0]
		$firstContents = Get-Content -LiteralPath $firstFile.FullName -Raw
		$null = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false
		$files = @(Get-RunLogs)
		$files | Should -HaveCount 2
		(Get-Content -LiteralPath $firstFile.FullName -Raw) | Should -BeExactly $firstContents
		$runIds = @(foreach ($file in $files) {
			$entries = Get-LogEntries -Path $file.FullName
			$entries | Should -HaveCount 3
			$entries.Event | Should -Be @('RunStarted', 'GroupResult', 'RunCompleted')
			@($entries.RunId | Select-Object -Unique) | Should -HaveCount 1
			$file.Name | Should -Match "^Create-AccessReview_\d{8}T\d{9}Z_$($entries[0].RunId)\.jsonl$"
			([datetime]$entries[0].Timestamp).Kind | Should -Be ([System.DateTimeKind]::Utc)
			$entries[1].Result.ReviewId | Should -Be $script:reviewId
			$entries[2].Created | Should -Be 1
			$entries[0].Syntax | Should -Match '-Confirm:\$false'
			$script:fixture.HostMessages | Should -Contain "Creation log: $($file.FullName)"
			$entries[0].RunId
		})
		@($runIds | Select-Object -Unique) | Should -HaveCount 2
	}

	It 'does not create logs for <Mode> previews (inherited WhatIf: <Inherited>)' -TestCases @(
		@{ Mode = 'Single'; Inherited = $false },
		@{ Mode = 'Csv'; Inherited = $false },
		@{ Mode = 'Application'; Inherited = $false },
		@{ Mode = 'Single'; Inherited = $true },
		@{ Mode = 'Csv'; Inherited = $true },
		@{ Mode = 'Application'; Inherited = $true }
	) {
		param($Mode, $Inherited)
		$parameters = switch ($Mode) {
			'Single' { @{ GroupId = $script:groupId } }
			'Csv' {
				$csv = Join-Path $script:testScriptRoot 'preview.csv'
				"GroupId`n$script:groupId" | Set-Content $csv
				@{ CsvPath = $csv }
			}
			'Application' { @{ ServicePrincipalId = $script:groupId } }
		}
		if ($Inherited) { $WhatIfPreference = $true }
		else { $parameters.WhatIf = $true }
		$results = @(& $script:testScript @parameters)
		$results | Should -HaveCount 1
		$results[0].Status | Should -Be 'Would be created'
		$results[0].MemberCount | Should -Be 150
		Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse
		($script:fixture.HostMessages -join "`n") | Should -Not -Match 'Creation log:'
		$script:fixture.HostMessages | Should -Contain 'Total direct users in group reviews that would be created: 150'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'leaves legacy logs untouched during real runs and all existing logs untouched during previews' {
		$directory = Join-Path $script:testScriptRoot 'logs'
		$null = New-Item -ItemType Directory -Path $directory
		$legacy = Join-Path $directory 'Create-AccessReview.jsonl'
		$legacyContents = '{"Event":"ExistingRun"}'
		Set-Content -LiteralPath $legacy -Value $legacyContents -NoNewline
		$null = & $script:testScript -GroupId $script:groupId -Confirm:$false
		(Get-Content -LiteralPath $legacy -Raw) | Should -BeExactly $legacyContents
		$runFile = @(Get-RunLogs)[0]
		$runContents = Get-Content -LiteralPath $runFile.FullName -Raw
		$runWriteTime = $runFile.LastWriteTimeUtc
		$legacyWriteTime = (Get-Item -LiteralPath $legacy).LastWriteTimeUtc
		$null = & $script:testScript -GroupId $script:groupId -WhatIf
		@(Get-ChildItem -LiteralPath $directory -File) | Should -HaveCount 2
		(Get-Content -LiteralPath $runFile.FullName -Raw) | Should -BeExactly $runContents
		(Get-Item -LiteralPath $runFile.FullName).LastWriteTimeUtc | Should -Be $runWriteTime
		(Get-Content -LiteralPath $legacy -Raw) | Should -BeExactly $legacyContents
		(Get-Item -LiteralPath $legacy).LastWriteTimeUtc | Should -Be $legacyWriteTime
	}

	It 'does not require a writable log directory during WhatIf' {
		$blockedPath = Join-Path $script:testScriptRoot 'logs'
		Set-Content -LiteralPath $blockedPath -Value 'Blocked.' -NoNewline
		$result = & $script:testScript -GroupId $script:groupId -WhatIf
		$result.Status | Should -Be 'Would be created'
		(Get-Content -LiteralPath $blockedPath -Raw) | Should -BeExactly 'Blocked.'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'preserves preview failures without writing diagnostic logs' {
		$script:fixture.OwnerFailure = $true
		{ & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -WhatIf } |
			Should -Throw '*Owner read forbidden*'
		Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'logs a real run when explicit WhatIf false overrides inherited WhatIf' {
		$WhatIfPreference = $true
		$result = & $script:testScript -GroupId $script:groupId -WhatIf:$false -Confirm:$false
		$result.Status | Should -Be 'Created'
		$entries = Get-LogEntries
		$entries | Should -HaveCount 3
		$entries[0].Syntax | Should -Match '-WhatIf:\$false'
		$entries[2].Created | Should -Be 1
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 1 -Exactly
	}

	It 'logs each CSV failure and completes the batch' {
		$csv = Join-Path $script:testScriptRoot 'groups.csv'
		"GroupId`n$script:groupId`n44444444-4444-4444-4444-444444444444" | Set-Content $csv
		$script:fixture.CreateFailure = $true
		$null = & $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -Confirm:$false
		$entries = Get-LogEntries
		$entries.Count | Should -Be 4
		$entries[1].Result.Status | Should -Be 'Failed'
		$entries[2].Result.Status | Should -Be 'Failed'
		$entries[3].Total | Should -Be 2
		$entries[3].Failed | Should -Be 2
	}

	It 'fails before Graph creation if the log directory is not writable' {
		Set-Content (Join-Path $script:testScriptRoot 'logs') 'A file blocks the log directory.'
		{ & $script:testScript -GroupId $script:groupId } | Should -Throw '*Unable to write creation log*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'preserves report-only behavior for synchronized groups' {
		$script:fixture.Synced = $true
		$null = & $script:testScript -GroupId $script:groupId -ReviewerType ApplicationOwner -Confirm:$false -WarningAction SilentlyContinue
		$script:fixture.Bodies[0].settings.autoApplyDecisionsEnabled | Should -BeFalse
		$script:fixture.Bodies[0].settings.ContainsKey('applyActions') | Should -BeFalse
	}

	It 'discovers assigned groups once and retains application names and sign-off stages' {
		$script:fixture.Assignments += $script:fixture.Assignments[0]
		$script:fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -ReviewerType ApplicationOwner -Confirm:$false)
		$results | Should -HaveCount 2
		$results[0].ReviewName | Should -Be 'AR Demo app - Demo group'
		$results[0].Status | Should -Be 'Created'
		$script:fixture.Bodies[0].reviewers[0].query | Should -Be "/users/$script:ownerId"
		$results[1].ReviewName | Should -Be 'AR Demo app'
		$results[1].SubjectCount | Should -Be 1
	}

	It 'logs skipped application groups for <Kind>' -TestCases @(
		@{ Kind = 'dynamic' }, @{ Kind = 'existing' }, @{ Kind = 'empty' }
	) {
		param($Kind)
		switch ($Kind) {
			'dynamic' { $script:fixture.Dynamic = $true }
			'existing' { $script:fixture.ExistingNames = @('AR Demo app - Demo group') }
			'empty' { $script:fixture.Count = 0; $script:fixture.Owners = @() }
		}
		$result = & $script:testScript -ServicePrincipalId $script:groupId -Confirm:$false
		$result.Status | Should -Be 'Skipped'
		$entries = Get-LogEntries
		($entries | Where-Object Event -EQ 'GroupResult').Result.Status | Should -Be 'Skipped'
		($entries | Where-Object Event -EQ 'RunCompleted').Skipped | Should -Be 1
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'creates an application-scoped review for direct users without reading groups' {
		$script:fixture.Assignments = @(
			[pscustomobject]@{ Id = 'assignment-1'; PrincipalType = 'User'; PrincipalId = $script:ownerId },
			[pscustomobject]@{ Id = 'assignment-2'; PrincipalType = 'User'; PrincipalId = $script:ownerId },
			[pscustomobject]@{ Id = 'assignment-3'; PrincipalType = 'User'; PrincipalId = $script:groupOwnerId }
		)
		$result = & $script:testScript -ServicePrincipalId $script:groupId -Confirm:$false
		$result.Status | Should -Be 'Created'
		$result.TargetType | Should -Be 'ApplicationDirectUsers'
		$result.ReviewName | Should -Be 'AR Demo app'
		$result.GroupId | Should -BeNullOrEmpty
		$result.MemberCount | Should -BeNullOrEmpty
		$result.SubjectCount | Should -Be 2
		$result.PrincipalIds | Should -Be @($script:ownerId, $script:groupOwnerId)
		$result.AssignmentIds | Should -Be @('assignment-1', 'assignment-2', 'assignment-3')
		$script:fixture.HostMessages | Should -Contain 'Discovered app-role assignments: 0 distinct group(s), 2 distinct direct user(s), 0 other assignment(s).'
		$script:fixture.HostMessages | Should -Contain "Application: Demo app (service-principal Object ID: $script:groupId)"
		$body = $script:fixture.Bodies[0]
		$body.DisplayName | Should -Be 'AR Demo app'
		$scope = $body.scope.AdditionalProperties
		$scope.'@odata.type' | Should -Be '#microsoft.graph.principalResourceMembershipsScope'
		$scope.principalScopes | Should -HaveCount 1
		$scope.principalScopes[0].query | Should -Be '/v1.0/users'
		$scope.resourceScopes[0].query | Should -Be "/v1.0/servicePrincipals/$script:groupId"
		$body.settings.autoApplyDecisionsEnabled | Should -BeTrue
		$body.settings.applyActions[0].AdditionalProperties.'@odata.type' | Should -Be '#microsoft.graph.removeAccessApplyAction'
		$result.Stages | Should -Be 1
		$body.reviewers[0].query | Should -Be "/users/$script:ownerId"
		$entries = Get-LogEntries
		$applicationResult = $entries | Where-Object Event -EQ 'ApplicationResult'
		$applicationResult.Result.SubjectCount | Should -Be 2
		$applicationResult.Result.AssignmentIds | Should -Be $result.AssignmentIds
		$script:fixture.HostMessages | Should -Contain 'Total direct users in group created reviews: 0'
		$script:fixture.HostMessages | Should -Contain 'Directly assigned users in created reviews: 2'
		Should -Invoke Get-MgGroup -Times 0 -Exactly
		Should -Invoke Get-MgGroupMember -Times 0 -Exactly
		Should -Invoke Get-MgServicePrincipalAppRoleAssignedTo -Times 1 -Exactly -ParameterFilter {
			$All -and $Property -contains 'principalType' -and $Property -contains 'principalId' -and $Property -contains 'id'
		}
		Should -Invoke Get-MgServicePrincipalOwner -Times 1 -Exactly -ParameterFilter { $All }
	}

	It 'reports owners and workload or unknown assignments without pretending they are reviewed (preview: <Preview>)' -TestCases @(
		@{ Preview = $false }, @{ Preview = $true }
	) {
		param($Preview)
		$script:fixture.ApplicationOwners = @(
			[pscustomobject]@{ Id = $script:ownerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user'; displayName = 'App owner' } },
			[pscustomobject]@{ Id = $script:reviewId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.servicePrincipal'; displayName = 'Automation owner' } }
		)
		$script:fixture.Assignments = @(
			[pscustomobject]@{ Id = 'workload-role-1'; PrincipalType = 'ServicePrincipal'; PrincipalId = $script:reviewId; PrincipalDisplayName = 'Automation'; AppRoleId = 'role-1' },
			[pscustomobject]@{ Id = 'workload-role-2'; PrincipalType = 'ServicePrincipal'; PrincipalId = $script:reviewId; PrincipalDisplayName = 'Automation'; AppRoleId = 'role-2' },
			[pscustomobject]@{ Id = 'unknown-assignment'; PrincipalType = 'FutureIdentity'; PrincipalId = $script:groupOwnerId; AppRoleId = 'role-3' }
		)
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -WhatIf:$Preview -Confirm:$false -WarningAction SilentlyContinue)
		$results | Should -HaveCount 5
		@($results.Status | Select-Object -Unique) | Should -Be @('Manual review required')
		@($results | Where-Object TargetType -eq 'ServicePrincipalOwner') | Should -HaveCount 2
		$results[0].PrincipalDisplayName | Should -Be 'App owner'
		$results[1].PrincipalType | Should -Be 'servicePrincipal'
		$results[2].AssignmentId | Should -Be 'workload-role-1'
		$results[3].AppRoleId | Should -Be 'role-2'
		$results[4].PrincipalType | Should -Be 'FutureIdentity'
		$results[4].Reason | Should -Match 'manually'
		$script:fixture.HostMessages | Should -Contain 'Manual review required: 5 ownership/assignment item(s). These are not covered by the created reviews.'
		if ($Preview) {
			Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse
		}
		else {
			$entries = Get-LogEntries
			@($entries | Where-Object Event -eq 'ManualReviewRequired') | Should -HaveCount 5
			$entries[-1].Created | Should -Be 0
			$entries[-1].ManualReviewRequired | Should -Be 5
		}
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
		Should -Invoke Get-MgGroup -Times 0 -Exactly
	}

	It 'uses actual service-principal user owners for ResourceOwner routing and still reports ownership separately' {
		$script:fixture.ApplicationOwners = @(
			[pscustomobject]@{ Id = $script:groupOwnerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } },
			[pscustomobject]@{ Id = $script:reviewId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.servicePrincipal' } }
		)
		$script:fixture.Assignments = @(1..150 | ForEach-Object {
			[pscustomobject]@{ PrincipalType = 'User'; PrincipalId = ([guid]::NewGuid().ToString()) }
		})
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -ReviewerType ResourceOwner -Confirm:$false -WarningAction SilentlyContinue)
		$results | Should -HaveCount 3
		$results[-1].StageNames | Should -Be 'Resource owners (service-principal owners)'
		$script:fixture.Bodies[0].reviewers.query | Should -Be @("/users/$script:groupOwnerId")
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
	}

	It 'uses service-principal user owners after managers for <Count> direct users without application owners' -TestCases @(
		@{ Count = 151 }, @{ Count = 500 }
	) {
		param($Count)
		$script:fixture.ApplicationOwners = @(
			[pscustomobject]@{ Id = $script:groupOwnerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
		)
		$script:fixture.Assignments = @(1..$Count | ForEach-Object {
			[pscustomobject]@{ PrincipalType = 'User'; PrincipalId = ([guid]::NewGuid().ToString()) }
		})
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -ReviewerType ResourceOwner `
			-AppOwnerIds @() -Confirm:$false -WarningAction SilentlyContinue)
		$results[-1].StageNames | Should -Be $(if ($Count -ge 500) { 'Self-review -> Line managers -> Resource owners (service-principal owners)' }
			else { 'Line managers -> Resource owners (service-principal owners)' })
		$script:fixture.Bodies[0].stageSettings[-2].reviewers[0].query | Should -Be './manager'
		$script:fixture.Bodies[0].stageSettings[-2].fallbackReviewers.query | Should -Be @("/users/$script:groupOwnerId")
		$script:fixture.Bodies[0].stageSettings[-1].reviewers.query | Should -Be @("/users/$script:groupOwnerId")
	}

	It 'previews mixed subjects without creating reviews or disk logs' {
		$script:fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
		$script:fixture.ApplicationOwners = @(
			[pscustomobject]@{ Id = $script:ownerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
		)
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -WhatIf -WarningAction SilentlyContinue)
		$results.Status | Should -Be @('Manual review required', 'Would be created', 'Would be created')
		$manualFields = @($results[0].PSObject.Properties.Name | Sort-Object)
		foreach ($result in $results) {
			@($result.PSObject.Properties.Name | Sort-Object) | Should -Be $manualFields
		}
		Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse
		$script:fixture.HostMessages | Should -Contain '2 review target(s): 2 would be created, 0 skipped, 0 failed.'
		$script:fixture.HostMessages | Should -Contain 'Directly assigned users in reviews that would be created: 1'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'skips an existing direct-user review without skipping ownership reporting' {
		$script:fixture.Assignments = @([pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId })
		$script:fixture.ExistingNames = @('AR Demo app')
		$script:fixture.ApplicationOwners = @(
			[pscustomobject]@{ Id = $script:ownerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
		)
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -Confirm:$false -WarningAction SilentlyContinue)
		$results.Status | Should -Be @('Manual review required', 'Skipped')
		$results[-1].Reason | Should -Match 'already exists'
		$script:fixture.HostMessages | Should -Contain 'Discovered app-role assignments: 0 distinct group(s), 1 distinct direct user(s), 0 other assignment(s).'
		($script:fixture.HostMessages -join "`n") | Should -Match 'Direct-user review: Skipped; discovered users: 1;.*already exists'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'does not create direct-user reviews after incomplete group discovery' {
		$script:fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
		Mock Get-MgGroup { throw 'Group unavailable.' }
		{ & $script:testScript -ServicePrincipalId $script:groupId -Confirm:$false } |
			Should -Throw '*Group discovery failed; no reviews were created*Group unavailable*'
		(Get-LogEntries)[-1].Event | Should -Be 'RunFailed'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'does not confuse a group named Direct user assignments with the application review' {
		$script:fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
		Mock Get-MgGroup {
			[pscustomobject]@{ DisplayName = 'Direct user assignments'; GroupTypes = @(); OnPremisesSyncEnabled = $false }
		}
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -Confirm:$false)
		$results.Status | Should -Be @('Created', 'Created')
		$results[0].ReviewName | Should -Not -Be $results[1].ReviewName
	}

	It 'preserves stage policy and recurrence for large direct-user reviews with <Type> routing' -TestCases @(
		@{ Type = 'LineManager'; Stages = 3 },
		@{ Type = 'ApplicationOwner'; Stages = 3 }
	) {
		param($Type, $Stages)
		$script:fixture.Assignments = @(1..500 | ForEach-Object {
			[pscustomobject]@{ PrincipalType = 'User'; PrincipalId = ([guid]::NewGuid().ToString()) }
		})
		$start = (Get-Date).Date
		$result = & $script:testScript -ServicePrincipalId $script:groupId -ReviewerType $Type -Frequency Quarterly -EndDate $start.AddYears(1) -Confirm:$false
		$result.Stages | Should -Be $Stages
		$result.SubjectCount | Should -Be 500
		$body = $script:fixture.Bodies[0]
		$body.settings.instanceDurationInDays | Should -Be ($Stages * 5)
		$body.settings.recurrence.pattern.interval | Should -Be 3
		if ($Stages -eq 3) {
			$body.stageSettings[0].reviewers | Should -HaveCount 0
			$body.stageSettings[-1].reviewers[0].query | Should -Be "/users/$script:ownerId"
		}
		else {
			$body.reviewers[0].query | Should -Be "/users/$script:ownerId"
		}
	}

	It 'reports failed direct-user creation without adding users to the created total' {
		$script:fixture.Assignments = @([pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId })
		$script:fixture.CreateFailure = $true
		$result = & $script:testScript -ServicePrincipalId $script:groupId -Confirm:$false
		$result.Status | Should -Be 'Failed'
		$result.SubjectCount | Should -Be 1
		$result.Reason | Should -Match 'Create review forbidden'
		$script:fixture.HostMessages | Should -Contain 'Directly assigned users in created reviews: 0'
		(Get-LogEntries)[-1].Failed | Should -Be 1
	}

	It 'fails discovery explicitly if owners cannot be read, before any reviews are created' {
		Mock Get-MgServicePrincipalOwner { throw 'Owner access denied.' }
		{ & $script:testScript -ServicePrincipalId $script:groupId -Confirm:$false } | Should -Throw '*Owner access denied*'
		(Get-LogEntries)[-1].Event | Should -Be 'RunFailed'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'logs empty application discovery without creating a review' {
		$script:fixture.Assignments = @()
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -WarningAction SilentlyContinue)
		$results | Should -HaveCount 0
		((Get-LogEntries) | Where-Object Event -EQ 'RunCompleted').Total | Should -Be 0
		$application = ((Get-LogEntries) | Where-Object Event -EQ 'ApplicationDiscovered').Application
		$application.ServicePrincipalId | Should -Be $script:groupId
		$application.DirectUserCount | Should -Be 0
		@($application.GroupIds) | Should -HaveCount 0
		$script:fixture.HostMessages | Should -Contain '0 review target(s): 0 created, 0 skipped, 0 failed.'
		$script:fixture.HostMessages | Should -Contain 'Total direct users in group created reviews: 0'
	}

	It 'logs discovery errors and stops' {
		Mock Get-MgServicePrincipal { throw 'Application read forbidden.' }
		{ & $script:testScript -ServicePrincipalId $script:groupId } | Should -Throw '*Application read forbidden*'
		(Get-LogEntries)[1].Event | Should -Be 'RunFailed'
	}

	It 'does not create application-discovered reviews during WhatIf' {
		$result = & $script:testScript -ServicePrincipalId $script:groupId -WhatIf
		$result.Status | Should -Be 'Would be created'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'requires explicit sign-off owners for application discovery' {
		{ & $script:testScript -ServicePrincipalId $script:groupId -AppOwnerIds @() } | Should -Throw '*Supply -AppOwnerIds*'
		Should -Invoke Get-MgServicePrincipal -Times 0 -Exactly
	}

	It 'requires application read scope only for application discovery' {
		Mock Get-MgContext { @{ Scopes = @('AccessReview.ReadWrite.All') } }
		$null = & $script:testScript -GroupId $script:groupId -WhatIf
		{ & $script:testScript -ServicePrincipalId $script:groupId -WhatIf } | Should -Throw '*Application.Read.All*'
	}

	It 'stops before creating any application group reviews if group discovery fails' {
		$second = '44444444-4444-4444-4444-444444444444'
		$script:fixture.Assignments += [pscustomobject]@{ PrincipalType = 'Group'; PrincipalId = $second }
		Mock Get-MgGroup { throw 'Group unavailable.' } -ParameterFilter { "$GroupId" -eq '11111111-1111-1111-1111-111111111111' }
		{ & $script:testScript -ServicePrincipalId $script:groupId -Confirm:$false } |
			Should -Throw '*Group discovery failed; no reviews were created*Group unavailable*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'preserves quarterly recurrence and three-stage duration' {
		$script:fixture.Count = 500
		$start = (Get-Date).Date.AddDays(1)
		$end = $start.AddYears(1)
		$null = & $script:testScript -GroupId $script:groupId -Frequency Quarterly -StartDate $start -EndDate $end -Confirm:$false
		$settings = $script:fixture.Bodies[0].settings
		$settings.instanceDurationInDays | Should -Be 15
		$settings.recurrence.pattern.interval | Should -Be 3
		([datetime]$settings.recurrence.range.endDate).Date | Should -Be $end.Date
	}

	It 'rejects a recurring schedule too short for every stage' {
		$script:fixture.Count = 151
		$start = (Get-Date).Date
		{ & $script:testScript -GroupId $script:groupId -Frequency Monthly -StartDate $start -EndDate $start.AddDays(9) } |
			Should -Throw '*at least 10 days*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'rejects combining application and CSV modes' {
		{ & $script:testScript -ServicePrincipalId $script:groupId -CsvPath 'groups.csv' } |
			Should -Throw '*Parameter set cannot be resolved*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'detects <Kind> column delimiters with quoted headers and multiple owners' -TestCases @(
		@{ Kind = 'comma'; Delimiter = ',' },
		@{ Kind = 'semicolon'; Delimiter = ';' },
		@{ Kind = 'tab'; Delimiter = "`t" }
	) {
		param($Kind, $Delimiter)
		$csv = Join-Path $script:testScriptRoot 'delimited.csv'
		$header = @('"GroupId"', '"AppOwnerIds"', '"ReviewName"') -join $Delimiter
		$data = @($script:groupId, "`"$script:ownerId;$script:reviewId`"", '"Finance, reporting; review"') -join $Delimiter
		"$header`n$data" | Set-Content -LiteralPath $csv -Encoding utf8BOM
		$result = & $script:testScript -CsvPath $csv -AppOwnerIds @() -ReviewerType ApplicationOwner -Confirm:$false
		$result.Status | Should -Be 'Created'
		$result.ReviewName | Should -Be 'Finance, reporting; review'
		$script:fixture.Bodies[0].reviewers.query |
			Should -Be @("/users/$script:ownerId", "/users/$script:reviewId")
	}

	It 'uses every comma-separated owner with <Kind> columns and <Type> reviewers' -TestCases @(
		@{ Kind = 'comma'; Delimiter = ','; Type = 'ApplicationOwner' },
		@{ Kind = 'comma'; Delimiter = ','; Type = 'LineManager' },
		@{ Kind = 'comma'; Delimiter = ','; Type = 'ResourceOwner' },
		@{ Kind = 'semicolon'; Delimiter = ';'; Type = 'ApplicationOwner' },
		@{ Kind = 'semicolon'; Delimiter = ';'; Type = 'LineManager' },
		@{ Kind = 'semicolon'; Delimiter = ';'; Type = 'ResourceOwner' },
		@{ Kind = 'tab'; Delimiter = "`t"; Type = 'ApplicationOwner' },
		@{ Kind = 'tab'; Delimiter = "`t"; Type = 'LineManager' },
		@{ Kind = 'tab'; Delimiter = "`t"; Type = 'ResourceOwner' }
	) {
		param($Kind, $Delimiter, $Type)
		$csv = Join-Path $script:testScriptRoot 'comma-owners.csv'
		$header = @('GroupId', 'AppOwnerIds') -join $Delimiter
		$owners = "$script:ownerId, $script:reviewId, $script:ownerId"
		if ($Delimiter -eq ',') { $owners = "`"$owners`"" }
		$data = @($script:groupId, $owners) -join $Delimiter
		"$header`n$data" | Set-Content -LiteralPath $csv
		$result = & $script:testScript -CsvPath $csv -AppOwnerIds @() -ReviewerType $Type -Confirm:$false
		$result.Status | Should -Be 'Created'
		$result.AppOwnerIds | Should -Be @($script:ownerId, $script:reviewId)
		$body = $script:fixture.Bodies[0]
		if ($Type -eq 'ResourceOwner') {
			$body.reviewers.query | Should -Be @("/users/$script:groupOwnerId")
		}
		else {
			$body.reviewers.query | Should -Be @("/users/$script:ownerId", "/users/$script:reviewId")
		}
		(Get-LogEntries)[1].Result.AppOwnerIds | Should -Be @($script:ownerId, $script:reviewId)
	}

	It 'rejects malformed comma-separated owners without creating a partial review' {
		$csv = Join-Path $script:testScriptRoot 'invalid-owners.csv'
		"GroupId;AppOwnerIds`n$script:groupId;$script:ownerId,not-a-guid" | Set-Content -LiteralPath $csv
		$result = & $script:testScript -CsvPath $csv -Confirm:$false
		$result.Status | Should -Be 'Failed'
		$result.Reason | Should -Match 'invalid user object GUID'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'explains unsupported headers instead of claiming a valid delimiter was parsed' {
		$csv = Join-Path $script:testScriptRoot 'invalid.csv'
		"GroupId|AppOwnerIds`n$script:groupId|$script:ownerId" | Set-Content $csv
		{ & $script:testScript -CsvPath $csv } | Should -Throw '*Supported column delimiters*Detected header: GroupId|AppOwnerIds*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'rejects an empty CSV explicitly' {
		$csv = Join-Path $script:testScriptRoot 'empty.csv'
		Set-Content $csv ''
		{ & $script:testScript -CsvPath $csv } | Should -Throw '*empty or has a blank header*'
	}
}
