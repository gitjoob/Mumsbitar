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
}

Describe 'Unified access review creation and local logging' {
	BeforeEach {
		$script:testScriptRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString())
		$null = New-Item -ItemType Directory -Path $script:testScriptRoot
		$script:testScript = Join-Path $script:testScriptRoot 'Create-AccessReview.ps1'
		Copy-Item $script:sourcePath $script:testScript
		Copy-Item (Join-Path $PSScriptRoot '../AccessReview.Logging.ps1') $script:testScriptRoot
		$PSDefaultParameterValues['Create-AccessReview.ps1:AppOwnerIds'] = @($script:ownerId)
		$script:fixture = @{
			Count = 150
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
			[pscustomobject]@{
				Id = "$GroupId"; DisplayName = 'Demo group'; OnPremisesSyncEnabled = $fixture.Synced
				GroupTypes = $(if ($fixture.Dynamic) { @('DynamicMembership') } else { @() })
			}
		}.GetNewClosure())
		Mock Get-MgGroupTransitiveMemberCount ({ $fixture.Count }.GetNewClosure())
		Mock Get-MgGroupOwner ({
			if ($fixture.OwnerFailure) { throw 'Owner read forbidden.' }
			$fixture.Owners
		}.GetNewClosure())
		Mock Get-MgServicePrincipal { [pscustomobject]@{ DisplayName = 'Demo app' } }
		Mock Get-MgServicePrincipalAppRoleAssignedTo ({ $fixture.Assignments }.GetNewClosure())
		Mock Get-MgServicePrincipalOwner ({ $fixture.ApplicationOwners }.GetNewClosure())
		Mock Get-MgIdentityGovernanceAccessReviewDefinition ({
			foreach ($name in $fixture.ExistingNames) { [pscustomobject]@{ DisplayName = $name } }
		}.GetNewClosure())
		Mock New-MgIdentityGovernanceAccessReviewDefinition ({
			param($BodyParameter)
			if ($fixture.CreateFailure) { throw 'Create review forbidden.' }
			$fixture.Bodies.Add($BodyParameter)
			[pscustomobject]@{ Id = $fixture.ReviewId }
		}.GetNewClosure())
	}

	It 'uses independent durations for <Case> routing' -TestCases @(
		@{ Case = 'single'; Count = 22; Type = 'ResourceOwner'; Force = $false; Durations = @(5) },
		@{ Case = 'forced'; Count = 22; Type = 'ResourceOwner'; Force = $true; Durations = @(7, 5) },
		@{ Case = 'managers'; Count = 150; Type = 'LineManager'; Force = $false; Durations = @(7, 5) },
		@{ Case = 'three-stage'; Count = 500; Type = 'ResourceOwner'; Force = $false; Durations = @(3, 7, 5) },
		@{ Case = 'appowners'; Count = 500; Type = 'ApplicationOwner'; Force = $false; Durations = @(5) },
		@{ Case = 'merged'; Count = 500; Type = 'ResourceOwner'; Force = $false; Durations = @(3, 5) },
		@{ Case = 'ownerless'; Count = 500; Type = 'ResourceOwner'; Force = $false; Durations = @(3, 5) },
		@{ Case = 'fallback'; Count = 150; Type = 'ResourceOwner'; Force = $false; Durations = @(7, 5) }
	) {
		param($Case, $Count, $Type, $Force, $Durations)
		$script:fixture.Count = $Count
		$extra = @{}
		if ($Case -eq 'merged') {
			$script:fixture.Owners = @([pscustomobject]@{
				Id = $script:ownerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' }
			})
		}
		if ($Case -in @('ownerless', 'fallback')) { $script:fixture.Owners = @() }
		if ($Case -eq 'fallback') { $extra.FallbackReviewerUserId = $script:reviewId }
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType $Type -ForceTwoStage:$Force @extra `
			-SelfReviewDurationInDays 3 -SelectedReviewerDurationInDays 7 -ApplicationOwnerDurationInDays 5 -Confirm:$false -WarningAction SilentlyContinue
		$body = $script:fixture.Bodies[0]
		$total = ($Durations | Measure-Object -Sum).Sum
		$result.Stages | Should -Be $Durations.Count
		$body.settings.instanceDurationInDays | Should -Be $total
		([datetime]$body.settings.recurrence.range.endDate).Date | Should -Be ((Get-Date).Date.AddDays($total))
		if ($Durations.Count -gt 1) {
			@($body.stageSettings.durationInDays) | Should -Be $Durations
		}
		else { $body.ContainsKey('stageSettings') | Should -BeFalse }
		$entry = (Get-LogEntries)[0]
		$entry.SelfReviewDurationInDays | Should -Be 3
		$entry.SelectedReviewerDurationInDays | Should -Be 7
		$entry.ApplicationOwnerDurationInDays | Should -Be 5
		$entry.Syntax | Should -Match "-SelectedReviewerDurationInDays '7'"
	}

	It 'uses the application-owner duration for a single-stage review' {
		$script:fixture.Count = 22
		$null = & $script:testScript -GroupId $script:groupId -ApplicationOwnerDurationInDays 9 -Confirm:$false
		$script:fixture.Bodies[0].settings.instanceDurationInDays | Should -Be 9
	}

	It 'applies custom durations in <Mode> mode (preview: <Preview>)' -TestCases @(
		@{ Mode = 'Csv'; Preview = $false },
		@{ Mode = 'Application'; Preview = $false },
		@{ Mode = 'Application'; Preview = $true }
	) {
		param($Mode, $Preview)
		if ($Mode -eq 'Csv') {
			$csv = Join-Path $script:testScriptRoot 'durations.csv'
			"GroupId`n$script:groupId" | Set-Content $csv
			$parameters = @{ CsvPath = $csv }
		}
		else {
			$script:fixture.Assignments = @([pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId })
			$parameters = @{ ServicePrincipalId = $script:groupId }
		}
		$result = & $script:testScript @parameters -ForceTwoStage -SelectedReviewerDurationInDays 7 -ApplicationOwnerDurationInDays 4 -WhatIf:$Preview -Confirm:$false
		$result.Stages | Should -Be 2
		if ($Preview) {
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
			Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse
		}
		else {
			@($script:fixture.Bodies[0].stageSettings.durationInDays) | Should -Be @(7, 4)
			$script:fixture.Bodies[0].settings.instanceDurationInDays | Should -Be 11
		}
	}

	It 'validates recurring schedules against the actual duration sum (days: <Days>)' -TestCases @(
		@{ Days = 11; Valid = $false }, @{ Days = 12; Valid = $true }
	) {
		param($Days, $Valid)
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

	It 'sums stage durations without measuring hashtable properties' {
		Mock Measure-Object { throw 'Hashtable property measurement is unavailable.' } -ParameterFilter {
			$Property -contains 'durationInDays'
		}
		$result = & $script:testScript -GroupId $script:groupId -SelectedReviewerDurationInDays 7 -ApplicationOwnerDurationInDays 4 -Confirm:$false
		$result.Stages | Should -Be 2
		$script:fixture.Bodies[0].settings.instanceDurationInDays | Should -Be 11
	}

	It 'uses fallback reviewers for synced groups without querying owners (explicit: <Explicit>, forced: <Forced>, count: <Count>)' -TestCases @(
		@{ Explicit = $true; Forced = $false; Count = 150; Stages = 2 },
		@{ Explicit = $false; Forced = $false; Count = 150; Stages = 1 },
		@{ Explicit = $false; Forced = $true; Count = 22; Stages = 2 },
		@{ Explicit = $true; Forced = $true; Count = 22; Stages = 2 },
		@{ Explicit = $true; Forced = $false; Count = 500; Stages = 3 }
	) {
		param($Explicit, $Forced, $Count, $Stages)
		$script:fixture.Synced = $true
		$script:fixture.Count = $Count
		$script:fixture.OwnerFailure = $true
		$parameters = @{}
		$fallbackId = $script:ownerId
		if ($Explicit) {
			$parameters.FallbackReviewerUserId = $script:reviewId
			$fallbackId = $script:reviewId
		}
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -ForceTwoStage:$Forced @parameters `
			-SelectedReviewerDurationInDays 7 -ApplicationOwnerDurationInDays 4 -Confirm:$false -WarningAction SilentlyContinue
		$result.Stages | Should -Be $Stages
		$body = $script:fixture.Bodies[0]
		$body.settings.autoApplyDecisionsEnabled | Should -BeFalse
		$body.settings.ContainsKey('applyActions') | Should -BeFalse
		if ($Stages -gt 1) {
			$body.stageSettings[-2].reviewers[0].query | Should -Be "/users/$fallbackId"
			$body.stageSettings[-2].durationInDays | Should -Be 7
			$result.StageNames | Should -Match 'Fallback reviewer \(synced group\)'
		}
		else {
			$body.reviewers[0].query | Should -Be "/users/$script:ownerId"
			$body.settings.instanceDurationInDays | Should -Be 4
		}
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
	}

	It 'preserves default manager routing at the two-stage threshold' {
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
		$script:fixture.HostMessages[-1] | Should -Be "Total transitive members in ${label}: $Count"
		Should -Invoke Get-MgGroupTransitiveMemberCount -Times 1 -Exactly -ParameterFilter {
			"$GroupId" -eq $script:groupId -and $ConsistencyLevel -eq 'eventual'
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
		Mock Get-MgGroupTransitiveMemberCount { 75 } -ParameterFilter {
			"$GroupId" -eq '44444444-4444-4444-4444-444444444444'
		}
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
		Should -Invoke Get-MgGroupTransitiveMemberCount -Times 2 -Exactly
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

	It 'excludes failed and skipped application groups from the member total' {
		$second = '44444444-4444-4444-4444-444444444444'
		$third = '55555555-5555-5555-5555-555555555555'
		$script:fixture.Assignments += @(
			[pscustomobject]@{ PrincipalType = 'Group'; PrincipalId = $second },
			[pscustomobject]@{ PrincipalType = 'Group'; PrincipalId = $third }
		)
		Mock Get-MgGroup {
			[pscustomobject]@{ DisplayName = 'Dynamic group'; GroupTypes = @('DynamicMembership') }
		} -ParameterFilter { "$GroupId" -eq '44444444-4444-4444-4444-444444444444' }
		Mock Get-MgGroupTransitiveMemberCount { throw 'Count unavailable.' } -ParameterFilter {
			"$GroupId" -eq '55555555-5555-5555-5555-555555555555'
		}
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -Confirm:$false)
		$results.Status | Should -Be @('Created', 'Skipped', 'Failed')
		$results[1].MemberCount | Should -Be 150
		$results[2].MemberCount | Should -BeNullOrEmpty
		$results[2].Reason | Should -Match 'Count unavailable'
		$script:fixture.HostMessages | Should -Contain '3 review target(s): 1 created, 1 skipped, 1 failed.'
		$script:fixture.HostMessages | Should -Contain 'Total transitive members in created reviews: 150'
	}

	It 'excludes groups whose review creation failed after retrieving a count' {
		$script:fixture.CreateFailure = $true
		$csv = Join-Path $script:testScriptRoot 'failed-counts.csv'
		"GroupId`n$script:groupId" | Set-Content $csv
		$result = & $script:testScript -CsvPath $csv -Confirm:$false
		$result.Status | Should -Be 'Failed'
		$result.MemberCount | Should -Be 150
		$script:fixture.HostMessages[-1] | Should -Be 'Total transitive members in created reviews: 0'
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
		$script:fixture.HostMessages[-1] | Should -Be 'Total transitive members in created reviews: 0'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'selects only human group owners and retains application owners as the final stage' {
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false
		$result.StageNames | Should -Be 'Resource owners (group owners) -> Application owners'
		$body = $script:fixture.Bodies[0]
		$body.stageSettings[0].reviewers.Count | Should -Be 1
		$body.stageSettings[0].reviewers[0].query | Should -Be "/users/$script:groupOwnerId"
		$body.stageSettings[1].dependsOn | Should -Be @('1')
	}

	It 'merges identical explicit reviewers at <Count> members (fallback: <UseFallback>, multiple: <Multiple>)' -TestCases @(
		@{ Count = 150; UseFallback = $false; Multiple = $false },
		@{ Count = 500; UseFallback = $false; Multiple = $false },
		@{ Count = 150; UseFallback = $false; Multiple = $true },
		@{ Count = 500; UseFallback = $false; Multiple = $true },
		@{ Count = 150; UseFallback = $true; Multiple = $false },
		@{ Count = 500; UseFallback = $true; Multiple = $false }
	) {
		param($Count, $UseFallback, $Multiple)
		$script:fixture.Count = $Count
		$ids = @('abcdefab-abcd-abcd-abcd-abcdefabcdef')
		if ($Multiple) { $ids += $script:ownerId }
		$script:fixture.Owners = @(
			foreach ($id in @($ids | Sort-Object -Descending) + @($ids[0])) {
				[pscustomobject]@{ Id = $id.ToUpperInvariant(); AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
			}
		)
		$parameters = @{ AppOwnerIds = $ids }
		if ($UseFallback) {
			$script:fixture.Owners = @()
			$parameters.FallbackReviewerUserId = $ids[0]
		}
		$start = (Get-Date).Date
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner @parameters -StartDate $start -Frequency Monthly -EndDate $start.AddDays(10) -Confirm:$false -WarningAction SilentlyContinue
		$body = $script:fixture.Bodies[0]
		$expectedQueries = @($ids | ForEach-Object { "/users/$_" })
		if ($Count -lt 500) {
			$result.Stages | Should -Be 1
			$result.StageNames | Should -Be 'Application owners'
			$body.ContainsKey('stageSettings') | Should -BeFalse
			@($body.reviewers.query) | Should -Be $expectedQueries
		}
		else {
			$result.Stages | Should -Be 2
			$result.StageNames | Should -Be 'Self-review -> Application owners'
			$body.stageSettings[0].reviewers | Should -HaveCount 0
			@($body.stageSettings[1].reviewers.query) | Should -Be $expectedQueries
			$body.stageSettings[1].dependsOn | Should -Be @('1')
			$body.stageSettings.stageId | Should -Be @('1', '2')
		}
		$body.settings.instanceDurationInDays | Should -Be ($result.Stages * 5)
		(Get-LogEntries)[1].Result.Stages | Should -Be $result.Stages
		(Get-LogEntries)[1].Result.StageNames | Should -Be $result.StageNames
	}

	It 'does not merge partially overlapping reviewer sets (extra group owner: <ExtraGroupOwner>)' -TestCases @(
		@{ ExtraGroupOwner = $false }, @{ ExtraGroupOwner = $true }
	) {
		param($ExtraGroupOwner)
		$script:fixture.Owners = @(
			[pscustomobject]@{ Id = $script:ownerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
		)
		$ids = @($script:ownerId)
		if ($ExtraGroupOwner) {
			$script:fixture.Owners += [pscustomobject]@{ Id = $script:reviewId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } }
		}
		else { $ids += $script:reviewId }
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -AppOwnerIds $ids -Confirm:$false
		$result.Stages | Should -Be 2
		$script:fixture.Bodies[0].stageSettings | Should -HaveCount 2
	}

	It 'merges matching CSV reviewers during preview without creating a review' {
		$csv = Join-Path $script:testScriptRoot 'matching-owners.csv'
		"GroupId;AppOwnerIds`n$script:groupId;$script:groupOwnerId" | Set-Content -LiteralPath $csv
		$result = & $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -WhatIf -WarningAction SilentlyContinue
		$result.Status | Should -Be 'Would be created'
		$result.Stages | Should -Be 1
		$result.StageNames | Should -Be 'Application owners'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'preserves self-review in three-stage owner reviews' {
		$script:fixture.Count = 500
		$null = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false
		$body = $script:fixture.Bodies[0]
		$body.stageSettings.Count | Should -Be 3
		$body.stageSettings[0].reviewers | Should -HaveCount 0
		$body.stageSettings[1].reviewers[0].query | Should -Be "/users/$script:groupOwnerId"
		$body.settings.instanceDurationInDays | Should -Be 15
	}

	It 'uses only application-owner sign-off below 150 members' -TestCases @(
		@{ Type = 'LineManager'; Count = 1 }, @{ Type = 'ResourceOwner'; Count = 1 },
		@{ Type = 'LineManager'; Count = 49 }, @{ Type = 'ResourceOwner'; Count = 50 },
		@{ Type = 'LineManager'; Count = 149 }, @{ Type = 'ResourceOwner'; Count = 149 }
	) {
		param($Type, $Count)
		$script:fixture.Count = $Count
		$null = & $script:testScript -GroupId $script:groupId -ReviewerType $Type -Confirm:$false
		$body = $script:fixture.Bodies[0]
		$body.ContainsKey('stageSettings') | Should -BeFalse
		$body.reviewers[0].query | Should -Be "/users/$script:ownerId"
		$body.settings.instanceDurationInDays | Should -Be 5
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
	}

	It 'forces two stages for 22 members in <Mode> mode with <Type> reviewers (preview: <Preview>)' -TestCases @(
		@{ Mode = 'Single'; Type = 'ResourceOwner'; Preview = $false },
		@{ Mode = 'Single'; Type = 'LineManager'; Preview = $false },
		@{ Mode = 'Csv'; Type = 'ResourceOwner'; Preview = $false },
		@{ Mode = 'Application'; Type = 'ResourceOwner'; Preview = $false },
		@{ Mode = 'Single'; Type = 'ResourceOwner'; Preview = $true }
	) {
		param($Mode, $Type, $Preview)
		$script:fixture.Count = 22
		$parameters = switch ($Mode) {
			'Single' { @{ GroupId = $script:groupId } }
			'Csv' {
				$csv = Join-Path $script:testScriptRoot 'forced.csv'
				"GroupId`n$script:groupId" | Set-Content $csv
				@{ CsvPath = $csv }
			}
			'Application' {
				$script:fixture.ApplicationOwners = @($script:fixture.Owners[0])
				$script:fixture.Assignments = @(1..22 | ForEach-Object {
					[pscustomobject]@{ PrincipalType = 'User'; PrincipalId = ([guid]::NewGuid().ToString()) }
				})
				@{ ServicePrincipalId = $script:groupId }
			}
		}
		$results = @(& $script:testScript @parameters -ReviewerType $Type -ForceTwoStage -WhatIf:$Preview -Confirm:$false -WarningAction SilentlyContinue)
		$result = $results[-1]
		$result.Stages | Should -Be 2
		if ($Preview) {
			$result.Status | Should -Be 'Would be created'
			Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
			Test-Path (Join-Path $script:testScriptRoot 'logs') | Should -BeFalse
		}
		else {
			$result.Status | Should -Be 'Created'
			$body = $script:fixture.Bodies[0]
			$body.stageSettings | Should -HaveCount 2
			$body.stageSettings[0].reviewers[0].query | Should -Be $(if ($Type -eq 'ResourceOwner') { "/users/$script:groupOwnerId" } else { './manager' })
			$body.stageSettings[1].reviewers[0].query | Should -Be "/users/$script:ownerId"
			$body.stageSettings[1].dependsOn | Should -Be @('1')
			$body.settings.instanceDurationInDays | Should -Be 10
			([datetime]$body.settings.recurrence.range.endDate).Date | Should -Be ((Get-Date).Date.AddDays(10))
			$entries = Get-LogEntries
			$entries[0].ForceTwoStage | Should -BeTrue
			$entries[0].Syntax | Should -Match '-ForceTwoStage:\$true'
			@($entries | Where-Object Event -in @('GroupResult', 'ApplicationResult'))[-1].Result.Stages | Should -Be 2
		}
	}

	It 'preserves default single-stage behavior with ForceTwoStage false' {
		$script:fixture.Count = 22
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -ForceTwoStage:$false -Confirm:$false
		$result.Stages | Should -Be 1
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
	}

	It 'retains identical owner sets as separate stages only when forced' {
		$script:fixture.Count = 22
		$script:fixture.Owners = @([pscustomobject]@{
			Id = $script:ownerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' }
		})
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -ForceTwoStage -Confirm:$false
		$result.Stages | Should -Be 2
		$body = $script:fixture.Bodies[0]
		$body.stageSettings[0].reviewers[0].query | Should -Be "/users/$script:ownerId"
		$body.stageSettings[1].reviewers[0].query | Should -Be "/users/$script:ownerId"
	}

	It 'requires owners or an explicit fallback for a forced ResourceOwner stage' {
		$script:fixture.Count = 22
		$script:fixture.Owners = @()
		{ & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -ForceTwoStage -Confirm:$false } |
			Should -Throw '*ForceTwoStage requires user owners or FallbackReviewerUserId*'
		(Get-LogEntries)[1].Result.Status | Should -Be 'Failed'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'uses a fallback for an ownerless forced review even if it matches application owners' {
		$script:fixture.Count = 22
		$script:fixture.Owners = @()
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -ForceTwoStage -FallbackReviewerUserId $script:ownerId -Confirm:$false -WarningAction SilentlyContinue
		$result.Stages | Should -Be 2
		$script:fixture.Bodies[0].stageSettings[0].reviewers[0].query | Should -Be "/users/$script:ownerId"
	}

	It 'retains self-review at 500 members with the override' {
		$script:fixture.Count = 500
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -ForceTwoStage -Confirm:$false
		$result.Stages | Should -Be 3
		$script:fixture.Bodies[0].stageSettings[0].reviewers | Should -HaveCount 0
	}

	It 'still skips empty groups when forcing two stages' {
		$script:fixture.Count = 0
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -ForceTwoStage -Confirm:$false
		$result.Status | Should -Be 'Skipped'
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'rejects a contradictory ApplicationOwner-only override' {
		{ & $script:testScript -GroupId $script:groupId -ReviewerType ApplicationOwner -ForceTwoStage } |
			Should -Throw '*ForceTwoStage requires ReviewerType LineManager or ResourceOwner*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'validates recurrence against both forced stages' {
		$script:fixture.Count = 22
		{ & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -ForceTwoStage -Frequency Monthly -EndDate ((Get-Date).Date.AddDays(9)) -Confirm:$false } |
			Should -Throw '*at least 10 days*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'uses application-owner-only mode regardless of member count' -TestCases @(
		@{ Count = 1 }, @{ Count = 500 }
	) {
		param($Count)
		$script:fixture.Count = $Count
		$null = & $script:testScript -GroupId $script:groupId -ReviewerType ApplicationOwner -Confirm:$false
		$body = $script:fixture.Bodies[0]
		$body.reviewers[0].query | Should -Be "/users/$script:ownerId"
		$body.ContainsKey('stageSettings') | Should -BeFalse
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
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
		$result.Reason | Should -Be 'No members; no access review needed'
		$result.MemberCount | Should -Be 0
		$result.Stages | Should -BeNullOrEmpty
		$script:fixture.HostMessages | Should -Contain 'The group Demo group has no members; no access review needed.'
		if (-not $Preview) {
			(Get-LogEntries)[1].Result.Status | Should -Be 'Skipped'
			(Get-LogEntries)[-1].Skipped | Should -Be 1
		}
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'applies stage thresholds to <Count> members in <Mode> mode' -TestCases @(
		@{ Mode = 'Single'; Count = 10; Stages = 1 },
		@{ Mode = 'Single'; Count = 149; Stages = 1 },
		@{ Mode = 'Single'; Count = 150; Stages = 2 },
		@{ Mode = 'Single'; Count = 499; Stages = 2 },
		@{ Mode = 'Single'; Count = 500; Stages = 3 },
		@{ Mode = 'Application'; Count = 10; Stages = 1 },
		@{ Mode = 'Application'; Count = 149; Stages = 1 },
		@{ Mode = 'Application'; Count = 150; Stages = 2 },
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

	It 'uses per-row CSV owners and falls back to the parameter for blank cells' {
		$csv = Join-Path $script:testScriptRoot 'owners.csv'
		$second = '44444444-4444-4444-4444-444444444444'
		$third = '55555555-5555-5555-5555-555555555555'
		"GroupId,AppOwnerIds,ReviewName`n$script:groupId,$second;$third;$second,First`n$second,,Second" | Set-Content $csv
		$null = & $script:testScript -CsvPath $csv -ReviewerType ResourceOwner -Confirm:$false
		$script:fixture.Bodies[0].stageSettings[1].reviewers.query | Should -Be @("/users/$second", "/users/$third")
		$script:fixture.Bodies[1].stageSettings[1].reviewers[0].query | Should -Be "/users/$script:ownerId"
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
		$null = & $script:testScript -CsvPath $csv -AppOwnerIds @() -ReviewerType ResourceOwner -Confirm:$false
		$script:fixture.Bodies[0].stageSettings[1].reviewers[0].query | Should -Be "/users/$script:ownerId"
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

	It 'skips an ownerless stage while preserving sign-off at <Count> members (service principal only: <ServicePrincipalOnly>)' -TestCases @(
		@{ Count = 150; ServicePrincipalOnly = $false },
		@{ Count = 500; ServicePrincipalOnly = $false },
		@{ Count = 150; ServicePrincipalOnly = $true },
		@{ Count = 500; ServicePrincipalOnly = $true }
	) {
		param($Count, $ServicePrincipalOnly)
		$script:fixture.Count = $Count
		$script:fixture.Owners = @($script:fixture.Owners | Where-Object {
			$ServicePrincipalOnly -and $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.servicePrincipal'
		})
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false -WarningAction SilentlyContinue
		$body = $script:fixture.Bodies[0]
		if ($Count -lt 500) {
			$result.Stages | Should -Be 1
			$result.StageNames | Should -Be 'Application owners'
			$body.ContainsKey('stageSettings') | Should -BeFalse
			$body.reviewers.query | Should -Be "/users/$script:ownerId"
		}
		else {
			$result.Stages | Should -Be 2
			$result.StageNames | Should -Be 'Self-review -> Application owners'
			$body.stageSettings[0].reviewers | Should -HaveCount 0
			$body.stageSettings[1].reviewers.query | Should -Be "/users/$script:ownerId"
			$body.stageSettings[1].dependsOn | Should -Be @('1')
		}
		$body.settings.instanceDurationInDays | Should -Be ($result.Stages * 5)
		([datetime]$body.settings.recurrence.range.endDate).Date | Should -Be ((Get-Date).Date.AddDays($result.Stages * 5))
		(Get-LogEntries)[1].Result.StageNames | Should -Be $result.StageNames
		($body | ConvertTo-Json -Depth 20) | Should -Not -Match '/users/00000000-0000-0000-0000-000000000000'
	}

	It 'uses an explicit fallback before application owners at <Count> members' -TestCases @(
		@{ Count = 150 }, @{ Count = 500 }
	) {
		param($Count)
		$script:fixture.Count = $Count
		$script:fixture.Owners = @()
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -FallbackReviewerUserId $script:reviewId -Confirm:$false -WarningAction SilentlyContinue
		$body = $script:fixture.Bodies[0]
		$body.stageSettings[-2].reviewers.query | Should -Be "/users/$script:reviewId"
		$body.stageSettings[-1].reviewers.query | Should -Be "/users/$script:ownerId"
		$result.Stages | Should -Be $(if ($Count -ge 500) { 3 } else { 2 })
		(Get-LogEntries)[0].Syntax | Should -Match "-FallbackReviewerUserId '$script:reviewId'"
	}

	It 'uses per-row application owners as manager fallback instead of the zero GUID' {
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
		$script:fixture.HostMessages | Should -Contain 'Total transitive members in reviews that would be created: 150'
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
		$null = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false -WarningAction SilentlyContinue
		$script:fixture.Bodies[0].settings.autoApplyDecisionsEnabled | Should -BeFalse
		$script:fixture.Bodies[0].settings.ContainsKey('applyActions') | Should -BeFalse
	}

	It 'discovers assigned groups once and retains application names and sign-off stages' {
		$script:fixture.Assignments += $script:fixture.Assignments[0]
		$script:fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -ReviewerType ResourceOwner -Confirm:$false)
		$results | Should -HaveCount 2
		$results[0].ReviewName | Should -Be 'AR Demo app - Demo group'
		$results[0].Status | Should -Be 'Created'
		$script:fixture.Bodies[0].stageSettings | Should -HaveCount 2
		$script:fixture.Bodies[0].stageSettings[1].reviewers[0].query | Should -Be "/users/$script:ownerId"
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
		$entries[1].Result.Status | Should -Be 'Skipped'
		$entries[2].Skipped | Should -Be 1
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
		$entries[1].Event | Should -Be 'ApplicationResult'
		$entries[1].Result.SubjectCount | Should -Be 2
		$entries[1].Result.AssignmentIds | Should -Be $result.AssignmentIds
		$script:fixture.HostMessages | Should -Contain 'Total transitive members in created reviews: 0'
		$script:fixture.HostMessages | Should -Contain 'Directly assigned users in created reviews: 2'
		Should -Invoke Get-MgGroup -Times 0 -Exactly
		Should -Invoke Get-MgGroupTransitiveMemberCount -Times 0 -Exactly
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
		$results[-1].StageNames | Should -Be 'Resource owners (service-principal owners) -> Application owners'
		$script:fixture.Bodies[0].stageSettings[0].reviewers.query | Should -Be @("/users/$script:groupOwnerId")
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
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

	It 'does not let a failed group hide direct user assignments' {
		$script:fixture.Assignments += [pscustomobject]@{ PrincipalType = 'User'; PrincipalId = $script:ownerId }
		Mock Get-MgGroup { throw 'Group unavailable.' }
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -Confirm:$false)
		$results.Status | Should -Be @('Failed', 'Created')
		$results[-1].TargetType | Should -Be 'ApplicationDirectUsers'
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
		@{ Type = 'ApplicationOwner'; Stages = 1 }
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
		(Get-LogEntries)[1].Total | Should -Be 0
		$script:fixture.HostMessages | Should -Contain '0 review target(s): 0 created, 0 skipped, 0 failed.'
		$script:fixture.HostMessages | Should -Contain 'Total transitive members in created reviews: 0'
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

	It 'continues application discovery after a group fails' {
		$second = '44444444-4444-4444-4444-444444444444'
		$script:fixture.Assignments += [pscustomobject]@{ PrincipalType = 'Group'; PrincipalId = $second }
		Mock Get-MgGroup { throw 'Group unavailable.' } -ParameterFilter { "$GroupId" -eq '11111111-1111-1111-1111-111111111111' }
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -Confirm:$false)
		$results | Should -HaveCount 2
		$results[0].Status | Should -Be 'Failed'
		$results[1].Status | Should -Be 'Created'
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
		$result = & $script:testScript -CsvPath $csv -AppOwnerIds @() -ReviewerType ResourceOwner -Confirm:$false
		$result.Status | Should -Be 'Created'
		$result.ReviewName | Should -Be 'Finance, reporting; review'
		$script:fixture.Bodies[0].stageSettings[1].reviewers.query |
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
		$reviewers = if ($Type -eq 'ApplicationOwner') { $body.reviewers } else { $body.stageSettings[-1].reviewers }
		$reviewers.query | Should -Be @("/users/$script:ownerId", "/users/$script:reviewId")
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
