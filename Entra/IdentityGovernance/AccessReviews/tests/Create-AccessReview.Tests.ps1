BeforeAll {
	Import-Module Microsoft.Graph.Authentication
	Import-Module Microsoft.Graph.Groups
	Import-Module Microsoft.Graph.Identity.Governance
	Import-Module Microsoft.Graph.Applications
	$script:sourcePath = Join-Path $PSScriptRoot '../Create-AccessReview.ps1'
	$script:groupId = '11111111-1111-1111-1111-111111111111'
	$script:ownerId = '22222222-2222-2222-2222-222222222222'
	$script:reviewId = '33333333-3333-3333-3333-333333333333'
	function Get-LogEntries {
		@(Get-Content (Join-Path $script:testScriptRoot 'logs/Create-AccessReview.jsonl') |
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
				[pscustomobject]@{ Id = $script:ownerId; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' } },
				[pscustomobject]@{ Id = 'service-principal'; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.servicePrincipal' } }
			)
			OwnerFailure = $false
			CreateFailure = $false
			Bodies = [System.Collections.Generic.List[object]]::new()
			ReviewId = $script:reviewId
			Dynamic = $false
			Assignments = @([pscustomobject]@{ PrincipalType = 'Group'; PrincipalId = $script:groupId })
			ExistingNames = @()
		}
		$fixture = $script:fixture
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

	It 'selects only human group owners and retains application owners as the final stage' {
		$result = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false
		$result.StageNames | Should -Be 'Resource owners (group owners) -> Application owners'
		$body = $script:fixture.Bodies[0]
		$body.stageSettings[0].reviewers.Count | Should -Be 1
		$body.stageSettings[0].reviewers[0].query | Should -Be "/users/$script:ownerId"
		$body.stageSettings[1].dependsOn | Should -Be @('1')
	}

	It 'preserves self-review in three-stage owner reviews' {
		$script:fixture.Count = 500
		$null = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false
		$body = $script:fixture.Bodies[0]
		$body.stageSettings.Count | Should -Be 3
		$body.stageSettings[0].reviewers | Should -HaveCount 0
		$body.stageSettings[1].reviewers[0].query | Should -Be "/users/$script:ownerId"
		$body.settings.instanceDurationInDays | Should -Be 15
	}

	It 'always includes selected reviewers before application-owner sign-off for small groups' -TestCases @(
		@{ Type = 'LineManager'; Count = 0 }, @{ Type = 'ResourceOwner'; Count = 1 },
		@{ Type = 'LineManager'; Count = 49 }, @{ Type = 'ResourceOwner'; Count = 50 },
		@{ Type = 'LineManager'; Count = 149 }, @{ Type = 'ResourceOwner'; Count = 499 }
	) {
		param($Type, $Count)
		$script:fixture.Count = $Count
		$null = & $script:testScript -GroupId $script:groupId -ReviewerType $Type -Confirm:$false
		$body = $script:fixture.Bodies[0]
		$body.stageSettings | Should -HaveCount 2
		$body.stageSettings[1].reviewers[0].query | Should -Be "/users/$script:ownerId"
		$body.stageSettings[0].reviewers[0].query | Should -Be $(if ($Type -eq 'ResourceOwner') { "/users/$script:ownerId" } else { './manager' })
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
		$null = & $script:testScript -GroupId $script:groupId -AppOwnerIds $ids -WhatIf
		(Get-LogEntries)[0].Syntax | Should -Match "-AppOwnerIds @\('$script:ownerId', '$script:reviewId'\)"
	}

	It 'uses fallback when no human group owners exist' {
		$script:fixture.Owners = @()
		$null = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false -WarningAction SilentlyContinue
		$body = $script:fixture.Bodies[0]
		$body.stageSettings[0].reviewers[0].query | Should -Be '/users/00000000-0000-0000-0000-000000000000'
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

	It 'logs preview syntax with escaped quotes and no Graph creation' {
		$result = & $script:testScript -GroupId $script:groupId -ReviewName "Review O'Brien" -ReviewerType ResourceOwner -WhatIf
		$result.Status | Should -Be 'Would be created'
		$entries = Get-LogEntries
		$entries.Count | Should -Be 3
		$entries[0].Syntax | Should -Match "-ReviewName 'Review O''Brien'"
		$entries[0].Syntax | Should -Match '-WhatIf:\$true'
		$entries[0].Syntax | Should -Match "-ReviewerType 'ResourceOwner'"
		$entries[1].Result.Status | Should -Be 'Would be created'
		$entries[2].Previewed | Should -Be 1
		$entries[2].Created | Should -Be 0
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'appends independent runs with timestamps, invocation and created review IDs' {
		$null = & $script:testScript -GroupId $script:groupId -Confirm:$false
		$null = & $script:testScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false
		$entries = Get-LogEntries
		$entries.Count | Should -Be 6
		$entries[0].RunId | Should -Not -Be $entries[3].RunId
		([datetime]$entries[0].Timestamp).Kind | Should -Be ([System.DateTimeKind]::Utc)
		$entries[1].RunId | Should -Be $entries[0].RunId
		$entries[1].Result.ReviewId | Should -Be $script:reviewId
		$entries[2].Created | Should -Be 1
		$entries[0].Syntax | Should -Match '-Confirm:\$false'
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
		$results | Should -HaveCount 1
		$results[0].ReviewName | Should -Be 'AR Demo app - Demo group'
		$results[0].Status | Should -Be 'Created'
		$script:fixture.Bodies[0].stageSettings | Should -HaveCount 2
		$script:fixture.Bodies[0].stageSettings[1].reviewers[0].query | Should -Be "/users/$script:ownerId"
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

	It 'logs empty application discovery without creating a review' {
		$script:fixture.Assignments = @()
		$results = @(& $script:testScript -ServicePrincipalId $script:groupId -WarningAction SilentlyContinue)
		$results | Should -HaveCount 0
		(Get-LogEntries)[1].Total | Should -Be 0
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
}
