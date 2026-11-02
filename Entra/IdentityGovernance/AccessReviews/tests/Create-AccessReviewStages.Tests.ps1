BeforeAll {
	Import-Module Microsoft.Graph.Authentication
	Import-Module Microsoft.Graph.Groups
	Import-Module Microsoft.Graph.Identity.Governance
	$script:sourcePath = Join-Path $PSScriptRoot '../Create-AccessReviewStages.ps1'
	$script:groupId = '11111111-1111-1111-1111-111111111111'
	$script:ownerId = '22222222-2222-2222-2222-222222222222'
	$script:reviewId = '33333333-3333-3333-3333-333333333333'
	function Get-LogEntries {
		@(Get-Content (Join-Path $script:testScriptRoot 'logs/Create-AccessReviewStages.jsonl') |
			ForEach-Object { $_ | ConvertFrom-Json })
	}
}

Describe 'Staged review selection and local logging' {
	BeforeEach {
		$script:testScriptRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString())
		$null = New-Item -ItemType Directory -Path $script:testScriptRoot
		$script:testScript = Join-Path $script:testScriptRoot 'Create-AccessReviewStages.ps1'
		Copy-Item $script:sourcePath $script:testScript
		Copy-Item (Join-Path $PSScriptRoot '../AccessReview.Logging.ps1') $script:testScriptRoot
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
		}
		$fixture = $script:fixture
		Mock Get-MgContext { @{ Scopes = @('AccessReview.ReadWrite.All', 'Application.Read.All') } }
		Mock Get-MgGroup ({
			param($GroupId)
			[pscustomobject]@{ Id = "$GroupId"; DisplayName = 'Demo group'; OnPremisesSyncEnabled = $fixture.Synced }
		}.GetNewClosure())
		Mock Get-MgGroupTransitiveMemberCount ({ $fixture.Count }.GetNewClosure())
		Mock Get-MgGroupOwner ({
			if ($fixture.OwnerFailure) { throw 'Owner read forbidden.' }
			$fixture.Owners
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
		$body.stageSettings[1].reviewers[0].query | Should -Be '/users/00000000-0000-0000-0000-000000000000'
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

	It 'preserves single-stage application-owner behavior for either selection' -TestCases @(
		@{ Type = 'LineManager' }, @{ Type = 'ResourceOwner' }
	) {
		param($Type)
		$script:fixture.Count = 149
		$null = & $script:testScript -GroupId $script:groupId -ReviewerType $Type -Confirm:$false
		$body = $script:fixture.Bodies[0]
		$body.reviewers[0].query | Should -Be '/users/00000000-0000-0000-0000-000000000000'
		$body.ContainsKey('stageSettings') | Should -BeFalse
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
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
}
