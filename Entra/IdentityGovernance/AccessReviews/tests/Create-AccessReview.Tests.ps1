BeforeAll {
	Import-Module Microsoft.Graph.Authentication
	Import-Module Microsoft.Graph.Applications
	Import-Module Microsoft.Graph.Groups
	Import-Module Microsoft.Graph.Identity.Governance
	$script:groupId = '11111111-1111-1111-1111-111111111111'
	$script:ownerId = '22222222-2222-2222-2222-222222222222'
	$script:reviewId = '33333333-3333-3333-3333-333333333333'
	function Get-ReviewLog {
		param([string]$ScriptName)
		@(Get-Content (Join-Path $script:testRoot "logs/$ScriptName.jsonl") |
			ForEach-Object { $_ | ConvertFrom-Json })
	}
}

Describe 'Reviewer selection and logging across review scripts' {
	BeforeEach {
		$script:testRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString())
		$null = New-Item -ItemType Directory -Path $script:testRoot
		foreach ($name in 'Create-AccessReview.ps1', 'Create-AccessReviewAppGroups.ps1', 'AccessReview.Logging.ps1') {
			Copy-Item (Join-Path $PSScriptRoot "../$name") $script:testRoot
		}
		$script:singleScript = Join-Path $script:testRoot 'Create-AccessReview.ps1'
		$script:appScript = Join-Path $script:testRoot 'Create-AccessReviewAppGroups.ps1'
		$script:fixture = @{
			Count = 1
			Synced = $false
			Dynamic = $false
			Owners = @([pscustomobject]@{
				Id = $script:ownerId
				AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.user' }
			})
			OwnerFailure = $false
			CreateFailure = $false
			ExistingNames = @()
			Assignments = @([pscustomobject]@{ PrincipalType = 'Group'; PrincipalId = $script:groupId })
			Bodies = [System.Collections.Generic.List[object]]::new()
			ReviewId = $script:reviewId
		}
		$fixture = $script:fixture
		Mock Get-MgContext { @{ Scopes = @('AccessReview.ReadWrite.All') } }
		Mock Get-MgGroup ({
			param($GroupId)
			[pscustomobject]@{
				Id = "$GroupId"
				DisplayName = 'Demo group'
				OnPremisesSyncEnabled = $fixture.Synced
				GroupTypes = $(if ($fixture.Dynamic) { @('DynamicMembership') } else { @() })
			}
		}.GetNewClosure())
		Mock Get-MgGroupTransitiveMemberCount ({ $fixture.Count }.GetNewClosure())
		Mock Get-MgGroupOwner ({
			if ($fixture.OwnerFailure) { throw 'Owner read forbidden.' }
			$fixture.Owners
		}.GetNewClosure())
		Mock New-MgIdentityGovernanceAccessReviewDefinition ({
			param($BodyParameter)
			if ($fixture.CreateFailure) { throw 'Review creation forbidden.' }
			$fixture.Bodies.Add($BodyParameter)
			[pscustomobject]@{ Id = $fixture.ReviewId }
		}.GetNewClosure())
		Mock Get-MgServicePrincipal { [pscustomobject]@{ DisplayName = 'Demo app' } }
		Mock Get-MgServicePrincipalAppRoleAssignedTo ({ $fixture.Assignments }.GetNewClosure())
		Mock Get-MgIdentityGovernanceAccessReviewDefinition ({
			foreach ($name in $fixture.ExistingNames) { [pscustomobject]@{ DisplayName = $name } }
		}.GetNewClosure())
	}

	It 'preserves single-review application-owner defaults' {
		$result = & $script:singleScript -GroupId $script:groupId -Confirm:$false
		$result.ReviewerType | Should -Be 'ApplicationOwner'
		$result.ReviewId | Should -Be $script:reviewId
		$script:fixture.Bodies[0].reviewers[0].query | Should -Be '/users/00000000-0000-0000-0000-000000000000'
		Should -Invoke Get-MgGroupOwner -Times 0 -Exactly
	}

	It 'uses managers in single reviews' {
		$null = & $script:singleScript -GroupId $script:groupId -ReviewerType LineManager -Confirm:$false
		$body = $script:fixture.Bodies[0]
		$body.reviewers[0].query | Should -Be './manager'
		$body.reviewers[0].queryRoot | Should -Be 'decisions'
		$body.fallbackReviewers[0].query | Should -Be '/users/00000000-0000-0000-0000-000000000000'
	}

	It 'resolves human group owners in single reviews' {
		$script:fixture.Owners += [pscustomobject]@{
			Id = 'service-principal'; AdditionalProperties = @{ '@odata.type' = '#microsoft.graph.servicePrincipal' }
		}
		$null = & $script:singleScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false
		$script:fixture.Bodies[0].reviewers | Should -HaveCount 1
		$script:fixture.Bodies[0].reviewers[0].query | Should -Be "/users/$script:ownerId"
	}

	It 'falls back for a single review without user owners' {
		$script:fixture.Owners = @()
		$null = & $script:singleScript -GroupId $script:groupId -ReviewerType ResourceOwner -Confirm:$false -WarningAction SilentlyContinue
		$script:fixture.Bodies[0].reviewers[0].query | Should -Be '/users/00000000-0000-0000-0000-000000000000'
	}

	It 'logs single-review failures before rethrowing' {
		$script:fixture.OwnerFailure = $true
		{ & $script:singleScript -GroupId $script:groupId -ReviewerType ResourceOwner } | Should -Throw '*Owner read forbidden*'
		$entries = Get-ReviewLog 'Create-AccessReview'
		$entries[1].Result.Status | Should -Be 'Failed'
		$entries[1].Result.Reason | Should -Match 'Owner read forbidden'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'logs single-review CSV failures and final totals' {
		$csv = Join-Path $script:testRoot 'groups.csv'
		"GroupId`n$script:groupId`n44444444-4444-4444-4444-444444444444" | Set-Content $csv
		$script:fixture.CreateFailure = $true
		$null = & $script:singleScript -CsvPath $csv -ReviewerType LineManager -Confirm:$false
		$entries = Get-ReviewLog 'Create-AccessReview'
		$entries.Count | Should -Be 4
		$entries[3].Failed | Should -Be 2
	}

	It 'preserves group-owner defaults in app-group reviews' {
		$null = & $script:appScript -ServicePrincipalId $script:groupId -Confirm:$false
		$script:fixture.Bodies[0].reviewers[0].query | Should -Be "/groups/$script:groupId/owners"
		$entries = Get-ReviewLog 'Create-AccessReviewAppGroups'
		$entries[1].Result.ReviewId | Should -Be $script:reviewId
		$entries[1].Result.ReviewerType | Should -Be 'ResourceOwner'
		$entries[2].Created | Should -Be 1
	}

	It 'selects managers in <Count>-member app-group reviews while retaining self-review thresholds' -TestCases @(
		@{ Count = 150 }, @{ Count = 151 }
	) {
		param($Count)
		$script:fixture.Count = $Count
		$null = & $script:appScript -ServicePrincipalId $script:groupId -ReviewerType LineManager -Confirm:$false
		$body = $script:fixture.Bodies[0]
		$reviewers = @(if ($Count -gt 150) { $body.stageSettings[1].reviewers } else { $body.reviewers })
		$reviewers[0].query | Should -Be './manager'
		$reviewers[0].queryRoot | Should -Be 'decisions'
		if ($Count -gt 150) {
			$body.stageSettings | Should -HaveCount 2
			$body.stageSettings[0].reviewers | Should -HaveCount 0
		}
	}

	It 'logs dynamic and existing app groups as skipped' -TestCases @(
		@{ Kind = 'dynamic' }, @{ Kind = 'existing' }
	) {
		param($Kind)
		if ($Kind -eq 'dynamic') { $script:fixture.Dynamic = $true }
		else { $script:fixture.ExistingNames = @('AR Demo app - Demo group') }
		$null = & $script:appScript -ServicePrincipalId $script:groupId -ReviewerType LineManager
		$entries = Get-ReviewLog 'Create-AccessReviewAppGroups'
		$entries[1].Result.Status | Should -Be 'Skipped'
		$entries[2].Skipped | Should -Be 1
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'logs an empty app-group run' {
		$script:fixture.Assignments = @()
		$null = & $script:appScript -ServicePrincipalId $script:groupId -WarningAction SilentlyContinue
		$entries = Get-ReviewLog 'Create-AccessReviewAppGroups'
		$entries.Count | Should -Be 2
		$entries[1].Total | Should -Be 0
	}

	It 'logs an app-group creation failure' {
		$script:fixture.CreateFailure = $true
		$null = & $script:appScript -ServicePrincipalId $script:groupId -Confirm:$false
		$entries = Get-ReviewLog 'Create-AccessReviewAppGroups'
		$entries[1].Result.Status | Should -Be 'Failed'
		$entries[2].Failed | Should -Be 1
	}

	It 'logs WhatIf syntax without Graph writes in <ScriptName>' -TestCases @(
		@{ ScriptName = 'Create-AccessReview'; Parameter = 'GroupId' },
		@{ ScriptName = 'Create-AccessReviewAppGroups'; Parameter = 'ServicePrincipalId' }
	) {
		param($ScriptName, $Parameter)
		$path = Join-Path $script:testRoot "$ScriptName.ps1"
		$parameters = @{ $Parameter = $script:groupId; ReviewerType = 'LineManager'; WhatIf = $true }
		$null = & $path @parameters
		$entries = Get-ReviewLog $ScriptName
		$entries[0].Syntax | Should -Match "-ReviewerType 'LineManager'"
		$entries[0].Syntax | Should -Match '-WhatIf:\$true'
		$entries[1].Result.Status | Should -Be 'Would be created'
		$entries[2].Previewed | Should -Be 1
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'stops <ScriptName> before creation when logging fails' -TestCases @(
		@{ ScriptName = 'Create-AccessReview'; Parameter = 'GroupId' },
		@{ ScriptName = 'Create-AccessReviewAppGroups'; Parameter = 'ServicePrincipalId' }
	) {
		param($ScriptName, $Parameter)
		Set-Content (Join-Path $script:testRoot 'logs') 'A file blocks logging.'
		$parameters = @{ $Parameter = $script:groupId }
		$path = Join-Path $script:testRoot "$ScriptName.ps1"
		{ & $path @parameters } | Should -Throw '*Unable to write creation log*'
		Should -Invoke New-MgIdentityGovernanceAccessReviewDefinition -Times 0 -Exactly
	}

	It 'retains report-only behavior in <ScriptName>' -TestCases @(
		@{ ScriptName = 'Create-AccessReview'; Parameter = 'GroupId' },
		@{ ScriptName = 'Create-AccessReviewAppGroups'; Parameter = 'ServicePrincipalId' }
	) {
		param($ScriptName, $Parameter)
		$script:fixture.Synced = $true
		$path = Join-Path $script:testRoot "$ScriptName.ps1"
		$parameters = @{ $Parameter = $script:groupId; ReviewerType = 'LineManager'; Confirm = $false; WarningAction = 'SilentlyContinue' }
		$null = & $path @parameters
		$script:fixture.Bodies[0].settings.autoApplyDecisionsEnabled | Should -BeFalse
		$script:fixture.Bodies[0].settings.ContainsKey('applyActions') | Should -BeFalse
	}
}
