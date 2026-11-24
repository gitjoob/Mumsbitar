BeforeAll {
	Import-Module Microsoft.Graph.Authentication
	. (Join-Path $PSScriptRoot '../AccessPackage.Coverage.ps1')
	$script:groupA = '11111111-1111-1111-1111-111111111111'
	$script:groupB = '22222222-2222-2222-2222-222222222222'
	$script:packageId = '33333333-3333-3333-3333-333333333333'
}

Describe 'Discovery package coverage' {
	BeforeEach {
		$script:fixture = @{
			Policies = @(); Assignments = @(); Requests = @()
			Scopes = @()
			RoleAssignments = @(@{ directoryScopeId = '/'; roleDefinitionId = '45d8d3c5-c802-45c6-b32a-1d70b5e1e86e' })
		}
		$fixture = $script:fixture
		Mock Get-MgContext { @{ AuthType = 'Delegated' } }
		Mock Invoke-MgGraphRequest ({
			param($Uri)
			if ($Uri -like '/v1.0/me?*') { return @{ id = '44444444-4444-4444-4444-444444444444' } }
			if ($Uri -like '*/roleAssignments?*') { return @{ value = $fixture.RoleAssignments } }
			if ($Uri -like '*/accessPackages') { return @{ value = @(@{ id = '33333333-3333-3333-3333-333333333333'; displayName = 'Existing' }) } }
			if ($Uri -like '*/accessPackages/*') { return @{ resourceRoleScopes = $fixture.Scopes } }
			if ($Uri -like '*/assignmentPolicies?*') { return @{ value = $fixture.Policies } }
			if ($Uri -like '*/assignments?*') { return @{ value = $fixture.Assignments } }
			if ($Uri -like '*/assignmentRequests?*') { return @{ value = $fixture.Requests } }
			throw "Unexpected request $Uri"
		}.GetNewClosure())
	}

	It 'skips an exact resource set with an active policy' {
		$script:fixture.Scopes = @(
			@{ role = @{ originSystem = 'AadGroup'; originId = "Member_$script:groupA" }; scope = @{ originSystem = 'AadGroup'; originId = $script:groupA; isRootScope = $true } }
			@{ role = @{ originSystem = 'AadGroup'; originId = "Member_$script:groupB" }; scope = @{ originSystem = 'AadGroup'; originId = $script:groupB; isRootScope = $true } }
		)
		$script:fixture.Policies = @(@{ allowedTargetScope = 'allMemberUsers'; expiration = @{ type = 'afterDuration'; duration = 'P365D' }; requestorSettings = @{ enableTargetsToSelfAddAccess = $true } })
		$inventory = @(Get-PackageCoverageInventory)
		$result = Get-PackageCoverageResult -GroupIds @($script:groupB, $script:groupA) -Inventory $inventory
		$result.Status | Should -Be 'Skipped'
		$result.Reason | Should -Match $script:packageId
	}

	It 'distinguishes currently effective and expired assignments and fails closed on unknown schedules' {
		$script:fixture.Assignments = @(@{
			state = 'delivered'
			schedule = @{ startDateTime = [datetimeoffset]::UtcNow.AddDays(-1).ToString('o'); expiration = @{ type = 'afterDuration'; duration = 'P365D' } }
		})
		Get-PackageCoverageActivity -PackageId $script:packageId | Should -Be 'Active'
		$script:fixture.Assignments[0].schedule.startDateTime = [datetimeoffset]::UtcNow.AddDays(-400).ToString('o')
		Get-PackageCoverageActivity -PackageId $script:packageId | Should -Be 'Unknown'
		$script:fixture.Assignments[0].state = 'expired'
		Get-PackageCoverageActivity -PackageId $script:packageId | Should -Be 'Inactive'
		$script:fixture.Assignments[0].state = 'delivered'
		$script:fixture.Assignments[0].schedule = $null
		Get-PackageCoverageActivity -PackageId $script:packageId | Should -Be 'Unknown'
	}

	It 'requires manual review for extra roles or ambiguous policies' {
		$script:fixture.Scopes = @(
			@{ role = @{ originSystem = 'AadGroup'; originId = "Member_$script:groupA" }; scope = @{ originSystem = 'AadGroup'; originId = $script:groupA; isRootScope = $true } }
			@{ role = @{ originSystem = 'AadGroup'; originId = "Owner_$script:groupA" }; scope = @{ originSystem = 'AadGroup'; originId = $script:groupA; isRootScope = $true } }
		)
		$inventory = @(Get-PackageCoverageInventory)
		(Get-PackageCoverageResult -GroupIds @($script:groupA) -Inventory $inventory).Status | Should -Be 'Manual review'
		$script:fixture.Scopes = @($script:fixture.Scopes[0])
		$script:fixture.Policies = @(@{ allowedTargetScope = 'specificDirectoryUsers' })
		$inventory = @(Get-PackageCoverageInventory)
		(Get-PackageCoverageResult -GroupIds @($script:groupA) -Inventory $inventory).Status | Should -Be 'Manual review'
	}

	It 'rejects catalog-scoped or incomplete visibility instead of assuming no overlap' {
		$script:fixture.RoleAssignments = @(@{ directoryScopeId = '/some-scope'; roleDefinitionId = '45d8d3c5-c802-45c6-b32a-1d70b5e1e86e' })
		{ Get-PackageCoverageInventory } | Should -Throw '*cannot prove complete visibility*'
		$script:fixture.RoleAssignments = @(@{ directoryScopeId = '/'; roleDefinitionId = '45d8d3c5-c802-45c6-b32a-1d70b5e1e86e' })
		$script:fixture.Scopes = @(@{ role = @{ originSystem = 'AadGroup' }; scope = @{} })
		{ Get-PackageCoverageInventory } | Should -Throw '*incomplete resource role or scope*'
	}

	It 'allows a shared subset while retaining exact-match and extra-role safeguards' {
		$inventory = @([pscustomobject]@{
			Id = $script:packageId; DisplayName = 'Existing'; HasExtra = $false
			GroupIds = @($script:groupA, $script:groupB); MemberGroupIds = @($script:groupA, $script:groupB)
		})
		(Get-PackageCoverageResult -GroupIds @($script:groupA) -Inventory $inventory).Status | Should -Be 'Manual review'
		(Get-PackageCoverageResult -GroupIds @($script:groupA) -Inventory $inventory -AllowSharedGroups).Status | Should -Be 'Ready'
		$inventory[0].HasExtra = $true
		(Get-PackageCoverageResult -GroupIds @($script:groupA, $script:groupB) -Inventory $inventory -AllowSharedGroups).Status | Should -Be 'Manual review'
		$inventory[0].MemberGroupIds = @($script:groupA)
		(Get-PackageCoverageResult -GroupIds @($script:groupA, $script:groupB) -Inventory $inventory -AllowSharedGroups).Status | Should -Be 'Manual review'
	}

	It 'uses app-only tenant-wide application permission <Permission> to check package coverage' -TestCases @(
		@{ Permission = 'EntitlementManagement.Read.All' }
		@{ Permission = 'EntitlementManagement.ReadWrite.All' }
	) {
		param($Permission)
		Mock Get-MgContext { @{ AuthType = 'AppOnly'; Scopes = @($Permission) } }
		$script:fixture.Scopes = @(
			@{ role = @{ originSystem = 'AadGroup'; originId = "Member_$script:groupA" }; scope = @{ originSystem = 'AadGroup'; originId = $script:groupA; isRootScope = $true } }
		)
		$script:fixture.Policies = @(@{ allowedTargetScope = 'allMemberUsers'; expiration = @{ type = 'afterDuration'; duration = 'P365D' }; requestorSettings = @{ enableTargetsToSelfAddAccess = $true } })
		$inventory = @(Get-PackageCoverageInventory)
		$inventory | Should -HaveCount 1
		$inventory[0].Id | Should -Be $script:packageId
		(Get-PackageCoverageResult -GroupIds @($script:groupA) -Inventory $inventory).Status | Should -Be 'Skipped'
		Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly -ParameterFilter { $Uri -like '/v1.0/me?*' -or $Uri -like '*/roleAssignments?*' }
	}

	It 'rejects app-only connections without a tenant-wide application permission' {
		Mock Get-MgContext { @{ AuthType = 'AppOnly'; Scopes = @('Group.Read.All') } }
		{ Get-PackageCoverageInventory } | Should -Throw '*EntitlementManagement.Read.All*application permission*'
		Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly
	}

	It 'does not assume app-only visibility if Graph rejects package enumeration' {
		Mock Get-MgContext { @{ AuthType = 'AppOnly'; Scopes = @('EntitlementManagement.Read.All') } }
		Mock Invoke-MgGraphRequest { throw 'Forbidden' }
		{ Get-PackageCoverageInventory } | Should -Throw '*Forbidden*'
	}

	It 'detects overlapping roles even when the scope refers to a different group' {
		$script:fixture.Scopes = @(
			@{ role = @{ originSystem = 'AadGroup'; originId = "Member_$script:groupA" }; scope = @{ originSystem = 'AadGroup'; originId = $script:groupB; isRootScope = $true } }
		)
		$inventory = @(Get-PackageCoverageInventory)
		(Get-PackageCoverageResult -GroupIds @($script:groupA) -Inventory $inventory).Status | Should -Be 'Manual review'
		(Get-PackageCoverageResult -GroupIds @($script:groupB) -Inventory $inventory).Status | Should -Be 'Manual review'
	}

	It 'follows collection pagination and refuses malformed pages' {
		Mock Invoke-MgGraphRequest {
			param($Uri)
			if ($Uri -eq '/first') { return @{ value = @('first'); '@odata.nextLink' = '/second' } }
			if ($Uri -eq '/second') { return @{ value = @('second') } }
			return @{}
		}
		@(Get-PackageCoverageCollection '/first') | Should -Be @('first', 'second')
		{ Get-PackageCoverageCollection '/missing' } | Should -Throw '*complete collection*'
	}
}
