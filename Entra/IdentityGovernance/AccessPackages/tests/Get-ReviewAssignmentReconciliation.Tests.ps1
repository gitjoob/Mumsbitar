BeforeAll {
	Import-Module Microsoft.Graph.Authentication
	$script:path = Join-Path $PSScriptRoot '../Get-ReviewAssignmentReconciliation.ps1'
	$script:catalog = '11111111-1111-1111-1111-111111111111'
	$script:group = '22222222-2222-2222-2222-222222222222'
	$script:userId = '33333333-3333-3333-3333-333333333333'
	$script:base = 'https://graph.microsoft.com/v1.0'
	function Invoke-Report {
		param([string]$Path)
		& $script:path -CatalogId $script:catalog `
			-PackageCreatedFrom '2026-01-01T00:00:00Z' -PackageCreatedBefore '2026-01-03T00:00:00Z' `
			-ReviewPrefix 'AR ' -ReviewStartedFrom '2025-12-01T00:00:00Z' -ReviewStartedBefore '2026-01-04T00:00:00Z' `
			-RequestCreatedFrom '2026-01-01T00:00:00Z' -RequestCreatedBefore '2026-01-03T00:00:00Z' -ReportPath $Path
	}
}

Describe 'Read-only review assignment inventory' {
	BeforeEach {
		$script:fixture = @{
			Decisions = @(@{
				id = 'decision-1'; decision = 'Deny'; applyResult = 'AppliedSuccessfully'
				principalLink = "$script:base/users/$script:userId"; resourceLink = "$script:base/groups/$script:group"
				principal = @{ id = $script:userId; displayName = 'Example' }
			})
			Roles = @(@{ role = @{ originSystem = 'AadGroup'; originId = "Member_$script:group" }
				scope = @{ originSystem = 'AadGroup'; originId = $script:group; isRootScope = $true } })
			Assignments = @(@{ id = 'assignment-1'; state = 'delivered'; target = @{ objectId = $script:userId } })
			Requests = @(@{ id = 'request-1'; requestType = 'adminAdd'; state = 'delivered'; createdDateTime = '2026-01-02T00:00:00Z'; assignment = @{ id = 'assignment-1' } })
			Broken = $false; NextPage = $null
		}
		Mock Get-MgContext { @{ TenantId = 'tenant'; Environment = 'Global' } }
		Mock Get-MgEnvironment { @{ GraphEndpoint = 'https://graph.microsoft.com' } }
		$fixture = $script:fixture
		$catalog = $script:catalog
		Mock Invoke-MgGraphRequest ({
			param($Method, $Uri, $OutputType)
			if ($Method -ne 'GET') { throw 'Unexpected write' }
			switch -Regex ($Uri) {
				'/accessReviews/definitions$' { return @{ value = @(@{ id = 'review-1'; displayName = 'AR example' }); '@odata.nextLink' = $fixture.NextPage } }
				'/accessReviews/definitions\?page=2$' { return @{ value = @() } }
				'/definitions/review-1/instances$' { return @{ value = @(@{ id = 'instance-1'; startDateTime = '2025-12-15T00:00:00Z'; status = 'Completed' }) } }
				'/instances/instance-1/decisions$' {
					if ($fixture.Broken) { return @{} }
					return @{ value = $fixture.Decisions }
				}
				'/accessPackages\?\$expand=catalog$' { return @{ value = @(@{ id = 'package-1'; displayName = 'Package'; createdDateTime = '2026-01-02T00:00:00Z'; catalog = @{ id = $catalog } }) } }
				'/accessPackages/package-1\?' { return @{ resourceRoleScopes = $fixture.Roles } }
				'/assignments\?' { return @{ value = $fixture.Assignments } }
				'/assignmentRequests\?' { return @{ value = $fixture.Requests } }
				default { throw "Unexpected Graph URI: $Uri" }
			}
		}.GetNewClosure())
		$script:reportPath = Join-Path $TestDrive "$([guid]::NewGuid()).json"
	}

	It 'matches an applied group decision to the exact assignment and incident adminAdd' {
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.Results.Count | Should -Be 1
		$report.Results[0].Review.ApplyResult | Should -Be 'AppliedSuccessfully'
		$report.Results[0].Assignment.Id | Should -Be 'assignment-1'
		$report.Results[0].IncidentAdminAddRequests[0].Id | Should -Be 'request-1'
		$report.Results[0].Reason | Should -Match 'Manual review'
		Should -Invoke Invoke-MgGraphRequest -ParameterFilter { $Method -ne 'GET' } -Times 0
	}

	It 'reports unidentifiable decisions instead of inferring group membership' {
		$script:fixture.Decisions[0].resourceLink = "$script:base/servicePrincipals/$script:group"
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.Results.Count | Should -Be 0
		$report.UnmatchedDecisions.Count | Should -Be 1
	}

	It 'includes all package roles and does not mistake an unrelated request for the assignment origin' {
		$script:fixture.Roles += @{ role = @{ originSystem = 'AadApplication'; originId = 'app' }; scope = @{ originSystem = 'AadApplication'; originId = 'app' } }
		$script:fixture.Requests[0].assignment.id = 'different-assignment'
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.Results[0].ResourceRoles.Count | Should -Be 2
		$report.Results[0].IncidentAdminAddRequests.Count | Should -Be 0
		$report.Results[0].Reason | Should -Match 'Manual review'
	}

	It 'fails instead of saving an incomplete collection' {
		$script:fixture.Broken = $true
		{ Invoke-Report $script:reportPath } | Should -Throw '*incomplete collection*'
		Test-Path -LiteralPath $script:reportPath | Should -BeFalse
	}

	It 'follows Graph pages and excludes requests at the exclusive end of the window' {
		$script:fixture.NextPage = "$script:base/identityGovernance/accessReviews/definitions?page=2"
		$script:fixture.Requests[0].createdDateTime = '2026-01-03T00:00:00Z'
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.Results[0].IncidentAdminAddRequests.Count | Should -Be 0
		Should -Invoke Invoke-MgGraphRequest -ParameterFilter { $Uri -like '*definitions?page=2' } -Times 1
	}

	It 'refuses a paging URL outside the connected Graph endpoint' {
		$script:fixture.NextPage = 'https://example.invalid/v1.0/page'
		{ Invoke-Report $script:reportPath } | Should -Throw '*invalid or repeated paging URL*'
		Test-Path -LiteralPath $script:reportPath | Should -BeFalse
	}

	It 'reports missing assignment targets as an incomplete match' {
		$script:fixture.Assignments[0].target = @{}
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.PackageIssues[0].Reason | Should -Match 'target object ID'
		$report.Results[0].Assignment | Should -BeNullOrEmpty
	}

	It 'never overwrites a previous report' {
		Set-Content -LiteralPath $script:reportPath -Value 'previous report'
		{ Invoke-Report $script:reportPath } | Should -Throw '*already exists*'
		Get-Content -LiteralPath $script:reportPath | Should -Be 'previous report'
	}
}
