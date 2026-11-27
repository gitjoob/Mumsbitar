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
			-ReviewPrefix 'AR ' -ReviewStartedFrom '2025-12-01T00:00:00Z' -ReviewStartedBefore '2026-01-04T00:00:00Z' `
			-ReportPath $Path
	}
	function Invoke-Approval {
		param([switch]$Apply, [switch]$Preview)
		$scope = @{
			CatalogId = $script:catalog
			ReviewPrefix = 'AR '; ReviewStartedFrom = '2025-12-01T00:00:00Z'; ReviewStartedBefore = '2026-01-04T00:00:00Z'
			ApprovalCsvPath = $script:approvalPath
		}
		& $script:path @scope -ReportPath $script:reportPath -AuditPath $script:auditPath `
			-Remediate:$Apply -WhatIf:$Preview -Confirm:$false
	}
	function Read-Audit {
		@(Get-Content -LiteralPath $script:auditPath | ForEach-Object { $_ | ConvertFrom-Json })
	}
}

Describe 'Review assignment reconciliation' {
	BeforeEach {
		$script:fixture = @{
			Decisions = @(@{
				id = 'decision-1'; decision = 'Deny'; applyResult = 'AppliedSuccessfully'
				principalLink = "$script:base/users/$script:userId"; resourceLink = "$script:base/groups/$script:group"
				principal = @{ id = $script:userId; displayName = 'Example' }
				reviewedDateTime = '2026-01-02T12:00:00Z'
			})
			Instance = @{ id = 'instance-1'; startDateTime = '2025-12-15T00:00:00Z'; endDateTime = '2026-01-02T00:00:00Z'
				status = 'Completed'; scope = @{ query = "/groups/$script:group/members"; queryType = 'MicrosoftGraph' } }
			Roles = @(@{ role = @{ originSystem = 'AadGroup'; originId = "Member_$script:group" }
				scope = @{ originSystem = 'AadGroup'; originId = $script:group; isRootScope = $true } })
			Assignments = @(@{ id = 'assignment-1'; state = 'delivered'; target = @{ objectId = $script:userId } })
			Requests = @(@{ id = 'request-1'; requestType = 'adminAdd'; state = 'delivered'; createdDateTime = '2026-01-02T00:00:00Z'; assignment = @{ id = 'assignment-1'; target = @{ objectId = $script:userId } } })
			Package = @{ id = 'package-1'; displayName = 'Package'; createdDateTime = '2026-01-02T00:00:00Z'; catalog = @{ id = $script:catalog } }
			ExtraDefinitions = @(); OtherDecisions = @(); OtherStatus = 'Completed'; LiveRoles = $null; LiveDecision = $null
			ExtraPackages = @(); OtherPackageRoles = @()
			Broken = $false; NextPage = $null; FailPost = $false; ResponseState = 'submitted'
			Posts = [System.Collections.Generic.List[object]]::new()
		}
		Mock Get-MgContext { @{ TenantId = 'tenant'; Environment = 'Global'; AuthType = 'AppOnly'; Scopes = @('AccessReview.Read.All', 'EntitlementManagement.ReadWrite.All') } }
		Mock Get-MgEnvironment { @{ GraphEndpoint = 'https://graph.microsoft.com' } }
		Mock Write-Progress {}
		$fixture = $script:fixture
		Mock Invoke-MgGraphRequest ({
			param($Method, $Uri, $OutputType, $Body, $ContentType)
			if ($Method -eq 'POST') {
				$fixture.Posts.Add(($Body | ConvertFrom-Json))
				if ($fixture.FailPost) { throw 'Simulated network failure' }
				return @{ id = 'remove-request'; state = $fixture.ResponseState }
			}
			if ($Method -ne 'GET') { throw 'Unexpected write method' }
			switch -Regex ($Uri) {
				'/accessReviews/definitions$' { return @{ value = @(@{ id = 'review-1'; displayName = 'AR example' }) + $fixture.ExtraDefinitions; '@odata.nextLink' = $fixture.NextPage } }
				'/accessReviews/definitions\?page=2$' { return @{ value = @() } }
				'/definitions/review-1$' { return @{ id = 'review-1'; displayName = 'AR example' } }
				'/definitions/review-1/instances$' { return @{ value = @($fixture.Instance) } }
				'/definitions/review-1/instances/instance-1$' { return $fixture.Instance }
				'/instances/instance-1/decisions/([^/]+)$' {
					$id = $Matches[1]
					if ($id -eq 'decision-1' -and $fixture.LiveDecision) { return $fixture.LiveDecision }
					return @($fixture.Decisions | Where-Object { $_.id -eq $id })[0]
				}
				'/instances/instance-1/decisions$' {
					if ($fixture.Broken) { return @{} }
					return @{ value = $fixture.Decisions }
				}
				'/definitions/review-2/instances$' { return @{ value = @(@{ id = 'instance-2'; status = $fixture.OtherStatus; scope = $fixture.Instance.scope }) } }
				'/instances/instance-2/decisions$' { return @{ value = $fixture.OtherDecisions } }
				'/accessPackages\?\$filter=.*&\$expand=catalog$' { return @{ value = @($fixture.Package) + $fixture.ExtraPackages } }
				'/accessPackages/package-1\?\$expand=catalog,' {
					$roles = if ($null -ne $fixture.LiveRoles) { $fixture.LiveRoles } else { $fixture.Roles }
					return @{ id = $fixture.Package.id; createdDateTime = $fixture.Package.createdDateTime; catalog = $fixture.Package.catalog; resourceRoleScopes = $roles }
				}
				'/accessPackages/package-1\?' { return @{ id = 'package-1'; resourceRoleScopes = $fixture.Roles } }
				'/accessPackages/package-2\?' { return @{ id = 'package-2'; resourceRoleScopes = $fixture.OtherPackageRoles } }
				'/assignments\?' { return @{ value = $fixture.Assignments } }
				'/assignmentRequests\?' { return @{ value = $fixture.Requests } }
				default { throw "Unexpected Graph URI: $Uri" }
			}
		}.GetNewClosure())
		$script:reportPath = Join-Path $TestDrive "$([guid]::NewGuid()).json"
		$script:approvalPath = Join-Path $TestDrive "$([guid]::NewGuid()).csv"
		$script:auditPath = Join-Path $TestDrive "$([guid]::NewGuid()).jsonl"
		[pscustomobject]@{ TenantId = 'tenant'; AssignmentId = 'assignment-1'; PackageId = 'package-1'; UserId = $script:userId; GroupId = $script:group
			ReviewId = 'review-1'; InstanceId = 'instance-1'; DecisionId = 'decision-1'
			ReviewedAt = '2026-01-02T12:00:00Z'; AppliedAt = '' } | Export-Csv -LiteralPath $script:approvalPath -NoTypeInformation
	}

	It 'matches an applied group decision to an assignment without adminAdd history' {
		$script:fixture.Requests = @()
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.Results.Count | Should -Be 1
		$report.Results[0].Review.ApplyResult | Should -Be 'AppliedSuccessfully'
		$report.Results[0].Assignment.Id | Should -Be 'assignment-1'
		$report.Results[0].Reason | Should -Match 'Candidate for removal'
		$report.PreliminaryCandidateCount | Should -Be 1
		(Import-Csv -LiteralPath $report.PreliminarySelectionPath).AssignmentId | Should -Be 'assignment-1'
		Should -Invoke Invoke-MgGraphRequest -ParameterFilter { $Method -ne 'GET' } -Times 0
	}

	It 'remediates the exact CSV generated by preview without an adminAdd' {
		$script:fixture.Requests = @()
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$script:approvalPath = $report.PreliminarySelectionPath
		$script:reportPath = Join-Path $TestDrive "$([guid]::NewGuid()).json"
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 1
		$script:fixture.Posts[0].assignment.id | Should -Be 'assignment-1'
	}

	It 'reports unidentifiable decisions instead of inferring group membership' {
		$script:fixture.Decisions[0].resourceLink = "$script:base/servicePrincipals/$script:group"
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.Results.Count | Should -Be 0
		$report.UnmatchedDecisions.Count | Should -Be 1
	}

	It 'uses the direct group scope and typed user when decision links are absent' {
		$script:fixture.Decisions[0].resourceLink = $null
		$script:fixture.Decisions[0].principalLink = $null
		$script:fixture.Decisions[0].principal.'@odata.type' = '#microsoft.graph.userIdentity'
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.PreliminaryCandidateCount | Should -Be 1
		(Import-Csv -LiteralPath $report.PreliminarySelectionPath).UserId | Should -Be $script:userId
	}

	It 'does not infer an untyped principal when the user link is missing' {
		$script:fixture.Decisions[0].principalLink = $null
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.UnmatchedDecisions.Count | Should -Be 1
		$report.PreliminaryCandidateCount | Should -Be 0
	}

	It 'does not match another user assignment' {
		$script:fixture.Assignments[0].target.objectId = '44444444-4444-4444-4444-444444444444'
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.PreliminaryCandidateCount | Should -Be 0
	}

	It 'does not match an unrelated reviewed group' {
		$other = '44444444-4444-4444-4444-444444444444'
		$script:fixture.Decisions[0].resourceLink = "$script:base/groups/$other"
		$script:fixture.Instance.scope.query = "/groups/$other/members"
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.PreliminaryCandidateCount | Should -Be 0
		$report.DecisionsWithoutMatchingPackage.Count | Should -Be 1
	}

	It 'reports multiple resource roles for manual review' {
		$script:fixture.Roles += @{ role = @{ originSystem = 'AadApplication'; originId = 'app' }; scope = @{ originSystem = 'AadApplication'; originId = 'app' } }
		$script:fixture.Requests[0].assignment.id = 'different-assignment'
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.Results[0].ResourceRoles.Count | Should -Be 2
		$report.PreliminaryCandidateCount | Should -Be 0
		$report.Results[0].Reason | Should -Match 'manual review'
	}

	It 'fails instead of saving an incomplete collection' {
		$script:fixture.Broken = $true
		{ Invoke-Report $script:reportPath } | Should -Throw '*incomplete collection*'
		Test-Path -LiteralPath $script:reportPath | Should -BeFalse
	}

	It 'follows Graph pages' {
		$script:fixture.NextPage = "$script:base/identityGovernance/accessReviews/definitions?page=2"
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.PreliminaryCandidateCount | Should -Be 1
		Should -Invoke Invoke-MgGraphRequest -ParameterFilter { $Uri -like '*definitions?page=2' } -Times 2
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

	It 'checks all packages in the catalog and reviews that overlap the 30-day default' {
		$now = [datetimeoffset]::UtcNow
		$script:fixture.Package.createdDateTime = $now.AddDays(-400).ToString('o')
		$script:fixture.Instance.startDateTime = $now.AddDays(-60).ToString('o')
		$script:fixture.Instance.endDateTime = $now.AddDays(-1).ToString('o')
		$script:fixture.Decisions[0].reviewedDateTime = $now.AddDays(-1).ToString('o')
		$script:fixture.Requests = @()
		& $script:path -CatalogId $script:catalog -ReportPath $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.ReviewWindowBasis | Should -Be 'ScheduledPeriodOverlap'
		$report.Results.Count | Should -Be 1
		$report.PreliminaryCandidateCount | Should -Be 1
		Should -Invoke Write-Progress -ParameterFilter { $Activity -eq 'Reconciling access reviews and assignments' -and $Completed } -Times 1
	}

	It 'previews an approved eligible assignment without writes' {
		Invoke-Approval
		$audit = Read-Audit
		@($audit | Where-Object Type -EQ 'Preview').Count | Should -Be 1
		$script:fixture.Posts.Count | Should -Be 0
	}

	It 'submits an adminRemove request for a revalidated approved assignment and audits it' {
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 1
		$script:fixture.Posts[0].requestType | Should -Be 'adminRemove'
		$script:fixture.Posts[0].assignment.id | Should -Be 'assignment-1'
		$audit = Read-Audit
		@($audit | Where-Object Type -EQ 'Submitting').Count | Should -Be 1
		@($audit | Where-Object Type -EQ 'Submitted')[0].RequestId | Should -Be 'remove-request'
	}

	It 'submits multiple approved assignments without individual prompts' {
		$secondUser = '44444444-4444-4444-4444-444444444444'
		$script:fixture.Decisions += @{ id = 'decision-2'; decision = 'Deny'; applyResult = 'AppliedSuccessfully'
			principalLink = "$script:base/users/$secondUser"; resourceLink = "$script:base/groups/$script:group"
			principal = @{ id = $secondUser }; reviewedDateTime = '2026-01-02T12:00:00Z' }
		$script:fixture.Assignments += @{ id = 'assignment-2'; state = 'delivered'; target = @{ objectId = $secondUser } }
		$script:fixture.Requests += @{ id = 'request-2'; requestType = 'adminAdd'; state = 'delivered'; createdDateTime = '2026-01-02T00:00:00Z'; assignment = @{ id = 'assignment-2'; target = @{ objectId = $secondUser } } }
		$selection = @(Import-Csv -LiteralPath $script:approvalPath)
		$selection += [pscustomobject]@{ TenantId = 'tenant'; AssignmentId = 'assignment-2'; PackageId = 'package-1'; UserId = $secondUser; GroupId = $script:group
			ReviewId = 'review-1'; InstanceId = 'instance-1'; DecisionId = 'decision-2'
			ReviewedAt = '2026-01-02T12:00:00Z'; AppliedAt = '' }
		$selection | Export-Csv -LiteralPath $script:approvalPath -NoTypeInformation
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 2
		@((Read-Audit) | Where-Object Type -EQ 'Submitted').Count | Should -Be 2
	}

	It 'keeps WhatIf read-only even when remediation is enabled' {
		Invoke-Approval -Apply -Preview
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Preview').Count | Should -Be 1
	}

	It 'requires a reviewed candidate CSV for live remediation' {
		{ & $script:path -CatalogId $script:catalog -Remediate } | Should -Throw '*reviewed candidate CSV*'
		$script:fixture.Posts.Count | Should -Be 0
	}

	It 'refuses approval processing without tenant-wide entitlement permission' {
		Mock Get-MgContext { @{ TenantId = 'tenant'; Environment = 'Global'; AuthType = 'AppOnly'; Scopes = @('AccessReview.Read.All') } }
		{ Invoke-Approval } | Should -Throw '*EntitlementManagement.Read.All*'
		$script:fixture.Posts.Count | Should -Be 0
	}

	It 'skips a package containing another resource' {
		$script:fixture.Roles += @{ role = @{ originSystem = 'AadApplication'; originId = 'app' }; scope = @{ originSystem = 'AadApplication'; originId = 'app' } }
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Skipped')[0].Reason | Should -Match 'exactly one matching group Member role'
	}

	It 'skips a package that gains another resource after the inventory' {
		$script:fixture.LiveRoles = $script:fixture.Roles + @(@{ role = @{ originSystem = 'AadApplication' }; scope = @{ originSystem = 'AadApplication'; originId = 'app' } })
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Skipped')[0].Reason | Should -Match 'exactly one matching group Member role'
	}

	It 'does not exclude a group also used by a package outside the catalog' {
		$script:fixture.ExtraPackages = @(@{ id = 'package-2'; catalog = @{ id = '99999999-9999-9999-9999-999999999999' } })
		$script:fixture.OtherPackageRoles = $script:fixture.Roles
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 1
	}

	It 'skips an applied result other than AppliedSuccessfully' {
		$script:fixture.Decisions[0].applyResult = 'AppliedSuccessfullyButObjectNotFound'
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Skipped')[0].Reason | Should -Match 'applied Deny'
	}

	It 'keeps unapplied denials out of the preview CSV' {
		$script:fixture.Decisions[0].applyResult = 'New'
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.PreliminaryCandidateCount | Should -Be 0
		$report.Results[0].Reason | Should -Match 'applied Deny'
	}

	It 'removes an assignment without any adminAdd history' {
		$script:fixture.Requests = @()
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 1
	}

	It 'skips pending requests even if their target cannot be identified' {
		$script:fixture.Requests += @{ id = 'pending'; requestType = 'userAdd'; state = 'pendingApproval'; assignment = @{} }
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Skipped')[0].Reason | Should -Match 'pending or unresolved'
	}

	It 'skips a later review approval outside the selected prefix' {
		$script:fixture.ExtraDefinitions = @(@{ id = 'review-2'; displayName = 'Other review' })
		$script:fixture.OtherDecisions = @(@{ id = 'other'; decision = 'Approve'; reviewedDateTime = '2026-01-03T00:00:00Z'
			principalLink = "$script:base/users/$script:userId"; resourceLink = "$script:base/groups/$script:group"
			principal = @{ id = $script:userId } })
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Skipped')[0].Reason | Should -Match 'conflicting, later or unresolved'
	}

	It 'reports a later approval as a manual case in the initial preview' {
		$script:fixture.ExtraDefinitions = @(@{ id = 'review-2'; displayName = 'Other review' })
		$script:fixture.OtherDecisions = @(@{ id = 'other'; decision = 'Approve'; reviewedDateTime = '2026-01-03T00:00:00Z'
			principalLink = "$script:base/users/$script:userId"; resourceLink = "$script:base/groups/$script:group"
			principal = @{ id = $script:userId } })
		Invoke-Report $script:reportPath
		$report = Get-Content -LiteralPath $script:reportPath -Raw | ConvertFrom-Json
		$report.PreliminaryCandidateCount | Should -Be 0
		$report.Results[0].Reason | Should -Match 'conflicting, later or unresolved'
	}

	It 'rejects a CSV from another tenant' {
		$selection = Import-Csv -LiteralPath $script:approvalPath
		$selection.TenantId = 'another-tenant'
		$selection | Export-Csv -LiteralPath $script:approvalPath -NoTypeInformation
		{ Invoke-Approval -Apply } | Should -Throw '*different tenant*'
		$script:fixture.Posts.Count | Should -Be 0
	}

	It 'does not assume an earlier decision was applied before the denial' {
		$script:fixture.ExtraDefinitions = @(@{ id = 'review-2'; displayName = 'Other review' })
		$script:fixture.OtherDecisions = @(@{ id = 'other'; decision = 'Approve'; reviewedDateTime = '2026-01-01T00:00:00Z'
			principalLink = "$script:base/users/$script:userId"; resourceLink = "$script:base/groups/$script:group"
			principal = @{ id = $script:userId } })
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Skipped')[0].Reason | Should -Match 'conflicting, later or unresolved'
	}

	It 'accepts an older review only when its decision and application both precede the denial' {
		$script:fixture.Decisions[0].appliedDateTime = '2026-01-02T13:00:00Z'
		$selection = Import-Csv -LiteralPath $script:approvalPath
		$selection.AppliedAt = '2026-01-02T13:00:00Z'
		$selection | Export-Csv -LiteralPath $script:approvalPath -NoTypeInformation
		$script:fixture.ExtraDefinitions = @(@{ id = 'review-2'; displayName = 'Other review' })
		$script:fixture.OtherDecisions = @(@{ id = 'other'; decision = 'Approve'; reviewedDateTime = '2026-01-01T00:00:00Z'; appliedDateTime = '2026-01-01T23:00:00Z'
			principalLink = "$script:base/users/$script:userId"; resourceLink = "$script:base/groups/$script:group"
			principal = @{ id = $script:userId } })
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 1
	}

	It 'skips a group with another review still in progress even without returned decisions' {
		$script:fixture.ExtraDefinitions = @(@{ id = 'review-2'; displayName = 'Other review' })
		$script:fixture.OtherStatus = 'InProgress'
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Skipped')[0].Reason | Should -Match 'conflict checking is incomplete'
	}

	It 'skips when the decision changes after the inventory' {
		$script:fixture.LiveDecision = $script:fixture.Decisions[0].Clone()
		$script:fixture.LiveDecision.decision = 'Approve'
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Skipped')[0].Reason | Should -Match 'applied Deny'
	}

	It 'skips a later request affecting the same user' {
		$script:fixture.Requests += @{ id = 'later'; requestType = 'userAdd'; state = 'delivered'; createdDateTime = '2026-01-03T00:00:00Z'
			assignment = @{ id = 'other-assignment'; target = @{ objectId = $script:userId } } }
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Skipped')[0].Reason | Should -Match 'later assignment request'
	}

	It 'skips later requests whose target cannot be identified' {
		$script:fixture.Requests += @{ id = 'later'; requestType = 'userAdd'; state = 'delivered'; createdDateTime = '2026-01-03T00:00:00Z'
			assignment = @{ id = 'other-assignment' } }
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Skipped')[0].Reason | Should -Match 'unknown target'
	}

	It 'skips another active assignment for the same user and package' {
		$script:fixture.Assignments += @{ id = 'other-assignment'; state = 'delivered'; target = @{ objectId = $script:userId } }
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Skipped')[0].Reason | Should -Match 'Another active assignment'
	}

	It 'does not let an edited approval CSV bypass the live evidence checks' {
		$selection = Import-Csv -LiteralPath $script:approvalPath
		$selection.AssignmentId = 'unrelated-assignment'
		$selection | Export-Csv -LiteralPath $script:approvalPath -NoTypeInformation
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Skipped')[0].Reason | Should -Match 'no unique matching review'
	}

	It 'skips when the reviewed selection no longer matches the current evidence' {
		$selection = Import-Csv -LiteralPath $script:approvalPath
		$selection.GroupId = '55555555-5555-5555-5555-555555555555'
		$selection | Export-Csv -LiteralPath $script:approvalPath -NoTypeInformation
		Invoke-Approval -Apply
		$script:fixture.Posts.Count | Should -Be 0
		@((Read-Audit) | Where-Object Type -EQ 'Skipped')[0].Reason | Should -Match 'differs from the approved selection'
	}

	It 'records rejected requests without reporting them as submitted' {
		$script:fixture.ResponseState = 'denied'
		Invoke-Approval -Apply
		$audit = Read-Audit
		@($audit | Where-Object Type -EQ 'Rejected').Count | Should -Be 1
		@($audit | Where-Object Type -EQ 'Submitted').Count | Should -Be 0
	}

	It 'records an uncertain POST and stops instead of retrying it' {
		$script:fixture.FailPost = $true
		{ Invoke-Approval -Apply } | Should -Throw '*outcome is unknown*'
		$script:fixture.Posts.Count | Should -Be 1
		@((Read-Audit) | Where-Object Type -EQ 'Unknown').Count | Should -Be 1
	}
}
