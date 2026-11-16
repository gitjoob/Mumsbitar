BeforeAll {
	Import-Module Microsoft.Graph.Authentication
	$script:reporter = Join-Path $PSScriptRoot '../Get-AccessReviewStatusReport.ps1'
	$script:base = 'https://graph.microsoft.com/v1.0'
	$script:definitionsUri = "$base/identityGovernance/accessReviews/definitions"
	$script:definitionId = '11111111-1111-1111-1111-111111111111'
	$script:configuredOwner = '22222222-2222-2222-2222-222222222222'
	$script:currentOwner = '44444444-4444-4444-4444-444444444444'
	$script:fallback = '55555555-5555-5555-5555-555555555555'
	$script:groupId = '77777777-7777-7777-7777-777777777777'
	$script:applicationId = '88888888-8888-8888-8888-888888888888'
	$script:oldId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
	$script:newId = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
	$script:stageId = 'cccccccc-cccc-cccc-cccc-cccccccccccc'
	$script:instanceRoot = "$definitionsUri/$definitionId/instances"
	$script:decisionSelect = '?$select=id,decision,applyResult,reviewedBy,resourceLink'
	function New-StatusDecision {
		param([string]$Id, [string]$Decision)
		@{
			id = $Id; decision = $Decision; applyResult = 'AppliedSuccessfully'
			reviewedBy = @{ id = if ($Decision -in @('Approve', 'Deny', 'DontKnow')) { $configuredOwner } else { [guid]::Empty.ToString() } }
			resourceLink = "$base/groups/$groupId"
		}
	}
}

Describe 'Read-only historical access review status' {
	BeforeEach {
		$script:outputPath = Join-Path $TestDrive "$([guid]::NewGuid()).html"
		$script:responses = @{}
		$script:failures = @{}
		$script:requests = [System.Collections.Generic.List[object]]::new()
		$script:definition = @{
			id = $definitionId; displayName = 'AR Finance'; status = 'InProgress'; descriptionForReviewers = 'Finance access'
			scope = @{ query = "/groups/$groupId/members" }
			stageSettings = @(
				@{ stageId = '1'; durationInDays = 3; reviewers = @() },
				@{ stageId = '2'; durationInDays = 4; dependsOn = @('1'); reviewers = @(@{ query = './manager' }) },
				@{ stageId = '3'; durationInDays = 4; dependsOn = @('2'); reviewers = @(@{ query = "/users/$configuredOwner" }); fallbackReviewers = @(@{ query = "/users/$fallback" }) }
			)
		}
		$responses[$definitionsUri] = @{ value = @($definition, @{ id = 'unmatched'; displayName = 'Different campaign' }) }
		$responses[$instanceRoot] = @{ value = @(
			@{ id = $oldId; status = 'Completed'; startDateTime = '2025-01-01T00:00:00Z'; endDateTime = '2025-01-12T00:00:00Z'; reviewers = @(); scope = @{ query = "/groups/$groupId/members" } },
			@{ id = $newId; status = 'InProgress'; startDateTime = '2026-01-01T00:00:00Z'; endDateTime = '2026-01-12T00:00:00Z'; reviewers = @(@{ query = './owners' }); scope = @{ query = "/groups/$groupId/members" } }
		) }
		$responses["$base/groups/$groupId`?`$select=id,displayName"] = @{ id = $groupId; displayName = 'Finance group' }
		$responses["$base/groups/$groupId/owners?`$select=id,displayName"] = @{ value = @(
			@{ id = $currentOwner; displayName = 'Current resource owner'; '@odata.type' = '#microsoft.graph.user' },
			@{ id = '66666666-6666-6666-6666-666666666666'; displayName = 'Automation owner'; '@odata.type' = '#microsoft.graph.servicePrincipal' }
		) }
		foreach ($identity in @(
			@{ id = $configuredOwner; displayName = 'Configured owner'; userPrincipalName = 'configured.owner@example.test' },
			@{ id = $currentOwner; displayName = 'Current resource owner'; userPrincipalName = 'current.owner@example.test' },
			@{ id = $fallback; displayName = 'Fallback reviewer'; userPrincipalName = 'fallback@example.test' }
		)) {
			$responses["$base/users/$($identity.id)?`$select=id,displayName,userPrincipalName"] = $identity
		}
		foreach ($id in @($oldId, $newId)) {
			$responses["$instanceRoot/$id/contactedReviewers"] = @{ value = @(
				@{ id = $configuredOwner; displayName = 'Recorded historical owner'; userPrincipalName = 'historical.owner@example.test' }
			) }
			$responses["$instanceRoot/$id/stages"] = @{ value = @(
				@{ id = $stageId; status = 'Completed'; startDateTime = '2025-01-01T00:00:00Z'; reviewers = @(@{ query = "/users/$configuredOwner" }) }
			) }
			$responses["$instanceRoot/$id/stages/$stageId/decisions$decisionSelect"] = @{ value = @(
				(New-StatusDecision 'stage-decision' 'Approve')
			) }
		}
		$responses["$instanceRoot/$oldId/decisions$decisionSelect"] = @{ value = @(
			(New-StatusDecision 'a' 'Approve'), (New-StatusDecision 'b' 'Deny'), (New-StatusDecision 'c' 'DontKnow')
		) }
		$responses["$instanceRoot/$newId/decisions$decisionSelect"] = @{ value = @(
			(New-StatusDecision 'd' 'Approve'), (New-StatusDecision 'e' 'NotReviewed'),
			(New-StatusDecision 'f' 'NotEnabled'), (New-StatusDecision 'g' 'Unexpected')
		) }
		Mock Get-MgContext { [pscustomobject]@{ TenantId = 'fictional-tenant'; Environment = 'Global'; AuthType = 'Delegated'; Scopes = @('AccessReview.Read.All') } }
		Mock Get-MgEnvironment { [pscustomobject]@{ GraphEndpoint = 'https://graph.microsoft.com' } }
		Mock Invoke-MgGraphRequest {
			param($Method, $Uri)
			$requests.Add(@{ Method = "$Method"; Uri = "$Uri" })
			if ($failures.ContainsKey("$Uri")) { throw $failures["$Uri"] }
			if (-not $responses.ContainsKey("$Uri")) { throw "Unexpected request: $Method $Uri" }
			$responses["$Uri"] | ConvertTo-Json -Depth 20 | ConvertFrom-Json
		}
		Mock Start-Process {}
		Mock Write-Host {}
	}

	It 'includes every historical instance, future configured stages and distinct reviewer evidence' {
		$report = & $reporter -HtmlReportPath $outputPath
		$report.Completed | Should -BeTrue
		$report.Issues.Count | Should -Be 0
		$review = $report.Definitions[0]
		$review.Instances.Count | Should -Be 2
		$review.Configuration.Count | Should -Be 3
		$review.Configuration[2].Reviewers[0].Users[0].UserPrincipalName | Should -Be 'configured.owner@example.test'
		$review.Configuration[2].Fallbacks[0].Users[0].Id | Should -Be $fallback
		$review.Configuration[0].Fallbacks.Count | Should -Be 0
		$instance = $review.Instances[0]
		$instance.Stages.Count | Should -Be 1
		$instance.Stages[0].Id | Should -Be $stageId
		$instance.Stages[0].Reviewers[0].Mode | Should -Be 'Named user'
		$instance.Stages[0].Fallbacks.Count | Should -Be 0
		$instance.ContactedReviewers[0].UserPrincipalName | Should -Be 'historical.owner@example.test'
		$instance.Resources[0].Owners[0].UserPrincipalName | Should -Be 'current.owner@example.test'
		$instance.Resources[0].OtherOwners.Count | Should -Be 1
		$instance.Reviewers[0].Users[0].Id | Should -Be $currentOwner
		$html = Get-Content -LiteralPath $outputPath -Raw
		foreach ($text in @('Configured stage 3', 'configured.owner@example.test', 'historical.owner@example.test', 'current.owner@example.test', $stageId, 'not historical evidence', 'not mapped to stages')) {
			$html | Should -Match ([regex]::Escape($text))
		}
		@($requests | Where-Object Method -NE GET).Count | Should -Be 0
		@($requests | Where-Object Uri -Like "*/users/$configuredOwner`?*").Count | Should -Be 1
		@($requests | Where-Object Uri -Like "*/groups/$groupId/owners*").Count | Should -Be 1
		Should -Invoke Start-Process -Times 0
	}

	It 'paginates definitions, instances, recorded contacts and owners and uses the first literal prefix' {
		$definition.displayName = 'ar[*] Finance'
		$responses[$definitionsUri] = @{ value = @(@{ id = 'decoy'; displayName = 'ARRR Finance' }); '@odata.nextLink' = "$definitionsUri`?page=2" }
		$responses["$definitionsUri`?page=2"] = @{ value = @($definition) }
		$second = $responses[$instanceRoot].value[1]
		$responses[$instanceRoot].value = @($responses[$instanceRoot].value[0])
		$responses[$instanceRoot]['@odata.nextLink'] = "$instanceRoot`?page=2"
		$responses["$instanceRoot`?page=2"] = @{ value = @($second) }
		$responses["$instanceRoot/$oldId/contactedReviewers"]['@odata.nextLink'] = "$instanceRoot/$oldId/contactedReviewers?page=2"
		$responses["$instanceRoot/$oldId/contactedReviewers?page=2"] = @{ value = @(@{ id = $fallback; displayName = 'Recorded fallback'; userPrincipalName = 'recorded.fallback@example.test' }) }
		$responses["$base/groups/$groupId/owners?`$select=id,displayName"]['@odata.nextLink'] = "$base/groups/$groupId/owners?page=2"
		$responses["$base/groups/$groupId/owners?page=2"] = @{ value = @(@{ id = $fallback; displayName = 'Additional owner'; '@odata.type' = '#microsoft.graph.user' }) }
		$report = & $reporter -Prefixes @('AR[*]', 'AR[') -HtmlReportPath $outputPath
		$report.Issues.Count | Should -Be 0
		$report.Definitions.Count | Should -Be 1
		$report.Definitions[0].Prefix | Should -Be 'AR[*]'
		$report.Definitions[0].Instances.Count | Should -Be 2
		($report.Definitions[0].Instances | Where-Object Id -EQ $oldId).ContactedReviewers.Count | Should -Be 2
		$report.Definitions[0].CurrentResources[0].Owners.Count | Should -Be 2
	}

	It 'counts only instance outcomes in overview metrics and treats unknown outcomes as other' {
		$report = & $reporter -HtmlReportPath $outputPath
		$current = $report.Definitions[0].Instances[0].Summary
		$current.Total | Should -Be 4
		$current.Recorded | Should -Be 1
		$current.NotReviewed | Should -Be 1
		$current.Other | Should -Be 2
		$current.Completion | Should -Be 25
		$current.ApplyResults.AppliedSuccessfully | Should -Be 4
		(Get-Content $outputPath -Raw) | Should -Match '<strong>7</strong><span>Available instance decision items'
		(Get-Content $outputPath -Raw) | Should -Match '<strong>4</strong><span>Recorded outcomes'
	}

	It 'paginates and deduplicates decision items and does not resolve a zero-GUID actor' {
		$uri = "$instanceRoot/$newId/decisions$decisionSelect"
		$duplicate = $responses[$uri].value[0]
		$responses[$uri]['@odata.nextLink'] = "$instanceRoot/$newId/decisions?page=2"
		$responses["$instanceRoot/$newId/decisions?page=2"] = @{ value = @($duplicate, (New-StatusDecision 'h' 'DontKnow')) }
		$report = & $reporter -HtmlReportPath $outputPath
		$report.Definitions[0].Instances[0].Summary.Total | Should -Be 5
		$report.Definitions[0].Instances[0].Summary.DontKnow | Should -Be 1
		@($requests | Where-Object Uri -Like '*/users/00000000-*').Count | Should -Be 0
	}

	It 'keeps failed decision pages unavailable rather than showing partial counts as complete' {
		$responses["$instanceRoot/$newId/decisions$decisionSelect"]['@odata.nextLink'] = "$instanceRoot/$newId/decisions?page=2"
		$failures["$instanceRoot/$newId/decisions?page=2"] = 'Decision page forbidden'
		$report = & $reporter -HtmlReportPath $outputPath -WarningAction SilentlyContinue
		$report.Definitions[0].Instances[0].Summary | Should -BeNullOrEmpty
		$report.Definitions[0].Instances[0].DecisionError | Should -Be 'Decision page forbidden'
		$html = Get-Content $outputPath -Raw
		$html | Should -Match 'Partial data / collection issues'
		$html | Should -Match 'Decision-count coverage: 1 of 2'
		$html | Should -Match '<strong>3</strong><span>Available instance decision items'
	}

	It 'surfaces stage-list failures without inventing a single runtime stage' {
		$failures["$instanceRoot/$newId/stages"] = 'Stages forbidden'
		$report = & $reporter -HtmlReportPath $outputPath -WarningAction SilentlyContinue
		$report.Definitions[0].Instances[0].StageError | Should -Be 'Stages forbidden'
		$report.Definitions[0].Instances[0].Stages.Count | Should -Be 0
		$report.Definitions[0].Configuration.Count | Should -Be 3
		(Get-Content $outputPath -Raw) | Should -Match 'Stages forbidden'
	}

	It 'retains reviewer IDs on denied profile reads while preserving complete historical records' {
		$failures["$base/users/$configuredOwner`?`$select=id,displayName,userPrincipalName"] = 'Profile forbidden'
		$report = & $reporter -HtmlReportPath $outputPath -WarningAction SilentlyContinue
		$user = $report.Definitions[0].Configuration[2].Reviewers[0].Users[0]
		$user.Id | Should -Be $configuredOwner
		$user.Error | Should -Be 'Profile forbidden'
		$report.Definitions[0].Instances[0].ContactedReviewers[0].DisplayName | Should -Be 'Recorded historical owner'
		@($requests | Where-Object Uri -Like "*/users/$configuredOwner`?*").Count | Should -Be 1
	}

	It 'does not interpret missing runtime scopes as an explicit self-review list' {
		$responses["$instanceRoot/$newId/stages"].value[0].Remove('reviewers')
		$report = & $reporter -HtmlReportPath $outputPath
		$report.Definitions[0].Instances[0].Stages[0].Reviewers[0].Mode | Should -Be 'Reviewer scopes unavailable'
		$report.Definitions[0].Instances[1].Reviewers[0].Mode | Should -Be 'Self-review / default reviewer behavior'
	}

	It 'reports current application owners and preserves filtered owner eligibility instead of guessing it' {
		$definition.scope.query = "/servicePrincipals/$applicationId/appRoleAssignedTo"
		$definition.stageSettings[2].reviewers = @(@{ query = "/servicePrincipals/$applicationId/owners?`$filter=userType eq 'Member'" })
		$responses["$base/servicePrincipals/$applicationId`?`$select=id,displayName"] = @{ id = $applicationId; displayName = 'Finance application' }
		$responses["$base/servicePrincipals/$applicationId/owners?`$select=id,displayName"] = @{ value = @(@{ id = $currentOwner; displayName = 'Application owner'; '@odata.type' = '#microsoft.graph.user' }) }
		$report = & $reporter -HtmlReportPath $outputPath
		$review = $report.Definitions[0]
		$review.CurrentResources[0].Kind | Should -Be 'servicePrincipals'
		$review.CurrentResources[0].Owners[0].Id | Should -Be $currentOwner
		$review.Configuration[2].Reviewers[0].Mode | Should -Be 'Filtered owner rule'
		$review.Configuration[2].Reviewers[0].Users.Count | Should -Be 0
		$review.Configuration[2].Reviewers[0].Query | Should -Match '\$filter='
	}

	It 'resolves reviewer-group members as current membership rather than historical assignments' {
		$query = "/groups/$groupId/transitiveMembers"
		$definition.stageSettings[2].reviewers = @(@{ query = $query })
		$responses["$base$query/microsoft.graph.user?`$select=id,displayName,userPrincipalName"] = @{ value = @(
			@{ id = $fallback; displayName = 'Reviewer group member'; userPrincipalName = 'member@example.test' }
		) }
		$report = & $reporter -HtmlReportPath $outputPath
		$scope = $report.Definitions[0].Configuration[2].Reviewers[0]
		$scope.Users[0].UserPrincipalName | Should -Be 'member@example.test'
		$scope.Users[0].ProfileSource | Should -Be 'Current reviewer-group membership'
		$scope.Note | Should -Match 'not proof of historical'
	}

	It 'keeps stage decision failures and missing contacts explicit without losing instance summaries' {
		$failures["$instanceRoot/$newId/stages/$stageId/decisions$decisionSelect"] = 'Stage decisions forbidden'
		$failures["$instanceRoot/$newId/contactedReviewers"] = 'Contacts forbidden'
		$report = & $reporter -HtmlReportPath $outputPath -WarningAction SilentlyContinue
		$report.Definitions[0].Instances[0].Stages[0].Summary | Should -BeNullOrEmpty
		$report.Definitions[0].Instances[0].Stages[0].Error | Should -Be 'Stage decisions forbidden'
		$report.Definitions[0].Instances[0].ContactError | Should -Be 'Contacts forbidden'
		$report.Definitions[0].Instances[0].Summary.Total | Should -Be 4
	}

	It 'shows zero-item completion as unavailable, not 100 percent' {
		$responses["$instanceRoot/$newId/decisions$decisionSelect"] = @{ value = @() }
		$report = & $reporter -HtmlReportPath $outputPath
		$report.Definitions[0].Instances[0].Summary.Completion | Should -BeNullOrEmpty
		(Get-Content $outputPath -Raw) | Should -Match 'n/a \(no items\)'
	}

	It 'supports a single-stage definition without misclassifying omitted fallbacks' {
		$definition.Remove('stageSettings')
		$definition.reviewers = @(@{ query = "/groups/$groupId/owners" })
		$definition.settings = @{ instanceDurationInDays = 11 }
		$report = & $reporter -HtmlReportPath $outputPath
		$configuration = $report.Definitions[0].Configuration
		$configuration.Count | Should -Be 1
		$configuration[0].Reviewers[0].Mode | Should -Be 'Resource owners'
		$configuration[0].Reviewers[0].Users[0].Id | Should -Be $currentOwner
		$configuration[0].Fallbacks.Count | Should -Be 0
	}

	It 'shows inaccessible owners as unavailable, not as a successful empty roster' {
		$failures["$base/groups/$groupId/owners?`$select=id,displayName"] = 'Owner read forbidden'
		$report = & $reporter -HtmlReportPath $outputPath -WarningAction SilentlyContinue
		$report.Definitions[0].CurrentResources[0].OwnersError | Should -Be 'Owner read forbidden'
		$report.Definitions[0].Instances[0].Reviewers[0].Note | Should -Match 'unavailable or partial'
		(Get-Content $outputPath -Raw) | Should -Match 'Owner/resource lookup incomplete'
	}

	It 'includes definitions without instances and distinguishes instance-list failures' {
		$responses[$instanceRoot] = @{ value = @() }
		$report = & $reporter -HtmlReportPath $outputPath
		$report.Definitions[0].Instances.Count | Should -Be 0
		$html = Get-Content $outputPath -Raw
		$html | Should -Match 'No instances / unavailable'
		$detailJson = [regex]::Match($html, '(?s)<script type="application/json" id="status-details">(.*?)</script>').Groups[1].Value
		$detailJson.TrimStart().StartsWith('[') | Should -BeTrue
		@($detailJson | ConvertFrom-Json).Count | Should -Be 1
		$failures[$instanceRoot] = 'Instances forbidden'
		$report = & $reporter -HtmlReportPath (Join-Path $TestDrive 'failed-instances.html') -WarningAction SilentlyContinue
		$report.Definitions[0].Error | Should -Be 'Instances forbidden'
		(Get-Content (Join-Path $TestDrive 'failed-instances.html') -Raw) | Should -Match 'Instance listing unavailable: Instances forbidden'
	}

	It 'writes an empty matched report without issuing resource reads' {
		$report = & $reporter -Prefixes @('No matches') -HtmlReportPath $outputPath
		$report.Definitions.Count | Should -Be 0
		$requests.Count | Should -Be 1
		Test-Path -LiteralPath $outputPath | Should -BeTrue
	}

	It 'rejects unsafe or repeated collection paging and still exports an explicitly partial report' -TestCases @(
		@{ NextLink = 'https://outside.example.test/steal' },
		@{ NextLink = 'https://graph.microsoft.com/v1.0/identityGovernance/accessReviews/definitions' }
	) {
		param($NextLink)
		$responses[$definitionsUri]['@odata.nextLink'] = $NextLink
		{ & $reporter -HtmlReportPath $outputPath -WarningAction SilentlyContinue } | Should -Throw
		$requests.Count | Should -Be 1
		(Get-Content $outputPath -Raw) | Should -Match 'Partial data / collection issues'
	}

	It 'rejects existing output before making Graph calls' {
		Set-Content -LiteralPath $outputPath -Value 'preserve this file'
		{ & $reporter -HtmlReportPath $outputPath } | Should -Throw '*already exists*'
		(Get-Content $outputPath -Raw).Trim() | Should -Be 'preserve this file'
		Should -Invoke Get-MgContext -Times 0
		$requests.Count | Should -Be 0
	}

	It 'encodes names, identities, scopes and errors without activating markup' {
		$payload = '</script><img src=x onerror=alert(1)> "quoted" & text'
		$definition.displayName = "AR $payload"
		$definition.descriptionForReviewers = $payload
		$definition.stageSettings[1].reviewers[0].query = $payload
		$responses["$instanceRoot/$newId/contactedReviewers"].value[0].displayName = $payload
		$failures["$instanceRoot/$newId/stages"] = $payload
		$null = & $reporter -HtmlReportPath $outputPath -WarningAction SilentlyContinue
		$html = Get-Content $outputPath -Raw
		$html.Contains($payload) | Should -BeFalse
		$html | Should -Not -Match '<img'
		$html | Should -Match '&lt;/script&gt;&lt;img'
		$html | Should -Match 'data-target="status-0"'
		$detailJson = [regex]::Match($html, '(?s)<script type="application/json" id="status-details">(.*?)</script>').Groups[1].Value
		$detailJson | Should -Not -Match '<|>|&'
		$pages = @($detailJson | ConvertFrom-Json)
		$pages.Count | Should -Be 2
		$pages[0] | Should -Match '&lt;/script&gt;&lt;img'
		$pages[0] | Should -Not -Match '<img'
	}
}
