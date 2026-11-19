BeforeAll {
	Import-Module Microsoft.Graph.Authentication
	Import-Module Microsoft.Graph.Users
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
	$script:subjectId = '99999999-9999-9999-9999-999999999999'
	$script:managerId = 'dddddddd-dddd-dddd-dddd-dddddddddddd'
	$script:decisionSelect = '?$select=id,decision,applyResult,reviewedBy,resourceLink,principal,principalLink,reviewedDateTime'
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
		$script:managerProfiles = @{}
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
		Mock Get-MgUser {
			param($UserId, $Property, $ExpandProperty)
			if ($ExpandProperty) {
				if (-not $managerProfiles.ContainsKey("$UserId")) { throw "Unexpected manager lookup: $UserId" }
				if ($managerProfiles["$UserId"] -is [string]) { throw $managerProfiles["$UserId"] }
				return $managerProfiles["$UserId"]
			}
			$uri = "$base/users/$UserId`?`$select=id,displayName,userPrincipalName"
			$requests.Add(@{ Method = 'GET'; Uri = $uri })
			if ($failures.ContainsKey($uri)) { throw $failures[$uri] }
			if (-not $responses.ContainsKey($uri)) { throw "Unexpected user lookup: $UserId" }
			$responses[$uri] | ConvertTo-Json -Depth 20 | ConvertFrom-Json
		}
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

	It 'includes current departments for configured and runtime owner reviewers and recorded reviewer identities' {
		$responses["$base/users/$currentOwner`?`$select=id,displayName,userPrincipalName"].department = 'Identity & Operations'
		$responses["$base/users/$configuredOwner`?`$select=id,displayName,userPrincipalName"].department = 'Finance'
		$definition.stageSettings[2].reviewers = @(@{ query = './owners' }, @{ query = "/users/$configuredOwner" })
		foreach ($id in @($oldId, $newId)) {
			$responses["$instanceRoot/$id/stages"].value[0].reviewers = @(@{ query = './owners' })
		}
		$report = & $reporter -HtmlReportPath $outputPath
		$review = $report.Definitions[0]
		$review.Configuration[2].Reviewers[0].Users[0].Department | Should -Be 'Identity & Operations'
		$review.Instances[0].Stages[0].Reviewers[0].Users[0].Department | Should -Be 'Identity & Operations'
		$review.Instances[0].ContactedReviewers[0].Department | Should -Be 'Finance'
		$review.Instances[0].ContactedReviewers[0].DisplayName | Should -Be 'Recorded historical owner'
		Should -Invoke Get-MgUser -Exactly -Times 1 -ParameterFilter { $UserId -eq $currentOwner -and $Property -contains 'Department' }
		Should -Invoke Get-MgUser -Exactly -Times 1 -ParameterFilter { $UserId -eq $configuredOwner -and $Property -contains 'Department' }
		$pages = @([regex]::Match((Get-Content $outputPath -Raw), '(?s)<script type="application/json" id="status-details">(.*?)</script>').Groups[1].Value | ConvertFrom-Json)
		$pages[0] | Should -Match 'Department \(current directory\): Identity &amp; Operations'
		$pages[0] | Should -Match 'Department \(current directory\): Finance'
		$pages[0] | Should -Match 'Department \(current directory\): Not set / not returned'
		$report.Issues.Count | Should -Be 0
	}

	It 'resolves versioned and absolute named reviewer queries with Get-MgUser rather than displaying only paths' -TestCases @(
		@{ QueryPrefix = '/v1.0' },
		@{ QueryPrefix = 'https://graph.microsoft.com/v1.0' }
	) {
		param($QueryPrefix)
		$definition.stageSettings[2].reviewers[0].query = "$QueryPrefix/users/$configuredOwner"
		$report = & $reporter -HtmlReportPath $outputPath
		$reviewer = $report.Definitions[0].Configuration[2].Reviewers[0]
		$reviewer.Mode | Should -Be 'Named user'
		$reviewer.Users[0].DisplayName | Should -Be 'Configured owner'
		Should -Invoke Get-MgUser -Exactly -Times 1 -ParameterFilter { $UserId -eq $configuredOwner -and -not $ExpandProperty }
		Should -Invoke Invoke-MgGraphRequest -Times 0 -ParameterFilter { "$Uri" -like '*/users/*' }
		$html = Get-Content $outputPath -Raw
		$pages = @([regex]::Match($html, '(?s)<script type="application/json" id="status-details">(.*?)</script>').Groups[1].Value | ConvertFrom-Json)
		$pages[0] | Should -Match '<strong>Configured owner</strong> configured.owner@example.test'
		$pages[0] | Should -Match '<details><summary>Technical reviewer rule</summary><code>'
		$report.Issues.Count | Should -Be 0
	}

	It 'includes pending and completed self-review subjects and current managers separately from historical stage actors' {
		$otherSubject = 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'
		$responses["$base/users/$subjectId`?`$select=id,displayName,userPrincipalName"] = @{
			id = $subjectId; displayName = 'Alex Subject'; userPrincipalName = 'alex@example.test'
		}
		$responses["$base/users/$otherSubject`?`$select=id,displayName,userPrincipalName"] = @{
			id = $otherSubject; displayName = 'Sam Pending'; userPrincipalName = 'sam@example.test'
		}
		$responses["$base/users/$managerId`?`$select=id,displayName,userPrincipalName"] = @{
			id = $managerId; displayName = 'Current Line Manager'; userPrincipalName = 'manager@example.test'; department = 'Service Delivery'
		}
		$managerProfiles[$subjectId] = [pscustomobject]@{ Id = $subjectId; Manager = [pscustomobject]@{ Id = $managerId } }
		$managerProfiles[$otherSubject] = [pscustomobject]@{ Id = $otherSubject; Manager = $null }
		$completed = New-StatusDecision 'self-completed' 'Approve'
		$completed.principal = @{ id = $subjectId; '@odata.type' = '#microsoft.graph.userIdentity' }
		$completed.principalLink = "$base/users/$subjectId"
		$completed.reviewedBy = @{ id = $subjectId }
		$completed.reviewedDateTime = '2026-01-02T00:00:00Z'
		$pending = New-StatusDecision 'self-pending' 'NotReviewed'
		$pending.principalLink = "$base/users/$otherSubject"
		$managerDecision = New-StatusDecision 'manager-completed' 'Deny'
		$managerDecision.principal = $completed.principal
		$managerDecision.reviewedBy = @{ id = $configuredOwner }
		$managerStageId = 'ffffffff-ffff-ffff-ffff-ffffffffffff'
		$responses["$instanceRoot/$newId/decisions$decisionSelect"] = @{ value = @($completed, $pending) }
		$responses["$instanceRoot/$newId/stages"].value[0].reviewers = @()
		$responses["$instanceRoot/$newId/stages/$stageId/decisions$decisionSelect"] = @{ value = @($completed, $pending) }
		$responses["$instanceRoot/$newId/stages"].value += @{
			id = $managerStageId; status = 'InProgress'; startDateTime = '2026-01-04T00:00:00Z'; reviewers = @(@{ query = './manager' })
		}
		$responses["$instanceRoot/$newId/stages/$managerStageId/decisions$decisionSelect"] = @{ value = @($managerDecision) }
		$report = & $reporter -HtmlReportPath $outputPath
		$instance = $report.Definitions[0].Instances[0]
		$instance.Summary.Total | Should -Be 2
		$selfStage = $instance.Stages | Where-Object Id -EQ $stageId
		$selfStage.Summary.Items.Count | Should -Be 2
		$selfStage.Summary.Items[0].Subject.UserPrincipalName | Should -Be 'alex@example.test'
		$selfStage.Summary.Items[1].Subject.UserPrincipalName | Should -Be 'sam@example.test'
		$selfStage.Summary.Items[1].ReviewedBy | Should -BeNullOrEmpty
		$selfStage.Summary.Items[1].CurrentManager.Note | Should -Be 'No current manager returned by the directory.'
		$managerItem = ($instance.Stages | Where-Object Id -EQ $managerStageId).Summary.Items[0]
		$managerItem.ReviewedBy.Id | Should -Be $configuredOwner
		$managerItem.CurrentManager.User.Id | Should -Be $managerId
		$managerItem.CurrentManager.User.Department | Should -Be 'Service Delivery'
		Should -Invoke Get-MgUser -Exactly -Times 1 -ParameterFilter { $UserId -eq $subjectId -and $ExpandProperty -eq 'manager' }
		Should -Invoke Get-MgUser -Exactly -Times 1 -ParameterFilter { $UserId -eq $subjectId -and -not $ExpandProperty }
		$html = Get-Content $outputPath -Raw
		$pages = @([regex]::Match($html, '(?s)<script type="application/json" id="status-details">(.*?)</script>').Groups[1].Value | ConvertFrom-Json)
		$pages[0] | Should -Match 'Individual users and decisions'
		$pages[0] | Should -Match 'Sam Pending'
		$pages[0] | Should -Match 'Current Line Manager'
		$pages[0] | Should -Match 'Department \(current directory\): Service Delivery'
		$pages[0] | Should -Match 'not proof of historical or pending stage assignments'
		$report.Issues.Count | Should -Be 0
	}

	It 'keeps user decisions available when current manager lookup fails and does not repeat failed lookups' {
		$decision = $responses["$instanceRoot/$newId/decisions$decisionSelect"].value[0]
		$decision.principal = @{ id = $subjectId; displayName = 'Alex'; userPrincipalName = 'alex@example.test'; '@odata.type' = '#microsoft.graph.userIdentity' }
		$responses["$instanceRoot/$newId/stages/$stageId/decisions$decisionSelect"].value = @($decision)
		$managerProfiles[$subjectId] = 'Manager lookup forbidden'
		$report = & $reporter -HtmlReportPath $outputPath -WarningAction SilentlyContinue
		$item = $report.Definitions[0].Instances[0].Summary.Items[0]
		$item.Subject.DisplayName | Should -Be 'Alex'
		$item.CurrentManager.Error | Should -Be 'Manager lookup forbidden'
		$report.Definitions[0].Instances[0].Summary.Total | Should -Be 4
		$report.Issues.Source | Should -Contain "Current manager of $subjectId"
		Should -Invoke Get-MgUser -Exactly -Times 1 -ParameterFilter { $UserId -eq $subjectId -and $ExpandProperty -eq 'manager' }
	}

	It 'does not look up service principals or missing subject identities as users' {
		$decision = $responses["$instanceRoot/$newId/decisions$decisionSelect"].value[0]
		$decision.principal = @{ id = $applicationId; displayName = 'Automation'; '@odata.type' = '#microsoft.graph.servicePrincipalIdentity' }
		$decision.principalLink = "$base/servicePrincipals/$applicationId"
		$report = & $reporter -HtmlReportPath $outputPath
		$items = $report.Definitions[0].Instances[0].Summary.Items
		$items[0].Subject.DisplayName | Should -Be 'Automation'
		$items[0].CurrentManager | Should -BeNullOrEmpty
		$items[1].Subject | Should -BeNullOrEmpty
		Should -Invoke Get-MgUser -Times 0 -ParameterFilter { $UserId -eq $applicationId }
		$report.Issues.Count | Should -Be 0
	}

	It 'does not request current managers when no manager-review stage is configured' {
		$definition.stageSettings = @($definition.stageSettings[2])
		$decision = $responses["$instanceRoot/$newId/decisions$decisionSelect"].value[0]
		$decision.principal = @{ id = $subjectId; displayName = 'Alex'; userPrincipalName = 'alex@example.test'; '@odata.type' = '#microsoft.graph.userIdentity' }
		$report = & $reporter -HtmlReportPath $outputPath
		$report.Definitions[0].Instances[0].Summary.Items[0].Subject.DisplayName | Should -Be 'Alex'
		$report.Definitions[0].Instances[0].Summary.IncludeManagers | Should -BeFalse
		Should -Invoke Get-MgUser -Times 0 -ParameterFilter { $ExpandProperty }
		(Get-Content $outputPath -Raw) | Should -Not -Match 'Current line manager \(directory snapshot\)'
	}

	It 'does not follow named-reviewer URLs outside the connected Graph endpoint' {
		$definition.stageSettings[2].reviewers[0].query = "https://outside.example.test/v1.0/users/$managerId"
		$report = & $reporter -HtmlReportPath $outputPath
		$report.Definitions[0].Configuration[2].Reviewers[0].Mode | Should -Be 'Unexpanded reviewer scope'
		Should -Invoke Get-MgUser -Times 0 -ParameterFilter { $UserId -eq $managerId }
		@($requests | Where-Object Uri -Like 'https://outside.example.test/*').Count | Should -Be 0
	}

	It 'encodes individual subject, manager and decision-maker metadata inside decision tables' {
		$payload = '</script><img src=x onerror=alert(1)>'
		$decision = $responses["$instanceRoot/$newId/decisions$decisionSelect"].value[0]
		$decision.principal = @{ id = $subjectId; displayName = $payload; userPrincipalName = 'subject@example.test'; '@odata.type' = '#microsoft.graph.userIdentity' }
		$decision.reviewedBy = @{ id = $configuredOwner; displayName = $payload }
		$decision.reviewedDateTime = $payload
		$responses["$base/users/$managerId`?`$select=id,displayName,userPrincipalName"] = @{
			id = $managerId; displayName = $payload; userPrincipalName = 'manager@example.test'; department = $payload
		}
		$managerProfiles[$subjectId] = [pscustomobject]@{ Id = $subjectId; Manager = [pscustomobject]@{ Id = $managerId } }
		$null = & $reporter -HtmlReportPath $outputPath
		$html = Get-Content $outputPath -Raw
		$pages = @([regex]::Match($html, '(?s)<script type="application/json" id="status-details">(.*?)</script>').Groups[1].Value | ConvertFrom-Json)
		$pages[0] | Should -Not -Match '<img'
		$pages[0] | Should -Match '&lt;/script&gt;&lt;img'
		$pages[0] | Should -Match 'Department \(current directory\): &lt;/script&gt;&lt;img'
	}

	It 'uses instance outcomes when no stage is in progress and treats unknown outcomes as other' {
		$report = & $reporter -HtmlReportPath $outputPath
		$current = $report.Definitions[0].Instances[0].Summary
		$current.Total | Should -Be 4
		$current.Recorded | Should -Be 1
		$current.NotReviewed | Should -Be 1
		$current.Other | Should -Be 2
		$current.Completion | Should -Be 25
		$current.ApplyResults.AppliedSuccessfully | Should -Be 4
		(Get-Content $outputPath -Raw) | Should -Match '<strong>7</strong><span>Available review progress items'
		(Get-Content $outputPath -Raw) | Should -Match '<strong>4</strong><span>Recorded outcomes'
		$dashboard = [regex]::Match((Get-Content $outputPath -Raw), '(?s)<section id="status-dashboard".*?</section>').Value
		$dashboard | Should -Match 'data-review-status="InProgress" data-count="1"'
		$dashboard | Should -Not -Match 'data-review-status="Completed"'
	}

	It 'paginates and deduplicates decision items and does not resolve a zero-GUID actor' {
		$uri = "$instanceRoot/$newId/decisions$decisionSelect"
		$duplicate = $responses[$uri].value[0]
		$responses[$uri]['@odata.nextLink'] = "$instanceRoot/$newId/decisions?page=2"
		$responses["$instanceRoot/$newId/decisions?page=2"] = @{ value = @($duplicate, (New-StatusDecision 'h' 'DontKnow')) }
		$report = & $reporter -HtmlReportPath $outputPath
		$report.Definitions[0].Instances[0].Summary.Total | Should -Be 5
		$report.Definitions[0].Instances[0].Summary.Items.Count | Should -Be 5
		$report.Definitions[0].Instances[0].Summary.DontKnow | Should -Be 1
		@($requests | Where-Object Uri -Like '*/users/00000000-*').Count | Should -Be 0
	}

	It 'shows pending active-stage tasks separately when a matched prefix has zero instance-level NotReviewed outcomes' {
		$responses["$instanceRoot/$newId/decisions$decisionSelect"].value = @((New-StatusDecision 'a' 'Approve'))
		$responses["$instanceRoot/$newId/stages"].value[0].status = 'InProgress'
		$responses["$instanceRoot/$newId/stages/$stageId/decisions$decisionSelect"].value = @(
			(New-StatusDecision 'pending-one' 'NotReviewed'), (New-StatusDecision 'pending-two' 'notReviewed')
		)
		$responses["$instanceRoot/$oldId/stages/$stageId/decisions$decisionSelect"].value = @((New-StatusDecision 'historical' 'NotReviewed'))
		$report = & $reporter -Prefixes @('ar', 'AR Finance') -HtmlReportPath $outputPath
		$report.Definitions[0].Prefix | Should -BeExactly 'ar'
		$report.Definitions[0].Instances[0].Summary.NotReviewed | Should -Be 0
		$html = Get-Content $outputPath -Raw
		$html | Should -Match 'data-show-chart="ar"[^>]*>ar</button></td><td>1</td><td>2</td><td>2 / 2</td><td>5</td><td>3</td><td>2</td><td>2</td>'
		$html | Should -Match '<strong>2</strong><span>Not reviewed \(review progress\)'
	}

	It 'shows active-stage progress separately from instance outcomes in the inventory' {
		$responses["$instanceRoot/$newId/decisions$decisionSelect"].value = @((New-StatusDecision 'i' 'NotReviewed'))
		$responses["$instanceRoot/$newId/stages"].value[0].status = 'InProgress'
		$responses["$instanceRoot/$newId/stages/$stageId/decisions$decisionSelect"].value = @(
			(New-StatusDecision 's1' 'Approve'), (New-StatusDecision 's2' 'NotReviewed')
		)
		$null = & $reporter -HtmlReportPath $outputPath
		$html = Get-Content $outputPath -Raw
		$html | Should -Match '<th>Review progress recorded / items</th><th>Progress source</th>'
		$html | Should -Match "<td>1 / 2</td><td>Active stage $stageId</td><td>1</td>"
		$html | Should -Match '<td>Instance outcomes</td>'
	}

	It 'shows unavailable active-stage progress when its decisions cannot be read' {
		$responses["$instanceRoot/$newId/stages"].value[0].status = 'InProgress'
		$failures["$instanceRoot/$newId/stages/$stageId/decisions$decisionSelect"] = 'Stage decisions unavailable'
		$null = & $reporter -HtmlReportPath $outputPath -WarningAction SilentlyContinue
		(Get-Content $outputPath -Raw) | Should -Match "<td>Unavailable / not materialized</td><td>Active stage $stageId</td>"
		(Get-Content $outputPath -Raw) | Should -Match '<td>1 / 2</td><td>3</td><td>3</td><td>0 \(partial\)</td><td>Unavailable</td>'
	}

	It 'uses fully recorded RVDC stage results in inventory and totals without adding stale instance outcomes or completed stages' {
		$definition.displayName = 'RVDC Fictional review'
		$responses[$instanceRoot].value = @($responses[$instanceRoot].value[1])
		$responses["$instanceRoot/$newId/decisions$decisionSelect"].value = @(1..55 | ForEach-Object { New-StatusDecision "i$_" 'NotReviewed' })
		$responses["$instanceRoot/$newId/stages"].value[0].status = 'InProgress'
		$responses["$instanceRoot/$newId/stages/$stageId/decisions$decisionSelect"].value = @(1..55 | ForEach-Object { New-StatusDecision "s$_" 'Approve' })
		$earlier = 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'
		$responses["$instanceRoot/$newId/stages"].value += @{ id = $earlier; status = 'Completed' }
		$responses["$instanceRoot/$newId/stages/$earlier/decisions$decisionSelect"] = @{ value = @((New-StatusDecision 'earlier' 'Approve')) }
		$report = & $reporter -Prefixes RVDC -HtmlReportPath $outputPath
		$report.Definitions[0].Instances[0].Summary.Recorded | Should -Be 0
		$html = Get-Content $outputPath -Raw
		$html | Should -Match "<td>55 / 55</td><td>Active stage $stageId</td>"
		$html | Should -Match '<strong>55</strong><span>Available review progress items'
		$html | Should -Match '<strong>55</strong><span>Recorded outcomes'
		$html | Should -Match 'data-show-chart="RVDC"[^>]*>RVDC</button></td><td>1</td><td>1</td><td>1 / 1</td><td>55</td><td>55</td><td>0</td><td>0</td>'
		$pages = @([regex]::Match($html, '(?s)<script type="application/json" id="status-details">(.*?)</script>').Groups[1].Value | ConvertFrom-Json)
		$pages[0] | Should -Match '<h4>Instance-level outcomes</h4>.*?<strong>0 / 55</strong>'
	}

	It 'keeps completed-instance outcomes even if an old stage still reports InProgress' {
		$responses[$instanceRoot].value[1].status = 'Completed'
		$responses["$instanceRoot/$newId/stages"].value[0].status = 'InProgress'
		$null = & $reporter -HtmlReportPath $outputPath
		(Get-Content $outputPath -Raw) | Should -Match '<td>1 / 4</td><td>Instance outcomes</td>'
	}

	It 'does not combine multiple active stages into misleading instance progress' {
		$responses["$instanceRoot/$newId/stages"].value[0].status = 'InProgress'
		$secondStage = 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee'
		$responses["$instanceRoot/$newId/stages"].value += @{ id = $secondStage; status = 'InProgress' }
		$responses["$instanceRoot/$newId/stages/$secondStage/decisions$decisionSelect"] = @{ value = @((New-StatusDecision 's2' 'Approve')) }
		$null = & $reporter -HtmlReportPath $outputPath
		$html = Get-Content $outputPath -Raw
		$html | Should -Match '<td>Unavailable / not materialized</td><td>Active-stage progress unavailable</td>'
		$html | Should -Match '<strong>3</strong><span>Available review progress items'
	}

	It 'shows unavailable rather than a zero when both decision sources fail' {
		$responses["$instanceRoot/$newId/stages"].value[0].status = 'InProgress'
		foreach ($id in @($oldId, $newId)) {
			$failures["$instanceRoot/$id/decisions$decisionSelect"] = 'Instance decisions unavailable'
		}
		$failures["$instanceRoot/$newId/stages/$stageId/decisions$decisionSelect"] = 'Stage decisions unavailable'
		$null = & $reporter -Prefixes 'AR' -HtmlReportPath $outputPath -WarningAction SilentlyContinue
		(Get-Content $outputPath -Raw) | Should -Match 'data-show-chart="AR"[^>]*>AR</button></td><td>1</td><td>2</td><td>0 / 2</td>(<td>Unavailable</td>){4}'
	}

	It 'renders the combined chart and separate prefix counts and percentages' {
		$prefix = 'Ops & "IT"'
		$responses[$definitionsUri].value += @{ id = 'ops'; displayName = "$prefix access"; status = 'Completed' }
		$responses["$definitionsUri/ops/instances"] = @{ value = @() }
		$null = & $reporter -Prefixes @('AR ', $prefix) -HtmlReportPath $outputPath
		$html = Get-Content $outputPath -Raw
		$html | Should -Match '<option value="">All prefixes</option>'
		$html | Should -Match 'data-show-chart="Ops &amp; &quot;IT&quot;"'
		$all = [regex]::Match($html, '(?s)<div data-chart-prefix="">.*?</ul>').Value
		$ops = [regex]::Match($html, '(?s)<div data-chart-prefix="Ops &amp; &quot;IT&quot;" hidden>.*?</ul>').Value
		$all | Should -Match 'Completed: 1 \(50%\); InProgress: 1 \(50%\)'
		$ops | Should -Match 'Completed: 1 \(100%\)'
		$ops | Should -Not -Match 'InProgress'
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
		$html | Should -Match '<strong>3</strong><span>Available review progress items'
		$html | Should -Match '<td>1 / 2</td><td>3</td><td>3</td><td>0 \(partial\)</td><td>No active stages returned</td>'
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
		$report.Definitions[0].Instances[0].ContactedReviewers[0].Department | Should -BeNullOrEmpty
		$pages = @([regex]::Match((Get-Content $outputPath -Raw), '(?s)<script type="application/json" id="status-details">(.*?)</script>').Groups[1].Value | ConvertFrom-Json)
		$pages[0] | Should -Match 'Department \(current directory\): Unavailable \(profile lookup failed\)'
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
		$responses["$base/users/$fallback`?`$select=id,displayName,userPrincipalName"].department = 'People Operations'
		$report = & $reporter -HtmlReportPath $outputPath
		$scope = $report.Definitions[0].Configuration[2].Reviewers[0]
		$scope.Users[0].UserPrincipalName | Should -Be 'member@example.test'
		$scope.Users[0].ProfileSource | Should -Be 'Current reviewer-group membership'
		$scope.Users[0].Department | Should -Be 'People Operations'
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
		(Get-Content $outputPath -Raw) | Should -Match 'No review definitions available for the status chart'
		(Get-Content $outputPath -Raw) | Should -Not -Match 'class="status-pie"'
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
