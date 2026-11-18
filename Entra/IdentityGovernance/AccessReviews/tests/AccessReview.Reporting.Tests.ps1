BeforeAll {
	. (Join-Path $PSScriptRoot '../AccessReview.Reporting.ps1')
	$script:exporter = Join-Path $PSScriptRoot '../Export-AccessReviewReport.ps1'
	function New-ReportFixture {
		param([string]$Name = 'AR Demo group', [int[]]$Durations = @(6, 5))
		$date = [datetime]'2026-11-16'
		$stages = @(for ($i = 0; $i -lt $Durations.Count; $i++) {
			$end = $date.AddDays($Durations[$i])
			[pscustomobject]@{
				Name = if ($i -eq $Durations.Count - 1) { 'Resource owners (group owners)' } else { 'Line managers' }
				DurationInDays = $Durations[$i]
				StartDate = $date.ToString('yyyy-MM-dd')
				EndDate = $end.ToString('yyyy-MM-dd')
				Reviewers = @(@{ query = '/users/22222222-2222-2222-2222-222222222222' })
				FallbackReviewers = @()
				DecisionsThatWillMoveToNextStage = if ($i -lt $Durations.Count - 1) { @('Approve', 'NotReviewed') } else { @() }
			}
			$date = $end
		})
		[pscustomobject]@{
			TargetType = 'GroupMembership'; Target = $Name; Group = $Name
			ReviewName = $Name; GroupId = '11111111-1111-1111-1111-111111111111'
			MemberCount = 470; GroupCountBasis = 'DirectUsers'; SubjectCount = $null
			Status = 'Created'; Reason = $null; MatchedReviewName = $null
			IsNested = $false; IsSyncedGroup = $false; GroupPaths = @(); ParentGroupIds = @()
			StageNames = $stages.Name -join ' -> '
			ReviewConfiguration = [pscustomobject]@{
				Frequency = 'OneTime'; StartDate = '2026-11-16'; InstanceDurationInDays = 11
				SeriesEndDate = $null; AutoApplyDecisionsEnabled = $true; DefaultDecision = 'Deny'
				Scope = @{ query = '/groups/11111111-1111-1111-1111-111111111111/members' }
				Stages = $stages
			}
		}
	}
}

Describe 'Offline access review reporting' {
	BeforeEach {
		$script:context = @{ Mode = 'Test preview'; RunId = 'test-run'; Completed = $true; Failure = $null }
		$script:logPath = Join-Path $TestDrive ([guid]::NewGuid().ToString() + '.jsonl')
		$script:outputPath = Join-Path $TestDrive ([guid]::NewGuid().ToString() + '.html')
	}

	It 'renders supplied stage boundaries and separates a recurring series expiry' {
		$result = New-ReportFixture
		$result.ReviewConfiguration.Frequency = 'Quarterly'
		$result.ReviewConfiguration.SeriesEndDate = '2027-11-30'
		$html = New-AccessReviewHtml -Results @($result) -Context $script:context
		$html | Should -Match '2026-11-16 &rarr; 2026-11-22'
		$html | Should -Match '2026-11-22 &rarr; 2026-11-27'
		$html | Should -Match 'Recurring series expires: 2027-11-30'
		$html | Should -Match 'Approve, NotReviewed'
		$html | Should -Match '/users/22222222-2222-2222-2222-222222222222'
	}

	It 'encodes untrusted names, role text and errors rather than embedding active markup' {
		$payload = '</script><img src=x onerror=alert(1)> "quoted" & text'
		$result = New-ReportFixture -Name $payload
		$result | Add-Member Role $payload
		$result.Reason = $payload
		$html = New-AccessReviewHtml -Results @($result) -Context $script:context
		$html.Contains($payload) | Should -BeFalse
		$html | Should -Match '&lt;/script&gt;&lt;img'
		$html | Should -Not -Match '<img'
	}

	It 'shows recorded descriptions and application names without interpreting markup' {
		$result = New-ReportFixture
		$result.ReviewConfiguration | Add-Member DescriptionForReviewers "Member of 'Finance <Readers>'."
		$result.ReviewConfiguration | Add-Member DescriptionForAdmins 'Parent -> Finance <Readers>'
		$result.GroupPaths = @(@{
			GroupIds = @('parent', $result.GroupId); GroupNames = @('Parent', 'Finance <Readers>')
			ServicePrincipalId = 'application-id'; ServicePrincipalName = 'Payroll <Application>'
		})
		$html = New-AccessReviewHtml -Results @($result) -Context $script:context
		$html | Should -Match 'Reviewer description:</strong> Member of &#39;Finance &lt;Readers&gt;&#39;'
		$html | Should -Match 'Administrator description:</strong> Parent -&gt; Finance &lt;Readers&gt;'
		$html | Should -Match 'Payroll &lt;Application&gt; \(application-id\)'
		$html | Should -Not -Match '<Readers>|<Application>'
	}

	It 'adds safe Entra navigation only for a valid group Object ID (<Id>)' -TestCases @(
		@{ Id = 'AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE'; Valid = $true },
		@{ Id = 'javascript:alert(1)'; Valid = $false },
		@{ Id = '"><img src=x onerror=alert(1)>'; Valid = $false },
		@{ Id = '00000000-0000-0000-0000-000000000000'; Valid = $false }
	) {
		param($Id, $Valid)
		$result = New-ReportFixture
		$result.GroupId = $Id
		$html = New-AccessReviewHtml -Results @($result) -Context $script:context
		if ($Valid) {
			$links = [regex]::Matches($html, '<a class="entra-link" href="([^"]+)" target="_blank" rel="noopener noreferrer">')
			$links.Count | Should -Be 3
			$links[0].Groups[1].Value | Should -Be 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
		}
		else {
			$html | Should -Not -Match '<a class="entra-link"'
			$html | Should -Match 'Entra link unavailable'
			$html | Should -Not -Match '<img'
		}
	}

	It 'renders each shared group once with all directed edges through deeper nesting' {
		$root = New-ReportFixture -Name 'Root one'
		$other = New-ReportFixture -Name 'Root two'
		$other.GroupId = '77777777-7777-7777-7777-777777777777'
		$child = New-ReportFixture -Name 'Shared child'
		$child.GroupId = '44444444-4444-4444-4444-444444444444'
		$child.ParentGroupIds = @($root.GroupId, $other.GroupId)
		$child.GroupPaths = @(
			@{ GroupIds = @($root.GroupId, $child.GroupId); GroupNames = @('Root one', 'Shared child') },
			@{ GroupIds = @($other.GroupId, $child.GroupId); GroupNames = @('Root two', 'Shared child') }
		)
		$leaf = New-ReportFixture -Name 'Leaf'
		$leaf.GroupId = '55555555-5555-5555-5555-555555555555'
		$leaf.GroupPaths = @(
			@{ GroupIds = @($root.GroupId, $child.GroupId, $leaf.GroupId); GroupNames = @('Root one', 'Shared child', 'Leaf') },
			@{ GroupIds = @($other.GroupId, $child.GroupId, $leaf.GroupId); GroupNames = @('Root two', 'Shared child', 'Leaf') }
		)
		$html = New-AccessReviewHtml -Results @($root, $other, $child, $leaf) -Context $script:context
		[regex]::Matches($html, 'class="topology-node"').Count | Should -Be 4
		[regex]::Matches($html, 'class="topology-edge"').Count | Should -Be 3
		$html | Should -Match 'data-from="node-0" data-to="node-1"'
		$html | Should -Match 'data-from="node-3" data-to="node-1"'
		$html | Should -Match 'data-from="node-1" data-to="node-2"'
		$html | Should -Match 'Shared child / 2 parents'
	}

	It 'shows referenced parents without fabricating a review or population' {
		$child = New-ReportFixture
		$parentId = '77777777-7777-7777-7777-777777777777'
		$child.GroupPaths = @(@{ GroupIds = @($parentId, $child.GroupId); GroupNames = @('<Parent>', 'Child') })
		$html = New-AccessReviewHtml -Results @($child) -Context $script:context
		[regex]::Matches($html, 'class="topology-node"').Count | Should -Be 2
		[regex]::Matches($html, 'class="review-detail" hidden').Count | Should -Be 1
		$html | Should -Match '&lt;Parent&gt;'
		$html | Should -Match 'no review result recorded. Population and stages unavailable'
		$html | Should -Match "Overview/groupId/$parentId"
	}

	It 'shows applications as distinct roots connected only to directly assigned groups' {
		$root = New-ReportFixture -Name 'Assigned group'
		$child = New-ReportFixture -Name 'Nested child'
		$child.GroupId = '44444444-4444-4444-4444-444444444444'
		$child.ParentGroupIds = @($root.GroupId)
		$first = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
		$second = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
		foreach ($result in @($root, $child)) {
			$result.GroupPaths = @(foreach ($app in @($first, $second)) {
				@{
					ServicePrincipalId = $app; ServicePrincipalName = "Payroll <$app>"
					GroupIds = if ($result -eq $root) { @($root.GroupId) } else { @($root.GroupId, $child.GroupId) }
					GroupNames = if ($result -eq $root) { @('Assigned group') } else { @('Assigned group', 'Nested child') }
				}
			})
		}
		$direct = [pscustomobject]@{
			TargetType = 'ApplicationDirectUsers'; ServicePrincipalId = $first; Target = "Payroll <$first>"
			SubjectCount = 3; Status = 'Created'; ReviewName = 'AR Payroll direct users'
		}
		$html = New-AccessReviewHtml -Results @($root, $child, $direct) -Context $script:context
		[regex]::Matches($html, 'class="topology-node application-node"').Count | Should -Be 2
		[regex]::Matches($html, 'class="topology-edge application-edge"').Count | Should -Be 2
		[regex]::Matches($html, 'class="topology-edge"').Count | Should -Be 1
		$html | Should -Match 'data-from="app-0" data-to="node-0"'
		$html | Should -Match 'data-from="app-1" data-to="node-0"'
		$html | Should -Not -Match 'data-from="app-[01]" data-to="node-1"'
		$html | Should -Match '3 directly assigned users'
		$html | Should -Match 'Direct-user review: Created'
		$html | Should -Match 'data-target="review-2"'
		$html | Should -Match "objectId/$first"
		$html | Should -Not -Match '<aaaaaaaa'
	}

	It 'retains discovered applications with no review targets in live and offline reports (offline: <Offline>)' -TestCases @(
		@{ Offline = $false }, @{ Offline = $true }
	) {
		param($Offline)
		$app = @{ ServicePrincipalId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; ServicePrincipalName = 'Empty application'; GroupIds = @(); DirectUserCount = 0 }
		if ($Offline) {
			@(
				@{ Event = 'RunStarted'; RunId = 'empty-app' }
				@{ Event = 'ApplicationDiscovered'; RunId = 'empty-app'; Application = $app }
				@{ Event = 'RunCompleted'; RunId = 'empty-app' }
			) | ForEach-Object { $_ | ConvertTo-Json -Depth 10 -Compress } | Set-Content $script:logPath
			& $script:exporter -LogPath $script:logPath -OutputPath $script:outputPath
			$html = Get-Content $script:outputPath -Raw
		}
		else {
			$script:context.Applications = @($app)
			$html = New-AccessReviewHtml -Results @() -Context $script:context
		}
		[regex]::Matches($html, 'class="topology-node application-node"').Count | Should -Be 1
		$html | Should -Match 'Empty application'
		$html | Should -Match '0 directly assigned users discovered'
		$html | Should -Not -Match '<tr data-status=|class="topology-edge'
	}

	It 'does not invent edges for legacy groups without recorded paths' {
		$result = New-ReportFixture
		$html = New-AccessReviewHtml -Results @($result) -Context $script:context
		[regex]::Matches($html, 'class="topology-node"').Count | Should -Be 1
		[regex]::Matches($html, 'class="topology-edge"').Count | Should -Be 0
		$html | Should -Match 'Resources without recorded connections are not assumed to be unrelated'
	}

	It 'retains a self-membership arrow without duplicating the group card' {
		$result = New-ReportFixture
		$result.GroupPaths = @(@{ GroupIds = @($result.GroupId, $result.GroupId); GroupNames = @('Self-member', 'Self-member') })
		$html = New-AccessReviewHtml -Results @($result) -Context $script:context
		[regex]::Matches($html, 'class="topology-node"').Count | Should -Be 1
		[regex]::Matches($html, 'class="topology-edge"').Count | Should -Be 1
		$html | Should -Match 'data-from="node-0" data-to="node-0"'
	}

	It 'reports incomplete legacy connections and retains other nodes' {
		$result = New-ReportFixture
		$result.GroupPaths = @(@{ GroupIds = @($null, $result.GroupId); GroupNames = @('Unknown parent', 'Child') })
		$html = New-AccessReviewHtml -Results @($result) -Context $script:context
		$html | Should -Match 'Some recorded path IDs are missing'
		[regex]::Matches($html, 'class="topology-node"').Count | Should -Be 1
		[regex]::Matches($html, 'class="topology-edge"').Count | Should -Be 0
	}

	It 'does not invent configuration or direct-user counts for legacy logs' {
		$legacy = [pscustomobject]@{
			GroupId = 'old-group'; ReviewName = 'Old review'; MemberCount = 520
			StageNames = 'Self-review -> Line managers -> Application owners'; Status = 'Created'
		}
		@(
			@{ Event = 'RunStarted'; RunId = 'old-run' },
			@{ Event = 'GroupResult'; RunId = 'old-run'; Result = $legacy },
			@{ Event = 'RunCompleted'; RunId = 'old-run' }
		) | ForEach-Object { $_ | ConvertTo-Json -Compress -Depth 10 } | Set-Content $script:logPath
		& $script:exporter -LogPath $script:logPath -OutputPath $script:outputPath
		$html = Get-Content $script:outputPath -Raw
		$html | Should -Match '520 recorded members \(basis unknown\)'
		$html | Should -Not -Match '520 direct users'
		$html | Should -Match 'Detailed configuration unavailable'
		$html | Should -Not -Match 'class="timeline"'
	}

	It 'shows resource duplicate evidence without calling it a name match or inventing a timeline' {
		$result = New-ReportFixture
		$result.Status = 'Skipped'
		$result.ReviewConfiguration = $null
		$result.MatchedReviewName = 'Other campaign'
		$result | Add-Member DuplicateMatchType Resource
		$result | Add-Member MatchedReviewId '99999999-9999-9999-9999-999999999999'
		$result | Add-Member MatchedReviewScope '/groups/11111111-1111-1111-1111-111111111111/members'
		$html = New-AccessReviewHtml -Results @($result) -Context $script:context
		$html | Should -Match 'Existing resource match'
		$html | Should -Match '<span class="badge">Existing resource match</span>'
		$html | Should -Not -Match '<span class="badge">Existing name match</span>'
		$html | Should -Match '<strong>1</strong><span>Existing resource match'
		$html | Should -Match 'Other campaign'
		$html | Should -Match '99999999-9999-9999-9999-999999999999'
		$html | Should -Match 'Existing stage/reviewer policy not compared or changed'
		$html | Should -Not -Match 'Submitted configuration'
	}

	It 'leaves existing-name matches without an invented running-review timeline' {
		$result = New-ReportFixture
		$result.Status = 'Skipped'
		$result.MatchedReviewName = 'AR Old name'
		$result.ReviewConfiguration = $null
		$html = New-AccessReviewHtml -Results @($result) -Context $script:context
		$html | Should -Match 'Existing name match'
		$html | Should -Match 'Existing configuration not fetched'
		$html | Should -Not -Match 'class="timeline"'
	}

	It 'separates synced, direct-application and manual coverage from group counts' {
		$synced = New-ReportFixture
		$synced.IsSyncedGroup = $true
		$synced.ReviewConfiguration.AutoApplyDecisionsEnabled = $false
		$app = [pscustomobject]@{
			TargetType = 'ApplicationDirectUsers'; Target = 'Payroll'; ReviewName = 'AR Payroll'
			SubjectCount = 3; Status = 'Created'
		}
		$manual = [pscustomobject]@{
			TargetType = 'ServicePrincipalOwner'; Target = 'Payroll owner'; Status = 'Manual review required'
			PrincipalId = 'owner-id'; Reason = 'Ownership requires manual review.'
		}
		$html = New-AccessReviewHtml -Results @($synced, $app, $manual) -Context $script:context
		$html | Should -Match '<strong>470</strong> verified direct-user counts'
		$html | Should -Match '<strong>3</strong> direct-application user counts, separately'
		$html | Should -Match 'Synced / report only'
		$html | Should -Match 'Remediate denied access in on-premises AD'
		$html | Should -Match 'Ownership requires manual review'
	}

	It 'groups owners by application without merging same-named applications or changing audit results (offline: <Offline>)' -TestCases @(
		@{ Offline = $false }, @{ Offline = $true }
	) {
		param($Offline)
		$owners = @(foreach ($item in @(
			@{ Id = 'owner-one'; App = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; Name = '<Owner & One>' }
			@{ Id = 'owner-two'; App = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; Name = 'Owner Two' }
			@{ Id = 'owner-three'; App = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'; Name = 'Owner Three' }
		)) {
			[pscustomobject]@{
				TargetType = 'ServicePrincipalOwner'; Target = 'Payroll'; ServicePrincipalId = $item.App
				Status = 'Manual review required'; PrincipalId = $item.Id; PrincipalType = 'user'
				PrincipalDisplayName = $item.Name; Reason = 'Ownership requires manual review.'
			}
		})
		$results = @((New-ReportFixture)) + $owners
		$before = $results | ConvertTo-Json -Depth 10
		if ($Offline) {
			@(
				@{ Event = 'RunStarted'; RunId = 'owner-run' }
				@{ Event = 'GroupResult'; RunId = 'owner-run'; Result = $results[0] }
				foreach ($owner in $owners) { @{ Event = 'ManualReviewRequired'; RunId = 'owner-run'; Result = $owner } }
				@{ Event = 'RunCompleted'; RunId = 'owner-run' }
			) | ForEach-Object { $_ | ConvertTo-Json -Compress -Depth 10 } | Set-Content $script:logPath
			& $script:exporter -LogPath $script:logPath -OutputPath $script:outputPath
			$html = Get-Content $script:outputPath -Raw
		}
		else { $html = New-AccessReviewHtml -Results $results -Context $script:context }
		$inventory = [regex]::Match($html, '(?s)<tbody id="rows">(.*?)</tbody>').Groups[1].Value
		[regex]::Matches($inventory, '<tr data-status=').Count | Should -Be 3
		[regex]::Matches($inventory, 'data-type="ServicePrincipalOwner"').Count | Should -Be 2
		[regex]::Matches($html, 'class="review-detail" hidden').Count | Should -Be 3
		$html | Should -Match '<strong>2</strong><span>Manual review required'
		$html | Should -Match 'Recorded owners \(2\)'
		$html | Should -Match '2 recorded owners'
		$html | Should -Match '&lt;Owner &amp; One&gt;'
		$html | Should -Not -Match '<Owner'
		foreach ($id in @('owner-one', 'owner-two', 'owner-three')) {
			$html | Should -Match $id
			$inventory | Should -Match $id
		}
		($results | ConvertTo-Json -Depth 10) | Should -Be $before
	}

	It 'keeps legacy owner entries separate when the application ID was not recorded' {
		$results = @(foreach ($name in @('First application', 'Second application')) {
			[pscustomobject]@{
				TargetType = 'ServicePrincipalOwner'; Target = $name; Status = 'Manual review required'
				PrincipalId = 'owner-id'; Reason = 'Ownership requires manual review.'
			}
		})
		$html = New-AccessReviewHtml -Results $results -Context $script:context
		$inventory = [regex]::Match($html, '(?s)<tbody id="rows">(.*?)</tbody>').Groups[1].Value
		[regex]::Matches($inventory, '<tr data-status=').Count | Should -Be 2
		$inventory | Should -Match 'First application'
		$inventory | Should -Match 'Second application'
	}

	It 'rolls up 58 non-user assignments into one searchable row per application in offline export' {
		$app = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
		$other = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
		@(
			@{ Event = 'RunStarted'; RunId = 'assignment-run' }
			foreach ($i in 1..59) {
				@{ Event = 'ManualReviewRequired'; RunId = 'assignment-run'; Result = @{
					TargetType = 'ApplicationNonUserAssignment'; Target = 'Payroll'; Status = 'Manual review required'
					ServicePrincipalId = $(if ($i -eq 59) { $other } else { $app })
					PrincipalDisplayName = "Automation <$i>"; PrincipalType = 'ServicePrincipal'
					PrincipalId = "principal-$i"; AssignmentId = "assignment-$i"; AppRoleId = "role-$i"
				} }
			}
			@{ Event = 'RunCompleted'; RunId = 'assignment-run' }
		) | ForEach-Object { $_ | ConvertTo-Json -Compress -Depth 10 } | Set-Content $script:logPath
		& $script:exporter -LogPath $script:logPath -OutputPath $script:outputPath
		$html = Get-Content $script:outputPath -Raw
		$inventory = [regex]::Match($html, '(?s)<tbody id="rows">(.*?)</tbody>').Groups[1].Value
		[regex]::Matches($inventory, 'data-type="ApplicationNonUserAssignment"').Count | Should -Be 2
		$inventory | Should -Match '58 recorded non-user assignments'
		$inventory | Should -Match 'assignment-58'
		$html | Should -Match 'Recorded non-user assignments \(58\)'
		$html | Should -Match '<td>Automation &lt;58&gt;</td><td>ServicePrincipal</td><td>principal-58</td><td>assignment-58</td><td>role-58</td>'
		$html | Should -Not -Match 'class="stage-flow"'
	}

	It 'visualizes recorded stage order in inventory and topology without inventing legacy dates' {
		$result = New-ReportFixture -Durations @(3, 5, 6)
		$result.ReviewConfiguration.Stages[0].Name = '<Self review>'
		$app = [pscustomobject]@{
			TargetType = 'ApplicationDirectUsers'; Target = 'Payroll'; Status = 'Created'
			ServicePrincipalId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
			StageNames = 'Line managers -> Application owners'
		}
		$html = New-AccessReviewHtml -Results @($result, $app) -Context $script:context
		$inventory = [regex]::Match($html, '(?s)<tbody id="rows">(.*?)</tbody>').Groups[1].Value
		$topology = [regex]::Match($html, '(?s)<section id="topology".*?</section>').Value
		foreach ($section in @($inventory, $topology)) {
			$section | Should -Match '<ol class="stage-flow"[^>]*><li><span>&lt;Self review&gt;</span></li><li><span>Line managers</span></li><li><span>Resource owners \(group owners\)</span></li></ol>'
			$section | Should -Match '<ol class="stage-flow"[^>]*><li><span>Line managers</span></li><li><span>Application owners</span></li></ol>'
		}
		$html | Should -Match '<small>Stage 3</small>'
		[regex]::Matches($html, 'class="timeline"').Count | Should -Be 1
	}

	It 'shows review status proportions including unavailable statuses without counting instances or stages' {
		$definitions = @(
			@{ Status = 'InProgress'; Instances = @(1, 2, 3) }
			@{ Status = 'InProgress'; Instances = @() }
			@{ Status = 'Applied'; Instances = @(1) }
			@{ Status = $null; Instances = @(1, 2) }
		)
		$html = ConvertTo-AccessReviewStatusChartHtml -Definitions $definitions
		$html | Should -Match 'data-review-status="InProgress" data-count="2"'
		$html | Should -Match 'data-review-status="Applied" data-count="1"'
		$html | Should -Match 'data-review-status="Unavailable" data-count="1"'
		$html | Should -Match 'InProgress: 2 \(50%\)'
		$html | Should -Match 'Applied: 1 \(25%\)'
		$html | Should -Match 'conic-gradient\(#386641 0% 25%, #b51f3b 25% 75%, #737373 75% 100%\)'
	}

	It 'keeps chart percentages culture independent and encodes unfamiliar status labels' {
		$culture = [System.Globalization.CultureInfo]::CurrentCulture
		try {
			[System.Globalization.CultureInfo]::CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('sv-SE')
			$html = ConvertTo-AccessReviewStatusChartHtml -Definitions @(
				@{ Status = 'InProgress' }, @{ Status = 'Completed' }, @{ Status = '</div><img src=x onerror=alert(1)>' }
			)
			$html | Should -Match '33\.333333%'
			$html | Should -Match '33\.3%'
			$html | Should -Not -Match '33,3|<img'
			$html | Should -Match '&lt;/div&gt;&lt;img'
		}
		finally { [System.Globalization.CultureInfo]::CurrentCulture = $culture }
	}

	It 'exports recorded configuration and a failed run without importing Graph' {
		Mock Import-Module { throw 'Offline export must not import modules.' }
		$result = New-ReportFixture
		@(
			@{ Event = 'RunStarted'; RunId = 'run-1' },
			@{ Event = 'GroupResult'; RunId = 'run-1'; Result = $result },
			@{ Event = 'RunFailed'; RunId = 'run-1'; Reason = 'Membership denied.' }
		) | ForEach-Object { $_ | ConvertTo-Json -Compress -Depth 10 } | Set-Content $script:logPath
		& $script:exporter -LogPath $script:logPath -OutputPath $script:outputPath
		$html = Get-Content $script:outputPath -Raw
		$html | Should -Match 'Partial or failed run'
		$html | Should -Match 'Membership denied'
		$html | Should -Match '2026-11-22 &rarr; 2026-11-27'
		Should -Invoke Import-Module -Times 0 -Exactly
	}

	It 'rejects <Kind> input without producing an HTML file' -TestCases @(
		@{ Kind = 'malformed JSON' }, @{ Kind = 'mixed runs' }, @{ Kind = 'unidentified mixed runs' }, @{ Kind = 'empty data' }
	) {
		param($Kind)
		switch ($Kind) {
			'malformed JSON' { "{`"Event`":`"RunStarted`"}`nnot-json" | Set-Content $script:logPath }
			'mixed runs' { '{"Event":"RunStarted","RunId":"one"}', '{"Event":"RunStarted","RunId":"two"}' | Set-Content $script:logPath }
			'unidentified mixed runs' { '{"Event":"RunStarted"}', '{"Event":"RunStarted"}' | Set-Content $script:logPath }
			'empty data' { '' | Set-Content $script:logPath }
		}
		{ & $script:exporter -LogPath $script:logPath -OutputPath $script:outputPath } | Should -Throw
		Test-Path $script:outputPath | Should -BeFalse
		if ($Kind -eq 'malformed JSON') {
			{ & $script:exporter -LogPath $script:logPath -OutputPath $script:outputPath } | Should -Throw '*line 2*'
		}
	}

	It 'never overwrites an existing report' {
		Set-Content $script:outputPath 'Existing report'
		{ Write-AccessReviewReport -Path $script:outputPath -Results @() -Context $script:context } | Should -Throw '*already exists*'
		Get-Content $script:outputPath | Should -Be 'Existing report'
	}

	It 'preserves every target and shared path in a 2000-target report' {
		$results = @(for ($i = 0; $i -lt 2000; $i++) {
			$result = New-ReportFixture -Name "AR Scale $i"
			$result.GroupId = "group-$i"
			$result.GroupPaths = @(
				@{ GroupIds = @('root-one', "group-$i"); GroupNames = @('Root one', "Scale $i") },
				@{ GroupIds = @('root-two', "group-$i"); GroupNames = @('Root two', "Scale $i") }
			)
			$result
		})
		$html = New-AccessReviewHtml -Results $results -Context $script:context
		[regex]::Matches($html, '<tr data-status=').Count | Should -Be 2000
		[regex]::Matches($html, 'class="review-detail" hidden').Count | Should -Be 2000
		[regex]::Matches($html, 'class="topology-node"').Count | Should -Be 2002
		[regex]::Matches($html, 'class="topology-edge"').Count | Should -Be 4000
		[regex]::Matches($html, 'Root one -&gt; Scale').Count | Should -Be 2000
		[regex]::Matches($html, 'Root two -&gt; Scale').Count | Should -Be 2000
		$html | Should -Match 'AR Scale 1999'
	}
}
