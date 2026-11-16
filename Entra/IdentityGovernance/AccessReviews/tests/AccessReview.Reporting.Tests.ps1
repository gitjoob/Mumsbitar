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

	It 'does not invent edges for legacy groups without recorded paths' {
		$result = New-ReportFixture
		$html = New-AccessReviewHtml -Results @($result) -Context $script:context
		[regex]::Matches($html, 'class="topology-node"').Count | Should -Be 1
		[regex]::Matches($html, 'class="topology-edge"').Count | Should -Be 0
		$html | Should -Match 'Groups without recorded paths are not assumed to be unrelated'
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
