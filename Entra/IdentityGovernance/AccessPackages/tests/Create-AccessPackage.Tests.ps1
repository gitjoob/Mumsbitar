BeforeAll {
	Import-Module Microsoft.Graph.Authentication
	$script:scriptPath = Join-Path $PSScriptRoot '../Create-AccessPackage.ps1'
	$script:groupId = '11111111-1111-1111-1111-111111111111'
	$script:ownerId = '22222222-2222-2222-2222-222222222222'
	$script:fallbackId = '33333333-3333-3333-3333-333333333333'

	function Invoke-Creation {
		param([switch]$Preview, [guid[]]$DefaultOwners = @())

		& $script:scriptPath -CsvPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
			-FallbackApproverIds $script:fallbackId -DefaultAppOwnerIds $DefaultOwners `
			-WhatIf:$Preview -Confirm:$false -WarningAction SilentlyContinue
	}

	function Get-PostedBody {
		param([string]$Suffix)
		@($script:graphFixture.Writes | Where-Object { $_.Method -eq 'POST' -and $_.Uri.EndsWith($Suffix) })[-1].Body
	}

	function New-WorkbookFixture {
		param([hashtable]$Overrides)

		$columns = [ordered]@{
			Role = 'Reader'; Brand = 'Brand A'; Region = 'EMEA'; 'Potential Teams' = 'Finance'
			'Access Group' = $script:groupId; Description = 'View financial reports.'
			'Access Package Name' = 'Finance reporting - Read reports'
			'Approver 1 List' = '<Employee Manager>'; 'Approver 2' = ''; 'Approver 3' = ''
		}
		foreach ($key in $Overrides.Keys) { $columns[$key] = $Overrides[$key] }
		$path = Join-Path $TestDrive 'package [template].xlsx'
		[pscustomobject]$columns | Export-Excel -Path $path -WorksheetName 'Packages'
		$path
	}
}

Describe 'Create-AccessPackage' {
	BeforeEach {
		$script:csvPath = Join-Path $TestDrive 'packages.csv'
		@"
Name,Description,GroupId,AppOwnerIds
Finance reporting - Read reports,View financial reports.,$script:groupId,$script:ownerId
"@ | Set-Content -LiteralPath $script:csvPath
		$script:graphFixture = @{
			CatalogName = 'Business applications'
			CatalogExists = $true
			DuplicateCatalog = $false
			CatalogState = 'published'
			ExistingPackage = $false
			PaginateExisting = $false
			ResourceExists = $true
			ResourceDelay = 0
			NeverReady = $false
			UserCount = '500'
			DeniedUserId = $null
			NamedGroups = @()
			AddressUsers = @(
				@{ id = $script:ownerId; userPrincipalName = 'owner@example.test'; mail = 'owner-mail@example.test' }
				@{ id = $script:fallbackId; userPrincipalName = 'fallback@example.test'; mail = 'fallback-mail@example.test' }
			)
			Group = @{
				id = $script:groupId
				displayName = 'SG_Finance_Reports'
				groupTypes = @()
				onPremisesSyncEnabled = $false
				mailEnabled = $false
				securityEnabled = $true
				isAssignableToRole = $false
			}
			FailOn = $null
			RequestFailurePattern = $null
			FailureFormat = 'ErrorDetails'
			FailureStatus = 400
			FailRollback = $false
			Writes = [System.Collections.Generic.List[object]]::new()
			Deleted = [System.Collections.Generic.List[string]]::new()
		}

		Mock Get-MgContext {
			@{
				AuthType = 'Delegated'
				Scopes = @('EntitlementManagement.ReadWrite.All', 'Group.Read.All', 'User.ReadBasic.All')
			}
		}
		Mock Start-Sleep {}
		$fixture = $script:graphFixture
		$fixtureGroupId = $script:groupId
		Mock Invoke-MgGraphRequest ({
			param($Uri, $Method, $Body)

			$uriText = [string]$Uri
			$decodedUri = [uri]::UnescapeDataString($uriText)
			if ($fixture.RequestFailurePattern -and $uriText -match $fixture.RequestFailurePattern) {
				$response = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]$fixture.FailureStatus)
				$response.Headers.Add('request-id', 'test-request-id')
				$response.Content = [System.Net.Http.StringContent]::new(
					'{"error":{"code":"InvalidRequest","message":"Detailed Graph rejection for this request."}}'
				)
				$exception = [Microsoft.Graph.PowerShell.Authentication.Helpers.HttpResponseException]::new(
					'Response status code does not indicate success.', $response
				)
				$record = [System.Management.Automation.ErrorRecord]::new(
					$exception, 'GraphRequestFailed', [System.Management.Automation.ErrorCategory]::InvalidOperation, $Uri
				)
				if ($fixture.FailureFormat -eq 'ErrorDetails') {
					$record.ErrorDetails = [System.Management.Automation.ErrorDetails]::new(
						'{"error":{"code":"InvalidRequest","message":"Detailed Graph rejection for this request."}}'
					)
				}
				throw $record
			}
			if ($Method -eq 'GET') {
				if ($uriText -match '/users/([^?]+)') {
					$id = $Matches[1]
					if ($id -eq $fixture.DeniedUserId) { throw 'User not found or forbidden.' }
					return @{ id = $id; displayName = 'Approver' }
				}
				if ($decodedUri -match "/users\?\`$filter=userPrincipalName eq '([^']+)' or mail eq '\1'&\`$select=id,userPrincipalName,mail$") {
					$address = $Matches[1]
					return @{ value = @($fixture.AddressUsers | Where-Object { $_.userPrincipalName -eq $address -or $_.mail -eq $address }) }
				}
				if ($uriText -match '/transitiveMembers/microsoft.graph.user/\$count$') {
					return $fixture.UserCount
				}
				if ($uriText -match '/groups/[^/?]+\?') {
					return $fixture.Group
				}
				if ($decodedUri.Contains("/groups?`$filter=displayName eq 'Finance''s Readers'&")) {
					return @{ value = $fixture.NamedGroups }
				}
				if ($uriText -match '/catalogs\?') {
					$items = @()
					if ($fixture.CatalogExists) {
						$items += @{ id = 'catalog-id'; state = $fixture.CatalogState }
						if ($fixture.DuplicateCatalog) { $items += @{ id = 'other-catalog-id'; state = 'published' } }
					}
					return @{ value = $items }
				}
				if ($uriText -match '/accessPackages\?') {
					if ($fixture.PaginateExisting -and $decodedUri -notmatch 'next=1') {
						return @{ value = @(); '@odata.nextLink' = '/v1.0/identityGovernance/entitlementManagement/accessPackages?next=1' }
					}
					$items = if ($fixture.ExistingPackage) { @(@{ id = 'existing-package-id' }) } else { @() }
					return @{ value = $items }
				}
				if ($uriText -match '/resourceRoles\?') {
					if ($decodedUri -notmatch "originSystem eq 'AadGroup' and resource/id eq 'resource-id'") {
						throw 'Resource-role lookup must filter by origin system and resource ID.'
					}
					return @{
						value = @(
							@{ id = 'owner-role'; originId = "Owner_$fixtureGroupId"; originSystem = 'AadGroup' },
							@{ id = 'member-role'; originId = "Member_$fixtureGroupId"; originSystem = 'AadGroup' }
						)
					}
				}
				if ($uriText -match '/resources\?') {
					if ($fixture.ResourceDelay -gt 0) {
						$fixture.ResourceDelay--
						return @{ value = @() }
					}
					if (-not $fixture.ResourceExists -or $fixture.NeverReady) { return @{ value = @() } }
					return @{
						value = @(
							@{
								id = 'resource-id'
								originId = $fixtureGroupId
								originSystem = 'AadGroup'
								roles = @()
								scopes = @(@{ id = 'root-scope'; originId = $fixtureGroupId; originSystem = 'AadGroup'; isRootScope = $true })
							}
						)
					}
				}
			}
			elseif ($Method -eq 'POST') {
				$parsedBody = $Body | ConvertFrom-Json
				$fixture.Writes.Add([pscustomobject]@{ Method = "$Method"; Uri = $uriText; Body = $parsedBody })
				if ($fixture.FailOn -and $uriText.EndsWith($fixture.FailOn)) { throw "Simulated $($fixture.FailOn) failure." }
				if ($uriText.EndsWith('/catalogs')) {
					$fixture.CatalogExists = $true
					return @{ id = 'new-catalog-id'; state = 'published' }
				}
				if ($uriText.EndsWith('/resourceRequests')) {
					$fixture.ResourceExists = $true
					return @{ id = 'resource-request-id' }
				}
				if ($uriText.EndsWith('/accessPackages')) { return @{ id = 'new-package-id' } }
				if ($uriText.EndsWith('/resourceRoleScopes')) { return @{ id = 'binding-id' } }
				if ($uriText.EndsWith('/assignmentPolicies')) { return @{ id = 'new-policy-id' } }
			}
			elseif ($Method -eq 'DELETE') {
				$fixture.Deleted.Add($uriText)
				if ($fixture.FailRollback) { throw 'Simulated delete failure.' }
				return
			}
			throw "Unexpected Graph request: $Method $uriText"
		}.GetNewClosure())
	}

	Context 'Excel workbook' -Skip:(-not (Get-Module -ListAvailable -Name ImportExcel)) {
		BeforeAll { Import-Module ImportExcel -WarningAction SilentlyContinue }

		It 'resolves UPN and mail values to user IDs for ordered approval stages' {
			$script:graphFixture.NamedGroups = @(@{ id = $script:groupId; displayName = "Finance's Readers" })
			$path = New-WorkbookFixture @{
				'Access Group' = "Finance's Readers"; 'Approver 1 List' = ' < employee   MANAGER > '
				'Approver 2' = "owner@example.test;`nowner-mail@example.test"; 'Approver 3' = 'fallback-mail@example.test'
			}
			$result = & $script:scriptPath -InputPath $path -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -Confirm:$false -WarningAction SilentlyContinue
			$result.Status | Should -Be 'Created'
			$result.GroupId | Should -Be $script:groupId
			$result.ApprovalStages | Should -Be 3
			$package = Get-PostedBody '/accessPackages'
			$package.displayName | Should -Be 'Finance reporting - Read reports'
			$package.description -split '\r?\n' | Should -Be @('View financial reports.', 'Role: Reader', 'Region: EMEA', 'Brand: Brand A', 'Potential Teams: Finance')
			$stages = (Get-PostedBody '/assignmentPolicies').requestApprovalSettings.stages
			$stages[0].primaryApprovers[0].'@odata.type' | Should -Be '#microsoft.graph.targetManager'
			$stages[0].fallbackPrimaryApprovers[0].userId | Should -Be $script:fallbackId
			$stages[1].primaryApprovers | Should -HaveCount 1
			$stages[1].primaryApprovers[0].userId | Should -Be $script:ownerId
			$stages[2].primaryApprovers[0].userId | Should -Be $script:fallbackId
		}

		It 'stops on an unreadable workbook before worksheet validation or Graph access' {
			$script:csvPath = Join-Path $TestDrive 'not-a-workbook.xlsx'
			'not an Excel package' | Set-Content -LiteralPath $script:csvPath
			{ Invoke-Creation -Preview } | Should -Throw '*Could not read Excel workbook*Import error:*'
			Should -Invoke Get-MgContext -Exactly -Times 0
			Should -Invoke Invoke-MgGraphRequest -Exactly -Times 0
		}

		It 'rejects addresses with <MatchCount> exact matches without provisioning' -TestCases @(
			@{ MatchCount = 0 }, @{ MatchCount = 2 }
		) {
			param($MatchCount)
			if ($MatchCount -eq 2) {
				$script:graphFixture.AddressUsers[1].mail = 'owner@example.test'
			}
			else { $script:graphFixture.AddressUsers = @() }
			$script:csvPath = New-WorkbookFixture @{ 'Approver 2' = 'owner@example.test' }
			$result = Invoke-Creation
			$result.Status | Should -Be 'Failed'
			$result.Reason | Should -Match "matched $MatchCount Entra users"
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'previews a manager-only workbook row without Graph writes' {
			$script:csvPath = New-WorkbookFixture @{}
			$result = Invoke-Creation -Preview
			$result.Status | Should -Be 'Would be created'
			$result.ApprovalStages | Should -Be 1
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'retains the second-stage requirement above 500 users (default owner: <HasDefault>)' -TestCases @(
			@{ HasDefault = $false }, @{ HasDefault = $true }
		) {
			param($HasDefault)
			$script:graphFixture.UserCount = '501'
			$script:csvPath = New-WorkbookFixture @{}
			if ($HasDefault) {
				$result = Invoke-Creation -Preview -DefaultOwners $script:ownerId
				$result.Status | Should -Be 'Would be created'
				$result.ApprovalStages | Should -Be 2
			}
			else {
				$result = Invoke-Creation -Preview
				$result.Status | Should -Be 'Failed'
				$result.Reason | Should -Match 'require Approver 2 or DefaultAppOwnerIds'
			}
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'rejects unsafe workbook input before provisioning (<Case>)' -TestCases @(
			@{ Case = 'unsupported manager token'; Overrides = @{ 'Approver 1 List' = '<Employee Manager> plus owner' }; Reason = 'must be <Employee Manager>' }
			@{ Case = 'missing stage two'; Overrides = @{ 'Approver 3' = 'owner@example.test' }; Reason = 'Approver 3 requires Approver 2' }
			@{ Case = 'ambiguous group'; Overrides = @{ 'Access Group' = "Finance's Readers" }; Reason = 'matched 2 groups' }
		) {
			param($Case, $Overrides, $Reason)
			$script:graphFixture.NamedGroups = @(
				@{ id = $script:groupId; displayName = "Finance's Readers" }
				@{ id = '77777777-7777-7777-7777-777777777777'; displayName = "Finance's Readers" }
			)
			$script:csvPath = New-WorkbookFixture $Overrides
			$result = Invoke-Creation
			$result.Status | Should -Be 'Failed'
			$result.Reason | Should -Match $Reason
			$script:graphFixture.Writes | Should -HaveCount 0
		}
	}

	It 'creates a one-stage policy at exactly 500 users with all requested Graph settings' {
		$results = @(Invoke-Creation)
		$results.Count | Should -Be 1
		$results[0].Status | Should -Be 'Created'
		$results[0].AccessPackageId | Should -Be 'new-package-id'
		$results[0].PolicyId | Should -Be 'new-policy-id'
		$results[0].UserCount | Should -Be 500
		$results[0].ApprovalStages | Should -Be 1
		$results[0].ManualAction | Should -Match 'Bypass approval stage'
		$policy = Get-PostedBody '/assignmentPolicies'
		$policy.allowedTargetScope | Should -Be 'allMemberUsers'
		$policy.description | Should -Match 'excluding guests'
		$policy.accessPackage.id | Should -Be 'new-package-id'
		$policy.expiration.type | Should -Be 'afterDuration'
		$policy.expiration.duration | Should -Be 'P365D'
		$policy.requestorSettings.enableTargetsToSelfAddAccess | Should -BeTrue
		$policy.requestorSettings.enableTargetsToSelfUpdateAccess | Should -BeTrue
		$policy.requestorSettings.allowCustomAssignmentSchedule | Should -BeTrue
		$policy.requestorSettings.enableOnBehalfRequestorsToAddAccess | Should -BeTrue
		$policy.requestorSettings.enableOnBehalfRequestorsToUpdateAccess | Should -BeTrue
		$policy.requestorSettings.onBehalfRequestors.Count | Should -Be 1
		$policy.requestorSettings.onBehalfRequestors[0].'@odata.type' | Should -Be '#microsoft.graph.targetManager'
		$policy.requestorSettings.onBehalfRequestors[0].managerLevel | Should -Be 1
		$policy.requestApprovalSettings.isApprovalRequiredForAdd | Should -BeTrue
		$policy.requestApprovalSettings.isApprovalRequiredForUpdate | Should -BeTrue
		$policy.requestApprovalSettings.isRequestorJustificationRequired | Should -BeTrue
		$policy.requestApprovalSettings.stages.Count | Should -Be 1
		$stage = $policy.requestApprovalSettings.stages[0]
		$stage.durationBeforeAutomaticDenial | Should -Be 'P14D'
		$stage.isEscalationEnabled | Should -BeFalse
		$stage.durationBeforeEscalation | Should -Be 'PT0S'
		$stage.escalationApprovers | Should -HaveCount 0
		$stage.fallbackEscalationApprovers | Should -HaveCount 0
		$stage.primaryApprovers[0].'@odata.type' | Should -Be '#microsoft.graph.targetManager'
		$stage.fallbackPrimaryApprovers[0].userId | Should -Be $script:fallbackId
		$policy.notificationSettings.isAssignmentNotificationDisabled | Should -BeFalse
		$policy.reviewSettings | Should -BeNullOrEmpty
		$policy.PSObject.Properties.Name | Should -Contain 'reviewSettings'
		$policy.questions | Should -HaveCount 0
		$binding = Get-PostedBody '/resourceRoleScopes'
		$binding.role.id | Should -Be 'member-role'
		$binding.role.originId | Should -Be "Member_$script:groupId"
		$binding.role.resource.id | Should -Be 'resource-id'
		$binding.scope.id | Should -Be 'root-scope'
		Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
			$Method -eq 'GET' -and [uri]::UnescapeDataString("$Uri") -match "/resourceRoles\?.*resource/id eq 'resource-id'"
		}
		Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
			$Method -eq 'GET' -and "$Uri" -match '/transitiveMembers/microsoft.graph.user/\$count$' -and $Headers.ConsistencyLevel -eq 'eventual'
		}
	}

	It 'uses two stages at 501 users and supports multiple explicit owners' {
		$script:graphFixture.UserCount = '501'
		$secondOwner = '44444444-4444-4444-4444-444444444444'
		(Get-Content -LiteralPath $script:csvPath) -replace "$script:ownerId$", "$script:ownerId;$secondOwner" |
			Set-Content -LiteralPath $script:csvPath
		$result = Invoke-Creation
		$result.ApprovalStages | Should -Be 2
		$stages = (Get-PostedBody '/assignmentPolicies').requestApprovalSettings.stages
		$stages.Count | Should -Be 2
		$stages[1].primaryApprovers.userId | Should -Be @($script:ownerId, $secondOwner)
		$stages[1].durationBeforeAutomaticDenial | Should -Be 'P14D'
		$stages[1].durationBeforeEscalation | Should -Be 'PT0S'
		$stages[1].escalationApprovers | Should -HaveCount 0
		$stages[1].fallbackEscalationApprovers | Should -HaveCount 0
	}

	It 'uses one manager stage for an empty group without requiring owners' {
		$script:graphFixture.UserCount = '0'
		(Get-Content -LiteralPath $script:csvPath) -replace "$script:ownerId$", '' | Set-Content -LiteralPath $script:csvPath
		$result = Invoke-Creation
		$result.Status | Should -Be 'Created'
		$result.ApprovalStages | Should -Be 1
	}

	It 'uses default owners when the row is empty' {
		$script:graphFixture.UserCount = '501'
		(Get-Content -LiteralPath $script:csvPath) -replace "$script:ownerId$", '' | Set-Content -LiteralPath $script:csvPath
		$result = Invoke-Creation -DefaultOwners $script:ownerId
		$result.Status | Should -Be 'Created'
		(Get-PostedBody '/assignmentPolicies').requestApprovalSettings.stages[1].primaryApprovers[0].userId | Should -Be $script:ownerId
	}

	It 'overrides default owners rather than merging them' {
		$script:graphFixture.UserCount = '501'
		$null = Invoke-Creation -DefaultOwners $script:fallbackId
		$owners = (Get-PostedBody '/assignmentPolicies').requestApprovalSettings.stages[1].primaryApprovers
		$owners.Count | Should -Be 1
		$owners[0].userId | Should -Be $script:ownerId
	}

	It 'fails a large group without application owners before any writes' {
		$script:graphFixture.UserCount = '501'
		(Get-Content -LiteralPath $script:csvPath) -replace "$script:ownerId$", '' | Set-Content -LiteralPath $script:csvPath
		$result = Invoke-Creation
		$result.Status | Should -Be 'Failed'
		$result.Reason | Should -Match 'require AppOwnerIds'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'creates a missing catalog and imports the group with bounded readiness retries' {
		$script:graphFixture.CatalogExists = $false
		$script:graphFixture.ResourceExists = $false
		$script:graphFixture.ResourceDelay = 2
		$result = Invoke-Creation
		$result.Status | Should -Be 'Created'
		$result.CatalogId | Should -Be 'new-catalog-id'
		$catalog = Get-PostedBody '/catalogs'
		$catalog.displayName | Should -Be $script:graphFixture.CatalogName
		$catalog.state | Should -Be 'published'
		$catalog.isExternallyVisible | Should -BeTrue
		$resourceRequest = Get-PostedBody '/resourceRequests'
		$resourceRequest.requestType | Should -Be 'adminAdd'
		$resourceRequest.resource.originId | Should -Be $script:groupId
		$resourceRequest.catalog.id | Should -Be 'new-catalog-id'
		Should -Invoke Start-Sleep -Exactly -Times 2
	}

	It 'does not rewrite an existing catalog or import an existing resource' {
		$result = Invoke-Creation
		$result.Status | Should -Be 'Created'
		@($script:graphFixture.Writes | Where-Object { $_.Uri -match '/(catalogs|resourceRequests)$' }).Count | Should -Be 0
	}

	It 'makes no writes in WhatIf with a <CatalogExists> existing catalog' -TestCases @(
		@{ CatalogExists = $true },
		@{ CatalogExists = $false }
	) {
		param($CatalogExists)
		$script:graphFixture.CatalogExists = $CatalogExists
		$script:graphFixture.ResourceExists = $false
		$result = Invoke-Creation -Preview
		$result.Status | Should -Be 'Would be created'
		$result.ApprovalStages | Should -Be 1
		$result.ManualAction | Should -Match 'Bypass approval'
		Should -Invoke Invoke-MgGraphRequest -Exactly -Times 0 -ParameterFilter { $Method -ne 'GET' }
	}

	It 'skips an existing package found on a later page without modifying it' {
		$script:graphFixture.ExistingPackage = $true
		$script:graphFixture.PaginateExisting = $true
		$result = Invoke-Creation
		$result.Status | Should -Be 'Skipped'
		$result.AccessPackageId | Should -Be 'existing-package-id'
		$script:graphFixture.Writes.Count | Should -Be 0
		$script:graphFixture.Deleted.Count | Should -Be 0
	}

	It 'skips duplicate names within the CSV, including during preview' {
		Add-Content -LiteralPath $script:csvPath -Value "Finance reporting - Read reports,Duplicate row.,$script:groupId,$script:ownerId"
		$results = @(Invoke-Creation -Preview)
		$results.Count | Should -Be 2
		$results[0].Status | Should -Be 'Would be created'
		$results[1].Status | Should -Be 'Skipped'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'continues after an invalid row without creating a catalog for that row' {
		$script:graphFixture.CatalogExists = $false
		$script:graphFixture.ResourceExists = $false
		@"
Name,Description,GroupId,AppOwnerIds
Invalid group,Invalid.,not-a-guid,$script:ownerId
Finance reporting - Read reports,View reports.,$script:groupId,$script:ownerId
"@ | Set-Content -LiteralPath $script:csvPath
		$results = @(Invoke-Creation)
		$results.Count | Should -Be 2
		$results[0].Status | Should -Be 'Failed'
		$results[1].Status | Should -Be 'Created'
		@($script:graphFixture.Writes | Where-Object { $_.Uri.EndsWith('/catalogs') }).Count | Should -Be 1
	}

	It 'rejects forbidden naming term <Name> before any writes' -TestCases @(
		@{ Name = 'AP Finance reporting' },
		@{ Name = 'SG_Finance_Reports' },
		@{ Name = 'Finance Access Package' },
		@{ Name = 'Finance access-package' }
	) {
		param($Name)
		@"
Name,Description,GroupId,AppOwnerIds
$Name,View reports.,$script:groupId,$script:ownerId
"@ | Set-Content -LiteralPath $script:csvPath
		$result = Invoke-Creation
		$result.Status | Should -Be 'Failed'
		$result.Reason | Should -Match 'must not contain'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'preserves business wording and correctly escapes apostrophes and query characters' {
		$script:graphFixture.CatalogName = "People's applications & services"
		@"
Name,Description,GroupId,AppOwnerIds
People's reporting & insights,Read people's reports.,$script:groupId,$script:ownerId
"@ | Set-Content -LiteralPath $script:csvPath
		$result = Invoke-Creation
		$result.Status | Should -Be 'Created'
		$package = Get-PostedBody '/accessPackages'
		$package.displayName | Should -Be "People's reporting & insights"
		$package.description | Should -Be "Read people's reports."
		Should -Invoke Invoke-MgGraphRequest -Exactly -Times 1 -ParameterFilter {
			$Method -eq 'GET' -and [uri]::UnescapeDataString("$Uri") -match "displayName eq 'People''s applications & services'"
		}
	}

	It 'names the package from Application and Role and puts remaining details in the description' {
		@"
Name,Description,GroupId,AppOwnerIds,Application,Role,Region,Comments
Ignored name,View reports.,$script:groupId,$script:ownerId,Finance reporting,Report reader,Nordics,Needs manager sign-off
"@ | Set-Content -LiteralPath $script:csvPath
		$result = Invoke-Creation
		$result.Status | Should -Be 'Created'
		$result.Name | Should -Be 'Finance reporting - Report reader'
		$package = Get-PostedBody '/accessPackages'
		$package.displayName | Should -Be 'Finance reporting - Report reader'
		$package.description | Should -Be ("View reports.", "Region: Nordics", "Comments: Needs manager sign-off" -join [Environment]::NewLine)
	}

	It 'works without Name or Description columns and omits blank optional values' {
		@"
GroupId,AppOwnerIds,Application,Role,Region,Comments
$script:groupId,$script:ownerId,Finance reporting,Report reader,,
"@ | Set-Content -LiteralPath $script:csvPath
		$result = Invoke-Creation
		$result.Status | Should -Be 'Created'
		$package = Get-PostedBody '/accessPackages'
		$package.displayName | Should -Be 'Finance reporting - Report reader'
		$package.description | Should -Be 'Finance reporting - Report reader'
	}

	It 'keeps Name and records a lone <Column> in the description' -TestCases @(
		@{ Column = 'Application'; Value = 'Finance reporting' },
		@{ Column = 'Role'; Value = 'Report reader' }
	) {
		param($Column, $Value)
		@"
Name,Description,GroupId,AppOwnerIds,$Column
Finance reporting - Read reports,View reports.,$script:groupId,$script:ownerId,$Value
"@ | Set-Content -LiteralPath $script:csvPath
		$result = Invoke-Creation
		$result.Status | Should -Be 'Created'
		$package = Get-PostedBody '/accessPackages'
		$package.displayName | Should -Be 'Finance reporting - Read reports'
		$package.description | Should -Be ("View reports.", "${Column}: $Value" -join [Environment]::NewLine)
	}

	It 'applies naming rules to the generated name' {
		@"
GroupId,AppOwnerIds,Application,Role
$script:groupId,$script:ownerId,SG Finance,Reader
"@ | Set-Content -LiteralPath $script:csvPath
		$result = Invoke-Creation
		$result.Status | Should -Be 'Failed'
		$result.Reason | Should -Match 'must not contain'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'rejects a CSV with neither Name nor Application and Role columns' {
		"GroupId,AppOwnerIds,Application`n$script:groupId,$script:ownerId,Finance" | Set-Content -LiteralPath $script:csvPath
		{ Invoke-Creation } | Should -Throw "*'Name' column, or both*"
	}

	It 'rejects unmanageable group type <Kind>' -TestCases @(
		@{ Kind = 'synced' },
		@{ Kind = 'dynamic' },
		@{ Kind = 'mail-enabled security' },
		@{ Kind = 'role-assignable' }
	) {
		param($Kind)
		switch ($Kind) {
			'synced' { $script:graphFixture.Group.onPremisesSyncEnabled = $true }
			'dynamic' { $script:graphFixture.Group.groupTypes = @('DynamicMembership') }
			'mail-enabled security' { $script:graphFixture.Group.mailEnabled = $true }
			'role-assignable' { $script:graphFixture.Group.isAssignableToRole = $true }
		}
		$result = Invoke-Creation
		$result.Status | Should -Be 'Failed'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'accepts assigned Microsoft 365 groups' {
		$script:graphFixture.Group.groupTypes = @('Unified')
		$script:graphFixture.Group.mailEnabled = $true
		$script:graphFixture.Group.securityEnabled = $false
		(Invoke-Creation).Status | Should -Be 'Created'
	}

	It 'fails explicitly instead of assuming zero when the count is invalid' {
		$script:graphFixture.UserCount = 'not-a-count'
		$result = Invoke-Creation
		$result.Status | Should -Be 'Failed'
		$result.Reason | Should -Match 'invalid transitive user count'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'validates owner existence before provisioning' {
		$script:graphFixture.DeniedUserId = $script:ownerId
		$result = Invoke-Creation
		$result.Status | Should -Be 'Failed'
		$result.Reason | Should -Match 'User not found'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'stops when the mandatory fallback user cannot be read' {
		$script:graphFixture.DeniedUserId = $script:fallbackId
		{ Invoke-Creation } | Should -Throw '*User not found*'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'stops before writes for ambiguous catalog names' {
		$script:graphFixture.DuplicateCatalog = $true
		{ Invoke-Creation } | Should -Throw '*Multiple catalogs*'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'stops before writes for an unpublished catalog' {
		$script:graphFixture.CatalogState = 'unpublished'
		{ Invoke-Creation } | Should -Throw '*not published*'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'fails before creating a package when resource import never becomes ready' {
		$script:graphFixture.ResourceExists = $false
		$script:graphFixture.NeverReady = $true
		$result = Invoke-Creation
		$result.Status | Should -Be 'Failed'
		$result.Reason | Should -Match 'retry window'
		@($script:graphFixture.Writes | Where-Object { $_.Uri.EndsWith('/accessPackages') }).Count | Should -Be 0
		Should -Invoke Start-Sleep -Exactly -Times 5
	}

	It 'rolls back only its newly created package after <Failure> fails' -TestCases @(
		@{ Failure = '/resourceRoleScopes' },
		@{ Failure = '/assignmentPolicies' }
	) {
		param($Failure)
		$script:graphFixture.FailOn = $Failure
		$result = Invoke-Creation
		$result.Status | Should -Be 'Failed'
		$result.AccessPackageId | Should -BeNullOrEmpty
		$result.Reason | Should -Match 'newly created package was removed'
		$script:graphFixture.Deleted.Count | Should -Be 1
		$script:graphFixture.Deleted[0] | Should -Be '/v1.0/identityGovernance/entitlementManagement/accessPackages/new-package-id'
	}

	It 'retains the incomplete package ID and surfaces rollback failures' {
		$script:graphFixture.FailOn = '/assignmentPolicies'
		$script:graphFixture.FailRollback = $true
		$result = Invoke-Creation
		$result.Status | Should -Be 'Failed'
		$result.AccessPackageId | Should -Be 'new-package-id'
		$result.Reason | Should -Match 'ROLLBACK FAILED'
		$result.Reason | Should -Match 'manually'
	}

	It 'preserves Graph errors from <Format> with the failed request, status and request ID' -TestCases @(
		@{ Format = 'ErrorDetails' },
		@{ Format = 'ResponseContent' }
	) {
		param($Format)
		$script:graphFixture.RequestFailurePattern = '/resources\?'
		$script:graphFixture.FailureFormat = $Format
		$result = Invoke-Creation
		$result.Status | Should -Be 'Failed'
		$result.Reason | Should -Match 'GET /v1.0/identityGovernance/entitlementManagement/catalogs/catalog-id/resources'
		$result.Reason | Should -Match 'HTTP 400'
		$result.Reason | Should -Match 'Request ID: test-request-id'
		$result.Reason | Should -Match '"code":"InvalidRequest"'
		$result.Reason | Should -Match 'Detailed Graph rejection for this request'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'identifies a missing approver as a user GET failure rather than a package creation failure' {
		$script:graphFixture.RequestFailurePattern = "/users/$script:ownerId"
		$script:graphFixture.FailureStatus = 404
		$result = Invoke-Creation
		$result.Status | Should -Be 'Failed'
		$result.Reason | Should -Match "GET /v1.0/users/$script:ownerId"
		$result.Reason | Should -Match 'HTTP 404'
		$result.Reason | Should -Match 'Detailed Graph rejection'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'retains detailed policy rejection information after rollback' {
		$script:graphFixture.RequestFailurePattern = '/assignmentPolicies$'
		$result = Invoke-Creation
		$result.Status | Should -Be 'Failed'
		$result.Reason | Should -Match 'POST /v1.0/identityGovernance/entitlementManagement/assignmentPolicies'
		$result.Reason | Should -Match 'Detailed Graph rejection'
		$result.Reason | Should -Match 'newly created package was removed'
		$script:graphFixture.Deleted.Count | Should -Be 1
	}

	It 'preserves detailed errors for catalog lookup before processing rows' {
		$script:graphFixture.RequestFailurePattern = '/catalogs\?'
		$errorRecord = { Invoke-Creation } | Should -Throw '*Detailed Graph rejection*' -PassThru
		$errorRecord.Exception.Message | Should -Match 'GET /v1.0/identityGovernance/entitlementManagement/catalogs'
		$errorRecord.Exception.Message | Should -Match 'Request ID: test-request-id'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'rejects a missing CSV header before provisioning' {
		"Name,Description,GroupId`nFinance,Read reports.,$script:groupId" | Set-Content -LiteralPath $script:csvPath
		{ Invoke-Creation } | Should -Throw "*'AppOwnerIds' column*"
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'rejects an empty CSV before provisioning' {
		'Name,Description,GroupId,AppOwnerIds' | Set-Content -LiteralPath $script:csvPath
		{ Invoke-Creation } | Should -Throw '*contains no access packages*'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'requires a Graph connection' {
		Mock Get-MgContext { $null }
		{ Invoke-Creation } | Should -Throw '*Connect to Microsoft Graph first*'
		$script:graphFixture.Writes.Count | Should -Be 0
	}

	It 'rejects insufficient delegated scopes' {
		Mock Get-MgContext { @{ AuthType = 'Delegated'; Scopes = @('User.Read') } }
		$errorRecord = { Invoke-Creation } | Should -Throw '*permission check failed*' -PassThru
		$message = $errorRecord.Exception.Message
		$message | Should -Match 'Create catalogs, access packages and policies: EntitlementManagement.ReadWrite.All'
		$message | Should -Match 'Read groups and count their members: Group.Read.All OR Group.ReadWrite.All'
		$message | Should -Match 'Validate approver users: User.ReadBasic.All OR User.Read.All'
		$message | Should -Match 'Current scopes: User.Read'
		$message | Should -Match 'Disconnect-MgGraph'
		$message | Should -Match "Connect-MgGraph -Scopes 'EntitlementManagement.ReadWrite.All','Group.Read.All','User.ReadBasic.All'"
		$message | Should -Match 'administrator consent'
		$message | Should -Match 'appropriate active role'
		$script:graphFixture.Writes.Count | Should -Be 0
		Should -Invoke Invoke-MgGraphRequest -Exactly -Times 0
	}

	It 'reports only missing permission categories' {
		Mock Get-MgContext {
			@{ AuthType = 'Delegated'; Scopes = @('EntitlementManagement.ReadWrite.All', 'Group.Read.All') }
		}
		$errorRecord = { Invoke-Creation } | Should -Throw '*permission check failed*' -PassThru
		$errorRecord.Exception.Message | Should -Match 'Validate approver users:'
		$errorRecord.Exception.Message | Should -Not -Match 'Create catalogs, access packages and policies:'
		$errorRecord.Exception.Message | Should -Not -Match 'Read groups and count their members:'
		Should -Invoke Invoke-MgGraphRequest -Exactly -Times 0
	}

	It 'handles empty delegated scopes with a clear diagnostic' {
		Mock Get-MgContext { @{ AuthType = 'Delegated'; Scopes = @() } }
		$errorRecord = { Invoke-Creation } | Should -Throw '*permission check failed*' -PassThru
		$errorRecord.Exception.Message | Should -Match 'Current scopes: \(none\)'
		Should -Invoke Invoke-MgGraphRequest -Exactly -Times 0
	}

	It 'accepts higher-privileged alternatives for group and user reads' {
		Mock Get-MgContext {
			@{ AuthType = 'Delegated'; Scopes = @('EntitlementManagement.ReadWrite.All', 'Directory.Read.All') }
		}
		(Invoke-Creation).Status | Should -Be 'Created'
	}
}
