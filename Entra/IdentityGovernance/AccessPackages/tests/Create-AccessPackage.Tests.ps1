BeforeAll {
	Import-Module Microsoft.Graph.Authentication
	Import-Module Microsoft.Graph.Groups
	$script:scriptPath = Join-Path $PSScriptRoot '../Create-AccessPackage.ps1'
	$script:noBackfillScriptPath = Join-Path $TestDrive 'Create-AccessPackage.NoBackfill.ps1'
	(Get-Content -LiteralPath $script:scriptPath -Raw).Replace('$CreateAssignments = $true', '$CreateAssignments = $false') |
		Set-Content -LiteralPath $script:noBackfillScriptPath
	Copy-Item (Join-Path $PSScriptRoot '../AccessPackage.Group.ps1') $TestDrive
	$script:groupId = '11111111-1111-1111-1111-111111111111'
	$script:ownerId = '22222222-2222-2222-2222-222222222222'
	$script:fallbackId = '33333333-3333-3333-3333-333333333333'

	function Invoke-Creation {
		param([switch]$Preview, [guid[]]$DefaultOwners = @(), [switch]$Backfill)

		$path = if ($Backfill) { $script:scriptPath } else { $script:noBackfillScriptPath }
		& $path -CsvPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
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
		$path = Join-Path $TestDrive "package [template] $([guid]::NewGuid()).xlsx"
		[pscustomobject]$columns | Export-Excel -Path $path -WorksheetName 'Packages'
		$path
	}

	function New-DiscoveryWorkbookFixture {
		param([switch]$Combined, [switch]$Current, [switch]$Shared)

		$base = Join-Path $TestDrive "discovery-$([guid]::NewGuid())"
		$packages = Join-Path $base 'AccessPackages'
		$reviews = Join-Path $base 'AccessReviews'
		$null = New-Item -ItemType Directory -Path $packages, $reviews
		foreach ($file in 'Create-AccessPackage.ps1', 'AccessPackage.Group.ps1', 'AccessPackage.Coverage.ps1', 'AccessPackage.Discovery.ps1') {
			Copy-Item (Join-Path $PSScriptRoot "../$file") $packages
		}
		$app = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
		$secondApp = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
		$secondGroup = '44444444-4444-4444-4444-444444444444'
		$sharedGroup = '55555555-5555-5555-5555-555555555555'
		if ($Shared) { $Combined = $true; $Current = $true }
		$data = @{
			SchemaVersion = 1; ServicePrincipalId = $app
			Applications = @(@{ ServicePrincipalId = $app; ServicePrincipalName = 'First app'; GroupIds = @($script:groupId) })
			Groups = @(@{ Id = $script:groupId; DisplayName = 'SG_Finance_Reports'; ChildIds = @() })
		}
		if ($Shared) {
			$data.Groups[0].ChildIds = @($sharedGroup)
			$data.Groups += @{ Id = $sharedGroup; DisplayName = 'Shared Group'; ChildIds = @() }
		}
		$data | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $reviews 'snapshot.json')
		if ($Combined) {
			$secondData = @{
				SchemaVersion = 1; ServicePrincipalId = $secondApp
				Applications = @(@{ ServicePrincipalId = $secondApp; ServicePrincipalName = 'Second app'; GroupIds = @($secondGroup) })
				Groups = @(@{ Id = $secondGroup; DisplayName = 'SG_Operations'; ChildIds = @() })
			}
			if ($Shared) {
				$secondData.Groups[0].ChildIds = @($sharedGroup)
				$secondData.Groups += @{ Id = $sharedGroup; DisplayName = 'Shared Group'; ChildIds = @() }
			}
			$secondData | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $reviews 'snapshot-2.json')
		}
		@'
[CmdletBinding(SupportsShouldProcess)]
param([guid]$ServicePrincipalId, [string]$ReviewerType, [string]$DiscoveryOutputPath, [switch]$QuietDiscovery)
if (-not $WhatIfPreference) { throw 'Discovery requires WhatIf.' }
$snapshot = if ("$ServicePrincipalId" -eq 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb') { 'snapshot-2.json' } else { 'snapshot.json' }
Copy-Item -LiteralPath (Join-Path $PSScriptRoot $snapshot) -Destination $DiscoveryOutputPath -WhatIf:$false
'@ | Set-Content -LiteralPath (Join-Path $reviews 'Create-AccessReview.ps1')
		$warning = 'Assigning this package adds the user to every listed group and may grant access through other groups or applications. Nested membership does not establish effective application access.'
		if ($Current) {
			$warning = if ($Shared) {
				'Included groups: SG_Finance_Reports; Shared Group. Direct Entra application assignments for these groups: First app; Shared app.'
			}
			else { 'Included groups: SG_Finance_Reports. Direct Entra application assignments for these groups: First app.' }
			$script:graphFixture.AppAssignments[$script:groupId] = @(@{ resourceId = $app; resourceDisplayName = 'First app' })
		}
		$rows = @([pscustomobject][ordered]@{
			Role = 'Reader'; Brand = 'Brand A'; Region = 'EMEA'; 'Potential Teams' = 'Finance'
			GroupId = $(if ($Shared) { "$script:groupId;$sharedGroup" } else { $script:groupId })
			Description = $warning; 'Access Package Name' = 'First app - Finance'
			'Approver 1 List' = '<Employee Manager>'; 'Approver 2' = ''; 'Approver 3' = ''
			'Discovery ServicePrincipalId' = $app; 'Discovery RootGroupId' = $script:groupId
		})
		if ($Combined) {
			$rows += [pscustomobject][ordered]@{
				Role = 'Reader'; Brand = 'Brand B'; Region = 'EMEA'; 'Potential Teams' = 'Operations'
				GroupId = $(if ($Shared) { "$secondGroup;$sharedGroup" } else { $secondGroup })
				Description = $(if ($Shared) { 'Included groups: SG_Operations; Shared Group. Direct Entra application assignments for these groups: Second app; Shared app.' } else { $warning })
				'Access Package Name' = 'Second app - Operations'
				'Approver 1 List' = '<Employee Manager>'; 'Approver 2' = ''; 'Approver 3' = ''
				'Discovery ServicePrincipalId' = $secondApp; 'Discovery RootGroupId' = $secondGroup
			}
			$script:graphFixture.AdditionalGroups[$secondGroup] = @{
				id = $secondGroup; displayName = 'SG_Operations'; groupTypes = @()
				onPremisesSyncEnabled = $false; mailEnabled = $false; securityEnabled = $true; isAssignableToRole = $false
			}
			if ($Current) {
				$script:graphFixture.AppAssignments[$secondGroup] = @(@{ resourceId = $secondApp; resourceDisplayName = 'Second app' })
			}
		}
		if ($Shared) {
			$script:graphFixture.AppAssignments[$sharedGroup] = @(@{ resourceId = 'cccccccc-cccc-cccc-cccc-cccccccccccc'; resourceDisplayName = 'Shared app' })
			$script:graphFixture.AdditionalGroups[$sharedGroup] = @{
				id = $sharedGroup; displayName = 'Shared Group'; groupTypes = @()
				onPremisesSyncEnabled = $false; mailEnabled = $false; securityEnabled = $true; isAssignableToRole = $false
			}
		}
		$script:csvPath = Join-Path $TestDrive "discovery-packages-$([guid]::NewGuid()).xlsx"
		$rows | Export-Excel -Path $script:csvPath -WorksheetName 'Packages'
		$excel = [OfficeOpenXml.ExcelPackage]::new([IO.FileInfo]::new($script:csvPath))
		$metadata = $excel.Workbook.Worksheets.Add('Discovery metadata')
		$metadata.Cells[1, 1].Value = 'Discovery workbook version'
		$metadata.Cells[1, 2].Value = if ($Current) { '3' } elseif ($Combined) { '2' } else { '1' }
		foreach ($index in 0..($rows.Count - 1)) {
			$metadata.Cells[($index + 2), 1].Value = 'Service principal ID'
			$metadata.Cells[($index + 2), 2].Value = if ($index -eq 0) { $app } else { $secondApp }
		}
		Close-ExcelPackage -ExcelPackage $excel
		Join-Path $packages 'Create-AccessPackage.ps1'
	}

	function New-SingleGroupDiscoveryWorkbook {
		$path = New-DiscoveryWorkbookFixture -Current
		$reviews = Join-Path (Split-Path (Split-Path $path -Parent) -Parent) 'AccessReviews'
		$script:csvPath = Join-Path $TestDrive "single-group-$([guid]::NewGuid()).xlsx"
		& (Join-Path $PSScriptRoot '../Export-DiscoveryAccessPackages.ps1') `
			-DiscoveryPath (Join-Path $reviews 'snapshot.json') -WorkbookPath $script:csvPath -OnePackagePerGroup
		$excel = Open-ExcelPackage -Path $script:csvPath
		$sheet = $excel.Workbook.Worksheets['Packages']
		$sheet.Cells[2, 7].Value = 'Finance reporting - Read reports'
		$sheet.Cells[2, 8].Value = '<Employee Manager>'
		Close-ExcelPackage -ExcelPackage $excel
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
			GroupCounts = @{}
			AdditionalGroups = @{}
			GroupOwners = @{}
			OwnerNextPage = @{}
			DirectMembers = @{}
			AppAssignments = @{}
			PackageGroups = @($script:groupId)
			OmitPackageResourceRoles = $false
			PolicyDuration = 'P365D'
			Assignments = @()
			AssignmentRequests = @()
			DeniedUserId = $null
			NamedGroups = @()
			AddressUsers = @(
				@{ id = $script:ownerId; displayName = 'Application Owner'; userPrincipalName = 'owner@example.test'; mail = 'owner-mail@example.test' }
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
			CoverageMode = 'none'
			Writes = [System.Collections.Generic.List[object]]::new()
			Deleted = [System.Collections.Generic.List[string]]::new()
		}

		Mock Get-MgContext {
			@{
				AuthType = 'Delegated'
				Scopes = @('EntitlementManagement.ReadWrite.All', 'Group.Read.All', 'User.ReadBasic.All', 'Directory.Read.All')
			}
		}
		Mock Start-Sleep {}
		$fixture = $script:graphFixture
		Mock Get-MgGroup ({
			param($GroupId)
			if ($fixture.AdditionalGroups.ContainsKey($GroupId)) {
				return $fixture.AdditionalGroups[$GroupId]
			}
			$fixture.Group
		}.GetNewClosure())
		Mock Get-MgGroupMemberAsUser ({
			param($GroupId)
			$fixture.DirectMembers[$GroupId]
		}.GetNewClosure())
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
				if ($uriText -match '/groups/([^/]+)/appRoleAssignments$') {
					return @{ value = @($fixture.AppAssignments[$Matches[1]] | Where-Object { $_ }) }
				}
				if ($uriText -like '/v1.0/me?*') { return @{ id = '44444444-4444-4444-4444-444444444444' } }
				if ($uriText -like '*/roleManagement/directory/roleAssignments?*') {
					return @{ value = @(@{ directoryScopeId = '/'; roleDefinitionId = '45d8d3c5-c802-45c6-b32a-1d70b5e1e86e' }) }
				}
				if ($uriText -eq '/v1.0/identityGovernance/entitlementManagement/accessPackages') {
					if ($fixture.CoverageMode -eq 'forbidden') { throw 'Coverage denied' }
					$items = @()
					if ($fixture.CoverageMode -eq 'exact') { $items = @(@{ id = '99999999-9999-9999-9999-999999999999'; displayName = 'Existing' }) }
					return @{ value = $items }
				}
				if ($uriText -like '*/accessPackages/99999999-9999-9999-9999-999999999999?*') {
					return @{ resourceRoleScopes = @(@{
						role = @{ originSystem = 'AadGroup'; originId = "Member_$fixtureGroupId" }
						scope = @{ originSystem = 'AadGroup'; originId = $fixtureGroupId; isRootScope = $true }
					}) }
				}
				if ($decodedUri -like "*/assignmentPolicies?*accessPackage/id eq '99999999-9999-9999-9999-999999999999'*") {
					return @{ value = @(@{ allowedTargetScope = 'allMemberUsers'; expiration = @{ type = 'afterDuration'; duration = 'P365D' }; requestorSettings = @{ enableTargetsToSelfAddAccess = $true } }) }
				}
				if ($decodedUri -like "*/assignments?*accessPackage/id eq '99999999-9999-9999-9999-999999999999'*" -or
					$decodedUri -like "*/assignmentRequests?*accessPackage/id eq '99999999-9999-9999-9999-999999999999'*") { return @{ value = @() } }
				if ($uriText -match '/users/([^?]+)') {
					$id = $Matches[1]
					if ($id -eq $fixture.DeniedUserId) { throw 'User not found or forbidden.' }
					return @{ id = $id; displayName = 'Approver' }
				}
				if ($decodedUri -match "/users\?\`$filter=userPrincipalName eq '([^']+)' or mail eq '\1'&\`$select=id,userPrincipalName,mail,displayName$") {
					$address = $Matches[1]
					return @{ value = @($fixture.AddressUsers | Where-Object { $_.userPrincipalName -eq $address -or $_.mail -eq $address }) }
				}
				if ($uriText -match '/groups/([^/]+)/transitiveMembers/microsoft.graph.user/\$count$') {
					if ($fixture.GroupCounts.ContainsKey($Matches[1])) { return $fixture.GroupCounts[$Matches[1]] }
					return $fixture.UserCount
				}
				if ($uriText -match '/groups/([^/]+)/owners\?') {
					$ownerGroupId = $Matches[1]
					if ($uriText -match 'page=2') { return @{ value = @($fixture.OwnerNextPage[$ownerGroupId]) } }
					$page = @{ value = @($fixture.GroupOwners[$ownerGroupId]) }
					if ($fixture.OwnerNextPage.ContainsKey($ownerGroupId)) { $page['@odata.nextLink'] = "/v1.0/groups/$ownerGroupId/owners?page=2" }
					return $page
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
				if ($uriText -match '/accessPackages/([^?]+)\?') {
					$packageId = $Matches[1]
					$packageGroups = @($fixture.PackageGroups)
					$policyId = 'existing-policy-id'
					if ($packageId -eq 'new-package-id') {
						$packageGroups = @($fixture.Writes | Where-Object Uri -Like '*/resourceRoleScopes' | ForEach-Object { $_.Body.scope.originId })
						$policyId = 'new-policy-id'
					}
					$package = @{
						id = $packageId; displayName = 'Finance reporting - Read reports'; catalogId = 'catalog-id'
						assignmentPolicies = @(@{
							id = $policyId; displayName = 'Finance reporting - Read reports - Requests'
							allowedTargetScope = 'allMemberUsers'; expiration = @{ type = 'afterDuration'; duration = $fixture.PolicyDuration }
						})
						resourceRoleScopes = @(foreach ($groupId in $packageGroups) {
							@{ role = @{ originSystem = 'AadGroup'; originId = "Member_$groupId" }; scope = @{ originSystem = 'AadGroup'; originId = $groupId; isRootScope = $true } }
						})
					}
					if ($fixture.OmitPackageResourceRoles) { $package.Remove('resourceRoleScopes') }
					return $package
				}
				if ($decodedUri -match "/assignments\?\`$filter=accessPackage/id eq '(existing|new)-package-id'&") { return @{ value = $fixture.Assignments } }
				if ($decodedUri -match "/assignmentRequests\?\`$filter=accessPackage/id eq '(existing|new)-package-id'&") { return @{ value = $fixture.AssignmentRequests } }
				if ($uriText -match '/resourceRoles\?') {
					if ($decodedUri -notmatch "originSystem eq 'AadGroup' and resource/id eq 'resource-id(?:-([^']+))?'") {
						throw 'Resource-role lookup must filter by origin system and resource ID.'
					}
					$roleGroupId = if ($Matches[1]) { $Matches[1] } else { $fixtureGroupId }
					return @{
						value = @(
							@{ id = 'owner-role'; originId = "Owner_$roleGroupId"; originSystem = 'AadGroup' },
							@{ id = 'member-role'; originId = "Member_$roleGroupId"; originSystem = 'AadGroup' }
						)
					}
				}
				if ($uriText -match '/resources\?') {
					if ($fixture.ResourceDelay -gt 0) {
						$fixture.ResourceDelay--
						return @{ value = @() }
					}
					if (-not $fixture.ResourceExists -or $fixture.NeverReady) { return @{ value = @() } }
					$resourceGroupId = if ($decodedUri -match "originId eq '([^']+)'") { $Matches[1] } else { $fixtureGroupId }
					$resourceId = if ($resourceGroupId -eq $fixtureGroupId) { 'resource-id' } else { "resource-id-$resourceGroupId" }
					return @{
						value = @(
							@{
								id = $resourceId
								originId = $resourceGroupId
								originSystem = 'AadGroup'
								roles = @()
								scopes = @(@{ id = 'root-scope'; originId = $resourceGroupId; originSystem = 'AadGroup'; isRootScope = $true })
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
				if ($uriText.EndsWith('/assignmentRequests')) { return @{ id = 'new-request-id'; state = 'submitted' } }
			}
			elseif ($Method -eq 'DELETE') {
				$fixture.Deleted.Add($uriText)
				if ($fixture.FailRollback) { throw 'Simulated delete failure.' }
				return
			}
			throw "Unexpected Graph request: $Method $uriText"
		}.GetNewClosure())
	}

	Context 'Assignment backfill' {
		BeforeEach {
			$script:graphFixture.ExistingPackage = $true
			$script:graphFixture.DirectMembers[$script:groupId] = @(
				[pscustomobject]@{ Id = $script:ownerId; UserType = 'Member' }
				[pscustomobject]@{ Id = $script:fallbackId; UserType = 'Member' }
				[pscustomobject]@{ Id = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; UserType = 'Guest' }
			)
		}

		It 'does not scan group membership when the script variable is false' {
			$result = Invoke-Creation
			$result.Status | Should -Be 'Skipped'
			Should -Invoke Get-MgGroupMemberAsUser -Exactly -Times 0
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'backfills by default without recreating the package and skips live and pending users' {
			$pending = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
			$script:graphFixture.DirectMembers[$script:groupId] += [pscustomobject]@{ Id = $pending; UserType = 'Member' }
			$script:graphFixture.Assignments = @(@{ state = 'delivered'; target = @{ objectId = $script:ownerId } })
			$script:graphFixture.AssignmentRequests = @(@{ requestType = 'adminAdd'; state = 'submitted'; assignment = @{ target = @{ objectId = $pending } } })
			$result = Invoke-Creation -Backfill
			$result.Status | Should -Be 'Skipped'
			$result.AssignmentsSubmitted | Should -Be 1
			$result.AssignmentsAlreadyCovered | Should -Be 2
			$script:graphFixture.Writes | Should -HaveCount 1
			$request = Get-PostedBody '/assignmentRequests'
			$request.requestType | Should -Be 'adminAdd'
			$request.assignment.accessPackageId | Should -Be 'existing-package-id'
			$request.assignment.assignmentPolicyId | Should -Be 'existing-policy-id'
			$request.assignment.targetId | Should -Be $script:fallbackId
		}

		It 'previews eligible assignments on an existing package without writes' {
			$result = Invoke-Creation -Backfill -Preview
			$result.AssignmentBackfill | Should -Be 'Preview'
			$result.AssignmentsWouldSubmit | Should -Be 2
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'fails closed when the existing Requests policy does not last 365 days' {
			$script:graphFixture.PolicyDuration = 'P30D'
			$result = Invoke-Creation -Backfill
			$result.AssignmentBackfill | Should -Be 'Failed'
			$result.AssignmentBackfillError | Should -Match '365-day Requests policy'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'does not backfill a package that grants a group absent from the input' {
			$script:graphFixture.PackageGroups += '77777777-7777-7777-7777-777777777777'
			$result = Invoke-Creation -Backfill
			$result.AssignmentBackfill | Should -Be 'Failed'
			$result.AssignmentBackfillError | Should -Match 'role outside'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'reports missing expanded resource roles rather than a group mismatch' {
			$script:graphFixture.OmitPackageResourceRoles = $true
			$result = Invoke-Creation -Backfill
			$result.AssignmentBackfill | Should -Be 'Failed'
			$result.AssignmentBackfillError | Should -Match 'Graph did not return the package resource roles'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'does not roll back a new package if assignment submission fails' {
			$script:graphFixture.ExistingPackage = $false
			$script:graphFixture.FailOn = '/assignmentRequests'
			$result = Invoke-Creation -Backfill
			$result.Status | Should -Be 'Created'
			$result.AssignmentsFailed | Should -Be 2
			$script:graphFixture.Deleted | Should -HaveCount 0
		}
	}

	It 'reports row progress and ETA and clears it after processing' {
		@"
Name,Description,GroupId,AppOwnerIds
Finance reporting - Read reports,View financial reports.,$script:groupId,$script:ownerId
Finance reporting - Edit reports,Edit financial reports.,$script:groupId,$script:ownerId
"@ | Set-Content -LiteralPath $script:csvPath
		$progress = [System.Collections.Generic.List[object]]::new()
		Mock Write-Progress {
			$progress.Add(@{ Status = $Status; Percent = $PercentComplete; Remaining = $SecondsRemaining; Completed = $Completed.IsPresent })
		}
		$results = @(Invoke-Creation -Preview)
		$results | Should -HaveCount 2
		@($progress | Where-Object { $_.Status -match '0 of 2 rows completed.*Processing row 1' }).Count | Should -Be 1
		$half = @($progress | Where-Object { $_.Status -match '1 of 2 rows completed.*Finished' })[0]
		$half.Percent | Should -Be 50
		$half.Remaining | Should -BeGreaterOrEqual 0
		$half.Status | Should -Match 'Last update: \d\d:\d\d:\d\d'
		$progress[-1].Completed | Should -BeTrue
		$script:graphFixture.Writes | Should -HaveCount 0
	}

	It 'clears progress if discovery refresh fails' {
		$path = New-DiscoveryWorkbookFixture
		$reviews = Join-Path (Split-Path (Split-Path $path -Parent) -Parent) 'AccessReviews'
		Remove-Item -LiteralPath (Join-Path $reviews 'snapshot.json')
		Mock Write-Progress {}
		{ & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName -WhatIf -WarningAction SilentlyContinue } |
			Should -Throw '*Unable to refresh application*'
		Should -Invoke Write-Progress -ParameterFilter { $Completed } -Exactly -Times 1
	}

	Context 'Excel workbook' -Skip:(-not (Get-Module -ListAvailable -Name ImportExcel)) {
		It 'rechecks discovery and previews a safe discovery workbook without Graph writes' {
			$path = New-DiscoveryWorkbookFixture
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -WhatIf -WarningAction SilentlyContinue
			$result.Status | Should -Be 'Would be created'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'rechecks each application in a combined discovery workbook' {
			$path = New-DiscoveryWorkbookFixture -Combined
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -WhatIf -WarningAction SilentlyContinue
			$result | Should -HaveCount 2
			$result.Status | Should -Be @('Would be created', 'Would be created')
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'previews separate packages containing the same shared group' {
			$path = New-DiscoveryWorkbookFixture -Shared
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -WhatIf -WarningAction SilentlyContinue
			$result | Should -HaveCount 2
			$result.Status | Should -Be @('Would be created', 'Would be created')
			foreach ($item in $result) { $item.GroupIds | Should -Contain '55555555-5555-5555-5555-555555555555' }
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'previews a workbook produced by the discovery exporter' {
			$path = New-DiscoveryWorkbookFixture -Current
			$reviews = Join-Path (Split-Path (Split-Path $path -Parent) -Parent) 'AccessReviews'
			$script:csvPath = Join-Path $TestDrive "exported-$([guid]::NewGuid()).xlsx"
			& (Join-Path $PSScriptRoot '../Export-DiscoveryAccessPackages.ps1') `
				-DiscoveryPath (Join-Path $reviews 'snapshot.json') -WorkbookPath $script:csvPath
			$excel = Open-ExcelPackage -Path $script:csvPath
			$sheet = $excel.Workbook.Worksheets['Packages']
			$sheet.Cells[2, 7].Value = 'Finance reports'
			$sheet.Cells[2, 8].Value = '<Employee Manager>'
			$sheet.Cells[2, 6].Value = 'Rewritten business description for finance reporting.'
			Close-ExcelPackage -ExcelPackage $excel
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -WhatIf -WarningAction SilentlyContinue
			$result.Status | Should -Be 'Would be created'
			$result.GroupIds | Should -Be @($script:groupId)
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'creates a single-group export and backfills direct members without application assignment reads' {
			Mock Get-MgContext { @{ AuthType = 'AppOnly'; Scopes = @('EntitlementManagement.ReadWrite.All', 'Group.ReadWrite.All', 'User.Read.All') } }
			$path = New-SingleGroupDiscoveryWorkbook
			$script:graphFixture.RequestFailurePattern = '/appRoleAssignments$'
			$script:graphFixture.DirectMembers[$script:groupId] = @(
				[pscustomobject]@{ Id = $script:ownerId; UserType = 'Member' },
				[pscustomobject]@{ Id = $script:fallbackId; UserType = 'Guest' }
			)
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -Confirm:$false -WarningAction SilentlyContinue
			$result.Status | Should -Be 'Created'
			$result.AssignmentBackfill | Should -Be 'Processed'
			$result.AssignmentsSubmitted | Should -Be 1
			(Get-PostedBody '/assignmentRequests').assignment.targetId | Should -Be $script:ownerId
			Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly -ParameterFilter { $Uri -like '*/appRoleAssignments' }
		}

		It 'validates a nested single-group package against refreshed application discovery' {
			Mock Get-MgContext { @{ AuthType = 'AppOnly'; Scopes = @('EntitlementManagement.ReadWrite.All', 'Group.ReadWrite.All', 'User.Read.All') } }
			$path = New-DiscoveryWorkbookFixture -Shared
			$reviews = Join-Path (Split-Path (Split-Path $path -Parent) -Parent) 'AccessReviews'
			$script:csvPath = Join-Path $TestDrive "single-nested-$([guid]::NewGuid()).xlsx"
			& (Join-Path $PSScriptRoot '../Export-DiscoveryAccessPackages.ps1') `
				-DiscoveryPath (Join-Path $reviews 'snapshot.json') -WorkbookPath $script:csvPath -OnePackagePerGroup
			$excel = Open-ExcelPackage -Path $script:csvPath
			$sheet = $excel.Workbook.Worksheets['Packages']
			for ($i = 2; $i -le 3; $i++) {
				$sheet.Cells[$i, 7].Value = "Group package $i"
				$sheet.Cells[$i, 8].Value = '<Employee Manager>'
			}
			Close-ExcelPackage -ExcelPackage $excel
			$preview = @(& $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -WhatIf -WarningAction SilentlyContinue)
			$preview.Status | Should -Be @('Would be created', 'Would be created')
			$data = Get-Content -LiteralPath (Join-Path $reviews 'snapshot.json') -Raw | ConvertFrom-Json
			$data.Groups[0].ChildIds = @()
			$data | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $reviews 'snapshot.json')
			$preview = @(& $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -WhatIf -WarningAction SilentlyContinue)
			$nested = @($preview | Where-Object GroupId -eq '55555555-5555-5555-5555-555555555555')[0]
			$nested.Status | Should -Be 'Manual review'
			$nested.Reason | Should -Match 'no longer in the application discovery'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'allows a renamed group when its original single-group resource ID is unchanged' {
			$path = New-SingleGroupDiscoveryWorkbook
			$reviews = Join-Path (Split-Path (Split-Path $path -Parent) -Parent) 'AccessReviews'
			$data = Get-Content -LiteralPath (Join-Path $reviews 'snapshot.json') -Raw | ConvertFrom-Json
			$data.Groups[0].DisplayName = 'Renamed group'
			$data | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $reviews 'snapshot.json')
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -WhatIf -WarningAction SilentlyContinue
			$result.Status | Should -Be 'Would be created'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'blocks a single-group workbook if an extra group is added' {
			$path = New-SingleGroupDiscoveryWorkbook
			$extra = '55555555-5555-5555-5555-555555555555'
			$script:graphFixture.AdditionalGroups[$extra] = @{
				id = $extra; displayName = 'Other Group'; groupTypes = @()
				onPremisesSyncEnabled = $false; mailEnabled = $false; securityEnabled = $true; isAssignableToRole = $false
			}
			$excel = Open-ExcelPackage -Path $script:csvPath
			$excel.Workbook.Worksheets['Packages'].Cells[2, 5].Value = "$script:groupId;$extra"
			Close-ExcelPackage -ExcelPackage $excel
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -Confirm:$false -WarningAction SilentlyContinue
			$result.Status | Should -Be 'Manual review'
			$result.Reason | Should -Match 'exactly its original discovered group'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'creates separate packages for roots sharing a group' {
			$path = New-DiscoveryWorkbookFixture -Shared
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -Confirm:$false -WarningAction SilentlyContinue
			$result.Status | Should -Be @('Created', 'Created')
			$created = @($script:graphFixture.Writes | Where-Object Uri -like '*/accessPackages')
			$created | Should -HaveCount 2
			$created[0].Body.description | Should -Match 'First app; Shared app'
			$created[1].Body.description | Should -Match 'Second app; Shared app'
		}

		It 'does not use application descriptions as a resource safety check' {
			$path = New-DiscoveryWorkbookFixture -Current
			$script:graphFixture.AppAssignments[$script:groupId] += @{
				resourceId = 'cccccccc-cccc-cccc-cccc-cccccccccccc'; resourceDisplayName = 'New app'
			}
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -WhatIf -WarningAction SilentlyContinue
			$result.Status | Should -Be 'Would be created'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'does not request group application assignments when validating editable descriptions' {
			$path = New-DiscoveryWorkbookFixture -Current
			$script:graphFixture.RequestFailurePattern = '/appRoleAssignments$'
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -WhatIf -WarningAction SilentlyContinue
			$result.Status | Should -Be 'Would be created'
			Should -Invoke Invoke-MgGraphRequest -Times 0 -ParameterFilter { "$Uri" -match '/appRoleAssignments$' }
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'creates a safe discovery workbook candidate after the final coverage check' {
			$path = New-DiscoveryWorkbookFixture
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -Confirm:$false -WarningAction SilentlyContinue
			$result.Status | Should -Be 'Created'
			@($script:graphFixture.Writes | Where-Object Uri -like '*/accessPackages') | Should -HaveCount 1
		}

		It 'does not create or backfill an exact active package or an unreadable inventory' -TestCases @(
			@{ Coverage = 'exact'; Expected = 'Skipped' }
			@{ Coverage = 'forbidden'; Expected = 'Manual review' }
		) {
			param($Coverage, $Expected)
			$path = New-DiscoveryWorkbookFixture
			$script:graphFixture.CoverageMode = $Coverage
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -Confirm:$false -WarningAction SilentlyContinue
			$result.Status | Should -Be $Expected
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'blocks edited groups when refreshed discovery has changed' {
			$path = New-DiscoveryWorkbookFixture
			$reviews = Join-Path (Split-Path (Split-Path $path -Parent) -Parent) 'AccessReviews'
			$data = Get-Content -LiteralPath (Join-Path $reviews 'snapshot.json') -Raw | ConvertFrom-Json
			$data.Groups[0].ChildIds = @('55555555-5555-5555-5555-555555555555')
			$data.Groups += [pscustomobject]@{ Id = '55555555-5555-5555-5555-555555555555'; ChildIds = @() }
			$data | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $reviews 'snapshot.json')
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -Confirm:$false -WarningAction SilentlyContinue
			$result.Status | Should -Be 'Manual review'
			$result.Reason | Should -Match 'do not match'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'does not create a discovery package when its risk warning is removed' {
			$path = New-DiscoveryWorkbookFixture
			$excel = [OfficeOpenXml.ExcelPackage]::new([IO.FileInfo]::new($script:csvPath))
			$sheet = $excel.Workbook.Worksheets['Packages']
			$column = 1..$sheet.Dimension.End.Column | Where-Object { $sheet.Cells[1, $_].Value -eq 'Description' }
			$sheet.Cells[2, $column].Value = 'Replaced description'
			Close-ExcelPackage -ExcelPackage $excel
			$result = & $path -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-FallbackApproverIds $script:fallbackId -Confirm:$false -WarningAction SilentlyContinue
			$result.Status | Should -Be 'Failed'
			$result.Reason | Should -Match 'retain the generated warning'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		BeforeAll { Import-Module ImportExcel -WarningAction SilentlyContinue }

		It 'creates a manager stage without fallback users and warns about missing managers' -TestCases @(
			@{ Workbook = $false },
			@{ Workbook = $true }
		) {
			param($Workbook)
			if ($Workbook) { $script:csvPath = New-WorkbookFixture @{ 'Approver 1 List' = '<Employee Manager>' } }
			$warnings = @()
			$result = & $script:noBackfillScriptPath -InputPath $script:csvPath -CatalogName $script:graphFixture.CatalogName `
				-Confirm:$false -WarningAction SilentlyContinue -WarningVariable warnings
			$result.Status | Should -Be 'Created'
			$stages = (Get-PostedBody '/assignmentPolicies').requestApprovalSettings.stages
			$stages[0].primaryApprovers[0].'@odata.type' | Should -Be '#microsoft.graph.targetManager'
			$stages[0].fallbackPrimaryApprovers | Should -HaveCount 0
			($warnings -join ' ') | Should -Match 'without fallback approvers.*Users without a manager'
			Should -Invoke Invoke-MgGraphRequest -Times 0 -ParameterFilter { "$Uri" -like "/v1.0/users/$script:fallbackId`?*" }
		}

		It 'identifies a group in the workbook that was removed from the existing package' {
			$removed = '77777777-7777-7777-7777-777777777777'
			$script:csvPath = New-WorkbookFixture @{ GroupId = "$script:groupId;$removed" }
			$script:graphFixture.ExistingPackage = $true
			$result = Invoke-Creation -Backfill
			$result.AssignmentBackfill | Should -Be 'Failed'
			$result.Name | Should -Be 'Finance reporting - Read reports'
			$result.AssignmentBackfillError | Should -Match $removed
			$result.AssignmentBackfillError | Should -Match 'Update the input'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'backfills a user in both groups once and does not grant extra group access' {
			$otherGroup = '77777777-7777-7777-7777-777777777777'
			$script:csvPath = New-WorkbookFixture @{ GroupId = "$script:groupId,$otherGroup" }
			$script:graphFixture.ExistingPackage = $true
			$script:graphFixture.PackageGroups = @($script:groupId, $otherGroup)
			$sharedUser = [pscustomobject]@{ Id = $script:ownerId; UserType = 'Member' }
			$script:graphFixture.DirectMembers[$script:groupId] = @($sharedUser, [pscustomobject]@{ Id = $script:fallbackId; UserType = 'Member' })
			$script:graphFixture.DirectMembers[$otherGroup] = @($sharedUser, $sharedUser)
			$result = Invoke-Creation -Backfill
			$result.AssignmentsSubmitted | Should -Be 1
			$result.AssignmentPartialMembership | Should -Be 1
			$script:graphFixture.Writes | Should -HaveCount 1
			(Get-PostedBody '/assignmentRequests').assignment.targetId | Should -Be $script:ownerId
		}

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

		It 'submits exactly two stages when Approver 3 is NA' {
			$script:csvPath = New-WorkbookFixture @{ GroupId = $script:groupId; 'Approver 2' = 'owner-mail@example.test'; 'Approver 3' = ' na ' }
			$result = Invoke-Creation
			$result.Status | Should -Be 'Created'
			$result.ApprovalStages | Should -Be 2
			$stages = (Get-PostedBody '/assignmentPolicies').requestApprovalSettings.stages
			$stages | Should -HaveCount 2
			$stages[0].primaryApprovers[0].'@odata.type' | Should -Be '#microsoft.graph.targetManager'
			$stages[1].primaryApprovers[0].userId | Should -Be $script:ownerId
		}

		It 'prints the two planned approvers and manager fallback in WhatIf without writing' {
			$description = "Read financial reports.`nMembership in the reporting root group only."
			$script:csvPath = New-WorkbookFixture @{
				GroupId = $script:groupId; 'Approver 2' = 'owner-mail@example.test'; 'Approver 3' = 'NA'
				Description = $description; Role = 'Reader'; Region = 'EMEA'; Brand = 'Finance'; 'Potential Teams' = 'Reporting'
			}
			$output = @(Invoke-Creation -Preview 6>&1)
			$result = @($output | Where-Object { $_ -isnot [System.Management.Automation.InformationRecord] })
			$console = @($output | Where-Object { $_ -is [System.Management.Automation.InformationRecord] }) -join "`n"
			$result | Should -HaveCount 1
			$result[0].Status | Should -Be 'Would be created'
			$expectedDescription = @($description, 'Role: Reader', 'Region: EMEA', 'Brand: Finance', 'Potential Teams: Reporting') -join [Environment]::NewLine
			$result[0].Description | Should -Be $expectedDescription
			$console | Should -Match ([regex]::Escape("Description:`n$expectedDescription"))
			$expected = @(
				"Stage 1: Recipient's direct manager (if no manager: Approver <$script:fallbackId>)"
				'Stage 2: Application Owner <owner@example.test>'
			) -join [Environment]::NewLine
			$result[0].ApprovalStageDetails | Should -Be $expected
			$console | Should -Match ([regex]::Escape($expected))
			$console | Should -Match 'new and extension requests for directory members \(excluding guests\)'
			$console | Should -Not -Match 'Stage 3:'
			$script:graphFixture.Writes | Should -HaveCount 0
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

		It 'identifies the approver column when Graph returns a matching user without an id' {
			$script:graphFixture.AddressUsers = @(@{ mail = 'owner-mail@example.test' })
			$script:csvPath = New-WorkbookFixture @{ 'Approver 2' = 'owner-mail@example.test' }
			$result = Invoke-Creation -Preview
			$result.Status | Should -Be 'Failed'
			$result.Reason | Should -Match 'Approver 2 .*Graph did not return a valid user object ID'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'provisions one package with every distinct listed group and uses the largest group for approval' {
			$second = '77777777-7777-7777-7777-777777777777'
			$script:graphFixture.CatalogExists = $false
			$script:graphFixture.ResourceExists = $false
			$script:graphFixture.AdditionalGroups[$second] = $script:graphFixture.Group.Clone()
			$script:graphFixture.AdditionalGroups[$second].id = $second
			$script:graphFixture.AdditionalGroups[$second].displayName = 'Second group from Entra'
			$script:graphFixture.GroupCounts[$second] = '501'
			$script:csvPath = New-WorkbookFixture @{
				GroupId = "$script:groupId,$second;`n$script:groupId"
				'Access Group' = "Finance's Readers"
				'Approver 2' = 'owner-mail@example.test'
			}
			$result = Invoke-Creation
			$result.Status | Should -Be 'Created'
			$result.GroupIds | Should -Be @($script:groupId, $second)
			$result.GroupNames | Should -Be @('SG_Finance_Reports', 'Second group from Entra')
			Should -Invoke Get-MgGroup -Times 1 -Exactly -ParameterFilter { $GroupId -eq $script:groupId }
			Should -Invoke Get-MgGroup -Times 1 -Exactly -ParameterFilter { $GroupId -eq '77777777-7777-7777-7777-777777777777' }
			Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly -ParameterFilter { $Uri -like '/v1.0/groups?*' -and $Uri -like '*displayName%20eq*' }
			$result.UserCount | Should -Be 501
			$result.ApprovalStages | Should -Be 2
			$writes = $script:graphFixture.Writes
			@($writes | Where-Object Uri -Like '*/accessPackages') | Should -HaveCount 1
			$imports = @($writes | Where-Object Uri -Like '*/resourceRequests')
			$imports.Body.resource.originId | Should -Be @($script:groupId, $second)
			$bindings = @($writes | Where-Object Uri -Like '*/new-package-id/resourceRoleScopes')
			$bindings.Body.role.originId | Should -Be @("Member_$script:groupId", "Member_$second")
		}

		It 'does not fall back to Access Group when GroupId is <Case>' -TestCases @(
			@{ Case = 'blank'; Value = ''; Reason = 'GroupId must not be blank' }
			@{ Case = 'a group name'; Value = "Finance's Readers"; Reason = 'must contain only valid Entra group object IDs' }
		) {
			param($Case, $Value, $Reason)
			$script:csvPath = New-WorkbookFixture @{ GroupId = $Value }
			$result = Invoke-Creation -Preview
			$result.Status | Should -Be 'Failed'
			$result.Reason | Should -Match $Reason
			Should -Invoke Get-MgGroup -Times 0 -Exactly
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'fails the entire row when the second listed group cannot be managed' {
			$second = '77777777-7777-7777-7777-777777777777'
			$script:graphFixture.AdditionalGroups[$second] = $script:graphFixture.Group.Clone()
			$script:graphFixture.AdditionalGroups[$second].id = $second
			$script:graphFixture.AdditionalGroups[$second].displayName = 'Synced second group'
			$script:graphFixture.AdditionalGroups[$second].onPremisesSyncEnabled = $true
			$script:csvPath = New-WorkbookFixture @{ 'Access Group' = "$script:groupId,$second" }
			$result = Invoke-Creation
			$result.Status | Should -Be 'Failed'
			$result.Reason | Should -Match 'Synced second group.*synchronized'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'previews a manager-only workbook row without Graph writes' {
			$script:csvPath = New-WorkbookFixture @{}
			$result = Invoke-Creation -Preview
			$result.Status | Should -Be 'Would be created'
			$result.ApprovalStages | Should -Be 1
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'supports named users first and manager in a later workbook stage' {
			$script:csvPath = New-WorkbookFixture @{
				'Approver 1 List' = 'owner@example.test; fallback@example.test'
				'Approver 2' = '<Line Manager>'; 'Approver 3' = 'NA'
			}
			$result = Invoke-Creation
			$result.Status | Should -Be 'Created'
			$stages = (Get-PostedBody '/assignmentPolicies').requestApprovalSettings.stages
			$stages | Should -HaveCount 2
			$stages[0].primaryApprovers.userId | Should -Contain $script:ownerId
			$stages[0].primaryApprovers.userId | Should -Contain $script:fallbackId
			$stages[0].fallbackPrimaryApprovers | Should -HaveCount 0
			$stages[1].primaryApprovers[0].'@odata.type' | Should -Be '#microsoft.graph.targetManager'
			$stages[1].fallbackPrimaryApprovers[0].userId | Should -Be $script:fallbackId
		}

		It 'supports a resource owner and named approvers in later stages with paginated owners' {
			$script:graphFixture.GroupOwners[$script:groupId] = @(@{ id = $script:ownerId; '@odata.type' = '#microsoft.graph.user' })
			$script:graphFixture.OwnerNextPage[$script:groupId] = @(@{ id = $script:fallbackId; '@odata.type' = '#microsoft.graph.user' })
			$script:csvPath = New-WorkbookFixture @{
				'Approver 1 List' = '<Employee Manager>'; 'Approver 2' = '<Resource Owner>'
				'Approver 3' = 'owner@example.test'
			}
			$result = Invoke-Creation
			$result.Status | Should -Be 'Created'
			$stages = (Get-PostedBody '/assignmentPolicies').requestApprovalSettings.stages
			$stages[1].primaryApprovers | Should -HaveCount 2
			$stages[1].primaryApprovers.userId | Should -Contain $script:fallbackId
			$stages[2].primaryApprovers[0].userId | Should -Be $script:ownerId
		}

		It 'does not require a manager-bypass manual step for named-user-only stages' {
			$script:csvPath = New-WorkbookFixture @{ 'Approver 1 List' = 'owner@example.test' }
			$result = Invoke-Creation -Preview
			$result.Status | Should -Be 'Would be created'
			$result.ManualAction | Should -BeNullOrEmpty
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'fails before writes when group owner lookup is forbidden' {
			$script:graphFixture.RequestFailurePattern = '/groups/.*/owners'
			$script:csvPath = New-WorkbookFixture @{ 'Approver 1 List' = '<Resource Owner>' }
			$result = Invoke-Creation
			$result.Status | Should -Be 'Failed'
			$result.Reason | Should -Match 'GET /v1.0/groups/.*/owners'
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'captures and deduplicates all human group owners as named approvers in any workbook stage' {
			$second = '77777777-7777-7777-7777-777777777777'
			$script:graphFixture.AdditionalGroups[$second] = $script:graphFixture.Group.Clone()
			$script:graphFixture.AdditionalGroups[$second].id = $second
			$script:graphFixture.GroupOwners[$script:groupId] = @(
				@{ id = $script:ownerId; '@odata.type' = '#microsoft.graph.user' },
				@{ id = '88888888-8888-8888-8888-888888888888'; '@odata.type' = '#microsoft.graph.servicePrincipal' }
			)
			$script:graphFixture.GroupOwners[$second] = @(
				@{ id = $script:ownerId; '@odata.type' = '#microsoft.graph.user' },
				@{ id = $script:fallbackId; '@odata.type' = '#microsoft.graph.user' }
			)
			$script:csvPath = New-WorkbookFixture @{
				GroupId = "$script:groupId,$second"; 'Approver 1 List' = '<Resource Owner>'
				'Approver 2' = 'owner@example.test'; 'Approver 3' = '<Line Manager>'
			}
			$result = Invoke-Creation
			$result.Status | Should -Be 'Created'
			$stages = (Get-PostedBody '/assignmentPolicies').requestApprovalSettings.stages
			$stages[0].primaryApprovers | Should -HaveCount 2
			$stages[0].primaryApprovers.userId | Should -Contain $script:ownerId
			$stages[0].primaryApprovers.userId | Should -Contain $script:fallbackId
			$stages[2].primaryApprovers[0].'@odata.type' | Should -Be '#microsoft.graph.targetManager'
			$result.ApprovalStageDetails | Should -Match 'current group owners captured as named users'
		}

		It 'fails closed before writes when a resource-owner group has <Case>' -TestCases @(
			@{ Case = 'no human owners'; Owner = @{ id = '88888888-8888-8888-8888-888888888888'; '@odata.type' = '#microsoft.graph.servicePrincipal' }; Reason = 'no human owners' }
			@{ Case = 'an unclassified owner'; Owner = @{ id = '88888888-8888-8888-8888-888888888888' }; Reason = 'owner type unavailable' }
			@{ Case = 'an invalid human owner'; Owner = @{ id = 'bad'; '@odata.type' = '#microsoft.graph.user' }; Reason = 'without a valid user ID' }
		) {
			param($Case, $Owner, $Reason)
			$script:graphFixture.GroupOwners[$script:groupId] = @($Owner)
			$script:csvPath = New-WorkbookFixture @{ 'Approver 1 List' = '<Group Owner>' }
			$result = Invoke-Creation
			$result.Status | Should -Be 'Failed'
			$result.Reason | Should -Match $Reason
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'rejects self-review and repeated manager stages without provisioning' -TestCases @(
			@{ First = '<Self Review>'; Second = ''; Reason = 'self-review is not supported' }
			@{ First = '<Employee Manager>'; Second = '<Line Manager>'; Reason = 'only be selected in one approval stage' }
		) {
			param($First, $Second, $Reason)
			$script:csvPath = New-WorkbookFixture @{ 'Approver 1 List' = $First; 'Approver 2' = $Second }
			$result = Invoke-Creation
			$result.Status | Should -Be 'Failed'
			$result.Reason | Should -Match $Reason
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'uses only workbook stages above 500 users (default owner: <HasDefault>)' -TestCases @(
			@{ HasDefault = $false }, @{ HasDefault = $true }
		) {
			param($HasDefault)
			$script:graphFixture.UserCount = '501'
			$script:csvPath = New-WorkbookFixture @{}
			if ($HasDefault) {
				$result = Invoke-Creation -Preview -DefaultOwners $script:ownerId
				$result.Status | Should -Be 'Would be created'
				$result.ApprovalStages | Should -Be 1
			}
			else {
				$result = Invoke-Creation -Preview
				$result.Status | Should -Be 'Would be created'
				$result.ApprovalStages | Should -Be 1
			}
			$script:graphFixture.Writes | Should -HaveCount 0
		}

		It 'rejects unsafe workbook input before provisioning (<Case>)' -TestCases @(
			@{ Case = 'unsupported manager token'; Overrides = @{ 'Approver 1 List' = '<Employee Manager> plus owner' }; Reason = 'use Entra user object IDs' }
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
		(Get-Content -LiteralPath $script:csvPath) -replace "$script:ownerId$", '' | Set-Content -LiteralPath $script:csvPath
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

	It 'uses CSV application-owner approvers at any group size' -TestCases @(
		@{ Count = '10' }, @{ Count = '500' }, @{ Count = '501' }
	) {
		param($Count)
		$script:graphFixture.UserCount = $Count
		$result = Invoke-Creation
		$result.Status | Should -Be 'Created'
		$result.ApprovalStages | Should -Be 2
		(Get-PostedBody '/assignmentPolicies').requestApprovalSettings.stages[1].primaryApprovers[0].userId | Should -Be $script:ownerId
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

	It 'uses only the manager stage for a large group without application owners' {
		$script:graphFixture.UserCount = '501'
		(Get-Content -LiteralPath $script:csvPath) -replace "$script:ownerId$", '' | Set-Content -LiteralPath $script:csvPath
		$result = Invoke-Creation
		$result.Status | Should -Be 'Created'
		$result.ApprovalStages | Should -Be 1
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
		$result.ApprovalStages | Should -Be 2
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

	It 'stops when a supplied fallback user cannot be read' {
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
