BeforeAll {
	Import-Module Microsoft.Graph.Authentication
	Import-Module Microsoft.Graph.Groups
	Import-Module ImportExcel -WarningAction SilentlyContinue
	$script:exportPath = Join-Path $PSScriptRoot '../Export-DiscoveryAccessPackages.ps1'
	$script:app = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
	$script:root = '11111111-1111-1111-1111-111111111111'
	$script:child = '22222222-2222-2222-2222-222222222222'
	$script:other = '33333333-3333-3333-3333-333333333333'
}

Describe 'Discovery workbook export' {
	BeforeEach {
		$script:input = Join-Path $TestDrive "discovery-$([guid]::NewGuid()).json"
		$script:output = Join-Path $TestDrive "draft-$([guid]::NewGuid()).xlsx"
		$script:fixture = @{
			SchemaVersion = 1; ServicePrincipalId = $script:app
			Applications = @(@{ ServicePrincipalId = $script:app; ServicePrincipalName = '=unsafe'; GroupIds = @($script:root) })
			Groups = @(
				@{ Id = $script:root; DisplayName = 'Root SG'; ChildIds = @($script:child) }
				@{ Id = $script:child; DisplayName = 'Child'; ChildIds = @($script:root) }
				@{ Id = $script:other; DisplayName = 'Other'; ChildIds = @($script:child) }
			)
		}
		Mock Get-MgContext { @{ AuthType = 'Delegated'; Scopes = @('Directory.Read.All') } }
		Mock Get-MgGroup {
			param($GroupId)
			@{ id = $GroupId; displayName = 'Group'; groupTypes = @(); securityEnabled = $true }
		}
		Mock Invoke-MgGraphRequest {
			param($Uri)
			if ($Uri -like '/v1.0/me?*') { return @{ id = '44444444-4444-4444-4444-444444444444' } }
			if ($Uri -like '*/roleAssignments?*') { return @{ value = @(@{ directoryScopeId = '/'; roleDefinitionId = '45d8d3c5-c802-45c6-b32a-1d70b5e1e86e' }) } }
			if ($Uri -like '*/accessPackages') { return @{ value = @() } }
			if ($Uri -like '*/11111111-1111-1111-1111-111111111111/appRoleAssignments') { return @{ value = @(@{ resourceId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; resourceDisplayName = 'Application A' }) } }
			if ($Uri -like '*/22222222-2222-2222-2222-222222222222/appRoleAssignments') { return @{ value = @(@{ resourceId = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'; resourceDisplayName = 'Application B' }) } }
			if ($Uri -like '*/33333333-3333-3333-3333-333333333333/appRoleAssignments') { return @{ value = @() } }
			throw "Unexpected request $Uri"
		}
	}

	It 'writes a usable draft with literal text and the full cyclic descendant set' {
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output
		@(Import-Excel -Path $script:output -WorksheetName 'Manual review').Reason | Should -BeNullOrEmpty
		$rows = @(Import-Excel -Path $script:output -WorksheetName Packages)
		$rows | Should -HaveCount 1
		$rows[0].GroupId.Split(';') | Should -Contain $script:root
		$rows[0].GroupId.Split(';') | Should -Contain $script:child
		$rows[0].'Approver 1 List' | Should -BeNullOrEmpty
		$rows[0].Role | Should -BeNullOrEmpty
		$rows[0].Description | Should -Match 'Direct Entra application assignments for these groups: Application A; Application B'
		$rows[0].Description | Should -Match 'Included groups: Root SG; Child'
		$rows[0].Description | Should -Not -Match 'may grant'
		$topology = @(Import-Excel -Path $script:output -WorksheetName Topology)
		$topology | Should -HaveCount 2
		@($topology | Where-Object 'Group ID' -eq $script:child)[0].'Parent group IDs' | Should -Be $script:root
		$excel = Open-ExcelPackage -Path $script:output
		try {
			$excel.Workbook.Worksheets['Packages'].Cells[2, 7].Formula | Should -BeNullOrEmpty
			$excel.Workbook.Worksheets['Packages'].Cells[2, 7].Value | Should -Match '^=unsafe'
		}
		finally { Close-ExcelPackage -ExcelPackage $excel -NoSave }
		{ & $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output } | Should -Throw '*already exists*'
	}

	It 'suggests privileged roles from root and descendant group names' {
		$script:fixture.Groups[0].DisplayName = 'Application-Administrators'
		$script:fixture.Groups[1].DisplayName = 'ApplicationSuperUsersGroup'
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output
		$row = @(Import-Excel -Path $script:output -WorksheetName Packages)[0]
		$row.Role | Should -Match 'Potential privileged access'
		$row.Role | Should -Match 'Administrators'
		$row.Role | Should -Match 'SuperUsers'
		$row.'Action required' | Should -Match 'Verify the suggested privileged role'
	}

	It 'matches admin terms case-insensitively without matching unrelated words' {
		$script:fixture.Groups[0].DisplayName = 'Finance administration'
		$script:fixture.Groups[1].DisplayName = 'finance_admins'
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output
		$row = @(Import-Excel -Path $script:output -WorksheetName Packages)[0]
		$row.Role | Should -Be 'Potential privileged access (group name matches: admins)'
	}

	It 'does not suggest a role for privilege-like substrings in unrelated names' {
		$script:fixture.Groups[0].DisplayName = 'Finance administration'
		$script:fixture.Groups[1].DisplayName = 'Superuserland'
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output
		$row = @(Import-Excel -Path $script:output -WorksheetName Packages)[0]
		$row.Role | Should -BeNullOrEmpty
	}

	It 'matches separated super user terms without changing packages with no match' {
		$script:fixture.Applications[0].GroupIds = @($script:root, $script:other)
		$script:fixture.Groups[0].DisplayName = 'Application - SUPER USER'
		$script:fixture.Groups[2].ChildIds = @()
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output
		$rows = @(Import-Excel -Path $script:output -WorksheetName Packages)
		@($rows | Where-Object 'Discovery RootGroupId' -eq $script:root)[0].Role | Should -Match 'SUPER USER'
		@($rows | Where-Object 'Discovery RootGroupId' -eq $script:other)[0].Role | Should -BeNullOrEmpty
	}

	It 'combines multiple applications in one workbook while retaining each application ID' {
		$secondApp = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
		$script:fixture.SchemaVersion = 2
		$script:fixture.Remove('ServicePrincipalId')
		$script:fixture.Applications += @{
			ServicePrincipalId = $secondApp; ServicePrincipalName = 'Second app'; GroupIds = @($script:other)
		}
		$script:fixture.Groups[2].ChildIds = @()
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output
		$rows = @(Import-Excel -Path $script:output -WorksheetName Packages)
		$rows | Should -HaveCount 2
		$rows.'Discovery ServicePrincipalId' | Should -Contain $script:app
		$rows.'Discovery ServicePrincipalId' | Should -Contain $secondApp
		$rows.'Access Package Name' | Should -Contain 'Second app - Other'
		$metadata = Open-ExcelPackage -Path $script:output
		try {
			$metadata.Workbook.Worksheets['Discovery metadata'].Cells[1, 2].Value | Should -Be '3'
		}
		finally { Close-ExcelPackage -ExcelPackage $metadata -NoSave }
	}

	It 'keeps shared descendants in separate root packages' {
		$script:fixture.Applications[0].GroupIds = @($script:root, $script:other)
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output
		$rows = @(Import-Excel -Path $script:output -WorksheetName Packages)
		$rows | Should -HaveCount 2
		foreach ($row in $rows) {
			$row.GroupId.Split(';') | Should -Contain $script:child
			$row.Description | Should -Match 'Application B'
		}
		@(Import-Excel -Path $script:output -WorksheetName 'Manual review') | Should -HaveCount 0
	}

	It 'exports one package per unique discovered group without application assignment permission' {
		Mock Get-MgContext { @{ AuthType = 'AppOnly'; Scopes = @('EntitlementManagement.Read.All', 'Group.Read.All') } }
		$script:fixture.SchemaVersion = 2
		$script:fixture.Remove('ServicePrincipalId')
		$script:fixture.Applications += @{
			ServicePrincipalId = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'; ServicePrincipalName = 'Second app'; GroupIds = @($script:other)
		}
		$script:fixture.Groups[0].DirectUserCount = 0
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output -OnePackagePerGroup
		$rows = @(Import-Excel -Path $script:output -WorksheetName Packages)
		$rows | Should -HaveCount 3
		@($rows.GroupId | Sort-Object -Unique) | Should -HaveCount 3
		$rows.GroupId | Should -Contain $script:root
		$rows.GroupId | Should -Contain $script:child
		$rows.GroupId | Should -Contain $script:other
		foreach ($row in $rows) {
			$row.GroupId | Should -Be $row.'Discovery RootGroupId'
			$row.Description | Should -Match '^Included group: '
			$row.Description | Should -Not -Match 'application assignments'
		}
		@(Import-Excel -Path $script:output -WorksheetName 'Manual review') | Should -HaveCount 0
		@(Import-Excel -Path $script:output -WorksheetName Topology) | Should -HaveCount 5
		$excel = Open-ExcelPackage -Path $script:output
		try { $excel.Workbook.Worksheets['Discovery metadata'].Cells[1, 2].Value | Should -Be '4' }
		finally { Close-ExcelPackage -ExcelPackage $excel -NoSave }
		Should -Invoke Invoke-MgGraphRequest -Times 0 -Exactly -ParameterFilter { $Uri -like '*/appRoleAssignments' }
	}

	It 'supports an app-only connection with tenant-wide entitlement-management permission' {
		Mock Get-MgContext { @{ AuthType = 'AppOnly'; Scopes = @('EntitlementManagement.Read.All', 'Directory.Read.All') } }
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output
		@(Import-Excel -Path $script:output -WorksheetName Packages) | Should -HaveCount 1
		@(Import-Excel -Path $script:output -WorksheetName 'Manual review') | Should -HaveCount 0
	}

	It 'puts all candidates in manual review when app-only tenant-wide visibility cannot be established' {
		Mock Get-MgContext { @{ AuthType = 'AppOnly'; Scopes = @('Group.Read.All') } }
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output
		@(Import-Excel -Path $script:output -WorksheetName Packages) | Should -HaveCount 0
		$blocked = @(Import-Excel -Path $script:output -WorksheetName 'Manual review')
		$blocked | Should -HaveCount 1
		$blocked[0].Reason | Should -Match 'EntitlementManagement.Read.All'
	}

	It 'blocks a candidate when group application assignments cannot be read' {
		Mock Get-MgContext { @{ AuthType = 'AppOnly'; Scopes = @('EntitlementManagement.Read.All') } }
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output
		@(Import-Excel -Path $script:output -WorksheetName Packages) | Should -HaveCount 0
		$blocked = @(Import-Excel -Path $script:output -WorksheetName 'Manual review')
		$blocked[0].Reason | Should -Match 'Directory.Read.All'
	}
}
