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
		Mock Get-MgContext { @{ AuthType = 'Delegated' } }
		Mock Get-MgGroup {
			param($GroupId)
			@{ id = $GroupId; displayName = 'Group'; groupTypes = @(); securityEnabled = $true }
		}
		Mock Invoke-MgGraphRequest {
			param($Uri)
			if ($Uri -like '/v1.0/me?*') { return @{ id = '44444444-4444-4444-4444-444444444444' } }
			if ($Uri -like '*/roleAssignments?*') { return @{ value = @(@{ directoryScopeId = '/'; roleDefinitionId = '45d8d3c5-c802-45c6-b32a-1d70b5e1e86e' }) } }
			if ($Uri -like '*/accessPackages') { return @{ value = @() } }
			throw "Unexpected request $Uri"
		}
	}

	It 'writes a usable draft with literal text and the full cyclic descendant set' {
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output
		$rows = @(Import-Excel -Path $script:output -WorksheetName Packages)
		$rows | Should -HaveCount 1
		$rows[0].GroupId.Split(';') | Should -Contain $script:root
		$rows[0].GroupId.Split(';') | Should -Contain $script:child
		$rows[0].'Approver 1 List' | Should -BeNullOrEmpty
		$rows[0].Description | Should -Match 'adds the user to every listed group'
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
			$metadata.Workbook.Worksheets['Discovery metadata'].Cells[1, 2].Value | Should -Be '2'
		}
		finally { Close-ExcelPackage -ExcelPackage $metadata -NoSave }
	}

	It 'keeps shared descendants out of the import sheet' {
		$script:fixture.Applications[0].GroupIds = @($script:root, $script:other)
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output
		@(Import-Excel -Path $script:output -WorksheetName Packages) | Should -HaveCount 0
		@(Import-Excel -Path $script:output -WorksheetName 'Manual review') | Should -HaveCount 2
	}

	It 'puts all candidates in manual review when tenant-wide visibility cannot be established' {
		Mock Get-MgContext { @{ AuthType = 'AppOnly' } }
		$script:fixture | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:input
		& $script:exportPath -DiscoveryPath $script:input -WorkbookPath $script:output
		@(Import-Excel -Path $script:output -WorksheetName Packages) | Should -HaveCount 0
		$blocked = @(Import-Excel -Path $script:output -WorksheetName 'Manual review')
		$blocked | Should -HaveCount 1
		$blocked[0].Reason | Should -Match 'visibility cannot be verified'
	}
}
