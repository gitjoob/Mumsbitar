BeforeAll {
	. (Join-Path $PSScriptRoot '../AccessPackage.Discovery.ps1')
	. (Join-Path $PSScriptRoot '../AccessPackage.Coverage.ps1')
	$script:app = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'
	$script:root = '11111111-1111-1111-1111-111111111111'
	$script:child = '22222222-2222-2222-2222-222222222222'
	$script:other = '33333333-3333-3333-3333-333333333333'
}

Describe 'Discovery package creation safeguards' {
	BeforeEach {
		$script:data = [pscustomobject]@{
			SchemaVersion = 1; ServicePrincipalId = $script:app
			Applications = @([pscustomobject]@{ ServicePrincipalId = $script:app; GroupIds = @($script:root) })
			Groups = @(
				[pscustomobject]@{ Id = $script:root; ChildIds = @($script:child) }
				[pscustomobject]@{ Id = $script:child; ChildIds = @() }
				[pscustomobject]@{ Id = $script:other; ChildIds = @($script:child) }
			)
		}
		Mock Get-PackageCoverageInventory { @() }
		Mock Get-PackageCoverageResult { [pscustomobject]@{ Status = 'Ready'; Reason = '' } }
	}

	It 'accepts only the refreshed root and complete descendant set' {
		$discovery = Get-PackageDiscoveryCandidates $script:data
		(Test-DiscoveryPackageCandidate -Discovery $discovery -RootGroupId $script:root -GroupIds @($script:child, $script:root)).Status | Should -Be 'Ready'
		$result = Test-DiscoveryPackageCandidate -Discovery $discovery -RootGroupId $script:root -GroupIds @($script:root)
		$result.Status | Should -Be 'Manual review'
		$result.Reason | Should -Match 'do not match'
		(Test-DiscoveryPackageCandidate -Discovery $discovery -RootGroupId $script:other -GroupIds @($script:other)).Status | Should -Be 'Manual review'
	}

	It 'blocks shared descendants and failed coverage checks' {
		$script:data.Applications[0].GroupIds = @($script:root, $script:other)
		$discovery = Get-PackageDiscoveryCandidates $script:data
		(Test-DiscoveryPackageCandidate -Discovery $discovery -RootGroupId $script:root -GroupIds @($script:root, $script:child)).Status | Should -Be 'Manual review'
		$script:data.Applications[0].GroupIds = @($script:root)
		$discovery = Get-PackageDiscoveryCandidates $script:data
		Mock Get-PackageCoverageInventory { throw 'Forbidden' }
		$result = Test-DiscoveryPackageCandidate -Discovery $discovery -RootGroupId $script:root -GroupIds @($script:root, $script:child)
		$result.Status | Should -Be 'Manual review'
		$result.Reason | Should -Match 'Forbidden'
	}

	It 'supports combined application discovery and identifies each candidate by application' {
		$secondApp = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
		$script:data.SchemaVersion = 2
		$script:data.PSObject.Properties.Remove('ServicePrincipalId')
		$script:data.Applications += [pscustomobject]@{
			ServicePrincipalId = $secondApp; ServicePrincipalName = 'Second app'; GroupIds = @($script:other)
		}
		$script:data.Applications[0] | Add-Member -NotePropertyName ServicePrincipalName -NotePropertyValue 'First app'
		$script:data.Groups[2].ChildIds = @()
		$discovery = Get-PackageDiscoveryCandidates $script:data
		$discovery.ServicePrincipalId | Should -BeNullOrEmpty
		$discovery.ServicePrincipalIds | Should -Be @($script:app, $secondApp)
		$discovery.Candidates | Should -HaveCount 2
		$discovery.Candidates[0].ServicePrincipalName | Should -Be 'First app'
		(Test-DiscoveryPackageCandidate -Discovery $discovery -ServicePrincipalId $secondApp `
			-RootGroupId $script:other -GroupIds @($script:other)).Status | Should -Be 'Ready'
	}

	It 'marks roots shared across applications for manual review' {
		$secondApp = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb'
		$script:data.SchemaVersion = 2
		$script:data.PSObject.Properties.Remove('ServicePrincipalId')
		$script:data.Applications += [pscustomobject]@{
			ServicePrincipalId = $secondApp; GroupIds = @($script:root)
		}
		$discovery = Get-PackageDiscoveryCandidates $script:data
		$discovery.Candidates | Should -HaveCount 2
		$discovery.Candidates.Status | Should -Be @('Manual review', 'Manual review')
	}

	It 'can retain separate candidates when roots share a group' {
		$script:data.Applications[0].GroupIds = @($script:root, $script:other)
		$discovery = Get-PackageDiscoveryCandidates $script:data -AllowSharedGroups
		$discovery.Candidates | Should -HaveCount 2
		$discovery.Candidates.Status | Should -Be @('Ready', 'Ready')
		foreach ($candidate in $discovery.Candidates) { $candidate.GroupIds | Should -Contain $script:child }
	}

	It 'validates a root-only package resource against its discovered root' {
		$discovery = Get-PackageDiscoveryCandidates $script:data
		(Test-DiscoveryPackageCandidate -Discovery $discovery -RootGroupId $script:root `
			-GroupIds @($script:root) -RootOnly).Status | Should -Be 'Ready'
		$result = Test-DiscoveryPackageCandidate -Discovery $discovery -RootGroupId $script:root `
			-GroupIds @($script:root, $script:child) -RootOnly
		$result.Status | Should -Be 'Manual review'
		$result.Reason | Should -Match 'only its original discovered root group'
	}

	It 'describes only the applications directly assigned to the included groups' {
		Mock Get-MgContext { @{ AuthType = 'AppOnly'; Scopes = @('Directory.Read.All') } }
		Mock Get-PackageCoverageCollection {
			@(@{ resourceId = $script:app; resourceDisplayName = 'Reporting' })
		}
		$groups = @{
			$script:root = @{ DisplayName = 'Readers' }
			$script:child = @{ DisplayName = 'Shared' }
		}
		$cache = @{}
		$description = Get-DiscoveryPackageDescription -GroupIds @($script:root, $script:child) -Groups $groups -ApplicationCache $cache
		$description | Should -Be 'Included groups: Readers; Shared. Direct Entra application assignments for these groups: Reporting.'
		$null = Get-DiscoveryPackageDescription -GroupIds @($script:child) -Groups $groups -ApplicationCache $cache
		Should -Invoke Get-PackageCoverageCollection -Exactly -Times 2
	}

	It 'rejects incomplete group application assignment data' {
		Mock Get-MgContext { @{ AuthType = 'AppOnly'; Scopes = @('Directory.Read.All') } }
		Mock Get-PackageCoverageCollection { @(@{ resourceId = $script:app }) }
		{ Get-DiscoveryGroupApplications -GroupId $script:root } | Should -Throw '*incomplete application assignment*'
	}

	It 'refreshes discovery with WhatIf and removes only its own temporary files' {
		$fixture = Join-Path $TestDrive "preview-$([guid]::NewGuid()).ps1"
		$json = $script:data | ConvertTo-Json -Depth 10 -Compress
		@"
[CmdletBinding(SupportsShouldProcess)]
param([guid]`$ServicePrincipalId, [string]`$ReviewerType, [string]`$DiscoveryOutputPath)
if (-not `$WhatIfPreference) { throw 'Preview was not requested.' }
@'
$json
'@ | Set-Content -LiteralPath `$DiscoveryOutputPath -WhatIf:`$false
"@ | Set-Content -LiteralPath $fixture
		$WhatIfPreference = $true
		$discovery = Get-CurrentPackageDiscovery -ServicePrincipalId $script:app -ReviewScriptPath $fixture -TemporaryParent $TestDrive
		$discovery.Candidates[0].GroupIds | Should -Contain $script:child
		@(Get-ChildItem -LiteralPath $TestDrive -Filter '.package-discovery-*' -Force) | Should -HaveCount 0
	}
}
