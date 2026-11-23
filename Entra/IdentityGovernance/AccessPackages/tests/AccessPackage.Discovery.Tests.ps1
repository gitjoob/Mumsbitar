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
