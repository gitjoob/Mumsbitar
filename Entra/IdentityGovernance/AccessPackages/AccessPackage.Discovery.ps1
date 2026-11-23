function Get-DiscoveryPackageWarning {
	'Assigning this package adds the user to every listed group and may grant access through other groups or applications. Nested membership does not establish effective application access.'
}

function Get-PackageDiscoveryCandidates {
	param([Parameter(Mandatory)][object]$Data)
	if ($Data.SchemaVersion -ne 1 -or @($Data.Applications).Count -ne 1) { throw 'Expected a version 1 single-application discovery export.' }
	$appId = [guid]::Empty
	if (-not [guid]::TryParse([string]$Data.ServicePrincipalId, [ref]$appId) -or $appId -eq [guid]::Empty -or
		"$appId" -ne "$($Data.Applications[0].ServicePrincipalId)") { throw 'Discovery export has an invalid or inconsistent service principal ID.' }
	$groups = @{}
	foreach ($group in $Data.Groups) {
		$id = [guid]::Empty
		if (-not [guid]::TryParse([string]$group.Id, [ref]$id) -or $id -eq [guid]::Empty -or $groups.ContainsKey("$id")) {
			throw "Discovery export contains a duplicate or invalid group ID '$($group.Id)'."
		}
		$groups["$id"] = $group
	}
	$candidates = [System.Collections.Generic.List[object]]::new()
	$roots = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
	foreach ($rootValue in $Data.Applications[0].GroupIds) {
		$root = [guid]::Empty
		if (-not [guid]::TryParse([string]$rootValue, [ref]$root) -or $root -eq [guid]::Empty -or
			-not $roots.Add("$root") -or -not $groups.ContainsKey("$root")) { throw "Discovery export contains an invalid root group '$rootValue'." }
		$seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
		$queue = [System.Collections.Generic.Queue[string]]::new()
		$queue.Enqueue("$root")
		while ($queue.Count) {
			$id = $queue.Dequeue()
			if (-not $seen.Add($id)) { continue }
			if (-not $groups.ContainsKey($id)) { throw "Discovery export is missing descendant group '$id'." }
			foreach ($childValue in $groups[$id].ChildIds) {
				$child = [guid]::Empty
				if (-not [guid]::TryParse([string]$childValue, [ref]$child) -or $child -eq [guid]::Empty) {
					throw "Discovery export contains an invalid child group '$childValue'."
				}
				$queue.Enqueue("$child")
			}
		}
		$candidates.Add([pscustomobject]@{
			Root = "$root"; GroupIds = @($seen | Sort-Object); Status = 'Ready'; Reason = ''
		})
	}
	foreach ($candidate in $candidates) {
		$overlappingRoots = @($candidates | Where-Object {
			$other = $_
			$other.Root -ne $candidate.Root -and @($other.GroupIds | Where-Object { $_ -in $candidate.GroupIds }).Count
		} | ForEach-Object { $_.Root })
		if ($overlappingRoots.Count) {
			$candidate.Status = 'Manual review'
			$candidate.Reason = "Shares groups with proposed root(s): $($overlappingRoots -join ', ')."
		}
	}
	[pscustomobject]@{ ServicePrincipalId = "$appId"; Groups = $groups; Candidates = $candidates.ToArray() }
}

function Get-CurrentPackageDiscovery {
	param(
		[Parameter(Mandatory)][guid]$ServicePrincipalId,
		[Parameter(Mandatory)][string]$ReviewScriptPath,
		[Parameter(Mandatory)][string]$TemporaryParent
	)
	$directory = Join-Path $TemporaryParent ('.package-discovery-' + [guid]::NewGuid().ToString())
	$path = Join-Path $directory 'discovery.json'
	$createdDirectory = $false
	try {
		$null = New-Item -ItemType Directory -Path $directory -WhatIf:$false -ErrorAction Stop
		$createdDirectory = $true
		$null = @(& $ReviewScriptPath -ServicePrincipalId $ServicePrincipalId -ReviewerType ResourceOwner `
			-DiscoveryOutputPath $path -WhatIf -ErrorAction Stop)
		if (-not [IO.File]::Exists($path)) { throw 'The review preview did not produce complete discovery data.' }
		$data = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
		$discovery = Get-PackageDiscoveryCandidates -Data $data
		if ($discovery.ServicePrincipalId -ne "$ServicePrincipalId") { throw 'The refreshed discovery has a different application ID.' }
		$discovery
	}
	finally {
		if ($createdDirectory) {
			if ([IO.File]::Exists($path)) { [IO.File]::Delete($path) }
			[IO.Directory]::Delete($directory)
		}
	}
}

function Test-DiscoveryPackageCandidate {
	param(
		[Parameter(Mandatory)][object]$Discovery,
		[Parameter(Mandatory)][guid]$RootGroupId,
		[Parameter(Mandatory)][string[]]$GroupIds
	)
	try {
		$candidate = @($discovery.Candidates | Where-Object Root -eq "$RootGroupId")
		if ($candidate.Count -ne 1) { throw 'The root is no longer directly assigned to the application.' }
		if ($candidate[0].Status -ne 'Ready') { throw $candidate[0].Reason }
		$actual = @($GroupIds | Sort-Object -Unique)
		$expected = @($candidate[0].GroupIds | Sort-Object -Unique)
		if ($actual.Count -ne $expected.Count -or @(Compare-Object $actual $expected).Count) {
			throw 'The workbook groups do not match the current root and descendant groups. Regenerate the workbook.'
		}
		$inventory = @(Get-PackageCoverageInventory)
		Get-PackageCoverageResult -GroupIds $actual -Inventory $inventory
	}
	catch {
		[pscustomobject]@{ Status = 'Manual review'; Reason = "Unable to validate discovery and package coverage: $($_.Exception.Message)" }
	}
}
