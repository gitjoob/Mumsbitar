function Get-DiscoveryPackageWarning {
	'Assigning this package adds the user to every listed group and may grant access through other groups or applications. Nested membership does not establish effective application access.'
}

function Get-PackageDiscoveryCandidates {
	param([Parameter(Mandatory)][object]$Data, [switch]$AllowSharedGroups)
	$applications = @($Data.Applications)
	if ($Data.SchemaVersion -eq 1) {
		if ($applications.Count -ne 1) { throw 'Expected a version 1 single-application discovery export.' }
	}
	elseif ($Data.SchemaVersion -eq 2) {
		if ($applications.Count -lt 1) { throw 'Discovery export contains no applications.' }
	}
	else { throw "Unsupported discovery export version '$($Data.SchemaVersion)'." }
	$groups = @{}
	foreach ($group in $Data.Groups) {
		$id = [guid]::Empty
		if (-not [guid]::TryParse([string]$group.Id, [ref]$id) -or $id -eq [guid]::Empty -or $groups.ContainsKey("$id")) {
			throw "Discovery export contains a duplicate or invalid group ID '$($group.Id)'."
		}
		$groups["$id"] = $group
	}
	$candidates = [System.Collections.Generic.List[object]]::new()
	$applicationIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
	foreach ($application in $applications) {
		$appId = [guid]::Empty
		if (-not [guid]::TryParse([string]$application.ServicePrincipalId, [ref]$appId) -or $appId -eq [guid]::Empty -or
			-not $applicationIds.Add("$appId")) { throw 'Discovery export contains a duplicate or invalid service principal ID.' }
		if ($Data.SchemaVersion -eq 1) {
			$topLevelAppId = [guid]::Empty
			if (-not [guid]::TryParse([string]$Data.ServicePrincipalId, [ref]$topLevelAppId) -or
				"$topLevelAppId" -ne "$appId") { throw 'Discovery export has an invalid or inconsistent service principal ID.' }
		}
		$roots = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
		foreach ($rootValue in $application.GroupIds) {
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
				ServicePrincipalId = "$appId"
				ServicePrincipalName = [string]$application.ServicePrincipalName
				Root = "$root"; GroupIds = @($seen | Sort-Object); Status = 'Ready'; Reason = ''; Description = ''
			})
		}
	}
	if (-not $AllowSharedGroups) {
		foreach ($candidate in $candidates) {
			$overlappingRoots = @($candidates | Where-Object {
				$other = $_
				($other.ServicePrincipalId -ne $candidate.ServicePrincipalId -or $other.Root -ne $candidate.Root) -and
					@($other.GroupIds | Where-Object { $_ -in $candidate.GroupIds }).Count
			} | ForEach-Object { "$($_.ServicePrincipalId)/$($_.Root)" })
			if ($overlappingRoots.Count) {
				$candidate.Status = 'Manual review'
				$candidate.Reason = "Shares groups with proposed root(s): $($overlappingRoots -join ', ')."
			}
		}
	}
	[pscustomobject]@{
		ServicePrincipalId = if ($applicationIds.Count -eq 1) { @($applicationIds)[0] } else { $null }
		ServicePrincipalIds = @($applicationIds | Sort-Object)
		Groups = $groups
		Candidates = $candidates.ToArray()
	}
}

function Get-DiscoveryGroupApplications {
	param([Parameter(Mandatory)][guid]$GroupId)
	$context = Get-MgContext
	if (-not $context -or -not @($context.Scopes | Where-Object {
		$_ -in @('Directory.Read.All', 'Directory.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
	}).Count) {
		throw 'Reading group application assignments requires Directory.Read.All (or a supported higher Graph permission) in the current connection.'
	}
	$applications = @{}
	foreach ($assignment in @(Get-PackageCoverageCollection "/v1.0/groups/$GroupId/appRoleAssignments")) {
		$id = [guid]::Empty
		$name = [string]$assignment.resourceDisplayName
		if (-not [guid]::TryParse([string]$assignment.resourceId, [ref]$id) -or $id -eq [guid]::Empty -or
			[string]::IsNullOrWhiteSpace($name)) {
			throw "Graph returned an incomplete application assignment for group '$GroupId'."
		}
		if ($applications.ContainsKey("$id") -and $applications["$id"] -ne $name) {
			throw "Graph returned inconsistent application names for '$id'."
		}
		$applications["$id"] = $name
	}
	foreach ($id in @($applications.Keys | Sort-Object)) {
		[pscustomobject]@{ Id = "$id"; Name = $applications[$id] }
	}
}

function Get-DiscoveryPackageDescription {
	param(
		[Parameter(Mandatory)][string[]]$GroupIds,
		[Parameter(Mandatory)][hashtable]$Groups,
		[Parameter(Mandatory)][hashtable]$ApplicationCache,
		[string]$RootGroupId
	)
	if ($RootGroupId -and $RootGroupId -notin $GroupIds) {
		throw "Root group '$RootGroupId' is missing from the discovered group set."
	}
	$names = [System.Collections.Generic.List[string]]::new()
	$applications = @{}
	$groupApplications = @{}
	foreach ($id in @($GroupIds | Sort-Object -Unique)) {
		if (-not $Groups.ContainsKey($id) -or [string]::IsNullOrWhiteSpace([string]$Groups[$id].DisplayName)) {
			throw "Cannot establish the name of group '$id'."
		}
		$names.Add([string]$Groups[$id].DisplayName)
		if (-not $ApplicationCache.ContainsKey($id)) {
			$ApplicationCache[$id] = @(Get-DiscoveryGroupApplications -GroupId ([guid]$id))
		}
		$groupApplications[$id] = @($ApplicationCache[$id])
		foreach ($application in $ApplicationCache[$id]) {
			if ($applications.ContainsKey($application.Id) -and $applications[$application.Id] -ne $application.Name) {
				throw "Graph returned inconsistent application names for '$($application.Id)'."
			}
			$applications[$application.Id] = $application.Name
		}
	}
	$appNames = @($applications.Keys | ForEach-Object {
		$id = $_
		$name = $applications[$id]
		if (@($applications.Values | Where-Object { $_ -eq $name }).Count -gt 1) { "$name ($id)" }
		else { $name }
	} | Sort-Object)
	$appText = if ($appNames.Count) { $appNames -join '; ' } else { 'None' }
	if ($RootGroupId) {
		$rootName = [string]$Groups[$RootGroupId].DisplayName
		$nestedIds = @($GroupIds | Where-Object { $_ -ne $RootGroupId } | Sort-Object -Unique)
		$orderedIds = @($RootGroupId) + $nestedIds
		$nestedNames = @($nestedIds | ForEach-Object { [string]$Groups[$_].DisplayName })
		$assignmentLines = @(
			foreach ($id in $orderedIds) {
				$groupName = [string]$Groups[$id].DisplayName
				$assignedNames = @(
					foreach ($application in $groupApplications[$id]) {
						if (@($applications.Values | Where-Object { $_ -eq $application.Name }).Count -gt 1) {
							"$($application.Name) ($($application.Id))"
						}
						else { $application.Name }
					}
				) | Sort-Object
				$assignedText = if ($assignedNames.Count) { $assignedNames -join '; ' } else { 'None' }
				"$groupName`: $assignedText"
			}
		) -join '. '
		$nestedText = if ($nestedNames.Count) {
			" Nested groups in the root group's membership hierarchy: $($nestedNames -join '; '). Existing members of these groups are transitively members of the root group and may receive access through the root group's assignments."
		}
		else { ' No nested groups were discovered.' }
		$nestedAssignmentText = if ($nestedNames.Count) {
			' Assignments shown for nested groups apply to those groups and their own members; being assigned to the root through this package does not add a user to a nested group.'
		}
		else { '' }
		"Package resource: $rootName. This package grants assigned users membership in the root group only; it does not add them to nested groups.$nestedText Direct Entra enterprise application assignments by group: $assignmentLines.$nestedAssignmentText"
	}
	else {
		"Included groups: $($names -join '; '). Direct Entra application assignments for these groups: $appText."
	}
}

function Get-DiscoverySingleGroupDescription {
	param([Parameter(Mandatory)][string]$GroupId, [Parameter(Mandatory)][hashtable]$Groups)
	if (-not $Groups.ContainsKey($GroupId) -or [string]::IsNullOrWhiteSpace([string]$Groups[$GroupId].DisplayName)) {
		throw "Cannot establish the name of group '$GroupId'."
	}
	"Included group: $($Groups[$GroupId].DisplayName)."
}

function Get-CurrentPackageDiscovery {
	param(
		[Parameter(Mandatory)][guid]$ServicePrincipalId,
		[Parameter(Mandatory)][string]$ReviewScriptPath,
		[Parameter(Mandatory)][string]$TemporaryParent,
		[switch]$AllowSharedGroups
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
		$discovery = Get-PackageDiscoveryCandidates -Data $data -AllowSharedGroups:$AllowSharedGroups
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
		[Parameter(Mandatory)][string[]]$GroupIds,
		[guid]$ServicePrincipalId,
		[switch]$AllowSharedGroups,
		[switch]$SingleGroup,
		[switch]$RootOnly
	)
	try {
		$actual = @($GroupIds | Sort-Object -Unique)
		if ($RootOnly) {
			if ($actual.Count -ne 1 -or $actual[0] -ne "$RootGroupId") {
				throw 'The workbook must contain only its original discovered root group. Regenerate the workbook.'
			}
			$candidate = @($Discovery.Candidates | Where-Object {
				$_.Root -eq "$RootGroupId" -and (-not $ServicePrincipalId -or $_.ServicePrincipalId -eq "$ServicePrincipalId")
			})
			if ($candidate.Count -ne 1) { throw 'The root is no longer directly assigned to the application.' }
			if ($candidate[0].Status -ne 'Ready') { throw $candidate[0].Reason }
		}
		elseif ($SingleGroup) {
			if ($actual.Count -ne 1 -or $actual[0] -ne "$RootGroupId") {
				throw 'The workbook must contain exactly its original discovered group. Regenerate the workbook.'
			}
			$reachable = @($Discovery.Candidates | Where-Object {
				(-not $ServicePrincipalId -or $_.ServicePrincipalId -eq "$ServicePrincipalId") -and
					$_.GroupIds -contains "$RootGroupId"
			})
			if (-not $reachable.Count) { throw 'The group is no longer in the application discovery. Regenerate the workbook.' }
		}
		else {
			$candidate = @($Discovery.Candidates | Where-Object {
				$_.Root -eq "$RootGroupId" -and (-not $ServicePrincipalId -or $_.ServicePrincipalId -eq "$ServicePrincipalId")
			})
			if ($candidate.Count -ne 1) { throw 'The root is no longer directly assigned to the application.' }
			if ($candidate[0].Status -ne 'Ready') { throw $candidate[0].Reason }
			$expected = @($candidate[0].GroupIds | Sort-Object -Unique)
			if ($actual.Count -ne $expected.Count -or @(Compare-Object $actual $expected).Count) {
				throw 'The workbook groups do not match the current root and descendant groups. Regenerate the workbook.'
			}
		}
		$inventory = @(Get-PackageCoverageInventory)
		Get-PackageCoverageResult -GroupIds $actual -Inventory $inventory -AllowSharedGroups:$AllowSharedGroups
	}
	catch {
		[pscustomobject]@{ Status = 'Manual review'; Reason = "Unable to validate discovery and package coverage: $($_.Exception.Message)" }
	}
}
