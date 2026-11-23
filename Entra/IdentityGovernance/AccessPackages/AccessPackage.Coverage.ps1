function Get-PackageCoverageCollection {
	param([Parameter(Mandatory)][string]$Uri)
	do {
		$page = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Hashtable -ErrorAction Stop
		if ($page -isnot [System.Collections.IDictionary] -or -not $page.Contains('value') -or $null -eq $page.value) {
			throw "Graph did not return a complete collection for '$Uri'."
		}
		$page.value
		$Uri = $page.'@odata.nextLink'
	} while ($Uri)
}

function Assert-PackageCoverageVisibility {
	$context = Get-MgContext
	if (-not $context) {
		throw 'Tenant-wide package visibility cannot be verified without a Microsoft Graph connection.'
	}
	if ($context.AuthType -eq 'AppOnly') {
		if ($context.Scopes -notcontains 'EntitlementManagement.Read.All' -and
			$context.Scopes -notcontains 'EntitlementManagement.ReadWrite.All') {
			throw 'Tenant-wide package visibility requires the EntitlementManagement.Read.All or EntitlementManagement.ReadWrite.All Microsoft Graph application permission in the current connection. Grant admin consent and reconnect.'
		}
		return
	}
	if ($context.AuthType -ne 'Delegated') {
		throw "Tenant-wide package visibility cannot be verified for Graph authentication type '$($context.AuthType)'. Use an AppOnly or delegated connection."
	}
	$me = Invoke-MgGraphRequest -Method GET -Uri '/v1.0/me?$select=id' -OutputType Hashtable -ErrorAction Stop
	$id = [guid]::Empty
	if (-not [guid]::TryParse([string]$me.id, [ref]$id) -or $id -eq [guid]::Empty) {
		throw 'Graph did not return the signed-in user ID; tenant-wide package visibility cannot be verified.'
	}
	$filter = [uri]::EscapeDataString("principalId eq '$id'")
	$assignments = @(Get-PackageCoverageCollection "/v1.0/roleManagement/directory/roleAssignments?`$filter=$filter")
	$allowedRoles = @('62e90394-69f5-4237-9190-012177145e10', '45d8d3c5-c802-45c6-b32a-1d70b5e1e86e')
	if (-not @($assignments | Where-Object { $_.directoryScopeId -eq '/' -and $_.roleDefinitionId -in $allowedRoles }).Count) {
		throw 'Tenant-wide package visibility requires a directly assigned, active directory-wide Identity Governance Administrator or Global Administrator role. Catalog-only roles cannot prove complete visibility.'
	}
}

function Get-PackageCoverageActivity {
	param([Parameter(Mandatory)][string]$PackageId)
	$base = '/v1.0/identityGovernance/entitlementManagement'
	$filter = [uri]::EscapeDataString("accessPackage/id eq '$PackageId'")
	$policies = @(Get-PackageCoverageCollection "$base/assignmentPolicies?`$filter=$filter")
	$assignments = @(Get-PackageCoverageCollection "$base/assignments?`$filter=$filter")
	$requests = @(Get-PackageCoverageCollection "$base/assignmentRequests?`$filter=$filter")
	foreach ($assignment in $assignments) {
		if ($assignment.state -in @('delivering', 'partiallyDelivered', 'delivered')) {
			$start = [datetimeoffset]::MinValue
			if (-not $assignment.schedule -or
				-not [datetimeoffset]::TryParse([string]$assignment.schedule.startDateTime, [ref]$start)) { return 'Unknown' }
			if ($start -gt [datetimeoffset]::UtcNow) { return 'Unknown' }
			$expiration = $assignment.schedule.expiration
			if (-not $expiration) { return 'Unknown' }
			switch ($expiration.type) {
				'noExpiration' { return 'Active' }
				'afterDateTime' {
					$end = [datetimeoffset]::MinValue
					if (-not [datetimeoffset]::TryParse([string]$expiration.endDateTime, [ref]$end)) { return 'Unknown' }
					if ($end -gt [datetimeoffset]::UtcNow) { return 'Active' }
				}
				'afterDuration' {
					try { $duration = [System.Xml.XmlConvert]::ToTimeSpan([string]$expiration.duration) }
					catch [System.FormatException] { return 'Unknown' }
					if ($duration -le [timespan]::Zero) { return 'Unknown' }
					if ($start.Add($duration) -gt [datetimeoffset]::UtcNow) { return 'Active' }
				}
				default { return 'Unknown' }
			}
			return 'Unknown'
		}
		if ($assignment.state -ne 'expired') { return 'Unknown' }
	}
	foreach ($policy in $policies) {
		if ($policy.requestorSettings.enableTargetsToSelfAddAccess -eq $true -and
			$policy.allowedTargetScope -in @('allMemberUsers', 'allDirectoryUsers')) {
			switch ($policy.expiration.type) {
				'noExpiration' { return 'Active' }
				'afterDuration' {
					try { $duration = [System.Xml.XmlConvert]::ToTimeSpan([string]$policy.expiration.duration) }
					catch [System.FormatException] { continue }
					if ($duration -gt [timespan]::Zero) { return 'Active' }
				}
				'afterDateTime' {
					$end = [datetimeoffset]::MinValue
					if ([datetimeoffset]::TryParse([string]$policy.expiration.endDateTime, [ref]$end) -and $end -gt [datetimeoffset]::UtcNow) {
						return 'Active'
					}
				}
			}
		}
	}
	if ($policies.Count) { return 'Unknown' }
	foreach ($request in $requests) {
		if ($request.state -notin @('delivered', 'denied', 'canceled')) { return 'Unknown' }
	}
	'Inactive'
}

function Get-PackageCoverageInventory {
	Assert-PackageCoverageVisibility
	$base = '/v1.0/identityGovernance/entitlementManagement'
	$packages = @(Get-PackageCoverageCollection "$base/accessPackages")
	foreach ($package in $packages) {
		$id = [guid]::Empty
		if (-not [guid]::TryParse([string]$package.id, [ref]$id) -or $id -eq [guid]::Empty) {
			throw 'Graph returned an access package without a valid ID; package coverage cannot be assessed.'
		}
		$detail = Invoke-MgGraphRequest -Method GET -Uri "$base/accessPackages/$id`?`$expand=resourceRoleScopes(`$expand=role,scope)" -OutputType Hashtable -ErrorAction Stop
		if ($detail -isnot [System.Collections.IDictionary] -or -not $detail.Contains('resourceRoleScopes') -or $null -eq $detail.resourceRoleScopes) {
			throw "Graph did not return resource roles and scopes for package '$id'."
		}
		$scopes = @($detail.resourceRoleScopes)
		if ($detail.'resourceRoleScopes@odata.nextLink') {
			$scopes += @(Get-PackageCoverageCollection $detail.'resourceRoleScopes@odata.nextLink')
		}
		$groupIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
		$memberIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
		$hasExtra = $false
		foreach ($item in $scopes) {
			if (-not $item.role -or -not $item.scope -or -not $item.role.originSystem -or -not $item.role.originId -or
				-not $item.scope.originSystem -or -not $item.scope.originId -or $null -eq $item.scope.isRootScope) {
				throw "Graph returned an incomplete resource role or scope for package '$id'."
			}
			$groupId = [guid]::Empty
			if ($item.scope.originSystem -eq 'AadGroup' -and [guid]::TryParse([string]$item.scope.originId, [ref]$groupId) -and $groupId -ne [guid]::Empty) {
				$null = $groupIds.Add("$groupId")
				if ($item.role.originSystem -eq 'AadGroup') {
					$roleGroup = [guid]::Empty
					if ($item.role.originId -notmatch '^[^_]+_([0-9a-fA-F-]{36})$' -or
						-not [guid]::TryParse($Matches[1], [ref]$roleGroup) -or $roleGroup -eq [guid]::Empty) {
						throw "Graph returned an unidentifiable group role for package '$id'."
					}
					$null = $groupIds.Add("$roleGroup")
				}
				if ($item.scope.isRootScope -eq $true -and $item.role.originSystem -eq 'AadGroup' -and
					$item.role.originId -eq "Member_$groupId") {
					if (-not $memberIds.Add("$groupId")) { $hasExtra = $true }
				}
				else { $hasExtra = $true }
			}
			elseif ($item.role.originSystem -eq 'AadGroup' -or $item.scope.originSystem -eq 'AadGroup') {
				throw "Graph returned an unidentifiable group resource role or scope for package '$id'."
			}
			else { $hasExtra = $true }
		}
		[pscustomobject]@{
			Id = "$id"
			DisplayName = [string]$package.displayName
			CatalogId = [string]$package.catalogId
			GroupIds = @($groupIds)
			MemberGroupIds = @($memberIds)
			HasExtra = $hasExtra
		}
	}
}

function Get-PackageCoverageResult {
	param([Parameter(Mandatory)][string[]]$GroupIds, [AllowEmptyCollection()][object[]]$Inventory)
	$wanted = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
	foreach ($id in $GroupIds) {
		$parsed = [guid]::Empty
		if (-not [guid]::TryParse($id, [ref]$parsed) -or $parsed -eq [guid]::Empty) { throw "Invalid group ID '$id'." }
		$null = $wanted.Add("$parsed")
	}
	if (-not $wanted.Count) { throw 'At least one group is required for a package coverage check.' }
	$overlaps = @(foreach ($package in $Inventory) {
		if (@($package.GroupIds | Where-Object { $wanted.Contains($_) }).Count) { $package }
	})
	$manual = [System.Collections.Generic.List[string]]::new()
	foreach ($package in $overlaps) {
		$exact = -not $package.HasExtra -and $package.MemberGroupIds.Count -eq $wanted.Count -and
			@($package.MemberGroupIds | Where-Object { -not $wanted.Contains($_) }).Count -eq 0
		if (-not $exact) {
			$manual.Add("Partial resource overlap with '$($package.DisplayName)' ($($package.Id)).")
			continue
		}
		$activity = Get-PackageCoverageActivity -PackageId $package.Id
		if ($activity -eq 'Active') {
			return [pscustomobject]@{ Status = 'Skipped'; Reason = "Active exact resource match: '$($package.DisplayName)' ($($package.Id))." }
		}
		if ($activity -ne 'Inactive') {
			$manual.Add("Exact resource match '$($package.DisplayName)' ($($package.Id)) has ambiguous activity.")
		}
	}
	if ($manual.Count) {
		return [pscustomobject]@{ Status = 'Manual review'; Reason = ($manual -join ' ') }
	}
	[pscustomobject]@{ Status = 'Ready'; Reason = '' }
}
