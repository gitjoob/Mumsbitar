<#
.SYNOPSIS
Creates group-based access packages from a CSV file.
.DESCRIPTION
Connect to Microsoft Graph first. Requires EntitlementManagement.ReadWrite.All,
group read permissions, user read permissions, and appropriate licensing and roles.
CatalogName is mandatory. An exact matching catalog is reused; a missing catalog
is created. Multiple matching catalogs cause the script to stop.

CSV columns: Name, Description, GroupId, AppOwnerIds (semicolon-separated user IDs).
DefaultAppOwnerIds supplies owners when a row's AppOwnerIds is empty.
FallbackApproverIds is mandatory and is used when a recipient has no manager.
Names and descriptions must use business-friendly application and access terms.
Names must not contain "Access Package", "AP", or "SG", or identify the owner.

All directory member users, excluding guests, can request access. Managers can
request on behalf of their employees; authorized administrators retain their
standard assignment permissions. Justification and approval are required.
Up to 500 transitive user members uses manager approval; more than 500 adds
the specified application owners as a second stage. Each stage allows 14 days.
Assignments expire after 365 days. Custom schedules and approved extensions
are enabled. Assignment emails are enabled; access reviews are disabled.

MANUAL STEP: Enable "Bypass approval stage if manager is the requester and
approver" in the Entra portal for each new policy. The public Graph policy
schema does not expose this switch. The script does not claim to configure it.

Existing packages with the same name in the catalog are skipped, not updated.
Each row produces a result object; skipped/failed rows and a summary are printed.
WhatIf performs reads and validation but never writes to Microsoft Graph.
.EXAMPLE
.\Create-AccessPackage.ps1 -CsvPath '.\access-packages.csv' -CatalogName 'Business applications' -FallbackApproverIds '11111111-1111-1111-1111-111111111111' -WhatIf
.EXAMPLE
.\Create-AccessPackage.ps1 -CsvPath '.\access-packages.csv' -CatalogName 'Business applications' -FallbackApproverIds '11111111-1111-1111-1111-111111111111' -DefaultAppOwnerIds '22222222-2222-2222-2222-222222222222'
#>
[CmdletBinding(SupportsShouldProcess)]
param(
	[Parameter(Mandatory)]
	[ValidateNotNullOrEmpty()]
	[string]$CsvPath,
	[Parameter(Mandatory)]
	[ValidateNotNullOrEmpty()]
	[string]$CatalogName,
	[Parameter(Mandatory)]
	[ValidateNotNullOrEmpty()]
	[guid[]]$FallbackApproverIds,
	[guid[]]$DefaultAppOwnerIds = @()
)

$ErrorActionPreference = 'Stop'
$CatalogName = $CatalogName.Trim()
if (-not $CatalogName) {
	throw 'CatalogName must not be whitespace.'
}
if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
	throw "CSV file '$CsvPath' was not found."
}
$rows = @(Import-Csv -LiteralPath $CsvPath)
if ($rows.Count -eq 0) {
	throw 'The CSV file contains no access packages.'
}
foreach ($column in 'Name', 'Description', 'GroupId', 'AppOwnerIds') {
	if (-not $rows[0].PSObject.Properties[$column]) {
		throw "The CSV file must contain a '$column' column."
	}
}

Import-Module Microsoft.Graph.Authentication
$context = Get-MgContext
if (-not $context) {
	throw 'Connect to Microsoft Graph first.'
}
if ($context.AuthType -eq 'Delegated') {
	$permissionAlternatives = @(
		@{ Purpose = 'Create catalogs, access packages and policies'; Scopes = @('EntitlementManagement.ReadWrite.All') },
		@{ Purpose = 'Read groups and count their members'; Scopes = @('Group.Read.All', 'Group.ReadWrite.All', 'Directory.Read.All', 'Directory.ReadWrite.All') },
		@{ Purpose = 'Validate approver users'; Scopes = @('User.ReadBasic.All', 'User.Read.All', 'User.ReadWrite.All', 'Directory.Read.All', 'Directory.ReadWrite.All') }
	)
	$missingPermissions = @(foreach ($requirement in $permissionAlternatives) {
		if (-not @($requirement.Scopes | Where-Object { $context.Scopes -contains $_ }).Count) {
			"  - $($requirement.Purpose): $($requirement.Scopes -join ' OR ')"
		}
	})
	if ($missingPermissions.Count) {
		$currentScopes = if ($context.Scopes) { $context.Scopes -join ', ' } else { '(none)' }
		throw @"
Microsoft Graph permission check failed. The current Get-MgContext is missing:
$($missingPermissions -join [Environment]::NewLine)

Current scopes: $currentScopes

Reconnect with the required delegated permissions:
Disconnect-MgGraph
Connect-MgGraph -Scopes 'EntitlementManagement.ReadWrite.All','Group.Read.All','User.ReadBasic.All'

Obtain administrator consent if prompted, then rerun the script.
These scopes do not grant Entra roles. An appropriate active role is also required; a Graph 403 can still indicate missing role permissions.
"@
	}
}

$baseUri = '/v1.0/identityGovernance/entitlementManagement'
$manualAction = 'In the Entra portal, enable "Bypass approval stage if manager is the requester and approver" for this policy.'
$validatedUsers = @{}

function Invoke-PackageGraphRequest {
	param(
		[Parameter(Mandatory)][string]$Uri,
		[Parameter(Mandatory)][ValidateSet('GET', 'POST', 'DELETE')][string]$Method,
		[string]$Body,
		[string]$ContentType,
		[hashtable]$Headers,
		[string]$OutputType = 'Hashtable'
	)

	try {
		Invoke-MgGraphRequest @PSBoundParameters -ErrorAction Stop
	}
	catch {
		$failure = $_
		$details = @("$Method $Uri", $failure.Exception.Message)
		if ($failure.ErrorDetails.Message) {
			$details += $failure.ErrorDetails.Message
		}
		$response = $failure.Exception.Response
		if ($response -is [System.Net.Http.HttpResponseMessage]) {
			$details += "HTTP $([int]$response.StatusCode) ($($response.ReasonPhrase))"
			if ($response.Headers.Contains('request-id')) {
				$details += "Request ID: $($response.Headers.GetValues('request-id') -join ', ')"
			}
			if (-not $failure.ErrorDetails.Message -and $response.Content) {
				try {
					$responseBody = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
					if ($responseBody) { $details += $responseBody }
				}
				catch {
					$details += "Unable to read Graph error response body: $($_.Exception.Message)"
				}
			}
		}
		throw [System.InvalidOperationException]::new(($details -join [Environment]::NewLine), $failure.Exception)
	}
}

function Get-GraphCollection {
	param(
		[Parameter(Mandatory)][string]$Uri,
		[string]$Filter,
		[string]$Expand
	)

	$query = @()
	if ($Filter) { $query += '$filter=' + [uri]::EscapeDataString($Filter) }
	if ($Expand) { $query += '$expand=' + [uri]::EscapeDataString($Expand) }
	if ($query.Count) { $Uri += '?' + ($query -join '&') }
	do {
		$page = Invoke-PackageGraphRequest -Method GET -Uri $Uri -OutputType Hashtable
		if (-not $page.ContainsKey('value')) {
			throw "Microsoft Graph returned no collection for '$Uri'."
		}
		$page.value
		$Uri = $page.'@odata.nextLink'
	} while ($Uri)
}

function Send-GraphBody {
	param(
		[Parameter(Mandatory)][string]$Uri,
		[Parameter(Mandatory)][hashtable]$Body
	)

	Invoke-PackageGraphRequest -Method POST -Uri $Uri -Body ($Body | ConvertTo-Json -Depth 20) -ContentType 'application/json' -OutputType Hashtable
}

function ConvertTo-UserApprovers {
	param([Parameter(Mandatory)][guid[]]$Ids)

	foreach ($id in @($Ids | Select-Object -Unique)) {
		if ($id -eq [guid]::Empty) {
			throw 'Approver IDs must not be the empty GUID.'
		}
		$key = "$id"
		if (-not $validatedUsers.ContainsKey($key)) {
			$user = Invoke-PackageGraphRequest -Method GET -Uri "/v1.0/users/$id`?`$select=id,displayName" -OutputType Hashtable
			if ($user.id -ne $key) {
				throw "Unable to validate approver user '$id'."
			}
			$validatedUsers[$key] = $true
		}
		@{ '@odata.type' = '#microsoft.graph.singleUser'; userId = $key }
	}
}

function Get-GroupResource {
	param(
		[Parameter(Mandatory)][string]$CatalogId,
		[Parameter(Mandatory)][guid]$GroupId
	)

	$resources = @(Get-GraphCollection -Uri "$baseUri/catalogs/$CatalogId/resources" `
		-Filter "originSystem eq 'AadGroup' and originId eq '$GroupId'" -Expand 'roles,scopes')
	if ($resources.Count -gt 1) {
		throw "Multiple catalog resources match group '$GroupId'."
	}
	if ($resources.Count -eq 1) {
		$resource = $resources[0]
		$resource.roles = @(Get-GraphCollection -Uri "$baseUri/catalogs/$CatalogId/resourceRoles" `
			-Filter "originSystem eq 'AadGroup' and resource/id eq '$($resource.id)'")
		$resource
	}
}

function Get-MemberRoleScope {
	param([Parameter(Mandatory)][hashtable]$Resource)

	# Match the stable origin ID, not a localized display name or the Owner role.
	$roles = @($Resource.roles | Where-Object {
		$_.originSystem -eq 'AadGroup' -and $_.originId -eq "Member_$($Resource.originId)"
	})
	$scopes = @($Resource.scopes | Where-Object { $_.isRootScope -eq $true -and $_.originSystem -eq 'AadGroup' })
	if ($roles.Count -gt 1 -or $scopes.Count -gt 1) {
		throw "Ambiguous Member role or root scope for group '$($Resource.originId)'."
	}
	if ($roles.Count -eq 1 -and $scopes.Count -eq 1) {
		@{
			role = @{
				id = $roles[0].id
				originId = $roles[0].originId
				originSystem = $roles[0].originSystem
				resource = @{ id = $Resource.id }
			}
			scope = @{
				id = $scopes[0].id
				originId = $scopes[0].originId
				originSystem = $scopes[0].originSystem
				isRootScope = $true
			}
		}
	}
}

$fallbackApprovers = @(ConvertTo-UserApprovers -Ids $FallbackApproverIds)
$catalogs = @(Get-GraphCollection -Uri "$baseUri/catalogs" -Filter "displayName eq '$($CatalogName.Replace("'", "''"))'")
if ($catalogs.Count -gt 1) {
	throw "Multiple catalogs have the name '$CatalogName'. Use a unique catalog name."
}
$catalog = if ($catalogs.Count -eq 1) { $catalogs[0] } else { $null }
if ($catalog -and $catalog.state -ne 'published') {
	throw "Catalog '$CatalogName' is not published. Publish it before creating requestable access packages."
}

Write-Warning "Manual step required for every new policy: $manualAction"
$results = [System.Collections.Generic.List[object]]::new()
$processedNames = @{}

foreach ($row in $rows) {
	$result = [ordered]@{
		Name = "$($row.Name)".Trim()
		GroupId = "$($row.GroupId)".Trim()
		CatalogId = if ($catalog) { $catalog.id } else { $null }
		AccessPackageId = $null
		PolicyId = $null
		UserCount = $null
		ApprovalStages = $null
		Status = $null
		Reason = $null
		ManualAction = $null
	}
	$newPackageId = $null
	try {
		$name = $result.Name
		$description = "$($row.Description)".Trim()
		if (-not $name -or -not $description) {
			throw 'Name and Description must both contain business-friendly text.'
		}
		if ($name -match '(?i)(?<![a-z0-9])(?:access[\s_-]*package|ap|sg)(?![a-z0-9])') {
			throw "Name '$name' must not contain 'Access Package', 'AP', or 'SG'."
		}
		$groupId = [guid]::Empty
		if (-not [guid]::TryParse($result.GroupId, [ref]$groupId) -or $groupId -eq [guid]::Empty) {
			throw "GroupId '$($result.GroupId)' must be a nonempty GUID."
		}
		$result.GroupId = "$groupId"

		if ($processedNames.ContainsKey($name)) {
			$result.Status = 'Skipped'
			$result.Reason = 'A package with this name was already processed in this CSV.'
		}
		elseif ($catalog) {
			$existing = @(Get-GraphCollection -Uri "$baseUri/accessPackages" `
				-Filter "catalog/id eq '$($catalog.id)' and displayName eq '$($name.Replace("'", "''"))'")
			if ($existing.Count) {
				$result.Status = 'Skipped'
				$result.Reason = "$($existing.Count) package(s) with this name already exist in the catalog; no changes made."
				if ($existing.Count -eq 1) { $result.AccessPackageId = $existing[0].id }
			}
		}

		if (-not $result.Status) {
			$group = Invoke-PackageGraphRequest -Method GET `
				-Uri "/v1.0/groups/$groupId`?`$select=id,displayName,groupTypes,onPremisesSyncEnabled,mailEnabled,securityEnabled,isAssignableToRole" -OutputType Hashtable
			if ($group.id -ne "$groupId") {
				throw "Unable to validate group '$groupId'."
			}
			if ($group.onPremisesSyncEnabled -eq $true -or $group.groupTypes -contains 'DynamicMembership') {
				throw "Group '$($group.displayName)' has synchronized or dynamic membership; entitlement management cannot manage its members."
			}
			if ($group.groupTypes -notcontains 'Unified' -and ($group.mailEnabled -eq $true -or $group.securityEnabled -ne $true)) {
				throw "Group '$($group.displayName)' must be a cloud-managed security group or Microsoft 365 group."
			}
			if ($group.isAssignableToRole -eq $true) {
				throw "Group '$($group.displayName)' is role-assignable; this script does not provision privileged group access."
			}

			$countResponse = Invoke-PackageGraphRequest -Method GET `
				-Uri "/v1.0/groups/$groupId/transitiveMembers/microsoft.graph.user/`$count" `
				-Headers @{ ConsistencyLevel = 'eventual' } -OutputType PSObject
			$userCount = 0
			if (-not [int]::TryParse("$countResponse", [ref]$userCount) -or $userCount -lt 0) {
				throw "Graph returned an invalid transitive user count for group '$groupId'."
			}
			$result.UserCount = $userCount
			$result.ApprovalStages = if ($userCount -gt 500) { 2 } else { 1 }
			$ownerIds = @($DefaultAppOwnerIds)
			if (-not [string]::IsNullOrWhiteSpace($row.AppOwnerIds)) {
				$ownerIds = @(foreach ($value in $row.AppOwnerIds.Split(';')) {
					$ownerId = [guid]::Empty
					if (-not [guid]::TryParse($value.Trim(), [ref]$ownerId) -or $ownerId -eq [guid]::Empty) {
						throw "AppOwnerIds contains an invalid user GUID: '$value'."
					}
					$ownerId
				})
			}
			$ownerApprovers = if ($ownerIds.Count) { @(ConvertTo-UserApprovers -Ids $ownerIds) } else { @() }
			if ($userCount -gt 500 -and -not $ownerApprovers.Count) {
				throw 'Groups with more than 500 users require AppOwnerIds or DefaultAppOwnerIds.'
			}

			$stages = @(
				@{
					durationBeforeAutomaticDenial = 'P14D'
					isApproverJustificationRequired = $false
					isEscalationEnabled = $false
					durationBeforeEscalation = 'PT0S'
					escalationApprovers = @()
					fallbackEscalationApprovers = @()
					primaryApprovers = @(@{ '@odata.type' = '#microsoft.graph.targetManager'; managerLevel = 1 })
					fallbackPrimaryApprovers = $fallbackApprovers
				}
			)
			if ($userCount -gt 500) {
				$stages += @{
					durationBeforeAutomaticDenial = 'P14D'
					isApproverJustificationRequired = $false
					isEscalationEnabled = $false
					durationBeforeEscalation = 'PT0S'
					escalationApprovers = @()
					fallbackEscalationApprovers = @()
					primaryApprovers = @($ownerApprovers)
					fallbackPrimaryApprovers = @()
				}
			}

			$resource = if ($catalog) { Get-GroupResource -CatalogId $catalog.id -GroupId $groupId } else { $null }
			$roleScope = if ($resource) { Get-MemberRoleScope -Resource $resource } else { $null }
			if ($resource -and -not $roleScope) {
				throw "Existing catalog resource for '$groupId' has no Member role/root scope. Check its provisioning state."
			}
			if ($PSCmdlet.ShouldProcess(
				"$name (group $groupId, catalog '$CatalogName')",
				"Create missing catalog/resource, access package and $($result.ApprovalStages)-stage policy; manager bypass requires a manual portal step"
			)) {
				if (-not $catalog) {
					$catalog = Send-GraphBody -Uri "$baseUri/catalogs" -Body @{
						displayName = $CatalogName
						description = "Business application and service access for $CatalogName."
						state = 'published'
						isExternallyVisible = $true
					}
					if (-not $catalog.id) { throw 'Graph did not return an ID for the new catalog. Check the portal before rerunning.' }
				}
				$result.CatalogId = $catalog.id
				if (-not $resource) {
					$null = Send-GraphBody -Uri "$baseUri/resourceRequests" -Body @{
						requestType = 'adminAdd'
						resource = @{ originId = "$groupId"; originSystem = 'AadGroup'; displayName = $group.displayName }
						catalog = @{ id = $catalog.id }
					}
					# Resource import can complete before its role and scope become readable.
					for ($attempt = 0; $attempt -lt 6; $attempt++) {
						if ($attempt -gt 0) { Start-Sleep -Seconds 2 }
						$resource = Get-GroupResource -CatalogId $catalog.id -GroupId $groupId
						$roleScope = if ($resource) { Get-MemberRoleScope -Resource $resource } else { $null }
						if ($roleScope) { break }
					}
					if (-not $roleScope) {
						throw "The catalog resource for group '$groupId' did not expose a Member role/root scope within the retry window. Check provisioning before rerunning."
					}
				}
				$package = Send-GraphBody -Uri "$baseUri/accessPackages" -Body @{
					displayName = $name
					description = $description
					isHidden = $false
					catalog = @{ id = $catalog.id }
				}
				if (-not $package.id) { throw 'Graph did not return an ID for the new package. Check the portal before rerunning.' }
				$newPackageId = $package.id
				$result.AccessPackageId = $newPackageId
				$null = Send-GraphBody -Uri "$baseUri/accessPackages/$newPackageId/resourceRoleScopes" -Body $roleScope
				$policy = Send-GraphBody -Uri "$baseUri/assignmentPolicies" -Body @{
					displayName = "$name - Requests"
					description = 'All directory members excluding guests; self, administrator and manager requests.'
					accessPackage = @{ id = $newPackageId }
					allowedTargetScope = 'allMemberUsers'
					specificAllowedTargets = @()
					questions = @()
					expiration = @{ type = 'afterDuration'; duration = 'P365D' }
					requestorSettings = @{
						enableTargetsToSelfAddAccess = $true
						enableTargetsToSelfUpdateAccess = $true
						enableTargetsToSelfRemoveAccess = $true
						allowCustomAssignmentSchedule = $true
						enableOnBehalfRequestorsToAddAccess = $true
						enableOnBehalfRequestorsToUpdateAccess = $true
						enableOnBehalfRequestorsToRemoveAccess = $false
						onBehalfRequestors = @(@{ '@odata.type' = '#microsoft.graph.targetManager'; managerLevel = 1 })
					}
					requestApprovalSettings = @{
						isApprovalRequiredForAdd = $true
						isApprovalRequiredForUpdate = $true
						isRequestorJustificationRequired = $true
						stages = $stages
					}
					notificationSettings = @{ isAssignmentNotificationDisabled = $false }
					reviewSettings = $null
				}
				if (-not $policy.id) { throw 'Graph did not return an ID for the new policy.' }
				$result.PolicyId = $policy.id
				$result.Status = 'Created'
			}
			else {
				$result.Status = if ($WhatIfPreference) { 'Would be created' } else { 'Not created' }
			}
			if ($result.Status -in 'Created', 'Would be created') {
				$result.ManualAction = $manualAction
				$processedNames[$name] = $true
			}
		}
	}
	catch {
		$result.Status = 'Failed'
		$result.Reason = $_.Exception.Message
		if ($newPackageId) {
			try {
				$null = Invoke-PackageGraphRequest -Method DELETE -Uri "$baseUri/accessPackages/$newPackageId"
				$result.AccessPackageId = $null
				$result.Reason += ' The newly created package was removed; catalog resources were retained.'
			}
			catch {
				$result.Reason += " ROLLBACK FAILED for package '$newPackageId': $($_.Exception.Message). Inspect and remove the incomplete package manually."
			}
		}
		Write-Warning "Failed '$($result.Name)': $($result.Reason)"
	}

	$resultObject = [pscustomobject]$result
	$results.Add($resultObject)
	$resultObject
}

$notCreated = @($results | Where-Object { $_.Status -in 'Skipped', 'Failed' })
if ($notCreated.Count) {
	Write-Host ($notCreated | Select-Object Name, Status, Reason | Format-Table -AutoSize -Wrap | Out-String -Width 200)
}
$createdStatus = if ($WhatIfPreference) { 'Would be created' } else { 'Created' }
$createdCount = @($results | Where-Object { $_.Status -eq $createdStatus }).Count
$skippedCount = @($results | Where-Object { $_.Status -eq 'Skipped' }).Count
$failedCount = @($results | Where-Object { $_.Status -eq 'Failed' }).Count
$declinedCount = @($results | Where-Object { $_.Status -eq 'Not created' }).Count
Write-Host "$($results.Count) package(s): $createdCount $($createdStatus.ToLower()), $skippedCount skipped, $failedCount failed, $declinedCount not created."
