function Get-PackageGroup {
	param([guid]$GroupId)

	$group = Get-MgGroup -GroupId "$GroupId" -Property id, displayName, groupTypes, onPremisesSyncEnabled, mailEnabled, securityEnabled, isAssignableToRole -ErrorAction Stop
	if ($group.id -ne "$GroupId") {
		throw "Unable to validate group '$GroupId'."
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
	$group
}
