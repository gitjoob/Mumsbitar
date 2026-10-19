# Requires: Microsoft.Graph PowerShell module
Connect-MgGraph -Scopes "AccessReview.Read.All","EntitlementManagement.Read.All"

$accessPackages = Get-MgEntitlementManagementAccessPackage

$accessPolicys = Get-MgEntitlementManagementAssignmentPolicy -Property AccessPackage,Id,DisplayName,ReviewSettings | Select-Object AccessPackage,Id,DisplayName,ReviewSettings -ExpandProperty AccessPackage

$accessPolicysPackage = Get-MgEntitlementManagementAssignmentPolicy -Property AccessPackage | Select-Object AccessPackage

