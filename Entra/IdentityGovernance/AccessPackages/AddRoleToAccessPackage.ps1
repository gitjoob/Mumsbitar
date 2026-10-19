Connect-MgGraph -Scopes "EntitlementManagement.ReadWrite.All"

$accessPackageId = ""
$catalog = Get-MgEntitlementManagementCatalog -Filter "displayName eq 'Roles'" -All
if ($catalog -eq $null) { throw "catalog not found" }

# Find the Global Reader directory role resource
$resource = Get-MgEntitlementManagementCatalogResource -AccessPackageCatalogId $catalog.id `
    -Filter "originSystem eq 'DirectoryRole' and displayName eq 'Global Reader'" -ExpandProperty roles,scopes
if ($resource -eq $null) { throw "Global Reader resource not found" }

# Find the Member role definition
$role = $resource.roles | Where-Object { $_.displayName -eq "Member" }
if ($role -eq $null) { throw "Member role definition not found" }

# Find the default scope (should be the tenant)
$scope = $resource.scopes | Where-Object { $_.isRootScope -eq $true }
if ($scope -eq $null) { throw "Scope not found" }

# Add the Member role to the access package
$params = @{
    AccessPackageId = $accessPackageId
    AccessPackageResourceRoleId = $role.id
    AccessPackageResourceScopeId = $scope.id
}

New-MgEntitlementManagementAccessPackageResourceRoleScope @params

Write-Host "Member role added to access package $accessPackageId."