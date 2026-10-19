<#
.SYNOPSIS
    Automated RBAC Group Creation and Access Package Management for Azure Subscriptions

.DESCRIPTION
    This script automates the creation and management of Azure RBAC groups with integrated 
    Entra ID access packages for subscription-level role assignments.
    
    📋 Main Functions:
    1. 🔗 Connects to Azure and Microsoft Graph with required permissions
    2. 👥 Creates RBAC security groups for Reader, Contributor, and Owner roles
    3. 🔐 Enables PIM (Privileged Identity Management) for Owner groups (all environments) 
       and Contributor groups (prod/unknown environments)
    4. 🎯 Optionally assigns Azure RBAC roles at subscription level (controlled by -RBAC switch)
    5. 🛡️ Applies ABAC (Attribute-Based Access Control) conditions to Owner role assignments
       to prevent assignment/removal of privileged roles
    6. 📦 Optionally creates and configures Entra ID access packages (controlled by -AccessPackages switch):
       - Reader: Member role, 1/3/6 month policies
       - Contributor Dev/Stg: Member role, 1/3/6 month policies
       - Contributor Prod/Unknown: Eligible Member role (PIM), 1/3/6 month policies
       - Owner: Eligible Member role (PIM), 8-hour policy
    7. 📊 Automatic catalog resource management with role availability checking
    
    🔐 Security Features:
    - PIM-enabled groups for privileged access (Owner all environments, Contributor prod only)
    - ABAC conditions on Owner roles preventing assignment/deletion of:
      * Owner (8e3af657-a8ff-443c-a75c-2fe8c4bcb635)
      * Contributor (b24988ac-6180-42a0-ab88-20f7382dd24c)
      * User Access Administrator (18d7d88d-d35e-4fb5-a5c3-7773c20a72d9)
      * [Role GUID: f58310d9-a9f6-439a-9e8d-f62e7b41a168]
    - Automatic detection and update of existing Owner assignments lacking ABAC conditions
    
    📝 Naming Conventions:
    - Groups: grp-rbac-{subscription-name}-{Role}
    - Access Packages:
      * "Subscription - {ServiceName} Reader"
      * "Subscription - {ServiceName} Contributor Dev/Stg" (dev/stg only)
      * "Subscription - {ServiceName} Contributor Prod" (prod only)
      * "Subscription - {ServiceName} Owner"
    
    🎯 Environment Detection:
    The script automatically detects environment (dev/stg/prod/unknown) from subscription name:
    - Example: "sub-landingzone-dev-02" → environment = "dev"
    - Example: "sub-management" → environment = "unknown" (CAF/platform subscriptions)
    - Determines PIM enablement and access package creation accordingly:
      * dev/stg: Contributor as Member role
      * prod/unknown: Contributor as Eligible Member role (PIM)
      * all: Owner as Eligible Member role (PIM)
    
    ⚙️ Required Permissions:
    Microsoft Graph API:
    - Group.ReadWrite.All (create/manage security groups)
    - PrivilegedAccess.ReadWrite.AzureAD (PIM registration)
    - EntitlementManagement.ReadWrite.All (access packages and policies)
    
    Entra ID Directory Roles (at least one required):
    - Identity Governance Administrator (for access package management)
    - Privileged Role Administrator (for PIM operations)
    - Groups Administrator (for group management)
    - Global Administrator (all operations)
    
    Azure RBAC:
    - Owner or User Access Administrator on target subscription (to assign roles)
    
    Catalog Access:
    - Owner or contributor access to Entitlement Management catalogs:
      * Subscriptions catalog (d9b08eb9-b613-439e-8ea2-ba65fcafd51d)
      * Subscriptions Owner catalog (1cf97038-6532-4ca0-aac4-3cdc63982afa)
    
    📦 Required Modules:
    - Az.Accounts (Azure authentication and subscription management)
    - Az.Resources (Azure RBAC role assignments)
    - Microsoft.Graph.Authentication (Microsoft Graph connection)
    - Microsoft.Graph.Groups (Group operations)
    - Microsoft.Graph.Users (User operations)
    - Microsoft.Graph.Identity.Governance (Entitlement Management)
    
    🔄 Workflow:
    1. Script validates all required permissions
    2. Creates three security groups (Reader, Contributor, Owner)
    3. Registers Contributor (prod/unknown) and Owner groups for PIM
    4. Optionally assigns Azure RBAC role assignments with ABAC conditions for Owner (-RBAC switch)
    5. Optionally creates/updates access packages with environment-appropriate policies (-AccessPackages switch)
    6. Adds groups to access packages with correct role type (Member vs Eligible Member)
    
    ⚠️ Important Notes:
    - ABAC conditions are applied only to subscription-level Owner role assignments
    - PIM registration is idempotent and safe to run multiple times
    - Use -RBAC:$false to skip role assignment (groups only)
    - Use -AccessPackages:$false to skip access package creation (groups + RBAC only)
    - Access package policies include manager approval requirements
    - Catalog role propagation may take a few seconds after adding PIM groups

.PARAMETER Subscription
    The name of the Azure subscription.
    Example: "sub-landingzone-prod-02", "sub-management", "sub-security"

.PARAMETER RBAC
    Switch to control Azure RBAC role assignment.
    Default: $true (assigns roles)
    Use -RBAC:$false to skip role assignments and only create groups

.PARAMETER AccessPackages
    Switch to control Entra ID access package creation.
    Default: $true (creates access packages)
    Use -AccessPackages:$false to skip access package operations

.EXAMPLE
    .\Create-rbac-Groups.ps1 -Subscription "sub-management"
    
    Creates groups, assigns RBAC roles, and creates access packages for sub-management.

.EXAMPLE
    .\Create-rbac-Groups.ps1 -Subscription "sub-landingzone-prod-02" -RBAC:$false
    
    Creates groups and access packages but skips Azure RBAC role assignments.

.EXAMPLE
    .\Create-rbac-Groups.ps1 -Subscription "sub-security" -AccessPackages:$false
    
    Creates groups and assigns RBAC roles but skips access package creation.

.EXAMPLE
    .\Create-rbac-Groups.ps1 -Subscription "sub-identity" -RBAC:$false -AccessPackages:$false
    
    Creates groups only without role assignments or access packages.

.NOTES
    Created by: Johan Öberg
    Last Updated: January 2026
    Version: 4.0
    Tenant ID: bcee8bc2-b710-439d-b5bd-f5b83846ddee
    
    Change Log:
    - v4.0: Added parameter support, CAF subscription support, RBAC/AccessPackages switches
    - v3.0: Added ABAC conditions for Owner role assignments
    - v2.0: Added PIM registration and access package management
    - v1.0: Initial version with basic group creation
#>

param(
    [Parameter(Mandatory=$true, HelpMessage="The name of the Azure subscription (e.g., sub-management, sub-landingzone-prod-02)")]
    [string]$Subscription,
    
    [Parameter(Mandatory=$false, HelpMessage="Switch to control Azure RBAC role assignment. Default: true")]
    [bool]$RBAC = $true,
    
    [Parameter(Mandatory=$false, HelpMessage="Switch to control Entra ID access package creation. Default: true")]
    [bool]$AccessPackages = $true
)

$subname = $Subscription
$tenantId = "bcee8bc2-b710-439d-b5bd-f5b83846ddee"

# Check if Az.Accounts module is available
if (-not (Get-Module -ListAvailable -Name Az.Accounts)) {
    Write-Error "Az.Accounts module is not installed. Please install it using: Install-Module Az.Accounts"
    exit
}

# Check if Microsoft.Graph module is available
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Write-Error "Microsoft.Graph.Users module is not installed. Please install it using: Install-Module Microsoft.Graph.Users"
    exit
}

# Connect to Azure with specific tenant
try {
    $azContext = Get-AzContext
    if (-not $azContext -or $azContext.Tenant.Id -ne $tenantId) {
        Write-Host "Connecting to Azure (Tenant: $tenantId)..." -ForegroundColor Yellow
        Connect-AzAccount -TenantId $tenantId -ErrorAction Stop
        Write-Host "Connected to Azure as: $((Get-AzContext).Account.Id)" -ForegroundColor Green
    } else {
        Write-Host "Already connected to Azure as: $($azContext.Account.Id)" -ForegroundColor Green
    }
} catch {
    Write-Error "Failed to connect to Azure: $_"
    exit
}

# Validate subscription exists
Write-Host "`nValidating subscription..." -ForegroundColor Cyan
$validatedSubscription = Get-AzSubscription -SubscriptionName $subname -ErrorAction SilentlyContinue

if (-not $validatedSubscription) {
    Write-Error "Subscription '$subname' not found. Please check the subscription name and ensure you have access to it."
    Write-Host "`nAvailable subscriptions:" -ForegroundColor Yellow
    Get-AzSubscription | Format-Table Name, Id, State -AutoSize
    exit
}

Write-Host "✓ Subscription validated: $($validatedSubscription.Name) ($($validatedSubscription.Id))" -ForegroundColor Green

# Connect to Microsoft Graph if not already connected
try {
    $context = Get-MgContext
    if (-not $context) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes "Group.ReadWrite.All","PrivilegedAccess.ReadWrite.AzureAD","EntitlementManagement.ReadWrite.All" -TenantId $tenantId -ErrorAction Stop
    }
} catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit
}

# Catalog IDs
$readerContributorCatalogId = "d9b08eb9-b613-439e-8ea2-ba65fcafd51d"
$ownerCatalogId = "1cf97038-6532-4ca0-aac4-3cdc63982afa"
$fallbackReviewersGroupId = "fa3d81d6-b7f1-4068-bc11-ec5972f86859"

# Check if user has required permissions for Entitlement Management
Write-Host "`nChecking Entitlement Management permissions..." -ForegroundColor Cyan

$currentUser = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me"
$userId = $currentUser.id

# Check for Identity Governance Administrator role
$identityGovernanceAdminRole = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=displayName eq 'Identity Governance Administrator'"

$hasIdentityGovernanceRole = $false
if ($identityGovernanceAdminRole.value.Count -gt 0) {
    $roleMembers = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($identityGovernanceAdminRole.value[0].id)/members"
    $hasIdentityGovernanceRole = $roleMembers.value | Where-Object { $_.id -eq $userId }
}

# Check for Global Administrator role
$globalAdminRole = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=displayName eq 'Global Administrator'"

$hasGlobalAdminRole = $false
if ($globalAdminRole.value.Count -gt 0) {
    $roleMembers = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($globalAdminRole.value[0].id)/members"
    $hasGlobalAdminRole = $roleMembers.value | Where-Object { $_.id -eq $userId }
}

# Check for Privileged Role Administrator role
$privilegedRoleAdminRole = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=displayName eq 'Privileged Role Administrator'"

$hasPrivilegedRoleAdminRole = $false
if ($privilegedRoleAdminRole.value.Count -gt 0) {
    $roleMembers = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($privilegedRoleAdminRole.value[0].id)/members"
    $hasPrivilegedRoleAdminRole = $roleMembers.value | Where-Object { $_.id -eq $userId }
}

# Check for Groups Administrator role
$groupsAdminRole = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=displayName eq 'Groups Administrator'"

$hasGroupsAdminRole = $false
if ($groupsAdminRole.value.Count -gt 0) {
    $roleMembers = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($groupsAdminRole.value[0].id)/members"
    $hasGroupsAdminRole = $roleMembers.value | Where-Object { $_.id -eq $userId }
}

# Check catalog ownership for both catalogs
function Test-CatalogOwnership {
    param([string]$CatalogId, [string]$UserId)
    
    try {
        $catalogOwners = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$CatalogId/accessPackageResourceRoles?`$filter=originSystem eq 'AadGroup' or originSystem eq 'DirectoryUser'&`$expand=accessPackageResource"
        
        # Check if user is a catalog owner or resource owner
        foreach ($owner in $catalogOwners.value) {
            if ($owner.accessPackageResource.originId -eq $UserId) {
                return $true
            }
        }
        
        # Also check catalog-level permissions
        $catalog = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$CatalogId"
        return $catalog.isExternallyVisible -eq $false # Internal catalogs may allow more permissions
    } catch {
        return $false
    }
}

$hasReaderContributorCatalogAccess = Test-CatalogOwnership -CatalogId $readerContributorCatalogId -UserId $userId
$hasOwnerCatalogAccess = Test-CatalogOwnership -CatalogId $ownerCatalogId -UserId $userId

# Display permission status
Write-Host "`nPermission Check Results:" -ForegroundColor Yellow
Write-Host "  Current User: $($currentUser.userPrincipalName)" -ForegroundColor White
Write-Host "  Identity Governance Administrator: $(if($hasIdentityGovernanceRole){'✓ Yes'}else{'✗ No'})" -ForegroundColor $(if($hasIdentityGovernanceRole){'Green'}else{'Red'})
Write-Host "  Privileged Role Administrator: $(if($hasPrivilegedRoleAdminRole){'✓ Yes'}else{'✗ No'})" -ForegroundColor $(if($hasPrivilegedRoleAdminRole){'Green'}else{'Red'})
Write-Host "  Groups Administrator: $(if($hasGroupsAdminRole){'✓ Yes'}else{'✗ No'})" -ForegroundColor $(if($hasGroupsAdminRole){'Green'}else{'Red'})
Write-Host "  Global Administrator: $(if($hasGlobalAdminRole){'✓ Yes'}else{'✗ No'})" -ForegroundColor $(if($hasGlobalAdminRole){'Green'}else{'Red'})
Write-Host "  Subscriptions Catalog Access: $(if($hasReaderContributorCatalogAccess){'✓ Yes'}else{'✗ No'})" -ForegroundColor $(if($hasReaderContributorCatalogAccess){'Green'}else{'Red'})
Write-Host "  Subscriptions Owner Catalog Access: $(if($hasOwnerCatalogAccess){'✓ Yes'}else{'✗ No'})" -ForegroundColor $(if($hasOwnerCatalogAccess){'Green'}else{'Red'})

$hasRequiredPermissions = $hasIdentityGovernanceRole -or $hasGlobalAdminRole -or $hasPrivilegedRoleAdminRole -or $hasGroupsAdminRole -or $hasReaderContributorCatalogAccess -or $hasOwnerCatalogAccess

if (-not $hasRequiredPermissions) {
    Write-Warning "`nYou do not have sufficient permissions for Entitlement Management operations."
    Write-Warning "Required: Identity Governance Administrator OR Privileged Role Administrator OR Groups Administrator OR Global Administrator role OR catalog owner permissions."
    Write-Warning "Access package operations will likely fail. Contact your administrator to grant the required permissions.`n"
    
    $continue = Read-Host "Do you want to continue anyway? (Y/N)"
    if ($continue -ne 'Y' -and $continue -ne 'y') {
        Write-Host "Exiting script." -ForegroundColor Yellow
        exit
    }
} else {
    Write-Host "`n✓ You have sufficient permissions to manage access packages.`n" -ForegroundColor Green
    
    # Additional note about group ownership
    if ($hasGroupsAdminRole -or $hasGlobalAdminRole) {
        Write-Host "✓ You have administrative rights to manage all groups.`n" -ForegroundColor Green
    } else {
        Write-Host "Note: To add groups to catalogs, you may need Groups Administrator or be an owner of specific groups." -ForegroundColor Yellow
        Write-Host "The 'CallerNotResourceOwner' error occurs when you lack group ownership/management rights.`n" -ForegroundColor Yellow
    }
}

# Function to check and apply role assignment
function Set-RoleAssignment {
    param(
        [string]$GroupName,
        [string]$Role,
        [string]$SubscriptionName
    )
    
    # ABAC condition for Owner role to prevent assigning/removing privileged roles
    $ownerCondition = @"
(
 (
  !(ActionMatches{'Microsoft.Authorization/roleAssignments/write'})
 )
 OR 
 (
  @Request[Microsoft.Authorization/roleAssignments:RoleDefinitionId] ForAnyOfAllValues:GuidNotEquals {8e3af657-a8ff-443c-a75c-2fe8c4bcb635, b24988ac-6180-42a0-ab88-20f7382dd24c, f58310d9-a9f6-439a-9e8d-f62e7b41a168, 18d7d88d-d35e-4fb5-a5c3-7773c20a72d9}
 )
)
AND
(
 (
  !(ActionMatches{'Microsoft.Authorization/roleAssignments/delete'})
 )
 OR 
 (
  @Resource[Microsoft.Authorization/roleAssignments:RoleDefinitionId] ForAnyOfAllValues:GuidNotEquals {8e3af657-a8ff-443c-a75c-2fe8c4bcb635, b24988ac-6180-42a0-ab88-20f7382dd24c, f58310d9-a9f6-439a-9e8d-f62e7b41a168, 18d7d88d-d35e-4fb5-a5c3-7773c20a72d9}
 )
)
"@
    
    $subscription = Get-AzSubscription -SubscriptionName $SubscriptionName -ErrorAction SilentlyContinue
    if ($subscription) {
        $group = Get-MgGroup -Filter "displayName eq '$GroupName'"
        $roleAssignment = Get-AzRoleAssignment -ObjectId $group.Id -Scope "/subscriptions/$($subscription.Id)" -RoleDefinitionName $Role -ErrorAction SilentlyContinue
        
        if (-not $roleAssignment) {
            if ($script:RBAC) {
                if ($Role -eq 'Owner') {
                    # Create Owner assignment with ABAC condition
                    New-AzRoleAssignment -ObjectId $group.Id -RoleDefinitionName $Role -Scope "/subscriptions/$($subscription.Id)" -Condition $ownerCondition -ConditionVersion "2.0"
                    Write-Output "✓ Owner role assignment applied with ABAC conditions for $GroupName on $SubscriptionName"
                } else {
                    # Create assignment without conditions for other roles
                    New-AzRoleAssignment -ObjectId $group.Id -RoleDefinitionName $Role -Scope "/subscriptions/$($subscription.Id)"
                    Write-Output "✓ Role assignment applied: $Role for $GroupName on $SubscriptionName"
                }
            } else {
                Write-Output "⊘ Skipping role assignment for $GroupName (RBAC switch is disabled)"
            }
        } else {
            # Check if existing Owner assignment has conditions
            if ($Role -eq 'Owner' -and [string]::IsNullOrWhiteSpace($roleAssignment.Condition)) {
                if ($script:RBAC) {
                    Write-Host "⚠️  Owner role assignment exists but lacks ABAC conditions for $GroupName" -ForegroundColor Yellow
                    Write-Host "Updating Owner assignment with ABAC conditions..." -ForegroundColor Yellow
                    # Remove old assignment
                    Remove-AzRoleAssignment -ObjectId $group.Id -RoleDefinitionName $Role -Scope "/subscriptions/$($subscription.Id)" | Out-Null
                    # Create new assignment with conditions
                    New-AzRoleAssignment -ObjectId $group.Id -RoleDefinitionName $Role -Scope "/subscriptions/$($subscription.Id)" -Condition $ownerCondition -ConditionVersion "2.0"
                    Write-Output "✓ Owner role assignment updated with ABAC conditions for $GroupName on $SubscriptionName"
                } else {
                    Write-Output "⚠️  Owner assignment exists without ABAC conditions for $GroupName (RBAC switch disabled, skipping update)"
                }
            } else {
                Write-Output "✓ Role assignment already exists for $GroupName"
            }
        }
    }
}

# Function to set or update access package with group
function Set-AccessPackageWithGroup {
    param(
        [string]$AccessPackageName,
        [string]$CatalogId,
        [string]$GroupId,
        [string]$GroupName,
        [string]$PolicyType = "Reader",
        [string]$Environment = "unknown"
    )
    
    # Determine role type based on PolicyType and Environment
    # Eligible Member for: Owner (all), Contributor Prod
    # Member for: Reader (all), Contributor Dev/Stg
    $roleType = if ($PolicyType -eq "Owner" -or ($PolicyType -eq "Contributor" -and $Environment -eq "prod")) {
        "Eligible Member"
    } else {
        "Member"
    }
    
    # Check if access package exists
    $existingPackages = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/accessPackages?`$filter=displayName eq '$AccessPackageName'&`$expand=catalog"
    
    # Filter by catalog ID manually
    $accessPackage = $existingPackages.value | Where-Object { $_.catalog.id -eq $CatalogId } | Select-Object -First 1
    
    if ($accessPackage) {
        Write-Output "Access package already exists: $AccessPackageName"
        
        # Check if group is already a resource in the access package
        $resources = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/$($accessPackage.id)/accessPackageResourceRoleScopes"
        
        $groupAlreadyAdded = $false
        foreach ($resource in $resources.value) {
            if ($resource.accessPackageResourceRole.accessPackageResource.originId -eq $GroupId) {
                $groupAlreadyAdded = $true
                Write-Output "✓ Group $GroupName is already a resource in access package $AccessPackageName"
                break
            }
        }
        
        if (-not $groupAlreadyAdded) {
            # Add group as resource to existing access package
            Add-GroupToAccessPackage -AccessPackageId $accessPackage.id -CatalogId $CatalogId -GroupId $GroupId -GroupName $GroupName -RoleType $roleType
        }
        
        # Check if policies exist for this access package
        $existingPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/assignmentPolicies?`$filter=accessPackage/id eq '$($accessPackage.id)'"
        
        if ($existingPolicies.value.Count -eq 0) {
            Write-Output "No policies found for access package. Creating policies..."
            Create-AccessPackagePolicies -AccessPackageId $accessPackage.id -AccessPackageName $AccessPackageName -PolicyType $PolicyType
        } else {
            Write-Output "✓ Access package already has $($existingPolicies.value.Count) policy/policies configured"
        }
    } else {
        Write-Output "Creating new access package: $AccessPackageName"
        
        # Create new access package
        $newPackageBody = @{
            displayName = $AccessPackageName
            description = "Access package for $AccessPackageName"
            catalog = @{
                id = $CatalogId
            }
            isHidden = $false
        }
        
        try {
            $accessPackage = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/accessPackages" -Body ($newPackageBody | ConvertTo-Json) -ErrorAction Stop
            Write-Output "Access package created: $AccessPackageName"
            
            # Add group as resource
            Add-GroupToAccessPackage -AccessPackageId $accessPackage.id -CatalogId $CatalogId -GroupId $GroupId -GroupName $GroupName -RoleType $roleType
            
            # Create policies based on policy type
            Create-AccessPackagePolicies -AccessPackageId $accessPackage.id -AccessPackageName $AccessPackageName -PolicyType $PolicyType
        } catch {
            Write-Warning "Failed to create access package: $_"
        }
    }
}

# Function to add group to access package
function Add-GroupToAccessPackage {
    param(
        [string]$AccessPackageId,
        [string]$CatalogId,
        [string]$GroupId,
        [string]$GroupName,
        [string]$RoleType = "Member"  # "Member" or "Eligible Member"
    )
    
    # Check if current user is owner of the group
    $groupOwners = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId/owners"
    $isGroupOwner = $groupOwners.value | Where-Object { $_.id -eq $script:userId }
    
    if (-not $isGroupOwner -and -not $script:hasGlobalAdminRole -and -not $script:hasGroupsAdminRole) {
        Write-Warning "Cannot add group $GroupName to catalog: You are not an owner of this group and don't have Groups Administrator role."
        Write-Warning "To add this group to the catalog, you need one of:"
        Write-Warning "  - Be listed as an owner of the group"
        Write-Warning "  - Have Groups Administrator role"
        Write-Warning "  - Have Global Administrator role"
        Write-Warning "Current owners can be checked with: Get-MgGroupOwner -GroupId $GroupId"
        return
    }
    
    # First, add the group to the catalog if not already there
    $catalogResources = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$CatalogId/accessPackageResources?`$filter=originId eq '$GroupId'"
    
    # Check if group is PIM-enabled and we need Eligible Member role
    $needsEligibleMember = $RoleType -eq "Eligible Member"
    $groupInCatalog = $catalogResources.value.Count -gt 0
    
    # If group is already in catalog but we need Eligible Member, check if role is available
    if ($groupInCatalog -and $needsEligibleMember) {
        $resourceWithRoles = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$CatalogId/accessPackageResources?`$filter=originId eq '$GroupId'&`$expand=accessPackageResourceRoles"
        $availableRoles = $resourceWithRoles.value[0].accessPackageResourceRoles
        $hasEligibleMember = $availableRoles | Where-Object { $_.displayName -eq "Eligible Member" }
        
        if (-not $hasEligibleMember) {
            Write-Host "Group is in catalog but Eligible Member role not available yet. Refreshing catalog..." -ForegroundColor Yellow
            # Remove and re-add to refresh available roles
            $resourceId = $catalogResources.value[0].id
            Write-Host "  Removing group from catalog..." -ForegroundColor Gray
            Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$CatalogId/accessPackageResources/$resourceId"
            Start-Sleep -Seconds 5
            $groupInCatalog = $false  # Force re-add
        }
    }
    
    if (-not $groupInCatalog) {
        Write-Output "Adding group $GroupName to catalog..."
        
        $resourceRequestBody = @{
            catalogId = $CatalogId
            requestType = "AdminAdd"
            accessPackageResource = @{
                originId = $GroupId
                originSystem = "AadGroup"
            }
        }
        
        try {
            $response = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageResourceRequests" -Body ($resourceRequestBody | ConvertTo-Json -Depth 10) -ErrorAction Stop
            Write-Output "Group added to catalog"
            
            # Brief wait for catalog to update
            Start-Sleep -Seconds 5
        } catch {
            Write-Warning "Failed to add group to catalog (may already exist): $_"
        }
        
        # Get the catalog resource with retry logic
        $catalogResources = $null
        $retryCount = 0
        $maxRetries = 3
        
        while ($retryCount -lt $maxRetries -and ($null -eq $catalogResources -or $catalogResources.value.Count -eq 0)) {
            Start-Sleep -Seconds 5
            $catalogResources = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$CatalogId/accessPackageResources?`$filter=originId eq '$GroupId'"
            $retryCount++
            if ($null -eq $catalogResources -or $catalogResources.value.Count -eq 0) {
                Write-Host "Resource not yet available, retrying ($retryCount/$maxRetries)..." -ForegroundColor Yellow
            }
        }
    }
    
    if ($catalogResources.value.Count -gt 0) {
        $catalogResource = $catalogResources.value[0]
        
        # Use the requested role type directly
        Write-Host "Adding group to access package as '$RoleType'..." -ForegroundColor Yellow
        
        # Construct the resource role scope
        # For AadGroup resources, the role originId pattern is: "Member_<GroupId>" or "EligibleMember_<GroupId>"
        $roleOriginId = if ($RoleType -eq "Eligible Member") { "EligibleMember_$GroupId" } else { "Member_$GroupId" }
        
        $scopeBody = @{
            accessPackageResourceRole = @{
                displayName = $RoleType
                originId = $roleOriginId
                originSystem = "AadGroup"
                accessPackageResource = @{
                    id = $catalogResource.id
                    originId = $GroupId
                    originSystem = "AadGroup"
                }
            }
            accessPackageResourceScope = @{
                id = $catalogResource.id
                originId = $GroupId
                originSystem = "AadGroup"
            }
        }
        
        try {
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/$AccessPackageId/accessPackageResourceRoleScopes" -Body ($scopeBody | ConvertTo-Json -Depth 10) -ErrorAction Stop
            Write-Output "✓ Group $GroupName added as '$RoleType' to access package"
        } catch {
            Write-Warning "Failed to add group to access package: $_"
        }
    } else {
        Write-Warning "Could not find catalog resource for group $GroupName"
    }
}

# Function to create access package policies
function Create-AccessPackagePolicies {
    param(
        [string]$AccessPackageId,
        [string]$AccessPackageName,
        [string]$PolicyType = "Reader"  # Reader, Contributor, or Owner
    )
    
    # Define durations based on policy type
    if ($PolicyType -eq "Owner") {
        $durations = @(
            @{ Name = "8 Hour Policy"; Duration = "PT8H"; Description = "Assignment expire after 8 hours" }
        )
    } else {
        $durations = @(
            @{ Name = "1 Month Policy"; Duration = "P30D"; Description = "Assignment expire after 30 days with no review" },
            @{ Name = "3 Month Policy"; Duration = "P90D"; Description = "Assignment expire after 90 days" },
            @{ Name = "6 Month Policy"; Duration = "P180D"; Description = "Assignment expire after 180 days" }
        )
    }
    
    foreach ($duration in $durations) {
        Write-Output "Creating policy: $($duration.Name) for $AccessPackageName"
        
        # Disable extension for 1 Month Policy (P30D) and 8 Hour Policy (PT8H)
        $allowExtension = $duration.Duration -notin @("P30D", "PT8H")
        
        $policyBody = @{
            displayName = $duration.Name
            description = $duration.Description
            allowedTargetScope = "specificDirectoryUsers"
            specificAllowedTargets = @(
                @{
                    "@odata.type" = "#microsoft.graph.groupMembers"
                    groupId = "e91ecace-f851-429f-aee1-04e856299833"
                    description = "Allowed Requesters"
                },
                @{
                    "@odata.type" = "#microsoft.graph.groupMembers"
                    groupId = "ef9d0338-76da-4eb8-9246-b3b818d6001c"
                    description = "Allowed Requesters 2"
                }
            )
            expiration = @{
                type = "afterDuration"
                duration = $duration.Duration
            }
            requestorSettings = @{
                enableTargetsToSelfAddAccess = $true
                enableTargetsToSelfUpdateAccess = $allowExtension
                enableTargetsToSelfRemoveAccess = $true
                enableOnBehalfRequestorsToAddAccess = $true
                enableOnBehalfRequestorsToUpdateAccess = $allowExtension
                enableOnBehalfRequestorsToRemoveAccess = $true
                onBehalfRequestors = @()
            }
            requestApprovalSettings = @{
                isApprovalRequiredForAdd = $true
                isApprovalRequiredForUpdate = if ($PolicyType -eq "Owner") { $false } else { $true }
                isRequestorJustificationRequired = $true
                stages = @(
                    @{
                        durationBeforeAutomaticDenial = "P7D"
                        durationBeforeEscalation = "PT0S"
                        isEscalationEnabled = $false
                        isApproverJustificationRequired = $true
                        fallbackPrimaryApprovers = @(
                            @{
                                "@odata.type" = "#microsoft.graph.groupMembers"
                                groupId = $script:fallbackReviewersGroupId
                                description = "Fallback Reviewers"
                            }
                        )
                        fallbackEscalationApprovers = @()
                        escalationApprovers = @()
                        primaryApprovers = @(
                            @{
                                "@odata.type" = "#microsoft.graph.requestorManager"
                                managerLevel = 1
                            }
                        )
                    }
                )
            }
            accessPackage = @{
                id = [string]$AccessPackageId
            }
        }
        
        try {
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/assignmentPolicies" -Body ($policyBody | ConvertTo-Json -Depth 10) -ErrorAction Stop
            Write-Output "Policy created: $($duration.Name)"
        } catch {
            Write-Warning "Failed to create policy $($duration.Name): $_"
        }
    }
}

$Roles = @(
    'Reader'
    'Contributor'
    'Owner'
)

# Extract environment from subscription name to determine PIM requirements
# Example: sub-landingzone-dev-02 -> dev, sub-landingzone-prod-01 -> prod
$environment = if ($subname -match '-(dev|stg|prod)-') { $matches[1] } else { "unknown" }

foreach ($Role in $Roles) {
    $GroupName = "grp-rbac-$subname-$Role"
    
    if (!(Get-MgGroup -Filter "displayName eq '$GroupName'")) {
        $mailNickname = $GroupName -replace '\s','' -replace '[^a-zA-Z0-9]',''
        
        # Create role-assignable groups (IsAssignableToRole) for:
        # - Owner groups (all environments)
        # - Contributor groups (all environments - for security purposes)
        # - Reader groups remain as regular security groups
        $isRoleAssignable = ($Role -in @('Owner', 'Contributor'))
        
        if ($isRoleAssignable) {
            $Group = New-MgGroup -DisplayName $GroupName -SecurityEnabled:$true -IsAssignableToRole:$true -MailEnabled:$false -MailNickname $mailNickname
            Write-Output "Group created with IsAssignableToRole: $GroupName"
            Start-Sleep -Seconds 10
        } else {
            $Group = New-MgGroup -DisplayName $GroupName -SecurityEnabled:$true -MailEnabled:$false -MailNickname $mailNickname
            Write-Output "Group created: $GroupName"
            Start-Sleep -Seconds 10
        }
        
        # Apply role assignment for newly created group
        Set-RoleAssignment -GroupName $GroupName -Role $Role -SubscriptionName $subname
    }
    else {
        Write-Output "Group already exists: $GroupName"
        
        # Check and apply role assignment for existing group
        Set-RoleAssignment -GroupName $GroupName -Role $Role -SubscriptionName $subname
    }
    
    # Register for PIM only for:
    # - Owner groups (all environments)
    # - Contributor groups (prod environment and CAF/platform subscriptions)
    $shouldEnablePim = ($Role -eq 'Owner') -or ($Role -eq 'Contributor' -and $environment -in @('prod', 'unknown'))
    
    if ($shouldEnablePim) {
        $Group = Get-MgGroup -Filter "displayName eq '$GroupName'"
        
        Write-Output "Checking PIM registration for: $GroupName"
        
        # Try to register for PIM with retry logic (this is idempotent - safe to call even if already registered)
        $pimRegistered = $false
        $maxRetries = 3
        $retryCount = 0
        
        while (-not $pimRegistered -and $retryCount -lt $maxRetries) {
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/privilegedAccess/aadGroups/resources/register" -Body (@{externalId=$Group.Id} | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop
                Write-Output "✓ Group registered for PIM: $GroupName"
                $pimRegistered = $true
            } catch {
                $retryCount++
                if ($retryCount -lt $maxRetries) {
                    Write-Warning "PIM registration attempt $retryCount failed, retrying in 5 seconds..."
                    Start-Sleep -Seconds 5
                } else {
                    Write-Host "`n⚠️  PIM REGISTRATION FAILED for $GroupName after $maxRetries attempts ⚠️" -ForegroundColor Red
                    Write-Warning "Error: $_"
                    Write-Host "`nTo manually register this group for PIM:" -ForegroundColor Yellow
                    Write-Host "  1. Go to Azure Portal > Microsoft Entra ID > Groups" -ForegroundColor Yellow
                    Write-Host "  2. Search for and select: $GroupName" -ForegroundColor Yellow
                    Write-Host "  3. Click 'Privileged Identity Management' in the left menu" -ForegroundColor Yellow
                    Write-Host "  4. Click 'Enable PIM' or ensure it shows as enabled" -ForegroundColor Yellow
                    Write-Host "`nPress Enter after manually enabling PIM to continue, or Ctrl+C to cancel..." -ForegroundColor Cyan
                    Read-Host
                    Write-Output "Continuing with script..."
                }
            }
        }
    }
}

# Process Access Packages (if enabled)
if ($AccessPackages) {
    Write-Host "`n=== Processing Access Packages ===" -ForegroundColor Cyan

    # Extract service name and environment from subscription name
    # Example: sub-landingzone-dev-02 -> Landingzone, dev
    $serviceName = $subname -replace '^sub-', '' -replace '-(dev|stg|prod)-\d+$', ''
    $serviceName = (Get-Culture).TextInfo.ToTitleCase($serviceName)

# Extract environment (dev, stg, prod)
$environment = if ($subname -match '-(dev|stg|prod)-') { $matches[1] } else { "unknown" }

# 1. Reader Access Package
Write-Host "`n--- Processing Reader Access Package ---" -ForegroundColor Yellow
$readerAccessPackageName = "Subscription - $serviceName Reader"
$readerGroupName = "grp-rbac-$subname-Reader"
$readerGroup = Get-MgGroup -Filter "displayName eq '$readerGroupName'"

if ($readerGroup) {
    Set-AccessPackageWithGroup -AccessPackageName $readerAccessPackageName -CatalogId $readerContributorCatalogId -GroupId $readerGroup.Id -GroupName $readerGroupName -PolicyType "Reader" -Environment $environment
} else {
    Write-Warning "Reader group not found: $readerGroupName. Skipping access package creation."
}

# 2. Contributor Dev/Stg Access Package (only for dev and stg environments)
if ($environment -in @('dev', 'stg')) {
    Write-Host "`n--- Processing Contributor Dev/Stg Access Package ---" -ForegroundColor Yellow
    $contributorDevStgAccessPackageName = "Subscription - $serviceName Contributor Dev/Stg"
    $contributorGroupName = "grp-rbac-$subname-Contributor"
    $contributorGroup = Get-MgGroup -Filter "displayName eq '$contributorGroupName'"
    
    if ($contributorGroup) {
        Set-AccessPackageWithGroup -AccessPackageName $contributorDevStgAccessPackageName -CatalogId $readerContributorCatalogId -GroupId $contributorGroup.Id -GroupName $contributorGroupName -PolicyType "Contributor" -Environment $environment
    } else {
        Write-Warning "Contributor group not found: $contributorGroupName. Skipping access package creation."
    }
}

# 3. Contributor Prod Access Package (only for prod environment)
if ($environment -eq 'prod') {
    Write-Host "`n--- Processing Contributor Prod Access Package ---" -ForegroundColor Yellow
    $contributorProdAccessPackageName = "Subscription - $serviceName Contributor Prod"
    $contributorGroupName = "grp-rbac-$subname-Contributor"
    $contributorGroup = Get-MgGroup -Filter "displayName eq '$contributorGroupName'"
    
    if ($contributorGroup) {
        Set-AccessPackageWithGroup -AccessPackageName $contributorProdAccessPackageName -CatalogId $readerContributorCatalogId -GroupId $contributorGroup.Id -GroupName $contributorGroupName -PolicyType "Contributor" -Environment $environment
    } else {
        Write-Warning "Contributor group not found: $contributorGroupName. Skipping access package creation."
    }
}

# 3b. Contributor Access Package (for CAF/platform subscriptions with unknown environment)
if ($environment -eq 'unknown') {
    Write-Host "`n--- Processing Contributor Access Package (CAF/Platform) ---" -ForegroundColor Yellow
    $contributorAccessPackageName = "Subscription - $serviceName Contributor"
    $contributorGroupName = "grp-rbac-$subname-Contributor"
    $contributorGroup = Get-MgGroup -Filter "displayName eq '$contributorGroupName'"
    
    if ($contributorGroup) {
        Set-AccessPackageWithGroup -AccessPackageName $contributorAccessPackageName -CatalogId $readerContributorCatalogId -GroupId $contributorGroup.Id -GroupName $contributorGroupName -PolicyType "Contributor" -Environment $environment
    } else {
        Write-Warning "Contributor group not found: $contributorGroupName. Skipping access package creation."
    }
}

# 4. Owner Access Package (all environments, uses Owner catalog)
Write-Host "`n--- Processing Owner Access Package ---" -ForegroundColor Yellow
$ownerAccessPackageName = "Subscription - $serviceName Owner"
$ownerGroupName = "grp-rbac-$subname-Owner"
$ownerGroup = Get-MgGroup -Filter "displayName eq '$ownerGroupName'"

if ($ownerGroup) {
    Set-AccessPackageWithGroup -AccessPackageName $ownerAccessPackageName -CatalogId $ownerCatalogId -GroupId $ownerGroup.Id -GroupName $ownerGroupName -PolicyType "Owner" -Environment $environment
} else {
    Write-Warning "Owner group not found: $ownerGroupName. Skipping access package creation."
}

} else {
    Write-Host "`n⊘ Skipping Access Package creation (AccessPackages switch is disabled)" -ForegroundColor Yellow
}
