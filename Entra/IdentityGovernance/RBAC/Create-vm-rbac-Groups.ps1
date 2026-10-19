<#
.SYNOPSIS
    Automated RBAC Group Creation and Access Package Management for Azure Virtual Machines

.DESCRIPTION
    This script automates the creation and management of Azure RBAC groups with integrated 
    Entra ID access packages for VM-level access assignments.
    
    📋 Main Functions:
    1. 🔗 Connects to Azure and Microsoft Graph (detects runbook vs local execution)
    2. ✅ Validates VM exists and is accessible
    3. 🔐 Validates required permissions (Groups Administrator, Privileged Role Administrator)
    4. 👥 Creates RBAC security groups for VM access:
       - Users group: Normal security group for standard VM users
       - Admins group: Privileged security group (IsAssignableToRole) for VM administrators
    5. 📁 Creates or validates "Virtual Machines" access package catalog
    6. 📦 Creates and configures Entra ID access packages:
       - Users Access Package: 1 Year, 6 Month, 1 Month policies
       - Admins Access Package: 1 Month policy only
    7. 🔗 Adds groups as resources to respective access packages
    
    🔐 Security Features:
    - PIM-enabled groups for administrative access (IsAssignableToRole)
    - Two-stage approval process with manager and team approval
    - Automatic access expiration with configurable durations
    - Fallback reviewers for approval workflow
    
    📝 Naming Conventions:
    - Groups: grp-rbac-vm-{vm-name}-{Users|Admins}
    - Catalog: "Virtual Machines"
    - Access Packages:
      * "Virtual Machine - {vm-name} - Users"
      * "Virtual Machine - {vm-name} - Admins"
    
    ⚙️ Required Permissions:
    Microsoft Graph API:
    - Group.ReadWrite.All (create/manage security groups)
    - PrivilegedAccess.ReadWrite.AzureAD (PIM registration)
    - EntitlementManagement.ReadWrite.All (access packages and policies)
    
    Entra ID Directory Roles (required):
    - Groups Administrator (for group management including privileged groups)
    - Privileged Role Administrator (for PIM operations)
    
    Azure RBAC:
    - Reader access to the resource group or subscription containing the VM
    
    📦 Required Modules:
    - Az.Accounts (Azure authentication and VM validation)
    - Az.Compute (VM operations)
    - Microsoft.Graph.Authentication (Microsoft Graph connection)
    - Microsoft.Graph.Groups (Group operations)
    - Microsoft.Graph.Users (User operations)
    - Microsoft.Graph.Identity.Governance (Entitlement Management)
    
    🔄 Workflow:
    1. Detects execution context (runbook vs local)
    2. Connects to Azure and Microsoft Graph appropriately
    3. Validates VM exists
    4. Validates user permissions
    5. Creates Users and Admins groups
    6. Creates or validates "Virtual Machines" catalog
    7. Adds groups to catalog
    8. Creates Users access package with 1 Year, 6 Month, 1 Month policies
    9. Creates Admins access package with 1 Month policy
    10. Adds groups as resources to access packages
    
    ⚠️ Important Notes:
    - Admins group is created as privileged (IsAssignableToRole)
    - Access packages require two-stage approval (manager + team)
    - 1 Month policy does not allow extension
    - Script is idempotent - safe to run multiple times

.PARAMETER VM
    The name of the Azure Virtual Machine.
    Example: "vm-cloudockit", "vm-webapp-prod-01"

.EXAMPLE
    .\Create-vm-rbac-Groups.ps1 -VM "vm-cloudockit"
    
    Creates RBAC groups and access packages for vm-cloudockit.

.EXAMPLE
    .\Create-vm-rbac-Groups.ps1 -VM "vm-webapp-prod-01"
    
    Creates RBAC groups and access packages for vm-webapp-prod-01.

.NOTES
    Created by: Johan Öberg
    Last Updated: January 2026
    Version: 1.0
    Tenant ID: 
    
    Change Log:
    - v1.0: Initial version with VM RBAC group and access package automation
#>

param(
    [Parameter(Mandatory=$true, HelpMessage="The name of the Azure Virtual Machine (e.g., vm-cloudockit)")]
    [string]$VM
)

$tenantId = ""
$vmName = $VM
$fallbackReviewersGroupId = ""

# Track what was actually created/configured
$script:createdResources = @{
    UsersGroup = $false
    AdminsGroup = $false
    Catalog = $false
    UsersAccessPackage = $false
    AdminsAccessPackage = $false
}
$allowedRequestersGroup1 = ""
$allowedRequestersGroup2 = ""

# Detect if running in Azure Automation runbook
$isRunbook = $false
if ($env:AUTOMATION_ASSET_ACCOUNTID) {
    $isRunbook = $true
    Write-Host "🤖 Running in Azure Automation Runbook" -ForegroundColor Cyan
} else {
    Write-Host "💻 Running locally" -ForegroundColor Cyan
}

# Check required modules are installed
Write-Host "📦 Checking required modules..." -ForegroundColor Cyan

$requiredModules = @(
    @{Name='Az.Accounts'; MinVersion='5.3.2'},
    @{Name='Az.Compute'; MinVersion='1.0.0'},
    @{Name='Microsoft.Graph.Authentication'; MinVersion='2.0.0'},
    @{Name='Microsoft.Graph.Groups'; MinVersion='2.0.0'},
    @{Name='Microsoft.Graph.Users'; MinVersion='2.0.0'},
    @{Name='Microsoft.Graph.Identity.Governance'; MinVersion='2.0.0'}
)

$missingModules = @()
$versionConflicts = @()

foreach ($module in $requiredModules) {
    $installed = Get-Module -ListAvailable -Name $module.Name | Sort-Object Version -Descending | Select-Object -First 1
    
    if (-not $installed) {
        $missingModules += $module.Name
    } elseif ($installed.Version -lt [version]$module.MinVersion) {
        $versionConflicts += "$($module.Name) (installed: $($installed.Version), required: $($module.MinVersion))"
    }
}

if ($missingModules.Count -gt 0) {
    Write-Error "Missing required modules: $($missingModules -join ', ')"
    Write-Host "Install with: Install-Module $($missingModules -join ', ')" -ForegroundColor Yellow
    exit
}

if ($versionConflicts.Count -gt 0) {
    Write-Error "Module version conflicts detected: $($versionConflicts -join ', ')"
    Write-Host "Update with: Update-Module Az, Microsoft.Graph" -ForegroundColor Yellow
    exit
}

Write-Host "✓ All required modules are available" -ForegroundColor Green

# Try to load Az.Compute explicitly (it's not auto-loaded like Az.Accounts)
try {
    if (-not (Get-Module -Name Az.Compute)) {
        Import-Module Az.Compute -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Host "✓ Az.Compute module loaded" -ForegroundColor Green
    } else {
        Write-Host "✓ Az.Compute already loaded" -ForegroundColor Green
    }
} catch {
    Write-Error "Failed to load Az.Compute module. This is usually caused by a version mismatch with Az.Accounts already loaded in your session."
    Write-Host "`nTo fix this issue:" -ForegroundColor Yellow
    Write-Host "  1. Close this PowerShell session (type: exit)" -ForegroundColor Yellow
    Write-Host "  2. Start a new PowerShell session (type: pwsh)" -ForegroundColor Yellow
    Write-Host "  3. Run this script again" -ForegroundColor Yellow
    Write-Host "`nAlternatively, update all Az modules: Update-Module Az -Force" -ForegroundColor Yellow
    exit
}

# Connect to Azure
try {
    if ($isRunbook) {
        Write-Host "Connecting to Azure with Managed Identity..." -ForegroundColor Yellow
        Connect-AzAccount -Identity -TenantId $tenantId -ErrorAction Stop | Out-Null
        Write-Host "✓ Connected to Azure with Managed Identity" -ForegroundColor Green
    } else {
        $azContext = Get-AzContext
        if (-not $azContext -or $azContext.Tenant.Id -ne $tenantId) {
            Write-Host "Connecting to Azure (Tenant: $tenantId)..." -ForegroundColor Yellow
            Connect-AzAccount -TenantId $tenantId -ErrorAction Stop | Out-Null
            Write-Host "✓ Connected to Azure as: $((Get-AzContext).Account.Id)" -ForegroundColor Green
        } else {
            Write-Host "✓ Already connected to Azure as: $($azContext.Account.Id)" -ForegroundColor Green
        }
    }
} catch {
    Write-Error "Failed to connect to Azure: $_"
    exit
}

# Validate VM exists
Write-Host "`n🔍 Validating Virtual Machine..." -ForegroundColor Cyan
$currentContext = Get-AzContext
$validatedVM = Get-AzVM -Name $vmName -ErrorAction SilentlyContinue

if (-not $validatedVM) {
    Write-Host "VM not found in current subscription ($($currentContext.Subscription.Name))" -ForegroundColor Yellow
    Write-Host "Searching across all accessible subscriptions..." -ForegroundColor Yellow
    
    $allSubscriptions = Get-AzSubscription
    $foundVM = $null
    $foundInSubscription = $null
    
    foreach ($sub in $allSubscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        $vm = Get-AzVM -Name $vmName -ErrorAction SilentlyContinue
        if ($vm) {
            $foundVM = $vm
            $foundInSubscription = $sub
            break
        }
    }
    
    if ($foundVM) {
        Write-Host "✓ VM found in subscription: $($foundInSubscription.Name)" -ForegroundColor Green
        $validatedVM = $foundVM
    } else {
        # Restore original context
        Set-AzContext -SubscriptionId $currentContext.Subscription.Id | Out-Null
        
        Write-Error "Virtual Machine '$vmName' not found in any accessible subscription."
        Write-Host "`nSearching for VMs with similar names across all subscriptions..." -ForegroundColor Yellow
        
        $similarVMs = @()
        foreach ($sub in $allSubscriptions) {
            Set-AzContext -SubscriptionId $sub.Id | Out-Null
            $vms = Get-AzVM | Where-Object { $_.Name -like "*$vmName*" }
            foreach ($vm in $vms) {
                $similarVMs += [PSCustomObject]@{
                    Name = $vm.Name
                    ResourceGroupName = $vm.ResourceGroupName
                    Location = $vm.Location
                    Subscription = $sub.Name
                }
            }
        }
        
        if ($similarVMs.Count -gt 0) {
            Write-Host "Similar VMs found:" -ForegroundColor Yellow
            $similarVMs | Format-Table Name, ResourceGroupName, Location, Subscription -AutoSize
        } else {
            Write-Host "No similar VMs found. Listing all available VMs across subscriptions:" -ForegroundColor Yellow
            $allVMs = @()
            foreach ($sub in $allSubscriptions) {
                Set-AzContext -SubscriptionId $sub.Id | Out-Null
                $vms = Get-AzVM
                foreach ($vm in $vms) {
                    $allVMs += [PSCustomObject]@{
                        Name = $vm.Name
                        ResourceGroupName = $vm.ResourceGroupName
                        Location = $vm.Location
                        Subscription = $sub.Name
                    }
                }
            }
            $allVMs | Format-Table Name, ResourceGroupName, Location, Subscription -AutoSize
        }
        
        # Restore original context
        Set-AzContext -SubscriptionId $currentContext.Subscription.Id | Out-Null
        exit
    }
}

Write-Host "✓ VM validated: $($validatedVM.Name) in resource group $($validatedVM.ResourceGroupName)" -ForegroundColor Green
$vmSubscription = (Get-AzContext).Subscription.Name
Write-Host "  Subscription: $vmSubscription" -ForegroundColor Gray

# Connect to Microsoft Graph
try {
    if ($isRunbook) {
        Write-Host "`nConnecting to Microsoft Graph with Managed Identity..." -ForegroundColor Yellow
        Connect-MgGraph -Identity -TenantId $tenantId -ErrorAction Stop -NoWelcome
        Write-Host "✓ Connected to Microsoft Graph with Managed Identity" -ForegroundColor Green
    } else {
        $context = Get-MgContext
        if (-not $context) {
            Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Yellow
            Connect-MgGraph -Scopes "Group.ReadWrite.All","PrivilegedAccess.ReadWrite.AzureAD","EntitlementManagement.ReadWrite.All" -TenantId $tenantId -ErrorAction Stop -NoWelcome
            Write-Host "✓ Connected to Microsoft Graph" -ForegroundColor Green
        } else {
            Write-Host "✓ Already connected to Microsoft Graph" -ForegroundColor Green
        }
    }
} catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit
}

# Validate permissions
Write-Host "`n🔐 Checking required permissions..." -ForegroundColor Cyan

$currentUser = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me"
$userId = $currentUser.id

# Check for Groups Administrator role
$groupsAdminRole = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=displayName eq 'Groups Administrator'"
$hasGroupsAdminRole = $false
if ($groupsAdminRole.value.Count -gt 0) {
    $roleMembers = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($groupsAdminRole.value[0].id)/members"
    $hasGroupsAdminRole = $roleMembers.value | Where-Object { $_.id -eq $userId }
}

# Check for Privileged Role Administrator role
$privilegedRoleAdminRole = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=displayName eq 'Privileged Role Administrator'"
$hasPrivilegedRoleAdminRole = $false
if ($privilegedRoleAdminRole.value.Count -gt 0) {
    $roleMembers = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($privilegedRoleAdminRole.value[0].id)/members"
    $hasPrivilegedRoleAdminRole = $roleMembers.value | Where-Object { $_.id -eq $userId }
}

# Check for Global Administrator role
$globalAdminRole = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles?`$filter=displayName eq 'Global Administrator'"
$hasGlobalAdminRole = $false
if ($globalAdminRole.value.Count -gt 0) {
    $roleMembers = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryRoles/$($globalAdminRole.value[0].id)/members"
    $hasGlobalAdminRole = $roleMembers.value | Where-Object { $_.id -eq $userId }
}

# Display permission status
Write-Host "`nPermission Check Results:" -ForegroundColor Yellow
Write-Host "  Current User: $($currentUser.userPrincipalName)" -ForegroundColor White
Write-Host "  Groups Administrator: $(if($hasGroupsAdminRole){'✓ Yes'}else{'✗ No'})" -ForegroundColor $(if($hasGroupsAdminRole){'Green'}else{'Red'})
Write-Host "  Privileged Role Administrator: $(if($hasPrivilegedRoleAdminRole){'✓ Yes'}else{'✗ No'})" -ForegroundColor $(if($hasPrivilegedRoleAdminRole){'Green'}else{'Red'})
Write-Host "  Global Administrator: $(if($hasGlobalAdminRole){'✓ Yes'}else{'✗ No'})" -ForegroundColor $(if($hasGlobalAdminRole){'Green'}else{'Red'})

$hasRequiredPermissions = $hasGroupsAdminRole -and $hasPrivilegedRoleAdminRole -or $hasGlobalAdminRole

if (-not $hasRequiredPermissions) {
    Write-Error "`n❌ Insufficient permissions!"
    Write-Error "Required: Groups Administrator AND Privileged Role Administrator (or Global Administrator)"
    Write-Error "This is needed to create privileged groups (IsAssignableToRole) and manage access packages."
    exit
}

Write-Host "`n✓ You have sufficient permissions to proceed.`n" -ForegroundColor Green

# Create groups
Write-Host "👥 Creating RBAC Groups..." -ForegroundColor Cyan

# Remove 'vm-' prefix if present to avoid duplication in group name
$vmNameClean = $vmName -replace '^vm-', ''
$usersGroupName = "sg-rbac-vm-$vmNameClean-Users"
$adminsGroupName = "psg-rbac-vm-$vmNameClean-Admins"

# Create Users group (normal security group)
$usersGroup = Get-MgGroup -Filter "displayName eq '$usersGroupName'" -ErrorAction SilentlyContinue
if (-not $usersGroup) {
    $mailNickname = $usersGroupName -replace '\s','' -replace '[^a-zA-Z0-9]',''
    $usersGroup = New-MgGroup -DisplayName $usersGroupName -SecurityEnabled:$true -MailEnabled:$false -MailNickname $mailNickname
    Write-Host "✓ Users group created: $usersGroupName" -ForegroundColor Green
    $script:createdResources.UsersGroup = $true
    Start-Sleep -Seconds 7
} else {
    Write-Host "✓ Users group already exists: $usersGroupName" -ForegroundColor Green
}

# Create Admins group (privileged security group)
$adminsGroup = Get-MgGroup -Filter "displayName eq '$adminsGroupName'" -ErrorAction SilentlyContinue
if (-not $adminsGroup) {
    $mailNickname = $adminsGroupName -replace '\s','' -replace '[^a-zA-Z0-9]',''
    $adminsGroup = New-MgGroup -DisplayName $adminsGroupName -SecurityEnabled:$true -IsAssignableToRole:$true -MailEnabled:$false -MailNickname $mailNickname
    Write-Host "✓ Admins group created (privileged): $adminsGroupName" -ForegroundColor Green
    $script:createdResources.AdminsGroup = $true
    Start-Sleep -Seconds 7
} else {
    Write-Host "✓ Admins group already exists: $adminsGroupName" -ForegroundColor Green
}

# Create or validate Virtual Machines catalog
Write-Host "`n📁 Processing Virtual Machines Catalog..." -ForegroundColor Cyan

$catalogName = "Virtual Machines"
$existingCatalogs = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs?`$filter=displayName eq '$catalogName'"

if ($existingCatalogs.value.Count -gt 0) {
    $catalog = $existingCatalogs.value[0]
    Write-Host "✓ Catalog already exists: $catalogName (ID: $($catalog.id))" -ForegroundColor Green
} else {
    Write-Host "Creating new catalog: $catalogName" -ForegroundColor Yellow
    $catalogBody = @{
        displayName = $catalogName
        description = "Access packages for Virtual Machine RBAC groups"
        isExternallyVisible = $false
    }
    
    try {
        $catalog = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs" -Body ($catalogBody | ConvertTo-Json) -ErrorAction Stop
        Write-Host "✓ Catalog created: $catalogName (ID: $($catalog.id))" -ForegroundColor Green
        $script:createdResources.Catalog = $true
    } catch {
        Write-Error "Failed to create catalog: $_"
        exit
    }
}

$catalogId = $catalog.id

# Function to add group to catalog
function Add-GroupToCatalog {
    param(
        [string]$CatalogId,
        [string]$GroupId,
        [string]$GroupName
    )
    
    # Check if group is already in catalog
    $catalogResources = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$CatalogId/accessPackageResources?`$filter=originId eq '$GroupId'"
    
    if ($catalogResources.value.Count -gt 0) {
        Write-Host "  ✓ Group already in catalog: $GroupName" -ForegroundColor Green
        return $catalogResources.value[0]
    }
    
    Write-Host "  Adding group to catalog: $GroupName" -ForegroundColor Yellow
    
    $resourceRequestBody = @{
        catalogId = $CatalogId
        requestType = "AdminAdd"
        accessPackageResource = @{
            originId = $GroupId
            originSystem = "AadGroup"
        }
    }
    
    try {
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageResourceRequests" -Body ($resourceRequestBody | ConvertTo-Json -Depth 10) -ErrorAction Stop | Out-Null
        Write-Host "  ✓ Group added to catalog: $GroupName" -ForegroundColor Green
        
        # Wait for resource to be available
        Start-Sleep -Seconds 5
        
        # Retrieve the catalog resource
        $catalogResources = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$CatalogId/accessPackageResources?`$filter=originId eq '$GroupId'"
        return $catalogResources.value[0]
    } catch {
        Write-Warning "Failed to add group to catalog: $_"
        return $null
    }
}

# Add groups to catalog
Write-Host "`n📦 Adding groups to catalog..." -ForegroundColor Cyan
$usersGroup = Get-MgGroup -Filter "displayName eq '$usersGroupName'"
$adminsGroup = Get-MgGroup -Filter "displayName eq '$adminsGroupName'"

$usersCatalogResource = Add-GroupToCatalog -CatalogId $catalogId -GroupId $usersGroup.Id -GroupName $usersGroupName
$adminsCatalogResource = Add-GroupToCatalog -CatalogId $catalogId -GroupId $adminsGroup.Id -GroupName $adminsGroupName

# Function to create access package
function New-VMAccessPackage {
    param(
        [string]$DisplayName,
        [string]$Description,
        [string]$CatalogId
    )
    
    # Check if access package exists - get all packages and filter client-side
    try {
        Write-Host "  Checking for existing access package..." -ForegroundColor Gray
        $allPackages = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages"
        $accessPackage = $allPackages.value | Where-Object { $_.displayName -eq $DisplayName -and $_.catalogId -eq $CatalogId } | Select-Object -First 1
        
        if ($accessPackage) {
            Write-Host "  ✓ Access package already exists: $DisplayName" -ForegroundColor Green
            return $accessPackage
        }
    } catch {
        # If query fails, try creating anyway
        Write-Host "  Could not query existing packages, will attempt to create..." -ForegroundColor Gray
    }
    
    Write-Host "  Creating access package: $DisplayName" -ForegroundColor Yellow
    
    $packageBody = @{
        displayName = $DisplayName
        description = $Description
        catalogId = $CatalogId
        isHidden = $false
    }
    
    try {
        $accessPackage = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages" -Body ($packageBody | ConvertTo-Json) -ErrorAction Stop
        Write-Host "  ✓ Access package created: $DisplayName" -ForegroundColor Green
        return $accessPackage
    } catch {
        $errorMessage = $_.Exception.Message
        if ($errorMessage -like "*Duplicate*" -or $errorMessage -like "*409*" -or $errorMessage -like "*Conflict*") {
            # Package already exists, fetch it
            Write-Host "  ℹ️  Access package already exists, retrieving..." -ForegroundColor Yellow
            try {
                $allPackages = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages" -ErrorAction Stop
                $accessPackage = $allPackages.value | Where-Object { $_.displayName -eq $DisplayName -and $_.catalogId -eq $CatalogId } | Select-Object -First 1
                if ($accessPackage) {
                    Write-Host "  ✓ Access package found: $DisplayName" -ForegroundColor Green
                    return $accessPackage
                } else {
                    Write-Warning "Could not find existing access package: $DisplayName"
                    return $null
                }
            } catch {
                Write-Warning "Failed to retrieve existing access package: $_"
                return $null
            }
        } else {
            # Suppress the full error output, just show a concise warning
            Write-Warning "Failed to create access package '$DisplayName': $($_.Exception.Message -split "`n" | Select-Object -First 1)"
            return $null
        }
    }
}

# Function to add group to access package
function Add-GroupToAccessPackage {
    param(
        [string]$AccessPackageId,
        [string]$GroupId,
        [string]$GroupName,
        [string]$CatalogResourceId
    )
    
    # Check if group is already a resource in the access package
    $resources = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/$AccessPackageId/accessPackageResourceRoleScopes"
    
    $groupAlreadyAdded = $false
    foreach ($resource in $resources.value) {
        if ($resource.accessPackageResourceRole.accessPackageResource.originId -eq $GroupId) {
            $groupAlreadyAdded = $true
            Write-Host "  ✓ Group already in access package: $GroupName" -ForegroundColor Green
            return
        }
    }
    
    if (-not $groupAlreadyAdded) {
        Write-Host "  Adding group to access package: $GroupName" -ForegroundColor Yellow
        
        $roleOriginId = "Member_$GroupId"
        
        $scopeBody = @{
            accessPackageResourceRole = @{
                displayName = "Member"
                originId = $roleOriginId
                originSystem = "AadGroup"
                accessPackageResource = @{
                    id = $CatalogResourceId
                    originId = $GroupId
                    originSystem = "AadGroup"
                }
            }
            accessPackageResourceScope = @{
                id = $CatalogResourceId
                originId = $GroupId
                originSystem = "AadGroup"
            }
        }
        
        try {
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/$AccessPackageId/accessPackageResourceRoleScopes" -Body ($scopeBody | ConvertTo-Json -Depth 10) -ErrorAction Stop | Out-Null
            Write-Host "  ✓ Group added to access package: $GroupName" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to add group to access package: $_"
        }
    }
}

# Function to create access package policy
function New-AccessPackagePolicy {
    param(
        [string]$AccessPackageId,
        [string]$DisplayName,
        [string]$Description,
        [string]$Duration,
        [bool]$AllowExtension = $true
    )
    
    # Check if policy already exists
    $existingPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/assignmentPolicies?`$filter=accessPackage/id eq '$AccessPackageId' and displayName eq '$DisplayName'"
    
    if ($existingPolicies.value.Count -gt 0) {
        Write-Host "  ✓ Policy already exists: $DisplayName" -ForegroundColor Green
        return
    }
    
    Write-Host "  Creating policy: $DisplayName" -ForegroundColor Yellow
    
    $policyBody = @{
        displayName = $DisplayName
        description = $Description
        allowedTargetScope = "specificDirectoryUsers"
        specificAllowedTargets = @(
            @{
                "@odata.type" = "#microsoft.graph.groupMembers"
                groupId = $script:allowedRequestersGroup1
                description = "dyn-all-it"
            },
            @{
                "@odata.type" = "#microsoft.graph.groupMembers"
                groupId = $script:allowedRequestersGroup2
                description = "grp-all-pim-users"
            }
        )
        expiration = @{
            type = "afterDuration"
            duration = $Duration
        }
        requestorSettings = @{
            enableTargetsToSelfAddAccess = $true
            enableTargetsToSelfUpdateAccess = $AllowExtension
            enableTargetsToSelfRemoveAccess = $true
            enableOnBehalfRequestorsToAddAccess = $false
            enableOnBehalfRequestorsToUpdateAccess = $false
            enableOnBehalfRequestorsToRemoveAccess = $false
            onBehalfRequestors = @()
            allowCustomAssignmentSchedule = $true
        }
        requestApprovalSettings = @{
            isApprovalRequiredForAdd = $true
            isApprovalRequiredForUpdate = $true
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
                },
                @{
                    durationBeforeAutomaticDenial = "P7D"
                    durationBeforeEscalation = "PT0S"
                    isEscalationEnabled = $false
                    isApproverJustificationRequired = $false
                    fallbackPrimaryApprovers = @()
                    fallbackEscalationApprovers = @()
                    escalationApprovers = @()
                    primaryApprovers = @(
                        @{
                            "@odata.type" = "#microsoft.graph.groupMembers"
                            groupId = $script:teamTopNotchGroupId
                            description = "Team Top Notch"
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
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/assignmentPolicies" -Body ($policyBody | ConvertTo-Json -Depth 10) -ErrorAction Stop | Out-Null
        Write-Host "  ✓ Policy created: $DisplayName" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to create policy: $_"
    }
}

# Create Users Access Package
Write-Host "`n📦 Creating Users Access Package..." -ForegroundColor Cyan
$usersAccessPackageName = "Virtual Machine - $vmName - Users"
$usersAccessPackageDescription = "Gives User permission to login to $vmName via the AAD Login extension.`n`nRDP to server and login with AzureAD\<yourUPN>"

$usersAccessPackage = New-VMAccessPackage -DisplayName $usersAccessPackageName -Description $usersAccessPackageDescription -CatalogId $catalogId

if ($usersAccessPackage) {
    $script:createdResources.UsersAccessPackage = $true
}

if ($usersAccessPackage -and $usersCatalogResource) {
    Add-GroupToAccessPackage -AccessPackageId $usersAccessPackage.id -GroupId $usersGroup.Id -GroupName $usersGroupName -CatalogResourceId $usersCatalogResource.id
    
    # Create policies for Users access package
    Write-Host "`n  Creating policies for Users access package..." -ForegroundColor Cyan
    New-AccessPackagePolicy -AccessPackageId $usersAccessPackage.id -DisplayName "1 Year Policy" -Description "Virtual Machine $vmName`nAssignment expires after 365 days." -Duration "P365D" -AllowExtension $true
    New-AccessPackagePolicy -AccessPackageId $usersAccessPackage.id -DisplayName "6 Month Policy" -Description "Virtual Machine $vmName`nAssignment expires after 180 days." -Duration "P180D" -AllowExtension $true
    New-AccessPackagePolicy -AccessPackageId $usersAccessPackage.id -DisplayName "1 Month Policy" -Description "Virtual Machine $vmName`nAssignment expires after 30 days." -Duration "P30D" -AllowExtension $false
}

# Create Admins Access Package
Write-Host "`n📦 Creating Admins Access Package..." -ForegroundColor Cyan
$adminsAccessPackageName = "Virtual Machine - $vmName - Admins"
$adminsAccessPackageDescription = "Gives Administrator permission to $vmName via the AAD Login extension.`n`nRDP to server and login with AzureAD\<yourUPN>"

$adminsAccessPackage = New-VMAccessPackage -DisplayName $adminsAccessPackageName -Description $adminsAccessPackageDescription -CatalogId $catalogId

if ($adminsAccessPackage) {
    $script:createdResources.AdminsAccessPackage = $true
}

if ($adminsAccessPackage -and $adminsCatalogResource) {
    Add-GroupToAccessPackage -AccessPackageId $adminsAccessPackage.id -GroupId $adminsGroup.Id -GroupName $adminsGroupName -CatalogResourceId $adminsCatalogResource.id
    
    # Create policy for Admins access package (1 Month only)
    Write-Host "`n  Creating policy for Admins access package..." -ForegroundColor Cyan
    New-AccessPackagePolicy -AccessPackageId $adminsAccessPackage.id -DisplayName "1 Month Policy" -Description "Virtual Machine $vmName Admins`nAssignment expires after 30 days." -Duration "P30D" -AllowExtension $false
}

# Final summary
Write-Host "\n" -NoNewline
$hasErrors = -not ($script:createdResources.UsersAccessPackage -and $script:createdResources.AdminsAccessPackage)

if ($hasErrors) {
    Write-Host "⚠️  Script completed with warnings" -ForegroundColor Yellow
} else {
    Write-Host "✅ Script completed successfully!" -ForegroundColor Green
}

Write-Host "\nResources Status:" -ForegroundColor Cyan

Write-Host "  Groups:" -ForegroundColor Yellow
Write-Host "    $(if($script:createdResources.UsersGroup){'✅ Created'}else{'ℹ️  Existed'}): $usersGroupName" -ForegroundColor White
Write-Host "    $(if($script:createdResources.AdminsGroup){'✅ Created'}else{'ℹ️  Existed'}): $adminsGroupName (privileged)" -ForegroundColor White

Write-Host "  Catalog:" -ForegroundColor Yellow
Write-Host "    $(if($script:createdResources.Catalog){'✅ Created'}else{'ℹ️  Existed'}): $catalogName" -ForegroundColor White

Write-Host "  Access Packages:" -ForegroundColor Yellow
if ($script:createdResources.UsersAccessPackage) {
    Write-Host "    ✅ Created: $usersAccessPackageName (1Y, 6M, 1M policies)" -ForegroundColor White
} else {
    Write-Host "    ❌ Failed: $usersAccessPackageName" -ForegroundColor Red
}
if ($script:createdResources.AdminsAccessPackage) {
    Write-Host "    ✅ Created: $adminsAccessPackageName (1M policy)" -ForegroundColor White
} else {
    Write-Host "    ❌ Failed: $adminsAccessPackageName" -ForegroundColor Red
}

if ($hasErrors) {
    Write-Host "\n⚠️  Some access packages failed to create. Check the warnings above for details." -ForegroundColor Yellow
}

