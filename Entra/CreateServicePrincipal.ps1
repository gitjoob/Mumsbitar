<#
.SYNOPSIS
    Creates or updates a service principal for MFA automation with assessment or full permissions.

.DESCRIPTION
    This script supports two modes for flexible customer engagement workflows:
    
    ASSESSMENT MODE (-AssessmentMode):
    Creates a service principal with READ-ONLY permissions for running assessments:
    - User.Read.All (read user information)
    - Device.Read.All (read device information)
    - Group.Read.All (read group information)
    - Policy.Read.All (read policy configurations)
    - RoleManagement.Read.Directory (read role assignments)
    - Directory.Read.All (read directory data)
    
    FULL MODE (default):
    Creates or upgrades a service principal with READ-WRITE permissions for enrollment/enforcement:
    - All assessment mode permissions PLUS:
    - Group.ReadWrite.All (create and manage groups)
    - Policy.ReadWrite.ConditionalAccess (create and manage CAPs)
    - Policy.ReadWrite.AuthenticationMethod (manage authentication methods)
    - User.ReadWrite.All (GUID password reset for compliant users)
    
    The script intelligently:
    - Checks for existing application and reuses it
    - Detects valid non-expired secrets and reuses them
    - Upgrades permissions when switching from assessment to full mode
    - Grants admin consent automatically
    - Exports credentials to auth.json

.PARAMETER AssessmentMode
    When specified, creates service principal with read-only permissions suitable for assessments.
    When omitted, creates/upgrades to full read-write permissions for enrollment and enforcement.

.NOTES
    Requirements:
    - User must have Application Administrator or Global Administrator role
    - Microsoft.Graph.Applications module
    - Microsoft.Graph.Authentication module
    
    Generated auth.json location: Same directory as this script
    Secret expires after 30 days for security compliance
    
    RECOMMENDED WORKFLOW FOR CUSTOMER ENGAGEMENTS:
    1. Customer runs: .\CreateServicePrincipal.ps1 -AssessmentMode
    2. Customer shares auth.json securely with consultant
    3. Consultant runs: .\EnrollmentPhase.ps1 -WhatIf
    4. Consultant runs: .\PrivilegedAccountsPhase.ps1 -WhatIf
    5. Customer reviews assessment reports
    6. Customer runs: .\CreateServicePrincipal.ps1 (upgrades to full permissions)
    7. Consultant runs enrollment and enforcement

.EXAMPLE
    .\CreateServicePrincipal.ps1 -AssessmentMode
    # Creates service principal with read-only permissions for assessments
    
.EXAMPLE
    .\CreateServicePrincipal.ps1
    # Creates or upgrades service principal with full read-write permissions
#>

# Module dependencies (imported at runtime to avoid version conflicts with already-loaded assemblies)
# Microsoft.Graph.Applications, Microsoft.Graph.Authentication

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$AssessmentMode
)

# ============================================================================
# SCRIPT CONFIGURATION
# ============================================================================

$mode = if ($AssessmentMode) { "ASSESSMENT (Read-Only)" } else { "FULL (Read-Write)" }
Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  Service Principal Creation Utility                          ║" -ForegroundColor Cyan
Write-Host "║  Mode: $(($mode).PadRight(50)) ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

# Check for required modules
Write-Host "Checking required PowerShell modules..." -ForegroundColor Cyan

$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Applications"
)

foreach ($module in $requiredModules) {
    if (Get-Module -ListAvailable -Name $module) {
        Write-Host "  ✓ $module" -ForegroundColor Green
    }
    else {
        Write-Host "  ✗ $module is NOT installed" -ForegroundColor Red
        Write-Host "`n  Install it using: Install-Module -Name $module -Scope CurrentUser`n" -ForegroundColor Yellow
        exit 1
    }
}

# ============================================================================
# DEFINE PERMISSION SETS
# ============================================================================

Write-Host "`nDefining permission sets..." -ForegroundColor Cyan

# Microsoft Graph resource ID
$microsoftGraphResourceId = "00000003-0000-0000-c000-000000000000"

# Assessment Mode Permissions (Read-Only) - Using Application permissions (Role)
$assessmentPermissions = @(
    @{ Id = "df021288-bdef-4463-88db-98f22de89214"; Type = "Role"; Name = "User.Read.All" },
    @{ Id = "7ab1d382-f21e-4acd-a863-ba3e13f7da61"; Type = "Role"; Name = "Directory.Read.All" },
    @{ Id = "5b567255-7703-4780-807c-7be8301ae99b"; Type = "Role"; Name = "Group.Read.All" },
    @{ Id = "246dd0d5-5bd0-4def-940b-0421030a5b68"; Type = "Role"; Name = "Policy.Read.All" },
    @{ Id = "483bed4a-2ad3-4361-a73b-c83ccdbdc53c"; Type = "Role"; Name = "RoleManagement.Read.Directory" },
    @{ Id = "7438b122-aefc-4978-80ed-43db9fcc7715"; Type = "Role"; Name = "Device.Read.All" },
    @{ Id = "38d9df27-64da-44fd-b7c5-a6fbac20248f"; Type = "Role"; Name = "UserAuthenticationMethod.Read.All" },
    @{ Id = "b0afded3-3588-46d8-8b3d-9842eff778da"; Type = "Role"; Name = "AuditLog.Read.All" }
)

# Additional permissions for Full Mode (Read-Write) - Using Application permissions (Role)
$fullModeAdditionalPermissions = @(
    @{ Id = "62a82d76-70ea-41e2-9197-370581804d09"; Type = "Role"; Name = "Group.ReadWrite.All" },
    @{ Id = "01c0a623-fc9b-48e9-b794-0756f8e8f067"; Type = "Role"; Name = "Policy.ReadWrite.ConditionalAccess" },
    @{ Id = "29c18626-4985-4dcd-85c0-193eef327366"; Type = "Role"; Name = "Policy.ReadWrite.AuthenticationMethod" },
    @{ Id = "9e3f62cf-ca93-4989-b6ce-bf83c28f9fe8"; Type = "Role"; Name = "RoleManagement.ReadWrite.Directory" },
    @{ Id = "741f803b-c850-494e-b5df-cde7c675a1ca"; Type = "Role"; Name = "User.ReadWrite.All" },
    @{ Id = "9a5d68dd-52b0-4cc2-bd40-abcf44ac3a30"; Type = "Role"; Name = "Application.Read.All" }
)

# Select permission set based on mode
$selectedPermissions = if ($AssessmentMode) {
    Write-Host "  Mode: Assessment (Read-Only)" -ForegroundColor Yellow
    $assessmentPermissions
} else {
    Write-Host "  Mode: Full (Assessment + Enrollment/Enforcement)" -ForegroundColor Green
    $assessmentPermissions + $fullModeAdditionalPermissions
}

Write-Host "`n  Permissions to be assigned:" -ForegroundColor Gray
foreach ($perm in $selectedPermissions) {
    Write-Host "    • $($perm.Name)" -ForegroundColor Gray
}

# ============================================================================
# CONNECT TO MICROSOFT GRAPH
# ============================================================================

Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Cyan

# Disconnect any existing connection
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
catch {
    # Ignore errors if not connected
}

try {
    Connect-MgGraph -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All" -ErrorAction Stop | Out-Null
    Write-Host "  ✓ Successfully connected to Microsoft Graph" -ForegroundColor Green
}
catch {
    Write-Host "  ✗ Failed to connect to Microsoft Graph" -ForegroundColor Red
    Write-Host "    Error: $_" -ForegroundColor Red
    exit 1
}

# Get current context
$context = Get-MgContext
Write-Host "  ✓ Authenticated as: $($context.Account)" -ForegroundColor Green
Write-Host "  ✓ Tenant ID: $($context.TenantId)" -ForegroundColor Green

# ============================================================================
# VALIDATE PERMISSIONS
# ============================================================================

Write-Host "`nValidating your permissions..." -ForegroundColor Cyan

try {
    # Check if user has necessary permissions by attempting to read the required scopes
    $grantedScopes = $context.Scopes
    
    $requiredScopes = @(
        "Application.ReadWrite.All",
        "AppRoleAssignment.ReadWrite.All"
    )
    
    $missingScopes = @()
    foreach ($scope in $requiredScopes) {
        if ($grantedScopes -notcontains $scope) {
            $missingScopes += $scope
        }
    }
    
    if ($missingScopes.Count -gt 0) {
        Write-Host "`n  ✗ Insufficient permissions detected" -ForegroundColor Red
        Write-Host "    Missing scopes:" -ForegroundColor Red
        foreach ($scope in $missingScopes) {
            Write-Host "      • $scope" -ForegroundColor Red
        }
        Write-Host "`n  Required permissions:" -ForegroundColor Yellow
        Write-Host "    • Application.ReadWrite.All - To create/update app registrations" -ForegroundColor Yellow
        Write-Host "    • AppRoleAssignment.ReadWrite.All - To grant admin consent" -ForegroundColor Yellow
        Write-Host "`n  To resolve this issue:" -ForegroundColor Yellow
        Write-Host "    1. You need one of these roles:" -ForegroundColor Yellow
        Write-Host "       • Global Administrator (has both permissions)" -ForegroundColor Yellow
        Write-Host "       • Application Administrator + Privileged Role Administrator" -ForegroundColor Yellow
        Write-Host "    2. Contact your Global Administrator to assign the required role" -ForegroundColor Yellow
        Write-Host "    3. Then re-run this script" -ForegroundColor Yellow
        Write-Host "`n  Alternative: Manual Setup" -ForegroundColor Cyan
        Write-Host "    If you cannot get these permissions, ask your Global Administrator to:" -ForegroundColor Cyan
        Write-Host "    1. Create the app registration manually in Azure Portal" -ForegroundColor Cyan
        Write-Host "    2. Assign the required API permissions" -ForegroundColor Cyan
        Write-Host "    3. Grant admin consent" -ForegroundColor Cyan
        Write-Host "    4. Create a client secret" -ForegroundColor Cyan
        Write-Host "    5. Provide you with the credentials for auth.json" -ForegroundColor Cyan
        exit 1
    }
    
    Write-Host "  ✓ All required permissions granted" -ForegroundColor Green
    foreach ($scope in $requiredScopes) {
        Write-Host "    • $scope" -ForegroundColor Gray
    }
}
catch {
    Write-Host "`n  ⚠ Could not validate permissions" -ForegroundColor Yellow
    Write-Host "    Error: $_" -ForegroundColor Yellow
    Write-Host "    Continuing anyway - errors will occur if permissions are missing" -ForegroundColor Yellow
}

# ============================================================================
# STEP 1: CREATE OR UPDATE APPLICATION REGISTRATION
# ============================================================================

Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  STEP 1: Application Registration                            ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

$appDisplayName = "sp-onevinn-prmfa"
$app = $null
$existingSecret = $null
$reuseSecret = $false

try {
    # Check if app already exists
    $existingApp = Get-MgApplication -Filter "displayName eq '$appDisplayName'" -ErrorAction SilentlyContinue | Select-Object -First 1
    
    if ($existingApp) {
        $app = $existingApp
        Write-Host "`n⚠  Application already exists: $appDisplayName" -ForegroundColor Yellow
        Write-Host "   Application ID: $($app.AppId)" -ForegroundColor Gray
        Write-Host "   Object ID: $($app.Id)" -ForegroundColor Gray
        
        # Check for existing valid secrets
        $secrets = $app.PasswordCredentials
        if ($secrets -and $secrets.Count -gt 0) {
            # Find the most recent non-expired secret
            $validSecret = $secrets | Where-Object { $_.EndDateTime -gt (Get-Date) } | Sort-Object EndDateTime -Descending | Select-Object -First 1
            
            if ($validSecret) {
                $daysRemaining = [math]::Round(($validSecret.EndDateTime - (Get-Date)).TotalDays, 1)
                Write-Host "`n   ✓ Found valid secret" -ForegroundColor Green
                Write-Host "     Expires: $($validSecret.EndDateTime.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor Gray
                Write-Host "     Days remaining: $daysRemaining" -ForegroundColor Gray
                
                # Check if auth.json exists and contains a matching secret
                $authJsonPath = Join-Path $PSScriptRoot "auth.json"
                if (Test-Path $authJsonPath) {
                    try {
                        $existingAuth = Get-Content $authJsonPath -Raw | ConvertFrom-Json
                        
                        # Validate auth.json has required fields
                        if (-not $existingAuth.ClientSecret -or -not $existingAuth.ExpiryDate) {
                            Write-Host "`n   ⚠  auth.json is incomplete (missing ClientSecret or ExpiryDate)" -ForegroundColor Yellow
                            Write-Host "     A new secret will be created" -ForegroundColor Cyan
                            $reuseSecret = $false
                        }
                        # Validate the expiry date in auth.json matches the valid secret on the app
                        elseif ($existingAuth.ExpiryDate -ne $validSecret.EndDateTime.ToString('yyyy-MM-dd')) {
                            Write-Host "`n   ⚠  auth.json contains an outdated secret!" -ForegroundColor Yellow
                            Write-Host "     auth.json expiry: $($existingAuth.ExpiryDate)" -ForegroundColor Gray
                            Write-Host "     Valid secret expiry: $($validSecret.EndDateTime.ToString('yyyy-MM-dd'))" -ForegroundColor Gray
                            Write-Host "     The secret in auth.json does not match the current valid secret" -ForegroundColor Yellow
                            
                            $response = Read-Host "`n     Create new secret to replace the outdated one? (Y/N)"
                            if ($response -eq 'Y' -or $response -eq 'y') {
                                $reuseSecret = $false
                            } else {
                                Write-Host "`n   ✗ Cannot proceed with outdated secret. Exiting." -ForegroundColor Red
                                Disconnect-MgGraph | Out-Null
                                exit 1
                            }
                        }
                        else {
                            Write-Host "     auth.json exists with matching secret - will be reused" -ForegroundColor Green
                            $existingSecret = $validSecret
                            $reuseSecret = $true
                        }
                    }
                    catch {
                        Write-Host "`n   ⚠  auth.json exists but could not be parsed: $_" -ForegroundColor Yellow
                        Write-Host "     A new secret will be created" -ForegroundColor Cyan
                        $reuseSecret = $false
                    }
                } else {
                    Write-Host "`n   ⚠  auth.json not found!" -ForegroundColor Yellow
                    Write-Host "     Cannot retrieve secret value from Azure (it's write-only)" -ForegroundColor Yellow
                    Write-Host "     You must create a new secret or restore auth.json from backup`n" -ForegroundColor Yellow
                    
                    $response = Read-Host "     Create new secret? (Y/N)"
                    if ($response -eq 'Y' -or $response -eq 'y') {
                        $reuseSecret = $false
                    } else {
                        Write-Host "`n   ✗ Cannot proceed without secret. Exiting." -ForegroundColor Red
                        Disconnect-MgGraph | Out-Null
                        exit 1
                    }
                }
            } else {
                Write-Host "   ⚠  Existing secrets found but all are expired" -ForegroundColor Yellow
                Write-Host "     New secret will be created" -ForegroundColor Cyan
            }
        } else {
            Write-Host "   No existing secrets found - new secret will be created" -ForegroundColor Gray
        }
        
        # Update permissions
        Write-Host "`n   Updating permissions to match $mode mode..." -ForegroundColor Cyan
        
        $resourceAccess = @()
        foreach ($perm in $selectedPermissions) {
            $resourceAccess += @{
                Id = $perm.Id
                Type = $perm.Type
            }
        }
        
        $params = @{
            RequiredResourceAccess = @(
                @{
                    ResourceAppId = $microsoftGraphResourceId
                    ResourceAccess = $resourceAccess
                }
            )
        }
        
        try {
            Update-MgApplication -ApplicationId $app.Id -BodyParameter $params -ErrorAction Stop
            Write-Host "   ✓ Permissions updated" -ForegroundColor Green
        }
        catch {
            Write-Host "   ✗ Failed to update permissions" -ForegroundColor Red
            if ($_.Exception.Message -match "403|Forbidden|Insufficient privileges") {
                Write-Host "     Error: Insufficient privileges to update application permissions" -ForegroundColor Red
                Write-Host "     Required: Application.ReadWrite.All or Application.ReadWrite.OwnedBy" -ForegroundColor Yellow
                Write-Host "     Action: Ask your Global Administrator to:" -ForegroundColor Yellow
                Write-Host "       1. Grant you 'Application Administrator' role, OR" -ForegroundColor Yellow
                Write-Host "       2. Update the app permissions manually in Azure Portal" -ForegroundColor Yellow
                Write-Host "          https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($app.AppId)" -ForegroundColor Cyan
            }
            else {
                Write-Host "     Error: $_" -ForegroundColor Red
            }
            throw "Failed to update application permissions. Please resolve the error above and try again."
        }
    }
    else {
        # Create new application
        Write-Host "`nCreating new application: $appDisplayName..." -ForegroundColor Cyan
        
        $resourceAccess = @()
        foreach ($perm in $selectedPermissions) {
            $resourceAccess += @{
                Id = $perm.Id
                Type = $perm.Type
            }
        }
        
        $params = @{
            DisplayName = $appDisplayName
            SignInAudience = "AzureADMyOrg"
            RequiredResourceAccess = @(
                @{
                    ResourceAppId = $microsoftGraphResourceId
                    ResourceAccess = $resourceAccess
                }
            )
        }
        
        $app = New-MgApplication @params
        Write-Host "   ✓ Created application: $appDisplayName" -ForegroundColor Green
        Write-Host "     Application ID: $($app.AppId)" -ForegroundColor Gray
        Write-Host "     Object ID: $($app.Id)" -ForegroundColor Gray
    }
}
catch {
    Write-Host "`n   ✗ Failed to create/update application" -ForegroundColor Red
    Write-Host "     Error: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit 1
}

# Ensure current user is set as owner of the app registration
try {
    $currentUser = Get-MgUser -UserId $context.Account -ErrorAction Stop
    $existingOwners = @(Get-MgApplicationOwner -ApplicationId $app.Id -ErrorAction SilentlyContinue)
    $isAlreadyOwner = $existingOwners | Where-Object { $_.Id -eq $currentUser.Id }
    
    if ($isAlreadyOwner) {
        Write-Host "`n   ✓ Current user ($($context.Account)) is already owner of the app registration" -ForegroundColor Gray
    }
    else {
        $ownerParams = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($currentUser.Id)"
        }
        New-MgApplicationOwnerByRef -ApplicationId $app.Id -BodyParameter $ownerParams -ErrorAction Stop
        Write-Host "`n   ✓ Added current user ($($context.Account)) as owner of app registration" -ForegroundColor Green
    }
}
catch {
    Write-Host "`n   ⚠ Could not set app registration owner: $_" -ForegroundColor Yellow
    Write-Host "     The app registration will still work, but may not have an explicit owner." -ForegroundColor Yellow
}

# ============================================================================
# STEP 2: CREATE SERVICE PRINCIPAL
# ============================================================================

Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  STEP 2: Service Principal                                   ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

try {
    # Check if service principal already exists
    $existingSP = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue | Select-Object -First 1
    
    if ($existingSP) {
        $sp = $existingSP
        Write-Host "`n⚠  Service Principal already exists" -ForegroundColor Yellow
        Write-Host "   Service Principal ID: $($sp.Id)" -ForegroundColor Gray
    }
    else {
        Write-Host "`nCreating Service Principal..." -ForegroundColor Cyan
        $sp = New-MgServicePrincipal -AppId $app.AppId
        Write-Host "   ✓ Service Principal created" -ForegroundColor Green
        Write-Host "     Service Principal ID: $($sp.Id)" -ForegroundColor Gray
    }
}
catch {
    Write-Host "`n   ✗ Failed to create Service Principal" -ForegroundColor Red
    Write-Host "     Error: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit 1
}

# ============================================================================
# STEP 3: GRANT ADMIN CONSENT
# ============================================================================

Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  STEP 3: Admin Consent                                       ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

Write-Host "`nGranting admin consent for permissions..." -ForegroundColor Cyan

try {
    $microsoftGraphSP = Get-MgServicePrincipal -Filter "appId eq '$microsoftGraphResourceId'" | Select-Object -First 1
    
    # Get existing role assignments
    $existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All -ErrorAction SilentlyContinue
    
    $grantedCount = 0
    $skippedCount = 0
    
    foreach ($perm in $selectedPermissions) {
        # Check if already granted
        $alreadyGranted = $existingAssignments | Where-Object { $_.AppRoleId -eq $perm.Id }
        
        if ($alreadyGranted) {
            Write-Host "   ○ $($perm.Name) - already granted" -ForegroundColor Gray
            $skippedCount++
        }
        else {
            try {
                $params = @{
                    PrincipalId = $sp.Id
                    ResourceId = $microsoftGraphSP.Id
                    AppRoleId = $perm.Id
                }
                
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -BodyParameter $params -ErrorAction Stop | Out-Null
                Write-Host "   ✓ $($perm.Name) - granted" -ForegroundColor Green
                $grantedCount++
            }
            catch {
                if ($_.Exception.Message -match "403|Forbidden|Insufficient privileges") {
                    Write-Host "   ✗ $($perm.Name) - insufficient privileges" -ForegroundColor Red
                }
                else {
                    Write-Host "   ✗ $($perm.Name) - failed: $_" -ForegroundColor Red
                }
            }
        }
    }
    
    Write-Host "`n   Summary: $grantedCount granted, $skippedCount already existed" -ForegroundColor Cyan
    
    # Check if any permissions failed due to insufficient privileges
    $failedCount = $selectedPermissions.Count - $grantedCount - $skippedCount
    if ($failedCount -gt 0) {
        Write-Host "`n   ⚠ $failedCount permission(s) failed to grant" -ForegroundColor Yellow
        Write-Host "     Required: AppRoleAssignment.ReadWrite.All or RoleManagement.ReadWrite.Directory" -ForegroundColor Yellow
        Write-Host "     Action: Ask your Global Administrator to:" -ForegroundColor Yellow
        Write-Host "       1. Grant you 'Privileged Role Administrator' role, OR" -ForegroundColor Yellow
        Write-Host "       2. Grant admin consent manually in Azure Portal:" -ForegroundColor Yellow
        Write-Host "          https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/CallAnAPI/appId/$($app.AppId)" -ForegroundColor Cyan
        Write-Host "       3. Then re-run this script" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "`n   ✗ Failed to grant admin consent" -ForegroundColor Red
    Write-Host "     Error: $_" -ForegroundColor Red
    Write-Host "     You may need to grant consent manually in Azure Portal" -ForegroundColor Yellow
}

# ============================================================================
# STEP 4: MANAGE CLIENT SECRET
# ============================================================================

Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  STEP 4: Client Secret                                       ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

$secret = $null
$secretValue = $null

try {
    if ($reuseSecret) {
        Write-Host "`n✓ Reusing existing valid secret" -ForegroundColor Green
        Write-Host "  Expires: $($existingSecret.EndDateTime.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor Gray
        Write-Host "  KeyId: $($existingSecret.KeyId)" -ForegroundColor Gray
        
        # Load existing secret value from auth.json
        $authJsonPath = Join-Path $PSScriptRoot "auth.json"
        $existingAuth = Get-Content $authJsonPath -Raw | ConvertFrom-Json
        $secretValue = $existingAuth.ClientSecret
        $secret = $existingSecret
    }
    else {
        Write-Host "`nCreating new client secret..." -ForegroundColor Cyan
        Write-Host "  Validity: 30 days" -ForegroundColor Gray
        
        $passwordCredential = @{
            DisplayName = "Automation Secret - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
            EndDateTime = (Get-Date).AddDays(30)
        }
        
        $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential $passwordCredential
        $secretValue = $secret.SecretText
        
        Write-Host "  ✓ New client secret created" -ForegroundColor Green
        Write-Host "    Expires: $($secret.EndDateTime.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor Gray
    }
}
catch {
    Write-Host "`n  ✗ Failed to manage client secret" -ForegroundColor Red
    Write-Host "    Error: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null
    exit 1
}

# ============================================================================
# STEP 5: EXPORT TO AUTH.JSON
# ============================================================================

Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  STEP 5: Export Credentials                                  ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

try {
    $authJsonPath = Join-Path $PSScriptRoot "auth.json"
    
    $authData = @{
        TenantId = $context.TenantId
        ClientId = $app.AppId
        ClientSecret = $secretValue
        ExpiryDate = $secret.EndDateTime.ToString('yyyy-MM-dd')
        Mode = if ($AssessmentMode) { "Assessment" } else { "Full" }
        LastUpdated = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        SecretStatus = if ($reuseSecret) { "Reused" } else { "NewlyCreated" }
    }
    
    $authData | ConvertTo-Json -Depth 10 | Out-File -FilePath $authJsonPath -Encoding UTF8 -Force
    
    Write-Host "`n✓ Credentials exported successfully" -ForegroundColor Green
    Write-Host "  Location: $authJsonPath" -ForegroundColor Gray
    Write-Host "  Mode: $(if ($AssessmentMode) { 'Assessment (Read-Only)' } else { 'Full (Read-Write)' })" -ForegroundColor Gray
    
    # ============================================================================
    # AUTO-CREATE PASSWORD-PROTECTED ZIP
    # ============================================================================
    
    if ($reuseSecret) {
        # Secret was reused — auth.zip was already shared previously, no need to re-create
        Write-Host "`n✓ Skipping ZIP creation (existing secret reused, auth.zip was previously shared)" -ForegroundColor Green
        Write-Host "  auth.json updated in-place with current mode/permissions" -ForegroundColor Gray
        
        # Clean up auth.json if auth.zip already exists (credentials are in the existing ZIP)
        $zipPath = Join-Path $PSScriptRoot "auth.zip"
        if (Test-Path $zipPath) {
            try {
                Remove-Item $authJsonPath -Force
                Write-Host "  ✓ Unencrypted auth.json deleted (existing auth.zip still valid)" -ForegroundColor Green
            }
            catch {
                Write-Host "  ⚠ Could not delete auth.json - please delete it manually" -ForegroundColor Yellow
            }
        }
    }
    else {
    Write-Host "`nCreating password-protected ZIP file..." -ForegroundColor Cyan
    
    $zipPath = Join-Path $PSScriptRoot "auth.zip"
    # Password stored as Base64 to avoid plain text in code
    $password = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("T25lVmlubjRldmVyIw=="))
    $zipCreated = $false
    
    # Remove existing ZIP if present
    if (Test-Path $zipPath) {
        Remove-Item $zipPath -Force
    }
    
    # Detect OS and use appropriate method
    if ($IsWindows -or $env:OS -match "Windows") {
        # Windows: Try 7-Zip first, then provide fallback instructions
        $sevenZipPaths = @(
            "$env:ProgramFiles\7-Zip\7z.exe",
            "${env:ProgramFiles(x86)}\7-Zip\7z.exe",
            "C:\Program Files\7-Zip\7z.exe"
        )
        
        $sevenZip = $sevenZipPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
        
        if ($sevenZip) {
            # Use -w flag to set working directory, then use relative filename
            $arguments = "a", "-p$password", "-tzip", "-mem=AES256", "-w$PSScriptRoot", $zipPath, "auth.json"
            $process = Start-Process -FilePath $sevenZip -ArgumentList $arguments -NoNewWindow -Wait -PassThru -WorkingDirectory $PSScriptRoot
            if ($process.ExitCode -eq 0) {
                $zipCreated = $true
                Write-Host "  ✓ Password-protected ZIP created using 7-Zip" -ForegroundColor Green
            }
        }
        else {
            # Fallback: Use Compress-Archive (no password protection available)
            Write-Host "  ⚠ 7-Zip not found. Creating ZIP without password protection..." -ForegroundColor Yellow
            try {
                Compress-Archive -Path $authJsonPath -DestinationPath $zipPath -Force
                if (Test-Path $zipPath) {
                    $zipCreated = $true
                    Write-Host "  ✓ ZIP created (without password protection)" -ForegroundColor Green
                    Write-Host "  → For password-protected ZIP, install 7-Zip: https://www.7-zip.org/" -ForegroundColor Gray
                }
            }
            catch {
                Write-Host "  ⚠ Failed to create ZIP: $_" -ForegroundColor Yellow
                Write-Host "  → Contact consultant for manual ZIP creation instructions" -ForegroundColor Cyan
            }
        }
    }
    elseif ($IsMacOS -or $IsLinux -or (Test-Path "/usr/bin/zip")) {
        # macOS/Linux: Use built-in zip command with -j flag to junk paths (store just filename)
        try {
            # Change to script directory and use relative filename
            Push-Location $PSScriptRoot
            $zipCommand = "zip"
            $arguments = "-j", "-e", "-P", $password, "auth.zip", "auth.json"
            $process = Start-Process -FilePath $zipCommand -ArgumentList $arguments -NoNewWindow -Wait -PassThru -RedirectStandardError "/dev/null"
            Pop-Location
            if ($process.ExitCode -eq 0 -and (Test-Path $zipPath)) {
                $zipCreated = $true
                Write-Host "  ✓ Password-protected ZIP created" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "  ⚠ Failed to create ZIP automatically: $_" -ForegroundColor Yellow
            Write-Host "  → Contact consultant for manual ZIP creation instructions" -ForegroundColor Cyan
        }
    }
    
    if ($zipCreated) {
        Write-Host "  Location: $zipPath" -ForegroundColor Gray
        Write-Host "\n  ✓ Ready to share with consultant" -ForegroundColor Green
        
        # Delete unencrypted auth.json for security
        try {
            Remove-Item $authJsonPath -Force
            Write-Host "  ✓ Unencrypted auth.json deleted (kept only secure ZIP)" -ForegroundColor Green
        }
        catch {
            Write-Host "  ⚠ Could not delete auth.json - please delete it manually" -ForegroundColor Yellow
        }
    }
    } # end else (new secret - create ZIP)
}
catch {
    Write-Host "`n✗ Failed to export credentials" -ForegroundColor Red
    Write-Host "  Error: $_" -ForegroundColor Red
}

# ============================================================================
# SUMMARY
# ============================================================================

Disconnect-MgGraph | Out-Null

Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║  Setup Complete - $mode Mode$((' ' * [Math]::Max(0, 18 - $mode.Length))) ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Green

Write-Host "`n📋 Application Details:" -ForegroundColor Cyan
Write-Host "   • Application Name: $appDisplayName" -ForegroundColor White
Write-Host "   • Application ID: $($app.AppId)" -ForegroundColor White
Write-Host "   • Tenant ID: $($context.TenantId)" -ForegroundColor White
Write-Host "   • Mode: $(if ($AssessmentMode) { 'Assessment (Read-Only)' } else { 'Full (Read-Write)' })" -ForegroundColor White
Write-Host "   • Secret Expiry: $($secret.EndDateTime.ToString('yyyy-MM-dd'))" -ForegroundColor White
Write-Host "   • Secret Status: $(if ($reuseSecret) { 'Reused Existing' } else { 'Newly Created' })" -ForegroundColor White

Write-Host "`n🔐 Permissions Granted:" -ForegroundColor Cyan
foreach ($perm in $selectedPermissions) {
    Write-Host "   ✓ $($perm.Name)" -ForegroundColor Gray
}

if ($AssessmentMode) {
    Write-Host "`n📊 Next Steps (Assessment Mode):" -ForegroundColor Cyan
    Write-Host "   1. Email the auth.zip file to your consultant:" -ForegroundColor White
    Write-Host "      • File location: $(Join-Path $PSScriptRoot 'auth.zip')" -ForegroundColor Green
    Write-Host "      • Send via email or OneDrive/SharePoint" -ForegroundColor Gray
    Write-Host "      • No password needed - consultant will extract the file" -ForegroundColor Gray
    Write-Host "\n   2. Consultant runs assessments and provides reports" -ForegroundColor White
    Write-Host "\n   3. Review assessment reports from consultant" -ForegroundColor White
    Write-Host "\n   4. When ready for enrollment, upgrade permissions:" -ForegroundColor Yellow
    Write-Host "      • Run: .\CreateServicePrincipal.ps1" -ForegroundColor Yellow
    Write-Host "        (without -AssessmentMode flag)" -ForegroundColor Gray
    Write-Host "      • Email the new auth.zip to consultant" -ForegroundColor Gray
    Write-Host "\n   5. After engagement (IMPORTANT):" -ForegroundColor Yellow
    Write-Host "      • Delete auth.zip from your machine" -ForegroundColor Red
    Write-Host "      • Consultant will delete their copy" -ForegroundColor Gray
    Write-Host "      • Delete service principal from Azure AD (optional)" -ForegroundColor Gray
    Write-Host "      • Delete any email/cloud copies" -ForegroundColor Gray
} else {
    Write-Host "`n🚀 Next Steps (Full Mode):" -ForegroundColor Cyan
    Write-Host "   1. Email the auth.zip file to your consultant:" -ForegroundColor White
    Write-Host "      • File location: $(Join-Path $PSScriptRoot 'auth.zip')" -ForegroundColor Green
    Write-Host "      • Send via email or OneDrive/SharePoint" -ForegroundColor Gray
    Write-Host "      • No password needed - consultant will extract the file" -ForegroundColor Gray
    Write-Host "\n   2. Consultant performs enrollment and enforcement" -ForegroundColor White
    Write-Host "`n   3. Monitor secret expiration:" -ForegroundColor White
    Write-Host "      • Secret expires: $($secret.EndDateTime.ToString('yyyy-MM-dd'))" -ForegroundColor Gray
    Write-Host "      • Regenerate before expiry if needed" -ForegroundColor Gray
    Write-Host "`n   4. After engagement (IMPORTANT):" -ForegroundColor Yellow
    Write-Host "      • Delete auth.zip from your machine" -ForegroundColor Red
    Write-Host "      • Consultant will delete their copy" -ForegroundColor Gray
    Write-Host "      • Delete service principal from Azure AD (optional)" -ForegroundColor Gray
    Write-Host "      • Delete any email/cloud copies" -ForegroundColor Gray
}

if (-not $reuseSecret) {
Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
Write-Host "║  SECURITY BEST PRACTICES                                     ║" -ForegroundColor Yellow
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow

Write-Host "`n📁 Automated Security:" -ForegroundColor Cyan
Write-Host "   ✓ auth.zip created automatically with password protection" -ForegroundColor White
Write-Host "   ✓ Unencrypted auth.json deleted (only ZIP remains)" -ForegroundColor White
Write-Host "   ✓ 30-day credential expiration" -ForegroundColor White
Write-Host "   ✓ Never commit to Git/source control (.gitignore included)" -ForegroundColor White

Write-Host "`n🔐 Secure Sharing Method:" -ForegroundColor Cyan
Write-Host "   1. Email auth.zip to consultant:" -ForegroundColor White
Write-Host "      • Use email or OneDrive/SharePoint with 7-day link expiration" -ForegroundColor Gray
Write-Host "      • File is password-protected (consultant has password)" -ForegroundColor Gray
Write-Host "   2. Consultant extracts and uses credentials" -ForegroundColor White
Write-Host "      • No customer action required" -ForegroundColor Gray

Write-Host "`n⏰ Time-Limited Access:" -ForegroundColor Cyan
Write-Host "   • Secret valid for 30 days only" -ForegroundColor White
Write-Host "   • Typical engagement: 1-4 weeks" -ForegroundColor Gray
Write-Host "   • Automatic expiration adds security layer" -ForegroundColor Gray

Write-Host "`n🗑️  Cleanup Checklist (After Engagement):" -ForegroundColor Cyan
Write-Host "   ☐ Delete auth.zip from your machine" -ForegroundColor White
Write-Host "   ☐ Remove from email/OneDrive/SharePoint" -ForegroundColor White
Write-Host "   ☐ Empty Recycle Bin / Deleted Items" -ForegroundColor White
Write-Host "   ☐ (Optional) Delete service principal from Azure AD:" -ForegroundColor Gray
Write-Host "      Azure Portal > App registrations > $appDisplayName > Delete" -ForegroundColor Gray
}
