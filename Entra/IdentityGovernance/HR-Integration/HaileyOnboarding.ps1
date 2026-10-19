# Created by Johan Öberg

<#
.SYNOPSIS
    Automates employee onboarding by syncing data from Hailey HR API with Microsoft 365 and Exchange Online.

.DESCRIPTION
    This PowerShell script integrates with the Hailey HR system to automate employee onboarding tasks.
    It includes functions for:
    - Connecting to Microsoft Graph and Exchange Online (supports both Azure Automation and local execution)
    - Fetching employee data from Hailey HR API
    - Setting team assignments in Exchange Online based on Hailey team IDs
    - Managing company name attributes for users
    - Handling employee visibility in address lists based on joining dates
    - Processing both "WillJoin" and "InService" employees
    
    The script uses managed identity when running in Azure Automation and interactive authentication
    when running locally. Team assignments are mapped from Hailey team IDs to Exchange CustomAttribute1.
    Company name is set within a 14-day window before the employee's joining date.

.NOTES
    File Name      : run-HaileyOnboarding.ps1
    Author         : Johan Öberg
    Prerequisite   : Microsoft.Graph, ExchangeOnlineManagement modules
    Version        : 1.0
    
.EXAMPLE
    .\run-HaileyOnboarding.ps1
    Runs the script to process employee onboarding from Hailey HR system.
#>

# Connect using Managed Identity
    Write-Output "🔐 Connecting to Microsoft Graph..."
    if ($env:AZUREPS_HOST_ENVIRONMENT -or $env:AUTOMATION_ASSET_ACCOUNTID) {
        Write-Output "🤖 Detected Azure Automation environment - using Managed Identity"
        Connect-MgGraph -Identity -NoWelcome
        Connect-ExchangeOnline -ManagedIdentity -Organization .onmicrosoft.com
    }
    else {
        Write-Output "💻 Running locally - using interactive authentication"
        Connect-MgGraph -Scopes "Device.ReadWrite.All" -NoWelcome
        Connect-ExchangeOnline -Organization .onmicrosoft.com
    }
    Write-Output "✅ Successfully connected to Microsoft Graph and Exchange Online"

# Variables
# This date sets the team assignment
#$pastDate = (Get-Date).AddDays(-2).ToString('yyyy-MM-dd')

# This date is for adding Company Name 1 day before Team so that E5 licens can be deployed and Mailbox created to use Set-Mailbox for Team assignment
#$pastDateCompany = (Get-Date).AddDays(-3).ToString('yyyy-MM-dd')

$CompanyName = "" #Add your Company Name here

$TeamMapping = @{
    #"ID" = "TeamName" --- IGNORE ---
}

# Functions
function Set-Team {
    param (
        [Parameter(Mandatory=$true)]
        [array]$employees,
        [switch]$CheckJoiningDate
    )

    Write-Output "👥 Setting Team assignments..."

    foreach ($employee in $employees) {
        # Normalize teamIds to a string
        $normalizedTeamIds = ($employee.teamIds -join ",") -replace '[{}]', ''
        $pastDateCompany = ([datetime]$employee.dateOfJoining).AddDays(-2).ToString("yyyy-MM-dd")

        # Check if the employee's team ID matches any predefined team
        if ($TeamMapping.ContainsKey($normalizedTeamIds)) {
            $teamName = $TeamMapping[$normalizedTeamIds]

            if ($employee.CustomAttribute1 -eq $teamName) {
                Write-Output "✔️ Team already set to $teamName for $($employee.companyEmail)"
            } else {
                # If -CheckJoiningDate is enabled, only set the team if dateOfJoining matches pastDate
                if ($CheckJoiningDate) {
                    if ($employee.dateOfJoining -eq $pastDateCompany) {
                        Write-Output "🎯 Setting Team $teamName for $($employee.companyEmail) (Joining Date: $($employee.dateOfJoining))"
                        Set-Mailbox -Identity $employee.companyEmail -CustomAttribute1 $teamName
                    } else {
                        Write-Output "⏭️ Skipping $($employee.companyEmail) - Joining Date: $($employee.dateOfJoining) (Expected: $pastDateCompany)"
                    }
                } else {
                    Write-Output "🎯 Setting Team $teamName for $($employee.companyEmail) (Joining Date: $($employee.dateOfJoining))"
                    Set-Mailbox -Identity $employee.companyEmail -CustomAttribute1 $teamName
                }
            }
        } else {
            Write-Output "🔍 No matching team found for $($employee.companyEmail) (Team ID: $normalizedTeamIds)"
        }
    }
}

function Set-CompanyName {
    param (
        [Parameter(Mandatory=$true)]
        [array]$employees
    )

    Write-Output "🏢 Setting CompanyName..."
    $today = (Get-Date).Date

    foreach ($employee in $employees) {
        $joiningDate = ([datetime]$employee.dateOfJoining).Date
        $startWindow = $joiningDate.AddDays(-14)

        if ($employee.CompanyName -ne $CompanyName) {
            if ($today -ge $startWindow -and $today -lt $joiningDate) {
                Write-Output "📝 Updating CompanyName for $($employee.companyEmail) (Joining Date: $($employee.dateOfJoining))"
                Update-MgUser -UserId $employee.companyEmail -CompanyName $CompanyName
                Update-MgUser -UserId $employee.companyEmail -ShowInAddressList:$true
            } else {
                Write-Output "⏸️ Skipping $($employee.companyEmail) - Joining Date: $($employee.dateOfJoining) (Today not in 14-day window)"
                Update-MgUser -UserId $employee.companyEmail -ShowInAddressList:$false
            }
        } else {
            Write-Output "✔️ No update needed for $($employee.companyEmail) - CompanyName already set"
        }
    }
}


function Get-EmployeeDetails {
    param (
        [Parameter(Mandatory=$true)]
        [array]$HaileyAPIUsers
    )

    if ($HaileyAPIUsers.Count -eq 0) {
        throw "Error: The parameter 'HaileyAPIUsers' is empty. No users to process."
    }

    $entrausers = @()

    foreach ($user in $HaileyAPIUsers) {
        # Fetch user details from Microsoft Graph
        $mguser = Get-MgUser -UserId $user.companyEmail -Property UserPrincipalName,CompanyName -ErrorAction SilentlyContinue
        
        # Attempt to fetch mailbox to get CustomAttribute1
        $mailbox = Get-Mailbox -Identity $user.CompanyEmail -ErrorAction SilentlyContinue
        
        if ($mailbox) {
            $mguser | Add-Member -NotePropertyName CustomAttribute1 -NotePropertyValue $mailbox.CustomAttribute1 -Force
        } else {
            $mguser | Add-Member -NotePropertyName CustomAttribute1 -NotePropertyValue $null -Force
        }
        $entrausers += [pscustomobject]@{
            companyEmail      = $user.companyEmail
            employmentStatus  = $user.employmentStatus
            accountStatus     = $user.accountStatus
            teamIds           = ($user.teamIds -join ",") -replace '[{}]', ''  # Normalize teamIds
            dateOfJoining     = [datetime]$user.dateOfJoining
            UserPrincipalName = $mguser.UserPrincipalName
            CompanyName       = $mguser.CompanyName
            CustomAttribute1  = $mguser.CustomAttribute1
        }
    }

    return $entrausers
}

#Hailey Stuff
Write-Output "🔑 Retrieving Hailey API credentials..."
$credential = Get-AutomationPSCredential -Name 'HaileyAPIToken'
#Accesstoken updated 2025-12-16 available until 2026-12-16
$token = $credential.GetNetworkCredential().Password

$headers = @{
    'accept' = 'application/json'
    'Authorization' = "Bearer $token"
}

Write-Output "📡 Fetching employees from Hailey API..."
$employees = Invoke-RestMethod -Uri 'https://api.haileyhr.app/Employees' -Headers $headers -Method Get

Write-Output "🔎 Filtering employees..."
$employeesFiltered = $employees | Select-Object employeeId, companyEmail, firstName, lastName, dateOfJoining, employmentStatus, accountStatus, teamIds | Where-Object { $_.employmentStatus -eq "WillJoin" -or $_.employmentStatus -eq "InService" }

$WillJoin = $employeesFiltered | Where-Object { $_.employmentStatus -eq "WillJoin" -and $_.companyEmail -like "" }
$InService = $employeesFiltered | Where-Object { $_.employmentStatus -eq "InService" }

Write-Output "📋 Processing InService employees..."
try {
    $entrausers_inservice = Get-EmployeeDetails -HaileyAPIUsers $InService
} catch {
    Write-Error "❌ An error occurred: $_"
}

Write-Output "📋 Processing WillJoin employees..."
try {
    $entrausers_willjoin = Get-EmployeeDetails -HaileyAPIUsers $WillJoin
} catch {
    Write-Error "❌ An error occurred: $_"
}

# Apply Team Assignments
Write-Output "`n🚀 Starting team assignment process..."
if ($entrausers_willjoin){
    Set-Team -employees $entrausers_willjoin -CheckJoiningDate
}
else{
    Write-Output "😴 No new employees starting soon"
}

if ($entrausers_inservice){
    Set-Team -employees $entrausers_inservice
}

Write-Output "`n🏢 Starting company name assignment process..."
if ($entrausers_willjoin){
    Set-CompanyName -employees $entrausers_willjoin
}

Write-Output "`n🎉 Script completed successfully!"
