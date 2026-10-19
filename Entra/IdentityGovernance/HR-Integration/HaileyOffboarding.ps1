<#
.SYNOPSIS
    Automated Employee Offboarding Integration Script
    
.DESCRIPTION
    This script automates the offboarding process for employees managed in Hailey HR system.
    It performs the following operations:
    
    📋 Main Functions:
    1. 🔗 Connects to Microsoft Graph API (using Managed Identity in Azure or interactive auth locally)
    2. 📞 Fetches leaving employees from Hailey HR API
    3. 📅 Updates EmployeeLeaveDateTime in Microsoft Entra ID user accounts
    4. 🧹 Clears custom attributes (extensionAttribute1) on the employee's last day
    5. 📊 Generates DLP (Data Loss Prevention) configuration CSV for Microsoft Purview
    6. ☁️ Uploads DLP configuration to Azure Storage
    7. 📤 Pushes HR connector data to Microsoft 365 Compliance Center
    
    🔐 Required Permissions:
    - Microsoft Graph: User.Read.All, User-LifeCycleInfo.ReadWrite.All
    - Azure: Storage Account access via Managed Identity
    - Hailey API: Valid API token stored in Azure Automation
    
    🤖 Designed to run in:
    - Azure Automation (using Managed Identity)
    - Local environment (interactive authentication)
    
.NOTES
    Created by: Johan Öberg
    Last Updated: January 2026
    Version: 2.0
#>


# Connect using Managed Identity
    Write-Output "🔐 Connecting to Microsoft Graph..."
    if ($env:AZUREPS_HOST_ENVIRONMENT -or $env:AUTOMATION_ASSET_ACCOUNTID) {
        Write-Output "🤖 Detected Azure Automation environment - using Managed Identity"
        Connect-MgGraph -Identity -NoWelcome
    }
    else {
        Write-Output "💻 Running locally - using interactive authentication"
        Connect-MgGraph -Scopes "User.Read.All", "User-LifeCycleInfo.ReadWrite.All", "Group.Read.All", "RoleManagement.Read.Directory" -NoWelcome
    }
    Write-Output "✅ Successfully connected to Microsoft Graph"

#Variables
$Date = (Get-Date).AddDays(1).ToString('yyyy-MM-dd')
$storageaccountname = "" #Add your Storage Account Name here for DLP config upload

# ═══════════════════════════════════════════════════════════════
# 🎯 HAILEY HR API INTEGRATION
# ═══════════════════════════════════════════════════════════════
Write-Output ""
Write-Output "📡 Fetching employee data from Hailey HR..."

if ($env:AZUREPS_HOST_ENVIRONMENT -or $env:AUTOMATION_ASSET_ACCOUNTID) {
    Write-Output "🤖 Running in Azure Automation - retrieving token from credential store"
    $credential = Get-AutomationPSCredential -Name 'HaileyAPIToken'
    $token = $credential.GetNetworkCredential().Password
}
else {
    Write-Output "💻 Running locally - please enter Hailey API token"
    $token = Read-Host -Prompt "Enter Hailey API Token" -AsSecureString | ConvertFrom-SecureString -AsPlainText
}

$headers = @{
    'accept' = 'application/json'
    'Authorization' = "Bearer $token"
}

$employees = Invoke-RestMethod -Uri 'https://api.haileyhr.app/Employees' -Headers $headers -Method Get
Write-Output "✅ Successfully retrieved employee data from Hailey"

$leavingemployees = $employees | Select-Object employeeId, companyEmail, firstName, lastName, dateOfJoining, employmentStatus, employmentType, lastDayOfEmployment | Where-Object { $_.employmentStatus -eq "Leaving" -and $_.employmentType -ne "ProbationaryPeriod" }

Write-Output ""
Write-Output "👋 Found $($leavingemployees.Count) employees leaving the company:"
Write-Output $leavingemployees.companyEmail

# ═══════════════════════════════════════════════════════════════
# � CHECK FOR BLOCKING ISSUES (ROLES & GROUP OWNERSHIP)
# ═══════════════════════════════════════════════════════════════
Write-Output ""
Write-Output "🔍 Checking for accounts with blocking issues (Entra roles or sole group ownership)..."

$accountsWithIssues = @()

foreach ($emp in $leavingemployees) {
    $issues = @()
    
    try {
        # Get user ID first
        $userObj = Get-MgUser -UserId $emp.companyEmail -Property Id
        $userId = $userObj.Id
        
        # Check for active PIM role assignments
        $activeRoles = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$filter=principalId eq '$userId'&`$expand=roleDefinition"
        
        # Check for eligible PIM role assignments
        $eligibleRoles = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/roleManagement/directory/roleEligibilitySchedules?`$filter=principalId eq '$userId'&`$expand=roleDefinition"
        
        $roleDetails = @()
        
        if ($activeRoles.value.Count -gt 0) {
            foreach ($role in $activeRoles.value) {
                $roleDetails += "Active: $($role.roleDefinition.displayName)"
            }
        }
        
        if ($eligibleRoles.value.Count -gt 0) {
            foreach ($role in $eligibleRoles.value) {
                $roleDetails += "Eligible: $($role.roleDefinition.displayName)"
            }
        }
        
        if ($roleDetails.Count -gt 0) {
            $issues += "Has $($roleDetails.Count) Entra role(s): $($roleDetails -join ', ')"
        }
    } catch {
        Write-Output "  ⚠️ Could not check role assignments for $($emp.companyEmail): $_"
    }
    
    # Check for group ownership
    try {
        $ownedGroups = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($emp.companyEmail)/ownedObjects"
        $groups = $ownedGroups.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }
        
        if ($groups.Count -gt 0) {
            $soleOwnerGroups = @()
            
            foreach ($group in $groups) {
                $owners = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/owners"
                if ($owners.value.Count -eq 1) {
                    $soleOwnerGroups += $group.displayName
                }
            }
            
            if ($soleOwnerGroups.Count -gt 0) {
                $issues += "Sole owner of $($soleOwnerGroups.Count) group(s): $($soleOwnerGroups -join ', ')"
            }
        }
    } catch {
        Write-Output "  ⚠️ Could not check group ownership for $($emp.companyEmail): $_"
    }
    
    if ($issues.Count -gt 0) {
        $accountsWithIssues += [PSCustomObject]@{
            Email = $emp.companyEmail
            Name = "$($emp.firstName) $($emp.lastName)"
            LastDay = $emp.lastDayOfEmployment
            Issues = $issues
        }
    }
}

# Output summary of accounts with issues
if ($accountsWithIssues.Count -gt 0) {
    Write-Output ""
    Write-Output "════════════════════════════════════════════════════════════════"
    Write-Output "⚠️⚠️⚠️  CRITICAL: ACCOUNTS REQUIRING MANUAL ACTION  ⚠️⚠️⚠️"
    Write-Output "════════════════════════════════════════════════════════════════"
    Write-Output ""
    Write-Output "The following $($accountsWithIssues.Count) account(s) have blocking issues that will prevent"
    Write-Output "lifecycle workflow from completing successfully:"
    Write-Output ""
    
    foreach ($account in $accountsWithIssues) {
        Write-Output "❌ $($account.Name) ($($account.Email))"
        Write-Output "   📅 Last Day: $($account.LastDay)"
        foreach ($issue in $account.Issues) {
            Write-Output "   ⚠️  $issue"
        }
        Write-Output ""
    }
    
    Write-Output "════════════════════════════════════════════════════════════════"
    Write-Output "🔧 ACTION REQUIRED BY SUPPORT TEAM:"
    Write-Output "   1. Remove Entra role assignments from these users"
    Write-Output "   2. Add additional owners to groups or transfer ownership"
    Write-Output "   3. Re-run this script after resolving issues"
    Write-Output "════════════════════════════════════════════════════════════════"
    Write-Output ""
} else {
    Write-Output "✅ No blocking issues found - all accounts are ready for offboarding"
}

# ═══════════════════════════════════════════════════════════════
# �📅 UPDATE EMPLOYEE LEAVE DATETIME IN ENTRA ID
# ═══════════════════════════════════════════════════════════════
Write-Output ""
Write-Output "🔄 Processing employee leave dates in Entra ID..."
foreach ($emp in $leavingemployees) {
    # Retrieve the user object
    $user = Get-MgUser -UserId $emp.companyEmail -Property UserPrincipalName,EmployeeLeaveDateTime,extensionattribute2 | 
        Select-Object UserPrincipalName,EmployeeLeaveDateTime,extensionattribute2

    $leavedatetime = "$($emp.lastDayOfEmployment)T23:59:59Z"

    if ($null -eq $user.EmployeeLeaveDateTime -or $user.EmployeeLeaveDateTime -ne $leavedatetime) {
        Write-Output "  📝 Updating LeaveDateTime to $leavedatetime for $($emp.companyEmail)"
        Update-MgUser -UserId $user.UserPrincipalName -EmployeeLeaveDateTime $leavedatetime
    }
    else {
        Write-Output "  ✔️  User last date already set correctly for $($emp.companyEmail)"
    }

    if ($user.EmployeeLeaveDateTime -eq $Date){
        Write-Output "  🧹 Removing CustomAttribute for $($emp.companyEmail) (last day has passed)"
        Update-MgUser -UserId $user.UserPrincipalName -OnPremisesExtensionAttributes @{extensionAttribute1=""}
    }
}

# ═══════════════════════════════════════════════════════════════
# 📊 GENERATE DLP CONFIGURATION FOR MICROSOFT PURVIEW
# ═══════════════════════════════════════════════════════════════
Write-Output ""
Write-Output "📊 Generating DLP configuration for Microsoft Purview HR Connector..."

#Configure the DLP configuration for HR connector in Purview
$getdlpconfig = $leavingemployees | Select-Object CompanyEmail,lastDayOfEmployment,employmentStatus

$ldate = @()
$ldates = $getdlpconfig.lastDayOfEmployment 
foreach ($date in $ldates){
    $datetime = [datetime]::Parse($date)
    #format to ISO 8601 with milliseconds and offset
    $formattedldates = $datetime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffK")
    $ldate += $formattedldates
}

$rdate = @()
foreach ($date in $ldates) {
    $datetime = [datetime]::Parse($date)
    # Subtract 90 days and format to ISO 8601 with milliseconds and offset
    $formattedDate = $datetime.AddDays(-90).ToString("yyyy-MM-ddTHH:mm:ss.fffffffK")
    $rdate += $formattedDate
}

$companyEmails = $getdlpconfig.companyEmail

# Ensure the arrays have the same length
$dlpconfig = @()
for ($i = 0; $i -lt $companyEmails.Count; $i++) {
    $entry = [PSCustomObject]@{
        UserPrincipalName = $companyEmails[$i]
        ResignationDate   = $rdate[$i]
        LastWorkingDate   = $ldate[$i]
    }
    $dlpconfig += $entry
}

Write-Output "✅ DLP configuration prepared for $($dlpconfig.Count) users"
Write-Output $dlpconfig

# ═══════════════════════════════════════════════════════════════
# 💾 EXPORT TO CSV AND UPLOAD TO AZURE STORAGE
# ═══════════════════════════════════════════════════════════════
Write-Output ""
Write-Output "💾 Creating temporary CSV file..."
# Create a temporary CSV file path
$tempFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "TemporaryData.csv")

Write-Output $tempFilePath.GetType().FullName  # Should output "System.String"

# Export the content directly to a CSV file
$dlpconfig | Export-Csv -Path $tempFilePath -NoTypeInformation -Encoding UTF8

# Now, $tempFilePath holds the file path to the CSV file
Write-Output "✅ CSV file created: $tempFilePath"

# Variables for storage
$storageContainerName = "dlpconfig"
$blobName = "dlpconfig.csv"

# Authenticate to Azure using the Managed Identity
Write-Output ""
Write-Output "🔐 Connecting to Azure Storage..."
Connect-AzAccount -Identity -Subscription InfraStructure

$storageaccount = Get-AzStorageAccount -ResourceGroupName FileStorage -Name $storageaccountname
#$blobname = "dlpconfig/dlpconfig.csv"

# Upload the file to Azure Storage as a blob
Write-Output "☁️ Uploading CSV to Azure Storage ($storageContainerName/$blobName)..."
Set-AzStorageBlobContent -File $tempFilePath -Container $storageContainerName -Blob $blobName -Context $storageaccount.Context -Force
Write-Output "✅ Successfully uploaded to Azure Storage"

# ═══════════════════════════════════════════════════════════════
# 📤 PUSH DATA TO M365 COMPLIANCE CONNECTOR
# ═══════════════════════════════════════════════════════════════
Write-Output ""
Write-Output "📤 Preparing to push data to Microsoft 365 Compliance Connector..."

# Upload the CSV data to Azure Storage as a blob
#Write-Output "$dlpconfig"
#Write-Output "JUST GONNA SEND IT (to $storageaccountname)"

# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.
# 
# Sample Script to Push Data to M365 Compliance Connector 
# Examples
#
# Push data to M365 Compliance Connector:
# .\sample_script.ps1 -tenantId <Guid> -appId <App Id> -appSecret <App Secret> -jobId <Job id GUID> -filePath <File Path> -Verbose
<#
    .SYNOPSIS
        Sample Script to push Data to M365 Compliance Connector. Script Takes an Input file, chunks it to predefined size and does the data push
    .DESCRIPTION
        Take Input File. Use the metadata for checkpointing and then compute start point.
        Metadata has attributes corresponding to the last row for the file that was successfully sent. 
        If no metadata exists, create new metadata and start processing from start of File.
        If metadata exists start processing from point of last successful push.
        Create Temporary File Chunks of specified Line Count.
        Then fetch the access token and using the HttpClient generate a POST request
    .PARAMETER TenantId
        This is the Id for your Microsoft 365 organization. 
        This is used to identify your organization.
    .PARAMETER APPID
        This is the Azure AD application Id for the app that you created in Azure AD. 
        This is used by Azure AD for authentication when the script attempts to accesses your Microsoft 365 organization.
    .PARAMETER APPSECRET
        This is the Azure AD application secret for the app that you created in Azure AD. This also used for authentication.
    .PARAMETER JOBID
        The JobId retreived while setting up the Connector
    .PARAMETER FilePath
        This is the file path for the CSV file. Try to avoid spaces in the file path; otherwise use single quotation marks.
    .PARAMETER RecordsPerCall
        This is the chunk size to be used.
#>


$credential = Get-AutomationPSCredential -Name 'HR-ConnectorSecret'

#Variables
$tenantId = "bcee8bc2-b710-439d-b5bd-f5b83846ddee"
$appId = "2173c7d6-7d81-49ed-be56-4439dd8d47d6"
$appSecret = $credential.GetNetworkCredential().Password
$jobId = "a431061b-9417-4095-82fa-0ff823828414"
$FilePath = $tempFilePath
$RecordsPerCall = 50000
$retryTimeout = 60

# Access Token Config
$oAuthTokenEndpoint = "https://login.windows.net/$tenantId/oauth2/token"
$resource = 'https://microsoft.onmicrosoft.com/86dfdabb-5089-4a0c-880a-cfa5a790c5b1'
#old:4e476d41-2395-42be-89ff-34cb9186a1ac


# Csv upload config
$eventApiURl = "https://webhook.ingestion.office.com"
$eventApiEndpoint = "api/signals"

$serviceName = "PushConnector"
$TmpDirName = $env:TEMP + "ctr###D"

class FileMetdata {
    [string]$FileHash
    [string]$NoOfRowsWritten
    [string]$Service
    [string]$LastModTime
}

function Get-AccessToken () {
    # Token Authorization URI
    $uri = "$($oAuthTokenEndpoint)?api-version=1.0"

    # Access Token Body
    $formData = 
    @{
        client_id     = $appId;
        client_secret = $appSecret;
        grant_type    = 'client_credentials';
        resource      = $resource;
        tenant_id     = $tenantId;
    }

    # Parameters for Access Token call
    $params = 
    @{
        URI         = $uri
        Method      = 'Post'
        ContentType = 'application/x-www-form-urlencoded'
        Body        = $formData
    }

    $response = Invoke-RestMethod @params -ErrorAction Stop
    return $response.access_token
}

function RetryCommand {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter(Position = 1, Mandatory = $false)]
        [int]$Maximum = 15
    )

    Begin {
        $cnt = 0
    }

    Process {
        do {
            $cnt++
            try {
                $ScriptBlock.Invoke()
                return
            }
            catch {
                Write-Error $_.Exception.InnerException.Message -ErrorAction Continue
                write-output("Will retry in [{0}] seconds" -f $retryTimeout)
                Start-Sleep $retryTimeout
                if ($cnt -lt $Maximum) {
                    Write-Output "Retrying"
                }
            }
            
        } while ($cnt -lt $Maximum)

        throw 'Execution failed.'
    }
}

function Write-ErrorMessage($errorMessage) {
    $Exception = [Exception]::new($errorMessage)
    $ErrorRecord = [System.Management.Automation.ErrorRecord]::new(
        $Exception,
        "errorID",
        [System.Management.Automation.ErrorCategory]::NotSpecified,
        $TargetObject
    )
    $PSCmdlet.WriteError($ErrorRecord)
}

function Push-Data ($access_token, $FileName) {
    $nvCollection = [System.Web.HttpUtility]::ParseQueryString([String]::Empty) 
    $nvCollection.Add('jobid', $jobId)
    $uriRequest = [System.UriBuilder]"$eventApiURl/$eventApiEndpoint"
    $uriRequest.Query = $nvCollection.ToString()

    $fieldName = 'file'
    $url = $uriRequest.Uri.OriginalString

    Add-Type -AssemblyName 'System.Net.Http'

    $client = New-Object System.Net.Http.HttpClient
    $content = New-Object System.Net.Http.MultipartFormDataContent
    $fileStream = $null
    $fileContent = $null
	
    try {
		
        $fileStream = [System.IO.File]::OpenRead($FileName)
        $fileName = [System.IO.Path]::GetFileName($FileName)
        $fileContent = New-Object System.Net.Http.StreamContent($fileStream)
        $content.Add($fileContent, $fieldName, $fileName)
		
    }
    catch [System.IO.FileNotFoundException] {
        Write-Error "❌ CSV file not found"
        return
    }
    catch [System.IO.IOException] {
        Write-Error "❌ CSV file might be open"
        return
    }
    catch {
        Write-Error "❌ Error reading from CSV file"
        return
    }
	
    $client.DefaultRequestHeaders.Add("Authorization", "Bearer $access_token");
    $client.Timeout = New-Object System.TimeSpan(0, 0, 400)
    
    try {

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $result = $client.PostAsync($url, $content).Result
    }
    catch {
        Write-ErrorMessage("❌ Unknown failure while uploading")
        return
    }
    finally {
        # Properly dispose of streams and client to release file locks
        if ($null -ne $fileContent) { $fileContent.Dispose() }
        if ($null -ne $fileStream) { $fileStream.Dispose() }
        if ($null -ne $content) { $content.Dispose() }
        if ($null -ne $client) { $client.Dispose() }
    }
    
    $status_code = [int]$result.StatusCode
    if ($result.IsSuccessStatusCode) {
        Write-Output "✅ Upload Successful"
        $responseStr = $result.Content.ReadAsStringAsync().Result
        if (! [string]::IsNullOrWhiteSpace($responseStr)) {
            Write-Output("📋 Response: {0}" -f $responseStr)
        }
    }
    elseif ($status_code -eq 0 -or $status_code -eq 501 -or $status_code -eq 503) {
        throw "⚠️ Service unavailable"
    }
    else {
        $errorstring = "❌ Failure with StatusCode [{0}] and ReasonPhrase [{1}]" -f $result.StatusCode, $result.ReasonPhrase
        Write-ErrorMessage($errorstring)
        Write-ErrorMessage("Error body : {0}" -f $result.Content.ReadAsStringAsync().Result)
        throw $errorstring
    }
}

function GetOrCreateMetadata($FileName) {
    # Handle Obsolete Metadata
    HandleObsoleteMetadata
    $fileHash = ComputeHashForInputFile($FileName)
    $metaDataFileName = GetMetaDataFileName($fileHash)
    if ([System.IO.File]::Exists($metaDataFileName)) {
        # GET metadata from file
        $metadata = [FileMetdata](Get-Content $metaDataFileName | Out-String | ConvertFrom-Json)
        if ($metadata.FileHash -eq $fileHash) {
            #return Appropriate Metadata
            return $metadata
        }
    }
    
    $newmetadata = [FileMetdata]::new()
    $newmetadata.FileHash = $fileHash
    $newmetadata.NoOfRowsWritten = 0
    $newmetadata.LastModTime = Get-Date -format "yyyy-MM-ddTHH:mm:ss"
    $newmetadata.Service = $serviceName
    return $newmetadata
}

function GetMetaDataFileName($FileHash) {
    return $TmpDirName + "\." + $FileHash + "####mdata.txt"
}

function UpdateMetadata($FileName, $noOfRowsWritten) {
    # Update metadata
    $filemetaData = [FileMetdata]::new()
    $fileHash = ComputeHashForInputFile($FileName)
    $filemetaData.FileHash = $fileHash
    $filemetaData.Service = $serviceName
    $filemetaData.NoOfRowsWritten = $noOfRowsWritten
    $filemetaData.LastModTime = Get-Date -format "yyyy-MM-ddTHH:mm:ss"
    $metaDataFilePath = GetMetaDataFileName($fileHash)
    $filemetaData | ConvertTo-Json -Depth 100 | Out-File $metaDataFilePath
}

function HandleObsoleteMetadata() {
    # Delete metadata which are over 14 days old
    $timeLimit = (Get-Date).AddDays(-14)
    $filePath = $csvContent
    Get-ChildItem -Path $filePath -Force -Recurse | 
    Where-Object { !$_.PSIsContainer -and $_.LastWriteTime -lt $timeLimit } | 
    Where-Object { $_.Name -match '^.+\####mdata.txt$' } | 
    Remove-Item -Force
}


function ComputeHashForInputFile($FileName) {
    $stream = [System.IO.MemoryStream]::new()
    $writer = [System.IO.StreamWriter]::new($stream)
    $date = ([datetime](Get-ItemProperty -Path $FileName -Name LastWriteTime).lastwritetime).ToString("yyyy-MM-ddTHH:mm:ss")
    $writer.write($FileName + $date)
    $writer.Flush()
    $stream.Position = 0
    $filemetaData = Get-FileHash -InputStream $stream | Select-Object Hash
    $stream.Dispose()
    $writer.Dispose()
    return $filemetaData
}

<#
    .SYNOPSIS
        Checks for existing file metadata. Computes start point and then chunks file into specified size and send
    .DESCRIPTION
        Checks for existing file metadata. Metadata has attributes corresponding to the last row for the file that was successfully sent. 
        Use the metadata for checkpointing and then compute start point. 
        Chunks file into specified size and send
    .PARAMETER FileName
        The File to be used for consumption and pushing Data
    .PARAMETER LinesPerFile
        Defines the Chunk Size
#>
function Send-ChunkedData($FileName, $linesperFile) {

    if ( !(Test-Path $TmpDirName -PathType Container)) {
        New-Item -ItemType directory -Path $TmpDirName
    }

    $TmpFileName = "\tmp"
    $ext = ".txt"
    $filecount = 1
    $reader = $null
  
    try {
        $reader = [io.file]::OpenText($Filename)
        # Create/Get Metadata
        $metaData = GetOrCreateMetadata($FileName)

        try {        
            $header = $reader.ReadLine();
            $activeLineCount = 0

            # Skip no of rows already written as per metadata
            while ($activeLineCount -lt $metaData.NoOfRowsWritten -and $reader.EndOfStream -ne $true) {
                $reader.ReadLine() | Out-Null
                $activeLineCount++
            }
            
            write-output "⏭️ Rows already ingested from file: $activeLineCount"
            
            while ($reader.EndOfStream -ne $true) {              
                $linecount = 0
                $NewFile = "{0}{1}{2}{3}" -f ($TmpDirName, $TmpFileName, $filecount.ToString("0000"), $ext)
                write-Output "📦 Creating chunk file: $NewFile"
                $writer = [io.file]::CreateText($NewFile)
                $filecount++
                
                #"Adding header"
                $writer.WriteLine($header);

                #"Reading $linesperFile"
                while ( ($linecount -lt $linesperFile) -and ($reader.EndOfStream -ne $true)) {
                    $writer.WriteLine($reader.ReadLine());
                    $linecount++
                }

                # Update the active Linecount to be persisted in eventual metadata
                $activeLineCount = $activeLineCount + $linecount
                #"Closing file"
                $writer.Dispose();

                write-Output "✅ Created chunk with $linecount records"
                RetryCommand -ScriptBlock {
                    param($fileName)
                    $access_token = Get-AccessToken
                    Push-Data($access_token, $fileName) $NewFile
                }
                
                # Update metadata
                UpdateMetadata $FileName $activeLineCount
   
                write-Output "🗑️ Deleting temporary chunk file: $NewFile"
                Remove-Item $NewFile
            }
        }
        finally {     
            if ($null -ne $writer) {
                $writer.Dispose();
            }
        }
    }
    finally {
        if ($null -ne $reader) {
            $reader.Dispose();
        }
    }
}

Send-ChunkedData $FilePath $RecordsPerCall

# ═══════════════════════════════════════════════════════════════
# 🧹 CLEANUP
# ═══════════════════════════════════════════════════════════════
Write-Output ""
Write-Output "🧹 Cleaning up temporary files..."
# Clean up the temporary file
Remove-Item $tempFilePath
Write-Output "✨ Script completed successfully!"
Write-Output ""
