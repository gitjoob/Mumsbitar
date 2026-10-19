Import-Module Microsoft.Graph.Identity.Governance

$Devs = Get-MgUser -All -Property DisplayName,UserPrincipalName,JobTitle,Id,AccountEnabled | Where-Object { $_.JobTitle -like '*Utveckl*' -or  $_.jobTitle -like '*devel*' -and $_.UserPrincipalName -notlike '*ext@ssg*' -and $_.DisplayName -notlike '*EXT' -and $_.AccountEnabled -eq $true} | Select-Object DisplayName,UserPrincipalName,JobTitle,Id

$consultants = Get-MgUser -All -Property DisplayName,UserPrincipalName,JobTitle,Id,AccountEnabled | Where-Object { $_.JobTitle -like '*Utveckl*' -or  $_.jobTitle -like '*devel*' -and $_.UserPrincipalName -like '*ext@ssg*' -or $_.DisplayName -like '*EXT' -and $_.AccountEnabled -eq $true -and $_.JobTitle -like '*Utveckl*' -or  $_.jobTitle -like '*devel*'} | Select-Object DisplayName,UserPrincipalName,JobTitle,Id

$finance = Get-Mguser -All -Property DisplayName,UserPrincipalName,Department,JobTitle,Id,AccountEnabled | Where-Object {$_.UserPrincipalName -like '*ext@ssg*' -or $_.DisplayName -like '*EXT' -and $_.Department -eq 'Finance' -or $_.Department -eq "Finans" -or $_.Department -eq "Ekonomi" -and $_.AccountEnabled -eq $true} | Select-Object DisplayName,UserPrincipalName,JobTitle,Id

$printer = Import-Csv -Path '/Users/johan.oberg/Downloads/Printer Canon imagePROGRAF PRO-4000S_2025-3-11.csv'

foreach ($user in $printer){

    $id = $user.id
$params = @{
    requestType = "adminAdd"
    assignment = @{
        targetId = "$id"
        assignmentPolicyId = ""
        accessPackageId = ""        
    }
    justification = "Printer"
}
New-MgEntitlementManagementAssignmentRequest -BodyParameter $params
}