<#
.SYNOPSIS
Removes a user from Office 365 groups as specified in an Excel file.

.DESCRIPTION
This function removes a user from different types of Office 365 groups as specified in an Excel file.
The function uses the ExchangeOnlineManagement, ImportExcel, Microsoft.Graph.Groups, and Microsoft.Graph.Users modules.

The function first connects to Exchange Online and Microsoft Graph using the UserPrincipalName provided. 
It then imports data from an Excel file and iterates through each row. For each row, it removes the user from the group based on the group type.

.PARAMETER FilePath
This mandatory parameter specifies the path to the Excel file which contains information about the groups from which the user should be removed.

.PARAMETER UserEmail
This mandatory parameter specifies the email of the user who should be removed from the groups.

.PARAMETER Domain
This mandatory parameter specifies the domain of the user. The domain is used to construct the group name.

.PARAMETER UserRole
This mandatory parameter specifies the user's role. It should be either "NY-IT" or "NY-HR". This parameter is used to identify the worksheet in the Excel file to import.

.EXAMPLE
Remove-CT365GroupByTitle -FilePath "C:\Path\to\file.xlsx" -UserEmail "johndoe@example.com" -Domain "example.com" -UserRole "NY-IT"
This example removes the user "johndoe@example.com" from the groups specified in the "NY-IT" worksheet of the Excel file at "C:\Path\to\file.xlsx".

.NOTES
The Excel file should have columns for PrimarySMTP, GroupType, and DisplayName. These columns are used to get information about the groups from which the user should be removed.

The function supports the following group types: 365Group, 365Distribution, 365MailEnabledSecurity, and 365Security. For each group type, it uses a different cmdlet to remove the user from the group.

.LINK

https://docs.microsoft.com/en-us/powershell/module/microsoft.graph.groups/?view=graph-powershell-1.0

.LINK

https://www.powershellgallery.com/packages/ImportExcel

#>
function Remove-CT365GroupByTitle {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [String]$FilePath,
        
        [Parameter(Mandatory)]
        [string]$UserEmail,

        [Parameter(Mandatory)]
        [string]$Domain,
        
        [Parameter(Mandatory)]
        [ValidateSet("NY-IT", "NY-HR")]
        [string]$UserRole
    )

    Import-Module ExchangeOnlineManagement
    Import-Module ImportExcel
    Import-Module Microsoft.Graph.Groups
    Import-Module Microsoft.Graph.Users

    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "Group.ReadWrite.All","Directory.AccessAsUser.All"

    if (!(Test-Path $FilePath)) {
        Write-Error "Excel file not found at the specified path: $FilePath"
        return
    }

    $excelData = Import-Excel -Path $FilePath -WorksheetName $UserRole

    if ($PSCmdlet.ShouldProcess("Remove user from groups from Excel file")) {
        foreach ($row in $excelData) {
            $GroupName = $row.PrimarySMTP += "@$domain"
            $GroupType = $row.GroupType
            $DisplayName = $row.DisplayName

            if ($PSCmdlet.ShouldProcess("Remove user $UserEmail from $GroupType group $GroupName")) {
                try {
                    switch ($GroupType) {
                        '365Group' {
                            Remove-UnifiedGroupLinks -Identity $GroupName -LinkType "Members" -Links $UserEmail
                            Write-Host "User $UserEmail successfully removed from $GroupType group $GroupName"
                        }
                        '365Distribution' {
                            Remove-DistributionGroupMember -Identity $GroupName -Member $UserEmail -Confirm:$false
                            Write-Host "User $UserEmail successfully removed from $GroupType group $GroupName"
                        }
                        '365MailEnabledSecurity' {
                            Remove-DistributionGroupMember -Identity $GroupName -Member $UserEmail -Confirm:$false
                            Write-Host "User $UserEmail successfully removed from $GroupType group $GroupName"
                        }
                        '365Security' {
                            $user = Get-MgUser -Filter "userPrincipalName eq '$UserEmail'"
                            $ExistingGroup = Get-MgGroup -Filter "DisplayName eq '$($DisplayName)'"
                                if ($ExistingGroup) {
                                Remove-MgGroupMemberByRef -GroupId $ExistingGroup.Id -DirectoryObjectId $User.Id
                                Write-Host "User $UserEmail successfully removed from $GroupType group $GroupName"
                            }
                            else {
                                Write-Warning "No group found with the name: $GroupName"
                            }
                        
                        }
                        default {
                            Write-Warning "Unknown group type: $GroupType"
                        }
                        
                    }
                } catch {
                    Write-Error "Error removing user $UserEmail from $GroupType group $GroupName $_"
                }
            }
        }
    }

# Disconnect Exchange Online and Microsoft Graph sessions
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
}
