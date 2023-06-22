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

Connect-MgGraph -Scopes "Group.ReadWrite.All","Directory.AccessAsUser.All" - is needed to connect with correct scopes

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
        [string]$UserRole
    )

    # Import Required Modules
    Import-Module ExchangeOnlineManagement
    Import-Module ImportExcel
    Import-Module Microsoft.Graph.Groups
    Import-Module Microsoft.Graph.Users
    Import-Module PSFramework

    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "Group.ReadWrite.All","Directory.AccessAsUser.All"

    if (!(Test-Path $FilePath)) {
        Write-PSFMessage -Level Error -Message "Excel file not found at the specified path: $FilePath" -Target $FilePath
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
                            Write-PSFMessage -Level Output -Message "Removing $UserEmail from 365 Group $GroupName" -Target $UserEmail
                            Remove-UnifiedGroupLinks -Identity $GroupName -LinkType "Members" -Links $UserEmail -Confirm:$false
                            Write-PSFMessage -Level Output -Message "User $UserEmail successfully removed from $GroupType group $GroupName" -Target $UserEmail
                        }
                        '365Distribution' {
                            Write-PSFMessage -Level Output -Message "Removing $UserEmail from 365 Distribution Group $GroupName" -Target $UserEmail
                            Remove-DistributionGroupMember -Identity $GroupName -Member $UserEmail -Confirm:$false
                            Write-PSFMessage -Level Output -Message "User $UserEmail successfully removed from $GroupType group $GroupName" -Target $UserEmail
                        }
                        '365MailEnabledSecurity' {
                            Write-PSFMessage -Level Output -Message "Removing $UserEmail from 365 Mail-Enabled Security Group $GroupName" -Target $UserEmail
                            Remove-DistributionGroupMember -Identity $GroupName -Member $UserEmail -Confirm:$false
                            Write-PSFMessage -Level Output -Message "User $UserEmail successfully removed from $GroupType group $GroupName" -Target $UserEmail
                        }
                        '365Security' {
                            Write-PSFMessage -Level Output -Message "Removing $UserEmail from 365 Security Group $GroupName" -Target $UserEmail
                            $user = Get-MgUser -Filter "userPrincipalName eq '$UserEmail'"
                            $ExistingGroup = Get-MgGroup -Filter "DisplayName eq '$($DisplayName)'"
                                if ($ExistingGroup) {
                                Remove-MgGroupMemberByRef -GroupId $ExistingGroup.Id -DirectoryObjectId $User.Id
                                Write-PSFMessage -Level Output -Message "User $UserEmail successfully removed from $GroupType group $GroupName" -Target $UserEmail
                            }
                            else {
                                Write-PSFMessage -Level Warning -Message "No group found with the name: $GroupName" -Target $GroupName
                            }
                        
                        }
                        default {
                            Write-PSFMessage -Level Warning -Message "Unknown group type: $GroupType" -Target $GroupType
                        }
                        
                    }
                } catch {
                    Write-PSFMessage -Level Error -Message "Error removing user $UserEmail from $GroupType group $GroupName $_" -Target $UserEmail
                }
            }
        }
    }

# Disconnect Exchange Online and Microsoft Graph sessions
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
}
