<#
.SYNOPSIS
    Removes a user from specific Office 365 groups based on an Excel file.

.DESCRIPTION
    The Remove-CT365GroupByTitle function uses the Office 365 Exchange Online and Microsoft Graph APIs to remove a user 
    from specific Office 365 groups. The groups to remove the user from are listed in an Excel file.
    
.PARAMETER ExcelFilePath
    The path to an Excel file that contains the groups to remove the user from. Each row in the Excel file represents a group.
    The file must contain columns for PrimarySMTP, GroupType, and DisplayName.

.PARAMETER UserEmail
    The email address of the user to remove from the groups.

.PARAMETER Domain
    The domain of the user.

.PARAMETER UserRole
    The role of the user. This is used as the name of the worksheet in the Excel file to process.
    Valid values are 'NY-IT' and 'NY-HR'.

.EXAMPLE
    Remove-CT365GroupByTitle -ExcelFilePath 'C:\Path\to\file.xlsx' -UserEmail 'user@example.com' -Domain 'example.com' -UserRole 'NY-IT'

    This example removes the user 'user@example.com' from all groups listed in the 'NY-IT' worksheet of 'file.xlsx'.

.NOTES
    This function requires the ExchangeOnlineManagement, ImportExcel, and Microsoft.Graph.Groups PowerShell modules.
    It also requires administrative access to the Office 365 tenant.
#>
function Remove-CT365GroupByTitle {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [String]$ExcelFilePath,
        
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

    if (!(Test-Path $ExcelFilePath)) {
        Write-Error "Excel file not found at the specified path: $ExcelFilePath"
        return
    }

    $excelData = Import-Excel -Path $ExcelFilePath -WorksheetName $UserRole

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
