<#
.SYNOPSIS
Add a user to specified 365 groups as per the data in the Excel file.

.DESCRIPTION
The Add-CT365GroupByTitle function connects to Exchange Online and Microsoft Graph to add a user to specified 365 groups based on an Excel file. The Excel file should contain the group details.

.PARAMETER ExcelFilePath
The full file path to the Excel file that contains group details. This parameter is mandatory.

.PARAMETER UserEmail
The email of the user that needs to be added to the specified groups. This parameter is mandatory.

.PARAMETER Domain
The domain to be appended to the group names obtained from the Excel file. This parameter is mandatory.

.PARAMETER UserRole
The role of the user that needs to be added. This should be either "NY-IT" or "NY-HR". This parameter is mandatory.

.EXAMPLE
Add-CT365GroupByTitle -ExcelFilePath "C:\path\to\file.xlsx" -UserEmail "user@domain.com" -Domain "domain.com" -UserRole "NY-IT"

This will add the user "user@domain.com" to the 365 groups as per the data in the "C:\path\to\file.xlsx" file and with the role "NY-IT".

.NOTES
Make sure to have the required modules installed and the user running the script has necessary permissions in Exchange Online and Microsoft Graph.

.LINK
For more information about the ExchangeOnlineManagement, ImportExcel, and Microsoft.Graph.Groups modules, see their respective documentation.

#>
function Add-CT365GroupByTitle {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$ExcelFilePath,
        
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

    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "Group.ReadWrite.All"

    if (!(Test-Path $ExcelFilePath)) {
        Write-Error "Excel file not found at the specified path: $ExcelFilePath"
        return
    }

    $excelData = Import-Excel -Path $ExcelFilePath -WorksheetName $UserRole

    if ($PSCmdlet.ShouldProcess("Add user to groups from Excel file")) {
        foreach ($row in $excelData) {
            $GroupName = $row.PrimarySMTP += "@$domain"
            $GroupType = $row.GroupType
            $DisplayName = $row.DisplayName

            if ($PSCmdlet.ShouldProcess("Add user $UserEmail to $GroupType group $GroupName")) {
                try {
                    switch ($GroupType) {
                        '365Group' {
                            Add-UnifiedGroupLinks -Identity $GroupName -LinkType "Members"-Links $UserEmail
                            Write-Host "User $UserEmail successfully added to $GroupType group $GroupName"
                        }
                        '365Distribution' {
                            Add-DistributionGroupMember -Identity $GroupName -Member $UserEmail
                            Write-Host "User $UserEmail successfully added to $GroupType group $GroupName"
                        }
                        '365MailEnabledSecurity' {
                            Add-DistributionGroupMember -Identity $GroupName -Member $UserEmail
                            Write-Host "User $UserEmail successfully added to $GroupType group $GroupName"
                        }
                        '365Security' {
                            $user = Get-MgUser -Filter "userPrincipalName eq '$UserEmail'"
                            $ExistingGroup = Get-MgGroup -Filter "DisplayName eq '$($DisplayName)'"
                                if ($ExistingGroup) {
                                New-MgGroupMember -GroupId $ExistingGroup.Id -DirectoryObjectId $User.Id
                                Write-Host "User $UserEmail successfully added to $GroupType group $GroupName"
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
                    Write-Error "Error adding user $UserEmail to $GroupType group $GroupName $_"
                }
            }
        }
    }

# Disconnect Exchange Online and Microsoft Graph sessions
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
}
