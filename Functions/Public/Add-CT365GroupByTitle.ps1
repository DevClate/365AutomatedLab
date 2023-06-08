<#
.SYNOPSIS
This function adds a user to Microsoft 365 groups based on a provided Excel file.

.DESCRIPTION
The Add-CT365GroupByTitle function uses Microsoft Graph and Exchange Online Management modules to add a user to different types of Microsoft 365 groups. The group details are read from an Excel file. The group's SMTP, type, and display name are extracted from the Excel file and used to add the user to the group.

.PARAMETER FilePath
The path to the Excel file that contains the groups to which the user should be added. The file must contain a worksheet named as per the user role ("NY-IT" or "NY-HR"). The worksheet should contain details about the groups including the primary SMTP, group type, and display name.

.PARAMETER UserEmail
The email of the user to be added to the groups specified in the Excel file.

.PARAMETER Domain
The domain that is appended to the primary SMTP to form the group's email address.

.PARAMETER UserRole
The role of the user, which is also the name of the worksheet in the Excel file that contains the groups to which the user should be added. The possible values are "NY-IT" and "NY-HR".

.EXAMPLE
Add-CT365GroupByTitle -FilePath "C:\Users\Username\Documents\Groups.xlsx" -UserEmail "jdoe@example.com" -Domain "example.com" -UserRole "NY-IT"

This command reads the groups from the "NY-IT" worksheet in the Groups.xlsx file and adds the user "jdoe@example.com" to those groups.

.NOTES
This function requires the ExchangeOnlineManagement, ImportExcel, and Microsoft.Graph.Groups modules to be installed. It will import these modules at the start of the function. The function connects to Exchange Online and Microsoft Graph, and it will disconnect from them at the end of the function.

#>
function Add-CT365GroupByTitle {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,
        
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

    if (!(Test-Path $FilePath)) {
        Write-Error "Excel file not found at the specified path: $FilePath"
        return
    }

    $excelData = Import-Excel -Path $FilePath -WorksheetName $UserRole

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
