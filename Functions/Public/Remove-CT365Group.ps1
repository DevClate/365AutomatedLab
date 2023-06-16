<#
.SYNOPSIS
This function removes Office 365 groups based on information provided in an Excel file.

.DESCRIPTION
The Remove-CT365Group function is used to remove Office 365 groups. The function imports data from an Excel file and uses it to remove the Office 365 groups. The Excel file should contain a list of groups with their display names and types.

The function supports four types of groups: 
- 365Group
- 365Distribution
- 365MailEnabledSecurity
- 365Security

.PARAMETER FilePath
The full path to the Excel file that contains information about the groups that should be removed. The file should contain a worksheet named 'Groups'. The 'Groups' worksheet should contain the display names and types of the groups.

.PARAMETER UserPrincipalName
The User Principal Name (UPN) of the account to connect to Exchange Online and Microsoft Graph.

.EXAMPLE
Remove-CT365Group -FilePath "C:\Path\to\file.xlsx" -UserPrincipalName "admin@contoso.com"

This example removes the Office 365 groups listed in the 'Groups' worksheet of the 'file.xlsx' file, using the 'admin@contoso.com' UPN to connect to Exchange Online and Microsoft Graph.

.NOTES
1. The function checks if the specified file exists. If it doesn't exist, a warning message is displayed and the function returns.
2. The function imports the required modules: ExchangeOnlineManagement, Microsoft.Graph.Groups, Microsoft.Graph.Users, and ImportExcel.
3. The function connects to Exchange Online and Microsoft Graph.
4. The function imports data from the specified Excel file. It expects to find a worksheet named 'Groups' in the file.
5. The function iterates over the groups listed in the 'Groups' worksheet and removes them. If a group does not exist, a warning message is displayed. If an invalid group type is specified, a warning message is displayed.
6. After all groups have been processed, the function disconnects the Exchange Online and Microsoft Graph sessions.
#>
function Remove-CT365Group {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$FilePath,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$UserPrincialName
    )

    # Check if Excel file exists
    if (!(Test-Path $FilePath)) {
        Write-Warning "File $FilePath does not exist. Please check the file path and try again."
        return
    }

    # Import the required modules
    Import-Module ExchangeOnlineManagement
    Import-Module Microsoft.Graph.Groups
    Import-Module Microsoft.Graph.Users
    Import-Module ImportExcel

    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
    
    # Connect to Microsoft Graph - remove when done testing
    Connect-MgGraph -Scopes "Group.ReadWrite.All"

    # Import data from Excel
    $Groups = Import-Excel -Path $FilePath -WorksheetName Groups

    foreach ($Group in $Groups) {
        switch ($Group.Type) {
            "365Group" {
                try {
                    Write-Output "Removing 365 Group $Group.DisplayName"
                    Get-UnifiedGroup -Identity $Group.DisplayName -ErrorAction Stop
                    Remove-UnifiedGroup -Identity $Group.DisplayName -Confirm:$false
                    Write-Host "Removed 365 Group $($Group.DisplayName)" -ForegroundColor Green
                } catch {
                    Write-Host "365 Group $($Group.DisplayName) does not exist" -ForegroundColor Yellow
                }
            }
            "365Distribution" {
                try {
                    Write-Output "Removing 365 Distribution Group $Group.DisplayName"
                    Get-DistributionGroup -Identity $Group.DisplayName -ErrorAction Stop
                    Remove-DistributionGroup -Identity $Group.DisplayName -Confirm:$false
                    Write-Host "Removed Distribution Group $($Group.DisplayName)" -ForegroundColor Green
                } catch {
                    Write-Host "Distribution Group $($Group.DisplayName) does not exist" -ForegroundColor Yellow
                }
            }
            "365MailEnabledSecurity" {
                try {
                    Write-Output "Removing 365 Mail-Enabled Security Group $Group.DisplayName"
                    Get-DistributionGroup -Identity $Group.DisplayName -ErrorAction Stop
                    Remove-DistributionGroup -Identity $Group.DisplayName -Confirm:$false
                    Write-Host "Removed Mail-Enabled Security Group $($Group.DisplayName)" -ForegroundColor Green
                } catch {
                    Write-Host "Mail-Enabled Security Group $($Group.DisplayName) does not exist" -ForegroundColor Yellow
                }
            }
            "365Security" {
                Write-Output "Removing 365 Security Group $Group.DisplayName"
                $existingGroup = Get-MgGroup -Filter "DisplayName eq '$($Group.DisplayName)'"
                if ($existingGroup) {
                    Remove-MgGroup -GroupId $existingGroup.Id -Confirm:$false
                    Write-Host "Removed Security Group $($Group.DisplayName)" -ForegroundColor Green
                } else {
                    Write-Host "Security Group $($Group.DisplayName) does not exist" -ForegroundColor Yellow
                }
            }
            default {
                Write-Host "Invalid group type for $($Group.DisplayName)" -ForegroundColor Yellow
            }
        }
    }
    

    # Disconnect Exchange Online and Microsoft Graph sessions
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
}
