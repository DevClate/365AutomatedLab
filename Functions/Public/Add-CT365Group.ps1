<#
.SYNOPSIS
This function creates Office 365 Groups, Distribution Groups, Mail-Enabled Security Groups, and Security Groups based on the data provided in an Excel file.

.DESCRIPTION
The function Add-CT365Group takes the path of an Excel file, User Principal Name and a Domain as input. It creates Office 365 Groups, Distribution Groups, Mail-Enabled Security Groups, and Security Groups based on the data found in the Excel file. If a group already exists, the function will output a message and skip the creation of that group.

.PARAMETER FilePath
The full path to the Excel file containing the data for the groups to be created. The Excel file should have a worksheet named "Groups". Each row in this worksheet should represent a group to be created. The columns in the worksheet should include the DisplayName, Type (365Group, 365Distribution, 365MailEnabledSecurity, 365Security), PrimarySMTP (without domain), and Description of the group.

.PARAMETER UserPrincialName
The User Principal Name (UPN) used to connect to Exchange Online.

.PARAMETER Domain
The domain to be appended to the PrimarySMTP of each group to form the email address of the group.

.EXAMPLE
Add-CT365Group -FilePath "C:\Users\user\Desktop\GroupsData.xlsx" -UserPrincialName "admin@domain.com" -Domain "domain.com"

This will read the Excel file "GroupsData.xlsx" located at "C:\Users\user\Desktop\", use "admin@domain.com" to connect to Exchange Online, and append "@domain.com" to the PrimarySMTP of each group to form the email address of the group.

.INPUTS
System.String

.OUTPUTS
System.String
The function outputs strings informing about the creation of the groups or if the groups already exist.

.NOTES
The function uses the ExchangeOnlineManagement and Microsoft.Graph.Groups modules to interact with Office 365. Make sure these modules are installed before running the function.

.LINK
Get-UnifiedGroup
New-UnifiedGroup
Get-DistributionGroup
New-DistributionGroup
Get-MgGroup
New-MgGroup
#>
function Add-CT365Group {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$FilePath,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$UserPrincialName,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$Domain
    )

    # Check if Excel file exists
    if (!(Test-Path $FilePath)) {
        Write-Warning "File $FilePath does not exist. Please check the file path and try again."
        return
    }

    # Import the required modules
    Import-Module ExchangeOnlineManagement
    Import-Module Microsoft.Graph.Groups
    Import-Module ImportExcel

    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "Group.ReadWrite.All"

    # Import data from Excel
    $groups = Import-Excel -Path $FilePath -WorksheetName Groups

    foreach ($group in $groups) {
        # Append the domain to the PrimarySMTP
        $group.PrimarySMTP += "@$Domain"
        switch ($group.Type) {
            "365Group" {
                try {
                    Write-Output "Creating 365 Group $group.DisplayName"
                    Get-UnifiedGroup -Identity $group.DisplayName -ErrorAction Stop
                    Write-Host "365 Group $($group.DisplayName) already exists" -ForegroundColor Yellow
                } catch {
                    New-UnifiedGroup -DisplayName $group.DisplayName -PrimarySMTPAddress $group.PrimarySMTP -AccessType Private -Notes $group.Description -RequireSenderAuthenticationEnabled $False
                    Write-Host "Created 365 Group $($group.DisplayName)" -ForegroundColor Green
                }
            }
            "365Distribution" {
                try {
                    Write-Output "Creating 365 Distribution Group $group.DisplayName"
                    Get-DistributionGroup -Identity $group.DisplayName -ErrorAction Stop
                    Write-Host "Distribution Group $($group.DisplayName) already exists" -ForegroundColor Yellow
                } catch {
                    New-DistributionGroup -Name $group.DisplayName -DisplayName $group.DisplayName -PrimarySMTPAddress $group.PrimarySMTP -Description $group.Description -RequireSenderAuthenticationEnabled $False
                    Write-Host "Created Distribution Group $($group.DisplayName)" -ForegroundColor Green
                }
            }
            "365MailEnabledSecurity" {
                try {
                    Write-Output "Creating 365 Mail-Enabled Security Group $group.DisplayName"
                    Get-DistributionGroup -Identity $group.DisplayName -ErrorAction Stop
                    Write-Host "Mail-Enabled Security Group $($group.DisplayName) already exists" -ForegroundColor Yellow
                } catch {
                    New-DistributionGroup -Name $group.DisplayName -PrimarySMTPAddress $group.PrimarySMTP -Type "Security" -Description $group.Description -RequireSenderAuthenticationEnabled $False
                    Write-Host "Created Mail-Enabled Security Group $($group.DisplayName)" -ForegroundColor Green
                }
            }
            "365Security" {
                Write-Output "Creating 365 Security Group $group.DisplayName"
                $ExistingGroup = Get-MgGroup -Filter "DisplayName eq '$($group.DisplayName)'"
                if ($ExistingGroup) {
                    Write-Host "Security Group $($group.DisplayName) already exists" -ForegroundColor Yellow
                    continue
                }
                $mailNickname = $group.PrimarySMTP.Split('@')[0]
                New-MgGroup -DisplayName $group.DisplayName -Description $group.Description -MailNickName $mailNickname -SecurityEnabled:$true -MailEnabled:$false
                Write-Host "Created Security Group $($group.DisplayName)" -ForegroundColor Green
            }
            default {
                Write-Host "Invalid group type for $($group.DisplayName)" -ForegroundColor Yellow
            }
        }
    }

    # Disconnect Exchange Online and Microsoft Graph sessions
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
}
