<#
.SYNOPSIS
Creates Office 365 groups, distribution groups, mail-enabled security groups, and security groups based on an input Excel file.

.DESCRIPTION
The Add-CT365Group function reads an Excel file, connects to Exchange Online and Microsoft Graph, and creates various types of groups based on the data in the Excel file. The types of groups it can create are 365 Group, 365 Distribution, 365 Mail-Enabled Security, and 365 Security. If a group with the same name already exists, it will not be created and a warning message will be printed.

.PARAMETER FilePath
The path to the Excel file containing the data for the groups to be created.

.PARAMETER UserPrincipalName
The user principal name used to connect to Exchange Online. This parameter is mandatory.

.PARAMETER Domain
The domain to be appended to the PrimarySMTP. This parameter is mandatory.

.EXAMPLE
Add-CT365Group -FilePath C:\Data\365\365DataEnvironment.xlsx -UserPrincipalName john.doe@contoso.com -domain contoso.com

This example creates groups based on the data in '365DataEnvironment.xlsx' using the user principal name 'john.doe@contoso.com' and the domain 'contoso.com'.

.INPUTS
None. You cannot pipe objects to Add-CT365Group.

.OUTPUTS
This function does not produce any output. It will print to the console the progress of creating each group, and whether the group already exists or was created successfully.

.NOTES
This function requires the ExchangeOnlineManagement, Microsoft.Graph, and ImportExcel modules.
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
    Import-Module Microsoft.Graph
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
