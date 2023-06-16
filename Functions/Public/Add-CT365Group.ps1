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
Add-CT365Group -FilePath "C:\Path\to\file.xlsx" -UserPrincialName "admin@domain.com" -Domain "domain.com"

This will read the Excel file "file.xlsx" located at "C:\Path\to\", use "admin@domain.com" to connect to Exchange Online, and append "@domain.com" to the PrimarySMTP of each group to form the email address of the group.

.INPUTS
System.String

.OUTPUTS
System.String
The function outputs strings informing about the creation of the groups or if the groups already exist.

.NOTES
The function uses the ExchangeOnlineManagement and Microsoft.Graph.Groups modules to interact with Office 365. Make sure these modules are installed before running the function.

.LINK

Get-UnifiedGroup

.LINK
New-UnifiedGroup

.LINK

Get-DistributionGroup

.LINK

New-DistributionGroup

.LINK

Get-MgGroup

.LINK

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
    Import-Module PSFramework

    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "Group.ReadWrite.All"

    # Import data from Excel
    $Groups = Import-Excel -Path $FilePath -WorksheetName Groups

    foreach ($Group in $Groups) {
        # Append the domain to the PrimarySMTP
        $Group.PrimarySMTP += "@$Domain"
        switch ($Group.Type) {
            "365Group" {
                try {
                    Write-PSFMessage -Level Output -Message "Creating 365 Group $Group.DisplayName" -Target $Group.DisplayName
                    Get-UnifiedGroup -Identity $Group.DisplayName -ErrorAction Stop
                    Write-PSFMessage -Level Warning -Message "365 Group $($Group.DisplayName) already exists" -Target $Group.DisplayName
                } catch {
                    New-UnifiedGroup -DisplayName $Group.DisplayName -PrimarySMTPAddress $Group.PrimarySMTP -AccessType Private -Notes $Group.Description -RequireSenderAuthenticationEnabled $False
                    Write-PSFMessage -Level Output -Message "Created 365 Group $($Group.DisplayName)" -Target $Group.DisplayName
                }
            }
            "365Distribution" {
                try {
                    Write-PSFMessage -Level Output -Message "Creating 365 Distribution Group $Group.DisplayName" -Target $Group.DisplayName
                    Get-DistributionGroup -Identity $Group.DisplayName -ErrorAction Stop
                    Write-PSFMessage -Level Output -Message "Distribution Group $($Group.DisplayName) already exists" -Target $Group.DisplayName
                } catch {
                    New-DistributionGroup -Name $Group.DisplayName -DisplayName $Group.DisplayName -PrimarySMTPAddress $Group.PrimarySMTP -Description $Group.Description -RequireSenderAuthenticationEnabled $False
                    Write-PSFMessage -Level Output -Message "Created Distribution Group $($Group.DisplayName)" -Target $Group.DisplayName
                }
            }
            "365MailEnabledSecurity" {
                try {
                    Write-PSFMessage -Level Output -Message "Creating 365 Mail-Enabled Security Group $Group.DisplayName" -Target $Group.DisplayName
                    Get-DistributionGroup -Identity $Group.DisplayName -ErrorAction Stop
                    Write-PSFMessage -Level Output -Message "Mail-Enabled Security Group $($Group.DisplayName) already exists" -Target $Group.DisplayName
                } catch {
                    New-DistributionGroup -Name $Group.DisplayName -PrimarySMTPAddress $Group.PrimarySMTP -Type "Security" -Description $Group.Description -RequireSenderAuthenticationEnabled $False
                    Write-PSFMessage -Level Output -Message "Created Mail-Enabled Security Group $($Group.DisplayName)" -Target $Group.DisplayName
                }
            }
            "365Security" {
                Write-Output "Creating 365 Security Group $Group.DisplayName"
                $ExistingGroup = Get-MgGroup -Filter "DisplayName eq '$($Group.DisplayName)'"
                if ($ExistingGroup) {
                    Write-Host "Security Group $($Group.DisplayName) already exists" -ForegroundColor Yellow
                    continue
                }
                $mailNickname = $Group.PrimarySMTP.Split('@')[0]
                New-MgGroup -DisplayName $Group.DisplayName -Description $Group.Description -MailNickName $mailNickname -SecurityEnabled:$true -MailEnabled:$false
                Write-Host "Created Security Group $($Group.DisplayName)" -ForegroundColor Green
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
