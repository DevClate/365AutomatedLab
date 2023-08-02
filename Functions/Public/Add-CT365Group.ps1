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
        [ValidateScript({
            #making sure the Filepath leads to a file and not a folder and has a proper extension
            switch ($psitem){
                {-not([System.IO.File]::Exists($psitem))}{
                    throw "The file path '$PSitem' does not lead to an existing file. Please verify the 'FilePath' parameter and ensure that it points to a valid file (folders are not allowed).                "
                }
                {-not(([System.IO.Path]::GetExtension($psitem)) -match "(.xlsx|.xls)")}{
                    "The file path '$PSitem' does not have a valid Excel format. Please make sure to specify a valid file with a .xlsx or .xls extension and try again."
                }
                Default{
                    $true
                }
            }
        })]
        [string]$FilePath,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$UserPrincialName,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
            # Check if the domain fits the pattern
            switch ($psitem) {
                {$psitem -notmatch '^(((?!-))(xn--|_)?[a-z0-9-]{0,61}[a-z0-9]{1,1}\.)*(xn--)?[a-z]{2,}(?:\.[a-z]{2,})+$'}{
                    throw "The provided domain is not in the correct format."
                }
                Default {
                    $true
                }
            }
        })]
        [string]$Domain
    )

    # Import the required modules
    $ModulesToImport = "ImportExcel","Microsoft.Graph.Groups","PSFramework","ExchangeOnlineManagement"
    Import-Module $ModulesToImport

    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "Group.ReadWrite.All"

    # Import data from Excel
    $Groups = Import-Excel -Path $FilePath -WorksheetName Groups

    foreach ($Group in $Groups) {
        # Append the domain to the PrimarySMTP
        $Group.PrimarySMTP += "@$Domain"
        Write-PSFMessage -Level Output -Message "Creating $($Group.Type):'$($Group.DisplayName)'" -Target $Group.DisplayName
        try{
            switch ($Group.Type) {
                "365Group" {
                    $newUnifiedGroupSplat = @{
                        DisplayName = $Group.DisplayName
                        PrimarySMTPAddress = $Group.PrimarySMTP
                        AccessType = 'Private'
                        Notes = $Group.Description
                        RequireSenderAuthenticationEnabled = $False
                        ErrorAction = "Stop"
                    }

                    New-UnifiedGroup @newUnifiedGroupSplat
                }
                {"365Distribution" -or "365MailEnabledSecurity"}{
                    $newDistributionGroupSplat = @{
                        Name = $Group.DisplayName
                        DisplayName = $Group.DisplayName
                        PrimarySMTPAddress = $Group.PrimarySMTP
                        Description = $Group.Description
                        RequireSenderAuthenticationEnabled = $False
                        ErrorAction = "Stop"
                    }
                    if($Group.Type -eq "365MailEnabledSecurity" ){
                        $newDistributionGroupSplat["Type"] = "Security"
                    }

                    New-DistributionGroup @newDistributionGroupSplat
                }
                "365Security" {
                    $mailNickname = $Group.PrimarySMTP.Split('@')[0]
                    $newMgGroupSplat = @{
                        DisplayName = $Group.DisplayName
                        Description = $Group.Description
                        MailNickname = $mailNickname
                        SecurityEnabled = $true
                        ErrorAction = "Stop"
                    }
    
                    New-MgGroup @newMgGroupSplat
                }
                default {
                    Write-PSFMessage -Level Warning -Message "Invalid group type for $($Group.DisplayName)" -Target $Group.DisplayName
                }
            }
        }catch{
            Write-PSFMessage -Level Output -Message "Could not create $($Group.Type):'$($Group.DisplayName)' - maybe the group already exists" -Target $Group.DisplayName
        }
        Write-PSFMessage -Level Output -Message "Created $($Group.Type):'$($Group.DisplayName)' successfully" -Target $Group.DisplayName
    }

    # Disconnect Exchange Online and Microsoft Graph sessions
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
}
