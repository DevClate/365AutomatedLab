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
This function requires modules ExchangeOnlineManagement, Microsoft.Graph.Groups, Microsoft.Graph.Users, PSFramework, and ImportExcel.
#>
function Remove-CT365Group {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
            #making sure the Filepath leads to a file and not a folder and has a proper extension
            switch ($psitem){
                {-not([System.IO.File]::Exists($psitem))}{
                    throw "The file path '$PSitem' does not lead to an existing file. Please verify the 'FilePath' parameter and ensure that it points to a valid file (folders are not allowed).                "
                }
                {-not(([System.IO.Path]::GetExtension($psitem)) -match "(.xlsx)")}{
                    "The file path '$PSitem' does not have a valid Excel format. Please make sure to specify a valid file with a .xlsx extension and try again."
                }
                Default{
                    $true
                }
            }
        })]
        [string]$FilePath,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$UserPrincipalName
    )

    # Import the required modules
    $ModulesToImport = "ImportExcel","Microsoft.Graph.Groups","PSFramework","ExchangeOnlineManagement","Microsoft.Graph.Users"
    Import-Module $ModulesToImport

    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
    
    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "Group.ReadWrite.All"
    
    # Import data from Excel
    $Groups = Import-Excel -Path $FilePath -WorksheetName Groups
    foreach ($Group in $Groups) {
        switch ($Group.Type) {
            "365Group" {
                try {
                    Write-PSFMessage -Level Output -Message "Removing 365 Group: $($Group.DisplayName)" -Target $Group.DisplayName
                    Get-UnifiedGroup -Identity $Group.DisplayName -ErrorAction Stop
                    Remove-UnifiedGroup -Identity $Group.DisplayName -Confirm:$false
                    Write-PSFMessage -Level Output -Message "Removed 365 Group: $($Group.DisplayName)" -Target $Group.DisplayName
                } catch {
                    Write-PSFMessage -Level Warning -Message "365 Group $($Group.DisplayName) does not exist" -Target $Group.DisplayName -ErrorRecord $_
                    Continue
                }
            }
            "365Distribution" {
                try {
                    Write-PSFMessage -Level Output "Removing 365 Distribution Group: $($Group.DisplayName)" -Target $Group.DisplayName
                    Get-DistributionGroup -Identity $Group.DisplayName -ErrorAction Stop
                    Remove-DistributionGroup -Identity $Group.DisplayName -Confirm:$false
                    Write-PSFMessage -Level Output -Message "Removed Distribution Group: $($Group.DisplayName)" -Target $Group.DisplayName
                } catch {
                    Write-PSFMessage -Level Warning -Message "Distribution Group $($Group.DisplayName) does not exist" -Target $Group.DisplayName -ErrorRecord $_
                    Continue
                }
            }
            "365MailEnabledSecurity" {
                try {
                    Write-PSFMessage -Level Output -Message "Removing 365 Mail-Enabled Security Group: $($Group.DisplayName)" -Target $Group.DisplayName
                    Get-DistributionGroup -Identity $Group.DisplayName -ErrorAction Stop
                    Remove-DistributionGroup -Identity $Group.DisplayName -Confirm:$false
                    Write-PSFMessage -Level Output -Message "Removed Mail-Enabled Security Group $($Group.DisplayName)" -Target $Group.DisplayName
                } catch {
                    Write-PSFMessage -Level Warning -Message "Mail-Enabled Security Group $($Group.DisplayName) does not exist" -Target $Group.DisplayName -ErrorRecord $_
                    Continue
                }
            }
            "365Security" {
                Write-PSFMessage -Level Output -Message "Removing 365 Security Group $($Group.DisplayName)" -Target $Group.DisplayName
                $existingGroup = Get-MgGroup -Filter "DisplayName eq '$($Group.DisplayName)'"
                if ($existingGroup) {
                    Remove-MgGroup -GroupId $existingGroup.Id -Confirm:$false
                    Write-PSFMessage -Level Output -Message "Removed Security Group $($Group.DisplayName)" -Target $Group.DisplayName
                } else {
                    Write-PSFMessage -Level Warning -Message "Security Group $($Group.DisplayName) does not exist" -Target $Group.DisplayName
                }
            }
            default {
                Write-PSFMessage -Level Warning -Message "Invalid group type for $($Group.DisplayName)" -Target $Group.DisplayName
            }
        }
    }
    
    # Disconnect Exchange Online and Microsoft Graph sessions
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
}