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
This function requires modules ExchangeOnlineManagement, Microsoft.Graph.Groups, Microsoft.Graph.Users, and ImportExcel.

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
        [string]$UserPrincialName
    )

    # Import the required modules
    $ModulesToImport = "ImportExcel","Microsoft.Graph.Groups","PSFramework","ExchangeOnlineManagement","Microsoft.Graph.Users"
    Import-Module $ModulesToImport
    

    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
    
    # Connect to Microsoft Graph - remove when done testing
    $Scopes = @("Group.ReadWrite.All")
    $Context = Get-MgContext

    if ([string]::IsNullOrEmpty($Context) -or ($Context.Scopes -notmatch [string]::Join('|', $Scopes))) {
        Connect-MGGraph -Scopes $Scopes
    }

    # Import data from Excel
    $Groups = Import-Excel -Path $FilePath -WorksheetName Groups

    foreach ($Group in $Groups) {
        try {
        $writePSFMessageSplat = @{
            Level = 'Output'
            Message = "Removing $($Group.Type):'$($Group.DisplayName)'"
            Target = $Group.DisplayName
        }

        Write-PSFMessage @writePSFMessageSplat

            switch ($Group.Type) {
                "365Group" {
                    $removeUnifiedGroupSplat = @{
                        Identity = $Group.DisplayName
                        Confirm = $false
                        ErrorAction = 'Stop'
                    }

                    Remove-UnifiedGroup @removeUnifiedGroupSplat
                }
                {"365Distribution" -or "365MailEnabledSecurity"} {
                    $removeDistributionGroupSplat = @{
                        Identity = $Group.DisplayName
                        Confirm = $false
                        ErrorAction = 'Stop'
                    }

                    Remove-DistributionGroup @removeDistributionGroupSplat
                }
                "365Security" {
                    $getMgGroupSplat = @{
                        Filter = "DisplayName eq '$($Group.DisplayName)'"
                        ErrorAction = 'Stop'
                    }

                    $existingGroup = Get-MgGroup @getMgGroupSplat

                    $removeMgGroupSplat = @{
                        GroupId = $existingGroup.Id
                        ErrorAction = 'Stop'
                    }

                    Remove-MgGroup @removeMgGroupSplat
                }
                default {
                    $writePSFMessageSplat = @{
                        Level = 'Warning'
                        Message = "Invalid group type for $($Group.DisplayName)"
                        Target = $Group.DisplayName
                    }

                    Write-PSFMessage @writePSFMessageSplat
                }
            }
            $writePSFMessageSplat = @{
                Level = 'Output'
                Message = "Removed $($Group.Type):'$($Group.DisplayName)' successfully"
                Target = $Group.DisplayName
            }

            Write-PSFMessage @writePSFMessageSplat

        }
        catch {
            $writePSFMessageSplat = @{
                Level = 'Output'
                Message = "Could not remove $($Group.Type):'$($Group.DisplayName)' - maybe the group already exists"
                Target = $Group.DisplayName
            }

            Write-PSFMessage @writePSFMessageSplat
        }

    }
    

    # Disconnect Exchange Online and Microsoft Graph sessions
    Disconnect-ExchangeOnline -Confirm:$false
    if (-not [string]::IsNullOrEmpty($(Get-MgContext))) {
        Disconnect-MgGraph
    }
}
