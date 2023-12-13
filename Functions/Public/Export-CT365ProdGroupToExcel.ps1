<#
.SYNOPSIS
Exports Office 365 production group details to an Excel file.

.DESCRIPTION
This function connects to Microsoft Graph to fetch details of Office 365 groups based on certain criteria and then exports those details to an Excel file. The exported details include DisplayName, PrimarySMTP, Description, and Type.

.PARAMETER FilePath
The full path to the Excel file where the group details will be exported, including the file name and .xlsx extension.

.PARAMETER GroupLimit
(Optional) Specifies the maximum number of groups to export. If not provided, all groups will be exported.

.EXAMPLE
Export-CT365ProdGroupToExcel -FilePath 'C:\Exports\Groups.xlsx' -GroupLimit 100
This example exports the first 100 groups to an Excel file named 'Groups.xlsx' located at 'C:\Exports'.

.NOTES
This function requires the following modules to be installed:
- Microsoft.Graph.Authentication
- Microsoft.Graph.Groups
- ImportExcel
- PSFramework

The user executing this function should have the necessary permissions to read group details from Microsoft Graph.

.LINK
[Microsoft Graph PowerShell SDK](https://github.com/microsoftgraph/msgraph-sdk-powershell)

#>
function Export-CT365ProdGroupToExcel {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
                $isValid = $false
                $extension = [System.IO.Path]::GetExtension($_)
                $directory = [System.IO.Path]::GetDirectoryName($_)

                if ($extension -ne '.xlsx') {
                    throw "The file $_ is not an Excel file (.xlsx). Please specify a file with the .xlsx extension."
                }
                elseif (-not (Test-Path -Path $directory -PathType Container)) {
                    throw "The directory $directory does not exist. Please specify a valid directory."
                }
                else {
                    $isValid = $true
                }
                return $isValid
            })]
        [string]$FilePath,

        [Parameter()]
        [int]$GroupLimit = 0
    )

    begin {
        Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Groups, ImportExcel, PSFramework
        Write-PSFMessage -Level Output -Message "Preparing to export to $FilePath" -Target $FilePath
    }

    process {
        # Authenticate to Microsoft Graph
        $Scopes = @("Group.Read.All")
        $Context = Get-MgContext
            
        if ([string]::IsNullOrEmpty($Context) -or ($Context.Scopes -notmatch [string]::Join('|', $Scopes))) {
            Connect-MGGraph -Scopes $Scopes
        }

        $groupTypeFilters = @(
            "groupTypes/any(c:c eq 'Unified')",
            "(mailEnabled eq true and securityEnabled eq false)",
            "(mailEnabled eq true and securityEnabled eq true)",
            "(mailEnabled eq false and securityEnabled eq true)"
        )
        $filterQuery = $groupTypeFilters -join " or "

        $getMgGroupSplat = @{
            Filter = $filterQuery
            Select = 'DisplayName', 'MailNickname', 'Description', 'GroupTypes', 'MailEnabled', 'SecurityEnabled'
        }
        if ($GroupLimit -gt 0) { $getMgGroupSplat.Add("Top", $GroupLimit) }
        else { $getMgGroupSplat.Add("All", $true) }

        $selectProperties = @{
            Property = @(
                'DisplayName',
                @{Name = 'PrimarySMTP'; Expression = { $_.MailNickname } },
                'Description',
                @{Name = 'Type'; Expression = {
                        if ($_.GroupTypes -contains "Unified") { "365Group" }
                        elseif ($_.MailEnabled -and -not $_.SecurityEnabled) { "365Distribution" }
                        elseif ($_.MailEnabled -and $_.SecurityEnabled) { "365MailEnabledSecurity" }
                        else { "365Security" }
                    }
                }
            )
        }
        
        Get-MgGroup @getMgGroupSplat | Select-Object @selectProperties | Export-Excel -Path $FilePath -WorksheetName "Groups" -AutoSize

        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }

    end {
        Write-PSFMessage -Level Output -Message "Export completed. Check the file at $FilePath for the group details."
    }
}
