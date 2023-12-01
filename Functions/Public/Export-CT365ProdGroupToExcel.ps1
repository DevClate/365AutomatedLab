<#
.SYNOPSIS
Exports Office 365 production group details to an Excel file.

.DESCRIPTION
This function connects to Microsoft Graph to fetch details of Office 365 groups based on certain criteria and then exports those details to an Excel file. The exported details include DisplayName, PrimarySMTP, Description, and Type.

.PARAMETER WorkbookName
The name of the Excel workbook where the group details will be exported. It should have an extension of .xlsx.

.PARAMETER FilePath
The path to the folder where the Excel file will be created or saved.

.PARAMETER GroupLimit
(Optional) Specifies the maximum number of groups to export. If not provided, all groups will be exported.

.EXAMPLE
Export-CT365ProdGroupToExcel -WorkbookName 'Groups.xlsx' -FilePath 'C:\Exports' -GroupLimit 100
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
        [Parameter(Mandatory)]
        [ValidatePattern('^.*\.(xlsx)$')]
        [string]$WorkbookName,

        [Parameter(Mandatory)]
        [ValidateScript({
                Test-Path -Path $_ -PathType Container
            })]
        [string]$FilePath,

        [Parameter(Mandatory=$false)]
        [int]$GroupLimit = 0
    )

    begin {
        Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Groups, ImportExcel, PSFramework
        $Path = Join-Path -Path $FilePath -ChildPath $WorkbookName
        Write-PSFMessage -Level Output -Message "Preparing to export to $WorkbookName" -Target $WorkbookName
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
        
        Get-MgGroup @getMgGroupSplat | Select-Object @selectProperties | Export-Excel -Path $Path -WorksheetName "Groups" -AutoSize

        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }

    end {
        Write-PSFMessage -Level Output -Message "Export completed. Check the file at $Path for the group details."
    }
}