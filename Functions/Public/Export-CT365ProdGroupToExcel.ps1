<#
.SYNOPSIS
    Exports Office 365 group data to an Excel file.

.DESCRIPTION
    The Export-CT365ProdGroupToExcel function connects to Microsoft Graph, retrieves Office 365 group data based on specified filters, and exports the data to an Excel file. It supports limiting the number of groups to be retrieved.

.PARAMETER FilePath
    Specifies the path to the Excel file (.xlsx) where the group data will be exported. The directory must exist, and the file must have a .xlsx extension.

.PARAMETER GroupLimit
    Limits the number of groups to retrieve. If set to 0 (default), there is no limit.

.EXAMPLE
    Export-CT365ProdGroupToExcel -FilePath "C:\Groups\Groups.xlsx"

    Exports all Office 365 groups to the specified Excel file.

.EXAMPLE
    Export-CT365ProdGroupToExcel -FilePath "C:\Groups\LimitedGroups.xlsx" -GroupLimit 50

    Exports the first 50 Office 365 groups to the specified Excel file.

.NOTES
    Requires the Microsoft.Graph.Authentication, Microsoft.Graph.Groups, ImportExcel, and PSFramework modules.

    The user executing this script must have permissions to access group data via Microsoft Graph.

.LINK
    https://docs.microsoft.com/en-us/graph/api/resources/groups-overview?view=graph-rest-1.0

#>
function Export-CT365ProdGroupToExcel {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
                $extension = [System.IO.Path]::GetExtension($_)
                $directory = [System.IO.Path]::GetDirectoryName($_)

                if ($extension -ne '.xlsx') {
                    throw "The file $_ is not an Excel file (.xlsx). Please specify a file with the .xlsx extension."
                }
                if (-not (Test-Path -Path $directory -PathType Container)) {
                    throw "The directory $directory does not exist. Please specify a valid directory."
                }
                return $true
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
                        switch ($_) {
                            { $_.GroupTypes -contains "Unified" } { "365Group" }
                            { $_.MailEnabled -and -not $_.SecurityEnabled } { "365Distribution" }
                            { $_.MailEnabled -and $_.SecurityEnabled } { "365MailEnabledSecurity" }
                            default { "365Security" }
                        }
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