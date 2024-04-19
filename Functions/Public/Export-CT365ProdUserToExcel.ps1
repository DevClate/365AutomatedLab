<#
.SYNOPSIS
    Exports Office 365 production user data to an Excel file.

.DESCRIPTION
    The Export-CT365ProdUserToExcel function connects to Microsoft Graph, retrieves user data based on specified filters, and exports the data to an Excel file. It supports filtering by department, limiting the number of users, and an option to exclude license information.

.PARAMETER FilePath
    Specifies the path to the Excel file (.xlsx) where the user data will be exported. The directory must exist, and the file must have a .xlsx extension.

.PARAMETER DepartmentFilter
    Filters users by their department. If not specified, users from all departments are retrieved.

.PARAMETER UserLimit
    Limits the number of users to retrieve. If set to 0 (default), there is no limit.

.PARAMETER NoLicense
    If specified, the exported data will not include license information for the users.

.EXAMPLE
    Export-CT365ProdUserToExcel -FilePath "C:\Users\Export\Users.xlsx"

    Exports all Office 365 production users to the specified Excel file.

.EXAMPLE
    Export-CT365ProdUserToExcel -FilePath "C:\Users\Export\DeptUsers.xlsx" -DepartmentFilter "IT"

    Exports Office 365 production users from the IT department to the specified Excel file.

.EXAMPLE
    Export-CT365ProdUserToExcel -FilePath "C:\Users\Export\Users.xlsx" -UserLimit 100

    Exports the first 100 Office 365 production users to the specified Excel file.

.NOTES
    Requires the Microsoft.Graph.Authentication, Microsoft.Graph.Users, ImportExcel, and PSFramework modules.

    The user executing this script must have permissions to access user data via Microsoft Graph.

.LINK
    https://docs.microsoft.com/en-us/graph/api/resources/users?view=graph-rest-1.0

#>
function Export-CT365ProdUserToExcel {
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
        [string]$DepartmentFilter,

        [Parameter()]
        [int]$UserLimit = 0,

        [Parameter()]
        [switch]$NoLicense
    )

    begin {
        # Import Required Modules
        $ModulesToImport = "Microsoft.Graph.Authentication", "Microsoft.Graph.Users", "ImportExcel", "PSFramework"
        Import-Module $ModulesToImport

        Write-PSFMessage -Level Output -Message "Creating or Merging workbook at $FilePath"
    }

    process {
        # Authenticate to Microsoft Graph
        $Scopes = @("User.Read.All")
        $Context = Get-MgContext
    
        if ([string]::IsNullOrEmpty($Context) -or ($Context.Scopes -notmatch [string]::Join('|', $Scopes))) {
            Connect-MGGraph -Scopes $Scopes
        }

        # Build the user retrieval command
        $getMgUserSplat = @{
            Property = 'GivenName', 'SurName', 'UserPrincipalName', 
            'DisplayName', 'MailNickname', 'JobTitle', 
            'Department', 'StreetAddress', 'City', 
            'State', 'PostalCode', 'Country', 
            'BusinessPhones', 'MobilePhone', 'FaxNumber', 'UsageLocation',
            'CompanyName', 'EmployeeHireDate', 'EmployeeId', 'EmployeeType'
        }
        
        if (-not [string]::IsNullOrEmpty($DepartmentFilter)) {
            $getMgUserSplat['Filter'] = "Department eq '$DepartmentFilter'"
        }

        if ($UserLimit -gt 0) {
            $getMgUserSplat['Top'] = $UserLimit
        }
        else {
            $getMgUserSplat['All'] = $true
        }

        $selectProperties = @{
            Property = @(
                @{Name = 'FirstName'; Expression = { $_.GivenName } },
                @{Name = 'LastName'; Expression = { $_.SurName } },
                @{Name = 'UserName'; Expression = { $_.UserPrincipalName -replace '@.*' } },
                @{Name = 'Title'; Expression = { $_.JobTitle } },
                'Department', 'StreetAddress', 'City', 'State', 'PostalCode', 'Country',
                @{Name = 'PhoneNumber'; Expression = { $_.BusinessPhones } },
                'MobilePhone', 'FaxNumber', 'UsageLocation', 'CompanyName',
                'EmployeeHireDate', 'EmployeeId', 'EmployeeType',
                @{Name = 'License'; Expression = { if ($NoLicense) { "" } else { "DEVELOPERPACK_E5" } } }
            )
        }

        $userCommand = Get-MgUser @getMgUserSplat | Select-Object @selectProperties
        
        # Fetch and export users to Excel
        $userCommand | Export-Excel -Path $FilePath -WorksheetName "Users" -AutoSize

        # Disconnect from Microsoft Graph
        if (-not [string]::IsNullOrEmpty($(Get-MgContext))) {
            Disconnect-MgGraph
        }
    }

    end {
        Write-PSFMessage -Level Output -Message "Export completed. Check the file at $FilePath for the user details."
    }
}
