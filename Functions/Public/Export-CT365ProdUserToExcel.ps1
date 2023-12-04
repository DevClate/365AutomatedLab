<#
.SYNOPSIS
Exports Office 365 production user details to an Excel file.

.DESCRIPTION
This function connects to Microsoft Graph to fetch details of Office 365 users based on certain criteria and then exports those details to an Excel file. The exported details include GivenName, SurName, UserPrincipalName, DisplayName, MailNickname, JobTitle, Department, and address-related fields.

.PARAMETER FilePath
The full path to the Excel file where the user details will be exported, including the file name with an .xlsx extension.

.PARAMETER DepartmentFilter
(Optional) Filters users based on the specified department. If not provided, all users will be fetched.

.PARAMETER UserLimit
(Optional) Specifies the maximum number of users to export. If not provided, all users will be exported.

.EXAMPLE
Export-CT365ProdUserToExcel -FilePath 'C:\Exports\Users.xlsx' -DepartmentFilter 'IT' -UserLimit 100
This example exports the first 100 users from the IT department to an Excel file named 'Users.xlsx' located at 'C:\Exports'.

.NOTES
This function requires the following modules to be installed:
- Microsoft.Graph.Authentication
- Microsoft.Graph.Users
- ImportExcel
- PSFramework

The user executing this function should have the necessary permissions to read user details from Microsoft Graph.

.LINK
[Microsoft Graph PowerShell SDK](https://github.com/microsoftgraph/msgraph-sdk-powershell)

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
        [int]$UserLimit = 0
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
            Property = (
                'GivenName', 'SurName', 'UserPrincipalName', 
                'DisplayName', 'MailNickname', 'JobTitle', 
                'Department', 'StreetAddress', 'City', 
                'State', 'PostalCode', 'Country', 
                'BusinessPhones', 'MobilePhone', 'UsageLocation'
            )
        }
        
        # Apply department filter if provided as parameter
        if (-not [string]::IsNullOrEmpty($DepartmentFilter)) {
            $getMgUserSplat.Add('filter', "Department eq '$DepartmentFilter'")
        }

        # Limit the number of users if specified else get all users
        if ($UserLimit -eq 0) {
            $getMgUserSplat.Add("all", $true)
        }
        else {
            $getMgUserSplat.Add("Top", $UserLimit)
        }

        $selectProperties = @{
            Property = @(
                @{Name='FirstName'; Expression={$_.GivenName}},
                @{Name='LastName'; Expression={$_.SurName}},
                @{Name='UserName'; Expression={$_.UserPrincipalName -replace '@.*'}},
                @{Name='Title'; Expression={$_.JobTitle}},
                'Department', 'StreetAddress', 'City', 'State', 'PostalCode', 'Country',
                @{Name='PhoneNumber'; Expression={$_.BusinessPhones}},
                'MobilePhone', 'UsageLocation'
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
