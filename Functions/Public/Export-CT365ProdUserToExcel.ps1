<#
.SYNOPSIS
Exports Office 365 production user details to an Excel file.

.DESCRIPTION
This function connects to Microsoft Graph to fetch details of Office 365 users based on certain criteria and then exports those details to an Excel file. The exported details include GivenName, SurName, UserPrincipalName, DisplayName, MailNickname, JobTitle, Department, and address-related fields.

.PARAMETER WorkbookName
The name of the Excel workbook where the user details will be exported. It should have an extension of .xls or .xlsx.

.PARAMETER FilePath
The path to the folder where the Excel file will be created or saved.

.PARAMETER DepartmentFilter
(Optional) Filters users based on the specified department. If not provided, all users will be fetched.

.PARAMETER UserLimit
(Optional) Specifies the maximum number of users to export. If not provided, all users will be exported.

.EXAMPLE
Export-CT365ProdUserToExcel -WorkbookName 'Users.xlsx' -FilePath 'C:\Exports' -DepartmentFilter 'IT' -UserLimit 100
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
        [Parameter(Mandatory)]
        [ValidatePattern('^.*\.(xls|xlsx)$')]
        [string]$WorkbookName,

        [Parameter(Mandatory)]
        [ValidateScript({
            if (Test-Path -Path $_ -PathType Container) {
                $true
            } else {
                throw "Folder path $_ does not exist, please confirm path does exist"
            }
        })]
        [string]$FilePath,

        [Parameter()]
        [string]$DepartmentFilter,

        [Parameter()]
        [int]$UserLimit
    )

    begin {
        
        # Import Required Modules
        $ModulesToImport = "Microsoft.Graph.Authentication","Microsoft.Graph.Users","ImportExcel","PSFramework"
        Import-Module $ModulesToImport
    
        $Path = Join-Path -Path $filepath -ChildPath $workbookname
        
        Write-PSFMessage -Level Output -Message "Creating or Merging workbook $WorkbookName" -Target $WorkbookName
    }

    process {
        
        # Authenticate to Microsoft Graph
        $connection = Connect-MgGraph -Scopes "User.Read.All"
        if (-not $connection) {
            
            Write-PSFMessage -Level Warning -Message "Failed to connect to Microsoft Graph. Please ensure you have the necessary permissions."
            return
        }

        # Build the user retrieval command
        $userCommand = Get-MgUser -Property GivenName, SurName, UserPrincipalName, DisplayName, MailNickname, JobTitle, Department, StreetAddress, City, State, PostalCode, Country, BusinessPhones, MobilePhone, UsageLocation

        # Apply department filter if provided
        if ($DepartmentFilter) {
            $userCommand = $userCommand | Where-Object { $_.Department -eq $DepartmentFilter }
        }

        # Limit the number of users if specified
        if ($UserLimit) {
            $userCommand = $userCommand | Select-Object -First $UserLimit
        }
    
        # Fetch and export users to Excel
        $userCommand | 
        Select-Object GivenName, SurName, UserPrincipalName, DisplayName, MailNickname, JobTitle, Department, StreetAddress, City, State, PostalCode, Country, @{Name="BusinessPhones"; Expression={$_.BusinessPhones -join ", "}}, MobilePhone, UsageLocation |
        Export-Excel -Path $Path -WorksheetName "Users" -AutoSize


        # Disconnect from Microsoft Graph
        Disconnect-MgGraph
    }

    end {
        Write-PSFMessage -Level Output -Message "Export completed. Check the file at $Path for the user details."
    }
}