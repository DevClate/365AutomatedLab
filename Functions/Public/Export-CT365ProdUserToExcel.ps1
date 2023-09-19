<#
.SYNOPSIS
Exports Office 365 production user details to an Excel file.

.DESCRIPTION
This function connects to Microsoft Graph to fetch details of Office 365 users based on certain criteria and then exports those details to an Excel file. The exported details include GivenName, SurName, UserPrincipalName, DisplayName, MailNickname, JobTitle, Department, and address-related fields.

.PARAMETER WorkbookName
The name of the Excel workbook where the user details will be exported. It should have an extension of .xlsx.

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
        [ValidatePattern('^.*\.(xlsx)$')]
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
        [int]$UserLimit = 0
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
            #$userCommand = $userCommand | Where-Object { $_.Department -eq $DepartmentFilter }
            $getMgUserSplat.Add('filter',"Department eq '$DepartmentFilter'")
        }

        # Limit the number of users if specified else get all users
        if ($UserLimit-eq 0) {
            $getMgUserSplat.Add("all",$true)
        }else{
            $getMgUserSplat.Add("Top",$UserLimit)
        }

        $userCommand = Get-MgUser @getMgUserSplat | Select-Object -Property $getMgUserSplat.Property

        #alter the Businessphones property so they can be displayed in the excel file correctly
        foreach($User in $userCommand){
            $User.BusinessPhones = $User.BusinessPhones -join ", "
        }

        # Fetch and export users to Excel
        $userCommand | Export-Excel -Path $Path -WorksheetName "Users" -AutoSize


        # Disconnect from Microsoft Graph
        if (-not [string]::IsNullOrEmpty($(Get-MgContext))) {
            Disconnect-MgGraph
        }
    }

    end {
        Write-PSFMessage -Level Output -Message "Export completed. Check the file at $Path for the user details."
    }
}