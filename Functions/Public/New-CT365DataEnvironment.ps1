<#
.SYNOPSIS
    Creates a new Office 365 data environment template in an Excel workbook.

.DESCRIPTION
    The New-CT365DataEnvironment function creates a new Excel workbook with multiple worksheets 
    for Users, Groups, Teams, Sites, and specified Job Roles. Each worksheet is formatted 
    with predefined columns relevant to its content.

.PARAMETER FilePath
    Specifies the path to the Excel file (.xlsx) where the data environment will be created.
    The function checks if the file already exists, if it's a valid .xlsx file, and if the 
    folder path exists.

.PARAMETER JobRole
    An array of job roles to create individual worksheets for each role in the Excel workbook.
    Each job role will have a worksheet with predefined columns.

.EXAMPLE
    PS> New-CT365DataEnvironment -FilePath "C:\Data\O365Environment.xlsx" -JobRole "HR", "IT"
    This command creates an Excel workbook at the specified path with worksheets for Users, 
    Groups, Teams, Sites, and additional worksheets for 'HR' and 'IT' job roles.

.EXAMPLE
    PS> New-CT365DataEnvironment -FilePath "C:\Data\NewEnvironment.xlsx" -JobRole "Finance"
    This command creates an Excel workbook at the specified path with a worksheet for the 
    'Finance' job role, along with the standard Users, Groups, Teams, and Sites worksheets.

.INPUTS
    None. You cannot pipe objects to New-CT365DataEnvironment.

.OUTPUTS
    None. This function does not generate any output.

.NOTES
    Requires the modules ImportExcel and PSFramework to be installed.

.LINK
    https://www.powershellgallery.com/packages/ImportExcel
    https://www.powershellgallery.com/packages/PSFramework

#>

function New-CT365DataEnvironment {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
                # Check if the file has a valid Excel extension (.xlsx)
                if (-not(([System.IO.Path]::GetExtension($psitem)) -match "\.(xlsx)$")) {
                    throw "The file path '$PSitem' does not have a valid Excel format. Please specify a valid file with a .xlsx extension."
                }

                # Check if the file exists
                if (-not([System.IO.File]::Exists($psitem))) {
                    # Prompt the user to create the file path
                    $createFile = Read-Host "The file path '$PSitem' does not exist. Do you want to create it? (Y/N)"
                    if ($createFile -eq 'Y' -or $createFile -eq 'y') {
                        # Create the file and/or directory
                        $directory = [System.IO.Path]::GetDirectoryName($psitem)
                        if (-not [System.IO.Directory]::Exists($directory)) {
                            [System.IO.Directory]::CreateDirectory($directory) | Out-Null
                        }
                        [System.IO.File]::Create($psitem).Dispose()

                        Write-PSFMessage -Level Output -Message  "File created at '$PSitem'"
                    }
                    else {
                        throw "File path creation cancelled. Please provide an existing file path."
                    }
                }

                # Return true if the file exists or was created
                $true
            })]
        [string]$FilePath,

        [Parameter(Mandatory)]
        [string[]]$JobRole
    )

    begin {
        # Import Required Modules
        $ModulesToImport = "ImportExcel", "PSFramework"
        Import-Module $ModulesToImport

        Write-PSFMessage -Level Output -Message "Creating workbook at $FilePath"

        # Helper function
        function New-EmptyCustomObject {
            param (
                [string[]]$PropertyNames
            )
            
            $customObject = [PSCustomObject]@{}
            $customObject | Select-Object -Property $PropertyNames
        }
    }

    process {
        # Define properties for custom objects
        $propertyDefinitions = @{
            Users   = @(
                "FirstName", "LastName", "UserName", "Title", "Department",
                "StreetAddress", "City", "State", "PostalCode", "Country",
                "PhoneNumber", "MobilePhone", "UsageLocation", "License"
            )
            Groups  = @(
                "DisplayName", "PrimarySMTP", "Description", "Type"
            )
            JobRole = @(
                "DisplayName", "PrimarySMTP", "Description", "Type"
            )
            Teams   = @(
                "TeamName", "TeamDescription", "TeamType", "Channel1Name", "Channel1Description", "Channel1Type", "Channel2Name", "Channel2Description", "Channel2Type"
            )
            Sites   = @(
                "Url", "Template", "TimeZone", "Title", "Alias", "SiteType"
            )
        }
        
        # Define custom objects for each worksheet
        $usersObject = New-EmptyCustomObject -PropertyNames $propertyDefinitions.Users
        $groupsObject = New-EmptyCustomObject -PropertyNames $propertyDefinitions.Groups
        $teamsObject = New-EmptyCustomObject -PropertyNames $propertyDefinitions.Teams
        $sitesObject = New-EmptyCustomObject -PropertyNames $propertyDefinitions.Sites

        # Export each worksheet to the workbook
        $usersObject  | Export-Excel -Path $FilePath -WorksheetName "Users" -ClearSheet -AutoSize
        $groupsObject | Export-Excel -Path $FilePath -WorksheetName "Groups" -Append -AutoSize
        $teamsObject | Export-Excel -Path $FilePath -WorksheetName "Teams" -Append -AutoSize
        $sitesObject | Export-Excel -Path $FilePath -WorksheetName "Sites" -Append -AutoSize

        foreach ($JobRoleItem in $JobRole) {
            $RoleObject = New-EmptyCustomObject -PropertyNames $propertyDefinitions.JobRole
            $RoleObject | Export-Excel -Path $FilePath -WorksheetName $JobRoleItem -Append -AutoSize
        }
    }

    end {
        Write-PSFMessage -Level Output -Message "Workbook created successfully at $FilePath"
    }
}
