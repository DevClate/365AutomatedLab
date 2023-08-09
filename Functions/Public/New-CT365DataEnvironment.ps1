<#
.SYNOPSIS
Creates a new Excel workbook with predefined worksheets for Office 365 data.

.DESCRIPTION
The New-CT365DataEnvironment function creates a new Excel workbook with predefined worksheets. 
The workbook contains a "Users" worksheet and a "Groups" worksheet, and additional worksheets based on the provided JobRole parameter.

.PARAMETER WorkbookName
The name of the workbook file to be created. It must be a .xls or .xlsx file.

.PARAMETER FilePath
The path where the workbook will be created. The provided path must exist.

.PARAMETER JobRole
An array of strings specifying additional worksheets to be created in the workbook. 
Each string will be used as the name of a new worksheet.

.EXAMPLE
New-CT365DataEnvironment -WorkbookName "myworkbook.xlsx" -FilePath "C:\temp" -JobRole "Manager","Employee"

This command creates an Excel workbook named "myworkbook.xlsx" in the "C:\temp" directory.
The workbook will contain the "Users" and "Groups" worksheets, as well as a "Manager" and "Employee" worksheet.

.EXAMPLE
New-CT365DataEnvironment -WorkbookName "365DataEnvironment.xlsx" -FilePath "C:\365DevEnvironment" -JobRole "NY-ITManager","CA-AccountsPayable","FL-HumanResources"

This command creates an Excel workbook named "365DataEnvironment.xlsx" in the "C:\365DevEnvironment" directory.
The workbook will contain the "Users" and "Groups" worksheets, as well as a "NY-ITManager", "CA-AccountsPayable", and "FL-HumanResources" worksheet.

.INPUTS
System.String

.OUTPUTS
None. This cmdlet does not return any output.

.NOTES
The JobRole Parameter can be a single job role or a location and job role. If you are a company with only one location, you do not need to put a location in.
#>
function New-CT365DataEnvironment {
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

        [Parameter(Mandatory)]
        [string[]]$JobRole

    )

    begin {
        # Import Required Modules
        $ModulesToImport = "ImportExcel","PSFramework"
        Import-Module $ModulesToImport
        
        $Path = Join-Path -Path $filepath -ChildPath $workbookname
        
        Write-PSFMessage -Level Output -Message "Creating workbook $WorkbookName" -Target $WorkbookName
        
        #helper function
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
        $propertyNamesUsers = "FirstName", "LastName", "UserName", "Title", "Department", "StreetAddress", "City", "State", "PostalCode", "Country", "PhoneNumber", "MobilePhone", "UsageLocation", "License"
        $propertyNamesGroups = "DisplayName", "PrimarySMTP", "Description", "Owner", "Type"
        $propertyJobRole = "DisplayName", "PrimarySMTP", "Description", "Type"
        
        # Define custom objects for each worksheet
        $usersObject  = New-EmptyCustomObject -PropertyNames $propertyNamesUsers
        $groupsObject = New-EmptyCustomObject -PropertyNames $propertyNamesGroups

        # Export each worksheet to the workbook
        $usersObject | Export-Excel -Path $Path -WorksheetName "Users" -ClearSheet
        $groupsObject | Export-Excel -Path $Path -WorksheetName "Groups" -Append 

        foreach($JobRoleItem in $JobRole){
            $RoleObject = New-EmptyCustomObject -PropertyNames $propertyJobRole
            $RoleObject | Export-Excel -Path $Path -WorksheetName $JobRoleItem -Append
        }
    }

    end {
        Write-PSFMessage -Level Output -Message "Workbook $WorkbookName created successfully" -Target $WorkbookName
    }
}