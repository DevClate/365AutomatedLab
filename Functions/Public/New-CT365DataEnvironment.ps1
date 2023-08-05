# Import the module
Import-Module ImportExcel

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
    }

    process {
        # Define a custom object for each worksheet
        $usersObj = New-Object -TypeName PSCustomObject -Property ([ordered]@{
            "FirstName" = $null
            "LastName" = $null
            "UserName" = $null
            "Title" = $null
            "Department" = $null
            "StreetAddress" = $null
            "City" = $null
            "State" = $null
            "PostalCode" = $null
            "Country" = $null
            "PhoneNumber" = $null
            "MobilePhone" = $null
            "UsageLocation" = $null
            "License" = $null
        })

        $groupsObj = New-Object -TypeName PSCustomObject -Property ([ordered]@{
            "DisplayName" = $null
            "PrimarySMTP" = $null
            "Description" = $null
            "Owner" = $null
            "Type" = $null
        })

        # Export each worksheet to the workbook
        $usersObj | Export-Excel -Path $Path -WorksheetName "Users" -ClearSheet
        $groupsObj | Export-Excel -Path $Path -WorksheetName "Groups" -Append 

        foreach($JobRoleItem in $JobRole){
            $customObj = New-Object -TypeName PSCustomObject -Property ([ordered]@{
                "DisplayName" = $null
                "PrimarySMTP" = $null
                "Description" = $null
                "Type" = $null
            })

            $customObj | Export-Excel -Path $Path -WorksheetName $JobRoleItem -Append
        }

    }

    end {
        Write-PSFMessage -Level Output -Message "Workbook $WorkbookName created successfully" -Target $WorkbookName
    }
}
