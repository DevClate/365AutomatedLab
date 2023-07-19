<#
.SYNOPSIS
This function copies the names of all the worksheets in an Excel file and exports them into a CSV file.

.DESCRIPTION
The function Copy-WorksheetName takes two parameters, the file path of the Excel file and the output path of the CSV file. It reads the Excel file, extracts the names of all worksheets, and exports these names into a CSV file. 

.PARAMETER FilePath
The path to the Excel file. This is a mandatory parameter and it accepts pipeline input.

.PARAMETER outputCsvPath
The path where the CSV file will be created. This is a mandatory parameter and it accepts pipeline input.

.EXAMPLE
Copy-WorksheetName -FilePath "C:\path\to\your\excel\file.xlsx" -outputCsvPath "C:\path\to\your\output\file.csv"

This will read the Excel file located at "C:\path\to\your\excel\file.xlsx", get the names of all worksheets, and export these names to a CSV file at "C:\path\to\your\output\file.csv".

.NOTES
This function requires the ImportExcel module to be installed. If not already installed, you can install it by running Install-Module -Name ImportExcel.

.INPUTS
System.String. You can pipe a string that contains the file path to this cmdlet.

.OUTPUTS
System.String. This cmdlet outputs a CSV file containing the names of all worksheets in the Excel file.

.LINK
https://github.com/dfinke/ImportExcel
#>
function Copy-WorksheetName {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$FilePath,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$outputCsvPath
    )

if (!(Test-Path $FilePath)) {
    Write-PSFMessage -Level Error -Message "Excel file not found at the specified path: $FilePath" -Target $FilePath
    return
}

Import-Module ImportExcel
Import-Module PSFramework

# Import Excel file
$excel = Import-excel -ExcelPackage $FilePath

'"'+((Get-ExcelFileSummary $excel).WorksheetName -join '","')+'"' | Export-Csv -Path $outputCsvPath
}