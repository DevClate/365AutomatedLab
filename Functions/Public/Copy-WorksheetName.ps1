<#
.SYNOPSIS
    This function copies the names of all worksheets in an Excel file to a CSV file.

.DESCRIPTION
    The Copy-WorksheetName function copies the names of all worksheets from a specified Excel file and exports them to a CSV file. It uses the ImportExcel module to handle Excel files.

.PARAMETER FilePath
    The path to the Excel file whose worksheet names you want to copy. The function will throw an error if the Excel file doesn't exist at the specified path.

.PARAMETER outputCsvPath
    The path where the CSV file containing the worksheet names should be created.

.EXAMPLE
    PS C:\> Copy-WorksheetName -FilePath "C:\input.xlsx" -outputCsvPath "C:\output.csv"
    This command will copy all worksheet names from the input.xlsx file and output them to output.csv.

.INPUTS
    System.String
    You can pipe a string to Copy-WorksheetName.

.OUTPUTS
    None. This function does not return any output. It writes the worksheet names to a CSV file.

.NOTES
    This function requires the ImportExcel module. If the module is not installed, you can install it using the Install-Module cmdlet:
    PS C:\> Install-Module -Name ImportExcel
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
    Write-Error "Excel file not found at the specified path: $FilePath"
    return
}

Import-Module ImportExcel

# Import Excel file
$excel = Import-excel -ExcelPackage $FilePath

'"'+((Get-ExcelFileSummary $excel).WorksheetName -join '","')+'"' | Export-Csv -Path $outputCsvPath
}