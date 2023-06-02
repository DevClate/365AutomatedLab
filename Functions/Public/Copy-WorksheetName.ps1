# Copies worksheet names, and comma seperates them, with quotes to use in validateset
# Testing

function Copy-WorksheetName {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$excelFilePath,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$outputCsvPath
    )

if (!(Test-Path $ExcelFilePath)) {
    Write-Error "Excel file not found at the specified path: $ExcelFilePath"
    return
}

Import-Module ImportExcel

# Open the Excel file
$excel = Import-excel -ExcelPackage $excelFilePath

'"'+((Get-ExcelFileSummary $excel).WorksheetName -join '","')+'"' | Export-Csv -Path $outputCsvPath
}