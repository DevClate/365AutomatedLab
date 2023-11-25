<#
.SYNOPSIS
    Replaces specified numbers in an Excel worksheet, excluding certain columns.

.DESCRIPTION
    The Set-CT365SPDistinctNumber function opens an Excel file and replaces occurrences of a specified number in a given worksheet. 
    It excludes specific columns ("Template" and "TimeZone") from this operation. 
    The primary reason for this is to be able to create new Sharepoint Team sites while the others are deleting.

.PARAMETER FilePath
    The path to the Excel file that contains the worksheet to be modified.

.PARAMETER WorksheetName
    The name of the worksheet within the Excel file where the replacements will be made.

.PARAMETER FindNumber
    The number to find in the worksheet. This number will be replaced wherever it is found, except in the excluded columns.

.PARAMETER ReplaceNumber
    The number that will replace the FindNumber in the worksheet.

.EXAMPLE
    Set-CT365SPDistinctNumber -FilePath "C:\Documents\example.xlsx" -WorksheetName "Sheet1" -FindNumber "36" -ReplaceNumber "37"

    This command replaces all occurrences of the number 36 with 37 in the worksheet named "Sheet1" of the Excel file located at "C:\Documents\example.xlsx", excluding the "Template" and "TimeZone" columns.

.INPUTS
    None. You cannot pipe objects to Set-CT365SPDistinctNumber.

.OUTPUTS
    None. This function does not generate any output.

.NOTES
    This function requires the ImportExcel module to be installed.

.LINK
    https://github.com/dfinke/ImportExcel - The ImportExcel PowerShell module

#>
function Set-CT365SPDistinctNumber {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter(Mandatory)]
        [string]$WorksheetName,

        [Parameter(Mandatory)]
        [string]$FindNumber,

        [Parameter(Mandatory)]
        [string]$ReplaceNumber
    )

    # Import the ImportExcel module
    Import-Module ImportExcel

    # Open the Excel package
    $excelPackage = Open-ExcelPackage -Path $FilePath

    try {
        $worksheet = $excelPackage.Workbook.Worksheets[$WorksheetName]
        if ($null -eq $worksheet) {
            throw "Worksheet '$WorksheetName' not found."
        }

        # Get the indices of the columns to exclude
        $excludedColumns = @("Template", "TimeZone").ForEach({
            $worksheet.Dimension.Start.Column..$worksheet.Dimension.End.Column |
            Where-Object { $worksheet.Cells[1, $_].Text -eq $_ } | 
            ForEach-Object { [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($_) }
        })

        # Initialize a counter for replacements
        $replacementCount = 0

        # Find and replace the numbers, skipping the excluded columns
        $worksheet.Cells.Where({
            $_.Value -like "*$FindNumber*" -and 
            -not ($excludedColumns -contains [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($_.Start.Column))
        }).ForEach({
            if ($_ -ne $null -and $null -ne $_.Value -and $_.Value -like "*$FindNumber*") {
                $_.Value = $_.Value -replace $FindNumber, $ReplaceNumber
                $replacementCount++
            }
        })

        # Check if the number of replacements is as expected
        if ($replacementCount -eq 0) {
            throw "No replacements were made for the number '$FindNumber'."
        } elseif ($replacementCount -ne 12) {
            Write-PSFMessage -Message "Unexpected number of replacements: $replacementCount. Expected 12." -Level Error
        } else {
            Write-PSFMessage -Message "Exactly 12 replacements were made for the number '$FindNumber'." -Level Host
        }

        # Save and close the Excel package
        Close-ExcelPackage $excelPackage
    }
    catch {
        $excelPackage.Dispose()
        throw
    }
}