BeforeAll {
    # Call Cmdlet
    $commandScriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..\functions\public\Set-CT365SPDistinctNumber.ps1'

    . $commandScriptPath
}

Describe 'Set-CT365SPDistinctNumber Function' {
    Context 'When provided invalid parameters' {
        It 'Should throw an error for invalid file path' {
            $filePath = "C:\Invalid\Path\file.xlsx"
                        
            { Set-CT365SPDistinctNumber -FilePath $filePath -WorksheetName $WorksheetName -FindNumber 37 -ReplaceNumber 38 } | Should -Throw
        }
    }
}