BeforeAll {
    # Call Cmdlet
    $commandScriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..\functions\public\Remove-CT365Teams.ps1'

    . $commandScriptPath
}

Describe 'Remove-CT365Teams Function' {
    Context 'When provided invalid parameters' {
        It 'Should throw an error for invalid domain format' {
            $filePath = $commandScriptPath
            $AdminUrl = "invalid_domain"
            
            { Remove-CT365Teams -FilePath $filePath -AdminUrl $AdminUrl } | Should -Throw
        }
        It 'Should throw an error for invalid file path' {
            $filePath = "C:\Invalid\Path\file.xlsx"
            $AdminUrl = "contoso.com"
            
            { Remove-CT365Teams -FilePath $filePath -AdminUrl $AdminUrl } | Should -Throw
        }
    }
}