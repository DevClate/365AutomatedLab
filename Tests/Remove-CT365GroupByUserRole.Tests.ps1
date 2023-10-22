BeforeAll {
    # Call Cmdlet
    $commandScriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..\functions\public\Remove-CT365GroupByUserRole.ps1'

    . $commandScriptPath
}

Describe 'Remove-CT365GroupByUserRole Function' {
    Context 'When provided invalid parameters' {
        It 'Should throw an error for invalid domain format' {
            $filePath = $commandScriptPath
            $domain = "invalid_domain"
            
            { Remove-CT365GroupByUserRole -FilePath $filePath -Domain $domain } | Should -Throw
        }
        It 'Should throw an error for invalid file path' {
            $filePath = "C:\Invalid\Path\file.xlsx"
            $domain = "contoso.com"
            
            { Remove-CT365GroupByUserRole -FilePath $filePath -Domain $domain } | Should -Throw
        }
    }
}