BeforeAll {
    # Call Cmdlet
    $commandScriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..\functions\public\Remove-CT365User.ps1'

    . $commandScriptPath
}

Describe 'Remove-CT365User Function' {
    Context 'When provided invalid parameters' {
        It 'Should throw an error for invalid domain format' {
            $filePath = $commandScriptPath
            $domain = "invalid_domain"
            
            { Remove-CT365User -FilePath $filePath -Domain $domain } | Should -Throw
        }
        It 'Should throw an error for invalid file path' {
            $filePath = "C:\Invalid\Path\file.xlsx"
            $domain = "contoso.com"
            
            { Remove-CT365User -FilePath $filePath -Domain $domain } | Should -Throw
        }
    }
}