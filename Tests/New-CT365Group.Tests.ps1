BeforeAll {
    # Call Cmdlet
    $commandScriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..\functions\public\New-CT365Group.ps1'

    . $commandScriptPath
}

Describe 'New-CT365Group Function' {
    Context 'When provided invalid parameters' {
        It 'Should throw an error for invalid domain format' {
            $filePath = $commandScriptPath
            $domain = "invalid_domain"
            
            { New-CT365Group -FilePath $filePath -Domain $domain } | Should -Throw
        }
        It 'Should throw an error for invalid file path' {
            $filePath = "C:\Invalid\Path\file.xlsx"
            $domain = "contoso.com"
            
            { New-CT365Group -FilePath $filePath -Domain $domain } | Should -Throw
        }
    }
}