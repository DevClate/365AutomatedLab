BeforeAll {
    . $PSScriptRoot/Remove-CT365Group.ps1
}

Describe 'Remove-CT365Group Function' {
    Context 'When provided invalid parameters' {
        It 'Should throw an error for invalid domain format' {
            $filePath = "C:\Path\to\file.xlsx"
            $domain = "invalid_domain"
            
            { Add-CT365User -FilePath $filePath -Domain $domain } | Should -Throw
        }
    }
}