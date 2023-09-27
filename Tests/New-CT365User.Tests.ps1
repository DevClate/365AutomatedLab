BeforeAll {
    . $PSScriptRoot/New-CT365User.ps1
}

Describe 'New-CT365User Function' {
    Context 'When provided invalid parameters' {
        It 'Should throw an error for invalid domain format' {
            $filePath = "C:\Path\to\file.xlsx"
            $domain = "invalid_domain"
            
            { New-CT365User -FilePath $filePath -Domain $domain } | Should -Throw
        }
        It 'Should throw an error for invalid file path' {
            $filePath = "C:\Invalid\Path\file.xlsx"
            $domain = "contoso.com"
            
            { New-CT365User -FilePath $filePath -Domain $domain } | Should -Throw
        }
    }
}