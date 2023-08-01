BeforeAll {
    . $PSScriptRoot/Add-CT365GroupByTitle.ps1
}

Describe 'Add-CT365GroupByTitle Function' {
    Context 'When provided invalid parameters' {
        It 'Should throw an error for invalid domain format' {
            $filePath = "C:\Path\to\file.xlsx"
            $domain = "invalid_domain"
            
            { Add-CT365User -FilePath $filePath -Domain $domain } | Should -Throw
        }
        It 'Should throw an error for invalid file path' {
            $filePath = "C:\Invalid\Path\file.xlsx"
            $domain = "contoso.com"
            
            { Add-CT365User -FilePath $filePath -Domain $domain } | Should -Throw
        }
    }
}