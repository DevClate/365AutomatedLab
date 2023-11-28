BeforeAll {
    # Call Cmdlet
    $commandScriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..\functions\public\Remove-CT365AllDeletedM365Groups.ps1'

    . $commandScriptPath
}

Describe 'Remove-CT365AllDeletedM365Groups Function' {
    Context 'When provided invalid parameters' {
        It 'Should throw an error for invalid url format' {
            $AdminUrl = "invalid_url"
            
            { Remove-CT365AllDeletedM365Groups -AdminUrl $AdminUrl } | Should -Throw
        }
    }
}