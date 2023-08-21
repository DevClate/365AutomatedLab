$FunctionFiles = $("$PSScriptRoot\Functions\Public\","$PSScriptRoot\Functions\Private\")| Get-Childitem -file -Recurse -Include "*.ps1" -ErrorAction SilentlyContinue

foreach($FunctionFile in $FunctionFiles){
    try {
        . $FunctionFile.FullName
    }
    catch {
        Write-Error -Message "Failed to import function: '$($FunctionFile.FullName)': $_"
    }
}

Export-ModuleMember -Function $Public.BaseName