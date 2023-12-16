function Verify-CT365TeamsCreation {
    param(
        [string]$teamName,
        [int]$retryCount = 5,
        [int]$delayInSeconds = 10
    )

    for ($i = 0; $i -lt $retryCount; $i++) {
        $existingTeam = Get-PnPTeamsTeam | Where-Object { $_.DisplayName -eq $teamName }
        if ($existingTeam) {
            return $true
        }
        Start-Sleep -Seconds $delayInSeconds
    }
    return $false
}