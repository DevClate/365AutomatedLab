<#
.SYNOPSIS
Creates new Microsoft Teams and associated channels based on data from an Excel file.

.DESCRIPTION
The New-CT365Teams function connects to Microsoft Teams via PnP PowerShell, reads team and channel information from an Excel file, and creates new Teams and channels as specified. It supports retry logic for team and channel creation and allows specifying a default owner. The function requires the PnP.PowerShell, ImportExcel, and PSFramework modules.

.PARAMETER FilePath
Specifies the path to the Excel file containing the Teams and channel data. The file must be in .xlsx format.

.PARAMETER AdminUrl
Specifies the SharePoint admin URL for the tenant. The URL must match the format 'tenant.sharepoint.com'.

.PARAMETER DefaultOwnerUPN
Specifies the default owner's User Principal Name (UPN) for the Teams and channels.

.EXAMPLE
PS> New-CT365Teams -FilePath "C:\TeamsData.xlsx" -AdminUrl "contoso.sharepoint.com" -DefaultOwnerUPN "admin@contoso.com"

This example creates Teams and channels based on the data in 'C:\TeamsData.xlsx', using 'admin@contoso.com' as the default owner if none is specified in the Excel file.

.NOTES
- Requires the PnP.PowerShell, ImportExcel, and PSFramework modules.
- The Excel file should have a worksheet named 'teams' with appropriate columns for team and channel data.
- The function includes error handling and logging using PSFramework.

#>
function New-CT365Teams {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
                switch ($psitem) {
                    { -not(([System.IO.Path]::GetExtension($psitem)) -match "(.xlsx)") } {
                        "Invalid file format: '$PSitem'. Use .xlsx"
                    }
                    Default {
                        $true
                    }
                }
            })]
        [string]$FilePath,

        [Parameter(Mandatory)]
        [ValidateScript({
            if ($_ -match '^[a-zA-Z0-9]+\.sharepoint\.[a-zA-Z0-9]+$') {
                    $true
                }
                else {
                    throw "The URL $_ does not match the required format."
                }
            })]
        [string]$AdminUrl,


        [Parameter(Mandatory)]
        [string]$DefaultOwnerUPN
    )

    # Check and import required modules
    $requiredModules = @('PnP.PowerShell', 'ImportExcel', 'PSFramework')
    foreach ($module in $requiredModules) {
        try {
            if (-not (Get-Module -ListAvailable -Name $module)) {
                throw "Module $module is not installed."
            }
            Import-Module $module
        }
        catch {
            Write-PSFMessage -Level Warning -Message "[$(Get-Date -Format 'u')] $_.Exception.Message"
            return
        }
    }

    try {
        Connect-PnPOnline -Url $AdminUrl -Interactive
    }
    catch {
        Write-PSFMessage -Level Error -Message "[$(Get-Date -Format 'u')] Failed to connect to PnP Online: $($_.Exception.Message)"
        return
    }

    try {
        $teamsData = Import-Excel -Path $FilePath -WorksheetName "teams"
        $existingTeams = Get-PnPTeamsTeam
    }
    catch {
        Write-PSFMessage -Level Error -Message "[$(Get-Date -Format 'u')] Failed to import data from Excel or retrieve existing teams: $($_.Exception.Message)"
        return
    }

    foreach ($teamRow in $teamsData) {
        try {
            $teamOwnerUPN = if ($teamRow.TeamOwnerUPN) { $teamRow.TeamOwnerUPN } else { $DefaultOwnerUPN }
            $existingTeam = $existingTeams | Where-Object { $_.DisplayName -eq $teamRow.TeamName }

            if ($existingTeam) {
                Write-PSFMessage -Level Host -Message "[$(Get-Date -Format 'u')] Team $($teamRow.TeamName) already exists. Skipping creation."
                continue
            }

            $retryCount = 0
            $teamCreationSuccess = $false
            do {
                try {
                    $teamId = New-PnPTeamsTeam -DisplayName $teamRow.TeamName -Description $teamRow.TeamDescription -Visibility $teamRow.TeamType -Owners $teamOwnerUPN
                    if (Verify-CT365TeamsCreation -teamName $teamRow.TeamName) {
                        Write-PSFMessage -Level Host -Message "[$(Get-Date -Format 'u')] Verified creation of Team: $($teamRow.TeamName)"
                        $teamCreationSuccess = $true
                        break
                    }
                    else {
                        Write-PSFMessage -Level Warning -Message "[$(Get-Date -Format 'u')] Team $($teamRow.TeamName) creation reported but not verified. Retrying..."
                    }
                }
                catch {
                    Write-PSFMessage -Level Warning -Message "[$(Get-Date -Format 'u')] Attempt $retryCount to create team $($teamRow.TeamName) failed: $($_.Exception.Message)"
                }
                $retryCount++
                Start-Sleep -Seconds 5
            } while ($retryCount -lt 5)

            if (-not $teamCreationSuccess) {
                Write-PSFMessage -Level Error -Message "[$(Get-Date -Format 'u')] Failed to create and verify Team: $($teamRow.TeamName) after multiple retries."
                continue
            }

            for ($i = 1; $i -le 4; $i++) {
                $channelName = $teamRow."Channel${i}Name"
                $channelType = $teamRow."Channel${i}Type"
                $channelDescription = $teamRow."Channel${i}Description"
                $channelOwnerUPN = if ($teamRow."Channel${i}OwnerUPN") { $teamRow."Channel${i}OwnerUPN" } else { $DefaultOwnerUPN }

                if ($channelName -and $channelType) {
                    $retryCount = 1
                    $channelCreationSuccess = $false
                    do {
                        try {
                            Add-PnPTeamsChannel -Team $teamId -DisplayName $channelName -Description $channelDescription -ChannelType $channelType -OwnerUPN $channelOwnerUPN
                            Write-PSFMessage -Level Host -Message "[$(Get-Date -Format 'u')] Created Channel: $channelName in Team: $($teamRow.TeamName) with Type: $channelType and Description: $channelDescription"
                            $channelCreationSuccess = $true
                            break
                        }
                        catch {
                            Write-PSFMessage -Level Warning -Message "[$(Get-Date -Format 'u')] Attempt $retryCount to create channel $channelName in Team: $($teamRow.TeamName) failed: $($_.Exception.Message)"
                            $retryCount++
                            Start-Sleep -Seconds 10
                        }
                    } while ($retryCount -lt 5)

                    if (-not $channelCreationSuccess) {
                        Write-PSFMessage -Level Error -Message "[$(Get-Date -Format 'u')] Failed to create Channel: $channelName in Team: $($teamRow.TeamName) after multiple retries."
                    }
                }
            }
        }
        catch {
            Write-PSFMessage -Level Error -Message "[$(Get-Date -Format 'u')] Error processing team $($teamRow.TeamName): $($_.Exception.Message)"
        }
    }

    try {
        Disconnect-PnPOnline
    }
    catch {
        Write-PSFMessage -Level Error -Message "[$(Get-Date -Format 'u')] Error disconnecting PnP Online: $($_.Exception.Message)"
    }
}
