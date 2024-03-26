<#
.SYNOPSIS
    Exports Microsoft Teams and their channels to an Excel file.

.DESCRIPTION
    The Export-CT365ProdTeamsToExcel function connects to SharePoint Online and retrieves information about Microsoft Teams and their channels. It then exports this data to an Excel file. The function requires a valid SharePoint admin URL and the path to an Excel file for exporting the data.

.PARAMETER FilePath
    Specifies the path to the Excel file (.xlsx) where the Teams and Channels data will be exported.

.PARAMETER AdminUrl
    Specifies the SharePoint admin URL for connecting to Microsoft Teams. The URL should match the format 'tenant.sharepoint.com'.

.EXAMPLE
    Export-CT365ProdTeamsToExcel -FilePath "C:\Teams\TeamsData.xlsx" -AdminUrl "contoso.sharepoint.com"

    Exports Microsoft Teams and their channels information to the specified Excel file for the given SharePoint admin URL.

.NOTES
    Requires the PnP.PowerShell, ImportExcel, and PSFramework modules.

    The user executing this script must have SharePoint Online administration permissions.

    The function handles multiple channels per team and exports them in a structured format in the Excel file.

.LINK
    https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/connect-pnponline?view=sharepoint-ps

#>
function Export-CT365ProdTeamsToExcel {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
                $extension = [System.IO.Path]::GetExtension($_)
                $directory = [System.IO.Path]::GetDirectoryName($_)

                if ($extension -ne '.xlsx') {
                    throw "The file $_ is not an Excel file (.xlsx). Please specify a file with the .xlsx extension."
                }
                if (-not (Test-Path -Path $directory -PathType Container)) {
                    throw "The directory $directory does not exist. Please specify a valid directory."
                }
                return $true
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
        [string]$AdminUrl
    )

    begin {
        Import-Module PnP.PowerShell, ImportExcel, PSFramework
        Write-PSFMessage -Level Host -Message "Preparing to export to $(Split-Path -Path $FilePath -Leaf)"
    }

    process {
        try {
            $connectPnPOnlineSplat = @{
                Url = $AdminUrl
                Interactive = $true
                ErrorAction = 'Stop'
            }
            Connect-PnPOnline @connectPnPOnlineSplat
            Write-PSFMessage -Level Verbose -Message "Connected to SharePoint Online"

            # Fetch all teams
            $teams = Get-PnPTeamsTeam
            Write-PSFMessage -Level Verbose -Message "Retrieved Microsoft Teams information"

            # Create an array to hold team and channel information
            $exportData = @()

            foreach ($team in $teams) {
                # Fetch channels for the team, excluding 'General'
                $channels = Get-PnPTeamsChannel -Team $team.DisplayName | Where-Object { $_.DisplayName -ne 'General' }

                $teamObject = [PSCustomObject]@{
                    "TeamName" = $team.DisplayName
                    "TeamDescription" = $team.Description
                    "TeamType" = $team.Visibility
                }

                $channelCount = 1
                foreach ($channel in $channels) {
                    $channelPropertyName = "Channel${channelCount}Name"
                    $channelDescriptionPropertyName = "Channel${channelCount}Description"
                    $channelTypePropertyName = "Channel${channelCount}Type"

                    # Check if the channel type is 'unknownfuturevalue' and convert it to 'shared'
                    $channelType = if ($channel.MembershipType -eq 'unknownfuturevalue') { 'shared' } else { $channel.MembershipType }

                    $teamObject | Add-Member -NotePropertyName $channelPropertyName -NotePropertyValue $channel.DisplayName
                    $teamObject | Add-Member -NotePropertyName $channelDescriptionPropertyName -NotePropertyValue $channel.Description
                    $teamObject | Add-Member -NotePropertyName $channelTypePropertyName -NotePropertyValue $channelType
                    $channelCount++
                }

                $exportData += $teamObject
            }

            # Export data to Excel
            $exportData | Export-Excel -Path $FilePath -WorksheetName "Teams" -AutoSize
            Write-PSFMessage -Level Host -Message "Data exported to Excel successfully"

        } catch {
            Write-PSFMessage -Message "Failed to connect to SharePoint Online" -Level Error 
            return 
        } finally {
            # Disconnect the PnP session
            Disconnect-PnPOnline
            Write-PSFMessage -Level Verbose -Message "Disconnected from Microsoft 365"
        }
    }

    end {
        Write-PSFMessage -Level Host -Message "Export completed. Check the file at $FilePath for the Teams and Channels details."
    }
}
