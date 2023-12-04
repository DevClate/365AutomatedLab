<#
.SYNOPSIS
Exports Microsoft Teams and their channels to an Excel file, excluding the 'General' channel, and includes the type of each team and channel.

.DESCRIPTION
This function connects to Microsoft 365 to fetch details of Microsoft Teams and their channels, then exports those details to an Excel file, excluding the 'General' channel. The exported details include Team Name, Team Description, Team Type, Channel Names with count numbers, and Channel Types.

.PARAMETER FilePath
The full path to the Excel workbook where the Teams and Channels details will be exported, including the workbook name with an extension of .xlsx.

.PARAMETER AdminUrl
The URL of the SharePoint Online admin center.

.EXAMPLE
Export-CT365ProdTeamsToExcel -FilePath 'C:\Exports\Teams.xlsx' -AdminUrl 'https://yourtenant-admin.sharepoint.com'
This example exports Teams and their channels (excluding 'General') to an Excel file named 'Teams.xlsx' located at 'C:\Exports'.

.NOTES
This function requires the PnP.PowerShell, ImportExcel, and PSFramework modules to be installed.
The user executing this function should have the necessary permissions to read Teams and Channels details from Microsoft 365.

.LINK
[PnP PowerShell](https://pnp.github.io/powershell/)
[ImportExcel Module](https://github.com/dfinke/ImportExcel)
[PSFramework](https://psframework.org/)

#>
function Export-CT365ProdTeamsToExcel {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
            switch ($psitem){
                {-not(([System.IO.Path]::GetExtension($psitem)) -match "(.xlsx)")}{
                    "Invalid file format: '$PSitem'. Use .xlsx"
                }
                Default{
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
                    $channelTypePropertyName = "Channel${channelCount}Type"
                    $teamObject | Add-Member -NotePropertyName $channelPropertyName -NotePropertyValue $channel.DisplayName
                    $teamObject | Add-Member -NotePropertyName $channelTypePropertyName -NotePropertyValue $channel.MembershipType
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
