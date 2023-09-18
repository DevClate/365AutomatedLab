<#
.SYNOPSIS
Creates new Microsoft 365 Teams and channels based on data from an Excel file.

.DESCRIPTION
The New-CT365Teams function connects to SharePoint Online and creates new Microsoft 365 Teams and channels using the PnP PowerShell Module. 
The teams and channels are defined in an Excel file provided by the user.

.PARAMETER FilePath
Specifies the path to the Excel file that contains the teams and channels information. 
The Excel file should contain a worksheet named "Teams". 
This parameter is mandatory and can be passed through the pipeline.

.PARAMETER AdminUrl
Specifies the SharePoint Online admin URL. 
If not provided, the function will attempt to connect to SharePoint Online interactively.

.PARAMETER ChannelColumns
Specifies the columns in the Excel file that contain the channel names. 
By default, it looks for columns named "Channel1Name" and "Channel2Name".
You can specify other column names if your Excel file is structured differently.

.EXAMPLE
New-CT365Teams -FilePath "C:\path\to\teams.xlsx" -AdminUrl "https://contoso-admin.sharepoint.com"

This example connects to the specified SharePoint Online admin URL, reads the teams and channels from the provided Excel file, and then creates the teams and channels in Microsoft 365.

.EXAMPLE
$filePath = "C:\path\to\teams.xlsx"
$filePath | New-CT365Teams

This example uses pipeline input to provide the file path to the New-365Teams function.

.NOTES
Please submit any feedback and/or recommendations
Prerequisite   : PnP.PowerShell, ImportExcel, PSFramework modules should be installed.

#>
function New-CT365Teams {
    [CmdletBinding()]
    param (
        # Validate the Excel file path.
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
            switch ($psitem){
                {-not([System.IO.File]::Exists($psitem))}{
                    throw "Invalid file path: '$PSitem'."
                }
                {-not(([System.IO.Path]::GetExtension($psitem)) -match "(.xlsx|.xls)")}{
                    "Invalid file format: '$PSitem'. Use .xlsx or .xls."
                }
                Default{
                    $true
                }
            }
        })]
        [string]$FilePath,

        [Parameter(Mandatory=$false)]
        [ValidateScript({
            if ($_ -match '^[a-zA-Z0-9]+\.sharepoint\.[a-zA-Z0-9]+$') {
                $true
            } else {
                throw "The URL $_ does not match the required format."
            }
        })]
        [string]$AdminUrl,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Channel1Name", "Channel2Name")]
        [string[]]$ChannelColumns = @("Channel1Name", "Channel2Name")
    )

    begin {
        # Import required modules.
        $ModulesToImport = "ImportExcel","PnP.PowerShell","PSFramework"
        Import-Module $ModulesToImport
        
        try {
            # Connect to SharePoint Online.
            $connectPnPOnlineSplat = @{
                Url = $AdminUrl
                Interactive = $true
                ErrorAction = 'Stop'
            }
            Connect-PnPOnline @connectPnPOnlineSplat
        }
        catch {
            # Log an error and exit if the connection fails.
            Write-PSFMessage -Message "Failed to connect to SharePoint Online" -Level Error 
            return 
        }
        
        try {
            # Import site data from Excel.
            $SiteData = Import-Excel -Path $FilePath -WorksheetName "Teams"
        }
        catch {
            # Log an error and exit if importing site data fails.
            Write-PSFMessage -Message "Failed to import SharePoint Site data from Excel file." -Level Error 
            return
        }
    }

    process {
        foreach ($team in $SiteData) {
            Write-PSFMessage -Message "Processing team: $($team.TeamName)" -Level Host
        
            $existingTeam = Get-PnPTeamsTeam | Where-Object { $_.DisplayName -eq $team.TeamName }
            
            if (-not $existingTeam) {
                try {
                    # Create the Team
                    New-PnPTeamsTeam -DisplayName $team.TeamName -Description $team.TeamDescription -Visibility Private

                    Write-PSFMessage -Message "Successfully created Team: $($team.TeamName)" -Level Host
                    
                    # Retry mechanism to fetch team details up to 3 times
                    $retryCount = 0
                    $maxRetries = 3
                    $teamResult = $null

                    while ($retryCount -lt $maxRetries -and (-not $teamResult)) {
                        # Wait for a moment to ensure the team is fully provisioned before fetching
                        Start-Sleep -Seconds 15

                        # Attempt to fetch the team details after creation
                        $teamResult = Get-PnPTeamsTeam | Where-Object { $_.DisplayName -eq $team.TeamName }

                        $retryCount++
                    }

                    # If the team wasn't found after all retry attempts, log a warning and skip further processing
                    if (-not $teamResult) {
                        Write-PSFMessage -Message "Team $($team.TeamName) was not found after $maxRetries attempts." -Level Warning
                        continue
                    }
        
        } catch {
                    Write-PSFMessage -Message "Failed to create team $($team.TeamName): $_" -Level Error
                    continue
                }
            }
            else {
                Write-PSFMessage -Message "Team $($team.TeamName) already exists. Fetching team details..." -Level Host
                $teamResult = $existingTeam
                Write-PSFMessage -Message "Fetched team details: $($teamResult)" -Level Host  # Diagnostic message
            }
        
            # Check if teamResult is null
            if (-not $teamResult) {
                Write-PSFMessage -Message "Failed to retrieve or create team $($team.TeamName). Skipping channels creation." -Level Error
                continue
            }
    
            # Create channels based on the provided column names
            foreach ($column in $ChannelColumns) {
                $channelName = $team.$column
                if ($channelName) {
                    Write-PSFMessage -Message "Creating channel: $channelName for team: $($team.TeamName)" -Level Host
                    try {
                        Add-PnPTeamsChannel -Team $teamresult.GroupId -DisplayName $channelName -Description "Channel named $channelName for $($team.TeamName)"
                        Write-PSFMessage -Message "Successfully created channel: $channelName for team: $($team.TeamName)" -Level Host
                    }
                    catch {
                        Write-PSFMessage -Message "Failed to create channel $channelName for team $($team.TeamName): $_" -Level Error
                    }
                }
            }
        }
    }

    end {
        # Disconnect from PnP
        Disconnect-PnPOnline
        Write-PSFMessage "Teams and Channels creation completed."
    }
}