<#
.SYNOPSIS
Removes Microsoft 365 Teams based on the provided Excel data.

.DESCRIPTION
The Remove-CT365Teams function connects to SharePoint Online, reads a list of Teams from an Excel file, and then removes each team. The function provides feedback on the process using the Write-PSFMessage cmdlet.

.PARAMETER FilePath
The path to the Excel file containing the list of Teams to remove. The file should have a worksheet named "Teams" and must be in .xlsx format.

.PARAMETER AdminUrl
The URL of the SharePoint Online admin center. This is used for connecting to SharePoint Online.

.PARAMETER ChannelColumns
Array of channel column names. The default values are "Channel1Name" and "Channel2Name".

.EXAMPLE
Remove-CT365Teams -FilePath "C:\Path\To\File.xlsx" -AdminUrl "https://yourtenant-admin.sharepoint.com"

This example will connect to the SharePoint Online admin center using the provided AdminUrl, read the Teams from the specified Excel file, and proceed to remove each team.

.NOTES
- Ensure you have the necessary modules ("ImportExcel","PnP.PowerShell","PSFramework","Microsoft.Identity.Client") installed before running this function.
- Always backup your Teams data before using this function to avoid unintended data loss.
- This function has a built-in delay of 5 seconds between team removals to ensure proper deletion.

#>
function Remove-CT365Teams {
    [CmdletBinding()]
    param (
        # Validate the Excel file path.
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
            switch ($psitem){
                {-not([System.IO.File]::Exists($psitem))}{
                    throw "Invalid file path: '$PSitem'."
                }
                {-not(([System.IO.Path]::GetExtension($psitem)) -match "(.xlsx)")}{
                    "Invalid file format: '$PSitem'. Use .xlsx"
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
        [string[]]$ChannelColumns = @("Channel1Name", "Channel2Name")
    )

    begin {
        # Import required modules.
        $ModulesToImport = "ImportExcel","PnP.PowerShell","PSFramework","Microsoft.Identity.Client"
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
            # Get the GroupId for the Team based on its name
            $teamObj = Get-PnPTeamsTeam | Where-Object { $_.DisplayName -eq $team.TeamName }
            
            # Continue to the next iteration if no matching team is found
            if (-not $teamObj) { continue }
    
            $teamGroupId = $teamObj.GroupId
    
            # Display the team name that's being removed using Write-PSFMessage
            Write-PSFMessage -Message "Removing team: $($team.TeamName) with GroupId: $teamGroupId" -Level Host
                    
            # Remove the Team using the GroupId
            Remove-PnPTeamsTeam -Identity $teamGroupId -Force
    
            # Introduce a delay of 5 seconds
            Start-Sleep -Seconds 5
    
            # Check if the team still exists
            $teamCheck = Get-PnPTeamsTeam | Where-Object { $_.GroupId -eq $teamGroupId }
            
            # Provide feedback based on team removal status
            $messageLevel = if ($teamCheck) { "Warning" } else { "Host" }
            $messageContent = if ($teamCheck) { "Failed to remove team: $($team.TeamName)" } else { "Successfully removed team: $($team.TeamName)" }
            
            Write-PSFMessage -Message $messageContent -Level $messageLevel
        }
    }
    

    end {
        # Disconnect from PnP
        Disconnect-PnPOnline
        Write-PSFMessage -Message "Teams removal completed."
    }
}