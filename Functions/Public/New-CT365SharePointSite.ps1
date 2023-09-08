function New-365SharePointSite {
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

        # SharePoint Online admin URL.
        [Parameter(Mandatory)]
        [string]$AdminUrl,

        # Domain information.
        [Parameter(Mandatory)]
        [string]$Domain
    )

    begin {
        # Set default message parameters.
        $PSDefaultParameterValues = @{
            "Write-PSFMessage:Level"    = "OutPut"
            "Write-PSFMessage:Target"   = "Preparation"
        }

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
            $SiteData = Import-Excel -Path $FilePath -WorksheetName "Sites"
        }
        catch {
            # Log an error and exit if importing site data fails.
            Write-PSFMessage -Message "Failed to import SharePoint Site data from Excel file." -Level Error 
            return
        }
    }

    process {
        foreach ($site in $siteData) {
            # Set the message target to the site's title.
            $PSDefaultParameterValues["Write-PSFMessage:Target"] = $site.Title

            # Log a message indicating site creation.
            Write-PSFMessage -Message "Creating SharePoint Site: '$($site.Title)'"

            # Initialize parameters for creating a new SharePoint site.
            $newPnPSiteSplat = @{
                Type = $null
                TimeZone = $site.Timezone
                Title = $site.Title
                ErrorAction = "Stop"
            }

            switch -Regex ($site.SiteType) {
                "^TeamSite$" {
                    $newPnPSiteSplat.Type = $PSItem 
                    $newPnPSiteSplat.add("Alias",$site.Alias)
                }
                "^(CommunicationSite|TeamSiteWithoutMicrosoft365Group)$" {
                    $newPnPSiteSplat.Type = $PSItem 
                    $newPnPSiteSplat.add("Url",$site.Url)
                }
                default {
                    # Log an error for unknown site types and skip to the next site.
                    Write-PSFMessage "Unknown site type: $($site.SiteType) for site $($site.Title). Skipping." -Level Error
                    # Continue to the next site in the loop.
                    continue
                }
            }

            try {
                # Attempt to create a new SharePoint site using specified parameters.
                New-PnPSite @newPnPSiteSplat 
                Write-PSFMessage -Message "Created SharePoint Site: '$($site.Title)'"
            }
            catch {
                # Log an error message if site creation fails and continue to the next site.
                Write-PSFMessage -Message "Could not create SharePoint Site: '$($site.Title)' Skipping" -Level Error
                Write-PSFMessage -Message $Psitem.Exception.Message -Level Error
                Continue
            }
        }
    }

    end {
        # Log a message indicating completion of the SharePoint site creation process.
        Write-PSFMessage "SharePoint site creation process completed."
        
        # Disconnect from SharePoint Online.
        Disconnect-PnPOnline
    }
}