<#
.SYNOPSIS
Deletes SharePoint Online sites based on the data from an Excel file.

.DESCRIPTION
The `Remove-365CTSharePointSite` function connects to SharePoint Online(PnP) using the provided admin URL and imports site data from the specified Excel file. It then attempts to delete each site based on the data.

.PARAMETER FilePath
The path to the Excel file containing the SharePoint site data. The file must exist and have an .xlsx extension.

.PARAMETER AdminUrl
The SharePoint Online admin URL.

.PARAMETER Domain
The domain information required for the SharePoint site creation.

.EXAMPLE
Remove-CT365SharePointSite -FilePath "C:\path\to\file.xlsx" -AdminUrl "https://domainname.sharepoint.com" -Domain "contoso.com"

This example removes SharePoint sites using the data from the "file.xlsx" and connects to SharePoint Online using the provided admin URL.

.NOTES
To use this, please make sure that you the sites have been created at least 6 minutes prior, or it won't work. Also it will say "Group not found" but still works as of 10/23/2023. Open issue in GitHub for more information. 
Make sure you have the necessary modules installed: ImportExcel, PnP.PowerShell, and PSFramework.

.LINK
https://docs.microsoft.com/powershell/module/sharepoint-pnp/new-pnpsite
#>
function Remove-CT365SharePointSite {
    [CmdletBinding()]
    param (
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

        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
            switch ($psitem) {
                {$psitem -notmatch '^(((?!-))(xn--|_)?[a-z0-9-]{0,61}[a-z0-9]{1,1}\.)*(xn--)?[a-z]{2,}(?:\.[a-z]{2,})+$'}{
                    throw "The provided domain is not in the correct format."
                }
                Default {
                    $true
                }
            }
        })]
        [string]$Domain,

        [switch]$PermanentlyDelete
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
            
            # Join Admin URL and Site Url
            $siteUrl = "https://$AdminUrl/sites/$($site.Url)"
            
            try {
                # Set the message target to the site's title.
                $PSDefaultParameterValues["Write-PSFMessage:Target"] = $site.Title

                # Log a message indicating site creation.
                Write-PSFMessage -Message "Deleting SharePoint Site: '$($site.Title)'"

                # If PermanentlyDelete switch is set, prioritize those actions
                if ($PermanentlyDelete) {
                    switch -Regex ($site.SiteType) {
                        "^TeamSite$" {
                            $removePnPM365GroupPermSplat = @{
                                Identity = $site.Title
                                ErrorAction = 'Stop'
                            }
                            $removePnPM365GroupSplat = @{
                                Identity = $site.Title
                                ErrorAction = 'Stop'
                            }
                            $removePnPTenantSiteSplat = @{
                                Url = $siteUrl
                                ErrorAction = 'Stop'
                                Force = $true
                            }
                            Remove-PnPMicrosoft365Group @removePnPM365GroupSplat
                            Write-PSFMessage -Message "Soft deletion of Team Group: '$($site.Title)'"
                            Start-Sleep -Seconds 120
                            Remove-PnPDeletedMicrosoft365Group @removePnPM365GroupPermSplat
                            Write-PSFMessage -Message "Permanently deleted Team Group: '$($site.Title)'"
                            Remove-PnPTenantDeletedSite @removePnPTenantSiteSplat
                            Write-PSFMessage -Message "Permanently deleted SharePoint Site: '$($siteUrl)'"
                            continue
                        }
                        "^(CommunicationSite|TeamSiteWithoutMicrosoft365Group)$" {
                            $removePnPTenantSiteSplat = @{
                                Url = $siteUrl
                                ErrorAction = 'Stop'
                                Force = $true
                            }
                            Remove-PnPTenantDeletedSite @removePnPTenantSiteSplat
                            Write-PSFMessage -Message "Permanently deleted SharePoint Site: '$($siteUrl)'"
                            continue
                        }
                        default {
                            Write-PSFMessage "Unknown site type: $($site.SiteType) for site $($site.Title). Skipping." -Level Error
                            continue
                        }
                    }
                }
                
                # If not permanently deleting, proceed with regular deletion
                switch -Regex ($site.SiteType) {
                    "^TeamSite$" {
                        $removePnPM365GroupSplat = @{
                            Identity = $site.Title
                            ErrorAction = 'Stop'
                        }
                        Remove-PnPMicrosoft365Group @removePnPM365GroupSplat
                        Write-PSFMessage -Message "Successfully deleted Group Site: '$($site.Title)'"

                    }
                    "^(CommunicationSite|TeamSiteWithoutMicrosoft365Group)$" {
                        $removePnPSiteSplat = @{
                            Url = $siteUrl
                            ErrorAction = "Stop"
                            Force = $true
                        }
                        remove-PnPTenantSite @removePnPSiteSplat
                        Write-PSFMessage -Message "Successfully deleted SharePoint Site: '$($siteUrl)'" 
                    }
                    default {
                    # Log an error for unknown site types and skip to the next site.
                    Write-PSFMessage "Unknown site type: $($site.SiteType) for site $($site.Title). Skipping." -Level Error
                    # Continue to the next site in the loop.
                    continue
                }
            }
        }
        catch {
            # Log an error message if site removal fails and continue to the next site.
            Write-PSFMessage -Message "Could not delete SharePoint Site: '$($site.Title)' Skipping" -Level Error
            Write-PSFMessage -Message $Psitem.Exception.Message -Level Error
            Continue
        }
    }
}

    end {
        # Log a message indicating completion of the SharePoint site creation process.
        Write-PSFMessage "SharePoint site deletion process completed."
            
        # Disconnect from SharePoint Online.
        Disconnect-PnPOnline
    }
}