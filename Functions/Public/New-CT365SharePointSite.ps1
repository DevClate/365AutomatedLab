<#
.SYNOPSIS
Creates new SharePoint Online sites based on the data from an Excel file.

.DESCRIPTION
The `New-365CTSharePointSite` function connects to SharePoint Online(PnP) using the provided admin URL and imports site data from the specified Excel file. It then attempts to create each site based on the data.

.PARAMETER FilePath
The path to the Excel file containing the SharePoint site data. The file must exist and have an .xlsx extension.

.PARAMETER AdminUrl
The SharePoint Online admin URL.

.PARAMETER Domain
The domain information required for the SharePoint site creation.

.EXAMPLE
New-CT365SharePointSite -FilePath "C:\path\to\file.xlsx" -AdminUrl "https://admin.sharepoint.com" -Domain "contoso.com"

This example creates SharePoint sites using the data from the "file.xlsx" and connects to SharePoint Online using the provided admin URL.

.NOTES
Make sure you have the necessary modules installed: ImportExcel, PnP.PowerShell, and PSFramework.

.LINK
https://docs.microsoft.com/powershell/module/sharepoint-pnp/new-pnpsite
#>
function New-CT365SharePointSite {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
            #making sure the Filepath leads to a file and not a folder and has a proper extension
            switch ($psitem){
                {-not([System.IO.File]::Exists($psitem))}{
                    throw "The file path '$PSitem' does not lead to an existing file. Please verify the 'FilePath' parameter and ensure that it points to a valid file (folders are not allowed).                "
                }
                {-not(([System.IO.Path]::GetExtension($psitem)) -match "(.xlsx|.xls)")}{
                    "The file path '$PSitem' does not have a valid Excel format. Please make sure to specify a valid file with a .xlsx or .xls extension and try again."
                }
                Default{
                    $true
                }
            }
        })]
        [string]$FilePath,

        [Parameter(Mandatory)]
        [string]$AdminUrl,

        [Parameter(Mandatory)]
        [string]$Domain
    )
    begin{
        $PSDefaultParameterValues = @{
            "Write-PSFMessage:Level"    = "OutPut"
            "Write-PSFMessage:Target"   = "Preperation"
        }

        # Import the required modules
        $ModulesToImport = "ImportExcel","PnP.PowerShell","PSFramework"
        Import-Module $ModulesToImport

        try {
            $connectPnPOnlineSplat = @{
                Url = $AdminUrl
                Interactive = $true
                ErrorAction = 'Stop'
            }
            Connect-PnPOnline  @connectPnPOnlineSplat
        }
        catch {
            Write-PSFMessage -Message "Failed to connect to sharepoint online" -Level Error 
            return 
        }

        try {
            $SiteData = Import-Excel -Path $FilePath -WorksheetName "Sites"
        } catch {
            Write-PSFMessage -Message "Failed to import Sharepoint Site data from Excel file." -Level Error 
            return
        }

    }

    process {
        foreach ($site in $siteData) {
            
            $siteurl = "https://$AdminUrl/sites/$($site.Url)"
            $PSDefaultParameterValues["Write-PSFMessage:Target"] = $site.Title
            Write-PSFMessage -Message "Creating Sharepoint Site: '$($site.Title)'"
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
                    $newPnPSiteSplat.add("Url",$siteurl)
                }
                default {
                    Write-PSFMessage "Unknown site type: $($site.SiteType) for site $($site.Title). Skipping." -Level Error
                    continue
                }
            }
            try {
                New-PnPSite @newPnPSiteSplat 
                Write-PSFMessage -Message "Created Sharepoint Site: '$($site.Title)'"
            }
            catch {
                Write-PSFMessage -Message "Could not create Sharepoint Site: '$($site.Title)' Skipping" -Level Error
                Write-PSFMessage -Message $Psitem.Exception.Message -Level Error
                Continue
            }
        }
    }

    end {
        Write-PSFMessage "SharePoint site creation process completed."
        # Disconnect from SharePoint Online
        Disconnect-PnPOnline
    }
}