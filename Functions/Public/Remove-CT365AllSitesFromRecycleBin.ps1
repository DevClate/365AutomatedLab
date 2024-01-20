<#
.SYNOPSIS
Removes all sites from the SharePoint Online recycle bin.

.DESCRIPTION
The Remove-CT365AllSitesFromRecycleBin function connects to SharePoint Online using the provided admin URL and then removes all sites from the SharePoint Online recycle bin. This function requires the PSFramework and PnP.PowerShell modules.

.PARAMETER AdminUrl
Specifies the URL of the SharePoint Online admin center. The URL must follow the format 'tenantname.sharepoint.com'.

.EXAMPLE
Remove-CT365AllSitesFromRecycleBin -AdminUrl "contoso-admin.sharepoint.com"

Connects to the SharePoint Online admin center at 'https://contoso-admin.sharepoint.com' and removes all sites from the recycle bin.

.INPUTS
None

.OUTPUTS
None

.NOTES
Please add any suggestions or issues to GitHub Issues

#>
function Remove-CT365AllSitesFromRecycleBin {
    [CmdletBinding()]
    Param (
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

    Begin {
        # Check if required modules are available, otherwise install them
        foreach ($module in @('PSFramework', 'PnP.PowerShell')) {
            if (-not (Get-Module -ListAvailable -Name $module)) {
                Install-Module $module -Scope CurrentUser
            }
        }

        # Connect to SharePoint Online
        Connect-PnPOnline -Url $AdminUrl -Interactive
    }

    Process {
        try {
            # Retrieve sites from recycle bin
            $recycleBinItems = Get-PnPTenantRecycleBinItem

            # Delete sites from recycle bin
            foreach ($item in $recycleBinItems) {
                Remove-PnPTenantDeletedSite -Identity $item.Url -Force
                Write-PSFMessage -Message "Site removed from recycle bin: $($item.Url)" -Level Host
            }

            if ($recycleBinItems.Count -eq 0) {
                Write-PSFMessage -Message "No sites found in the recycle bin." -Level Host
            }
            else {
                Write-PSFMessage -Message "All sites have been removed from the recycle bin." -Level Host
            }
        }
        catch {
            Write-PSFMessage -Message "An error occurred: $_" -Level Error
        }
    }

    End {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
}
