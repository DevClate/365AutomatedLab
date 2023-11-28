<#
.SYNOPSIS
    Removes all deleted Microsoft 365 groups from the recycle bin.

.DESCRIPTION
    The Remove-CT365AllDeletedM365Groups function connects to a Microsoft 365 tenant and removes all Microsoft 365 groups that have been deleted and are currently in the recycle bin. It requires the PnP.PowerShell module and uses the Connect-PnPOnline cmdlet to establish the connection.

.PARAMETER AdminUrl
    The URL of the Microsoft 365 admin center. This parameter is mandatory and specifies the tenant to connect to.

.EXAMPLE
    PS C:\> Remove-CT365AllDeletedM365Groups -AdminUrl "https://contoso-admin.sharepoint.com"

    This example connects to the Microsoft 365 tenant at contoso-admin.sharepoint.com and removes all deleted Microsoft 365 groups.

.INPUTS
    None. You cannot pipe objects to Remove-CT365AllDeletedM365Groups.

.OUTPUTS
    String. The function outputs messages indicating the status of deletion operations and any errors encountered.

.NOTES
    This function requires the PnP.PowerShell module and PSFramework
    
#>
function Remove-CT365AllDeletedM365Groups {
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
        
        foreach ($module in @('PSFramework', 'PnP.PowerShell')) {
            if (-not (Get-Module -ListAvailable -Name $module)) {
                Install-Module $module -Scope CurrentUser
            }
        }
        
        Connect-PnPOnline -Url $AdminUrl -Interactive
    }

    Process {
        try {
            
            $deletedGroups = Get-PnPDeletedMicrosoft365Group

            foreach ($group in $deletedGroups) {
                Remove-PnPDeletedMicrosoft365Group -Identity $group.Id
                Write-PSFMessage -Message "Deleted group removed: $($group.DisplayName)" -Level Host
            }

            if ($deletedGroups.Count -eq 0) {
                Write-PSFMessage -Message "No deleted groups found." -Level Host
            }
            else {
                Write-PSFMessage -Message "All deleted groups have been removed." -Level Host
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
