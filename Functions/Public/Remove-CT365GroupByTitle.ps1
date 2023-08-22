<#
.SYNOPSIS
Removes a user from Office 365 groups as specified in an Excel file.

.DESCRIPTION
This function removes a user from different types of Office 365 groups as specified in an Excel file.
The function uses the ExchangeOnlineManagement, ImportExcel, Microsoft.Graph.Groups, and Microsoft.Graph.Users modules.

The function first connects to Exchange Online and Microsoft Graph using the UserPrincipalName provided. 
It then imports data from an Excel file and iterates through each row. For each row, it removes the user from the group based on the group type.

.PARAMETER FilePath
This mandatory parameter specifies the path to the Excel file which contains information about the groups from which the user should be removed.

.PARAMETER UserEmail
This mandatory parameter specifies the email of the user who should be removed from the groups.

.PARAMETER Domain
This mandatory parameter specifies the domain of the user. The domain is used to construct the group name.

.PARAMETER UserRole
This mandatory parameter specifies the user's role. It should be either "NY-IT" or "NY-HR". This parameter is used to identify the worksheet in the Excel file to import.

.EXAMPLE
Remove-CT365GroupByTitle -FilePath "C:\Path\to\file.xlsx" -UserEmail "johndoe@example.com" -Domain "example.com" -UserRole "NY-IT"
This example removes the user "johndoe@example.com" from the groups specified in the "NY-IT" worksheet of the Excel file at "C:\Path\to\file.xlsx".

.NOTES
The Excel file should have columns for PrimarySMTP, GroupType, and DisplayName. These columns are used to get information about the groups from which the user should be removed.

The function supports the following group types: 365Group, 365Distribution, 365MailEnabledSecurity, and 365Security. For each group type, it uses a different cmdlet to remove the user from the group.

Connect-MgGraph -Scopes "Group.ReadWrite.All","Directory.AccessAsUser.All" - is needed to connect with correct scopes

.LINK

https://docs.microsoft.com/en-us/powershell/module/microsoft.graph.groups/?view=graph-powershell-1.0

.LINK

https://www.powershellgallery.com/packages/ImportExcel

#>
function Remove-CT365GroupByTitle {
    [CmdletBinding(SupportsShouldProcess)]
    param(
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
        [string]$UserEmail,

        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
            # Check if the domain fits the pattern
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
        
        [Parameter(Mandatory)]
        [string]$UserRole
    )

    # Import Required Modules
    $ModulesToImport = "ImportExcel","Microsoft.Graph.Groups", "Microsoft.Graph.Users", "PSFramework","ExchangeOnlineManagement"
    Import-Module $ModulesToImport

    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true

    # Connect to Microsoft Graph
    $Scopes = @("Group.ReadWrite.All","Directory.AccessAsUser.All")
    $Context = Get-MgContext

    if ([string]::IsNullOrEmpty($Context) -or ($Context.Scopes -notmatch [string]::Join('|', $Scopes))) {
        Connect-MGGraph -Scopes $Scopes
    }

    $excelData = Import-Excel -Path $FilePath -WorksheetName $UserRole

    if ($PSCmdlet.ShouldProcess("Remove user from groups from Excel file")) {
        foreach ($row in $excelData) {
            $GroupName = $row.PrimarySMTP += "@$domain"
            $GroupType = $row.GroupType
            $DisplayName = $row.DisplayName

            if ($PSCmdlet.ShouldProcess("Remove user $UserEmail from $GroupType group $GroupName")) {
                try {
                    Write-PSFMessage -Level Output -Message "Removing $UserEmail from $($GroupType):'$GroupName'" -Target $UserEmail
                    switch ($GroupType) {
                        '365Group' {
                            Remove-UnifiedGroupLinks -Identity $GroupName -LinkType "Members" -Links $UserEmail -Confirm:$false
                        }
                        '365Distribution' {
                            Remove-DistributionGroupMember -Identity $GroupName -Member $UserEmail -Confirm:$false
                        }
                        '365MailEnabledSecurity' {
                            Remove-DistributionGroupMember -Identity $GroupName -Member $UserEmail -Confirm:$false
                        }
                        '365Security' {
                            $user = Get-MgUser -Filter "userPrincipalName eq '$UserEmail'"
                            $ExistingGroup = Get-MgGroup -Filter "DisplayName eq '$($DisplayName)'"
                                if ($ExistingGroup) {
                                Remove-MgGroupMemberByRef -GroupId $ExistingGroup.Id -DirectoryObjectId $User.Id
                                Write-PSFMessage -Level Output -Message "User $UserEmail successfully removed from $GroupType group $GroupName" -Target $UserEmail
                            }
                            else {
                                Write-PSFMessage -Level Warning -Message "No group found with the name: $GroupName" -Target $GroupName
                            }
                        
                        }
                        default {
                            Write-PSFMessage -Level Warning -Message "Unknown group type: $GroupType" -Target $GroupType
                        }
                        
                    }
                    Write-PSFMessage -Level Output -Message "Removed $UserEmail from $($GroupType):'$GroupName' sucessfully" -Target $UserEmail
                } catch {
                    Write-PSFMessage -Level Error -Message "Error removing user $UserEmail from $($GroupType):'$GroupName'" -Target $UserEmail
                }
            }
        }
    }

# Disconnect Exchange Online and Microsoft Graph sessions
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
}
