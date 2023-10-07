<#
.SYNOPSIS
This function adds a user to Microsoft 365 groups based on a provided Excel file.

.DESCRIPTION
The New-CT365GroupByTitle function uses Microsoft Graph and Exchange Online Management modules to add a user to different types of Microsoft 365 groups. The group details are read from an Excel file. The group's SMTP, type, and display name are extracted from the Excel file and used to add the user to the group.

.PARAMETER FilePath
The path to the Excel file that contains the groups to which the user should be added. The file must contain a worksheet named as per the user role ("NY-IT" or "NY-HR"). The worksheet should contain details about the groups including the primary SMTP, group type, and display name.

.PARAMETER UserEmail
The email of the user to be added to the groups specified in the Excel file.

.PARAMETER Domain
The domain that is appended to the primary SMTP to form the group's email address.

.PARAMETER UserRole
The role of the user, which is also the name of the worksheet in the Excel file that contains the groups to which the user should be added. The possible values are "NY-IT" and "NY-HR".

.EXAMPLE
New-CT365GroupByTitle -FilePath "C:\Users\Username\Documents\Groups.xlsx" -UserEmail "jdoe@example.com" -Domain "example.com" -UserRole "NY-IT"

This command reads the groups from the "NY-IT" worksheet in the Groups.xlsx file and adds the user "jdoe@example.com" to those groups.

.NOTES
This function requires the ExchangeOnlineManagement, ImportExcel, and Microsoft.Graph.Groups modules to be installed. It will import these modules at the start of the function. The function connects to Exchange Online and Microsoft Graph, and it will disconnect from them at the end of the function.

#>
function New-CT365GroupByTitle {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
            #making sure the Filepath leads to a file and not a folder and has a proper extension
            switch ($psitem){
                {-not([System.IO.File]::Exists($psitem))}{
                    throw "The file path '$PSitem' does not lead to an existing file. Please verify the 'FilePath' parameter and ensure that it points to a valid file (folders are not allowed)."
                }
                {-not(([System.IO.Path]::GetExtension($psitem)) -match "(.xlsx)")}{
                    "The file path '$PSitem' does not have a valid Excel format. Please make sure to specify a valid file with a .xlsx extension and try again."
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
    $ModulesToImport = "ImportExcel","Microsoft.Graph.Groups","PSFramework","ExchangeOnlineManagement"
    Import-Module $ModulesToImport

    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
    
    # Connect to Microsoft Graph
    $Scopes = @("Group.ReadWrite.All")
    $Context = Get-MgContext

    if ([string]::IsNullOrEmpty($Context) -or ($Context.Scopes -notmatch [string]::Join('|', $Scopes))) {
        Connect-MGGraph -Scopes $Scopes
    }

    $excelData = Import-Excel -Path $FilePath -WorksheetName $UserRole

    if ($PSCmdlet.ShouldProcess("Add user to groups from Excel file")) {
        foreach ($row in $excelData) {
            $GroupName = $row.PrimarySMTP += "@$domain"
            $GroupType = $row.Type
            $DisplayName = $row.DisplayName

            if ($PSCmdlet.ShouldProcess("Add user $UserEmail to $GroupType group $GroupName")) {
                try {
                    Write-PSFMessage -Level Output -Message "Adding $UserEmail to $($GroupType):'$GroupName'" -Target $UserEmail
                    switch ($GroupType) {
                        '365Group' {
                            Add-UnifiedGroupLinks -Identity $GroupName -LinkType "Members"-Links $UserEmail -erroraction Stop
                        }
                        '365Distribution' {
                            Add-DistributionGroupMember -Identity $GroupName -Member $UserEmail -Erroraction Stop
                        }
                        '365MailEnabledSecurity' {
                            Add-DistributionGroupMember -Identity $GroupName -Member $UserEmail -Erroraction Stop
                        }
                        '365Security' {
                            $user = Get-MgUser -Filter "userPrincipalName eq '$UserEmail'"
                            $ExistingGroup = Get-MgGroup -Filter "DisplayName eq '$($DisplayName)'"
                            if ($ExistingGroup) {
                                New-MgGroupMember -GroupId $ExistingGroup.Id -DirectoryObjectId $User.Id -ErrorAction Stop
                                Write-PSFMessage -Level Output -Message "User $UserEmail successfully added to $GroupType group $GroupName" -Target $UserEmail
                            }
                            else {
                                Write-PSFMessage -Level Warning -Message "No group found with the name: $GroupName" -Target $GroupName
                            }
                        
                        }
                        default {
                            Write-PSFMessage -Level Warning -Message "Unknown group type: $GroupType" -Target $GroupType
                        }  
                    }
                    Write-PSFMessage -Level Output -Message "Added $UserEmail to $($GroupType):'$GroupName' sucessfully" -Target $UserEmail
                } catch {
                    Write-PSFMessage -Level Error -Message "Error adding $UserEmail to $($GroupType):'$GroupName'" -Target $UserEmail
                }
            }
        }
    }
    # Disconnect Exchange Online and Microsoft Graph sessions
    Disconnect-ExchangeOnline -Confirm:$false
    if (-not [string]::IsNullOrEmpty($(Get-MgContext))) {
        Disconnect-MgGraph
    }
}
