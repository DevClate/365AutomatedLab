function Add-CT365GroupByTitle {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$ExcelFilePath,
        
        [Parameter(Mandatory)]
        [string]$UserEmail,

        [Parameter(Mandatory)]
        [string]$Domain,
        
        [Parameter(Mandatory)]
        [ValidateSet("NY-IT", "NY-HR")]
        [string]$UserRole
    )

    Import-Module ExchangeOnlineManagement
    Import-Module ImportExcel
    Import-Module Microsoft.Graph.Groups

    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "Group.ReadWrite.All"

    if (!(Test-Path $ExcelFilePath)) {
        Write-Error "Excel file not found at the specified path: $ExcelFilePath"
        return
    }

    $excelData = Import-Excel -Path $ExcelFilePath -WorksheetName $UserRole

    if ($PSCmdlet.ShouldProcess("Add user to groups from Excel file")) {
        foreach ($row in $excelData) {
            $GroupName = $row.PrimarySMTP += "@$domain"
            $GroupType = $row.GroupType
            $DisplayName = $row.DisplayName

            if ($PSCmdlet.ShouldProcess("Add user $UserEmail to $GroupType group $GroupName")) {
                try {
                    switch ($GroupType) {
                        '365Group' {
                            Add-UnifiedGroupLinks -Identity $GroupName -LinkType "Members"-Links $UserEmail
                            Write-Host "User $UserEmail successfully added to $GroupType group $GroupName"
                        }
                        '365Distribution' {
                            Add-DistributionGroupMember -Identity $GroupName -Member $UserEmail
                            Write-Host "User $UserEmail successfully added to $GroupType group $GroupName"
                        }
                        '365MailEnabledSecurity' {
                            Add-DistributionGroupMember -Identity $GroupName -Member $UserEmail
                            Write-Host "User $UserEmail successfully added to $GroupType group $GroupName"
                        }
                        '365Security' {
                            $user = Get-MgUser -Filter "userPrincipalName eq '$UserEmail'"
                            $ExistingGroup = Get-MgGroup -Filter "DisplayName eq '$($DisplayName)'"
                                if ($ExistingGroup) {
                                New-MgGroupMember -GroupId $ExistingGroup.Id -DirectoryObjectId $User.Id
                                Write-Host "User $UserEmail successfully added to $GroupType group $GroupName"
                            }
                            else {
                                Write-Warning "No group found with the name: $GroupName"
                            }
                        
                        }
                        default {
                            Write-Warning "Unknown group type: $GroupType"
                        }
                        
                    }
                } catch {
                    Write-Error "Error adding user $UserEmail to $GroupType group $GroupName $_"
                }
            }
        }
    }

# Disconnect Exchange Online and Microsoft Graph sessions
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph
}
