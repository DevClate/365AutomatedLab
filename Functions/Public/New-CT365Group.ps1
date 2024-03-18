<#
.SYNOPSIS
This function creates Office 365 Groups, Distribution Groups, Mail-Enabled Security Groups, and Security Groups based on the data provided in an Excel file.
.DESCRIPTION
The function New-CT365Group takes the path of an Excel file, User Principal Name and a Domain as input. It creates Office 365 Groups, Distribution Groups, Mail-Enabled Security Groups, and Security Groups based on the data found in the Excel file. If a group already exists, the function will output a message and skip the creation of that group.
.PARAMETER FilePath
The full path to the Excel file(.xlsx) containing the data for the groups to be created. The Excel file should have a worksheet named "Groups". Each row in this worksheet should represent a group to be created. The columns in the worksheet should include the DisplayName, Type (365Group, 365Distribution, 365MailEnabledSecurity, 365Security), PrimarySMTP (without domain), and Description of the group.
.PARAMETER UserPrincipalName
The User Principal Name (UPN) used to connect to Exchange Online.
.PARAMETER Domain
The domain to be appended to the PrimarySMTP of each group to form the email address of the group.
.EXAMPLE
New-CT365Group -FilePath "C:\Path\to\file.xlsx" -UserPrincialName "admin@domain.com" -Domain "domain.com"
This will read the Excel file "file.xlsx" located at "C:\Path\to\", use "admin@domain.com" to connect to Exchange Online, and append "@domain.com" to the PrimarySMTP of each group to form the email address of the group.
.INPUTS
System.String
.OUTPUTS
System.String
The function outputs strings informing about the creation of the groups or if the groups already exist.
.NOTES
The function uses the ExchangeOnlineManagement and Microsoft.Graph.Groups modules to interact with Office 365. Make sure these modules are installed before running the function.
.LINK
Get-UnifiedGroup
.LINK
New-UnifiedGroup
.LINK
Get-DistributionGroup
.LINK
New-DistributionGroup
.LINK
Get-MgGroup
.LINK
New-MgGroup
#>
function New-CT365Group {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
                # First, check if the file has a valid Excel extension (.xlsx)
                if (-not(([System.IO.Path]::GetExtension($psitem)) -match "\.(xlsx)$")) {
                    throw "The file path '$PSitem' does not have a valid Excel format. Please make sure to specify a valid file with a .xlsx extension and try again."
                }
        
                # Then, check if the file exists
                if (-not([System.IO.File]::Exists($psitem))) {
                    throw "The file path '$PSitem' does not lead to an existing file. Please verify the 'FilePath' parameter and ensure that it points to a valid file (folders are not allowed)."
                }
        
                # Return true if both conditions are met
                $true
            })]
        [string]$FilePath,

        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$UserPrincipalName,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateScript({
                # Check if the domain fits the pattern
                switch ($psitem) {
                    { $psitem -notmatch '^(((?!-))(xn--|_)?[a-z0-9-]{0,61}[a-z0-9]{1,1}\.)*(xn--)?[a-z]{2,}(?:\.[a-z]{2,})+$' } {
                        throw "The provided domain is not in the correct format."
                    }
                    Default {
                        $true
                    }
                }
            })]
        [string]$Domain
    )
    

    # Import the required modules
    $ModulesToImport = "ImportExcel", "Microsoft.Graph.Groups", "PSFramework", "ExchangeOnlineManagement"
    Import-Module $ModulesToImport
    
    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
    
    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "Group.ReadWrite.All", "User.Read.All"
    
    # Import data from Excel
    $groups = Import-Excel -Path $FilePath -WorksheetName Groups
    
    foreach ($group in $groups) {
        # Append the domain to the PrimarySMTP
        $group.PrimarySMTP += "@$Domain"
        if (-not [string]::IsNullOrEmpty($group.ManagedBy)) {
            $group.ManagedBy += "@$Domain"
        }
        switch -Regex ($group.Type) {
            "^365Group$" {
                try {
                    Write-PSFMessage -Level Output -Message "Creating 365 Group: $($group.DisplayName)" -Target $Group.DisplayName
                    Get-UnifiedGroup -Identity $group.DisplayName -ErrorAction Stop
                    Write-PSFMessage -Level Warning -Message "365 Group: $($group.DisplayName) already exists" -Target $Group.DisplayName
                }
                catch {
                    $ManagedBy = $group.ManagedBy
                    if ([string]::IsNullOrEmpty($ManagedBy)) {
                        $ManagedBy = $UserPrincipalName
                    }
                    New-UnifiedGroup -DisplayName $group.DisplayName -PrimarySMTPAddress $group.PrimarySMTP -AccessType Private -Notes $group.Description -RequireSenderAuthenticationEnabled $False -Owner $ManagedBy
                    Write-PSFMessage -Level Output -Message "Created 365 Group: $($Group.DisplayName) successfully" -Target $Group.DisplayName
                }
            }
            "^365Distribution$" {
                try {
                    Write-PSFMessage -Level Output -Message "Creating 365 Distribution Group: $($group.DisplayName)" -Target $Group.DisplayName
                    Get-DistributionGroup -Identity $group.DisplayName -ErrorAction Stop
                    Write-PSFMessage -Level Warning -Message "365 Distribution Group $($group.DisplayName) already exists" -Target $Group.DisplayName
                }
                catch {
                    $ManagedBy = $group.ManagedBy
                    if ([string]::IsNullOrEmpty($ManagedBy)) {
                        $ManagedBy = $UserPrincipalName
                    }
                    New-DistributionGroup -Name $group.DisplayName -DisplayName $($group.DisplayName) -PrimarySMTPAddress $group.PrimarySMTP -Description $group.Description -ManagedBy $ManagedBy -RequireSenderAuthenticationEnabled $False
                    Write-PSFMessage -Level Output -Message "Created 365 Distribution Group: $($group.DisplayName)"  -Target $Group.DisplayName
                }
            }
            "^365MailEnabledSecurity$" {
                try {
                    Write-PSFMessage -Level Output -Message "Creating 365 Mail-Enabled Security Group: $($group.DisplayName)" -Target $Group.DisplayName
                    Get-DistributionGroup -Identity $group.DisplayName -ErrorAction Stop
                    Write-PSFMessage -Level Warning -Message "365 Mail-Enabled Security Group: $($group.DisplayName) already exists" -Target $Group.DisplayName
                }
                catch {
                    $ManagedBy = $group.ManagedBy
                    if ([string]::IsNullOrEmpty($ManagedBy)) {
                        $ManagedBy = $UserPrincipalName
                    }
                    New-DistributionGroup -Name $group.DisplayName -PrimarySMTPAddress $group.PrimarySMTP -Type "Security" -Description $group.Description -ManagedBy $ManagedBy -RequireSenderAuthenticationEnabled $False
                    Write-PSFMessage -Level Output -Message "Created 365 Mail-Enabled Security Group: $($group.DisplayName)" -Target $Group.DisplayName
                }
            }
            "^365Security$" {
                Write-PSFMessage -Level Output -Message "Creating 365 Security Group: $($group.DisplayName)" -Target $Group.DisplayName
                $ExistingGroup = Get-MgGroup -Filter "DisplayName eq '$($group.DisplayName)'"
                if ($ExistingGroup) {
                    Write-PSFMessage -Level Warning -Message "365 Security Group: $($group.DisplayName) already exists" -Target $Group.DisplayName
                    continue
                }
                $ManagedBy = $group.ManagedBy
                if ([string]::IsNullOrEmpty($ManagedBy)) {
                    $ManagedBy = $UserPrincipalName
                }
                $GroupOwner = Get-MgUser -Filter "UserPrincipalName eq '$ManagedBy'"
                if ($null -eq $GroupOwner) {
                    Write-PSFMessage -Level Error -Message "User with UserPrincipalName '$ManagedBy' not found"
                    continue
                }

                $mailNickname = $group.PrimarySMTP.Split('@')[0]
                New-MgGroup -DisplayName $group.DisplayName -Description $group.Description -MailNickName $mailNickname -SecurityEnabled:$true -MailEnabled:$false
                $365group = Get-MgGroup -Filter "DisplayName eq '$($group.DisplayName)'"
                New-MgGroupOwner -GroupId $365group.Id -DirectoryObjectId $GroupOwner.Id
                Write-PSFMessage -Level Output -Message "Created 365 Security Group: $($group.DisplayName)" -Target $Group.DisplayName
            }
            default {
                Write-PSFMessage -Level Error -Message "Invalid group type for $($group.DisplayName)" -Target $Group.DisplayName
            }
        }
    }
    # Disconnect Exchange Online and Microsoft Graph sessions
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
}