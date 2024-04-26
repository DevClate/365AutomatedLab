<#
.SYNOPSIS
Creates a new user in Office 365.

.DESCRIPTION
The New-CT365User function creates a new user in Office 365 using the Microsoft Graph API. 
It imports user data from an Excel file and assigns licenses based on the user data. 
If the UseDeveloperPackE5 switch is set, it assigns the DEVELOPERPACK_E5 license to all users.

.PARAMETER FilePath
The path to the Excel file that contains the user data. This parameter is mandatory.

.PARAMETER Domain
The domain for the new users. This parameter is mandatory.

.PARAMETER UseDeveloperPackE5
A switch that, if set, assigns the DEVELOPERPACK_E5 license to all users.

.PARAMETER Password
The password for the new users. If not provided, the function will prompt for the password.

.EXAMPLE
New-CT365User -FilePath "C:\Users\admin\Documents\user_data.xlsx" -Domain "contoso.com" -UseDeveloperPackE5

This command creates new users in Office 365 using the user data in the "user_data.xlsx" file, assigns the DEVELOPERPACK_E5 license to all users, and prompts for the password.

.EXAMPLE
New-CT365User -FilePath "C:\Users\admin\Documents\user_data.xlsx" -Domain "contoso.com"

This command creates new users in Office 365 using the user data in the "user_data.xlsx" file, assigns named license from excel worksheet, and prompts for the password.

.NOTES
You need to have the necessary permissions to create users and assign licenses in Office 365.
#>
function New-CT365User {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [string]$Domain,

        [Parameter()]
        [switch]$UseDeveloperPackE5,

        [Parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [Security.SecureString]$Password = $(Read-Host -Prompt "Enter the password" -AsSecureString)
    )

    # Import Required Modules
    $ModulesToImport = "ImportExcel", "Microsoft.Graph.Users", "Microsoft.Graph.Groups", "Microsoft.Graph.Identity.DirectoryManagement", "Microsoft.Graph.Users.Actions", "PSFramework"
    Import-Module $ModulesToImport

    # Scopes
    $Scopes = @("Directory.ReadWrite.All")
    $Context = Get-MgContext

    if ([string]::IsNullOrEmpty($Context) -or ($Context.Scopes -notmatch [string]::Join('|', $Scopes))) {
        Connect-MGGraph -Scopes $Scopes
    }

    # Import user data from Excel file
    $userData = $null
    try {
        $userData = Import-Excel -Path $FilePath -WorksheetName Users
    }
    catch {
        Write-PSFMessage -Level Error -Message "Failed to import user data from Excel file."
        return
    }

    foreach ($user in $userData) {
        # Prepare user parameters for creation
        $userParams = @{
            UserPrincipalName = "$($user.UserName)@$Domain"
            GivenName         = $user.FirstName
            Surname           = $user.LastName
            DisplayName       = "$($user.FirstName) $($user.LastName)"
            MailNickname      = $user.UserName
            JobTitle          = $user.Title
            Department        = $user.Department
            StreetAddress     = $user.StreetAddress
            City              = $user.City
            State             = $user.State
            PostalCode        = $user.PostalCode
            Country           = $user.Country
            UsageLocation     = $user.UsageLocation
            CompanyName       = $user.CompanyName
            AccountEnabled    = $true
            PasswordProfile   = @{
                ForceChangePasswordNextSignIn = $false
                Password                      = $Password | ConvertFrom-SecureString -AsPlainText
            }
        }

        # Add optional properties if they exist
        foreach ($prop in @('MobilePhone', 'FaxNumber', 'EmployeeHireDate', 'EmployeeId', 'EmployeeType')) {
            if (-not [string]::IsNullOrEmpty($user.$prop)) {
                $userParams[$prop] = $user.$prop
            }
        }

        if (-not [string]::IsNullOrEmpty($user.PhoneNumber)) {
            $UserParams.BusinessPhones = @($user.PhoneNumber)
        }

        # Create the new user
        $createdUser = New-MgUser @userParams
        if ($null -ne $createdUser) {
            Write-PSFMessage -Level Host -Message "User created: $($userParams.UserPrincipalName)" -Target $user.UserName
        } else {
            Write-PSFMessage -Level Warning -Message "Failed to create user: $($userParams.UserPrincipalName)" -Target $user.UserName
            continue
        }

        # License assignment logic
        $licenseType = $UseDeveloperPackE5 ? 'DEVELOPERPACK_E5' : $user.License
        $licenses = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $licenseType }
        
        if ($licenses) {
            Set-MgUserLicense -UserId $createdUser.Id -AddLicenses @{ SkuId = $licenses.SkuId } -RemoveLicenses @()
            Write-PSFMessage -Level Host -Message "License assigned: $($licenses.SkuPartNumber) to user: $($userParams.UserPrincipalName)" -Target $user.UserName
        } else {
            Write-PSFMessage -Level Warning -Message "Failed to assign license: $licenseType to user: $($userParams.UserPrincipalName)" -Target $user.UserName
        }

        # Manager assignment if applicable
        if ($null -ne $user.ManagerUPN) {
            $managerUPNData = "$($user.ManagerUPN)@$Domain"
            $manager = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($managerUPNData)" }
        
            $assigningManagerMessage = "Assigning manager $($managerUPNData) to user: '$($UserParams.UserPrincipalName)'"
            Write-PSFMessage -Level Host -Message $assigningManagerMessage -Target $user.UserName
        
            Set-MgUserManagerByRef -UserId $createdUser.Id -BodyParameter $manager
        
            # Confirm manager assignment
            $managerDirectoryObject = Get-MgUserManager -UserId $createdUser.Id
            if ($null -ne $managerDirectoryObject) {
                $managerUser = Get-MgUser -UserId $managerDirectoryObject.Id
        
                $assignedManagerMessage = "Assigned manager $($managerUser.UserPrincipalName) to user: '$($NewUserParams.UserPrincipalName)'"
                Write-PSFMessage -Level Host -Message $assignedManagerMessage -Target $user.UserName
            }
            else {
                $failedToAssignManagerMessage = "Failed to assign manager $($user.Manager) to user: '$($NewUserParams.UserPrincipalName)'"
                Write-PSFMessage -Level Warning -Message $failedToAssignManagerMessage -Target $user.UserName
            }
        }
    }

    # Close the Microsoft Graph connection
    Disconnect-MgGraph
}
