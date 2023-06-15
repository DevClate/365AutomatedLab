<#
.SYNOPSIS
    This function creates new Microsoft 365 users from data in an Excel file and assigns them a license.

.DESCRIPTION
    The Add-CT365User function imports user data from an Excel file, creates new users in Microsoft 365, and assigns them a license. 
    It performs these tasks using the Microsoft.Graph.Users and Microsoft.Graph.Groups modules.

.PARAMETER FilePath
    The path of the Excel file containing user data. The file should have a worksheet named 'Users' with columns for UserName, FirstName, LastName, Title, Department, StreetAddress, City, State, PostalCode, Country, PhoneNumber, MobilePhone, UsageLocation, and License. 
    This parameter is mandatory and accepts pipeline input and property names.

.PARAMETER domain
    The domain to be appended to the UserName to create the UserPrincipalName for each user.
    This parameter is mandatory and accepts pipeline input and property names.

.EXAMPLE
    Add-CT365User -FilePath "C:\Path\to\file.xlsx" -domain "contoso.com"
    This command imports user data from the 'file.xlsx' file and creates new users in Microsoft 365 under the domain 'contoso.com'.

.NOTES
    The function connects to Microsoft Graph using 'Directory.ReadWrite.All' scope. Make sure the account running this script has the necessary permissions.
    The function sets the password for each new user to 'P@ssw0rd123' and does not require the user to change the password at the next sign-in. 
    Modify the password setting to meet your organization's security requirements.

    Connect-MgGraph -Scopes "Directory.ReadWrite.All" - is needed to connect to Graph
#>
function Add-CT365User {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$FilePath,
        
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$domain
    )

    # Import Required Modules
    Import-Module ImportExcel
    Import-Module Microsoft.Graph.Users
    Import-Module Microsoft.Graph.Groups
    Import-Module Microsoft.Graph.Identity.DirectoryManagement
    Import-Module Microsoft.Graph.Users.Actions
    Import-Module PSFramework

    # Check if Excel file exists
    if (!(Test-Path $FilePath)) {
        Write-Warning "File $FilePath does not exist. Please check the file path and try again."
        return
    }
    


    # Connect to Microsoft Graph - Pull these out eventually still in here for testing
    Connect-MgGraph -Scopes "Directory.ReadWrite.All"


    # Import user data from Excel file
    $userData = Import-Excel -Path $FilePath -WorksheetName Users

    foreach ($user in $userData) {

            $UserPrincipalName = $user.UserName
            $GivenName         = $user.FirstName
            $Surname           = $user.LastName
            $MailNickname      = $user.UserName
            $JobTitle          = $user.Title
            $Department        = $user.Department
            $Streetaddress     = $user.StreetAddress
            $City              = $user.City
            $State             = $user.State
            $PostalCode        = $user.PostalCode
            $Country           = $user.Country
            $BusinessPhones    = $user.PhoneNumber
            $MobilePhone       = $user.MobilePhone
            $UsageLocation     = $user.UsageLocation
            $License           = $user.License
            
        $NewUserParams = @{
            UserPrincipalName = "$userPrincipalName@$domain"
            GivenName         = $GivenName
            Surname           = $Surname
            DisplayName       = "$GivenName $Surname"
            MailNickname      = $MailNickname
            JobTitle          = $JobTitle
            Department        = $Department
            StreetAddress     = $Streetaddress
            City              = $City
            State             = $State
            PostalCode        = $PostalCode
            Country           = $Country
            BusinessPhones    = $BusinessPhones
            MobilePhone       = $MobilePhone
            UsageLocation     = $UsageLocation
            AccountEnabled    = $true
        }

        $PasswordProfile   = @{
            'ForceChangePasswordNextSignIn' = $false
            'Password'                      = 'P@ssw0rd123'
        }
        
        Write-Output "Creating user $userPrincipalName@$domain"

        $createdUser = New-MgUser @NewUserParams -PasswordProfile $PasswordProfile

        # Validate user creation
        if ($null -ne $createdUser) {
            Write-Output "User $userPrincipalName@$domain created successfully"
        } else {
            Write-Warning "Failed to create user $userPrincipalName@$domain"
            }

        $licenses = Get-MgSubscribedSku | Where-Object {$_.SkuPartNumber -eq $License }
        $user = Get-MgUser | Where-Object {$_.DisplayName -eq $NewUserParams.DisplayName}
        
        Write-Output "Assigning license $License to user $userPrincipalName@$domain"

        Set-MgUserLicense -UserId $user.Id -AddLicenses @{SkuId = ($licenses.SkuId)} -RemoveLicenses @()

        # Retrieve the user's licenses after assignment
        $assignedLicenses = Get-MgUserLicenseDetail -UserId $user.Id | Select-Object -ExpandProperty SkuId

        # Check if the assigned license ID is in the user's licenses
        if ($assignedLicenses -contains $licenses.SkuId) {
            Write-Output "License $License successfully assigned to user $userPrincipalName@$domain" 
        } else {
            Write-Warning "Failed to assign license $License to user $userPrincipalName@$domain"
        }
    }
}

