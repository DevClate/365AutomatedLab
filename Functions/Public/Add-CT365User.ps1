<#
.SYNOPSIS
    Creates and manages users in Office 365 from an Excel file.

.DESCRIPTION
    The Add-CT365User function creates and manages users in Office 365 using an Excel file as a data source. 
    The Excel file should have a worksheet named 'Users' containing the required user data. The function also assigns a license to the created users.

.PARAMETER FilePath
    The path to the Excel file which contains the user data.

.PARAMETER domain
    The domain of the Office 365 organization in which the users are to be created. This parameter is mandatory.

.EXAMPLE
    Add-CT365User -FilePath "C:\Data\365DataEnvironment.xlsx" -domain "example.com"
    This example creates users in the 'example.com' domain using the user data from the '365DataEnvironment.xlsx' file.

.INPUTS
    FilePath: A string representing the path to the Excel file.
    domain: A string representing the domain name.

.OUTPUTS
    The function outputs messages indicating the status of user creation and license assignment.

.NOTES
    1. The function uses the ImportExcel module. If the module is not installed, the function will install it.
    2. The function requires the 'Directory.ReadWrite.All' scope in Microsoft Graph for user creation and license assignment.
    3. User data should be in a worksheet named 'Users' in the Excel file.
    4. The password for all new users is set to 'P@ssw0rd123' by default. 
#>
function Add-CT365User {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$FilePath,
        
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$domain
    )

    # Check if Excel file exists
    if (!(Test-Path $FilePath)) {
        Write-Warning "File $FilePath does not exist. Please check the file path and try again."
        return
    }
    
    Import-Module ImportExcel
    Import-Module Microsoft.Graph.Users
    Import-Module Microsoft.Graph.Groups

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

