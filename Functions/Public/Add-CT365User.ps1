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

        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Security.SecureString]$Password = $(Read-Host -Prompt "Enter the password" -AsSecureString)

    )

    # Import Required Modules
    $ModulesToImport = "ImportExcel","Microsoft.Graph.Users","Microsoft.Graph.Groups","Microsoft.Graph.Identity.DirectoryManagement","Microsoft.Graph.Users.Actions","PSFramework"
    Import-Module $ModulesToImport

    # Connect to Microsoft Graph - Pull these out eventually still in here for testing
    Connect-MgGraph -Scopes "Directory.ReadWrite.All"

    # Import user data from Excel file
    $userData = $null
    try {
        $userData = Import-Excel -Path $FilePath -WorksheetName Users
    } catch {
        Write-PSFMessage -Level Error -Message "Failed to import user data from Excel file."
        return
    }

    # Iterate through each user in the Excel file and create them
    foreach ($user in $userData) {
        $NewUserParams = @{
            UserPrincipalName = "$($user.UserName)@$domain"
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
            BusinessPhones    = $user.PhoneNumber
            MobilePhone       = $user.MobilePhone
            UsageLocation     = $user.UsageLocation
            AccountEnabled    = $true
        }

        $PasswordProfile   = @{
            'ForceChangePasswordNextSignIn' = $false
            'Password'                      = $password | ConvertFrom-SecureString -AsPlainText
        }
        
        Write-PSFMessage -Level Output -Message "Creating user: '$($NewUserParams.UserPrincipalName)'" -Target $user.UserName

        $createdUser = New-MgUser @NewUserParams -PasswordProfile $PasswordProfile

        # Validate user creation
        if ($null -ne $createdUser) {
            Write-PSFMessage -Level Output -Message "User: '$($NewUserParams.UserPrincipalName)' created successfully" -Target $user.UserName
        } else {
            Write-PSFMessage -Level Warning -Message "Failed to create user: '$($NewUserParams.UserPrincipalName)'" -Target $user.UserName
            # if the creation failed go ahead with the next user and skip the license part
            continue
        }

        $licenses = Get-MgSubscribedSku | Where-Object {$_.SkuPartNumber -eq $user.License }
        $user = Get-MgUser -Filter "DisplayName eq '$($NewUserParams.DisplayName)'"
        
        Write-PSFMessage -Level Host -Message "Assigning license $($user.License) to user: '$($NewUserParams.UserPrincipalName)'" -Target $user.UserName

        Set-MgUserLicense -UserId $user.Id -AddLicenses @{SkuId = ($licenses.SkuId)} -RemoveLicenses @()

        # Retrieve the user's licenses after assignment
        $assignedLicenses = Get-MgUserLicenseDetail -UserId $user.Id | Select-Object -ExpandProperty SkuId

        # Check if the assigned license ID is in the user's licenses
        if ($assignedLicenses -contains $licenses.SkuId) {
            Write-PSFMessage -Level Output -Message "License $License successfully assigned to user: '$($NewUserParams.UserPrincipalName)'" -Target $user.UserName
        } else {
            Write-PSFMessage -Level Warning -Message "Failed to assign license $License to user: '$($NewUserParams.UserPrincipalName)'" -Target $user.UserName
        }
    }

    # Disconnect Exchange Online and Microsoft Graph sessions
    Disconnect-MgGraph
}