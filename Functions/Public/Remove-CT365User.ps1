<#
.SYNOPSIS
Removes users from Microsoft 365 using a provided Excel file with user data.

.DESCRIPTION
The Remove-CT365User function uses the Microsoft Graph API to remove users from a Microsoft 365 environment. It takes two parameters: a file path to an Excel file containing user data, and a mandatory domain parameter for user removal.

.PARAMETER FilePath
An mandatory parameter which specifies the location of an Excel file containing user data.

.PARAMETER Domain
A mandatory parameter which specifies the domain for user removal. This should be a valid Microsoft 365 domain.

.EXAMPLE
Remove-CT365User -FilePath "C:\Data\365DataEnvironment.xlsx" -Domain contoso.com

This example demonstrates how to remove users from the "contoso.com" domain using an Excel file located at the specified file path.

.NOTES
This function relies on the Microsoft.Graph and ImportExcel PowerShell modules. If these modules are not installed, the function will attempt to install them. 

It also requires Microsoft 365 Administrator permissions for the 'User.ReadWrite.All' scope to perform user removal operations. 

Make sure the provided Excel file contains the following fields for each user: MailNickname, FirstName, LastName, Title, Department. The Title and Department are for future use.

#>
function Remove-CT365User {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$FilePath,
        
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$domain
    )

    if (!(Test-Path $FilePath)) {
        Write-Warning "File $FilePath does not exist. Please check the file path and try again."
        return
    }

    # Import Required Modules
    Import-Module Microsoft.Graph.Users
    Import-Module ImportExcel

    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "User.ReadWrite.All"

    # Check if the Excel file exists before importing
    if (Test-Path $FilePath) {

        # Import user data from Excel file
        $userData = Import-Excel -Path $FilePath -WorksheetName Users

        foreach ($user in $userData) {

            $UserPrincipalName = $user.MailNickname
            $GivenName         = $user.FirstName
            $Surname           = $user.LastName
            $MailNickname      = $user.MailNickname
            $JobTitle          = $user.Title
            $Department        = $user.Department

            $NewUserParams = @{
                UserPrincipalName = "$userPrincipalName@$domain"
                GivenName         = $GivenName
                Surname           = $Surname
                DisplayName       = "$GivenName $Surname"
                MailNickname      = $MailNickname
                JobTitle          = $JobTitle
                Department        = $Department
            }

            $userToRemove = Get-MgUser | Where-Object {$_.DisplayName -eq $NewUserParams.DisplayName}

            Write-Output "Attemping to remove User $($NewUserParams.DisplayName)"

            # Validate if the user exists
            if ($userToRemove) {
                Remove-MgUser -UserId $userToRemove.id
                
                # Check the user's existence
                $removedUser = Get-MgUser | Where-Object {$_.DisplayName -eq $NewUserParams.DisplayName}
                
                # Confirm that the user was removed
                if (-not $removedUser) {
                    Write-Output "User $($NewUserParams.DisplayName) has been successfully removed."
                } else {
                    Write-Warning "Failed to remove user $($NewUserParams.DisplayName)."
                }
            } else {
                Write-Warning "User $($NewUserParams.DisplayName) does not exist."
            }
        }
    }
}