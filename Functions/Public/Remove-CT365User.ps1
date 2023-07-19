<#
.SYNOPSIS
Removes a user from Microsoft 365 based on the provided Excel data.

.DESCRIPTION
The Remove-CT365User function connects to the Microsoft Graph, reads user data from the provided Excel file, 
and attempts to remove each user listed in the file from Microsoft 365.

.PARAMETER FilePath
Specifies the full path to the Excel file that contains the user data. This parameter is mandatory.

.PARAMETER Domain
Specifies the domain that will be concatenated with the UserPrincipalName to form a valid email address. This parameter is mandatory.

.EXAMPLE
Remove-CT365User -FilePath "C:\Path\to\file.xlsx" -Domain "example.com"

This command attempts to remove the users listed in the "file.xlsx" Excel file from the "example.com" domain.

.INPUTS
System.String. You can pipe a string that contains the file path and domain to Remove-CT365User.

.OUTPUTS
System.String. Outputs a message for each attempted user removal, indicating success or failure.

.NOTES
This function requires the Microsoft.Graph.Users, ImportExcel, and PSFramework modules. Make sure to install them using Install-Module before running this function.

.LINK
https://docs.microsoft.com/en-us/powershell/module/microsoft.graph.users/?view=graph-powershell-1.0

.LINK
https://www.powershellgallery.com/packages/ImportExcel

.LINK
https://psframework.org/documentation/commands/PSFramework.html
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
    Import-Module PSFramework

    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "User.ReadWrite.All"

    # Import user data from Excel file
    $userData = $null
    try {
        $userData = Import-Excel -Path $FilePath -WorksheetName Users
    } catch {
        Write-PSFMessage -Level Error -Message "Failed to import user data from Excel file."
        return
    }

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

            Write-PSFMessage -Level Output -Message "Attemping to remove User $($NewUserParams.DisplayName)" -Target $NewUserParams.DisplayName

            # Validate if the user exists
            if ($userToRemove) {
                Remove-MgUser -UserId $userToRemove.id
                
                # Check the user's existence
                $removedUser = Get-MgUser | Where-Object {$_.DisplayName -eq $NewUserParams.DisplayName}
                
                # Confirm that the user was removed
                if (-not $removedUser) {
                    Write-PSFMessage -Level Output -Message "User $($NewUserParams.DisplayName) has been successfully removed." -Target $NewUserParams.DisplayName
                } else {
                    Write-PSFMessage -Level Warning -Message "Failed to remove user $($NewUserParams.DisplayName)." -Target $NewUserParams.DisplayName
                }
            } else {
                Write-PSFMessage -Level Warning -Message "User $($NewUserParams.DisplayName) does not exist." -Target $NewUserParams.DisplayName
            }
        }
    }
Disconnect-MgGraph
