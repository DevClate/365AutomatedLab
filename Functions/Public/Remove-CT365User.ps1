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

    # Import Required Modules
    $ModulesToImport = "ImportExcel", "Microsoft.Graph.Users", "Microsoft.Graph.Groups", "Microsoft.Graph.Identity.DirectoryManagement", "Microsoft.Graph.Users.Actions", "PSFramework"
    Import-Module $ModulesToImport

    # Connect to Microsoft Graph
    $Scopes = @("User.ReadWrite.All")
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

    # Iterate through each user in the Excel file and delete them
    foreach ($user in $userData) {
        $NewUserParams = @{
            UserPrincipalName = "$($user.UserName)@$domain"
            GivenName         = $user.FirstName
            Surname           = $user.LastName
            DisplayName       = "$($user.Firstname) $($user.Lastname)"
            MailNickname      = $user.UserName
            JobTitle          = $user.Title
            Department        = $user.Department
        }
            
        Write-PSFMessage -Level Output -Message "Removing user: '$($NewUserParams.UserPrincipalName)'" -Target $NewUserParams.UserName
            
        $userToRemove = Get-MgUser | Where-Object { $_.DisplayName -eq $NewUserParams.DisplayName }

        # Validate if the user exists
        if ($userToRemove) {
            Remove-MgUser -UserId $userToRemove.id
                
            # Check the user's existence
            $removedUser = Get-MgUser | Where-Object { $_.DisplayName -eq $NewUserParams.DisplayName }
                
            # Confirm that the user was removed
            if (-not $removedUser) {
                Write-PSFMessage -Level Output -Message "User $($NewUserParams.DisplayName) has been successfully removed." -Target $NewUserParams.DisplayName
            }
            else {
                Write-PSFMessage -Level Warning -Message "Failed to remove user $($NewUserParams.DisplayName)." -Target $NewUserParams.DisplayName
            }
        }
        else {
            Write-PSFMessage -Level Warning -Message "User $($NewUserParams.DisplayName) does not exist." -Target $NewUserParams.DisplayName
        }
    }
    
    # Disconnect Microsoft Graph Sessions
    if (-not [string]::IsNullOrEmpty($(Get-MgContext))) {
        Disconnect-MgGraph
    }
}
