<#
.SYNOPSIS
    Removes Office 365 groups from an Excel file.

.DESCRIPTION
    This function reads data from an Excel file and creates three different types of Office 365 groups: Unified Groups, Distribution Groups, Mail-Enabled Security Groups, and Security Groups. It uses the Exchange Online Management module (v2 or higher) and the Microsoft.Graph module.

.PARAMETER FilePath
    The full path to the Excel file containing the group information. The Excel file should have columns 'DisplayName', 'PrimarySMTP', 'Description', 'Type'

.EXAMPLE
    Remove-CT365Group -FilePath "C:\path\to\your\file.xlsx"

    Reads data from the specified Excel file and creates the corresponding Office 365 groups.

.NOTES
    This script assumes that you have the necessary permissions to create Office 365 groups in your organization.
    Before running the script, make sure you've installed the "ExchangeOnlineManagement", "Microsoft.Graph", and "ImportExcel" modules.
#>
function Remove-CT365Group {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$FilePath,
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$UserPrincialName
    )

    # Check if Excel file exists
    if (!(Test-Path $FilePath)) {
        Write-Warning "File $FilePath does not exist. Please check the file path and try again."
        return
    }

    # Import the required modules
    Import-Module ExchangeOnlineManagement
    Import-Module Microsoft.Graph
    Import-Module ImportExcel

    # Connect to Exchange Online
    $exoSession = Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true
    
    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "Group.ReadWrite.All"

    # Import data from Excel
    $groups = Import-Excel -Path $FilePath -WorksheetName Groups

    foreach ($group in $groups) {
        switch ($group.Type) {
            "365Group" {
                try {
                    Write-Output "Removing 365 Group $group.DisplayName"
                    Get-UnifiedGroup -Identity $group.DisplayName -ErrorAction Stop
                    Remove-UnifiedGroup -Identity $group.DisplayName -Confirm:$false
                    Write-Host "Removed 365 Group $($group.DisplayName)" -ForegroundColor Green
                } catch {
                    Write-Host "365 Group $($group.DisplayName) does not exist" -ForegroundColor Yellow
                }
            }
            "365Distribution" {
                try {
                    Write-Output "Removing 365 Distribution Group $group.DisplayName"
                    Get-DistributionGroup -Identity $group.DisplayName -ErrorAction Stop
                    Remove-DistributionGroup -Identity $group.DisplayName -Confirm:$false
                    Write-Host "Removed Distribution Group $($group.DisplayName)" -ForegroundColor Green
                } catch {
                    Write-Host "Distribution Group $($group.DisplayName) does not exist" -ForegroundColor Yellow
                }
            }
            "365MailEnabledSecurity" {
                try {
                    Write-Output "Removing 365 Mail-Enabled Security Group $group.DisplayName"
                    Get-DistributionGroup -Identity $group.DisplayName -ErrorAction Stop
                    Remove-DistributionGroup -Identity $group.DisplayName -Confirm:$false
                    Write-Host "Removed Mail-Enabled Security Group $($group.DisplayName)" -ForegroundColor Green
                } catch {
                    Write-Host "Mail-Enabled Security Group $($group.DisplayName) does not exist" -ForegroundColor Yellow
                }
            }
            "365Security" {
                Write-Output "Removing 365 Security Group $group.DisplayName"
                $existingGroup = Get-MgGroup -Filter "DisplayName eq '$($group.DisplayName)'"
                if ($existingGroup) {
                    Remove-MgGroup -GroupId $existingGroup.Id -Confirm:$false
                    Write-Host "Removed Security Group $($group.DisplayName)" -ForegroundColor Green
                } else {
                    Write-Host "Security Group $($group.DisplayName) does not exist" -ForegroundColor Yellow
                }
            }
            default {
                Write-Host "Invalid group type for $($group.DisplayName)" -ForegroundColor Yellow
            }
        }
    }
    

    # Disconnect Exchange Online and Microsoft Graph sessions
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
}
