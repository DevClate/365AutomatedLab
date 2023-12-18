# 365AutomatedLab Changelog

## 2.2.0

**New Features**

New-CT365Teams - added functionality to create channels and their descriptions. Currently you’ll set one owner for all Teams. Please create an issue if you would like to see the option for owners per Teams and Channels.

Verify-CT365VerifyTeamsCreation - internal cmdlet to verify Teams creation

**Breaking Changes**

None

## 2.1.0

**Fixed:** 
Changed function name inside code from Export-CT365GroupToExcel to Export-CT365ProdGroupToExcel.

## 2.0.0

**New Features**

Export-CT365Teams - This will export the teams from your production tenant to an Excel worksheet named Teams.

**Breaking Changes**

For the 3 functions below, there will no longer be the parameter for WorkbookName, it will only be filepath going forward. This is to keep it consistent with the other functions. If you would rather have the WorkbookName, please let me know and if there is enough interest, I'll change that to the standard.

- Export-CT365ProdGroupToExcel
- Export-CT365ProdUserToExcel
- New-CT365DataEnvironment

## 1.1.0

Export-CT365ProdUserToExcel function added to enable you to export your production groups to a template that is easily imported into your dev tenant.

## 1.0.0

Fixed Issues:

    Remove-CT365SharePointSite now behaves correctly. If you only want to delete the sites, run Remove-CT365SharePointSite, and if you want to permanently delete them, you have to run previous command, wait till SharePoint processes(10-20 minutes), then run Remove-CT365SharePointSite -PermanentlyDelete.

## 0.1.8

Added Remove-CT365AllDeletedM365Groups. This will permanently delete all deleted Modern Microsoft 365 Groups.

## 0.1.7

Added Set-CT365SPDistinctNumber. Currently I have it so Sharepoint Sites have a number after them for testing so I know which ones I'm working on and not having to create "real" names for each. This allows you to easily rename the site names in one quick line. I do this as SharePoint Team sites never can fully delete fast as I want while testing.

## 0.1.6

Minor formatting

Confirmed working upload to PowerShell Gallery GitHub Action

## 0.1.5

Added better tags and added tag to show it works on MacOS in PowerShell Gallery

## 0.1.4

Confirmed working on Mac OS

Added microsoft.identity.client v4.50.0.0 into required modules

Added microsoft.identity.client module to import for New-CT365Teams

Added microsoft.identity.client module to import for Remove-CT365Teams

Fixed spelling error for UserPrincipalName on New-CT365Teams

Export-CT365ProdUserToExcel now matches exactly for importing into Dev(only need to add licensing)

## 0.1.3

Fixed issue with New-CT365SharePointSite not creating each of the different sites correctly every time

## 0.1.2

Updated/Created Pester Tests

## 0.1.1

Updated URI for Icon and Documentation
