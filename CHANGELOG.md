# 365AutomatedLab Changelog

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
