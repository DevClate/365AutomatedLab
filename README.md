# 365AutomatedLab

This module will create a Microsoft 365 Test Environment using an excel workbook

## Project Summary

I started this module to create a test environment for 365 as there wasn't one for groups, only users which Microsoft provides in there [365 Developer Program Environment](https://developer.microsoft.com/en-us/microsoft-365/dev-program) . I was tired of having to come up with test user names, their information, and groups each time, or remembering which ones I had already created every time I wanted to do some testing. Let's be honest. we aren't in our test environment every day and all day. This started out as a quick function, then another function, then another....you can see how I quickly determined that this needed to be a module to keep everything organized and easily share out to others. With that said, I'd love feedback(Create issues) and community help on this to really expand this.

Please do not use this module in your production environment until tested in your test environment. I highly recommend using Microsoft's 365 Developer Program Environment, and use this module with it so you are truly using this in a test environment. It is free to you, and it renews every 90 days as long as you are using it.

### Requirements

* PowerShell Version:
  * 7.1+ (Untested on Mac OSX and Linux)
* Modules
  * ImportExcel v7.8.2+
  * ExchangeOnlineManagement v2.0.6+
  * Microsoft.Graph.Users v1.17.0+
  * Microsoft.Graph.Groups v1.17.0+
  * Microsoft.Graph.Identity.DirectoryManagement v1.17.0+
  * Microsoft.Graph.Users.Actions v1.17.0+

### Current Functions

* Create Users and assign license
  * Add-CT365User
* Remove Users
  * Remove-CT365User
* Create the 4 types of Office 365 Groups/Distribution Lists
  * Add-CT365Group
* Remove the 4 types of Office 365 Groups/Distribution Lists
  * Remove-CT365Group
* Assign User to any of their 4 Office 365 Groups/Distribution Lists by Job Title and Location
  * Add-CT365GroupByTitle
* Remove User from any of their 4 Office 365 Groups/Distribution Lists by Job Title and Location
  * Remove-CT365GroupByTitle
* Copy worksheets name to a csv file so you can copy those location-titles into your ValidateSet
  * Copy-WorkSheetName

### Data

In LabSources you can find an excel file named 365DataEnvironment.xlsx that has 4 tabs.
* Users: This will have all of the user's information you are creating including licensing information
  * If you do not have a UsageLocation set, the licenses will not be added
* Groups: This will have all the groups you want created
  * I do not have it assigning manager as of yet, but will in the future
* Location-JobTitle: This will have all the groups that location and jobtitle are suppose to have.
