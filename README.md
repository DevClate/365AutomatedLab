# 365AutomatedLab

This module will create a Microsoft 365 Test Environment using an excel workbook

## Project Summary

I started this module to create a test environment for 365 as there wasn't one for groups, only users which Microsoft provides in there [365 Developer Program Environment](https://developer.microsoft.com/en-us/microsoft-365/dev-program) . I was tired of having to come up with test user names, their information, and groups each time, or remembering which ones I had already created every time I wanted to do some testing. Let's be honest. we aren't in our test environment every day and all day. This started out as a quick function, then another function, then another....you can see how I quickly determined that this needed to be a module to keep everything organized and easily shareable to others. With that said, I'd love feedback(Create issues) and community help on this to really expand this.

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

In LabSources you will find an excel file named 365DataEnvironment.xlsx that has 3 main tabs. Any additional tabs will be for different location-jobtitle tabs. You can use this workbook as is in your test environment, or use it as a template for your own data.

* Users: This will have all of the user's information you are creating including licensing information
  * If you do not have a UsageLocation set, the licenses will not be added
* Groups: This will have all the groups you want created
  * I do not have it assigning manager as of yet, but will in the future
* Location-JobTitle: This will have all the groups that location and job title are suppose to have.
  * Originally I had these in a validateset, but opted out. Let me know in the issues if they should be brought back

In the future, I will have it so you can create random users using Doug Finke's PowerShellAI module and his ImportExcel module. Eventually, it will create the whole workbook! For now you can use ChatGPT with the prompt below to create your users. Feel free to customize the prompt for locations and departments that more match your environment if needed.

```
I need to create a Microsoft 365 test environment with 25 users. There must be a mixture of locations but they can only be in NY, FL, and CA. There must be a mixture of departments, but they can only be IT, HR, Accounting, and Marketing. The fields to create values for are FirstName, LastName, UserName, Title, Department, StreetAddress, City, State, PostalCode, Country, PhoneNumber, MobilePhone. The phone number and mobile number area codes should match the city and state they are in. This should be able to be pasted into an excel document.
```

### Getting Started

Once you have created your 365 Developer Program Environment, you can start adding users and groups.

1. Install the module by
2. Save the 365DataEnvironment.xlsx (located in LabSources) file on to your system
3. Run the below command to add users to your environment with their licensing

   1. ```powershell
      Add-CT365User -FilePath "C:\Path\to\365DataEnvironment.xlsx" -Domain "yourdomain.onmicrosoft.com"
      ```
4. Run the below command to add groups to your environment.

   1. ```powershell
      Add-CT365Group -FilePath "C:\Path\to\365DataEnvironment.xlsx" -UserPrincialName "user@yourdomain.onmicrosoft.com" -Domain "yourdomain.onmicrosoft.com"
      ```

   > The UserPrincipalName will be the email that you created for the admin of the tenant, or an account who has rights to create groups
   >
5. Run the below command to add a user to their groups per their location and title.

   1. ```powershell
      Add-CT365GroupByTitle -FilePath "C:\Path\to\365DataEnvironment.xlsx" -UserEmail "jdoe@yourdomain.onmicrosoft.com" -Domain "yourdomain.onmicrosoft.com" -UserRole "NY-IT"
      ```
