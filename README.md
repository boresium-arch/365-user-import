# 365-user-import
A set of PowerShell scripts for Office 365 user import.

# function
Installs and uses the ImportExcel PowerShell module from the PowerShell Gallery.
Installs and uses the Microsoft Graph PowerShell module.
Works in a Hybrid Exchange environment with mailboxes and users in Microsoft 365 and local AD domain users.
The orchestration script prompts the user for the domain name and the remote domain name.
Creates new Office 365 users from a list of users in an Excel file.

# excel file fields
FirstName, LastName, DisplayName, Department, Office, ManagerName, UserAlias, TemplateUser, Password

# calculated fields
OrganizationalUnit for a new user is found using the OU of the TemplateUser
