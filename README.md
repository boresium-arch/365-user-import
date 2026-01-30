# 365-user-import
A set of PowerShell scripts for Office 365 user import.

## Plan and usage guide
See [docs/365-user-import-guide.md](docs/365-user-import-guide.md) for input validation, schema mapping, prerequisites, error handling, and examples.

# function
Installs and uses the ImportExcel PowerShell module from the PowerShell Gallery.
Installs and uses the Microsoft Graph PowerShell module.
Works in a Hybrid Exchange environment with mailboxes and users in Microsoft 365 and local AD domain users.
The orchestration script prompts the user for the domain name and the remote domain name.
Creates new Office 365 users from a list of users in an Excel file.

# excel file fields
COMPANY, OU, FirstName, DisplayName, LastName, DEPARTMENT, OFFICE, JOBTITLE, DESCRIPTION, MANAGERNAME, UserAlias, UPN, TEMPLATEUSER, Password

# field notes
- OU: Organizational unit path (can be left empty to derive from TEMPLATEUSER)
- UPN: Full User Principal Name including domain (e.g., user@grymca.org)
- TEMPLATEUSER: UPN of template user to copy groups and licenses from
- DESCRIPTION: Additional user description field
