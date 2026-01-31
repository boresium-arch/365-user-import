# 365-user-import AI Coding Instructions

## Project Purpose
This project creates Office 365 users from Excel files in a **hybrid Exchange environment** (Microsoft 365 + local AD). The orchestration script reads user data from Excel, derives organizational units from template users in local AD, and provisions new users in Microsoft 365. Licenses are assigned based on the template user's existing Microsoft 365 license. The solution emphasizes robust input validation, error handling, and logging to ensure reliable bulk user creation.

## Technical Requirements and Constraints ##
- **PowerShell 7+**: Leverage modern PowerShell features and modules.
- **ImportExcel module**: For reading Excel files without needing Excel installed.
- **Microsoft Graph PowerShell**: For Microsoft 365 user and license management.
- **Local AD access**: To resolve TemplateUser OUs.
- **Excel file schema**: Fixed structure with required columns; no deviations allowed.
- **Error handling**: Fail fast on critical errors, skip invalid rows with warnings, log all actions.
- **Logging**: Comprehensive logging of all operations, including a final summary.
- **Documentation-first**: All scripts and modules must be well-documented; see the provided guide.
- **Hybrid environment support**: Must function correctly in environments with both Microsoft 365 and on-premises Active Directory.
- **Template-based provisioning**: New users inherit attributes and licenses from specified template usersa

## Architecture & Data Flow
1. **Input**: Excel file with user data including COMPANY, OU, FirstName, DisplayName, LastName, DEPARTMENT, OFFICE, JOBTITLE, DESCRIPTION, MANAGERNAME, UserAlias, UPN, TEMPLATEUSER, Password
2. **Processing**: Validate inputs → lookup TemplateUser OU in local AD → create M365 user → assign M365 license using a lookup of TemplateUser license → set manager 
3. **Integration points**: 
   - ImportExcel module (PSGallery) for reading Excel files
   - Microsoft Graph PowerShell for M365 user creation and license assignment
   - Local AD for TemplateUser OU resolution

## Excel Schema (Fixed Structure)
Required columns: `COMPANY`, `OU`, `FirstName`, `DisplayName`, `LastName`, `DEPARTMENT`, `OFFICE`, `JOBTITLE`, `DESCRIPTION`, `MANAGERNAME`, `UserAlias`, `UPN`, `TEMPLATEUSER`, `Password`

**Field notes**:
- `OU`: The organizational unit path where the user will be created (can be left empty to derive from TEMPLATEUSER)
- `UPN`: Full User Principal Name (e.g., user@grymca.org)
- `TEMPLATEUSER`: UPN of template user to copy group membership and license from
- `DESCRIPTION`: Additional user description field
- `Password`: Initial password (user must reset on first login)

See [examples/users-template.csv](../examples/users-template.csv) for the exact format.

## Critical Workflows

### Script Execution Flow
1. Prompt for: Excel path (domain and UPN are in the file)
2. **Validation** (fail fast):
   - Excel path exists with .xlsx extension
   - Schema matches required columns exactly (COMPANY, OU, FirstName, DisplayName, LastName, DEPARTMENT, OFFICE, JOBTITLE, DESCRIPTION, MANAGERNAME, UserAlias, UPN, TEMPLATEUSER, Password)
3. Per-user processing:
   - Use OU from Excel file (or derive from TEMPLATEUSER if OU is empty)
   - Create remote mailbox in M365 with UPN from Excel
   - Set user attributes: JOBTITLE, DEPARTMENT, OFFICE, DESCRIPTION, COMPANY
   - Copy group membership from TEMPLATEUSER
   - Resolve MANAGERNAME to set manager
   - Apply initial password with reset-on-first-login policy
   - Retrieve TEMPLATEUSER license SKU from M365 and assign to new user 

### Error Handling Pattern
- **Fail fast** on missing inputs or schema mismatch
- **Skip invalid rows** with warnings, continue processing
- **Catch and log** Graph/AD exceptions (user exists, invalid OU, permissions)
- **Emit final summary**: Created/Skipped/Failed counts

### Logging Requirements
- Start PowerShell transcript for entire run
- Log per-user result with action taken and reason
- Provide final summary count at end

## Prerequisites & Setup
```powershell
# PowerShell 7+ recommended
Install-Module ImportExcel -Scope CurrentUser
Install-Module Microsoft.Graph -Scope CurrentUser
```

Permissions needed:
- Microsoft 365 user creation rights
- Local AD read access for TemplateUser OU lookup
- Microsoft Graph scopes for user and license management

## Project Conventions
- **Documentation-first approach**: See comprehensive planning in [docs/365-user-import-guide.md](../docs/365-user-import-guide.md)
- **Validation before action**: All inputs validated before any changes
- **Hybrid environment assumption**: Scripts must handle both M365 and local AD
- **Template-based provisioning**: All user attributes (OU, licenses) derived from TemplateUser pattern
- **CSV as reference format**: While Excel (.xlsx) is processed, CSV template shows structure

## Key Files
- [docs/365-user-import-guide.md](../docs/365-user-import-guide.md): Complete implementation guide with validation, error handling, and examples
- [examples/users-template.csv](../examples/users-template.csv): Excel schema reference
- [examples/sample-run.md](../examples/sample-run.md): Expected execution flow
