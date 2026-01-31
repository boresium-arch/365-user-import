# 365-user-import plan and usage guide

## 1) Review existing PowerShell scripts and modules
- This repo currently contains documentation only. When scripts are added, ensure they:
  - Import the ImportExcel module (from PowerShell Gallery).
  - Import the Microsoft Graph PowerShell module.
  - Connect to Microsoft Graph and Exchange/AD as required for a hybrid environment
  - Read from the Excel input and create Microsoft 365 users.

## 2) Required input validation
Validate these inputs before any changes are made:
- Domain (e.g., contoso.com)
- Remote domain (e.g., contoso.mail.onmicrosoft.com)
- Excel path (must exist and have the required columns)

Validation checklist:
- Domain values are non-empty and DNS-like (contain at least one dot).
- Excel path exists and file extension is .xlsx.
- Excel worksheet contains the expected schema (see below).

## 3) Excel schema and calculated fields
Required columns (case-insensitive):
- COMPANY
- OU
- FirstName
- DisplayName
- LastName
- DEPARTMENT
- OFFICE
- JOBTITLE
- DESCRIPTION
- MANAGERNAME
- UserAlias
- UPN
- TEMPLATEUSER
- Password

Calculated fields:
- `OU` should be derived by looking up the OU of the `TemplateUser` in local AD when blank.

Mapping notes:
- `UPN` is required and must include the domain (e.g., user@domain).
- `UserAlias` becomes the mailbox/AD alias.
- `DisplayName` should be used as-is when provided; otherwise derive from FirstName + LastName.
- `ManagerName` should be resolved to an AD user for manager assignment.
- `Password` should be used only for initial set; enforce reset at first sign-in.

## 4) Prerequisites, setup, and execution steps
Prerequisites:
- PowerShell 7+ recommended.
- ImportExcel module installed from PSGallery.
- Microsoft Graph PowerShell module installed.
- Appropriate permissions for:
  - Creating users in Microsoft 365.
  - Reading local AD (to resolve OU for TemplateUser).

Setup:
1. Install modules:
   - Install-Module ImportExcel -Scope CurrentUser
   - Install-Module Microsoft.Graph -Scope CurrentUser
2. Ensure you can authenticate to Microsoft Graph with required scopes.

Execution (typical flow):
1. Prompt for domain, remote domain, and Excel file path (place live inputs in `runtime/`).
2. Validate inputs and schema.
3. Read Excel rows and map fields.
4. Resolve TemplateUser OU (if OU not provided).
5. Create on-premises users and set attributes.
6. Assign manager and copy groups from TemplateUser.
7. Trigger Azure AD Connect sync.
8. Assign a Microsoft 365 license SKU matching the TemplateUser.

## 5) Error handling and logging notes
Error handling:
- Fail fast on missing inputs or schema mismatch.
- Validate each row and skip invalid rows with a warning summary.
- Catch and log Graph/AD exceptions (user already exists, invalid OU, permission issues).

Logging:
- Start a transcript for the run.
- Log a per-user result (Created/Skipped/Failed) with reason.
- Emit a final summary count.

## 6) Example Excel template and sample run
See:
- [examples/users-template.csv](../examples/users-template.csv)
- [examples/sample-run.md](../examples/sample-run.md)
