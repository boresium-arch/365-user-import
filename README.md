# 365-user-import
A set of PowerShell scripts for Office 365 user import.

## Plan and usage guide
See [docs/365-user-import-guide.md](docs/365-user-import-guide.md) for input validation, schema mapping, prerequisites, error handling, and examples.

## Scripts
- [Bulk-HybridUser.ps1](Bulk-HybridUser.ps1): Orchestrates Excel import, validation, on-premises user creation with remote mailbox, group copying, manager setting, sync, and license assignment.

## Prerequisites
- PowerShell 7+
- Exchange Management Shell (for `New-RemoteMailbox`)
- Active Directory cmdlets (for user/group/manager operations)
- Azure AD Connect (for `Start-ADSyncSyncCycle`)
- ImportExcel module (PowerShell Gallery)
- Microsoft Graph PowerShell modules (Users + Users.Actions)

## Workflow
- Prompts for Excel file path (with file picker popup).
- Auto-discovers primary domain and remote routing domain from Microsoft 365 tenant.
- Validates Excel path, `.xlsx` extension, and strict schema (no extra columns).
- For each user: creates on-premises user with remote mailbox, sets attributes, copies groups, sets manager, assigns M365 license.
- Runs Azure AD Connect delta sync and waits for cloud availability.
- Writes a transcript to `logs/` and displays final summary in a popup.

## Runtime files
Place live Excel inputs in the `runtime/` folder (e.g., `runtime/GXP IMPORT.xlsx`).

## Excel file fields
COMPANY, OU, FirstName, DisplayName, LastName, DEPARTMENT, OFFICE, JOBTITLE, DESCRIPTION, MANAGERNAME, UserAlias, UPN, TEMPLATEUSER, Password

## Field notes
- OU: Organizational unit path (can be left empty to derive from TEMPLATEUSER)
- UPN: Full User Principal Name including domain (e.g., user@grymca.org)
- TEMPLATEUSER: UPN of template user to copy groups and licenses from
- DESCRIPTION: Additional user description field
- Domain enforcement: UPN domain must match the provided `Domain` value
