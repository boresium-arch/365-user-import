Plan

1. Review existing PowerShell scripts and modules used.
2. Validate required inputs (Excel path with GXP Import.xlsx schema).
3. Confirm Excel schema: COMPANY, OU, FirstName, DisplayName, LastName, DEPARTMENT, OFFICE, JOBTITLE, DESCRIPTION, MANAGERNAME, UserAlias, UPN, TEMPLATEUSER, Password.
4. Process each row: use OU from file (or derive from TEMPLATEUSER), create user with all attributes, copy groups, set manager, assign license.
5. Document prerequisites, setup, and execution steps.
6. Add error handling and logging notes.
7. Provide example Excel template and sample run.
