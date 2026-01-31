# Sample run (illustrative)

1. Script prompts for:
   - Excel path: C:\path\to\runtime\GXP IMPORT.xlsx

2. Validation:
   - File exists: ✓
   - Extension is .xlsx: ✓
   - Schema matches (COMPANY, OU, FirstName, DisplayName, LastName, DEPARTMENT, OFFICE, JOBTITLE, DESCRIPTION, MANAGERNAME, UserAlias, UPN, TEMPLATEUSER, Password): ✓
   - Found 2 users to process

3. Processing:
   - User 1: avery.johnson@grymca.org
     → Resolved OU: grymca.org/New OU Structure/Mary Free Bed Y/106 - Health & Wellness
     → Created remote mailbox: ✓
     → Set attributes (JOBTITLE, DEPARTMENT, OFFICE, DESCRIPTION, COMPANY): ✓
     → Copied 3 group(s) from template: ✓
     → Set manager (Jacob DeLosh): ✓
     → Assigned license (ENTERPRISEPACK): ✓
   
   - User 2: chris.lee@grymca.org
     → Resolved OU: grymca.org/New OU Structure/Mary Free Bed Y/106 - Health & Wellness
     → Created remote mailbox: ✓
     → Set attributes: ✓
     → Copied 3 group(s) from template: ✓
     → Set manager (Jacob DeLosh): ✓
     → Assigned license (ENTERPRISEPACK): ✓

4. Result summary:
   - Created: 2
   - Skipped: 0
   - Failed: 0
