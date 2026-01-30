#On-premises section requirements: 
#Run this from your Hybrid Exchange Server 2013 or 2016 (requires Exchange management shell)
#Azure AD Connect should be present for the Start-ADSyncSyncCycle cmdlet to run

#Add the Exchange PS Snap-in
write-host "Loading the PowerShell Snap-in for Exchange Management"
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

#Bulk-create hybrid users on-premises using CSV import
write-host "Creating user(s) from CSV file"
IMPORT-CSV GXPimport.csv | ForEach-Object { ./New-HybridUser.PS1 -TemplateUser $_.TemplateUser -UserAlias $_.UserAlias -DisplayName $_.DisplayName -FirstName $_.FirstName -LastName $_.LastName -UPN $_.UPN -Password (ConvertTo-SecureString -String $_.Password -AsPlainText -Force) -OU $_.OU }

#Force a sync of Azure AD Connect so the new users show up in the MSOL tenant
#This command requires you to be on the server with Azure AD Connect installed
write-host "Running a delta sync from Azure AD Connect"
Start-ADSyncSyncCycle -PolicyType Delta

Start-Sleep -Seconds 60

#Cloud section requirements: 
#Install the Microsoft Graph PowerShell module
#Requires scopes: User.ReadWrite.All, Directory.ReadWrite.All

#Connect to Microsoft Graph
write-host "Connecting to Microsoft Graph"
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Users.Actions
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All" -NoWelcome

#Bulk-assign licenses in Office 365 using CSV import
write-host "Assigning user licenses and activating the mailboxes"
IMPORT-CSV GXPimport.csv | ForEach-Object { ./Assign-License.PS1 -UPN $_.UPN -TemplateUser $_.TemplateUser }