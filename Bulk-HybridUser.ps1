[CmdletBinding()]
Param(
	[string]$CsvPath,
	[string]$UsageLocation = "US"
)

# On-premises section requirements:
# Run this from your Hybrid Exchange Server 2013 or 2016 (requires Exchange management shell)
# Azure AD Connect should be present for the Start-ADSyncSyncCycle cmdlet to run

$ErrorActionPreference = "Stop"

# Check and install required modules
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12   
$RequiredModules = @('Microsoft.Graph.Users', 'Microsoft.Graph.Users.Actions')
foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -Name $Module -ListAvailable)) {
        Write-Host "Installing module: $Module"
        Install-Module -Name $Module -Scope CurrentUser -Force -ErrorAction Stop
    }
}

# Load Exchange snap-in
try {
    Write-Host "Loading the PowerShell Snap-in for Exchange Management"
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
}
catch {
    throw "Exchange snap-in could not be loaded. Ensure you are running in the Exchange Management Shell."
}

# Import modules
Write-Host "Loading Microsoft Graph modules"
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Users.Actions -ErrorAction Stop

if (-not $CsvPath) {
	try {
		Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
		$Dialog = New-Object System.Windows.Forms.OpenFileDialog
		$Dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
		$Dialog.Title = "Select CSV file"
		$Dialog.InitialDirectory = (Join-Path $PSScriptRoot "runtime")
		$Dialog.Multiselect = $false

		if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
			$CsvPath = $Dialog.FileName
		}
	}
	catch {
		Write-Warning "File picker unavailable. Falling back to input dialog."
	}

	if (-not $CsvPath) {
		Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
		$CsvPath = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the CSV path (.csv)", "CSV Path Input", "")
	}
}

if (-not (Test-Path -Path $CsvPath)) {
	throw "CSV path not found: $CsvPath"
}

if ([System.IO.Path]::GetExtension($CsvPath) -ne ".csv") {
	throw "CSV file must have .csv extension: $CsvPath"
}

$LogDir = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path -Path $LogDir)) {
	New-Item -Path $LogDir -ItemType Directory | Out-Null
}

$LogPath = Join-Path $LogDir ("bulk-run-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $LogPath -Append | Out-Null

Write-Host "Connecting to Microsoft Graph"
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All", "Domain.Read.All" -NoWelcome

Write-Host "Discovering domain and remote domain from Microsoft 365"
$Domains = Get-MgDomain -ErrorAction Stop
$PrimaryDomain = ($Domains | Where-Object { $_.IsVerified -and $_.IsDefault }).Id
if (-not $PrimaryDomain) {
    throw "Could not find primary verified domain in Microsoft 365 tenant."
}
if (-not $Domain) {
    $Domain = $PrimaryDomain
}
if (-not $RemoteDomain) {
    if ($PrimaryDomain -match '\.onmicrosoft\.com$') {
        $RemoteDomain = $PrimaryDomain -replace '\.onmicrosoft\.com$', '.mail.onmicrosoft.com'
    } else {
        $TenantName = ($PrimaryDomain -split '\.')[0]
        $RemoteDomain = "$TenantName.mail.onmicrosoft.com"
    }
}

$Rows = Import-Csv -Path $CsvPath -ErrorAction Stop

if (-not $Rows -or $Rows.Count -eq 0) {
	throw "No rows found in CSV file: $CsvPath"
}

$RequiredColumns = @(
	"COMPANY", "FIRSTNAME", "DISPLAYNAME", "LASTNAME",
	"DEPARTMENT", "OFFICE", "JOBTITLE", "DESCRIPTION", "MANAGERNAME",
	"USERALIAS", "UPN", "TEMPLATEUSER", "PASSWORD"
)

$ActualColumns = $Rows | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
$ActualUpper = $ActualColumns | ForEach-Object { $_.ToUpperInvariant() }

$MissingColumns = $RequiredColumns | Where-Object { $_ -notin $ActualUpper }
$UnexpectedColumns = $ActualUpper | Where-Object { $_ -notin $RequiredColumns }

if ($MissingColumns.Count -gt 0 -or $UnexpectedColumns.Count -gt 0) {
	if ($MissingColumns.Count -gt 0) {
		Write-Error "Missing required columns: $($MissingColumns -join ', ')"
	}
	if ($UnexpectedColumns.Count -gt 0) {
		Write-Error "Unexpected columns present: $($UnexpectedColumns -join ', ')"
	}
	throw "CSV schema does not match required columns."
}

$Created = 0
$Skipped = 0
$Failed = 0

Write-Host "Found $($Rows.Count) user(s) to process from CSV"

foreach ($Row in $Rows) {
	$Upn = $Row.UPN
	$TemplateUser = $Row.TEMPLATEUSER
	$UserAlias = $Row.UserAlias
	$PasswordPlain = $Row.Password

	if ([string]::IsNullOrWhiteSpace($Upn) -or
		[string]::IsNullOrWhiteSpace($TemplateUser) -or
		[string]::IsNullOrWhiteSpace($UserAlias) -or
		[string]::IsNullOrWhiteSpace($PasswordPlain)) {
		Write-Warning "Skipping row due to missing required values (UPN, TEMPLATEUSER, UserAlias, Password)."
		$Skipped++
		continue
	}

	if ($Upn -notmatch "@" -or ($Upn -split "@")[-1].ToLowerInvariant() -ne $Domain.ToLowerInvariant()) {
		Write-Warning "UPN '$Upn' does not match domain '$Domain'. Skipping."
		$Skipped++
		continue
	}

	$DisplayName = $Row.DisplayName
	if ([string]::IsNullOrWhiteSpace($DisplayName)) {
		if (-not [string]::IsNullOrWhiteSpace($Row.FirstName) -and -not [string]::IsNullOrWhiteSpace($Row.LastName)) {
			$DisplayName = "$($Row.FirstName) $($Row.LastName)"
		}
		else {
			Write-Warning "Skipping row: missing DisplayName and FirstName/LastName for $Upn."
			$Skipped++
			continue
		}
	}

	Write-Host "Processing $Upn" -ForegroundColor Cyan

	try {
		$SecurePassword = ConvertTo-SecureString -String $PasswordPlain -AsPlainText -Force

		# Resolve OU from template user
		Write-Host "  → Resolving OU from template user..."
		$TemplateAdUser = Get-AdUser -Filter {UserPrincipalName -eq $TemplateUser} -Properties DistinguishedName, UserPrincipalName -ErrorAction Stop
		$OrganizationalUnit = $TemplateAdUser.DistinguishedName -replace '^CN=.*?,'
		
		if ([string]::IsNullOrWhiteSpace($OrganizationalUnit)) {
			throw "Failed to resolve OU from template user: $TemplateUser"
		}
		
		# Convert DN to canonical name
		$parts = $OrganizationalUnit -split ','
		$domainParts = $parts | Where-Object { $_ -like 'DC=*' } | ForEach-Object { $_.Substring(3) }
		$domain = $domainParts -join '.'
		$ouParts = $parts | Where-Object { $_ -like 'OU=*' } | ForEach-Object { $_.Substring(3) }
		[array]::Reverse($ouParts)
		$OrganizationalUnit = $domain + '/' + ($ouParts -join '/')
		
		catch {
			throw "Invalid OU: $OrganizationalUnit. $_"
		}
		if ($TemplateUser -notmatch '@') {
			if (-not $TemplateAdUser) {
				$TemplateAdUser = Get-AdUser -Filter {UserPrincipalName -eq $TemplateUser} -Properties UserPrincipalName -ErrorAction Stop
			}
			$TemplateUserUpn = $TemplateAdUser.UserPrincipalName
		}
		
		# Create the new on-premises user with a remote mailbox
		Write-Host "  → Creating remote mailbox..."
		
		$NewMailboxParams = @{
			Alias = $UserAlias
			Name = $DisplayName
			FirstName = $Row.FirstName
			LastName = $Row.LastName
			OrganizationalUnit = $OrganizationalUnit
			UserPrincipalName = $Upn
			Password = $SecurePassword
			ResetPasswordOnNextLogon = $true  # Enforce password reset on first login
		}
		
		# Add optional parameters if provided
		if (-not [string]::IsNullOrWhiteSpace($RemoteDomain)) {
			$NewMailboxParams['RemoteRoutingAddress'] = "$UserAlias@$RemoteDomain"
		}
		
		New-RemoteMailbox @NewMailboxParams -ErrorAction Stop | Out-Null
		Write-Host "    ✓ Remote mailbox created" -ForegroundColor Green
		
		# Set additional attributes (JobTitle, Department, Office, Description, Company)
		if (-not [string]::IsNullOrWhiteSpace($Row.JOBTITLE) -or 
			-not [string]::IsNullOrWhiteSpace($Row.DEPARTMENT) -or 
			-not [string]::IsNullOrWhiteSpace($Row.OFFICE) -or
			-not [string]::IsNullOrWhiteSpace($Row.DESCRIPTION) -or
			-not [string]::IsNullOrWhiteSpace($Row.COMPANY)) {
			
			Write-Host "  → Setting user attributes..."
			
			$SetUserParams = @{
				Identity = $Upn
			}
			
			if (-not [string]::IsNullOrWhiteSpace($Row.JOBTITLE)) { $SetUserParams['Title'] = $Row.JOBTITLE }
			if (-not [string]::IsNullOrWhiteSpace($Row.DEPARTMENT)) { $SetUserParams['Department'] = $Row.DEPARTMENT }
			if (-not [string]::IsNullOrWhiteSpace($Row.OFFICE)) { $SetUserParams['Office'] = $Row.OFFICE }
			if (-not [string]::IsNullOrWhiteSpace($Row.DESCRIPTION)) { $SetUserParams['Description'] = $Row.DESCRIPTION }
			if (-not [string]::IsNullOrWhiteSpace($Row.COMPANY)) { $SetUserParams['Company'] = $Row.COMPANY }
			
			Set-AdUser @SetUserParams -ErrorAction Stop
			Write-Host "    ✓ Attributes set" -ForegroundColor Green
		}
		
		# Copy group membership from template user
		Write-Host "  → Copying group(s) from template..."
		
		$TemplateGroups = Get-AdUser -Identity $TemplateUser -ErrorAction Stop | Get-AdMemberOf -ErrorAction Stop
		
		$GroupCount = 0
		foreach ($Group in $TemplateGroups) {
			try {
				Add-AdGroupMember -Identity $Group -Members $Upn -ErrorAction Stop
				$GroupCount++
			}
			catch {
				Write-Warning "  ! Failed to add $Upn to group $($Group.Name): $_"
			}
		}
		Write-Host "    ✓ Copied $GroupCount group(s)" -ForegroundColor Green
		
		# Set manager if provided
		if (-not [string]::IsNullOrWhiteSpace($Row.MANAGERNAME)) {
			Write-Host "  → Setting manager..."
			
			try {
				$Manager = Get-AdUser -Identity $Row.MANAGERNAME -ErrorAction Stop
				Set-AdUser -Identity $Upn -Manager $Manager -ErrorAction Stop
				Write-Host "    ✓ Manager set ($($Manager.DisplayName))" -ForegroundColor Green
			}
			catch {
				Write-Warning "  ! Failed to set manager: $_"
			}
		}
		
		# Assign M365 license from template user
		Write-Host "  → Assigning M365 license..."
		
		try {
			# Wait briefly for user to sync to M365 if just created
			Write-Host "    → Checking if user exists in M365..."
			$MaxAttempts = 3
			$Attempt = 0
			$MgUser = $null
			
			while ($Attempt -lt $MaxAttempts -and -not $MgUser) {
				try {
					$MgUser = Get-MgUser -UserId $Upn -ErrorAction Stop
					Write-Host "      ✓ User found in M365" -ForegroundColor Green
				}
				catch {
					$Attempt++
					if ($Attempt -lt $MaxAttempts) {
						Write-Host "      ⓘ User not found, waiting 10 seconds... (Attempt $Attempt/$MaxAttempts)" -ForegroundColor Gray
						Start-Sleep -Seconds 10
					}
				}
			}
			
			if (-not $MgUser) {
				Write-Warning "    User not found in M365. Ensure Azure AD Connect sync has completed."
			}
			else {
				# Set usage location
				Write-Host "    → Setting usage location to: $UsageLocation"
				Update-MgUser -UserId $Upn -UsageLocation $UsageLocation -ErrorAction Stop
				Write-Host "      ✓ Usage location set" -ForegroundColor Green
				
				# Resolve template user and get their license
				Write-Host "    → Resolving template user license..."
				
				# Get template user's licenses from M365
				$TemplateUserMg = Get-MgUser -UserId $TemplateUserUpn -Property AssignedLicenses -ErrorAction Stop
				
				if ($TemplateUserMg.AssignedLicenses.Count -eq 0) {
					Write-Warning "    Template user has no licenses assigned"
				}
				else {
					# Get the first license SKU ID from template user
					$LicenseSkuId = $TemplateUserMg.AssignedLicenses[0].SkuId
					
					# Get the SKU details to display the name
					$Sku = Get-MgSubscribedSku | Where-Object { $_.SkuId -eq $LicenseSkuId }
					
					Write-Host "      ✓ Found license: $($Sku.SkuPartNumber)" -ForegroundColor Green
					
					# Assign the license
					Write-Host "    → Assigning license..."
					
					$AddLicenses = @{
						SkuId = $LicenseSkuId
					}
					
					Set-MgUserLicense -UserId $Upn -AddLicenses @($AddLicenses) -RemoveLicenses @() -ErrorAction Stop
					
					Write-Host "      ✓ License assigned successfully" -ForegroundColor Green
				}
			}
		}
		catch {
			Write-Warning "  ! Failed to assign license: $_"
		}

		$Created++
	}
	catch {
		Write-Warning "Failed to create user $Upn : $_"
		$Failed++
	}
}

Write-Host "\nSummary (On-Premises): Created=$Created  Skipped=$Skipped  Failed=$Failed" -ForegroundColor Yellow

Write-Host "Running a delta sync from Azure AD Connect"
try {
	Start-ADSyncSyncCycle -PolicyType Delta
}
catch {
	Write-Warning "Start-ADSyncSyncCycle failed or not available. Ensure Azure AD Connect is installed."
}

#Start-Sleep -Seconds 60

Write-Host "\nFinal Summary: Created=$Created  Skipped=$Skipped  Failed=$Failed" -ForegroundColor Green

Read-Host "Press Enter to exit"

Stop-Transcript | Out-Null