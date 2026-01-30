<#
.SYNOPSIS
    Creates a single hybrid user with remote mailbox in M365.

.DESCRIPTION
    Creates an on-premises user with a remote mailbox, copies group membership
    from template user, and sets manager.

.PARAMETER TemplateUser
    The username of the template user to copy groups from

.PARAMETER UserAlias
    The alias for the new user (used for mailbox and AD account)

.PARAMETER DisplayName
    The display name for the new user

.PARAMETER FirstName
    First name of the new user

.PARAMETER LastName
    Last name of the new user

.PARAMETER JobTitle
    Job title of the new user

.PARAMETER Department
    Department of the new user

.PARAMETER Office
    Office location of the new user

.PARAMETER Description
    Description of the new user

.PARAMETER Company
    Company name for the new user

.PARAMETER ManagerName
    Display name or UPN of the manager to assign

.PARAMETER OU
    Organizational Unit where the user will be created

.PARAMETER UPN
    User Principal Name (email address format)

.PARAMETER Password
    Initial password (SecureString)

.PARAMETER RemoteDomain
    Remote routing domain (e.g., contoso.mail.onmicrosoft.com)
#>

Param(
    [Parameter(Mandatory=$true)]
    [string]$TemplateUser,
    
    [Parameter(Mandatory=$true)]
    [string]$UserAlias,
    
    [Parameter(Mandatory=$true)]
    [string]$DisplayName,
    
    [Parameter(Mandatory=$true)]
    [string]$FirstName,
    
    [Parameter(Mandatory=$true)]
    [string]$LastName,
    
    [string]$JobTitle,
    
    [string]$Department,
    
    [string]$Office,
    
    [string]$Description,
    
    [string]$Company,
    
    [string]$ManagerName,
    
    [Parameter(Mandatory=$true)]
    [string]$OU,
    
    [Parameter(Mandatory=$true)]
    [string]$UPN,
    
    [Parameter(Mandatory=$true)]
    [SecureString]$Password,
    
    [Parameter(Mandatory=$true)]
    [string]$RemoteDomain
)

try {
    # Create the new on-premises user with a remote mailbox
    Write-Host "  → Creating remote mailbox..."
    
    $NewMailboxParams = @{
        Alias = $UserAlias
        Name = $DisplayName
        FirstName = $FirstName
        LastName = $LastName
        OrganizationalUnit = $OU
        UserPrincipalName = $UPN
        Password = $Password
        ResetPasswordOnNextLogon = $true  # Enforce password reset on first login
    }
    
    # Add optional parameters if provided
    if (-not [string]::IsNullOrWhiteSpace($RemoteDomain)) {
        $NewMailboxParams['RemoteRoutingAddress'] = "$UserAlias@$RemoteDomain"
    }
    
    New-RemoteMailbox @NewMailboxParams -ErrorAction Stop | Out-Null
    Write-Host "    ✓ Remote mailbox created" -ForegroundColor Green
    
    # Set additional attributes (JobTitle, Department, Office, Description, Company)
    if (-not [string]::IsNullOrWhiteSpace($JobTitle) -or 
        -not [string]::IsNullOrWhiteSpace($Department) -or 
        -not [string]::IsNullOrWhiteSpace($Office) -or
        -not [string]::IsNullOrWhiteSpace($Description) -or
        -not [string]::IsNullOrWhiteSpace($Company)) {
        
        Write-Host "  → Setting user attributes..."
        
        $SetUserParams = @{
            Identity = $UPN
        }
        
        if (-not [string]::IsNullOrWhiteSpace($JobTitle)) { $SetUserParams['Title'] = $JobTitle }
        if (-not [string]::IsNullOrWhiteSpace($Department)) { $SetUserParams['Department'] = $Department }
        if (-not [string]::IsNullOrWhiteSpace($Office)) { $SetUserParams['Office'] = $Office }
        if (-not [string]::IsNullOrWhiteSpace($Description)) { $SetUserParams['Description'] = $Description }
        if (-not [string]::IsNullOrWhiteSpace($Company)) { $SetUserParams['Company'] = $Company }
        
        Set-AdUser @SetUserParams -ErrorAction Stop
        Write-Host "    ✓ Attributes set" -ForegroundColor Green
    }
    
    # Copy group membership from template user
    Write-Host "  → Copying group(s) from template..."
    
    $TemplateGroups = Get-AdUser -Identity $TemplateUser -ErrorAction Stop | Get-AdMemberOf -ErrorAction Stop
    
    $GroupCount = 0
    foreach ($Group in $TemplateGroups) {
        try {
            Add-AdGroupMember -Identity $Group -Members $UPN -ErrorAction Stop
            $GroupCount++
        }
        catch {
            Write-Warning "  ! Failed to add $UPN to group $($Group.Name): $_"
        }
    }
    Write-Host "    ✓ Copied $GroupCount group(s)" -ForegroundColor Green
    
    # Set manager if provided
    if (-not [string]::IsNullOrWhiteSpace($ManagerName)) {
        Write-Host "  → Setting manager..."
        
        try {
            $Manager = Get-AdUser -Identity $ManagerName -ErrorAction Stop
            Set-AdUser -Identity $UPN -Manager $Manager -ErrorAction Stop
            Write-Host "    ✓ Manager set ($($Manager.DisplayName))" -ForegroundColor Green
        }
        catch {
            Write-Warning "  ! Failed to set manager: $_"
        }
    }
    
    # Assign M365 license from template user
    Write-Host "  → Assigning M365 license..."
    
    try {
        # Get template user's license SKU from M365
        $TemplateM365User = Get-MgUser -Filter "userPrincipalName eq '$TemplateUser'" -ErrorAction Stop
        $TemplateLicenses = Get-MgUserLicenseDetail -UserId $TemplateM365User.Id -ErrorAction Stop
        
        if ($TemplateLicenses.Count -gt 0) {
            # Use first license SKU from template
            $SkuId = $TemplateLicenses[0].SkuId
            
            # Wait for new user to appear in M365
            Start-Sleep -Seconds 5
            
            $NewM365User = Get-MgUser -Filter "userPrincipalName eq '$UPN'" -ErrorAction Stop
            Set-MgUserLicense -UserId $NewM365User.Id -AddLicenses @{SkuId = $SkuId} -RemoveLicenses @() -ErrorAction Stop
            Write-Host "    ✓ License assigned" -ForegroundColor Green
        }
        else {
            Write-Warning "  ! Template user has no licenses to copy"
        }
    }
    catch {
        Write-Warning "  ! Failed to assign license: $_"
    }
    
    return $true
}
catch {
    Write-Error "Failed to create user $UPN : $_"
    return $false
}