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

.PARAMETER ManagerName
    Display name or UPN of the manager to assign

.PARAMETER OU
    Organizational Unit where the user will be created

.PARAMETER UPN
    User Principal Name (email address format)

.PARAMETER Password
    Initial password (SecureString)

.PARAMETER Domain
    Primary domain (e.g., contoso.com)

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
    
    [string]$ManagerName,
    
    [Parameter(Mandatory=$true)]
    [string]$OU,
    
    [Parameter(Mandatory=$true)]
    [string]$UPN,
    
    [Parameter(Mandatory=$true)]
    [SecureString]$Password,
    
    [Parameter(Mandatory=$true)]
    [string]$Domain,
    
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
        OnPremisesOrganizationalUnit = $OU
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
    
    # Set additional attributes (JobTitle, Department, Office)
    if (-not [string]::IsNullOrWhiteSpace($JobTitle) -or 
        -not [string]::IsNullOrWhiteSpace($Department) -or 
        -not [string]::IsNullOrWhiteSpace($Office)) {
        
        Write-Host "  → Setting user attributes..."
        
        $SetUserParams = @{}
        if ($JobTitle) { $SetUserParams['Title'] = $JobTitle }
        if ($Department) { $SetUserParams['Department'] = $Department }
        if ($Office) { $SetUserParams['Office'] = $Office }
        
        Set-ADUser -Identity $UserAlias @SetUserParams -ErrorAction Stop
        Write-Host "    ✓ Attributes set" -ForegroundColor Green
    }
    
    # Get group membership from template user
    Write-Host "  → Copying group membership from template user..."
    
    $UserGroups = @()
    $UserGroups = (Get-ADUser -Identity $TemplateUser -Properties MemberOf -ErrorAction Stop).MemberOf
    
    if ($UserGroups.Count -gt 0) {
        # Add the new user into the same groups as the template user
        $GroupCount = 0
        foreach ($Group in $UserGroups) {
            try {
                Add-ADGroupMember -Identity $Group -Members $UserAlias -ErrorAction Stop
                $GroupCount++
            }
            catch {
                Write-Warning "    Could not add to group: $Group"
            }
        }
        Write-Host "    ✓ Added to $GroupCount group(s)" -ForegroundColor Green
    }
    else {
        Write-Host "    ⓘ Template user has no group memberships" -ForegroundColor Gray
    }
    
    # Set manager if provided
    if (-not [string]::IsNullOrWhiteSpace($ManagerName)) {
        Write-Host "  → Setting manager..."
        
        try {
            # Try to resolve manager by display name or UPN
            $Manager = $null
            
            # First try as UPN
            if ($ManagerName -match '@') {
                $Manager = Get-ADUser -Filter "UserPrincipalName -eq '$ManagerName'" -ErrorAction SilentlyContinue
            }
            
            # If not found, try as display name
            if (-not $Manager) {
                $Manager = Get-ADUser -Filter "DisplayName -eq '$ManagerName'" -ErrorAction SilentlyContinue
            }
            
            # If still not found, try as SAMAccountName
            if (-not $Manager) {
                $Manager = Get-ADUser -Identity $ManagerName -ErrorAction SilentlyContinue
            }
            
            if ($Manager) {
                Set-ADUser -Identity $UserAlias -Manager $Manager.DistinguishedName -ErrorAction Stop
                Write-Host "    ✓ Manager set to: $($Manager.DisplayName)" -ForegroundColor Green
            }
            else {
                Write-Warning "    Could not find manager: $ManagerName"
            }
        }
        catch {
            Write-Warning "    Failed to set manager: $_"
        }
    }
    
    # Display final group membership
    $FinalGroups = (Get-ADUser -Identity $UserAlias -Properties MemberOf).MemberOf
    Write-Host "    ⓘ $UPN is now a member of $($FinalGroups.Count) group(s)" -ForegroundColor Gray
    
}
catch {
    Write-Error "Failed to create user: $_"
    throw
}

# Note: After account provisioning, it can take several minutes for mailbox and other services to become available