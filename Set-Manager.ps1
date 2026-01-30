<#
.SYNOPSIS
    Resolves and sets the manager for a user.

.DESCRIPTION
    Resolves a manager by UPN, DisplayName, or SAMAccountName and assigns them
    to the specified user in Active Directory.

.PARAMETER UserAlias
    The SAM account name or identity of the user to set manager for

.PARAMETER ManagerName
    Display name, UPN, or SAM account name of the manager

.EXAMPLE
    .\Set-Manager.ps1 -UserAlias "jdoe" -ManagerName "jane.smith@contoso.com"
#>

Param(
    [Parameter(Mandatory=$true)]
    [string]$UserAlias,
    
    [Parameter(Mandatory=$true)]
    [string]$ManagerName
)

try {
    Write-Host "  → Setting manager..."
    
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
    throw
}
