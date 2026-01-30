<#
.SYNOPSIS
    Assigns M365 license to a user based on template user's license.

.DESCRIPTION
    Retrieves the license SKU from the template user and assigns it to the new user.
    Uses Microsoft Graph PowerShell for modern authentication and license management.

.PARAMETER UPN
    User Principal Name of the user to license

.PARAMETER TemplateUser
    Template user whose license SKU should be copied (AD username or UPN)

.PARAMETER UsageLocation
    Usage location for the user (default: US)

.NOTES
    Requires Microsoft Graph PowerShell module with User.ReadWrite.All scope.
    Ensure users have synced via Azure AD Connect before running.
#>

Param(
    [Parameter(Mandatory=$true)]
    [string]$UPN,
    
    [Parameter(Mandatory=$true)]
    [string]$TemplateUser,
    
    [string]$UsageLocation = "US"
)

try {
    Write-Host "  Processing license for: $UPN" -ForegroundColor Cyan
    
    # Wait briefly for user to sync to M365 if just created
    Write-Host "    → Checking if user exists in M365..."
    $MaxAttempts = 3
    $Attempt = 0
    $MgUser = $null
    
    while ($Attempt -lt $MaxAttempts -and -not $MgUser) {
        try {
            $MgUser = Get-MgUser -UserId $UPN -ErrorAction Stop
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
        return
    }
    
    # Set usage location
    Write-Host "    → Setting usage location to: $UsageLocation"
    Update-MgUser -UserId $UPN -UsageLocation $UsageLocation -ErrorAction Stop
    Write-Host "      ✓ Usage location set" -ForegroundColor Green
    
    # Resolve template user and get their license
    Write-Host "    → Resolving template user license..."
    
    try {
        # Get template user from AD to find their UPN
        $TemplateUserAD = Get-ADUser -Identity $TemplateUser -Properties UserPrincipalName -ErrorAction Stop
        $TemplateUserUPN = $TemplateUserAD.UserPrincipalName
        
        # Get template user's licenses from M365
        $TemplateUserMg = Get-MgUser -UserId $TemplateUserUPN -Property AssignedLicenses -ErrorAction Stop
        
        if ($TemplateUserMg.AssignedLicenses.Count -eq 0) {
            Write-Warning "    Template user has no licenses assigned"
            return
        }
        
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
        
        Set-MgUserLicense -UserId $UPN -AddLicenses @($AddLicenses) -RemoveLicenses @() -ErrorAction Stop
        
        Write-Host "      ✓ License assigned successfully" -ForegroundColor Green
        
    }
    catch {
        Write-Warning "    Failed to assign license: $_"
        throw
    }
    
}
catch {
    Write-Error "Failed to process license for $UPN : $_"
    throw
}