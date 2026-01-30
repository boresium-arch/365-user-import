<#
.SYNOPSIS
    Copies group membership from a template user to a new user.

.DESCRIPTION
    Retrieves all group memberships from a template user and adds the new user
    to the same groups in Active Directory.

.PARAMETER UserAlias
    The SAM account name or identity of the new user

.PARAMETER TemplateUser
    The SAM account name or identity of the template user to copy groups from

.EXAMPLE
    .\Set-Groups.ps1 -UserAlias "jdoe" -TemplateUser "jsmith"
#>

Param(
    [Parameter(Mandatory=$true)]
    [string]$UserAlias,
    
    [Parameter(Mandatory=$true)]
    [string]$TemplateUser
)

try {
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
    
    # Display final group membership
    $FinalGroups = (Get-ADUser -Identity $UserAlias -Properties MemberOf).MemberOf
    Write-Host "    ⓘ $UserAlias is now a member of $($FinalGroups.Count) group(s)" -ForegroundColor Gray
}
catch {
    Write-Error "Failed to copy groups for $UserAlias : $_"
    throw
}
