<#
.SYNOPSIS
    Restores soft-deleted user accounts in the Microsoft 365 Source Tenant.

.DESCRIPTION
    Restores soft-deleted users from the Entra ID recycle bin. Matching
    against UserMapping.csv uses exact UPN match (case-insensitive), never
    substring match. -RestoreAllFromMapping lists every user it will restore
    and requires typing YES to confirm. Supports -WhatIf / -Confirm via
    ShouldProcess.

    Requires (minimum scopes): Microsoft Graph delegated scopes
    User.ReadWrite.All and Directory.ReadWrite.All (Connect-MgGraph,
    Get-MgDirectoryDeletedItemAsUser, Restore-MgDirectoryDeletedItem).
.EXAMPLE
    .\Restore-DeletedUsers.ps1 -UserFilter "tomytra"
.EXAMPLE
    .\Restore-DeletedUsers.ps1 -RestoreAllFromMapping
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [string]$UserFilter,
    [switch]$RestoreAllFromMapping,
    [string]$MappingPath = '.\Output\UserMapping.csv',
    [string]$ConfigPath = '.\MigrationConfig.psd1'
)

. "$PSScriptRoot\MigrationHelpers.ps1"

$ErrorActionPreference = "Stop"
$env:AZURE_IDENTITY_DISABLE_WAM = "true"

if (-not (Test-Path $ConfigPath)) {
    Write-Host "Config not found at $ConfigPath" -ForegroundColor Red
    exit 1
}
$config = Import-PowerShellDataFile -Path $ConfigPath
$sourceTenantDomain = Get-RequiredConfigValue -Value $config.SourceTenantDomain -Name 'SourceTenantDomain'

Write-Host "[STEP 1] Connecting to Microsoft Graph (Directory.ReadWrite.All)..." -ForegroundColor Yellow
Connect-MgGraph -Scopes @('User.ReadWrite.All', 'Directory.ReadWrite.All') -TenantId $sourceTenantDomain -NoWelcome -ErrorAction SilentlyContinue
if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes @('User.ReadWrite.All', 'Directory.ReadWrite.All') -TenantId $sourceTenantDomain -UseDeviceCode -NoWelcome
}
Assert-Tenant -ExpectedTenantDomain $sourceTenantDomain
Write-Host "Connected as $((Get-MgContext).Account)`n" -ForegroundColor Green

Write-Host "[STEP 2] Searching for soft-deleted users in Azure AD / Entra ID..." -ForegroundColor Cyan
$deletedUsers = @(Get-MgDirectoryDeletedItemAsUser -All -ErrorAction Stop)

if ($deletedUsers.Count -eq 0) {
    Write-Host "No soft-deleted users found in directory recycle bin!" -ForegroundColor Yellow
    return
}

Write-Host "Found $($deletedUsers.Count) soft-deleted user(s) in recycle bin." -ForegroundColor DarkGray

$toRestore = @()
if ($RestoreAllFromMapping -and (Test-Path $MappingPath)) {
    $mapping = Import-Csv -Path $MappingPath -Encoding UTF8
    # Exact UPN match only (case-insensitive): a deleted user's UPN must equal
    # a SourceUPN from the mapping exactly, never merely contain it.
    $upnSet = @{}
    foreach ($u in @($mapping.SourceUPN)) {
        if (-not [string]::IsNullOrWhiteSpace($u)) {
            $upnSet[$u.Trim().ToLowerInvariant()] = $true
        }
    }
    $toRestore = @($deletedUsers | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.UserPrincipalName) -and
        $upnSet.ContainsKey($_.UserPrincipalName.Trim().ToLowerInvariant())
    })
} elseif ($UserFilter) {
    $toRestore = $deletedUsers | Where-Object { $_.UserPrincipalName -match $UserFilter -or $_.DisplayName -match $UserFilter }
} else {
    # If no filter specified, show all deleted users
    Write-Host "`nAll Soft-Deleted Users Available for Restore:" -ForegroundColor Cyan
    $deletedUsers | Select-Object Id, UserPrincipalName, DisplayName, DeletedDateTime | Format-Table -AutoSize
    Write-Host "`nTo restore specific user(s), run: .\Restore-DeletedUsers.ps1 -UserFilter <upn_or_name>" -ForegroundColor Yellow
    Write-Host "To restore ALL deleted users matching UserMapping.csv, run: .\Restore-DeletedUsers.ps1 -RestoreAllFromMapping" -ForegroundColor Yellow
    return
}

if ($toRestore.Count -eq 0) {
    Write-Host "No soft-deleted users matched your filter or UserMapping.csv." -ForegroundColor Yellow
    return
}

if ($RestoreAllFromMapping -and -not $WhatIfPreference) {
    Write-Host "`nThe following $($toRestore.Count) soft-deleted user(s) will be RESTORED:" -ForegroundColor Yellow
    foreach ($u in $toRestore) {
        Write-Host "  - $($u.UserPrincipalName) ($($u.DisplayName))" -ForegroundColor Yellow
    }
    $confirmation = Read-Host "Type YES in uppercase to confirm restoring these $($toRestore.Count) user(s)"
    if ($confirmation -cne 'YES') {
        Write-Host "Aborted: confirmation was not 'YES'. No users were restored." -ForegroundColor Yellow
        return
    }
}

foreach ($u in $toRestore) {
    if ($PSCmdlet.ShouldProcess($u.UserPrincipalName, "Restore soft-deleted user")) {
        Write-Host " -> Restoring account: $($u.UserPrincipalName) ($($u.DisplayName))... " -NoNewline -ForegroundColor Cyan
        try {
            Restore-MgDirectoryDeletedItem -DirectoryObjectId $u.Id -ErrorAction Stop | Out-Null
            Write-Host "[RESTORED SUCCESSFULLY]" -ForegroundColor Green
        } catch {
            Write-Host "[FAILED: $_]" -ForegroundColor Red
        }
    }
}

Write-Host "`nDone! Restored accounts and OneDrives will become active again within a few minutes." -ForegroundColor Green
