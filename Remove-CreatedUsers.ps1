<#
.SYNOPSIS
Removes newly created destination users from UserMapping.csv so you can recreate them with the correct domain.

.DESCRIPTION
Destructive script: only deletes users that BOTH have Status='Created' in the
mapping AND have a UPN in the destination domain (NewDomain from config).
Before deleting each user it checks that the user has no assigned licenses
and no OneDrive data left; users that still have a license or OneDrive data
are skipped with a warning, not deleted. Deletion requires typing YES at the
confirmation prompt, and supports -WhatIf / -Confirm via ShouldProcess.

Requires (minimum scopes): Microsoft Graph delegated scopes User.ReadWrite.All
and Directory.ReadWrite.All (Remove-MgUser), User.Read.All
(Get-MgUserLicenseDetail license check), Files.Read.All (OneDrive data check)
and Organization.Read.All (tenant check).
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param (
    [string]$ConfigPath = ".\MigrationConfig.psd1",
    [string]$MappingPath = ".\Output\UserMapping.csv"
)

. "$PSScriptRoot\MigrationHelpers.ps1"

$ErrorActionPreference = "Stop"

if (-not (Test-Path $ConfigPath)) {
    Write-Host "Config not found at $ConfigPath" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $MappingPath)) {
    Write-Host "UserMapping.csv not found at $MappingPath" -ForegroundColor Red
    exit 1
}

$config = Import-PowerShellDataFile -Path $ConfigPath
$destTenantDomain = Get-RequiredConfigValue -Value $config.DestTenantDomain -Name 'DestTenantDomain'
$newDomain = Get-RequiredConfigValue -Value $config.NewDomain -Name 'NewDomain'

$users = @(Import-Csv -Path $MappingPath | Where-Object {
    $_.Status -eq 'Created' -and -not [string]::IsNullOrWhiteSpace($_.DestUPN) -and
    (($_.DestUPN -split '@')[-1] -ieq $newDomain)
})

if ($users.Count -eq 0) {
    Write-Host "No users with Status='Created' and a UPN in the destination domain '$newDomain' found in $MappingPath." -ForegroundColor Yellow
    exit 0
}

Write-Host "The following $($users.Count) user(s) will be DELETED from the destination tenant '$destTenantDomain':" -ForegroundColor Red
foreach ($u in $users) {
    Write-Host "  - $($u.DestUPN)" -ForegroundColor Red
}
if (-not $WhatIfPreference) {
    $confirmation = Read-Host "Type YES in uppercase to confirm deletion of these $($users.Count) user(s)"
    if ($confirmation -cne 'YES') {
        Write-Host "Aborted: confirmation was not 'YES'. No users were deleted." -ForegroundColor Yellow
        exit 0
    }
}

Write-Host "Connecting to Destination Tenant ($destTenantDomain)..." -ForegroundColor Cyan
Disconnect-MgGraph -ErrorAction SilentlyContinue
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All", "User.Read.All", "Files.Read.All", "Organization.Read.All" -TenantId $destTenantDomain -UseDeviceCode -NoWelcome *>&1 | Write-Host
Assert-Tenant -ExpectedTenantDomain $destTenantDomain

foreach ($u in $users) {
    $upn = $u.DestUPN

    # Pre-check 1: user must have no assigned licenses.
    try {
        $licenses = @(Get-MgUserLicenseDetail -UserId $upn -ErrorAction Stop)
    } catch {
        Write-Host " -> Skipping ${upn}: could not check licenses ($_). Not deleting." -ForegroundColor Yellow
        continue
    }
    if ($licenses.Count -gt 0) {
        Write-Host " -> Skipping ${upn}: user still has $($licenses.Count) assigned license(s). Remove licenses first; not deleting." -ForegroundColor Yellow
        continue
    }

    # Pre-check 2: user must have no OneDrive data left.
    try {
        $drive = Get-MgUserDrive -UserId $upn -ErrorAction Stop
        if ($null -ne $drive) {
            $encodedUpn = [System.Uri]::EscapeDataString($upn)
            $children = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$encodedUpn/drive/root/children?`$top=1" -ErrorAction Stop
            if ($null -ne $children.value -and @($children.value).Count -gt 0) {
                Write-Host " -> Skipping ${upn}: OneDrive still contains data. Clear the OneDrive first; not deleting." -ForegroundColor Yellow
                continue
            }
        }
    } catch {
        # No drive found (404) means no OneDrive data; any other error is treated as "cannot verify", so skip.
        if ($_.Exception.Message -notmatch '404|NotFound|ResourceNotFound|itemNotFound') {
            Write-Host " -> Skipping ${upn}: could not check OneDrive data ($_). Not deleting." -ForegroundColor Yellow
            continue
        }
    }

    if ($PSCmdlet.ShouldProcess($upn, "Remove destination user")) {
        Write-Host "Removing user $upn..." -ForegroundColor Yellow
        try {
            Remove-MgUser -UserId $upn -ErrorAction Stop
            Write-Host " -> Deleted $upn successfully." -ForegroundColor Green
        } catch {
            Write-Host " -> Failed to remove ${upn}: $_" -ForegroundColor Red
        }
    }
}

Write-Host "`nCleanup completed! Update NewDomain in MigrationConfig.psd1 to your custom domain and re-run Phase 2." -ForegroundColor Green
