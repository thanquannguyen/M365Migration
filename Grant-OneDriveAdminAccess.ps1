<#
.SYNOPSIS
Grants an admin account Site Collection Administrator access to all SOURCE tenant
OneDrive sites in UserMapping.csv.

.DESCRIPTION
SOURCE tenant only, via the SharePoint Online Management Shell module
(Connect-SPOService / Set-SPOUser). This is NOT the destination-tenant script:
3b-GrantDestOneDriveAdminAccess.ps1 grants the same kind of access on the
destination tenant's OneDrive sites. Do not confuse the two -- running this
script against the wrong tenant grants admin access on the wrong tenant's sites.

By default this script does NOT change any storage quota. Changing quotas is
opt-in only via -SetStorageQuota (see below).

Requires (minimum scopes): SharePoint Online Management Shell module
(Connect-SPOService, Set-SPOUser, and Set-SPOSite only when -SetStorageQuota
is used) with a source-tenant SharePoint admin account; no Microsoft Graph
scopes are used.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Your Admin UPN in the Source Tenant (e.g. admin@source-domain.com)")]
    [string]$AdminUPN,

    [string]$ConfigPath = ".\MigrationConfig.psd1",
    [string]$MappingPath = ".\Output\UserMapping.csv",

    [switch]$SetStorageQuota
)

. "$PSScriptRoot\MigrationHelpers.ps1"

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

if ($SetStorageQuota) {
    Write-Host "WARNING: -SetStorageQuota was passed. This script will change the storage quota of every source OneDrive site it touches to 102400 MB (warning level 92160 MB) and reset the quota. This changes tenant storage policy -- only continue if you really intend to change quotas." -ForegroundColor Red
}

# Load Config
if (-not (Test-Path $ConfigPath)) {
    Write-Host "Config file not found at $ConfigPath" -ForegroundColor Red
    exit 1
}
$config = Import-PowerShellDataFile -Path $ConfigPath
$SourceTenantDomain = Get-RequiredConfigValue -Value $config.SourceTenantDomain -Name 'SourceTenantDomain'
$TenantName = $SourceTenantDomain.Split('.')[0] # e.g. contoso
$AdminUrl = Get-SpoAdminUrl -TenantDomain $SourceTenantDomain
$MySiteHost = "https://${TenantName}-my.sharepoint.com"

# Check for SharePoint module
if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
    Write-Host "Installing Microsoft.Online.SharePoint.PowerShell module..." -ForegroundColor Yellow
    Install-Module Microsoft.Online.SharePoint.PowerShell -Force -Scope CurrentUser -AllowClobber
}

Import-Module Microsoft.Online.SharePoint.PowerShell

Write-Host "Connecting to SharePoint Admin Center: $AdminUrl" -ForegroundColor Cyan
Connect-SPOService -Url $AdminUrl
Assert-SpoTenant -ExpectedTenantDomain $SourceTenantDomain -AdminUrl $AdminUrl

if (-not (Test-Path $MappingPath)) {
    Write-Host "UserMapping.csv not found at $MappingPath" -ForegroundColor Red
    exit 1
}

$users = Import-Csv -Path $MappingPath
foreach ($u in $users) {
    $upn = $u.SourceUPN
    $personalUrlPart = $upn.Replace('@', '_').Replace('.', '_')
    $siteUrl = "${MySiteHost}/personal/${personalUrlPart}"
    
    Write-Host "Unlocking $siteUrl..." -ForegroundColor Yellow
    try {
        Set-SPOSite -Identity $siteUrl -LockState Unlock -ErrorAction SilentlyContinue
    } catch {
        Write-Host " -> Notice while unlocking: $_" -ForegroundColor DarkYellow
    }

    if ($SetStorageQuota) {
        Write-Host "Changing storage quota to 102400 MB (warning level 92160 MB) on $siteUrl..." -ForegroundColor Red
        try {
            Set-SPOSite -Identity $siteUrl -StorageQuota 102400 -StorageQuotaWarningLevel 92160 -ErrorAction SilentlyContinue
            Set-SPOSite -Identity $siteUrl -StorageQuotaReset -ErrorAction SilentlyContinue
        } catch {
            Write-Host " -> Notice while changing quota: $_" -ForegroundColor DarkYellow
        }
    }

    Write-Host "Granting Site Collection Admin on $siteUrl to $AdminUPN..." -ForegroundColor Yellow
    try {
        Set-SPOUser -Site $siteUrl -LoginName $AdminUPN -IsSiteCollectionAdmin $true -ErrorAction Stop
        Write-Host " -> Successfully granted admin access on $upn's OneDrive." -ForegroundColor Green
    } catch {
        Write-Host " -> Failed for ${upn}: $_" -ForegroundColor Red
    }
}

Write-Host "`nAll done! You can now run .\Run-Migration.ps1 -StartFromPhase 4" -ForegroundColor Green
