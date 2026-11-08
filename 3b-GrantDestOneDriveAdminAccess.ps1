<#
.SYNOPSIS
Grants Site Collection Admin access on all destination OneDrive sites to the specified admin user.

.DESCRIPTION
By default, Global Admins in Microsoft 365 do not have permission to access users' OneDrive sites.
This script grants Site Collection Admin rights on the destination OneDrive sites listed in
UserMapping.csv so that:
1. You can visit them in the browser without getting Access Denied.
2. 4b-UploadWorker.ps1 can read and upload files without 403 Forbidden / Not Found errors.

Requires (minimum scopes): SharePoint Online Management Shell module
(Connect-SPOService, Set-SPOUser) with a destination-tenant SharePoint admin account;
no Microsoft Graph scopes are used.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Admin UPN in the destination tenant that receives Site Collection Admin access (e.g. admin@dest-domain.com)")]
    [string]$AdminUPN,

    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$MappingPath = '.\Output\UserMapping.csv'
)

. "$PSScriptRoot\MigrationHelpers.ps1"

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$env:AZURE_IDENTITY_DISABLE_WAM = 'true'
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location -Path $ScriptDir

if (-not (Test-Path $MappingPath)) {
    Write-Error "Mapping file not found at $MappingPath"
    exit
}
$Mapping = Import-Csv -Path $MappingPath -Encoding UTF8
$emails = @($Mapping.DestUPN | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

$config = @{}
if (Test-Path $ConfigPath) {
    $config = Import-PowerShellDataFile -Path $ConfigPath
}

$tenantDomain = Get-RequiredConfigValue -Value $config.DestTenantDomain -Name 'DestTenantDomain'
$tenantName = $tenantDomain.Split('.')[0]
$adminUrl = Get-SpoAdminUrl -TenantDomain $tenantDomain

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   Destination OneDrive Admin Access Grant Tool           " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Connecting to SharePoint Admin Center ($adminUrl)..." -ForegroundColor Cyan
Write-Host "Please sign in with your DESTINATION tenant admin account ($AdminUPN)." -ForegroundColor Yellow

try {
    Connect-SPOService -Url $adminUrl -ErrorAction Stop
    Assert-SpoTenant -ExpectedTenantDomain $tenantDomain -AdminUrl $adminUrl
    Write-Host "Connected to SharePoint Admin Center successfully!" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to SharePoint Admin Center: $_"
    exit
}

Write-Host "`nGranting Site Collection Admin access to ($AdminUPN) on $($emails.Count) OneDrive sites..." -ForegroundColor Cyan

foreach ($email in $emails) {
    $siteName = $email -replace '[@.]', '_'
    $mySiteUrl = "https://$tenantName-my.sharepoint.com/personal/$siteName"
    
    Write-Host " -> $email ($mySiteUrl) ..." -NoNewline
    try {
        Set-SPOUser -Site $mySiteUrl -LoginName $AdminUPN -IsSiteCollectionAdmin $true -ErrorAction Stop
        Write-Host " [ACCESS GRANTED]" -ForegroundColor Green
    } catch {
        Write-Host " [WARN: $_]" -ForegroundColor Yellow
    }
}

Write-Host "`nAll OneDrive sites processed! You now have full admin permissions." -ForegroundColor Cyan
Write-Host "You can now run .\4b-UploadWorker.ps1 to start parallel uploading!" -ForegroundColor Green
