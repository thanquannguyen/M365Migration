<#
.SYNOPSIS
Pre-provisions OneDrive for Business sites for all migrated users in the destination tenant.

.DESCRIPTION
Microsoft 365 does not automatically create OneDrive sites when users are created.
This script uses SharePoint Online Management Shell (Request-SPOPersonalSite) to trigger
OneDrive provisioning for all target users in UserMapping.csv so Phase 4 can upload immediately.

Requires (minimum scopes): SharePoint Online Management Shell module
(Connect-SPOService, Request-SPOPersonalSite, Get-SPOSite) with a destination-tenant
SharePoint admin account; no Microsoft Graph scopes are used.
#>

[CmdletBinding()]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$MappingPath = '.\Output\UserMapping.csv',
    [ValidateRange(1, 480)]
    [int]$ProvisionTimeoutMinutes = 60,
    [ValidateRange(10, 600)]
    [int]$PollIntervalSeconds = 60
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
Write-Host "   OneDrive Site Pre-Provisioning Tool                    " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Connecting to SharePoint Admin Center ($adminUrl)..." -ForegroundColor Cyan
Write-Host "Please sign in with your DESTINATION tenant admin account." -ForegroundColor Yellow

try {
    Connect-SPOService -Url $adminUrl -ErrorAction Stop
    Assert-SpoTenant -ExpectedTenantDomain $tenantDomain -AdminUrl $adminUrl
    Write-Host "Connected to SharePoint Admin Center successfully!" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to SharePoint Admin Center: $_"
    exit
}

Write-Host "`nRequesting OneDrive site provisioning for $($emails.Count) users one by one..." -ForegroundColor Cyan
foreach ($email in $emails) {
    Write-Host "   -> Requesting site for $email ..." -NoNewline
    try {
        Request-SPOPersonalSite -UserEmails @($email) -NoWait -ErrorAction Stop
        Write-Host " [QUEUED]" -ForegroundColor Green
    } catch {
        Write-Host " [WARN: $_]" -ForegroundColor Yellow
    }
}

Write-Host "`nPolling provisioning status (timeout: $ProvisionTimeoutMinutes minute(s), checking every $PollIntervalSeconds second(s))..." -ForegroundColor Cyan
Write-Host "Provisioning OneDrive can take minutes to hours; this script waits until every site is ready or the timeout expires." -ForegroundColor Yellow
$deadline = (Get-Date).AddMinutes($ProvisionTimeoutMinutes)
$pendingEmails = @($emails)
while ($pendingEmails.Count -gt 0) {
    $stillPending = @()
    foreach ($email in $pendingEmails) {
        $siteName = $email -replace '[@.]', '_'
        $mySiteUrl = "https://$tenantName-my.sharepoint.com/personal/$siteName"

        try {
            $site = Get-SPOSite -Identity $mySiteUrl -ErrorAction Stop
            if ($site) {
                Write-Host " [READY] $email -> $mySiteUrl" -ForegroundColor Green
            } else {
                $stillPending += $email
            }
        } catch {
            $stillPending += $email
            Write-Host " [PROVISIONING] $email is queued in Microsoft 365 background." -ForegroundColor Yellow
        }
    }
    $pendingEmails = $stillPending
    if ($pendingEmails.Count -eq 0) {
        break
    }
    if ((Get-Date) -ge $deadline) {
        break
    }
    Write-Host "Waiting for $($pendingEmails.Count) site(s) to finish provisioning; next check in $PollIntervalSeconds second(s)..." -ForegroundColor Yellow
    Start-Sleep -Seconds $PollIntervalSeconds
}

if ($pendingEmails.Count -gt 0) {
    Write-Warning "Provisioning timed out after $ProvisionTimeoutMinutes minute(s). The following site(s) are NOT yet provisioned:"
    foreach ($email in $pendingEmails) {
        Write-Warning "  - $email"
    }
    Write-Warning "Do NOT start the parallel upload worker for these users yet. Re-run this script later, or wait for the sites to finish provisioning before uploading to them."
    exit 1
}

Write-Host "`nAll OneDrive sites provisioned! You can now start the parallel upload worker." -ForegroundColor Cyan
