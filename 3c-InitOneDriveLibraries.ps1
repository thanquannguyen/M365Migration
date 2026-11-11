<#
.SYNOPSIS
Instantly initializes the default OneDrive 'Documents' library on all newly created SharePoint Personal Sites via Microsoft Graph API.

.DESCRIPTION
When Request-SPOPersonalSite creates a user's MySite, the default Documents library (OneDrive drive) is not created until first access or timer job completion.
This script uses Microsoft Graph API (POST /sites/{site-id}/lists with template='documentLibrary') to synchronously create the Documents library on every user's personal site.
Once created, 4b-UploadWorker.ps1 can immediately upload files to /users/{upn}/drive/root without waiting!

Requires (minimum scopes): Microsoft Graph delegated scopes Sites.Manage.All,
Files.ReadWrite.All and Organization.Read.All (Connect-MgGraph,
Invoke-MgGraphRequest). Organization.Read.All is needed by the tenant check.
#>

[CmdletBinding()]
param (
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

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   OneDrive Documents Library Initializer (Graph API)     " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Connecting to Microsoft Graph ($tenantDomain)..." -ForegroundColor Cyan
Write-Host "A browser window or Device Code prompt will appear. Sign in with your DESTINATION tenant admin account." -ForegroundColor Yellow

try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Connect-MgGraph -Scopes "Sites.Manage.All","Files.ReadWrite.All","Organization.Read.All" -TenantId $tenantDomain -UseDeviceCode -NoWelcome -ErrorAction Stop
    Assert-Tenant -ExpectedTenantDomain $tenantDomain
    Write-Host "Connected to Microsoft Graph successfully!" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit
}

Write-Host "`nInitializing Documents library on $($emails.Count) user OneDrive sites..." -ForegroundColor Cyan

foreach ($email in $emails) {
    $siteName = $email -replace '[@.]', '_'
    $sitePath = "https://graph.microsoft.com/v1.0/sites/$tenantName-my.sharepoint.com:/personal/$siteName"
    
    Write-Host " -> $email ..." -NoNewline
    try {
        $site = Invoke-MgGraphRequest -Method GET -Uri $sitePath -ErrorAction Stop
        
        # Check if drive already exists
        $drives = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives" -ErrorAction Stop
        if ($drives.value.Count -gt 0) {
            Write-Host " [READY: $($drives.value[0].name) exists]" -ForegroundColor Green
            continue
        }
        
        # Create default Documents library
        $body = @{
            displayName = "Documents"
            list = @{
                template = "documentLibrary"
            }
        } | ConvertTo-Json -Depth 5
        
        $newLib = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists" -Body $body -ContentType "application/json" -ErrorAction Stop
        Write-Host " [INITIALIZED LIBRARY: $($newLib.displayName)]" -ForegroundColor Green
    } catch {
        Write-Host " [WARN: $_]" -ForegroundColor Yellow
    }
}

Write-Host "`nAll OneDrive Documents libraries initialized!" -ForegroundColor Cyan
Write-Host "You can now start .\4b-UploadWorker.ps1 immediately!" -ForegroundColor Green
