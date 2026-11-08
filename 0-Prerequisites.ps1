<#
.SYNOPSIS
    Prerequisites check for M365 Migration project.

.DESCRIPTION
    Checks if Microsoft.Graph module is installed, imports sub-modules,
    tests connectivity to source and destination tenants, and lists available
    licenses in the destination tenant.

    Requires (minimum Graph scopes): User.Read.All, User.ReadWrite.All, Directory.Read.All, Directory.ReadWrite.All, Files.Read.All, Files.ReadWrite.All, Sites.Read.All, Sites.ReadWrite.All, Organization.Read.All (delegated, interactive sign-in). The app-only token test uses the source app registration's application permissions via https://graph.microsoft.com/.default (client credentials); the token is kept in memory only and never printed or written to disk.

.PARAMETER ConfigPath
    Path to the MigrationConfig.psd1 file. Default is '.\MigrationConfig.psd1'.

.EXAMPLE
    .\0-Prerequisites.ps1 -ConfigPath '.\MigrationConfig.psd1'
#>

[CmdletBinding()]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1'
)

. "$PSScriptRoot\MigrationHelpers.ps1"

# S5: any failed check increments this counter; the script exits 1 at the end
# instead of reporting SUCCESS when a check failed.
$failedChecks = 0

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$env:AZURE_IDENTITY_DISABLE_WAM = 'true'

# Create necessary directories
$baseDir = Split-Path $script:MyInvocation.MyCommand.Path -Parent
if (-not $baseDir) { $baseDir = "." }

$LogPath = Join-Path $baseDir "Logs"
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

$logFile = Join-Path $LogPath "0-Prerequisites_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $logFile -Value $logMessage
    
    $color = switch ($Level) {
        "INFO" { "Cyan" }
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
        default { "White" }
    }
    Write-Host $logMessage -ForegroundColor $color
}

Write-Log "Starting prerequisites check..." "INFO"

# Load Configuration to get tenant info
$config = @{}
if (Test-Path $ConfigPath) {
    $config = Import-PowerShellDataFile -Path $ConfigPath
}
$SourceTenant = Get-RequiredConfigValue -Value $config.SourceTenantDomain -Name 'SourceTenantDomain'
$DestTenant = Get-RequiredConfigValue -Value $config.DestTenantDomain -Name 'DestTenantDomain'

# 0b. Test app-only Graph token (client credentials) for the SOURCE tenant.
# S5: the migration actually runs app-only, so verify the app registration works.
# The token is kept in memory only and is never printed or written to disk.
try {
    Write-Log "Testing app-only Graph token (client credentials) for SOURCE tenant..." "INFO"
    $sourceClientId = Get-RequiredConfigValue -Value $config.SourceClientId -Name 'SourceClientId'
    $sourceClientSecret = Get-RequiredConfigValue -Value $config.SourceClientSecret -Name 'SourceClientSecret'
    $appOnlyToken = Get-AppOnlyGraphAccessToken -TenantDomain $SourceTenant -ClientId $sourceClientId -ClientSecret $sourceClientSecret
    $appOnlyToken = $null
    Write-Log "App-only token acquired successfully for SOURCE tenant." "SUCCESS"
} catch {
    Write-Log "Failed to acquire app-only Graph token for SOURCE tenant: $_" "ERROR"
    $failedChecks++
}

# 1. Check if Microsoft.Graph is installed
try {
    Write-Log "Checking for Microsoft.Graph module..." "INFO"
    $module = Get-Module -ListAvailable -Name Microsoft.Graph
    if (-not $module) {
        Write-Log "Microsoft.Graph module not found. Installing..." "WARNING"
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Log "Microsoft.Graph module installed successfully." "SUCCESS"
    } else {
        Write-Log "Microsoft.Graph module is already installed." "SUCCESS"
    }
} catch {
    Write-Log "Failed to check or install Microsoft.Graph module: $_" "ERROR"
    throw
}

# 1b. Check SharePoint Online Management Shell (needed by 3a/3b for Connect-SPOService)
try {
    Write-Log "Checking for Microsoft.Online.SharePoint.PowerShell module..." "INFO"
    $spoModule = Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell
    if (-not $spoModule) {
        Write-Log "Microsoft.Online.SharePoint.PowerShell module not found. It is required by 3a/3b (Connect-SPOService). Install it with: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser" "ERROR"
        $failedChecks++
    } else {
        Write-Log "Microsoft.Online.SharePoint.PowerShell module is already installed." "SUCCESS"
    }
} catch {
    Write-Log "Failed to check Microsoft.Online.SharePoint.PowerShell module: $_" "ERROR"
    $failedChecks++
}

# 2. Import sub-modules
try {
    Write-Log "Importing required Microsoft.Graph sub-modules..." "INFO"
    $modulesToImport = @(
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Identity.DirectoryManagement",
        "Microsoft.Graph.Files",
        "Microsoft.Graph.Users.Actions"
    )
    foreach ($mod in $modulesToImport) {
        Import-Module $mod -ErrorAction Stop
        Write-Log "Imported module: $mod" "SUCCESS"
    }
} catch {
    Write-Log "Failed to import sub-modules: $_" "ERROR"
    throw
}

# 3. Test connectivity - Source Tenant
try {
    Write-Log "Testing connectivity to SOURCE tenant (Read-Only)..." "INFO"
    Write-Log "A browser window will open. Sign in with your SOURCE tenant ($SourceTenant) admin account." "WARNING"
    $sourceScopes = @("User.Read.All", "Directory.Read.All", "Files.Read.All", "Sites.Read.All")
    Connect-MgGraph -Scopes $sourceScopes -TenantId $SourceTenant -ErrorAction Stop
    $sourceContext = Get-MgContext
    Write-Log "Successfully connected to SOURCE tenant: $($sourceContext.TenantId)" "SUCCESS"
    
    Write-Log "Disconnecting from SOURCE tenant..." "INFO"
    Disconnect-MgGraph | Out-Null
    Write-Log "Disconnected from SOURCE tenant." "SUCCESS"
} catch {
    Write-Log "Failed to connect to SOURCE tenant: $_" "ERROR"
    $failedChecks++
}

# 4. Test connectivity - Destination Tenant
try {
    Write-Log "Testing connectivity to DESTINATION tenant (ReadWrite)..." "INFO"
    Write-Log "A browser window will open. Sign in with your DESTINATION tenant ($DestTenant) admin account." "WARNING"
    $destScopes = @("User.ReadWrite.All", "Directory.ReadWrite.All", "Files.ReadWrite.All", "Organization.Read.All", "Sites.ReadWrite.All")
    Connect-MgGraph -Scopes $destScopes -TenantId $DestTenant -ErrorAction Stop
    $destContext = Get-MgContext
    Write-Log "Successfully connected to DESTINATION tenant: $($destContext.TenantId)" "SUCCESS"
    
    # 5. List available licenses
    Write-Log "Retrieving available licenses (SKUs) in DESTINATION tenant..." "INFO"
    $skus = Get-MgSubscribedSku -ErrorAction Stop
    if ($skus) {
        Write-Log "Available Licenses in DESTINATION tenant:" "SUCCESS"
        foreach ($sku in $skus) {
            $enabled = $sku.PrepaidUnits.Enabled
            $consumed = $sku.ConsumedUnits
            Write-Log "  - SkuPartNumber: $($sku.SkuPartNumber) | Total Enabled: $enabled | Consumed: $consumed" "INFO"
        }
    } else {
        Write-Log "No subscribed SKUs found in DESTINATION tenant." "WARNING"
    }
    
    Write-Log "Disconnecting from DESTINATION tenant..." "INFO"
    Disconnect-MgGraph | Out-Null
    Write-Log "Disconnected from DESTINATION tenant." "SUCCESS"
} catch {
    Write-Log "Failed to connect to DESTINATION tenant or retrieve licenses: $_" "ERROR"
    $failedChecks++
}

# 6. Summary (S5: any failed check fails the whole script)
if ($failedChecks -gt 0) {
    Write-Log "Prerequisites check completed with $failedChecks failed check(s). Fix the failures above before running the migration." "ERROR"
    exit 1
}
Write-Log "Prerequisites check completed." "SUCCESS"
