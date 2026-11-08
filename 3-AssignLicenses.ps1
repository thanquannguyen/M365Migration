<#
.SYNOPSIS
    Assigns licenses to users in the destination tenant.

.DESCRIPTION
    This script connects to the destination tenant and assigns a specific license
    (e.g., Office 365 A1 for Students) to the users based on the UserMapping.csv file.
    It ensures that a UsageLocation is set for each user before assigning the license,
    as required by Microsoft 365.

    Requires (minimum Graph scopes): User.ReadWrite.All, Organization.Read.All (delegated, interactive sign-in).

.PARAMETER ConfigPath
    Path to the migration configuration file. Defaults to '.\MigrationConfig.psd1'.

.PARAMETER MappingPath
    Path to the user mapping CSV file. Defaults to '.\Output\UserMapping.csv'.

.PARAMETER LicenseSku
    Optional. The SkuPartNumber of the license to assign. If not provided, it
    must be set in the config file (LicenseSku); there is no hard-coded fallback.

.EXAMPLE
    .\3-AssignLicenses.ps1 -ConfigPath '.\MigrationConfig.psd1' -MappingPath '.\Output\UserMapping.csv' -LicenseSku 'O365_BUSINESS_ESSENTIALS'
#>

[CmdletBinding()]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$MappingPath = '.\Output\UserMapping.csv',
    [string]$LicenseSku
)

. "$PSScriptRoot\MigrationHelpers.ps1"

# S7: fail fast on any error instead of continuing after a failed call.
$ErrorActionPreference = 'Stop'

# Ensure Logs directory exists
$InformationPreference = 'Continue'
$env:AZURE_IDENTITY_DISABLE_WAM = 'true'
$logDir = ".\Logs"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}
$logFile = Join-Path $logDir "AssignLicenses_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor White }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
    }
    
    Add-Content -Path $logFile -Value $logMessage -Encoding UTF8
}

Write-Log "Starting License Assignment process."

# 1. Load config
if (Test-Path $ConfigPath) {
    Write-Log "Loading configuration from $ConfigPath"
    $config = Import-PowerShellDataFile -Path $ConfigPath
} else {
    Write-Log "Configuration file not found at $ConfigPath" -Level "WARNING"
    $config = @{}
}

if (-not (Test-Path $MappingPath)) {
    Write-Log "Mapping file not found at $MappingPath. Please run the user creation step first." -Level "ERROR"
    exit 1
}

# Determine default UsageLocation from config or fallback
$defaultUsageLocation = if ($config.DefaultUsageLocation) { $config.DefaultUsageLocation } else { 'VN' }

# Determine target License SKU (S7: required from config, no hard-coded fallback)
$targetSkuPartNumber = $LicenseSku
if ([string]::IsNullOrEmpty($targetSkuPartNumber)) {
    $targetSkuPartNumber = Get-RequiredConfigValue -Value $config.LicenseSku -Name 'LicenseSku'
}

Write-Log "Connecting to destination tenant..."
Write-Log "A browser window will open. Sign in with your DESTINATION tenant admin account." -Level "WARNING"
try {
    $connectParams = @{ Scopes = @('User.ReadWrite.All', 'Organization.Read.All') }
    $destTenantDomain = Get-RequiredConfigValue -Value $config.DestTenantDomain -Name 'DestTenantDomain'
    $connectParams.TenantId = $destTenantDomain
    Connect-MgGraph @connectParams -ErrorAction Stop
    Assert-Tenant -ExpectedTenantDomain $destTenantDomain
    Write-Log "Connected successfully. Tenant verified against DestTenantDomain." -Level "SUCCESS"
} catch {
    Write-Log "Failed to connect to destination tenant: $_" -Level "ERROR"
    exit 1
}

# 3. Get all available SKUs
Write-Log "Retrieving subscribed SKUs..."
try {
    $skus = Get-MgSubscribedSku -All
    
    Write-Host "`nAvailable Licenses:" -ForegroundColor Cyan
    $skuTable = $skus | Select-Object SkuPartNumber, SkuId, ConsumedUnits, @{Name="PrepaidUnits";Expression={$_.PrepaidUnits.Enabled}}
    $skuTable | Format-Table -AutoSize
    
    # 4. Find the SKU matching target
    $targetSku = $skus | Where-Object { $_.SkuPartNumber -eq $targetSkuPartNumber }
    
    if (-not $targetSku) {
        Write-Log "Could not find a matching SKU for '$targetSkuPartNumber'." -Level "ERROR"
        Write-Log "Please check the available licenses above and provide a valid -LicenseSku parameter." -Level "ERROR"
        Disconnect-MgGraph
        exit 1
    }
    
    $skuId = $targetSku.SkuId
    Write-Log "Selected SKU: $($targetSku.SkuPartNumber) (${skuId})" -Level "INFO"
    
    if ($targetSku.PrepaidUnits.Enabled -le $targetSku.ConsumedUnits) {
        # S7: stop instead of warning and assigning anyway when no units are left.
        Write-Log "No available units left for SKU $($targetSku.SkuPartNumber). Stopping instead of assigning a license with no units left." -Level "ERROR"
        Disconnect-MgGraph
        exit 1
    }
} catch {
    Write-Log "Failed to retrieve subscribed SKUs: $_" -Level "ERROR"
    Disconnect-MgGraph
    exit 1
}

# 5. Read UserMapping.csv
Write-Log "Reading user mapping file from $MappingPath..."
$userMapping = Import-Csv -Path $MappingPath
# S7: guard against an empty/null mapping before touching $userMapping[0] below.
if (-not $userMapping -or $userMapping.Count -eq 0) {
    Write-Log "User mapping file '$MappingPath' is empty or contains no users. Nothing to license." -Level "ERROR"
    Disconnect-MgGraph
    exit 1
}
$eligibleUsers = $userMapping | Where-Object { $_.Status -in @('Created', 'Skipped') }

Write-Log "Found $($eligibleUsers.Count) users eligible for license assignment." -Level "INFO"

$assignedCount = 0
$alreadyLicensedCount = 0
$failedCount = 0

# Add LicenseStatus column if it doesn't exist
if (-not $userMapping[0].psobject.properties.match('LicenseStatus').Count) {
    $userMapping | Add-Member -MemberType NoteProperty -Name 'LicenseStatus' -Value ''
}

# 6. For each user
foreach ($userRow in $eligibleUsers) {
    $upn = $userRow.DestUPN
    $userId = $userRow.DestId
    
    if ([string]::IsNullOrEmpty($userId)) {
        $userId = $upn # Fallback to UPN if ID is empty
    }
    
    Write-Log "Processing user: $upn"
    
    try {
        # Get user details to check UsageLocation and current licenses
        $mgUser = Get-MgUser -UserId $userId -Property Id, UserPrincipalName, UsageLocation, AssignedLicenses -ErrorAction Stop
        
        # 6a. Ensure UsageLocation is set
        if ([string]::IsNullOrEmpty($mgUser.UsageLocation)) {
            Write-Log "  UsageLocation not set for $upn. Setting to '$defaultUsageLocation'..."
            Update-MgUser -UserId $userId -UsageLocation $defaultUsageLocation -ErrorAction Stop
        }
        
        # 6c. Handle already-licensed users
        $hasLicense = $false
        foreach ($assigned in $mgUser.AssignedLicenses) {
            if ($assigned.SkuId -eq $skuId) {
                $hasLicense = $true
                break
            }
        }
        
        if ($hasLicense) {
            Write-Log "  User $upn already has the target license." -Level "WARNING"
            $userRow.LicenseStatus = 'AlreadyLicensed'
            $alreadyLicensedCount++
        } else {
            # 6b. Assign license
            Write-Log "  Assigning license to $upn..."
            
            $addLicenses = @(
                @{
                    SkuId = $skuId
                }
            )
            
            Set-MgUserLicense -UserId $userId -AddLicenses $addLicenses -RemoveLicenses @() -ErrorAction Stop
            
            Write-Log "  Successfully assigned license to $upn." -Level "SUCCESS"
            $userRow.LicenseStatus = 'Assigned'
            $assignedCount++
        }
    } catch {
        Write-Log "  Failed to process license for ${upn}: $_" -Level "ERROR"
        $userRow.LicenseStatus = 'Failed'
        $failedCount++
    }
}

# 7. Update UserMapping.csv
Write-Log "Updating mapping file $MappingPath..."
$userMapping | Export-Csv -Path $MappingPath -NoTypeInformation -Encoding UTF8
Write-Log "Updated mapping file successfully." -Level "SUCCESS"

# 8. Display summary
Write-Host "`n--- License Assignment Summary ---" -ForegroundColor Cyan
Write-Host "Total processed  : $($eligibleUsers.Count)"
Write-Host "Assigned         : $assignedCount" -ForegroundColor Green
Write-Host "Already Licensed : $alreadyLicensedCount" -ForegroundColor Yellow
Write-Host "Failed           : $failedCount" -ForegroundColor Red
Write-Host "----------------------------------" -ForegroundColor Cyan

Write-Log "Summary: Total: $($eligibleUsers.Count), Assigned: $assignedCount, Already Licensed: $alreadyLicensedCount, Failed: $failedCount"

# 10. Disconnect
Write-Log "Disconnecting from Microsoft Graph..."
Disconnect-MgGraph

Write-Log "License assignment process completed."
