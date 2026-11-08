<#
.SYNOPSIS
Permanently empties OneDrive Recycle Bins for source tenant accounts.

Requires (minimum Graph scopes): User.Read.All, Files.ReadWrite.All,
Sites.ReadWrite.All, Directory.Read.All, Organization.Read.All.

Runs only after a recent passing 5-Validate.ps1 run (validation gate, B2)
and only against the configured source tenant (B1).

.DESCRIPTION
When files are deleted from OneDrive, they are moved to the Recycle Bin.
Because SharePoint/OneDrive includes Recycle Bin data in the account's storage quota,
accounts that deleted large amounts of data may still show "507 Insufficient Storage"
(Remaining: 0 B) until the Recycle Bin is permanently emptied.

This script queries each user's SharePoint site recycle bin via Microsoft Graph API
and permanently purges deleted items in high-speed batches (500 items per request).

Modes:
- What-If (Default): Scans each account's quota and checks how many items/bytes are in the Recycle Bin.
- Execute (-Confirm): Permanently purges items from the Recycle Bin.

Examples:
  # Check recycle bin usage across all accounts (preview only)
  .\5d-EmptyRecycleBin.ps1

  # Check deelf's recycle bin
  .\5d-EmptyRecycleBin.ps1 -UserFilter "deelf"

  # Permanently empty deelf's recycle bin
  .\5d-EmptyRecycleBin.ps1 -UserFilter "deelf" -Confirm

  # Permanently empty recycle bins for all accounts in UserMapping.csv
  .\5d-EmptyRecycleBin.ps1 -Confirm
#>

# S9: standard confirmation (SupportsShouldProcess). -WhatIf previews,
# -Confirm prompts once; the old self-defined [switch]$ExecuteDelete (which
# shadowed the standard parameter) is removed.
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$MappingPath = '.\Output\UserMapping.csv',
    [string]$UserFilter,
    
    # NOTE (S9): permanent deletion is now gated by the standard
    # -Confirm/-WhatIf common parameters (SupportsShouldProcess). Without
    # -Confirm the script prompts once; -WhatIf runs a safe preview.
    
    # Number of items to delete per API request (max/recommended: 500)
    [ValidateRange(50, 1000)]
    [int]$BatchSize = 500,
    
    # Stop clearing once quota is healthy (> 500 MB free), rather than emptying all items
    [switch]$UntilQuotaFree
)

. "$PSScriptRoot\MigrationHelpers.ps1"

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$env:AZURE_IDENTITY_DISABLE_WAM = "true"
$env:AZURE_IDENTITY_DISABLE_INTERACTIVE_AUTH = "true"

try {
    [System.Net.ServicePointManager]::DefaultConnectionLimit = 500
    [System.Net.ServicePointManager]::Expect100Continue = $false
    [System.Net.ServicePointManager]::UseNagleAlgorithm = $false
} catch {}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if ([string]::IsNullOrWhiteSpace($ScriptDir)) {
    $ScriptDir = (Get-Location).Path
}
Set-Location -Path $ScriptDir

# Setup Logging
$LogDir = Join-Path -Path $ScriptDir -ChildPath 'Logs'
if (-not (Test-Path -Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogSuffix = if ($UserFilter) { "_" + ($UserFilter -replace '[^a-zA-Z0-9]', '_') } else { "_master" }
$LogPath = Join-Path -Path $LogDir -ChildPath "EmptyRecycleBin_${Timestamp}${LogSuffix}_${PID}.log"

function Write-Log {
    param ([string]$Message, [string]$Type = "INFO")
    $LogMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Type] $Message"
    $color = 'Cyan'
    if ($Type -eq 'ERROR') { $color = 'Red' }
    elseif ($Type -eq 'WARN') { $color = 'Yellow' }
    elseif ($Type -eq 'SUCCESS') { $color = 'Green' }
    
    Write-Host $LogMsg -ForegroundColor $color
    Add-Content -Path $LogPath -Value $LogMsg -Encoding UTF8 -ErrorAction SilentlyContinue
}

function Format-FileSize {
    param ([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

# Resolve paths
if (-not [System.IO.Path]::IsPathRooted($MappingPath)) {
    $MappingPath = Join-Path -Path $ScriptDir -ChildPath $MappingPath
}
if (-not [System.IO.Path]::IsPathRooted($ConfigPath)) {
    $ConfigPath = Join-Path -Path $ScriptDir -ChildPath $ConfigPath
}

$config = @{}
if (Test-Path $ConfigPath) {
    $config = Import-PowerShellDataFile -Path $ConfigPath
}

if (-not (Test-Path $MappingPath)) {
    Write-Log "Mapping file not found at $MappingPath" -Type ERROR
    exit 1
}
$Mapping = @(Import-Csv -Path $MappingPath -Encoding UTF8)
$isFilter = ($null -ne $UserFilter -and $UserFilter -ne '')
if ($isFilter) {
    $exactMatches = @($Mapping | Where-Object { 
        $_.SourceUPN -eq $UserFilter -or 
        $_.SourceUPN -like "$UserFilter@*" -or 
        $_.DestUPN -eq $UserFilter -or 
        $_.DestUPN -like "$UserFilter@*" 
    })
    if ($exactMatches.Count -gt 0) {
        $Mapping = $exactMatches
    } else {
        $Mapping = @($Mapping | Where-Object { $_.SourceUPN -match $UserFilter -or $_.DestUPN -match $UserFilter })
    }
    Write-Log "Filtered to user(s): $UserFilter ($($Mapping.Count) user(s))"
}

# S9: -Confirm is now the standard common parameter (SupportsShouldProcess).
# -WhatIf previews without prompting; otherwise PowerShell prompts once here.
$ExecuteDelete = $false
if (-not $WhatIfPreference) {
    $ExecuteDelete = $PSCmdlet.ShouldProcess("Recycle bins of $($Mapping.Count) source user(s) (PERMANENT deletion)", "Empty OneDrive recycle bins")
}
if (-not $ExecuteDelete -and -not $WhatIfPreference) {
    Write-Log "Recycle-bin purge cancelled by operator. Nothing was deleted." -Type WARN
    exit 0
}
# B2 validation gate (before any delete). Preview mode deletes nothing.
if ($ExecuteDelete) { Assert-ValidationGate }

Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "   ONEDRIVE RECYCLE BIN PURGE UTILITY" -ForegroundColor Cyan
if ($ExecuteDelete) {
    Write-Host "   Execution Mode: PURGE (-Confirm) - Items will be PERMANENTLY deleted" -ForegroundColor Red
} else {
    Write-Host "   Execution Mode: PREVIEW (What-If) - NO items will be deleted" -ForegroundColor Green
}
Write-Host "   Target Users  : $($Mapping.Count)" -ForegroundColor Cyan
Write-Host "==========================================================`n" -ForegroundColor Cyan

# Graph Auth
$defaultSrcTenantId = $config.SourceTenantDomain
$defaultSrcClientId = $config.SourceClientId
$defaultSrcClientSecret = $config.SourceClientSecret

function Ensure-SourceGraphAuth {
    if (Get-MgContext) { return $true }
    
    $cId = if ($config.SourceClientId) { $config.SourceClientId } else { $defaultSrcClientId }
    $cSecret = if ($config.SourceClientSecret) { $config.SourceClientSecret } else { $defaultSrcClientSecret }
    $tId = if ($config.SourceTenantDomain) { $config.SourceTenantDomain } else { $defaultSrcTenantId }
    
    if ($cId -and $cSecret -and $tId) {
        try {
            $secSecret = ConvertTo-SecureString $cSecret -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential($cId, $secSecret)
            Connect-MgGraph -ClientSecretCredential $cred -TenantId $tId -NoWelcome -ErrorAction Stop | Out-Null
            if (Get-MgContext) { return $true }
        } catch {}
    }
    
    $sourceScopes = @('User.Read.All', 'Files.ReadWrite.All', 'Sites.ReadWrite.All', 'Directory.Read.All', 'Organization.Read.All')
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            if ($tId) {
                Connect-MgGraph -Scopes $sourceScopes -TenantId $tId -NoWelcome -ErrorAction Stop | Out-Null
            } else {
                Connect-MgGraph -Scopes $sourceScopes -NoWelcome -ErrorAction Stop | Out-Null
            }
            if (Get-MgContext) { return $true }
        } catch {
            Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 1000)
        }
    }
    return $false
}

try { Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction SilentlyContinue } catch {}

if (-not (Ensure-SourceGraphAuth)) {
    Write-Log "Failed to acquire Graph authorization. Aborting." -Type ERROR
    exit 1
}

$authIdentity = if ((Get-MgContext).Account) { (Get-MgContext).Account } else { (Get-MgContext).AppName }
Write-Log "Authenticated to Graph as $authIdentity" -Type SUCCESS

# B1: refuse to purge against the wrong tenant.
Assert-Tenant -ExpectedTenantDomain (Get-RequiredConfigValue -Value $config.SourceTenantDomain -Name 'SourceTenantDomain')

function Get-DriveQuota {
    param ([string]$DriveId)
    if (-not $DriveId) { return $null }
    try {
        $d = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId" -ErrorAction SilentlyContinue
        if ($d -and $d.quota) {
            $q = $d.quota
            $u = if ($q -is [hashtable]) { $q['used'] } else { $q.used }
            $t = if ($q -is [hashtable]) { $q['total'] } else { $q.total }
            $r = if ($q -is [hashtable]) { $q['remaining'] } else { $q.remaining }
            $del = if ($q -is [hashtable]) { $q['deleted'] } else { $q.deleted }
            $s = if ($q -is [hashtable]) { $q['state'] } else { $q.state }
            return [PSCustomObject]@{
                Used      = [long]$u
                Total     = [long]$t
                Remaining = [long]$r
                Deleted   = [long]$del
                State     = [string]$s
            }
        }
    } catch {}
    return $null
}

function Get-UserSharePointSiteId {
    # S10: use the standard Graph drive resource (parentReference.siteId)
    # instead of cutting the '/Documents' suffix off the drive webUrl.
    # Falls back to the /sites/{hostname}:/personal/{upn} address.
    param([string]$UserId)
    try {
        $drive = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UserId/drive?`$select=id,parentReference,webUrl" -ErrorAction Stop
        $siteId = $null
        if ($drive.parentReference -and $drive.parentReference.siteId) {
            $siteId = $drive.parentReference.siteId
        } else {
            $u = [System.Uri]$drive.webUrl
            $slug = ($UserId -replace '@', '_') -replace '\.', '_'
            $site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($u.Host):/personal/$slug" -ErrorAction Stop
            $siteId = $site.id
        }
        return @{ SiteId = $siteId; DriveId = $drive.id }
    } catch {
        return $null
    }
}

$SummaryReport = @()

foreach ($User in $Mapping) {
    $upn = $User.SourceUPN
    Write-Log "--------------------------------------------------"
    Write-Log "Processing: $upn"
    
    $siteInfo = Get-UserSharePointSiteId -UserId $upn
    if (-not $siteInfo -or -not $siteInfo.SiteId) {
        Write-Log "  Cannot resolve SharePoint SiteId for $upn" -Type WARN
        $SummaryReport += [PSCustomObject]@{
            User      = $upn
            Status    = "Failed (No Site)"
            DeletedMB = 0
            Purged    = 0
        }
        continue
    }
    
    $siteId = $siteInfo.SiteId
    $driveId = $siteInfo.DriveId
    
    $quotaBefore = Get-DriveQuota -DriveId $driveId
    if ($quotaBefore) {
        Write-Log "  [QUOTA BEFORE] Used: $(Format-FileSize $quotaBefore.Used) | Recycle Bin: $(Format-FileSize $quotaBefore.Deleted) | Remaining: $(Format-FileSize $quotaBefore.Remaining) | State: $($quotaBefore.State)"
    }
    
    if (-not $quotaBefore -or $quotaBefore.Deleted -eq 0) {
        Write-Log "  Recycle Bin is already empty for $upn. Skipping." -Type SUCCESS
        $SummaryReport += [PSCustomObject]@{
            User      = $upn
            Status    = "Already Empty"
            DeletedMB = 0
            Purged    = 0
        }
        continue
    }
    
    if (-not $ExecuteDelete) {
        Write-Log "  [WHAT-IF] Recycle Bin contains $(Format-FileSize $quotaBefore.Deleted). Run with -Confirm to purge." -Type INFO
        $SummaryReport += [PSCustomObject]@{
            User      = $upn
            Status    = "Preview ($(Format-FileSize $quotaBefore.Deleted))"
            # Preview deletes nothing, so DeletedMB is 0; the amount that
            # WOULD be freed is in the Status text above.
            DeletedMB = 0
            Purged    = 0
        }
        continue
    }
    
    # Execute Purge
    # NOTE (S10): the site recycleBin list/bulk-delete endpoints below are
    # Graph BETA - v1.0 has no equivalent endpoint for site recycleBin items
    # (checked 2026-11-08). Beta carries no SLA and can change without
    # notice; re-check v1.0 before each production run and switch if an
    # equivalent stable endpoint appears.
    Write-Log "  Purging Recycle Bin items in batches of $BatchSize..." -Type INFO
    $userPurged = 0
    $batchNum = 0
    
    while ($true) {
        $batchNum++
        $rb = $null
        for ($a = 1; $a -le 3; $a++) {
            try {
                $rb = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/sites/$siteId/recycleBin/items?`$top=$BatchSize&`$select=id" -ErrorAction Stop
                break
            } catch {
                Start-Sleep -Seconds 2
            }
        }
        
        if (-not $rb -or -not $rb.value -or $rb.value.Count -eq 0) {
            Write-Log "  No more items found in Recycle Bin for $upn." -Type SUCCESS
            break
        }
        
        $ids = @($rb.value | ForEach-Object { $_.id })
        $body = @{ ids = $ids } | ConvertTo-Json -Compress
        
        $delSuccess = $false
        for ($a = 1; $a -le 3; $a++) {
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/sites/$siteId/recycleBin/items/delete" -Body $body -ContentType "application/json" -ErrorAction Stop | Out-Null
                $delSuccess = $true
                break
            } catch {
                Start-Sleep -Seconds 2
            }
        }
        
        if ($delSuccess) {
            $userPurged += $ids.Count
            Write-Log "    Purged $userPurged items from Recycle Bin (Batch $batchNum)..."
        } else {
            Write-Log "    Failed purging batch $batchNum for $upn." -Type WARN
            break
        }
        
        if ($UntilQuotaFree -and $driveId) {
            $curQ = Get-DriveQuota -DriveId $driveId
            if ($curQ -and $curQ.Remaining -gt 500MB) {
                Write-Log "    [QUOTA RECLAIMED] Quota remaining is $(Format-FileSize $curQ.Remaining). Stopping early." -Type SUCCESS
                break
            }
        }
        
        if ($batchNum % 10 -eq 0) {
            [System.GC]::Collect()
        }
    }
    
    # Check quota after
    Start-Sleep -Seconds 2
    $quotaAfter = Get-DriveQuota -DriveId $driveId
    if ($quotaAfter) {
        Write-Log "  [QUOTA AFTER]  Used: $(Format-FileSize $quotaAfter.Used) | Recycle Bin: $(Format-FileSize $quotaAfter.Deleted) | Remaining: $(Format-FileSize $quotaAfter.Remaining) | State: $($quotaAfter.State)" -Type SUCCESS
        if ($quotaBefore -and $quotaAfter.Deleted -lt $quotaBefore.Deleted) {
            $diff = $quotaBefore.Deleted - $quotaAfter.Deleted
            Write-Log "  [FREED FROM BIN] Reclaimed $(Format-FileSize $diff) from Recycle Bin!" -Type SUCCESS
        }
    }
    
    $SummaryReport += [PSCustomObject]@{
        User      = $upn
        Status    = "Purged ($userPurged items)"
        # S10: DeletedMB is the amount actually REMOVED from the recycle bin
        # (quota before minus quota after), not what is left in the bin.
        DeletedMB = if ($quotaAfter -and $quotaBefore) { [math]::Round([Math]::Max(0, ($quotaBefore.Deleted - $quotaAfter.Deleted)) / 1MB, 1) } else { 0 }
        Purged    = $userPurged
    }
    
    [System.GC]::Collect()
}

Write-Log "=========================================================="
Write-Log "RECYCLE BIN PURGE SUMMARY"
$SummaryReport | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor Cyan
Write-Log "=========================================================="
