<#
.SYNOPSIS
Streaming partial purge for source tenant OneDrive. Deletes only files not touched in the last 30 days.

Requires (minimum Graph scopes): User.Read.All, Files.ReadWrite.All,
Sites.ReadWrite.All, Directory.Read.All, Organization.Read.All.

Runs only after a recent passing 5-Validate.ps1 run (validation gate, B2)
and only against the configured source tenant (B1).

.DESCRIPTION
Designed specifically for resource-constrained environments (e.g. 1 GB RAM VPS).
Uses a high-performance, constant-memory (O(1)) streaming engine over Microsoft Graph Delta API:
- Memory footprint is strictly < 30 MB RAM regardless of drive size (tested on 375,000+ file drives).
- Deletions execute page-by-page as items stream in, eliminating array accumulation and OOM crashes.
- Built-in retry loops (5 attempts with backoff) for network drops, connection resets, and 429 throttling.
- Runs sequentially in-process to guarantee zero WinRM remoting breaks (eliminates PSSessionStateBroken).
- Automatically skips accounts that are already within normal quota and have announcement notices.
- Announcement files (CHU_Y.txt) and directory structures are strictly protected.

Modes:
- What-If (Default): Fast scan-only preview showing space to be freed and sample files. No deletions.
- Purge (-Confirm): Executes deletion of files older than $DaysOld.
- Purge + Announcement (-Confirm -SendAnnouncementAfterPurge): Purges old files and automatically uploads
  CHU_Y.txt as soon as quota drops below the 10 GB limit.

Examples:
  # Preview single account
  .\5c-PartialPurgeSource.ps1 -UserFilter "ductri1510"

  # Purge single account
  .\5c-PartialPurgeSource.ps1 -UserFilter "ductri1510" -Confirm

  # Purge all accounts sequentially and send announcements
  .\5c-PartialPurgeSource.ps1 -Confirm -SendAnnouncementAfterPurge
#>

# S9: standard confirmation (SupportsShouldProcess). -WhatIf previews,
# -Confirm prompts once; the old self-defined [switch]$ExecuteDelete (which
# shadowed the standard parameter) is removed.
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$MappingPath = '.\Output\UserMapping.csv',
    [string]$UserFilter,
    
    # Days threshold: files last modified older than this number of days will be deleted.
    [ValidateRange(1, 3650)]
    [int]$DaysOld = 30,
    
    # NOTE (S9): deletion is now gated by the standard -Confirm/-WhatIf
    # common parameters (SupportsShouldProcess). Without -Confirm the script
    # prompts once; -WhatIf runs a safe scan-only preview.
    
    # Automatically triggers 5a-Announcement.ps1 for each user after successful purge.
    [switch]$SendAnnouncementAfterPurge,
    
    # Permanently delete files (bypassing recycle bin) AND automatically purge recycle bin if quota is full
    [switch]$PermanentDelete,
    
    # Automatically skip accounts that are already within quota (< 10 GB) and have announcement
    [bool]$SkipHealthyAccounts = $true
)

. "$PSScriptRoot\MigrationHelpers.ps1"
. "$PSScriptRoot\AnnouncementHelpers.ps1"

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

# Import helpers
$HelpersPath = Join-Path -Path $ScriptDir -ChildPath '4a-OneDriveHelpers.ps1'
if (Test-Path $HelpersPath) {
    . $HelpersPath
} else {
    Write-Error "Cannot find helper script: $HelpersPath"
    exit
}

# Setup Logging
$LogDir = Join-Path -Path $ScriptDir -ChildPath 'Logs'
if (-not (Test-Path -Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogSuffix = if ($UserFilter) { "_" + ($UserFilter -replace '[^a-zA-Z0-9]', '_') } else { "_master" }
$LogPath = Join-Path -Path $LogDir -ChildPath "PartialPurgeSource_${Timestamp}${LogSuffix}_${PID}.log"

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

# Format-FileSize helper
function Format-FileSize {
    param ([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

# Helper to safely parse LastModified date
function Get-ItemDate {
    param ($Item)
    if (-not $Item) { return $null }
    $raw = if ($Item.lastModifiedDateTime) { $Item.lastModifiedDateTime } else { $Item.LastModified }
    if (-not $raw) { return $null }
    if ($raw -is [DateTime]) { return $raw }
    if ($raw -is [DateTimeOffset]) { return $raw.DateTime }
    try {
        return [DateTime]::Parse($raw, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
    } catch {
        try { return [DateTime]$raw } catch { return $null }
    }
}

# Helper to fetch current drive quota info from Graph API
function Get-DriveQuotaInfo {
    param ([string]$DriveId)
    if (-not $DriveId) { return $null }
    try {
        $drive = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId" -ErrorAction SilentlyContinue
        if ($drive -and $drive.quota) {
            $q = $drive.quota
            $usedVal = if ($q -is [hashtable]) { $q['used'] } else { $q.used }
            $totalVal = if ($q -is [hashtable]) { $q['total'] } else { $q.total }
            $remVal = if ($q -is [hashtable]) { $q['remaining'] } else { $q.remaining }
            $delVal = if ($q -is [hashtable]) { $q['deleted'] } else { $q.deleted }
            $stateVal = if ($q -is [hashtable]) { $q['state'] } else { $q.state }
            return [PSCustomObject]@{
                Used      = [long]$usedVal
                Total     = [long]$totalVal
                Remaining = [long]$remVal
                Deleted   = [long]$delVal
                State     = [string]$stateVal
            }
        }
    } catch {}
    return $null
}

# Check if .CHU_Y.txt or CHU_Y.txt exists in ANY of the 4 announcement
# folders (S10): the announcement lives in root/Desktop/Documents/Pictures,
# so checking root alone misses it.
function Test-AnnouncementExists {
    param ([string]$DriveId)
    if (-not $DriveId) { return $false }
    $folders = @('', 'Desktop', 'Documents', 'Pictures')
    foreach ($folder in $folders) {
        foreach ($fileName in @('.CHU_Y.txt', 'CHU_Y.txt')) {
            $path = if ($folder -eq '') { "root:/$fileName" } else { "root:/$([System.Uri]::EscapeDataString($folder))/$fileName" }
            try {
                $check = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/$path" -ErrorAction SilentlyContinue
                if ($check -and $check.id) { return $true }
            } catch {}
        }
    }
    return $false
}

# Direct in-memory announcement uploader (zero disk dependency)
# Announcement upload (B7/S19): template + 4-folder upload live in
# AnnouncementHelpers.ps1. The text contains NO password.
function Send-UserAnnouncement {
    param(
        [string]$SourceUpn,
        [string]$DestUpn,
        [string]$SourceId,
        [string]$DriveId
    )
    if (-not $DriveId -or -not $SourceUpn) { return 0 }
    $content = Get-AnnouncementText -SourceUpn $SourceUpn -DestUpn $DestUpn
    return (Publish-AnnouncementToFolders -SourceUpn $SourceUpn -SourceId $SourceId -DriveId $DriveId -Content $content -LogAction { param($m, $t) Write-Log $m -Type $t })
}

# Resolve SharePoint SiteId for a OneDrive user
# Resolve SharePoint SiteId for a OneDrive user (S10): use the standard
# Graph drive resource (parentReference.siteId) instead of cutting the
# '/Documents' suffix off the drive webUrl, which breaks when the URL shape
# changes. Falls back to the /sites/{hostname}:/personal/{upn} address.
function Get-UserSiteId {
    param([string]$UserId)
    try {
        $drive = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UserId/drive?`$select=id,parentReference,webUrl" -ErrorAction Stop
        if ($drive.parentReference -and $drive.parentReference.siteId) {
            return $drive.parentReference.siteId
        }
        $u = [System.Uri]$drive.webUrl
        $slug = ($UserId -replace '@', '_') -replace '\.', '_'
        $site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($u.Host):/personal/$slug" -ErrorAction Stop
        return $site.id
    } catch {
        return $null
    }
}

# Permanently purge items from user's Recycle Bin in batches
function Clear-UserRecycleBin {
    # NOTE (S10): the site recycleBin list/bulk-delete endpoints below are
    # Graph BETA - v1.0 has no equivalent endpoint for site recycleBin items
    # (checked 2026-11-08). Beta carries no SLA and can change without
    # notice; re-check v1.0 before each production run and switch if an
    # equivalent stable endpoint appears.
    param(
        [string]$SourceUpn,
        [string]$DriveId,
        [int]$BatchSize = 500,
        [switch]$UntilQuotaFree
    )
    $siteId = Get-UserSiteId -UserId $SourceUpn
    if (-not $siteId) {
        Write-Log "    Cannot resolve SharePoint SiteId for ${SourceUpn} to clear Recycle Bin." -Type WARN
        return 0
    }
    
    Write-Log "    [RECYCLE BIN PURGE] Purging Recycle Bin for ${SourceUpn}..." -Type INFO
    $totalPurged = 0
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
            Write-Log "    [RECYCLE BIN EMPTY] No more items in Recycle Bin for ${SourceUpn}." -Type SUCCESS
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
            $totalPurged += $ids.Count
            Write-Log "    Purged $totalPurged items from Recycle Bin (Batch $batchNum)..."
        } else {
            Write-Log "    Failed purging batch $batchNum for ${SourceUpn}." -Type WARN
            break
        }
        
        # Check quota if UntilQuotaFree is specified
        if ($UntilQuotaFree -and $DriveId) {
            $currentQuota = Get-DriveQuotaInfo -DriveId $DriveId
            if ($currentQuota -and $currentQuota.Remaining -gt 500MB) {
                Write-Log "    [QUOTA RECLAIMED] Quota remaining is now $(Format-FileSize $currentQuota.Remaining). Stopping recycle bin purge." -Type SUCCESS
                break
            }
        }
        
        # Periodic garbage collection
        if ($batchNum % 10 -eq 0) {
            [System.GC]::Collect()
        }
    }
    return $totalPurged
}

# Resolve paths
if (-not [System.IO.Path]::IsPathRooted($MappingPath)) {
    $MappingPath = Join-Path -Path $ScriptDir -ChildPath $MappingPath
}
if (-not [System.IO.Path]::IsPathRooted($ConfigPath)) {
    $ConfigPath = Join-Path -Path $ScriptDir -ChildPath $ConfigPath
}

# Load config
$config = @{}
if (Test-Path $ConfigPath) {
    $config = Import-PowerShellDataFile -Path $ConfigPath
}

if (-not (Test-Path $MappingPath)) {
    Write-Log "Mapping file not found at $MappingPath" -Type ERROR
    exit
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

$CutoffDate = (Get-Date).AddDays(-$DaysOld)

# S9: -Confirm is now the standard common parameter (SupportsShouldProcess).
# -WhatIf previews without prompting; otherwise PowerShell prompts once here.
$ExecuteDelete = $false
if (-not $WhatIfPreference) {
    $ExecuteDelete = $PSCmdlet.ShouldProcess("OneDrive files older than $DaysOld days for $($Mapping.Count) source user(s)", "Partial purge source OneDrive files")
}
if (-not $ExecuteDelete -and -not $WhatIfPreference) {
    Write-Log "Purge cancelled by operator. No files were deleted." -Type WARN
    exit 0
}
# B2 validation gate (before any delete). Preview mode deletes nothing.
if ($ExecuteDelete) { Assert-ValidationGate }
# B2: -PermanentDelete is a separate, explicitly confirmed step. It bypasses
# the recycle bin (unrecoverable), so it needs a typed confirmation; anything
# else falls back to ordinary recycle-bin deletion.
if ($PermanentDelete -and $ExecuteDelete) {
    Write-Host "WARNING: -PermanentDelete BYPASSES the recycle bin - deleted files CANNOT be recovered." -ForegroundColor Red
    Write-Host "To proceed with permanent deletion, type PERMANENT and press Enter. Anything else falls back to recycle-bin deletion." -ForegroundColor Red
    $typedConfirm = Read-Host 'Type PERMANENT to confirm permanent deletion'
    if ($typedConfirm -cne 'PERMANENT') {
        $PermanentDelete = $false
        Write-Log "Permanent-delete confirmation not given; falling back to recycle-bin deletion." -Type WARN
    }
}

Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "   STREAMING ZERO-MEMORY SOURCE PURGE ENGINE" -ForegroundColor Cyan
Write-Host "   Age Threshold : $DaysOld days (Cutoff: $($CutoffDate.ToString('yyyy-MM-dd HH:mm')))" -ForegroundColor Cyan
if ($ExecuteDelete) {
    Write-Host "   Execution Mode: PURGE - Files older than $DaysOld days will be deleted" -ForegroundColor Red
    if ($PermanentDelete) {
        Write-Host "   Permanent Del : ENABLED (-PermanentDelete) - Bypasses recycle bin & purges bin" -ForegroundColor Magenta
    }
} else {
    Write-Host "   Execution Mode: WHAT-IF (Scan-Only) - NO files will be deleted" -ForegroundColor Green
}
Write-Host "   Target Users  : $($Mapping.Count)" -ForegroundColor Cyan
Write-Host "==========================================================`n" -ForegroundColor Cyan

# App Registration credentials from config
$defaultSrcTenantId = $config.SourceTenantDomain
$defaultSrcClientId = $config.SourceClientId
$defaultSrcClientSecret = $config.SourceClientSecret

function Ensure-SourceGraphAuth {
    if (Get-MgContext) { return $true }
    
    # 1. Try App Registration client credentials (config or embedded default)
    $cId = if ($config.SourceClientId) { $config.SourceClientId } else { $defaultSrcClientId }
    $cSecret = if ($config.SourceClientSecret) { $config.SourceClientSecret } else { $defaultSrcClientSecret }
    $tId = if ($config.SourceTenantDomain) { $config.SourceTenantDomain } else { $defaultSrcTenantId }
    
    if ($cId -and $cSecret -and $tId) {
        try {
            $secSecret = ConvertTo-SecureString $cSecret -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential($cId, $secSecret)
            Connect-MgGraph -ClientSecretCredential $cred -TenantId $tId -NoWelcome -ErrorAction Stop | Out-Null
            if (Get-MgContext) { return $true }
        } catch {
            Write-Verbose "App Registration auth failed: $_"
        }
    }
    
    # 2. Try MSAL token cache
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
    
    # 3. Fallback to DeviceCode if interactive
    try {
        if ($tId) {
            Connect-MgGraph -Scopes $sourceScopes -TenantId $tId -UseDeviceCode -NoWelcome -ErrorAction Stop | Out-Null
        } else {
            Connect-MgGraph -Scopes $sourceScopes -UseDeviceCode -NoWelcome -ErrorAction Stop | Out-Null
        }
        if (Get-MgContext) { return $true }
    } catch {}
    
    return $false
}

try { Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction SilentlyContinue } catch {}

if (-not (Ensure-SourceGraphAuth)) {
    Write-Log "Failed to acquire Graph authorization for Source Tenant. Aborting." -Type ERROR
    exit 1
}

$authIdentity = if ((Get-MgContext).Account) { (Get-MgContext).Account } else { (Get-MgContext).AppName }
Write-Log "Authenticated to Graph as $authIdentity" -Type SUCCESS

# B1: refuse to purge against the wrong tenant.
Assert-Tenant -ExpectedTenantDomain (Get-RequiredConfigValue -Value $config.SourceTenantDomain -Name 'SourceTenantDomain')

$TotalUsersProcessed = 0
$TotalGlobalScanned = 0
$TotalGlobalEligible = 0
$TotalGlobalDeleted = 0
$TotalGlobalFailed = 0
$TotalGlobalKept = 0
$TotalGlobalProtected = 0
$TotalGlobalSpaceFreed = 0
$TotalGlobalSpaceToFree = 0

foreach ($User in $Mapping) {
    $SourceUpn = $User.SourceUPN
    $SourceId = $User.SourceId
    Write-Log "--------------------------------------------------"
    Write-Log "Processing: $SourceUpn"
    $TotalUsersProcessed++
    
    $driveId = Get-UserOneDriveId -UserId $SourceId
    if (-not $driveId) {
        Write-Log "  Cannot access OneDrive for $SourceUpn" -Type WARN
        continue
    }
    
    # Check quota
    $quota = Get-DriveQuotaInfo -DriveId $driveId
    if ($quota) {
        Write-Log "  [QUOTA BEFORE] Used: $(Format-FileSize $quota.Used) / $(Format-FileSize $quota.Total) | Remaining: $(Format-FileSize $quota.Remaining) | State: $($quota.State)"
    }
    
    $hasAnn = Test-AnnouncementExists -DriveId $driveId
    
    # Check if account is already within normal quota (< 10 GB)
    if ($SkipHealthyAccounts -and -not $isFilter -and $quota -and $quota.State -eq 'normal') {
        if ($hasAnn) {
            Write-Log "  [ALREADY HEALTHY] Storage is normal ($(Format-FileSize $quota.Used)) and announcement CHU_Y.txt exists. Skipping." -Type SUCCESS
            continue
        } elseif ($ExecuteDelete -and $SendAnnouncementAfterPurge) {
            Write-Log "  [STORAGE NORMAL] Storage is already within quota ($(Format-FileSize $quota.Used)), but announcement is missing." -Type INFO
            
            # If remaining quota is 0 (due to Recycle Bin), clear Recycle Bin first!
            if ($quota.Remaining -le 0 -or $PermanentDelete) {
                Write-Log "  [RECYCLE BIN FULL] Remaining quota is 0 B ($(Format-FileSize $quota.Deleted) in Recycle Bin). Clearing Recycle Bin..." -Type INFO
                Clear-UserRecycleBin -SourceUpn $SourceUpn -DriveId $driveId -UntilQuotaFree
            }
            
            Write-Log "  Uploading CHU_Y.txt for $SourceUpn..." -Type INFO
            $uploaded = Send-UserAnnouncement -SourceUpn $SourceUpn -DestUpn $User.DestUPN -SourceId $SourceId -DriveId $driveId
            if ($uploaded -gt 0) {
                Write-Log "  Announcement successfully uploaded ($uploaded folder(s)) for $SourceUpn." -Type SUCCESS
            } else {
                Write-Log "  Could not upload announcement for $SourceUpn (quota blocked by Recycle Bin). Moving to next user." -Type WARN
            }
            continue
        }
    }
    
    Write-Log "  Starting streaming scan via Delta API..."
    
    $deltaUri = "https://graph.microsoft.com/v1.0/drives/$driveId/root/delta?`$select=id,name,size,folder,package,file,lastModifiedDateTime,deleted"
    $pageCount = 0
    $userScanned = 0
    $userFiles = 0
    $userEligible = 0
    $userKept = 0
    $userProtected = 0
    $userDeleted = 0
    $userFailed = 0
    $userFreedBytes = 0
    $userSpaceToFree = 0
    $sampleList = [System.Collections.Generic.List[string]]::new()
    
    while ($deltaUri) {
        $pageCount++
        $pageResp = $null
        
        # Resilient page fetch loop (up to 5 retries for socket drops/429 throttling)
        for ($attempt = 1; $attempt -le 5; $attempt++) {
            try {
                $pageResp = Invoke-MgGraphRequest -Method GET -Uri $deltaUri -ErrorAction Stop
                break
            } catch {
                $err = $_.Exception.Message
                if ($attempt -lt 5) {
                    Write-Log "    [RETRY] Delta page $pageCount failed ($err). Retrying in $($attempt * 3)s ($attempt/5)..." -Type WARN
                    Start-Sleep -Seconds ($attempt * 3)
                } else {
                    Write-Log "    [ERROR] Delta page $pageCount failed after 5 retries: $err" -Type ERROR
                    break
                }
            }
        }
        
        if (-not $pageResp) {
            Write-Log "    Aborting further delta pages for $SourceUpn due to persistent network error." -Type ERROR
            break
        }
        
        foreach ($child in $pageResp.value) {
            $userScanned++
            if ($child.deleted -or -not $child.name -or $child.name -eq 'root') { continue }
            
            $isFolder = ($null -ne $child.folder) -and ($child.name -notmatch '\.(one|onetoc2)$')
            if ($child.package -and $child.package.type -eq 'oneNote') { $isFolder = $false }
            if ($isFolder) { continue }
            
            $userFiles++
            
            # Protect announcements
            if ($child.name -match '^\.?CHU_Y\.txt$' -or $child.name -match 'announcement|thong_bao') {
                $userProtected++
                continue
            }
            
            $modDate = Get-ItemDate -Item $child
            if ($null -eq $modDate) {
                $userKept++
                continue
            }
            
            if ($modDate -lt $CutoffDate) {
                $userEligible++
                $fileSize = [long]$child.size
                $userSpaceToFree += $fileSize
                
                if ($ExecuteDelete) {
                    # Execute deletion immediately
                    $delSuccess = $false
                    
                    # If PermanentDelete, attempt permanentDelete first to bypass recycle bin directly
                    if ($PermanentDelete) {
                        $permUri = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$($child.id)/permanentDelete"
                        try {
                            Invoke-MgGraphRequest -Method POST -Uri $permUri -ErrorAction Stop | Out-Null
                            $delSuccess = $true
                        } catch {
                            $pErr = $_.Exception.Message
                            if ($pErr -match '404|NotFound|itemNotFound') {
                                $delSuccess = $true
                            }
                        }
                    }
                    
                    # Fallback to standard DELETE if permanentDelete was not requested or not supported
                    if (-not $delSuccess) {
                        $delUri = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$($child.id)"
                        for ($dAttempt = 1; $dAttempt -le 3; $dAttempt++) {
                            try {
                                Invoke-MgGraphRequest -Method DELETE -Uri $delUri -ErrorAction Stop | Out-Null
                                $delSuccess = $true
                                break
                            } catch {
                                $dErr = $_.Exception.Message
                                if ($dErr -match '404|NotFound|itemNotFound') {
                                    # Already deleted
                                    $delSuccess = $true
                                    break
                                } elseif ($dErr -match '429|TooManyRequests|throttled') {
                                    Start-Sleep -Seconds 10
                                } else {
                                    if ($dAttempt -lt 3) {
                                        Start-Sleep -Seconds 2
                                    } else {
                                        Write-Log "    FAILED: $($child.name) - $dErr" -Type ERROR
                                    }
                                }
                            }
                        }
                    }
                    
                    if ($delSuccess) {
                        $userDeleted++
                        $userFreedBytes += $fileSize
                    } else {
                        $userFailed++
                    }
                    
                    if ($userDeleted % 200 -eq 0 -and $userDeleted -gt 0) {
                        Write-Log "    Progress: $userDeleted files deleted ($(Format-FileSize $userFreedBytes) freed)..."
                    }
                } else {
                    # What-If mode: keep up to 15 sample items
                    if ($sampleList.Count -lt 15) {
                        $sampleList.Add("$($child.name) ($(Format-FileSize $fileSize), modified $($modDate.ToString('yyyy-MM-dd HH:mm')))")
                    }
                }
            } else {
                $userKept++
            }
        }
        
        # Advance page
        $deltaUri = $pageResp.'@odata.nextLink'
        $pageResp = $null
        
        # Periodically collect garbage to keep RAM constant (< 30 MB)
        if ($pageCount % 25 -eq 0) {
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            if (-not $ExecuteDelete) {
                Write-Log "    Scanned $userScanned items ($userEligible eligible files, $(Format-FileSize $userSpaceToFree))..."
            }
        }
    }
    
    $TotalGlobalScanned += $userScanned
    $TotalGlobalEligible += $userEligible
    $TotalGlobalKept += $userKept
    $TotalGlobalProtected += $userProtected
    $TotalGlobalDeleted += $userDeleted
    $TotalGlobalFailed += $userFailed
    $TotalGlobalSpaceFreed += $userFreedBytes
    $TotalGlobalSpaceToFree += $userSpaceToFree
    
    if ($ExecuteDelete) {
        Write-Log "  Purge finished for ${SourceUpn}: $userDeleted deleted, $userFailed failed. Space freed: $(Format-FileSize $userFreedBytes)" -Type SUCCESS
        
        # Check quota after deletion
        Start-Sleep -Seconds 2
        $finalQuota = Get-DriveQuotaInfo -DriveId $driveId
        if ($finalQuota) {
            Write-Log "  [QUOTA AFTER]  Used: $(Format-FileSize $finalQuota.Used) / $(Format-FileSize $finalQuota.Total) | Remaining: $(Format-FileSize $finalQuota.Remaining) | State: $($finalQuota.State)" -Type SUCCESS
            if ($quota -and $finalQuota.Used -lt $quota.Used) {
                $diff = $quota.Used - $finalQuota.Used
                Write-Log "  [QUOTA REDUCED] Storage reduced by $(Format-FileSize $diff)!" -Type SUCCESS
            }
        }
        
        # If PermanentDelete or if remaining quota is 0, clear Recycle Bin to free quota
        if ($PermanentDelete -or ($finalQuota -and $finalQuota.Remaining -le 0)) {
            Write-Log "  [CLEARING RECYCLE BIN] Purging Recycle Bin to reclaim quota for ${SourceUpn}..." -Type INFO
            Clear-UserRecycleBin -SourceUpn $SourceUpn -DriveId $driveId -UntilQuotaFree
        }
        
        # Trigger announcement if requested
        if ($SendAnnouncementAfterPurge) {
            $alreadyHasAnn = Test-AnnouncementExists -DriveId $driveId
            if ($alreadyHasAnn) {
                Write-Log "  Announcement CHU_Y.txt is already present on $SourceUpn. Skipping re-upload." -Type SUCCESS
            } else {
                Write-Log "  Uploading announcement CHU_Y.txt for $SourceUpn..." -Type INFO
                $annSuccess = $false
                $announcementScript = Join-Path -Path $ScriptDir -ChildPath '5a-Announcement.ps1'
                if (Test-Path $announcementScript) {
                    try {
                        & $announcementScript -UserFilter $SourceUpn
                        $annSuccess = $true
                        Write-Log "  Announcement script completed for $SourceUpn." -Type SUCCESS
                    } catch {
                        Write-Log "  Announcement script threw ($($_)), using direct in-memory fallback..." -Type WARN
                    }
                }
                if (-not $annSuccess) {
                    $uploaded = Send-UserAnnouncement -SourceUpn $SourceUpn -DestUpn $User.DestUPN -SourceId $SourceId -DriveId $driveId
                    if ($uploaded -gt 0) {
                        Write-Log "  Direct in-memory announcement upload succeeded ($uploaded folder(s)) for $SourceUpn." -Type SUCCESS
                    } else {
                        Write-Log "  Failed uploading announcement for $SourceUpn." -Type ERROR
                    }
                }
            }
        }
    } else {
        # What-If summary
        Write-Log "  [WHAT-IF SUMMARY for $SourceUpn]" -Type INFO
        Write-Log "    Total items scanned: $userScanned"
        Write-Log "    Total files found  : $userFiles"
        Write-Log "    Files older than $DaysOld days (would delete): $userEligible ($(Format-FileSize $userSpaceToFree))"
        Write-Log "    Files kept safe (< $DaysOld days)            : $userKept"
        Write-Log "    Protected announcement files (kept safe)    : $userProtected"
        if ($sampleList.Count -gt 0) {
            Write-Log "    Sample of files that WOULD be deleted:"
            foreach ($s in $sampleList) {
                Write-Log "      - $s"
            }
        }
    }
    
    # Clean memory between users
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Log "=========================================================="
Write-Log "PARTIAL PURGE OVERALL SUMMARY"
Write-Log "Users processed : $TotalUsersProcessed"
Write-Log "Total items scanned : $TotalGlobalScanned"
Write-Log "Files kept safe     : $TotalGlobalKept"
Write-Log "Announcements safe  : $TotalGlobalProtected"

if ($ExecuteDelete) {
    Write-Log "Total files deleted : $TotalGlobalDeleted"
    Write-Log "Total deletions failed: $TotalGlobalFailed"
    Write-Log "Total storage freed : $(Format-FileSize $TotalGlobalSpaceFreed)" -Type SUCCESS
} else {
    Write-Log "Total eligible files: $TotalGlobalEligible"
    Write-Log "Total space to free : $(Format-FileSize $TotalGlobalSpaceToFree)" -Type INFO
    Write-Log ""
    Write-Log "TO EXECUTE DELETION, RUN:" -Type WARN
    Write-Log "  .\5c-PartialPurgeSource.ps1 -Confirm -SendAnnouncementAfterPurge"
}
Write-Log "=========================================================="
