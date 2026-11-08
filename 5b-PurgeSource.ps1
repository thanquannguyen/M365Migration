<#
.SYNOPSIS
Purges ALL files and folders on each source user's OneDrive, EXCEPT CHU_Y.txt announcement files.

Requires (minimum Graph scopes): User.Read.All, Files.ReadWrite.All,
Sites.ReadWrite.All, Directory.Read.All, Organization.Read.All.

Runs only after a recent passing 5-Validate.ps1 run (validation gate, B2)
and only against the configured source tenant (B1).
#>

# S9: standard confirmation (SupportsShouldProcess). -WhatIf previews,
# -Confirm prompts once per user; the old self-defined Confirm switch
# (which shadowed the standard parameter) is removed.
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$MappingPath = '.\Output\UserMapping.csv',
    [string]$UserFilter
)


. "$PSScriptRoot\MigrationHelpers.ps1"

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$env:AZURE_IDENTITY_DISABLE_WAM = "true"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location -Path $ScriptDir

# Import helpers
$HelpersPath = Join-Path -Path $ScriptDir -ChildPath '4a-OneDriveHelpers.ps1'
if (Test-Path $HelpersPath) {
    . $HelpersPath
} else {
    Write-Error "Cannot find helper script: $HelpersPath"
    exit 1
}

# Setup Logging
$LogDir = Join-Path -Path $ScriptDir -ChildPath 'Logs'
if (-not (Test-Path -Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogSuffix = if ($UserFilter) { "_" + ($UserFilter -replace '[^a-zA-Z0-9]', '_') } else { "_master" }
$LogPath = Join-Path -Path $LogDir -ChildPath "PurgeSource_${Timestamp}${LogSuffix}_${PID}.log"

function Write-Log {
    param ([string]$Message, [string]$Type = "INFO")
    $LogMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Type] $Message"
    Write-Host $LogMsg -ForegroundColor $(if ($Type -eq 'ERROR') {'Red'} elseif ($Type -eq 'WARN') {'Yellow'} else {'Cyan'})
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

# Load config
$config = @{}
if (Test-Path $ConfigPath) {
    $config = Import-PowerShellDataFile -Path $ConfigPath
}

# B2 validation gate: purge only runs after a recent passing 5-Validate.ps1
# run. Preview (-WhatIf) deletes nothing, so the gate is skipped there.
if (-not $WhatIfPreference) { Assert-ValidationGate }

if (-not (Test-Path $MappingPath)) {
    Write-Log "Mapping file not found at $MappingPath" -Type ERROR
    exit 1
}
$Mapping = Import-Csv -Path $MappingPath -Encoding UTF8
$isFilter = ($null -ne $UserFilter -and $UserFilter -ne '')
if ($isFilter) {
    $Mapping = $Mapping | Where-Object { $_.SourceUPN -match $UserFilter }
    Write-Log "Filtered to user(s): $UserFilter ($($Mapping.Count) users)"
}

if ($WhatIfPreference) {
    Write-Host "`n==========================================================" -ForegroundColor Green
    Write-Host "   SCAN-ONLY MODE (WhatIf) - No files will be deleted" -ForegroundColor Green
    Write-Host "==========================================================`n" -ForegroundColor Green
} else {
    Write-Host "`n==========================================================" -ForegroundColor Red
    Write-Host "   PURGE MODE - FILES WILL BE DELETED (confirmation required)" -ForegroundColor Red
    Write-Host "==========================================================`n" -ForegroundColor Red
}

# Parallel Orchestrator
$MaxConcurrentUsers = if ($null -ne $config.MaxConcurrentUsers -and $config.MaxConcurrentUsers -gt 0) { [int]$config.MaxConcurrentUsers } else { 4 }
$MaxConcurrentUsers = [Math]::Min($MaxConcurrentUsers, $Mapping.Count)

if (-not $isFilter -and $Mapping.Count -gt 0) {
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host "   PARALLEL SOURCE PURGE ORCHESTRATOR" -ForegroundColor Cyan
    Write-Host "   Concurrent Threads: $MaxConcurrentUsers | Total Accounts: $($Mapping.Count)" -ForegroundColor Cyan
    Write-Host "==========================================================`n" -ForegroundColor Cyan
    
    $pendingUsers = [System.Collections.ArrayList]@($Mapping)
    $runningJobs = @{}
    $scriptPath = $MyInvocation.MyCommand.Definition
    
    Write-Host "[STEP 1] Authenticating with Source Tenant before launching parallel threads..." -ForegroundColor Yellow
    $sourceTenantId = $config.SourceTenantDomain
    $sourceScopes = @('User.Read.All', 'Files.ReadWrite.All', 'Sites.ReadWrite.All', 'Directory.Read.All', 'Organization.Read.All')
    Connect-MgGraph -Scopes $sourceScopes -NoWelcome -ErrorAction SilentlyContinue
    if (-not (Get-MgContext)) {
        if ($sourceTenantId) {
            Connect-MgGraph -Scopes $sourceScopes -TenantId $sourceTenantId -UseDeviceCode -NoWelcome *>&1 | Write-Host
        } else {
            Connect-MgGraph -Scopes $sourceScopes -UseDeviceCode -NoWelcome *>&1 | Write-Host
        }
    }
    Write-Host "Connected as $((Get-MgContext).Account). Launching parallel threads...`n" -ForegroundColor Green

    # B1: refuse to purge against the wrong tenant.
    Assert-Tenant -ExpectedTenantDomain (Get-RequiredConfigValue -Value $sourceTenantId -Name 'SourceTenantDomain')

    # S9: one confirmation for the whole orchestrator run. Children run with
    # -Confirm:$false (already confirmed here) or -WhatIf (preview mode).
    $orchestratorApproved = $false
    if (-not $WhatIfPreference) {
        $orchestratorApproved = $PSCmdlet.ShouldProcess("OneDrive contents of $($Mapping.Count) source user(s) (except CHU_Y.txt announcement files)", "Purge source OneDrive files")
    }
    if (-not $orchestratorApproved -and -not $WhatIfPreference) {
        Write-Host "Purge cancelled by operator. No files were deleted." -ForegroundColor Yellow
        return
    }

    try {
        while ($pendingUsers.Count -gt 0 -or $runningJobs.Count -gt 0) {
            while ($runningJobs.Count -lt $MaxConcurrentUsers -and $pendingUsers.Count -gt 0) {
                $user = $pendingUsers[0]
                $pendingUsers.RemoveAt(0)
                $upn = $user.SourceUPN
                Write-Host "[LAUNCHING] Thread for account: $upn" -ForegroundColor Green
                
                $jobArgs = @($scriptPath, $upn)
                if ($orchestratorApproved) {
                    $scriptBlock = {
                        param($script, $upnFilter)
                        & $script -UserFilter $upnFilter -Confirm:$false
                    }
                } else {
                    $scriptBlock = {
                        param($script, $upnFilter)
                        & $script -UserFilter $upnFilter -WhatIf
                    }
                }
                
                $job = Start-Job -Name "Purge-$upn" -ScriptBlock $scriptBlock -ArgumentList $jobArgs
                $runningJobs[$upn] = $job
            }
            
            $completed = @()
            foreach ($key in $runningJobs.Keys) {
                $j = $runningJobs[$key]
                if ($j.State -in @('Completed', 'Failed', 'Stopped')) {
                    $completed += $key
                    $output = Receive-Job -Job $j 2>&1
                    if ($j.State -eq 'Completed') {
                        Write-Host "[COMPLETED] Account: $key" -ForegroundColor Green
                    } else {
                        Write-Host "[FAILED/STOPPED] Account: $key - State: $($j.State)" -ForegroundColor Yellow
                    }
                    $output | ForEach-Object { Write-Host "   $_" }
                    Remove-Job -Job $j -Force -ErrorAction SilentlyContinue | Out-Null
                }
            }
            foreach ($c in $completed) {
                $runningJobs.Remove($c)
            }
            if ($runningJobs.Count -gt 0) {
                Start-Sleep -Seconds 3
            }
        }
    } finally {
        Get-Job | Where-Object { $_.Name -like "Purge-*" } | ForEach-Object { Stop-Job -Job $_ -ErrorAction SilentlyContinue; Remove-Job -Job $_ -Force -ErrorAction SilentlyContinue }
    }
    
    if ($WhatIfPreference) {
        Write-Host "`nPreview only - no files were deleted. Run without -WhatIf (and confirm the prompt) to execute the purge. Deleted files go to OneDrive recycle bin (recoverable for 93 days)." -ForegroundColor Yellow
    } else {
        Write-Host "`nPurge orchestrator completed." -ForegroundColor Green
    }
    
    return
}

# Single-user mode
$sourceTenantId = $config.SourceTenantDomain
$sourceScopes = @('User.Read.All', 'Files.ReadWrite.All', 'Sites.ReadWrite.All', 'Directory.Read.All', 'Organization.Read.All')

try { Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction SilentlyContinue } catch {}

# Resilient non-interactive token cache connect (retries to avoid MSAL token file lock contention across parallel jobs)
for ($attempt = 1; $attempt -le 5; $attempt++) {
    try {
        if ($sourceTenantId) {
            Connect-MgGraph -Scopes $sourceScopes -TenantId $sourceTenantId -NoWelcome -ErrorAction Stop
        } else {
            Connect-MgGraph -Scopes $sourceScopes -NoWelcome -ErrorAction Stop
        }
        if (Get-MgContext) { break }
    } catch {
        Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 1500)
    }
}

if (-not (Get-MgContext)) {
    if ($isFilter) {
        Write-Log "Background job failed to acquire token from MSAL cache for account. Aborting thread." -Type ERROR
        return
    }
    if ($sourceTenantId) {
        Connect-MgGraph -Scopes $sourceScopes -TenantId $sourceTenantId -UseDeviceCode -NoWelcome *>&1 | Write-Host
    } else {
        Connect-MgGraph -Scopes $sourceScopes -UseDeviceCode -NoWelcome *>&1 | Write-Host
    }
}

# B1: refuse to purge against the wrong tenant.
Assert-Tenant -ExpectedTenantDomain (Get-RequiredConfigValue -Value $sourceTenantId -Name 'SourceTenantDomain')

$TotalUsers = 0
$TotalScanned = 0
$TotalDeleted = 0
$TotalProtected = 0
$TotalSpaceFreed = 0
$TotalFailed = 0

foreach ($User in $Mapping) {
    $SourceUpn = $User.SourceUPN
    $SourceId = $User.SourceId
    Write-Log "Processing user $SourceUpn"
    $TotalUsers++
    
    $driveId = Get-UserOneDriveId -UserId $SourceId
    if (-not $driveId) {
        Write-Log "  Cannot access OneDrive for $SourceUpn" -Type WARN
        continue
    }
    
    $allItems = @(Get-AllDriveItemsFast -UserId $SourceId)
    $TotalScanned += $allItems.Count
    
    $protectedFiles = $allItems | Where-Object { $_.Name -match '^\.?CHU_Y\.txt$' -and -not $_.IsFolder }
    $protectedFileIds = @{}
    foreach ($pf in $protectedFiles) { $protectedFileIds[$pf.Id] = $true }
    
    $protectedPaths = @{}
    foreach ($pf in $protectedFiles) {
        $parts = $pf.ParentPath -split '/'
        $accumulated = ''
        foreach ($part in $parts) {
            if ($part) {
                $accumulated = if ($accumulated) { "$accumulated/$part" } else { $part }
                $protectedPaths[$accumulated] = $true
            }
        }
    }
    
    $protectedFolderIds = @{}
    foreach ($item in $allItems) {
        if ($item.IsFolder) {
            $folderPath = $item.Path
            if ($protectedPaths.ContainsKey($folderPath)) {
                $protectedFolderIds[$item.Id] = $true
            }
        }
    }
    
    $rootItems = $allItems | Where-Object { -not $_.ParentPath -or $_.ParentPath -eq '' }
    $toDelete = @()
    $toSkip = @()
    foreach ($item in $rootItems) {
        if ($protectedFileIds.ContainsKey($item.Id) -or $protectedFolderIds.ContainsKey($item.Id)) {
            $toSkip += $item
        } else {
            $toDelete += $item
        }
    }
    
    foreach ($folder in ($allItems | Where-Object { $_.IsFolder -and $protectedFolderIds.ContainsKey($_.Id) })) {
        $children = $allItems | Where-Object { $_.ParentPath -eq $folder.Path }
        foreach ($child in $children) {
            if (-not $protectedFileIds.ContainsKey($child.Id) -and -not $protectedFolderIds.ContainsKey($child.Id)) {
                $toDelete += $child
            }
        }
    }
    
    # S17: delete children before parents (deepest paths first) so removing
    # a folder does not turn its children's deletes into 404s.
    $toDelete = @($toDelete | Sort-Object { ($_.Path -split '/').Count } -Descending)

    Write-Log "  Items to delete: $($toDelete.Count)"
    $TotalDeleted += $toDelete.Count
    $protCount = $protectedFileIds.Count + $protectedFolderIds.Count
    $TotalProtected += $protCount
    Write-Log "  Protected items (CHU_Y.txt + ancestors): $protCount"
    $userDeleteSize = ($toDelete | Where-Object { -not $_.IsFolder } | Measure-Object -Property Size -Sum).Sum
    $TotalSpaceFreed += $userDeleteSize
    Write-Log "  Estimated space to free: $(Format-FileSize $userDeleteSize)"
    
    # S9: standard confirmation, once per user (not once per file).
    # -WhatIf makes ShouldProcess return $false (preview); -Confirm:$false
    # (used by the orchestrator after its own confirmation) proceeds silently.
    $deletedCount = 0
    $failedCount = 0
    $doDelete = if ($toDelete.Count -gt 0) { $PSCmdlet.ShouldProcess("$SourceUpn ($($toDelete.Count) items, except CHU_Y.txt)", "Delete OneDrive items") } else { $false }
    if ($doDelete) {
        foreach ($item in $toDelete) {
            try {
                $deleteUri = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$($item.Id)"
                Invoke-MgGraphRequest -Method DELETE -Uri $deleteUri -ErrorAction Stop
                $deletedCount++
                Write-Log "    DELETED: $($item.Path)" -Type INFO
            } catch {
                # S17: a 404 here means the item is already gone (e.g. its
                # parent folder was deleted first) - not a real failure.
                if ($_.Exception.Message -match '404|NotFound|itemNotFound') {
                    $deletedCount++
                    Write-Log "    ALREADY GONE: $($item.Path)" -Type INFO
                } else {
                    $failedCount++
                    Write-Log "    FAILED: $($item.Path) - $_" -Type ERROR
                }
            }
        }
        Write-Log "  Purge complete: $deletedCount deleted, $failedCount failed."
    } else {
        Write-Log "  [WHATIF] Would delete $($toDelete.Count) items. Run without -WhatIf (and confirm the prompt) to execute."
        $toDelete | Select-Object -First 20 | ForEach-Object {
            $sizeStr = if ($_.IsFolder) { '[FOLDER]' } else { Format-FileSize $_.Size }
            Write-Log "    [WHATIF] $($_.Path) ($sizeStr)"
        }
        if ($toDelete.Count -gt 20) {
            Write-Log "    ... and $($toDelete.Count - 20) more items"
        }
    }
    $TotalFailed += $failedCount
}

if ($isFilter) {
    if ($TotalFailed -gt 0) { exit 1 }
    return
}

Write-Log "=========================================================="
Write-Log "SUMMARY REPORT"
Write-Log "Total users processed: $TotalUsers"
Write-Log "Total items scanned: $TotalScanned"
Write-Log "Total items deleted (or would-be-deleted): $TotalDeleted"
Write-Log "Total protected items: $TotalProtected"
Write-Log "Total space freed: $(Format-FileSize $TotalSpaceFreed)"
if ($WhatIfPreference) {
    Write-Log "Preview only - no files were deleted. Run without -WhatIf (and confirm the prompt) to execute purge. Deleted files go to OneDrive recycle bin (recoverable for 93 days)."
}
Write-Log "=========================================================="
if ($TotalFailed -gt 0) { exit 1 }
