<#
.SYNOPSIS
    Detects and cleans duplicate files on destination OneDrive accounts.

.DESCRIPTION
    Scans each destination user's OneDrive via Delta API to find duplicate files
    created during migration (conflict-renamed copies like "report (1).docx" and
    exact path+size duplicates). By default runs in read-only scan mode and exports
    a report. Use -AutoClean to delete duplicates (moves to OneDrive recycle bin,
    recoverable for 93 days).

    Requires (minimum Graph scopes, app-only): Files.Read.All,
    Files.ReadWrite.All (delete in -AutoClean mode only), User.Read.All.

.PARAMETER ConfigPath
    Path to MigrationConfig.psd1 configuration file.

.PARAMETER MappingPath
    Path to UserMapping.csv with SourceUPN, DestUPN, SourceId, DestId columns.

.PARAMETER AutoClean
    If specified, deletes detected duplicates (moves to OneDrive recycle bin).
    Without this switch, the script only scans and reports.

.PARAMETER UserFilter
    Optional exact-UPN filter to process only the matching user (matched against SourceUPN or DestUPN with -eq).

.EXAMPLE
    .\4c-CleanDestDuplicates.ps1
    # Scan all destination users and generate a duplicate report (read-only)

.EXAMPLE
    .\4c-CleanDestDuplicates.ps1 -AutoClean -UserFilter "john@"
    # Scan and clean duplicates for users matching "john@"
#>

[CmdletBinding()]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$MappingPath = '.\Output\UserMapping.csv',
    [switch]$AutoClean,
    [string]$UserFilter
)

# Shared safety helpers (tenant guard, app-only token, required config values).
. "$PSScriptRoot\MigrationHelpers.ps1"

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$env:AZURE_IDENTITY_DISABLE_WAM = 'true'
try {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    [System.Net.ServicePointManager]::DefaultConnectionLimit = 500
    [System.Net.ServicePointManager]::Expect100Continue = $false
    [System.Net.ServicePointManager]::UseNagleAlgorithm = $false
} catch {}
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location -Path $ScriptDir

# -- Setup Logging --------------------------------------------------------------
$LogDir = Join-Path -Path $ScriptDir -ChildPath 'Logs'
if (-not (Test-Path -Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = Join-Path -Path $LogDir -ChildPath "CleanDuplicates_${Timestamp}.log"

function Write-Log {
    param ([string]$Message, [string]$Color = "Cyan")
    $LogTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$LogTimestamp] $Message"
    Add-Content -Path $LogFile -Value $LogMessage -Encoding UTF8 -ErrorAction SilentlyContinue
    Write-Host $LogMessage -ForegroundColor $Color
}

# -- Banner ---------------------------------------------------------------------
$Mode = if ($AutoClean) { "AUTO-CLEAN MODE (duplicates WILL be deleted)" } else { "SCAN-ONLY MODE (read-only, no changes)" }
Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   OneDrive Destination Duplicate Cleaner                 " -ForegroundColor Cyan
Write-Host "   $Mode" -ForegroundColor $(if ($AutoClean) { "Red" } else { "Green" })
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

# -- Load Config ----------------------------------------------------------------
if (-not (Test-Path $ConfigPath)) {
    $ConfigPath = Join-Path $ScriptDir 'MigrationConfig.psd1'
}
$config = @{}
if (Test-Path $ConfigPath) {
    $config = Import-PowerShellDataFile -Path $ConfigPath
} else {
    Write-Log "Config file not found at $ConfigPath" "Red"
    exit
}

$TenantDomain = Get-RequiredConfigValue -Value $config.DestTenantDomain -Name 'DestTenantDomain'
# B3: no user password is used any more (app-only auth); $DefaultPassword removed.

Write-Log "Config loaded. Destination tenant: $TenantDomain"

# -- Load User Mapping ----------------------------------------------------------
if (-not (Test-Path $MappingPath)) {
    Write-Log "Mapping file not found at $MappingPath" "Red"
    exit
}
$Mapping = Import-Csv -Path $MappingPath -Encoding UTF8

# NOTE (S16): the old '-notmatch' guard for a '/goal' prefix was dead code (no UPN
# ever starts with '/goal') and has been removed. UserFilter is now an exact UPN (-eq, not regex).
$isFilter = [bool]$UserFilter
if ($isFilter) {
    $Mapping = $Mapping | Where-Object { $_.SourceUPN -eq $UserFilter -or $_.DestUPN -eq $UserFilter }
    Write-Log "Filtered to user(s): $UserFilter ($($Mapping.Count) users)" "Yellow"
}

if ($Mapping.Count -eq 0) {
    Write-Log "No users found in mapping (after filter). Nothing to do." "Yellow"
    exit
}

Write-Log "Processing $($Mapping.Count) user(s)"

# -- Multi-account parallel orchestration via background jobs -------------------
$MaxConcurrentUsers = if ($null -ne $config.MaxConcurrentUsers -and $config.MaxConcurrentUsers -gt 0) { [int]$config.MaxConcurrentUsers } else { 4 }
$MaxConcurrentUsers = [Math]::Min($MaxConcurrentUsers, $Mapping.Count)

if (-not $isFilter -and $Mapping.Count -gt 1) {
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host "   PARALLEL DUPLICATE SCAN ORCHESTRATOR" -ForegroundColor Cyan
    Write-Host "   Concurrent Threads: $MaxConcurrentUsers | Total Accounts: $($Mapping.Count)" -ForegroundColor Cyan
    Write-Host "   Mode: $Mode" -ForegroundColor $(if ($AutoClean) { "Red" } else { "Green" })
    Write-Host "==========================================================`n" -ForegroundColor Cyan

    $pendingUsers = [System.Collections.ArrayList]@($Mapping)
    $runningJobs = @{}
    $runStart = Get-Date  # S4: only merge DuplicateReport files written after this run started
    $scriptPath = $MyInvocation.MyCommand.Definition
    $completedUsers = @{}
    $failedUsers = @{}

    try {
        while ($pendingUsers.Count -gt 0 -or $runningJobs.Count -gt 0) {
            # Launch jobs up to concurrency limit
            while ($runningJobs.Count -lt $MaxConcurrentUsers -and $pendingUsers.Count -gt 0) {
                $user = $pendingUsers[0]
                $pendingUsers.RemoveAt(0)
                $upn = $user.DestUPN
                Write-Host "[LAUNCHING] Scan thread for: $upn" -ForegroundColor Green
                $job = Start-Job -Name "DUP-$upn" -ScriptBlock {
                    param($script, $upnFilter, $cleanFlag)
                    if ($cleanFlag) {
                        & $script -UserFilter $upnFilter -AutoClean
                    } else {
                        & $script -UserFilter $upnFilter
                    }
                } -ArgumentList $scriptPath, $upn, $AutoClean.IsPresent
                $runningJobs[$upn] = $job
            }

            # Check for completed jobs
            $completed = @()
            foreach ($key in $runningJobs.Keys) {
                $j = $runningJobs[$key]
                if ($j.State -in @('Completed', 'Failed', 'Stopped')) {
                    $completed += $key
                    $output = Receive-Job -Job $j 2>&1
                    if ($j.State -eq 'Completed') {
                        Write-Host "[COMPLETED] $key" -ForegroundColor Green
                        $completedUsers[$key] = $true
                    } else {
                        Write-Host "[FAILED] $key - State: $($j.State)" -ForegroundColor Red
                        $output | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkRed }
                        $failedUsers[$key] = $true
                    }
                    # Print any stdout from the job
                    $output | ForEach-Object {
                        if ($_ -is [string]) { Write-Host "   $_" -ForegroundColor DarkGray }
                    }
                    Remove-Job -Job $j -Force -ErrorAction SilentlyContinue | Out-Null
                }
            }
            foreach ($c in $completed) {
                $runningJobs.Remove($c)
            }

            # Progress display
            if ($runningJobs.Count -gt 0) {
                Write-Host "`n--- LIVE PROGRESS ($($runningJobs.Count) threads active, $($completedUsers.Count) done, $($pendingUsers.Count) pending) ---" -ForegroundColor Cyan
                foreach ($upn in $runningJobs.Keys) {
                    Write-Host "   [$upn] : scanning..." -ForegroundColor Yellow
                }
                Write-Host "--------------------------------------------------------" -ForegroundColor Cyan
                Start-Sleep -Seconds 5
            }
        }
    } finally {
        foreach ($job in $runningJobs.Values) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }
        Get-Job | Where-Object { $_.Name -like "DUP-*" } | ForEach-Object {
            Stop-Job -Job $_ -ErrorAction SilentlyContinue
            Remove-Job -Job $_ -Force -ErrorAction SilentlyContinue
        }
    }

    # Merge per-user CSV reports into a combined report
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host "   ALL THREADS COMPLETED - MERGING REPORTS" -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan

    $OutputDir = Join-Path -Path $ScriptDir -ChildPath 'Output'
    # S4: only merge reports created by THIS run (written after $runStart).
    # Per-user jobs each write their own DuplicateReport_<timestamp>.csv, so
    # globbing all history would double-count earlier runs.
    $allReportFiles = Get-ChildItem -Path $OutputDir -Filter "DuplicateReport_*.csv" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $runStart -and $_.Name -notlike 'DuplicateReport_MERGED_*' } |
        Sort-Object LastWriteTime -Descending
    $mergedDuplicates = @()
    foreach ($rf in $allReportFiles) {
        try {
            $rows = Import-Csv -Path $rf.FullName -Encoding UTF8
            $mergedDuplicates += $rows
        } catch {}
    }

    if ($mergedDuplicates.Count -gt 0) {
        $MergedPath = Join-Path $OutputDir "DuplicateReport_MERGED_${Timestamp}.csv"
        $mergedDuplicates | Export-Csv -Path $MergedPath -NoTypeInformation -Encoding UTF8
        Write-Host "`n  Total duplicates found across all users: $($mergedDuplicates.Count)" -ForegroundColor Yellow
        Write-Host "  Merged report: $MergedPath" -ForegroundColor Green
    } else {
        Write-Host "`n  No duplicates found across any user." -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "  Completed: $($completedUsers.Count) | Failed: $($failedUsers.Count) | Total: $($Mapping.Count)" -ForegroundColor Cyan
    if (-not $AutoClean -and $mergedDuplicates.Count -gt 0) {
        Write-Host "  Run with -AutoClean to remove duplicates." -ForegroundColor Yellow
        Write-Host "  Files will be moved to OneDrive recycle bin (recoverable for 93 days)." -ForegroundColor Yellow
    }
    Write-Host ""
    return
}

# -- Token Cache ----------------------------------------------------------------
$TokenCache = @{}

function Get-UserGraphToken {
    # B3: app-only client-credentials token, in-memory cache only.
    # $UPN is only the cache key / log label: an app-only token is tenant-wide,
    # not a per-user token, so callers must use /users/{UPN}/... endpoints.
    param ([string]$UPN, [switch]$ForceRefresh)

    if (-not $ForceRefresh -and $TokenCache.ContainsKey($UPN)) {
        $cached = $TokenCache[$UPN]
        if ([DateTime]::UtcNow.AddMinutes(10) -lt $cached.Expires.ToUniversalTime()) {
            return $cached.Token
        }
    }

    try {
        $destClientId = Get-RequiredConfigValue -Value $config.DestClientId -Name 'DestClientId'
        $destClientSecret = Get-RequiredConfigValue -Value $config.DestClientSecret -Name 'DestClientSecret'
        $token = Get-AppOnlyGraphAccessToken -TenantDomain $TenantDomain -ClientId $destClientId -ClientSecret $destClientSecret
    } catch {
        Write-Log "  Failed to get app-only token for tenant $TenantDomain : $_" "Red"
        return $null
    }
    # Get-AppOnlyGraphAccessToken returns only the token string; Graph app-only
    # tokens live ~1h, so cache it for 50 minutes.
    $expiresDt = [DateTime]::UtcNow.AddMinutes(50)
    $TokenCache[$UPN] = @{
        Token   = $token
        Expires = $expiresDt
    }
    return $token
}

# -- Graph API Helper with auto token refresh ----------------------------------
function Invoke-GraphRequest {
    param (
        [string]$UPN,
        [string]$Method = 'GET',
        [string]$Uri,
        [hashtable]$Headers = @{}
    )

    $token = Get-UserGraphToken -UPN $UPN
    if (-not $token) { return $null }

    $Headers['Authorization'] = "Bearer $token"

    try {
        return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -ErrorAction Stop
    } catch {
        if ($_.Exception.Message -match '401|Unauthorized') {
            Write-Log "  [TOKEN EXPIRED] Refreshing token for $UPN..." "Yellow"
            $token = Get-UserGraphToken -UPN $UPN -ForceRefresh
            if ($token) {
                $Headers['Authorization'] = "Bearer $token"
                return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -ErrorAction Stop
            }
        }
        throw
    }
}

# -- Scan Drive Items via Delta API ---------------------------------------------
function Get-AllDriveItems {
    param (
        [string]$UPN,
        [string]$DriveId
    )

    $items = [System.Collections.Generic.List[PSCustomObject]]::new()
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root/delta"
    $pageCount = 0

    while ($uri) {
        $pageCount++
        if ($pageCount -ge 1500) {
            Write-Log "  [WARN] Reached safety limit of 1500 delta pages ($($items.Count) items found)" "Yellow"
            break
        }

        $token = Get-UserGraphToken -UPN $UPN
        if (-not $token) { break }
        $headers = @{ Authorization = "Bearer $token" }

        try {
            $response = Invoke-RestMethod -Uri $uri -Headers $headers -ErrorAction Stop
        } catch {
            if ($_.Exception.Message -match '401|Unauthorized') {
                $token = Get-UserGraphToken -UPN $UPN -ForceRefresh
                if ($token) {
                    $headers = @{ Authorization = "Bearer $token" }
                    $response = Invoke-RestMethod -Uri $uri -Headers $headers -ErrorAction Stop
                } else { break }
            } else { throw }
        }

        foreach ($child in $response.value) {
            $name = $child.name
            if (-not $name -or $name -eq 'root') { continue }
            if ($child.deleted) { continue }

            # Skip folders - we only care about files for duplicate detection
            $isFolder = ($null -ne $child.folder) -and ($name -notmatch '\.(one|onetoc2)$')
            if ($child.package -and $child.package.type -eq 'oneNote') { $isFolder = $false }
            if ($isFolder) { continue }

            $ParentPath = ""
            if ($child.parentReference -and $child.parentReference.path) {
                $ParentPath = $child.parentReference.path -replace '^.*root:?/?', ''
                $ParentPath = $ParentPath -replace '^/', '' -replace '/$', ''
            }

            $itemPath = if ($ParentPath) { "$ParentPath/$name" } else { $name }

            $itemObj = [PSCustomObject]@{
                Id         = $child.id
                Name       = $name
                Path       = $itemPath
                Size       = [long]$child.size
                ParentPath = $ParentPath
            }
            $items.Add($itemObj)
        }

        if ($pageCount % 10 -eq 0) {
            Write-Host "      Page $pageCount done, $($items.Count) files found so far..." -ForegroundColor DarkGray
        }

        # Next page or stop
        if ($response.'@odata.nextLink') {
            $uri = $response.'@odata.nextLink'
        } else {
            $uri = $null
        }
    }

    return $items
}

# -- Duplicate Detection -------------------------------------------------------
function Find-Duplicates {
    param (
        [System.Collections.Generic.List[PSCustomObject]]$Items
    )

    $duplicates = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Build lookup: ParentPath -> list of items for fast matching
    $folderIndex = @{}
    foreach ($item in $Items) {
        $key = $item.ParentPath
        if (-not $folderIndex.ContainsKey($key)) {
            $folderIndex[$key] = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        $folderIndex[$key].Add($item)
    }

    # -- Strategy A: Conflict-renamed duplicates --------------------------------
    # Pattern: "filename (1).ext", "report (2).docx", "photo (3).jpg"
    $conflictPattern = '^(.+?)\s+\(\d+\)(\.[^.]+)?$'

    foreach ($item in $Items) {
        if ($item.Name -match $conflictPattern) {
            $baseName = $Matches[1]
            $extension = if ($Matches[2]) { $Matches[2] } else { '' }
            $originalName = "${baseName}${extension}"

            # Check if a file with the original name exists in the same parent folder
            $siblings = $folderIndex[$item.ParentPath]
            if (-not $siblings) { continue }

            $original = $siblings | Where-Object { $_.Name -eq $originalName -and $_.Id -ne $item.Id } | Select-Object -First 1

            if ($original) {
                # S4: AutoClean still requires the SAME SIZE - a conflict-renamed
                # file with a different size is a different file, not a duplicate.
                if ($original.Size -eq $item.Size) {
                    $duplicates.Add([PSCustomObject]@{
                        DuplicateId   = $item.Id
                        DuplicateName = $item.Name
                        DuplicatePath = $item.Path
                        DuplicateSize = $item.Size
                        OriginalName  = $original.Name
                        OriginalPath  = $original.Path
                        OriginalSize  = $original.Size
                        Strategy      = 'A-ConflictRenamed'
                    })
                }
            }
        }
    }

    # Second strategy removed (S4): dead code - two files cannot share a name in the
    # same OneDrive folder, so grouping by ParentPath+Name+Size can never match.
    return $duplicates
}

# -- Format file size for display -----------------------------------------------
function Format-FileSize {
    param ([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

# ==============================================================================
#  MAIN PROCESSING LOOP
# ==============================================================================

$OutputDir = Join-Path -Path $ScriptDir -ChildPath 'Output'
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

$AllDuplicates = [System.Collections.Generic.List[PSCustomObject]]::new()
$TotalFilesScanned = 0
$TotalUsersProcessed = 0
$TotalUsersSkipped = 0

foreach ($User in $Mapping) {
    $DestUPN = $User.DestUPN
    $SourceUPN = $User.SourceUPN

    Write-Host ""
    Write-Host "----------------------------------------------------------" -ForegroundColor DarkGray
    Write-Log "Processing user: $DestUPN"

    # Assert the destination UPN belongs to this migration's dest tenant
    # (DestTenantDomain or NewDomain) before scanning/deleting anything in it.
    $destUpnDomain = ($DestUPN -split '@')[-1]
    $allowedDestDomains = @($TenantDomain, $config.NewDomain) | Where-Object { $_ }
    if ($destUpnDomain -notin $allowedDestDomains) {
        Write-Log "  Skipping $DestUPN - UPN domain '$destUpnDomain' is not in the destination tenant ($($allowedDestDomains -join ', '))." "Yellow"
        $TotalUsersSkipped++
        continue
    }

    # -- Step 1: Authenticate ---------------------------------------------------
    try {
        $token = Get-UserGraphToken -UPN $DestUPN
        if (-not $token) {
            Write-Log "  Skipping $DestUPN - authentication failed." "Red"
            $TotalUsersSkipped++
            continue
        }
        Write-Log "  Authenticated successfully." "Green"
    } catch {
        Write-Log "  Skipping $DestUPN - authentication error: $_" "Red"
        $TotalUsersSkipped++
        continue
    }

    # -- Step 2: Get DriveId ----------------------------------------------------
    # App-only token: /users/{UPN}/drives, never the per-user me endpoint.
    $escapedDestUpn = [System.Uri]::EscapeDataString($DestUPN)
    $DriveId = $null
    try {
        $headers = @{ Authorization = "Bearer $token" }
        $drivesResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$escapedDestUpn/drives" -Headers $headers -ErrorAction Stop
        if ($drivesResponse.value.Count -gt 0) {
            $DriveId = $drivesResponse.value[0].id
        }
    } catch {
        if ($_.Exception.Message -match '401|Unauthorized') {
            $token = Get-UserGraphToken -UPN $DestUPN -ForceRefresh
            if ($token) {
                $headers = @{ Authorization = "Bearer $token" }
                try {
                    $drivesResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$escapedDestUpn/drives" -Headers $headers -ErrorAction Stop
                    if ($drivesResponse.value.Count -gt 0) {
                        $DriveId = $drivesResponse.value[0].id
                    }
                } catch {}
            }
        }
    }

    if (-not $DriveId) {
        Write-Log "  Skipping $DestUPN - OneDrive not provisioned or inaccessible." "Yellow"
        $TotalUsersSkipped++
        continue
    }
    Write-Log "  DriveId: $DriveId"

    # -- Step 3: Scan all files via Delta API -----------------------------------
    Write-Log "  Scanning files via Delta API..."
    try {
        $items = Get-AllDriveItems -UPN $DestUPN -DriveId $DriveId
    } catch {
        Write-Log "  Skipping $DestUPN - Delta scan failed: $_" "Red"
        $TotalUsersSkipped++
        continue
    }

    $fileCount = $items.Count
    $TotalFilesScanned += $fileCount
    Write-Log "  Found $fileCount files." "Green"

    if ($fileCount -eq 0) {
        $TotalUsersProcessed++
        continue
    }

    # -- Step 4: Detect duplicates ----------------------------------------------
    Write-Log "  Analyzing for duplicates..."
    $userDuplicates = Find-Duplicates -Items $items

    if ($userDuplicates.Count -eq 0) {
        Write-Log "  No duplicates found." "Green"
        $TotalUsersProcessed++
        continue
    }

    $dupSize = ($userDuplicates | Measure-Object -Property DuplicateSize -Sum).Sum
    Write-Log "  Found $($userDuplicates.Count) duplicate(s) totaling $(Format-FileSize $dupSize)" "Yellow"

    # Display per-user duplicate table
    Write-Host ""
    $userDuplicates | Format-Table -Property @(
        @{ Label = 'Duplicate'; Expression = { $_.DuplicatePath }; Width = 50 },
        @{ Label = 'Original'; Expression = { $_.OriginalPath }; Width = 50 },
        @{ Label = 'Size'; Expression = { Format-FileSize $_.DuplicateSize }; Width = 12 },
        @{ Label = 'Strategy'; Expression = { $_.Strategy }; Width = 18 }
    ) -AutoSize -Wrap | Out-String | Write-Host -ForegroundColor DarkYellow

    # Add user context to each duplicate for the combined report
    foreach ($dup in $userDuplicates) {
        $AllDuplicates.Add([PSCustomObject]@{
            User          = $DestUPN
            DuplicateId   = $dup.DuplicateId
            DuplicatePath = $dup.DuplicatePath
            DuplicateSize = $dup.DuplicateSize
            OriginalPath  = $dup.OriginalPath
            OriginalSize  = $dup.OriginalSize
            Strategy      = $dup.Strategy
            DriveId       = $DriveId
        })
    }

    # -- Step 5: Clean duplicates (if -AutoClean) -------------------------------
    if ($AutoClean) {
        Write-Log "  Cleaning $($userDuplicates.Count) duplicate(s)..." "Red"
        $deletedCount = 0
        $failedCount = 0

        foreach ($dup in $userDuplicates) {
            $deleteUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($dup.DuplicateId)"

            try {
                $delToken = Get-UserGraphToken -UPN $DestUPN
                if (-not $delToken) {
                    Write-Log "    FAILED (no token): $($dup.DuplicatePath)" "Red"
                    $failedCount++
                    continue
                }
                $delHeaders = @{ Authorization = "Bearer $delToken" }
                Invoke-RestMethod -Uri $deleteUri -Method DELETE -Headers $delHeaders -ErrorAction Stop
                $deletedCount++
                Write-Log "    DELETED: $($dup.DuplicatePath)" "Green"
            } catch {
                if ($_.Exception.Message -match '401|Unauthorized') {
                    # Retry with refreshed token
                    try {
                        $delToken = Get-UserGraphToken -UPN $DestUPN -ForceRefresh
                        if ($delToken) {
                            $delHeaders = @{ Authorization = "Bearer $delToken" }
                            Invoke-RestMethod -Uri $deleteUri -Method DELETE -Headers $delHeaders -ErrorAction Stop
                            $deletedCount++
                            Write-Log "    DELETED (after refresh): $($dup.DuplicatePath)" "Green"
                        } else {
                            $failedCount++
                            Write-Log "    FAILED (auth): $($dup.DuplicatePath)" "Red"
                        }
                    } catch {
                        $failedCount++
                        Write-Log "    FAILED: $($dup.DuplicatePath) - $_" "Red"
                    }
                } else {
                    $failedCount++
                    Write-Log "    FAILED: $($dup.DuplicatePath) - $_" "Red"
                }
            }
        }

        Write-Log "  Cleanup complete: $deletedCount deleted, $failedCount failed." $(if ($failedCount -gt 0) { "Yellow" } else { "Green" })
    }

    $TotalUsersProcessed++
}

# ==============================================================================
#  SUMMARY & REPORT
# ==============================================================================

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   DUPLICATE SCAN SUMMARY                                 " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

$TotalDupSize = if ($AllDuplicates.Count -gt 0) { ($AllDuplicates | Measure-Object -Property DuplicateSize -Sum).Sum } else { 0 }

Write-Host ""
Write-Host "  Users processed:       $TotalUsersProcessed" -ForegroundColor Cyan
Write-Host "  Users skipped:         $TotalUsersSkipped" -ForegroundColor $(if ($TotalUsersSkipped -gt 0) { "Yellow" } else { "Cyan" })
Write-Host "  Total files scanned:   $TotalFilesScanned" -ForegroundColor Cyan
Write-Host "  Duplicates found:      $($AllDuplicates.Count)" -ForegroundColor $(if ($AllDuplicates.Count -gt 0) { "Yellow" } else { "Green" })
Write-Host "  Recoverable space:     $(Format-FileSize $TotalDupSize)" -ForegroundColor $(if ($TotalDupSize -gt 0) { "Yellow" } else { "Green" })

$stratA = @($AllDuplicates | Where-Object { $_.Strategy -eq 'A-ConflictRenamed' }).Count
if ($AllDuplicates.Count -gt 0) {
    Write-Host "    Strategy A (conflict-renamed): $stratA" -ForegroundColor DarkYellow
}

# Export report CSV
if ($AllDuplicates.Count -gt 0) {
    $ReportPath = Join-Path $OutputDir "DuplicateReport_${Timestamp}.csv"
    $AllDuplicates | Select-Object User, DuplicatePath, OriginalPath, DuplicateSize, OriginalSize, Strategy |
        Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Log "Report exported to: $ReportPath" "Green"
}

# Final message
Write-Host ""
if (-not $AutoClean) {
    if ($AllDuplicates.Count -gt 0) {
        Write-Host "  Run with -AutoClean to remove duplicates." -ForegroundColor Yellow
        Write-Host "  Files will be moved to OneDrive recycle bin (recoverable for 93 days)." -ForegroundColor Yellow
    } else {
        Write-Host "  No duplicates found. Your destination OneDrive accounts are clean!" -ForegroundColor Green
    }
} else {
    Write-Host "  Auto-clean completed. Deleted files are in OneDrive recycle bin (recoverable for 93 days)." -ForegroundColor Green
}

Write-Host ""
Write-Log "Script completed." "Green"
