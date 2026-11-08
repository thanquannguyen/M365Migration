<#
.SYNOPSIS
Parallel OneDrive Upload Worker using app-only authentication (client credentials).

.DESCRIPTION
Authenticates with the destination tenant app registration (DestClientId /
DestClientSecret from config) and uploads staged files directly to each
destination user's drive via /users/{UPN}/drive endpoints. No user password
is used and no token is written to disk (B3/B4): tokens live in memory only.
Auto-hydrates a destination user's OneDrive if needed. This avoids all WAM
token cache issues, admin permission errors, and language-dependent
/drive/root shortcuts.

Requires (minimum Graph scopes, app-only): Files.ReadWrite.All, User.Read.All.
SharePoint hydration also uses client credentials against
https://<tenant>-my.sharepoint.com/.default.
#>

[CmdletBinding()]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$MappingPath = '.\Output\UserMapping.csv',
    [string]$UserFilter = ''
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

if (-not (Test-Path $MappingPath)) {
    Write-Error "Mapping file not found at $MappingPath"
    exit
}
$Mapping = Import-Csv -Path $MappingPath -Encoding UTF8
# NOTE (S16): the old '-notmatch' guard for a '/goal' prefix was dead code (no UPN
# ever starts with '/goal') and has been removed. UserFilter is now an exact UPN (-eq, not regex).
$isFilter = [bool]$UserFilter
if ($isFilter) {
    $Mapping = $Mapping | Where-Object { $_.SourceUPN -eq $UserFilter -or $_.DestUPN -eq $UserFilter }
    Write-Host "Filtered to user(s): $UserFilter ($($Mapping.Count) users)" -ForegroundColor Cyan
}

if (-not (Test-Path $ConfigPath)) {
    $ConfigPath = Join-Path $ScriptDir 'MigrationConfig.psd1'
}
$config = @{}
if (Test-Path $ConfigPath) {
    $config = Import-PowerShellDataFile -Path $ConfigPath
}

# Multi-account parallel upload orchestration via background jobs
$MaxConcurrentUsers = if ($null -ne $config.MaxConcurrentUsers -and $config.MaxConcurrentUsers -gt 0) { [int]$config.MaxConcurrentUsers } else { 1 }
$MaxConcurrentUsers = [Math]::Min($MaxConcurrentUsers, $Mapping.Count)
if (-not $isFilter -and $Mapping.Count -gt 0) {
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host "   MULTI-ACCOUNT PARALLEL ONEDRIVE UPLOAD ORCHESTRATOR" -ForegroundColor Cyan
    Write-Host "   Concurrent Threads: $MaxConcurrentUsers | Total Accounts: $($Mapping.Count)" -ForegroundColor Cyan
    Write-Host "==========================================================`n" -ForegroundColor Cyan
    
    $pendingUsers = [System.Collections.ArrayList]@($Mapping)
    $runningJobs = @{}
    $uploadRetryCounts = @{}  # S2: Failed/Stopped re-queues per DestUPN, capped at 3
    $scriptPath = $MyInvocation.MyCommand.Definition

    function Get-UserUploadStatus {
        param($User)
        $srcUpn = $User.SourceUPN
        $manPath = Join-Path -Path $ScriptDir -ChildPath "MigrationStaging\$srcUpn\manifest.json"
        $upLogPath = Join-Path -Path $ScriptDir -ChildPath "MigrationStaging\$srcUpn\uploaded.log"
        
        if (-not (Test-Path $manPath)) { return @{ HasFiles = $false; IsDone = $false } }
        
        try {
            $manifestObj = Get-Content $manPath -Raw -ErrorAction Stop | ConvertFrom-Json
            $props = @($manifestObj.PSObject.Properties)
            if ($props.Count -eq 0) { return @{ HasFiles = $false; IsDone = $false } }
            
            $uploadedIds = @{}
            if (Test-Path $upLogPath) {
                Get-Content $upLogPath -Encoding UTF8 | Where-Object { $_.Trim() } | ForEach-Object { $uploadedIds[$_.Trim()] = $true }
            }
            
            $readyCount = 0
            $uncompletedCount = 0
            foreach ($p in $props) {
                $item = $p.Value
                $isSkipped = if ($null -ne $item.Skipped) { [bool]$item.Skipped } else { $false }
                if (-not $isSkipped -and -not $uploadedIds.ContainsKey($item.Id)) {
                    $uncompletedCount++
                    if ($item.Downloaded -eq $true) {
                        $readyCount++
                    }
                }
            }
            
            return @{ HasFiles = ($readyCount -gt 0); IsDone = ($uncompletedCount -eq 0) }
        } catch {
            return @{ HasFiles = $false; IsDone = $false }
        }
    }

    try {
        while ($pendingUsers.Count -gt 0 -or $runningJobs.Count -gt 0) {
            # 1. Prune any pending users who are already 100% uploaded
            for ($i = $pendingUsers.Count - 1; $i -ge 0; $i--) {
                $status = Get-UserUploadStatus $pendingUsers[$i]
                if ($status.IsDone) {
                    Write-Host "[ALREADY UPLOADED] Account: $($pendingUsers[$i].DestUPN)" -ForegroundColor DarkGreen
                    $pendingUsers.RemoveAt($i)
                }
            }

            # 2. Dynamically pick pending users who have downloaded files ready on disk
            while ($runningJobs.Count -lt $MaxConcurrentUsers -and $pendingUsers.Count -gt 0) {
                $readyIndex = -1
                for ($i = 0; $i -lt $pendingUsers.Count; $i++) {
                    $status = Get-UserUploadStatus $pendingUsers[$i]
                    if ($status.HasFiles) {
                        $readyIndex = $i
                        break
                    }
                }
                
                # If no pending user has downloaded files ready yet, wait for download phase
                if ($readyIndex -eq -1) { break }
                
                $user = $pendingUsers[$readyIndex]
                $pendingUsers.RemoveAt($readyIndex)
                $upn = $user.DestUPN
                Write-Host "[LAUNCHING] Upload thread for account: $upn (files ready on disk)" -ForegroundColor Green
                $job = Start-Job -Name "UP-$upn" -ScriptBlock {
                    param($script, $upnFilter)
                    & $script -UserFilter $upnFilter
                } -ArgumentList $scriptPath, $upn
                $runningJobs[$upn] = $job
            }
            
            $completed = @()
            foreach ($key in $runningJobs.Keys) {
                $j = $runningJobs[$key]
                if ($j.State -in @('Completed', 'Failed', 'Stopped')) {
                    $completed += $key
                    $output = Receive-Job -Job $j 2>&1
                    if ($j.State -eq 'Completed') {
                        $mappingRow = $Mapping | Where-Object { $_.DestUPN -eq $key } | Select-Object -First 1
                        $status = if ($mappingRow) { Get-UserUploadStatus $mappingRow } else { @{ IsDone = $true } }
                        if ($status.IsDone) {
                            Write-Host "[ALL FILES UPLOADED - COMPLETE] Account: $key" -ForegroundColor Green
                        } else {
                            Write-Host "[BATCH UPLOADED] Account: $key - waiting for more files to download..." -ForegroundColor Yellow
                            # NOTE (S2): this re-queue is not a failure retry - it waits for
                            # more files to finish downloading - so it is intentionally not capped.
                            if ($mappingRow) { [void]$pendingUsers.Add($mappingRow) }
                        }
                    } else {
                        Write-Host "[FAILED/STOPPED] Account: $key - State: $($j.State)" -ForegroundColor Red
                        $output | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkRed }
                        # S2: re-queue a failed/stopped user at most 3 times, then give up.
                        $uploadRetryCounts[$key] = ([int]$uploadRetryCounts[$key]) + 1
                        if ($uploadRetryCounts[$key] -le 3) {
                            Write-Host "[RE-QUEUEING] Account: $key - retry $($uploadRetryCounts[$key])/3..." -ForegroundColor Yellow
                            $mappingRow = $Mapping | Where-Object { $_.DestUPN -eq $key } | Select-Object -First 1
                            if ($mappingRow) { [void]$pendingUsers.Add($mappingRow) }
                        } else {
                            Write-Host "[FAILED] Account: $key - exceeded 3 retries, giving up (recorded as failed)." -ForegroundColor Red
                        }
                    }
                    Remove-Job -Job $j -Force -ErrorAction SilentlyContinue | Out-Null
                }
            }
            foreach ($c in $completed) {
                $runningJobs.Remove($c)
            }
            
            if ($runningJobs.Count -gt 0) {
                Write-Host "`n--- LIVE UPLOAD PROGRESS ($($runningJobs.Count) accounts active) ---" -ForegroundColor Cyan
                foreach ($destUpn in $runningJobs.Keys) {
                    $mappingRow = $Mapping | Where-Object { $_.DestUPN -eq $destUpn } | Select-Object -First 1
                    if ($mappingRow) {
                        $srcUpn = $mappingRow.SourceUPN
                        $manPath = Join-Path -Path $ScriptDir -ChildPath "MigrationStaging\$srcUpn\manifest.json"
                        $upLogPath = Join-Path -Path $ScriptDir -ChildPath "MigrationStaging\$srcUpn\uploaded.log"
                        $upCount = 0
                        if (Test-Path $upLogPath) {
                            $upCount = @(Get-Content $upLogPath -Encoding UTF8 | Where-Object { $_.Trim() } | Select-Object -Unique).Count
                        }
                        $total = 0
                        if (Test-Path $manPath) {
                            try {
                                $content = Get-Content $manPath -Raw -ErrorAction Stop | ConvertFrom-Json
                                $total = @($content.PSObject.Properties).Count
                            } catch {}
                        }
                        $displayTotal = [Math]::Max($total, $upCount)
                        Write-Host "   [$destUpn] : $upCount / $displayTotal files uploaded" -ForegroundColor Green
                    }
                }
                Write-Host "--------------------------------------------------------" -ForegroundColor Cyan
                Start-Sleep -Seconds 5
            }
        }
        Write-Host "`n[ALL ACCOUNTS UPLOAD FINISHED]" -ForegroundColor Green
    } finally {
        foreach ($job in $runningJobs.Values) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }
        Get-Job | Where-Object { $_.Name -like "UP-*" } | ForEach-Object { Stop-Job -Job $_ -ErrorAction SilentlyContinue; Remove-Job -Job $_ -Force -ErrorAction SilentlyContinue }
    }
    return
}

$TenantDomain = Get-RequiredConfigValue -Value $config.DestTenantDomain -Name 'DestTenantDomain'
$TenantName = $TenantDomain.Split('.')[0]
# B3/B4: no user password is used any more (app-only auth), so $DefaultPassword is gone.

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   OneDrive Parallel Upload Worker (Direct Auto-Login)    " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "No browser or Device Code required! Authenticating with the destination app registration (client credentials).`n" -ForegroundColor Green

$StagingDir = Join-Path -Path $ScriptDir -ChildPath 'MigrationStaging'
if (-not (Test-Path $StagingDir)) { New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null }

$TokenCache = @{}     # map TargetUPN -> @{ Token = "..."; Expires = ... }
$FolderCaches = @{}   # map TargetUPN -> hashtable of folder paths to FolderIds

function Get-UserGraphToken {
    # B3: app-only client-credentials token, in-memory cache only (B4 - no
    # on-disk token cache file, no working password stored anywhere).
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
        Write-Error "Failed to get app-only token for tenant $TenantDomain : $_"
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

function Ensure-OneDriveProvisioned {
    param ([string]$UPN)

    $token = Get-UserGraphToken -UPN $UPN
    if (-not $token) { return $null }
    $headers = @{ Authorization = "Bearer $token" }
    $escapedUpn = [System.Uri]::EscapeDataString($UPN)

    # 1. Query /users/{UPN}/drives with automatic 401 token refresh.
    # (App-only tokens have no per-user me identity, so the me endpoint must not be used.)
    $drives = $null
    try {
        $drives = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$escapedUpn/drives" -Headers $headers -ErrorAction Stop
    } catch {
        if ($_.Exception.Message -match '401|Unauthorized') {
            Write-Host "   [TOKEN EXPIRED] Refreshing token for $UPN..." -ForegroundColor Yellow
            $token = Get-UserGraphToken -UPN $UPN -ForceRefresh
            if ($token) {
                $headers = @{ Authorization = "Bearer $token" }
                try { $drives = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$escapedUpn/drives" -Headers $headers -ErrorAction Stop } catch {}
            }
        }
    }
    if ($drives -and $drives.value.Count -gt 0) {
        return $drives.value[0].id
    }

    # 2. Trigger hydration via SharePoint Online REST API if drive not ready.
    # Client credentials for the SPO resource (B3) - no user password, and the
    # host is built from this migration's tenant (B5), never hard-coded.
    Write-Host "   [AUTO-PROVISIONING] Hydrating OneDrive for ($UPN)..." -ForegroundColor Yellow
    try {
        $spoBody = @{
            client_id     = (Get-RequiredConfigValue -Value $config.DestClientId -Name 'DestClientId')
            client_secret = (Get-RequiredConfigValue -Value $config.DestClientSecret -Name 'DestClientSecret')
            scope         = "https://$TenantName-my.sharepoint.com/.default"
            grant_type    = 'client_credentials'
        }
        $spoToken = (Invoke-RestMethod -Uri "https://login.microsoftonline.com/$([System.Uri]::EscapeDataString($TenantDomain))/oauth2/v2.0/token" -Method POST -Body $spoBody -ErrorAction Stop).access_token
        $siteName = $UPN -replace '[@.]', '_'
        $spoHeaders = @{ Authorization = "Bearer $spoToken"; Accept = "application/json;odata=verbose" }
        $null = Invoke-RestMethod -Uri "https://$TenantName-my.sharepoint.com/personal/$siteName/_api/web/lists" -Headers $spoHeaders -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
    } catch {
        # ignore
    }

    # Retry query /users/{UPN}/drives
    try {
        $drives = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$escapedUpn/drives" -Headers $headers -ErrorAction Stop
    } catch {}
    if ($drives -and $drives.value.Count -gt 0) {
        return $drives.value[0].id
    }
    return $null
}

function New-DriveFolderDirect {
    param ([string]$Token, [string]$DriveId, [string]$ParentItemId, [string]$FolderName)
    
    $headers = @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json; charset=utf-8" }
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ParentItemId/children"
    $body = @{
        name = $FolderName
        folder = @{}
        "@microsoft.graph.conflictBehavior" = "replace"
    } | ConvertTo-Json -Depth 5
    $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($body)
    
    # First try to create directly (conflict=replace means it returns existing folder if present)
    try {
        $folder = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $bodyBytes -ErrorAction Stop
        return $folder
    } catch {
        # If 409 Conflict or nameAlreadyExists, list children to find the existing folder
        $errMsg = $_.ToString()
        if ($errMsg -match '409' -or $errMsg -match 'nameAlreadyExists') {
            try {
                $listUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ParentItemId/children"
                $children = Invoke-RestMethod -Method GET -Uri $listUri -Headers @{ Authorization = "Bearer $Token" } -ErrorAction Stop
                $match = $children.value | Where-Object { $_.name -eq $FolderName } | Select-Object -First 1
                if ($match) { return $match }
            } catch {}
        }
        Write-Warning "Failed to create folder $FolderName : $_"
        return $null
    }
}

function Upload-DriveItemDirect {
    param ([string]$Token, [string]$DriveId, [string]$ParentItemId, [string]$FilePath, [string]$FileName, [string]$UPN)
    
    $fileInfo = Get-Item -Path $FilePath
    $fileSize = $fileInfo.Length
    $fourMB = 4 * 1024 * 1024
    $ChunkSize = 10 * 1024 * 1024
    $encodedFileName = [System.Uri]::EscapeDataString($FileName)
    
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            if ($fileSize -le $fourMB) {
                $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
                $uri = "https://graph.microsoft.com/v1.0/drives/${DriveId}/items/${ParentItemId}:/${encodedFileName}:/content"
                $headers = @{ Authorization = "Bearer $Token"; "Content-Type" = "application/octet-stream" }
                $res = Invoke-RestMethod -Method PUT -Uri $uri -Headers $headers -Body $fileBytes -ErrorAction Stop
                return $res
            } else {
                $uri = "https://graph.microsoft.com/v1.0/drives/${DriveId}/items/${ParentItemId}:/${encodedFileName}:/createUploadSession"
                $headers = @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json; charset=utf-8" }
                $body = @{ item = @{ "@microsoft.graph.conflictBehavior" = "replace" } } | ConvertTo-Json -Depth 5
                $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($body)
                
                $session = Invoke-RestMethod -Method POST -Uri $uri -Headers $headers -Body $bodyBytes -ErrorAction Stop
                $uploadUrl = $session.uploadUrl
                
                $fileStream = [System.IO.File]::OpenRead($FilePath)
                $buffer = New-Object byte[] $ChunkSize
                $bytesRead = 0
                $startRange = 0
                
                while (($bytesRead = $fileStream.Read($buffer, 0, $ChunkSize)) -gt 0) {
                    $endRange = $startRange + $bytesRead - 1
                    $contentRange = "bytes $startRange-$endRange/$fileSize"
                    $chunk = $buffer
                    if ($bytesRead -lt $ChunkSize) {
                        $chunk = New-Object byte[] $bytesRead
                        [System.Array]::Copy($buffer, $chunk, $bytesRead)
                    }
                    $chunkHeaders = @{
                        "Content-Length" = $bytesRead.ToString()
                        "Content-Range"  = $contentRange
                    }
                    $chunkUploaded = $false
                    $chunkAttempt = 0
                    while (-not $chunkUploaded -and $chunkAttempt -lt 5) {
                        $chunkAttempt++
                        try {
                            $uploadRes = Invoke-WebRequest -Method PUT -Uri $uploadUrl -Body $chunk -Headers $chunkHeaders -UseBasicParsing -TimeoutSec 3600 -ErrorAction Stop
                            $chunkUploaded = $true
                        } catch {
                            if ($chunkAttempt -lt 5) {
                                Write-Host " [RETRY CHUNK] Network drop detected on chunk ($contentRange), retrying in 5s... ($chunkAttempt/5)" -ForegroundColor Yellow
                                Start-Sleep -Seconds 5
                            } else {
                                throw $_
                            }
                        }
                    }
                    $startRange += $bytesRead
                }
                $fileStream.Close()
                return $uploadRes
            }
        } catch {
            if ($null -ne $fileStream) { try { $fileStream.Close() } catch {} }
            if ($_.Exception.Message -match 'another process|access the file|being used') {
                return $null # skip silently, file is still being downloaded or locked
            }
            if ($_.Exception.Message -match '401|Unauthorized' -and $UPN) {
                Write-Host "   [TOKEN EXPIRED DURING UPLOAD] Automatically refreshing token for $UPN..." -ForegroundColor Yellow
                $Token = Get-UserGraphToken -UPN $UPN -ForceRefresh
                Start-Sleep -Seconds 1
                continue
            }
            Write-Warning "Attempt $attempt failed to upload ${FileName}: $_"
            Start-Sleep -Seconds 2
        }
    }
    return $null
}

# S11: same sanitization as Download-DriveItem in 4a-OneDriveHelpers.ps1 - the file
# on disk lives at the sanitized path, so the worker must check that path, not the
# raw manifest path. Keep this function in sync with the 4a copy.
function Sanitize-LocalPath([string]$path) {
    if (-not $path) { return $path }
    $prefix = ""
    $main = $path
    if ($path.StartsWith('\\?\')) {
        $prefix = '\\?\'
        $main = $path.Substring(4)
    } elseif ($path.Length -ge 3 -and $path[1] -eq ':' -and $path[2] -eq '\') {
        $prefix = $path.Substring(0, 3)
        $main = $path.Substring(3)
    }
    $parts = $main -split '[\\/]'
    for ($i = 0; $i -lt $parts.Count; $i++) {
        $parts[$i] = $parts[$i] -replace '[?*:"<>|]', '_'
        $parts[$i] = $parts[$i].TrimEnd('. ')
    }
    return $prefix + ($parts -join '\')
}

Write-Host "[WATCHER STARTED] Watching for downloaded files to upload & auto-delete..." -ForegroundColor Cyan
Write-Host "Press Ctrl+C to stop.`n" -ForegroundColor DarkGray

while ($true) {
    $uploadedInLoop = 0
    
    foreach ($User in $Mapping) {
        $SourceUpn = $User.SourceUPN
        $TargetUpn = $User.DestUPN

        # Assert the destination UPN belongs to this migration's dest tenant
        # (DestTenantDomain or NewDomain) before uploading anything to it.
        $destUpnDomain = ($TargetUpn -split '@')[-1]
        $allowedDestDomains = @($TenantDomain, $config.NewDomain) | Where-Object { $_ }
        if ($destUpnDomain -notin $allowedDestDomains) {
            Write-Host " [SKIP] $TargetUpn : UPN domain '$destUpnDomain' is not in the destination tenant ($($allowedDestDomains -join ', '))." -ForegroundColor Yellow
            continue
        }

        $UserStagingDir = Join-Path -Path $StagingDir -ChildPath $SourceUpn
        $ManifestPath = Join-Path -Path $UserStagingDir -ChildPath "manifest.json"
        $UploadedLogPath = Join-Path -Path $UserStagingDir -ChildPath "uploaded.log"
        
        if (-not (Test-Path $ManifestPath)) { continue }
        
        # Read the uploaded.log (one FileId per line) to know what's already done
        $UploadedIds = @{}
        if (Test-Path $UploadedLogPath) {
            Get-Content $UploadedLogPath -Encoding UTF8 | ForEach-Object {
                if ($_.Trim()) { $UploadedIds[$_.Trim()] = $true }
            }
        }
        
        # Read manifest (read-only, we never write to it)
        $Manifest = @{}
        try {
            $jsonObj = Get-Content $ManifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
            foreach ($prop in $jsonObj.PSObject.Properties) {
                $v = $prop.Value
                $Manifest[$prop.Name] = @{
                    Id           = $v.Id
                    Path         = $v.Path
                    Size         = $v.Size
                    ParentPath   = $v.ParentPath
                    Downloaded   = [bool]$v.Downloaded
                    Skipped      = if ($null -ne $v.Skipped) { [bool]$v.Skipped } else { $false }
                    Uploaded     = [bool]$v.Uploaded
                    DestFolderId = $v.DestFolderId
                }
            }
        } catch {
            continue # manifest currently being written by downloader
        }
        
        if (-not $FolderCaches.ContainsKey($TargetUpn)) {
            Write-Host " -> Authenticating and checking OneDrive for ($TargetUpn)... " -NoNewline
            $DriveId = Ensure-OneDriveProvisioned -UPN $TargetUpn
            if (-not $DriveId) {
                Write-Host "[WAITING FOR CLOUD]" -ForegroundColor Yellow
                continue
            }
            Write-Host "[READY: Drive $DriveId]" -ForegroundColor Green
            $FolderCaches[$TargetUpn] = @{ "_DriveId" = $DriveId; "root" = "root" }
        }
        
        $FolderCache = $FolderCaches[$TargetUpn]
        $UserDriveId = $FolderCache["_DriveId"]
        $UserToken   = Get-UserGraphToken -UPN $TargetUpn
        
        # Find files that are Downloaded but NOT in uploaded.log and NOT permanently skipped
        $PendingFiles = @($Manifest.Values | Where-Object {
            $_.Downloaded -eq $true -and -not $_.Skipped -and -not $UploadedIds.ContainsKey($_.Id)
        })
        
        if ($isFilter) {
            if ($PendingFiles.Count -eq 0) {
                Write-Host "`n[NO READY FILES ON DISK FOR $TargetUpn - EXITING WORKER JOB]" -ForegroundColor Yellow
                return
            }
        }

        $fileCounter = 0
        foreach ($Item in $PendingFiles) {
            $fileCounter++
            if ($fileCounter % 15 -eq 0) {
                $UserToken = Get-UserGraphToken -UPN $TargetUpn
            }
            $FilePath = Sanitize-LocalPath (Join-Path -Path $UserStagingDir -ChildPath $Item.Path)
            if ($FilePath -match '^[a-zA-Z]:\\' -and $FilePath -notmatch '^\\\\\?\\') { $FilePath = "\\?\" + $FilePath }
            if (-not (Test-Path -LiteralPath $FilePath)) { continue }
            try {
                $testStream = [System.IO.File]::Open($FilePath, 'Open', 'Read', 'ReadWrite')
                $testStream.Close()
                $testStream.Dispose()
            } catch {
                continue # File is currently locked/being downloaded - skip silently
            }
            
            $ParentPath = $Item.ParentPath
            $ParentId = "root"
            
            if ($ParentPath -ne '') {
                if (-not $FolderCache.ContainsKey($ParentPath)) {
                    $FolderParts = $ParentPath -split '/'
                    $CurrentPath = ""
                    $CurrentParentId = "root"
                    foreach ($Part in $FolderParts) {
                        if ($CurrentPath -eq "") { $CurrentPath = $Part } else { $CurrentPath = "$CurrentPath/$Part" }
                        if (-not $FolderCache.ContainsKey($CurrentPath)) {
                            $NewFolder = New-DriveFolderDirect -Token $UserToken -DriveId $UserDriveId -ParentItemId $CurrentParentId -FolderName $Part
                            if ($NewFolder) {
                                $FolderCache[$CurrentPath] = $NewFolder.id
                                $CurrentParentId = $NewFolder.id
                            } else {
                                break
                            }
                        } else {
                            $CurrentParentId = $FolderCache[$CurrentPath]
                        }
                    }
                }
                $ParentId = $FolderCache[$ParentPath]
            }
            
            if ($null -eq $ParentId) { continue }
            
            $FileName = Split-Path -Path $Item.Path -Leaf
            Write-Host " [UPLOADING] ($TargetUpn) $($Item.Path) ..." -ForegroundColor Cyan
            $UploadResult = Upload-DriveItemDirect -Token $UserToken -DriveId $UserDriveId -ParentItemId $ParentId -FilePath $FilePath -FileName $FileName -UPN $TargetUpn
            
            if ($UploadResult) {
                Write-Host " [SUCCESS & CLEANED] $($Item.Path)" -ForegroundColor Green
                # Append FileId to uploaded.log resiliently
                try {
                    Add-Content -Path $UploadedLogPath -Value $Item.Id -Encoding UTF8 -ErrorAction Stop
                } catch {
                    Start-Sleep -Milliseconds 300
                    try { Add-Content -Path $UploadedLogPath -Value $Item.Id -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
                }
                # Delete local file to free disk space
                Remove-Item -LiteralPath $FilePath -Force -ErrorAction SilentlyContinue
            } else {
                Write-Host " [FAILED] $($Item.Path)" -ForegroundColor Red
            }
        }
        
        if ($isFilter) {
            return
        }
        
        $uploadedInLoop++
    }
    
    if ($uploadedInLoop -eq 0) {
        Start-Sleep -Seconds 3
    }
}
