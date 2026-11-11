<#
.SYNOPSIS
Main script to migrate OneDrive files from source tenant to destination tenant.

Requires (minimum Graph scopes): source download - User.Read.All, Files.Read.All,
Sites.Read.All, Directory.Read.All, Organization.Read.All; destination upload
(delegated/device-code admin session) - User.Read.All, Files.ReadWrite.All,
Sites.ReadWrite.All, Directory.Read.All, Organization.Read.All; destination
scan (app-only, client credentials) - Files.Read.All, User.Read.All.
Organization.Read.All is needed by the tenant check (Assert-Tenant); on the
app-only paths the app registration must grant it as an Application permission.
#>

[CmdletBinding()]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$MappingPath = '.\Output\UserMapping.csv',
    [switch]$SkipDownload,
    [switch]$SkipUpload = $false,
    [string]$UserFilter
)

# Shared safety helpers (tenant guard, app-only token, required config values).
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
# NOTE (S16): the old '-notmatch' guard for a '/goal' prefix was dead code -
# no UPN ever starts with '/goal' (it looks left over from another project) -
# so it has been removed.
$LogSuffix = if ($UserFilter) { "_" + ($UserFilter -replace '[^a-zA-Z0-9]', '_') } else { "_master" }
$LogPath = Join-Path -Path $LogDir -ChildPath "MigrateOneDrive_${Timestamp}${LogSuffix}_${PID}.log"

function Write-Log {
    param ([string]$Message, [string]$Type = "INFO")
    $LogMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Type] $Message"
    Write-Host $LogMsg -ForegroundColor $(if ($Type -eq 'ERROR') {'Red'} elseif ($Type -eq 'WARN') {'Yellow'} else {'Cyan'})
    Add-Content -Path $LogPath -Value $LogMsg -Encoding UTF8 -ErrorAction SilentlyContinue
}

Write-Log "Starting OneDrive Migration"
Write-Log "ConfigPath: $ConfigPath, MappingPath: $MappingPath, SkipDownload: $SkipDownload, SkipUpload: $SkipUpload"

# Load config for tenant info
# Load config for tenant info
if (-not (Test-Path $ConfigPath)) {
    $ConfigPath = Join-Path $ScriptDir 'MigrationConfig.psd1'
}
$config = @{}
if (Test-Path $ConfigPath) {
    $config = Import-PowerShellDataFile -Path $ConfigPath
}

if (-not (Test-Path $MappingPath)) {
    Write-Log "Mapping file not found at $MappingPath" -Type ERROR
    exit
}
$Mapping = Import-Csv -Path $MappingPath -Encoding UTF8
# NOTE (S16): UserFilter is now an exact UPN, matched with -eq, not a regex.
$isFilter = [bool]$UserFilter
if ($isFilter) {
    $Mapping = $Mapping | Where-Object { $_.SourceUPN -eq $UserFilter -or $_.DestUPN -eq $UserFilter }
    Write-Log "Filtered to user(s): $UserFilter ($($Mapping.Count) users)"
}

# Multi-account parallel download orchestration via background jobs
$MaxConcurrentUsers = if ($null -ne $config.MaxConcurrentUsers -and $config.MaxConcurrentUsers -gt 0) { [int]$config.MaxConcurrentUsers } else { 1 }
$MaxConcurrentUsers = [Math]::Min($MaxConcurrentUsers, $Mapping.Count)
if (-not $isFilter -and $Mapping.Count -gt 0) {
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host "   MULTI-ACCOUNT PARALLEL ONEDRIVE DOWNLOAD ORCHESTRATOR" -ForegroundColor Cyan
    Write-Host "   Concurrent Threads: $MaxConcurrentUsers | Total Accounts: $($Mapping.Count)" -ForegroundColor Cyan
    Write-Host "==========================================================`n" -ForegroundColor Cyan
    
    $pendingUsers = [System.Collections.ArrayList]@($Mapping)
    $runningJobs = @{}
    $retryCounts = @{}  # S2: Failed/Stopped re-queues per SourceUPN, capped at 3
    $scriptPath = $MyInvocation.MyCommand.Definition

    Write-Host "[STEP 1] Authenticating with Source Tenant before launching parallel threads..." -ForegroundColor Yellow
    $env:AZURE_IDENTITY_DISABLE_WAM = 'true'
    $sourceTenantId = $config.SourceTenantDomain
    $sourceScopes = @('User.Read.All', 'Files.Read.All', 'Sites.Read.All', 'Directory.Read.All', 'Organization.Read.All')
    Connect-MgGraph -Scopes $sourceScopes -NoWelcome -ErrorAction SilentlyContinue
    if (-not (Get-MgContext)) {
        if ($sourceTenantId) {
            Connect-MgGraph -Scopes $sourceScopes -TenantId $sourceTenantId -UseDeviceCode -NoWelcome *>&1 | Write-Host
        } else {
            Connect-MgGraph -Scopes $sourceScopes -UseDeviceCode -NoWelcome *>&1 | Write-Host
        }
    }
    Assert-Tenant -ExpectedTenantDomain (Get-RequiredConfigValue -Value $config.SourceTenantDomain -Name 'SourceTenantDomain')
    Write-Host "Connected as $((Get-MgContext).Account). Launching parallel download threads...`n" -ForegroundColor Green

    try {
        while ($pendingUsers.Count -gt 0 -or $runningJobs.Count -gt 0) {
            while ($runningJobs.Count -lt $MaxConcurrentUsers -and $pendingUsers.Count -gt 0) {
                $user = $pendingUsers[0]
                $pendingUsers.RemoveAt(0)
                $upn = $user.SourceUPN
                Write-Host "[LAUNCHING] Thread for account: $upn" -ForegroundColor Green
                $job = Start-Job -Name "DL-$upn" -ScriptBlock {
                    param($script, $upnFilter)
                    & $script -UserFilter $upnFilter -SkipUpload:$false
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
                        Write-Host "[COMPLETED] Account: $key" -ForegroundColor Green
                    } else {
                        $output | ForEach-Object { Write-Host "   $_" -ForegroundColor DarkRed }
                        # S2: re-queue a failed/stopped user at most 3 times, then give up.
                        $retryCounts[$key] = ([int]$retryCounts[$key]) + 1
                        if ($retryCounts[$key] -le 3) {
                            Write-Host "[FAILED/STOPPED] Account: $key - State: $($j.State) - Re-queueing to retry ($($retryCounts[$key])/3)..." -ForegroundColor Yellow
                            $mappingRow = $Mapping | Where-Object { $_.SourceUPN -eq $key } | Select-Object -First 1
                            if ($mappingRow) { [void]$pendingUsers.Add($mappingRow) }
                        } else {
                            Write-Host "[FAILED] Account: $key - State: $($j.State) - exceeded 3 retries, giving up (recorded as failed)." -ForegroundColor Red
                        }
                    }
                    Remove-Job -Job $j -Force -ErrorAction SilentlyContinue | Out-Null
                }
            }
            foreach ($c in $completed) {
                $runningJobs.Remove($c)
            }
            
            if ($runningJobs.Count -gt 0) {
                Write-Host "`n--- LIVE DOWNLOAD PROGRESS ($($runningJobs.Count) accounts active) ---" -ForegroundColor Cyan
                foreach ($upn in $runningJobs.Keys) {
                    $manPath = Join-Path -Path $ScriptDir -ChildPath "MigrationStaging\$upn\manifest.json"
                    if (Test-Path $manPath) {
                        try {
                            $content = Get-Content $manPath -Raw -ErrorAction Stop | ConvertFrom-Json
                            $total = @($content.PSObject.Properties).Count
                            $dlCount = @($content.PSObject.Properties | Where-Object { $_.Value.Downloaded -eq $true }).Count
                            if ($total -eq 0) {
                                Write-Host "   [$upn] : Scanning drive items..." -ForegroundColor DarkGray
                            } else {
                                Write-Host "   [$upn] : $dlCount / $total files downloaded" -ForegroundColor Yellow
                            }
                        } catch {}
                    } else {
                        Write-Host "   [$upn] : Scanning drive items..." -ForegroundColor DarkGray
                    }
                }
                Write-Host "--------------------------------------------------------" -ForegroundColor Cyan
                Start-Sleep -Seconds 5
            }
        }
        Write-Host "`n[ALL ACCOUNTS DOWNLOAD FINISHED]" -ForegroundColor Green
    } finally {
        foreach ($job in $runningJobs.Values) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }
        Get-Job | Where-Object { $_.Name -like "DL-*" } | ForEach-Object { Stop-Job -Job $_ -ErrorAction SilentlyContinue; Remove-Job -Job $_ -Force -ErrorAction SilentlyContinue }
    }
    return
}

$StagingDir = Join-Path -Path $ScriptDir -ChildPath 'MigrationStaging'
if (-not (Test-Path $StagingDir)) { New-Item -ItemType Directory -Path $StagingDir -Force | Out-Null }

$Summary = @()

# ------------- DESTINATION SCAN HELPERS -------------
function Get-DestUserGraphToken {
    <#
    .SYNOPSIS
    Get a destination-tenant app-only Graph access token (client credentials,
    in-memory only). Replaces the old ROPC password grant (B3), which fails
    on tenants with MFA / Conditional Access.
    NOTE: an app-only token has no "/me" identity, so callers must use
    /users/{UPN}/... endpoints, never /me/... (see Get-DestExistingFiles).
    #>
    param (
        [Parameter(Mandatory=$true)][string]$TenantDomain
    )

    try {
        $destClientId = Get-RequiredConfigValue -Value $config.DestClientId -Name 'DestClientId'
        $destClientSecret = Get-RequiredConfigValue -Value $config.DestClientSecret -Name 'DestClientSecret'
        return Get-AppOnlyGraphAccessToken -TenantDomain $TenantDomain -ClientId $destClientId -ClientSecret $destClientSecret
    } catch {
        Write-Log "[DEST AUTH] Failed to get app-only token for tenant $TenantDomain : $_" -Type WARN
        return $null
    }
}
function Get-DestExistingFiles {
    <#
    .SYNOPSIS
    Scan a destination user's OneDrive via Delta API and return a hashtable of "relativePath|size" keys.
    Uses Invoke-RestMethod with Bearer token (NOT Invoke-MgGraphRequest which targets the source tenant admin session).
    #>
    param (
        [Parameter(Mandatory=$true)][string]$Token,
        [Parameter(Mandatory=$true)][string]$UPN
    )

    $existingFiles = @{}
    $headers = @{ Authorization = "Bearer $Token" }

    try {
        # Get the user's drive ID. App-only token: /users/{UPN}/drives, never the per-user me endpoint.
        $escapedUpn = [System.Uri]::EscapeDataString($UPN)
        $drivesResp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$escapedUpn/drives" -Headers $headers -ErrorAction Stop
        $driveId = $drivesResp.value[0].id
        if (-not $driveId) {
            Write-Log "[DEST SCAN] Could not determine destination drive ID." -Type WARN
            return $existingFiles
        }

        # Page through Delta API results
        $uri = "https://graph.microsoft.com/v1.0/drives/$driveId/root/delta"
        $pageCount = 0

        while ($uri) {
            $pageCount++
            if ($pageCount -ge 1500) {
                Write-Log "[DEST SCAN] Reached safety limit of 1500 delta pages ($($existingFiles.Count) items), breaking." -Type WARN
                break
            }
            $response = Invoke-RestMethod -Uri $uri -Headers $headers -ErrorAction Stop
            $deltaItems = $response.value

            foreach ($child in $deltaItems) {
                $name = $child.name
                if (-not $name -or $name -eq 'root') { continue }
                if ($child.deleted) { continue }
                # Skip folders
                if ($null -ne $child.folder) { continue }

                $parentPath = ""
                if ($child.parentReference -and $child.parentReference.path) {
                    $parentPath = $child.parentReference.path -replace '^.*root:?/?', ''
                    $parentPath = $parentPath -replace '^/', '' -replace '/$', ''
                }

                $itemPath = if ($parentPath) { "$parentPath/$name" } else { $name }
                $lookupKey = "$itemPath|$($child.size)"
                $existingFiles[$lookupKey] = $true
            }

            # Next page or done
            if ($response.'@odata.nextLink') {
                $uri = $response.'@odata.nextLink'
            } else {
                $uri = $null
            }
        }
    } catch {
        Write-Log "[DEST SCAN] Error scanning destination drive: $_" -Type WARN
    }

    return $existingFiles
}

# ------------- DOWNLOAD PHASE -------------
if (-not $SkipDownload) {
    $env:AZURE_IDENTITY_DISABLE_WAM = 'true'
    $env:AZURE_IDENTITY_DISABLE_INTERACTIVE_AUTH = 'true'
    try { Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction SilentlyContinue } catch {}
    $sourceTenantId = $config.SourceTenantDomain
    $sourceScopes = @('User.Read.All', 'Files.Read.All', 'Sites.Read.All', 'Directory.Read.All', 'Organization.Read.All')
    
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
            Write-Log "Background job failed to acquire token from MSAL cache for account. Aborting thread without interactive prompt." -Type ERROR
            return
        }
        Write-Log "CONNECTING TO SOURCE TENANT FOR DOWNLOAD PHASE..."
        if ($config.SourceClientId -and $config.SourceClientSecret) {
            Write-Log "Using App-Only Authentication (Client ID & Secret) for Source Tenant..." -Type INFO
            $secSecret = ConvertTo-SecureString $config.SourceClientSecret -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential($config.SourceClientId, $secSecret)
            Connect-MgGraph -ClientSecretCredential $cred -TenantId $sourceTenantId -NoWelcome *>&1 | Write-Host
        } else {
            Write-Log "A device code will appear below. Go to https://microsoft.com/devicelogin and enter it." -Type WARN
            Write-Log "Sign in with your SOURCE tenant admin account." -Type WARN
            if ($sourceTenantId) {
                Connect-MgGraph -Scopes $sourceScopes -TenantId $sourceTenantId -UseDeviceCode -NoWelcome *>&1 | Write-Host
            } else {
                Connect-MgGraph -Scopes $sourceScopes -UseDeviceCode -NoWelcome *>&1 | Write-Host
            }
        }
    }
    Assert-Tenant -ExpectedTenantDomain (Get-RequiredConfigValue -Value $config.SourceTenantDomain -Name 'SourceTenantDomain')
    Write-Log "Connected to source tenant successfully as $((Get-MgContext).Account)."
    
    foreach ($User in $Mapping) {
        $SourceUpn = $User.SourceUPN
        $SourceId = $User.SourceId
        Write-Log "Processing downloads for user $SourceUpn"
        
        $UserStagingDir = Join-Path -Path $StagingDir -ChildPath $SourceUpn
        if (-not (Test-Path $UserStagingDir)) { New-Item -ItemType Directory -Path $UserStagingDir -Force | Out-Null }
        
        $ManifestPath = Join-Path -Path $UserStagingDir -ChildPath "manifest.json"
        $UploadedLogPath = Join-Path -Path $UserStagingDir -ChildPath "uploaded.log"
        $Manifest = @{}
        if (Test-Path $ManifestPath) {
            $jsonObj = Get-Content $ManifestPath -Raw | ConvertFrom-Json
            foreach ($prop in $jsonObj.PSObject.Properties) {
                $v = $prop.Value
                # Convert PSCustomObject entries to hashtables so we can freely add new fields
                $Manifest[$prop.Name] = @{
                    Id            = $v.Id
                    Path          = $v.Path
                    Size          = $v.Size
                    ParentPath    = $v.ParentPath
                    Downloaded    = [bool]$v.Downloaded
                    Skipped       = if ($null -ne $v.Skipped) { [bool]$v.Skipped } else { $false }
                    Uploaded      = [bool]$v.Uploaded
                    DestFolderId  = $v.DestFolderId
                }
            }
        }
        # Load uploaded.log to skip files that were already uploaded & cleaned
        $UploadedIds = @{}
        if (Test-Path $UploadedLogPath) {
            Get-Content $UploadedLogPath -Encoding UTF8 | ForEach-Object {
                if ($_.Trim()) { $UploadedIds[$_.Trim()] = $true }
            }
        }
        try { $SourceDriveId = Get-UserOneDriveId -UserId $SourceId } catch { Write-Log "Cannot access OneDrive for $SourceUpn : $_" -Type WARN; continue }
        if (-not $SourceDriveId) {
            Write-Log "Cannot access OneDrive for $SourceUpn (may be unprovisioned or 0 files)" -Type WARN
            continue
        }
        $Items = Get-AllDriveItemsFast -UserId $SourceId
        
        $Files = $Items | Where-Object { -not $_.IsFolder }
        $TotalFiles = $Files.Count
        $CurrentFile = 0

        # Pre-populate complete Manifest and save immediately so orchestrator and uploaders see the true total count
        foreach ($File in $Files) {
            $FileId = $File.Id
            $DestPath = Sanitize-LocalPath (Join-Path -Path $UserStagingDir -ChildPath $File.Path)
            if ($DestPath -match '^[a-zA-Z]:\\' -and $DestPath -notmatch '^\\\\\?\\') { $DestPath = "\\?\" + $DestPath }

            $isUp = $UploadedIds.ContainsKey($FileId)
            if (-not $Manifest.ContainsKey($FileId)) {
                $Manifest[$FileId] = @{
                    Id = $File.Id
                    Path = $File.Path
                    Size = $File.Size
                    ParentPath = $File.ParentPath
                    Downloaded = ((Test-Path -LiteralPath $DestPath) -or $isUp)
                    Skipped = $false
                    Uploaded = $isUp
                    DestFolderId = $null
                }
            } elseif ($isUp) {
                $Manifest[$FileId].Downloaded = $true
                $Manifest[$FileId].Uploaded = $true
            }
        }
        try { $Manifest | ConvertTo-Json -Depth 5 | Set-Content -Path $ManifestPath -Encoding UTF8 -Force } catch {}
        
        # Scan destination OneDrive to skip files that already exist there
        $DestExistingFiles = @{}
        $DestUpn = $User.DestUPN
        $DestTenantDomain = Get-RequiredConfigValue -Value $config.DestTenantDomain -Name 'DestTenantDomain'
        $destToken = Get-DestUserGraphToken -TenantDomain $DestTenantDomain
        if ($destToken) {
            $DestExistingFiles = Get-DestExistingFiles -Token $destToken -UPN $DestUpn
            Write-Log "[DEST SCAN] Found $($DestExistingFiles.Count) files already on destination for $DestUpn"
        } else {
            Write-Log "[DEST SCAN] Could not authenticate as $DestUpn - skipping destination scan (will download all files)" -Type WARN
        }

        foreach ($File in $Files) {
            $CurrentFile++
            $pctComplete = if ($TotalFiles -gt 0) { [int](($CurrentFile / $TotalFiles) * 100) } else { 100 }
            Write-Progress -Activity "Downloading OneDrive for $SourceUpn" -Status "File $CurrentFile of $TotalFiles" -PercentComplete $pctComplete
            
            $FileId = $File.Id
            $DestPath = Sanitize-LocalPath (Join-Path -Path $UserStagingDir -ChildPath $File.Path)
            if ($DestPath -match '^[a-zA-Z]:\\' -and $DestPath -notmatch '^\\\\\?\\') { $DestPath = "\\?\" + $DestPath }

            # Skip if already uploaded (in uploaded.log)
            if ($UploadedIds.ContainsKey($FileId)) {
                continue
            }
            # Skip if permanently failed in a previous run (antivirus block, file deleted from source, etc.)
            if ($Manifest.ContainsKey($FileId) -and $Manifest[$FileId].Skipped) {
                continue
            }
            # Skip if already downloaded AND local file still exists on disk
            if ($Manifest.ContainsKey($FileId) -and $Manifest[$FileId].Downloaded -and (Test-Path -LiteralPath $DestPath)) {
                continue
            }
            # Skip if file already exists on destination OneDrive (same path + size)
            $destLookupKey = "$($File.Path)|$($File.Size)"
            if ($DestExistingFiles.ContainsKey($destLookupKey)) {
                $Manifest[$FileId].Downloaded = $true
                $Manifest[$FileId].Uploaded = $true
                # Append to uploaded.log so upload worker also skips
                try { Add-Content -Path $UploadedLogPath -Value $FileId -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
                Write-Log "[SKIP-DEST-EXISTS] $($File.Path) (already on destination)"
                continue
            }
            $Result = Download-DriveItem -UserId $SourceId -DriveId $SourceDriveId -DriveItemId $FileId -DestinationPath $DestPath -DownloadUrl $File.DownloadUrl
            $Success = $Result.Success
            $IsPermanent = $Result.PermanentlyFailed
            
            $Manifest[$FileId].Downloaded = $Success
            if ($IsPermanent) { $Manifest[$FileId].Skipped = $true }
            
            if ($Success) {
                Write-Log "Downloaded: $($File.Path)"
            } elseif ($IsPermanent) {
                Write-Log "Permanently skipped (deleted/blocked): $($File.Path)" -Type WARN
            } else {
                Write-Log "Failed to download: $($File.Path)" -Type ERROR
            }
            # Save manifest every 100 files or at the end to prevent 25MB JSON serialization bottleneck on every single file
            if ($CurrentFile % 100 -eq 0 -or $CurrentFile -eq $TotalFiles) {
                try { $Manifest | ConvertTo-Json -Depth 5 | Set-Content -Path $ManifestPath -Encoding UTF8 -Force } catch {}
            }
        }
        try { $Manifest | ConvertTo-Json -Depth 5 | Set-Content -Path $ManifestPath -Encoding UTF8 -Force } catch {}
    }
}

# ------------- UPLOAD PHASE -------------
if (-not $SkipUpload) {
    Write-Log "CONNECTING TO DESTINATION TENANT FOR UPLOAD PHASE..."
    Write-Log "A device code will appear below. Go to https://microsoft.com/devicelogin and enter it." -Type WARN
    Write-Log "Sign in with your DESTINATION tenant admin account." -Type WARN
    $destScopes = @('User.Read.All', 'Files.ReadWrite.All', 'Sites.ReadWrite.All', 'Directory.Read.All', 'Organization.Read.All')
    $destTenantId = $config.DestTenantDomain
    if ($destTenantId) {
        Connect-MgGraph -Scopes $destScopes -TenantId $destTenantId -UseDeviceCode -NoWelcome *>&1 | Write-Host
    } else {
        Connect-MgGraph -Scopes $destScopes -UseDeviceCode -NoWelcome *>&1 | Write-Host
    }
    Assert-Tenant -ExpectedTenantDomain (Get-RequiredConfigValue -Value $config.DestTenantDomain -Name 'DestTenantDomain')
    Write-Log "Connected to destination tenant successfully."
    
    foreach ($User in $Mapping) {
        $SourceUpn = $User.SourceUPN
        $TargetUpn = $User.DestUPN
        $TargetId = $User.DestId
        Write-Log "Processing uploads for user $TargetUpn"
        
        $UserStagingDir = Join-Path -Path $StagingDir -ChildPath $SourceUpn
        $ManifestPath = Join-Path -Path $UserStagingDir -ChildPath "manifest.json"
        
        if (-not (Test-Path $ManifestPath)) {
            Write-Log "Manifest not found for $SourceUpn. Skipping upload." -Type WARN
            continue
        }
        
        $Manifest = @{}
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
        
        # Trigger OneDrive provisioning
        try {
            $null = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$TargetId/drive/root" -ErrorAction Stop
        } catch {
            Write-Log "Drive for $TargetUpn not provisioned. Waiting 10 seconds..." -Type WARN
            Start-Sleep -Seconds 10
        }
        
        $FolderCache = @{"root" = "root"} # maps relative path to DriveItemId
        
        $DownloadedItems = $Manifest.Values | Where-Object { $_.Downloaded -eq $true }
        $TotalItems = @($DownloadedItems).Count
        $CurrentItem = 0
        
        $UploadedCount = 0
        $FailedCount = 0
        $TotalSize = 0
        
        foreach ($Item in $DownloadedItems) {
            $CurrentItem++
            Write-Progress -Activity "Uploading OneDrive for $TargetUpn" -Status "File $CurrentItem of $TotalItems" -PercentComplete (($CurrentItem/$TotalItems)*100)
            
            if ($Item.Uploaded -eq $true) {
                Write-Log "Skipping already uploaded file: $($Item.Path)"
                $UploadedCount++
                $TotalSize += $Item.Size
                continue
            }
            
            $FilePath = Sanitize-LocalPath (Join-Path -Path $UserStagingDir -ChildPath $Item.Path)
            if ($FilePath -match '^[a-zA-Z]:\\' -and $FilePath -notmatch '^\\\\\?\\') { $FilePath = "\\?\" + $FilePath }
            if (-not (Test-Path -LiteralPath $FilePath)) {
                Write-Log "Local file missing: $FilePath" -Type ERROR
                $FailedCount++
                continue
            }
            
            $ParentPath = $Item.ParentPath
            $ParentId = "root"
            
            if ($ParentPath -ne '') {
                if (-not $FolderCache.ContainsKey($ParentPath)) {
                    # We need to build folders
                    $FolderParts = $ParentPath -split '/'
                    $CurrentPath = ""
                    $CurrentParentId = "root"
                    foreach ($Part in $FolderParts) {
                        $PrevPath = $CurrentPath
                        if ($CurrentPath -eq "") { $CurrentPath = $Part } else { $CurrentPath = "$CurrentPath/$Part" }
                        if (-not $FolderCache.ContainsKey($CurrentPath)) {
                            $NewFolder = New-DriveFolder -UserId $TargetId -ParentItemId $CurrentParentId -FolderName $Part
                            if ($NewFolder) {
                                $FolderCache[$CurrentPath] = $NewFolder.id
                                $CurrentParentId = $NewFolder.id
                            } else {
                                Write-Log "Failed to create folder structure for $ParentPath" -Type ERROR
                                break
                            }
                        } else {
                            $CurrentParentId = $FolderCache[$CurrentPath]
                        }
                    }
                }
                $ParentId = $FolderCache[$ParentPath]
            }
            
            if ($null -eq $ParentId) {
                Write-Log "Skipping upload due to folder creation failure: $($Item.Path)" -Type ERROR
                $FailedCount++
                continue
            }
            
            $FileName = Split-Path -Path $Item.Path -Leaf
            $UploadResult = Upload-DriveItem -UserId $TargetId -DestDriveId "" -ParentItemId $ParentId -FilePath $FilePath -FileName $FileName
            
            if ($UploadResult) {
                Write-Log "Uploaded: $($Item.Path)"
                $Item.Uploaded = $true
                $UploadedCount++
                $TotalSize += $Item.Size
                # Clean up local file immediately to save laptop disk space
                Remove-Item -Path $FilePath -Force -ErrorAction SilentlyContinue
            } else {
                Write-Log "Failed to upload: $($Item.Path)" -Type ERROR
                $FailedCount++
            }
        }
        
        $Summary += [PSCustomObject]@{
            SourceUpn = $SourceUpn
            TargetUpn = $TargetUpn
            TotalDownloaded = $TotalItems
            Uploaded = $UploadedCount
            Failed = $FailedCount
            TotalSizeBytes = $TotalSize
        }
        
        # Save updated manifest
        try { $Manifest | ConvertTo-Json -Depth 5 | Set-Content -Path $ManifestPath -Encoding UTF8 -Force } catch {}
    }
}

Write-Log "MIGRATION SUMMARY"
$Summary | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor Cyan
$Summary | Export-Csv -Path (Join-Path -Path $LogDir -ChildPath "OneDriveMigrationSummary_$Timestamp.csv") -NoTypeInformation -Encoding UTF8

Write-Log "Migration script completed."
